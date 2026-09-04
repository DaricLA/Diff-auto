# -*- coding: utf-8 -*-
"""
excel-diff 自动运行（命令行/双击运行）
双击 auto_run.exe = 读取同目录 scenarios.json 自动完成：
  1) 跑对比引擎          -> out/DIAG_FULL_<case>.json  (完整诊断,大文件留测试机)
  2) Excel COM 显示层采集-> 嵌入每条差异 (com_old/com_new)
  3) 生成 AI 摘要        -> out/AI_SUMMARY_<case>.txt  (小文件,给AI)
  4) 合并打包            -> out/AI_PACKAGE.txt          (一次粘给AI的合集)
"""
import os, sys, json, time, re, threading

BASE = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, BASE)

import main  # 导入主程序模块（不会弹出图形界面）
from main import (OpenpyxlComparer, ExcelCOMVerifier, CheckProject,
                  merge_adjacent_diffs, _fmt_duration, DEFAULT_CHECK_OPTIONS)

try:
    import ai_summary  # 第三批文件，缺失时自动降级
except Exception:
    ai_summary = None

try:
    import pythoncom
except Exception:
    pythoncom = None

MAX_LEN_LIST = {'diffs', 'sheet_diffs'}
_log_buf = []


def log(msg):
    try:
        print("[%s] %s" % (time.strftime('%H:%M:%S'), msg))
    except Exception:
        pass
    _log_buf.append(msg)
    if len(_log_buf) > 5000:
        del _log_buf[:1000]


def get_version():
    for d in (getattr(sys, '_MEIPASS', None), BASE):
        if not d:
            continue
        p = os.path.join(d, 'version.txt')
        if os.path.isfile(p):
            try:
                return open(p, encoding='utf-8').read().strip() or 'unknown'
            except Exception:
                pass
    return getattr(main, 'VERSION', 'unknown')


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', '_', str(s))


def resolve_path(p):
    if not p:
        return ''
    return p if os.path.isabs(p) else os.path.join(BASE, p)


def load_project(path):
    if not path:
        return None
    p = resolve_path(path)
    if not os.path.isfile(p):
        log("!! 规则文件不存在，跳过: %s" % p)
        return None
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return CheckProject.from_dict(data)
    except Exception as e:
        log("!! 规则文件解析失败，跳过: %s" % e)
        return None


def rebuild_diag(comparer, old, new, check_options, project, error=None):
    """与 GUI 的 _finalize_diag 等价，产出 AI 摘要需要的完整 diag 字典"""
    diag = {'meta': {'version': get_version(), 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
                     'old_path': old, 'new_path': new, 'options': check_options,
                     'project': (project.to_dict() if project else None)},
            'stats': {}, 'sheet_diffs': [], 'diffs': [], 'log': '\n'.join(_log_buf), 'error': error}
    try:
        if comparer is not None:
            diag['stats'] = getattr(comparer, 'stats', {}) or {}
            diag['sheet_diffs'] = getattr(comparer, 'sheet_diffs', []) or []
            _owb = getattr(comparer, 'old_wb_ref', None)
            _nwb = getattr(comparer, 'new_wb_ref', None)
            _dd = []
            for d in getattr(comparer, 'diffs', []) or []:
                e = {k: v for k, v in d.items() if k != 'com_style'}
                if d.get('com_style'):
                    e['com_old'] = d['com_style'].get('old')
                    e['com_new'] = d['com_style'].get('new')
                try:
                    if _owb and _nwb and d.get('sheet') in _owb.sheetnames:
                        _m = getattr(comparer, '_last_shift_old_map', {}) if d.get('type') == '单元格删除' else getattr(comparer, '_last_shift_new_map', {})
                        _hit = _m.get((d.get('sheet'), d.get('address')))
                        _osh, _oad = (_hit[0][1], _hit[0][2]) if _hit else (d.get('sheet'), d.get('address'))
                        for _tag, _wb, _sh, _ad in (('old', _owb, _osh, _oad), ('new', _nwb, d.get('sheet'), d.get('address'))):
                            _cs = ''.join(ch for ch in _ad if ch.isalpha())
                            _rs = ''.join(ch for ch in _ad if ch.isdigit())
                            if _cs and _rs and _sh in _wb.sheetnames:
                                _r, _c = int(_rs), main.column_index_from_string(_cs)
                                e['%s_value' % _tag] = str(_wb[_sh].cell(_r, _c).value)
                                e['%s_numfmt' % _tag] = _wb[_sh].cell(_r, _c).number_format
                        if _hit:
                            e['shift_paired_old'] = "%s!%s" % (_osh, _oad)
                except Exception:
                    pass
                _dd.append(e)
            diag['diffs'] = _dd
    except Exception as e:
        log("!! 汇总诊断数据失败: %s" % e)
    return diag


def run_case(cfg, out_dir):
    name = str(cfg.get('name', '案例')) or '案例'
    old = resolve_path(cfg.get('old', ''))
    new = resolve_path(cfg.get('new', ''))
    log("=" * 12 + " 案例: %s " + "=" * 12)
    if not old or not new or not os.path.isfile(old) or not os.path.isfile(new):
        log("!! 文件不存在，跳过该案例（请检查 scenarios.json 路径）: %s → %s" % (old, new))
        return None
    project = load_project(cfg.get('project', ''))
    check_options = dict(DEFAULT_CHECK_OPTIONS)
    extra_opts = cfg.get('check_options')
    if isinstance(extra_opts, dict):
        check_options.update(extra_opts)
    do_com = bool(cfg.get('verify', True))
    do_focus = str(cfg.get('focus_pattern', '') or '') or None

    stop_event = threading.Event()
    comparer = None
    error = None
    try:
        if pythoncom:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        comparer = OpenpyxlComparer(old, new, log,
                                    lambda v, s=None: None,
                                    check_options=check_options,
                                    plugin_manager=None,
                                    progress_mode_fn=lambda m: None,
                                    check_project=project,
                                    stop_event=stop_event,
                                    mode='diff', color_tolerance=0)
        comparer.run()
        if do_com:
            log("启动 Excel COM 显示层采集（隐藏实例自动打开文件，无需人工操作）...")
            try:
                verifier = ExcelCOMVerifier(old, new, log, progress_fn=lambda v, s=None: None, progress_mode_fn=lambda m: None)
                old_fe = getattr(getattr(comparer, 'old_cache', None), 'font_equiv', None)
                new_fe = getattr(getattr(comparer, 'new_cache', None), 'font_equiv', None)
                ok, fail = verifier.collect_style_data(comparer.diffs, old_font_equiv=old_fe, new_font_equiv=new_fe)
                log("COM 采集完成：成功 %d 条，失败 %d 条" % (ok, fail))
            except Exception as e:
                log("!! COM 采集异常（跳过，不影响对比）: %s" % e)
        if comparer.check_project:
            log("执行进阶规则过滤...")
            try:
                comparer._apply_rule_filter(comparer.diffs, comparer.old_wb_ref, comparer.new_wb_ref)
            except Exception as e:
                log("!! 规则过滤异常: %s" % e)
        log("对比阶段耗时 %s | 差异: %d 处单元格, %d 处 Sheet" % (
            _fmt_duration(0), len(comparer.diffs), len(comparer.sheet_diffs)))
    except KeyboardInterrupt:
        log("!! 用户中止")
    except Exception as e:
        log("!! 运行异常: %s" % e)
        error = str(e)

    diag = rebuild_diag(comparer, old, new, check_options, project, error=error)
    os.makedirs(out_dir, exist_ok=True)
    full_path = os.path.join(out_dir, "DIAG_FULL_%s.json" % safe_name(name))
    try:
        with open(full_path, 'w', encoding='utf-8') as f:
            json.dump(diag, f, ensure_ascii=False, indent=1, default=str)
        log("完整诊断已保存: %s" % full_path)
    except Exception as e:
        log("!! 保存完整诊断失败: %s" % e)

    summary_path = os.path.join(out_dir, "AI_SUMMARY_%s.txt" % safe_name(name))
    try:
        if ai_summary is not None:
            ai_summary.write_summary(diag, summary_path,
                                     max_samples=int(cfg.get('max_samples_per_pattern', 3) or 3))
            log("AI 摘要已生成: %s" % summary_path)
        else:
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write("ai_summary 模块缺失，无法生成摘要\n")
            log("!! ai_summary 模块缺失（第三批文件未上传）")
    except Exception as e:
        log("!! 生成摘要失败: %s" % e)

    if do_focus and ai_summary is not None:
        try:
            focus_path = os.path.join(out_dir, "FOCUS_%s.txt" % safe_name(do_focus))
            ai_summary.write_focus(diag, do_focus, focus_path, sample_n=20)
            log("定向取样已生成: %s" % focus_path)
        except Exception as e:
            log("!! 定向取样失败: %s" % e)

    return {'case': name, 'summary': summary_path, 'full': full_path, 'diag': diag, 'error': error}


def main():
    import argparse
    ap = argparse.ArgumentParser(description="Excel 对比工具自动运行")
    ap.add_argument('--scenarios', default=os.path.join(BASE, 'scenarios.json'))
    ap.add_argument('--out', default='')
    ap.add_argument('--case', default='', help='只运行指定名称的案例')
    ap.add_argument('--no-verify', action='store_true', help='跳过 Excel COM 采集')
    args = ap.parse_args()

    sp = args.scenarios if os.path.isabs(args.scenarios) else os.path.join(BASE, args.scenarios)
    if not os.path.isfile(sp):
        log("!! 找不到 %s" % sp)
        log("请在程序同目录放置 scenarios.json（模板见仓库），然后重新双击。")
        try:
            input("按回车关闭...")
        except Exception:
            pass
        return 1

    try:
        with open(sp, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
    except Exception as e:
        log("!! scenarios.json 解析失败: %s" % e)
        try:
            input("按回车关闭...")
        except Exception:
            pass
        return 1

    out_dir = args.out or str(cfg.get('output_dir', 'out') or 'out')
    out_dir = out_dir if os.path.isabs(out_dir) else os.path.join(BASE, out_dir)
    os.makedirs(out_dir, exist_ok=True)

    cases = cfg.get('cases') or []
    if not cases:
        log("!! scenarios.json 中没有 cases")
        try:
            input("按回车关闭...")
        except Exception:
            pass
        return 1

    results = []
    for c in cases:
        nm = str(c.get('name', ''))
        if args.case and nm != args.case:
            continue
        if args.no_verify:
            c = dict(c); c['verify'] = False
        results.append(run_case(c, out_dir))

    try:
        if ai_summary is not None:
            pkg = os.path.join(out_dir, "AI_PACKAGE.txt")
            ai_summary.write_package([r for r in results if r], pkg)
            log("AI 打包文件已生成: %s" % pkg)
    except Exception as e:
        log("!! 打包失败: %s" % e)

    log("=" * 40)
    log("全部完成！请把 out 文件夹中的 AI_PACKAGE.txt（或各 AI_SUMMARY_*.txt）拷回手机，粘贴给 AI。")
    log("完整诊断 DIAG_FULL_*.json 留在本机即可（很大，不用传）。")
    try:
        input("按回车关闭窗口...")
    except Exception:
        pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
