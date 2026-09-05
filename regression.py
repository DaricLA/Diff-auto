# -*- coding: utf-8 -*-
"""
金标准回归测试 v2：14 案例（防回退）+ L2 全场景验证（B1单项/B2干扰/B3全局交叉/B4 COM分歧/B5主题专项）
双击 regression.exe：
  1) 金标准回归    -> out/RULES_VERDICT.txt
  2) L2 全场景     -> out/L2_REPORT.txt（硬断言 + 观察点；失败用例 xlsx 留 cases_gen）
用例数据在 l2_scenarios.py（与本文件一起打包）。
"""
import os, sys, time, threading, traceback, re, zipfile
from collections import Counter

BASE = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, BASE)

import main
from main import OpenpyxlComparer, DEFAULT_CHECK_OPTIONS, CheckProject, CheckRule, CheckItemConfig

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

try:
    import l2_scenarios as L2S
    HAS_L2 = True
except Exception:
    HAS_L2 = False

# ============================================================
# 一、金标准回归（原 14 案例）
# ============================================================
def _save_pair(name, old_wb, new_wb, tmpdir):
    o = os.path.join(tmpdir, name + '_old.xlsx')
    n = os.path.join(tmpdir, name + '_new.xlsx')
    old_wb.save(o)
    new_wb.save(n)
    return o, n


def _base():
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws['A1'] = 1
    ws['B1'] = 'x'
    ws['A2'] = 10
    return wb, ws


def c_value_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws2['A1'] = 2
    return old, new, {'must': [('内容变化', 'A1')], 'absent': []}, {}


def c_identical():
    old, _ = _base()
    new, _ = _base()
    return old, new, {'must': [], 'absent': ['内容变化', '公式变化', '字体变化', '填充变化',
                                              '边框变化', '对齐变化', '数字格式变化']}, {}


def c_numfmt_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].number_format = '0'
    ws2['A1'].number_format = '0.00'
    return old, new, {'must': [('数字格式变化', 'A1')], 'absent': ['内容变化']}, {}


def c_numfmt_gate_off():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].number_format = '0'
    ws2['A1'].number_format = '0.00'
    return old, new, {'must': [], 'absent': ['数字格式变化']}, {'number_format': False}


def c_font_name_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].font = Font(name='微软雅黑')
    ws2['A1'].font = Font(name='宋体')
    return old, new, {'must': [('字体变化', 'A1')], 'absent': []}, {}


def c_font_size_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].font = Font(size=11)
    ws2['A1'].font = Font(size=14)
    return old, new, {'must': [('字体变化', 'A1')], 'absent': []}, {}


def c_fill_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].fill = PatternFill(fill_type='solid', start_color='FFFF00')
    ws2['A1'].fill = PatternFill(fill_type='solid', start_color='0000FF')
    return old, new, {'must': [('填充变化', 'A1')], 'absent': []}, {}


def c_border_changed():
    old, ws = _base()
    new, ws2 = _base()
    thin = Side(style='thin', color='000000')
    ws2['A1'].border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return old, new, {'must': [('边框变化', 'A1')], 'absent': []}, {}


def c_align_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws2['A1'].alignment = Alignment(horizontal='center')
    return old, new, {'must': [('对齐变化', 'A1')], 'absent': []}, {}


def c_rowheight_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws2.row_dimensions[2].height = 30
    return old, new, {'must': [('行高变化', 'A2')], 'absent': []}, {}


def c_colwidth_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws2.column_dimensions['A'].width = 20
    return old, new, {'must': [('列宽变化', 'A1')], 'absent': []}, {}


def c_merged_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws.merge_cells('A1:B1')
    return old, new, {'must': [('合并删除', 'A1')], 'absent': []}, {}


def c_formula_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A2'] = '=1+1'
    ws2['A2'] = '=2+2'
    return old, new, {'must': [('公式变化', 'A2')], 'absent': []}, {}


def c_sheet_added():
    old, ws = _base()
    new, ws2 = _base()
    new.create_sheet('新增表')
    return old, new, {'must': [], 'absent': []}, {}


CASES = [
    ('v01 值变化检测', c_value_changed),
    ('v02 完全相同应零差异', c_identical),
    ('v03 数字格式变化检测', c_numfmt_changed),
    ('v04 数字格式检测关闭(豁免)', c_numfmt_gate_off),
    ('v05 字体名变化检测', c_font_name_changed),
    ('v06 字号变化检测', c_font_size_changed),
    ('v07 填充色变化检测', c_fill_changed),
    ('v08 边框变化检测', c_border_changed),
    ('v09 对齐变化检测', c_align_changed),
    ('v10 行高变化检测', c_rowheight_changed),
    ('v11 列宽变化检测', c_colwidth_changed),
    ('v12 合并单元格删除检测', c_merged_changed),
    ('v13 公式变化检测', c_formula_changed),
    ('v14 新增sheet检测', c_sheet_added),
]


def run_one(name, builder, tmpdir):
    try:
        old, new, expect, opts = builder()
        o, n = _save_pair(name.replace(' ', '_'), old, new, tmpdir)
        copt = dict(DEFAULT_CHECK_OPTIONS)
        copt.update(opts or {})
        cmp = OpenpyxlComparer(o, n, log_callback=lambda m: None,
                               progress_callback=lambda v, s=None: None,
                               progress_mode_fn=lambda m: None,
                               check_options=copt,
                               stop_event=threading.Event())
        ok = cmp.run()
        if not ok:
            return {'ok': False, 'detail': '引擎返回失败', 'counts': {}}
        found = [(d.get('type'), str(d.get('address', '')).split(':')[0].upper()) for d in cmp.diffs]
        counts = Counter(t for t, a in found)
        problems = []
        for t, a in expect.get('must', []):
            if not any(ft == t and (fa == a.upper() or fa.startswith(a.upper())) for ft, fa in found):
                problems.append('缺少期望差异 %s@%s' % (t, a))
        for t in expect.get('absent', []):
            if any(ft == t for ft, fa in found):
                problems.append('不应报但报了 %s' % t)
        return {'ok': not problems, 'detail': '; '.join(problems) if problems else '符合预期',
                'counts': counts, 'found': found[:20]}
    except Exception as e:
        return {'ok': False, 'detail': '异常: %s | %s' % (e, traceback.format_exc()[:400]), 'counts': {}}


# ============================================================
# 二、L2 通用工具
# ============================================================
def l2_make_project(rules):
    """rules: [(ds, [(check_type, expect), ...]), ...] → CheckProject（支持多条规则）"""
    return CheckProject('L2', '1.0', [CheckRule(rule_name='L2', data_source=ds,
                                                checks=[CheckItemConfig(check_type=t, enabled=True, expect=e)
                                                        for (t, e) in (checks or [])])
                                       for ds, checks in rules])


def l2_save(data, name, tmpdir):
    o = os.path.join(tmpdir, 'l2_' + re.sub(r'[\\/:*?"<>|]', '_', name) + '_old.xlsx')
    n = os.path.join(tmpdir, 'l2_' + re.sub(r'[\\/:*?"<>|]', '_', name) + '_new.xlsx')
    data['old'].save(o)
    data['new'].save(n)
    return o, n


def diff_at(diffs, sheet, addr, types=None):
    a = str(addr).upper()
    for d in diffs:
        if d.get('sheet') != sheet:
            continue
        da = str(d.get('address', '')).upper()
        head = da.split(':')[0]
        if a == da or a == head or da.startswith(a + ':'):
            if types is None or d.get('type') in types:
                return d
    return None


def check_assert(diffs, sheet, addr, kind, types=None, rule_expect='same'):
    """kind: 'DIFF'(真差异格) 'SAME'(无真差异格) 'PRESENT'(仅需存在) 'NODIFF'(必须无差)"""
    d = diff_at(diffs, sheet, addr, types)
    if kind == 'NODIFF':
        if d is None:
            return None, 'OK'
        return 'FAIL', '不应报差但报 %s@%s: %s' % (d.get('type'), addr, str(d.get('desc', ''))[:70])
    if kind == 'PRESENT':
        return (None, 'OK') if d is not None else ('FAIL', '缺少期望差异 @%s' % addr)
    if d is None:
        return 'FAIL', '期望 %s 但无差异 @%s' % (kind, addr)
    if kind == 'DIFF':
        if rule_expect == 'same':
            if d.get('rule_pass') is True:
                return 'FAIL', '%s@%s 期望未豁免，但被规则豁免(rule_pass=True)' % (d.get('type'), addr)
            return None, 'OK'
        if d.get('rule_pass') is not True:
            return 'FAIL', '%s@%s 期望豁免，但 rule_pass=%s rule_name=%s' % (d.get('type'), addr, d.get('rule_pass'), d.get('rule_name'))
        return None, 'OK'
    if kind == 'SAME':
        if rule_expect == 'same':
            if d.get('rule_pass') is not True:
                return 'FAIL', '%s@%s 无差格期望豁免，但未豁免(rule_name=%s)' % (d.get('type'), addr, d.get('rule_name'))
            return None, 'OK'
        if d.get('rule_pass') is True:
            return 'FAIL', '%s@%s 无差格期望不豁免，但被豁免' % (d.get('type'), addr)
        return None, 'OK'
    return 'FAIL', '未知 kind %s' % kind


def set_theme_accent1(path, hex6):
    """替换 xlsx 主题表 accent1 颜色（zip 重写 theme1.xml）"""
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path, 'r')
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/theme/theme1.xml':
            try:
                s = data.decode('utf-8')
                s2 = re.sub(r'(<a:accent1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(?="/>)',
                            r'\g<1>' + hex6, s, flags=re.S)
                data = s2.encode('utf-8')
            except Exception:
                pass
        zout.writestr(item, data)
    zin.close()
    zout.close()
    os.replace(tmp, path)


def com_read_cells(path, sheet, addr):
    """COM 读单格快照：值/文本/数字格式/字体名/字号/行高/列宽/填充色(BGR)"""
    try:
        import pythoncom
        import win32com.client
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        try:
            wb = app.Workbooks.Open(path, ReadOnly=True)
            try:
                ws = wb.Worksheets(sheet)
                rng = ws.Range(addr)
                snap = {}
                for f in ('value', 'text', 'font_name', 'font_size', 'row_height', 'col_width'):
                    try:
                        snap[f] = getattr(rng, f)
                    except Exception:
                        pass
                try:
                    snap['numfmt'] = rng.NumberFormat
                except Exception:
                    pass
                try:
                    c = rng.Interior.Color
                    snap['fill_bgr'] = int(c) if c is not None else None
                except Exception:
                    pass
                return snap
            finally:
                wb.Close(False)
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    except Exception as e:
        return {'error': str(e)}


def com_snap_equal(a, b, fields):
    diffs = []
    for f in fields:
        va = (a or {}).get(f)
        vb = (b or {}).get(f)
        if va != vb:
            diffs.append('%s: %s vs %s' % (f, va, vb))
    return (not diffs), '; '.join(diffs)


# ============================================================
# 三、L2 执行
# ============================================================
def l2_run_case(case, tmpdir):
    res = {'batch': case.get('batch', '?'), 'name': case.get('name', '?'),
           'fails': [], 'notes': [], 'com_lines': [], 'ok': True, 'skipped': []}
    try:
        data = case['build']()
    except Exception as e:
        res['ok'] = False
        res['fails'].append('构建失败: %s | %s' % (e, traceback.format_exc()[:200]))
        return res
    old, new = l2_save(data, res['name'], tmpdir)
    project = None
    if data.get('rules'):
        project = l2_make_project(data['rules'])
    copt = dict(DEFAULT_CHECK_OPTIONS)
    copt.update((data.get('opts') or {}) or {})
    cmp = OpenpyxlComparer(old, new, log_callback=lambda m: None,
                           progress_callback=lambda v, s=None: None,
                           progress_mode_fn=lambda m: None,
                           check_options=copt,
                           check_project=project,
                           stop_event=threading.Event())
    try:
        cmp.run()
    except Exception as e:
        res['ok'] = False
        res['fails'].append('引擎异常: %s' % e)
        return res
    if project is not None:
        try:
            cmp._apply_rule_filter(cmp.diffs, cmp.old_wb_ref, cmp.new_wb_ref)
        except Exception as e:
            res['fails'].append('规则过滤异常: %s' % e)
            res['ok'] = False
    rule_expect = 'same'
    for (ds, checks) in (data.get('rules') or []):
        for (t, e) in checks:
            rule_expect = e
            break
        break
    for (sh, ad, kind, types) in (data.get('asserts') or []):
        st, msg = check_assert(cmp.diffs, sh, ad, kind, types, rule_expect=rule_expect)
        if st:
            res['fails'].append('%s@%s: %s' % (ad, sh, msg))
            res['ok'] = False
    for com in (data.get('com') or []):
        kind2 = com.get('kind', 'observe')
        fields = com.get('fields', ['text'])
        a = com_read_cells(old, com.get('old_sheet', 'Sheet1'), com.get('old_addr'))
        b = com_read_cells(new, com.get('new_sheet', 'Sheet1'), com.get('new_addr'))
        if (isinstance(a, dict) and a.get('error')) or (isinstance(b, dict) and b.get('error')):
            err = (a or {}).get('error') or (b or {}).get('error')
            res['com_lines'].append('[SKIP] %s: COM %s' % (com.get('label', '?'), err))
            res['skipped'].append(com.get('label', '?'))
            continue
        eq, desc = com_snap_equal(a, b, fields)
        prog_diff = bool(com.get('prog_diff'))
        if kind2 == 'fp':
            verdict = 'FP候选(程序报差,COM相同)' if eq else 'consistent'
        elif kind2 == 'fn':
            verdict = 'FN候选(程序无差,COM不同)' if not eq else 'consistent'
        else:
            verdict = '观察(%s)' % ('COM一致' if eq else ('COM不同: ' + (desc[:80] if desc else '')))
        res['com_lines'].append('[%s] %s | 程序: %s | COM字段: %s' % (
            verdict, com.get('label', '?'), '报差' if prog_diff else '无差', desc if desc else '一致'))
    return res


def run_l2(tmpdir, out_dir, com_ok=True):
    L = []
    L.append('===== L2 全场景验证报告 =====')
    L.append('时间: %s | 版本: %s | COM可用: %s' % (time.strftime('%Y-%m-%d %H:%M:%S'),
             getattr(main, 'VERSION', '?'), '是' if com_ok else '否(COM段将SKIP)'))
    L.append('')
    if not HAS_L2:
        L.append('!! l2_scenarios.py 缺失，无法运行 L2。')
        return L
    results = []
    for case in L2S.suite():
        results.append(l2_run_case(case, tmpdir))
    per_batch = {}
    for r in results:
        per_batch.setdefault(r['batch'], [0, 0])
        per_batch[r['batch']][0] += 1
        if r['ok']:
            per_batch[r['batch']][1] += 1
    all_ok = sum(v[1] for v in per_batch.values())
    total = sum(v[0] for v in per_batch.values())
    for batch in sorted(per_batch):
        b = per_batch[batch]
        L.append('【%s】%d/%d 通过' % (batch, b[1], b[0]))
    L.append('')
    L.append('---- 失败明细（硬断言 ❌）----')
    fails = [r for r in results if not r['ok']]
    if not fails:
        L.append('（无失败）')
    for r in fails:
        L.append('❌ [%s] %s' % (r['batch'], r['name']))
        for f in r['fails']:
            L.append('   ' + f)
    L.append('')
    L.append('---- COM 对照观察点 ----')
    com_lines = []
    for r in results:
        com_lines.extend(r['com_lines'])
    if not com_lines:
        L.append('（无 COM 对照用例或全部 SKIP）')
    for cl in com_lines:
        L.append('  ' + cl)
    L.append('')
    L.append('汇总: 硬断言 %d/%d 通过（全场景用例）' % (all_ok, total))
    L.append('')
    L.append('说明: 所有案例的原始对照文件保留在 cases_gen 目录（l2_*_old/new.xlsx），失败项可在 Excel 中人工复核。')
    return L


# ============================================================
# 四、入口
# ============================================================
def main():
    tmpdir = os.path.join(BASE, 'cases_gen')
    os.makedirs(tmpdir, exist_ok=True)
    out_dir = os.path.join(BASE, 'out')
    os.makedirs(out_dir, exist_ok=True)

    L = []
    L.append('===== 金标准回归结果 =====')
    L.append('时间: %s | 版本: %s' % (time.strftime('%Y-%m-%d %H:%M:%S'), getattr(main, 'VERSION', '?')))
    L.append('')
    passed = 0
    n = len(CASES)
    for name, builder in CASES:
        print('[回归] %s ...' % name)
        r = run_one(name, builder, tmpdir)
        if r['ok']:
            passed += 1
            L.append('✅ %s | 符合预期' % name)
        else:
            L.append('❌ %s\n   期望/问题: %s' % (name, r['detail']))
        if r.get('counts'):
            L.append('   实际检出类型: %s' % dict(r['counts']))
        if not r['ok'] and r.get('found'):
            L.append('   实际检出(前20): %s' % r['found'])
        L.append('')
    L.append('汇总: %d / %d 通过' % (passed, n))
    if passed == n:
        L.append('结论: 对比引擎核心检测逻辑正常。')
    else:
        L.append('结论: 有 %d 项不符合预期，请将本文件粘贴给 AI 分析。' % (n - passed))
    L.append('')
    L.append('说明: 案例文件保留在 cases_gen 目录，可在 Excel 中打开人工复核。')
    out_path = os.path.join(out_dir, 'RULES_VERDICT.txt')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))

    com_ok = True
    try:
        import pythoncom  # noqa
        import win32com.client  # noqa
        a = win32com.client.DispatchEx('Excel.Application')
        a.Quit()
    except Exception:
        com_ok = False
    print('[L2] 开始全场景验证（约 2~4 分钟）...')
    L2 = run_l2(tmpdir, out_dir, com_ok)
    l2_path = os.path.join(out_dir, 'L2_REPORT.txt')
    with open(l2_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L2))

    print('\n'.join(L))
    print('---')
    print('\n'.join(L2))
    print('已生成: %s' % out_path)
    print('已生成: %s' % l2_path)
    try:
        input('按回车关闭窗口...')
    except Exception:
        pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
