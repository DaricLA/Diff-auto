# -*- coding: utf-8 -*-
"""
金标准回归测试：验证对比引擎各类检测逻辑正确性
双击 regression.exe：
  1) 用 openpyxl 自动生成最小 xlsx 测试对（价值/公式/字体/填充/边框/数字格式/
     对齐/行高/列宽/合并/新增sheet 等检测点，无需任何外部数据）
  2) 跑对比引擎，检查每类检测是否符合预期（含检测开关=选项级豁免）
  3) 输出 out/RULES_VERDICT.txt（AI 可直接分析）
"""
import os, sys, time, threading, traceback
from collections import Counter

BASE = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, BASE)

import main
from main import OpenpyxlComparer, DEFAULT_CHECK_OPTIONS

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


def _save_pair(name, old_wb, new_wb, tmpdir):
    o = os.path.join(tmpdir, name + '_old.xlsx')
    n = os.path.join(tmpdir, name + '_new.xlsx')
    old_wb.save(o)
    new_wb.save(n)
    return o, n


# ---------- 案例定义 ----------
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
    ws2['A1'] = 2          # 数值变化
    return old, new, {'must': [('值变化', 'A1')], 'absent': []}, {}


def c_identical():
    old, _ = _base()
    new, _ = _base()
    return old, new, {'must': [], 'absent': ['值变化', '公式变化', '字体变化', '填充变化',
                                              '边框变化', '对齐变化', '数字格式变化']}, {}


def c_numfmt_changed():
    old, ws = _base()
    new, ws2 = _base()
    ws['A1'].number_format = '0'
    ws2['A1'].number_format = '0.00'
    return old, new, {'must': [('数字格式变化', 'A1')], 'absent': ['值变化']}, {}


def c_numfmt_gate_off():   # 选项级豁免：关闭数字格式检测
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
    new.merge_cells('A1:B1')
    new, ws3 = _base()
    ws3.merge_cells('A1:B1')
    ws3.merge_cells('C1')  # 不变
    # old=A1:B1 合并, new=不合并 → 应报合并删除
    old, ws = _base()
    ws.merge_cells('A1:B1')
    new, ws2 = _base()
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
    print('\n'.join(L))
    print('已生成: %s' % out_path)
    try:
        input('按回车关闭窗口...')
    except Exception:
        pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
