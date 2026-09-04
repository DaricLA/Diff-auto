# -*- coding: utf-8 -*-
"""
AI 摘要生成器：把 DIAG 诊断数据压缩为高信号文本
- 模式归并：同类型/同列/同特征的差异合并为 P001/P002...
- COM 显示层判定：有 com_old/com_new 时自动给出「疑似误报/疑似真实」
- 枚举映射：数字格式、颜色、字体的 旧→新 映射计数
"""
import re, json, os
from collections import Counter

HEAD_KEYS = ['fill_color', 'fill_pattern', 'font_name', 'font_color', 'font_size',
             'font_bold', 'num_format', 'h_align', 'wrap_text',
             'row_height', 'col_width']


def _norm(desc):
    s = str(desc or '')
    s = re.sub(r'\$?[A-Za-z]{1,3}\$?\d+', 'CELL', s)
    s = re.sub(r'-?\d+(?:\.\d+)?', 'N', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s[:60]


def _addr_col(addr):
    a = str(addr or '')
    if ':' in a:
        a = a.split(':')[0]
    m = re.match(r'^([A-Za-z]+)', a)
    return m.group(1).upper() if m else '?'


def _short(v, n=36):
    s = str(v if v is not None else '')
    return s if len(s) <= n else s[:n] + '…'


def _com_fields(d):
    o = d.get('com_old') or {}
    n = d.get('com_new') or {}
    if not o and not n:
        return None
    same, diff = True, []
    for k in HEAD_KEYS:
        if k in o and k in n and o[k] != n[k]:
            same = False
            diff.append('%s:%s→%s' % (k, _short(o[k], 14), _short(n[k], 14)))
    if not diff:
        return True, []
    return False, diff


def group_patterns(diag):
    """返回按数量降序的模式列表"""
    groups = {}
    for d in diag.get('diffs') or []:
        if not d.get('type'):
            continue
        key = (str(d.get('type')), str(d.get('sheet')), _addr_col(d.get('address')), _norm(d.get('desc')))
        groups.setdefault(key, []).append(d)
    pats = []
    for i, (key, items) in enumerate(sorted(groups.items(), key=lambda kv: -len(kv[1])), 1):
        pats.append({'id': 'P%03d' % i, 'type': key[0], 'sheet': key[1], 'col': key[2],
                     'sig': key[3], 'count': len(items), 'items': items})
    return pats


def _pattern_head(p):
    return '%s | x%d | %s | Sheet"%s" 列%s | 特征: %s' % (
        p['id'], p['count'], p['type'], p['sheet'], p['col'], p['sig'] or '(无)')


def _sample_line(d, i, p):
    addr = d.get('address', '?')
    oldv = _short(d.get('old_value') or (d.get('com_old') or {}).get('fill_color', ''))
    newv = _short(d.get('new_value') or (d.get('com_new') or {}).get('fill_color', ''))
    part = '%s#%d | %s!%s | 旧:%s 新:%s' % (p['id'], i, d.get('sheet', '?'), addr, oldv, newv)
    cm = _com_fields(d)
    if cm is not None:
        if cm[0]:
            part += ' | COM显示层一致 ★疑似误报'
        else:
            part += ' | COM显示层不同[%s] ★疑似真实' % '; '.join(cm[1][:3])
    desc = _short(d.get('desc'), 160)
    if desc:
        part += ' | %s' % desc
    return part


def write_summary(diag, out_path, max_samples=3, max_patterns=40, max_chars=26000):
    L = []
    meta = diag.get('meta') or {}
    L.append('===== 对比诊断摘要 =====')
    L.append('版本: %s | 时间: %s' % (meta.get('version', '?'), meta.get('time', '?')))
    L.append('旧文件: %s' % os.path.basename(meta.get('old_path', '?')))
    L.append('新文件: %s' % os.path.basename(meta.get('new_path', '?')))
    if diag.get('error'):
        L.append('⚠ 运行错误: %s' % diag['error'])
    st = diag.get('stats') or {}
    if st:
        L.append('【统计】总单元格: %s | 差异: %s 处 | 新增sheet: %s | 删除sheet: %s' % (
            st.get('total_cells', '?'), st.get('diff_cells', '?'),
            len(st.get('added_sheets') or []), len(st.get('removed_sheets') or [])))
    diffs = diag.get('diffs') or []
    cnt = Counter(d.get('type', '?') for d in diffs)
    L.append('【按类型计数】' + ' | '.join('%s:%d' % (k, v) for k, v in cnt.most_common(20)))
    sd = diag.get('sheet_diffs') or []
    if sd:
        L.append('【Sheet差异】%d 处: %s' % (len(sd), '; '.join(_short(str(x), 60) for x in sd[:5])))

    pats = group_patterns(diag)

    # ---- 【类型概览】按类型统计：差异处数 | 模式数 | 代表位置（放在模式清单之前，避免截断丢失）----
    by_type = {}
    for p in pats:
        t = p['type']
        if t not in by_type:
            first = p['items'][0]
            by_type[t] = {'count': 0, 'pats': 0, 'id': p['id'],
                          'pos': '%s!%s' % (p['sheet'], first.get('address', '?'))}
        by_type[t]['count'] += p['count']
        by_type[t]['pats'] += 1
    if by_type:
        L.append('')
        L.append('【类型概览】差异处数 | 模式数 | 代表位置(该类型最大模式)')
        for t, info in sorted(by_type.items(), key=lambda kv: -kv[1]['count']):
            L.append('  %s | %d 处 | %d 模式 | %s [%s]' % (
                t, info['count'], info['pats'], info['pos'], info['id']))

    L.append('')
    L.append('【差异模式归并】共 %d 个模式（同类合并，样本展示前 %d 条）' % (len(pats), max_samples))
    written = 0
    for p in pats:
        if written > max_patterns:
            L.append('… 其余 %d 个模式略（需要时用 focus_pattern 深入）' % (len(pats) - written))
            break
        L.append(_pattern_head(p))
        for i, d in enumerate(p['items'][:max_samples], 1):
            L.append('  ' + _sample_line(d, i, p))
        written += 1
        if sum(len(x) for x in L) > max_chars:
            L.append('… (摘要超出长度上限，其余内容请用 focus_pattern 深入)')
            break

    rules = {}
    for d in diffs:
        rn = d.get('rule_name')
        if rn:
            r = rules.setdefault(rn, [0, 0])
            r[1] += 1
            if d.get('rule_pass'):
                r[0] += 1
    if rules:
        L.append('')
        L.append('【规则检查】')
        for rn, (ok, total) in rules.items():
            L.append('  %s: 通过 %d / %d' % (rn, ok, total))

    lg = (diag.get('log') or '').strip().splitlines()
    if lg:
        L.append('')
        L.append('【日志尾部】')
        L.extend('  ' + ln[:120] for ln in lg[-25:])

    L.append('')
    L.append('【下一步】如需某个模式的详细样本，在 scenarios.json 中加 "focus_pattern": "P002" 后重新双击 auto_run.exe。')
    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(L))
    except Exception as e:
        print('write_summary 失败: %s' % e)
    return out_path


def write_focus(diag, pattern_id, out_path, sample_n=20):
    pats = group_patterns(diag)
    p = next((x for x in pats if x['id'] == pattern_id), None)
    L = []
    if p is None:
        L.append('未找到模式 %s（可先用 AI_SUMMARY 查看现有模式 ID）' % pattern_id)
    else:
        L.append('===== 定向取样 %s =====' % pattern_id)
        L.append(_pattern_head(p))
        L.append('')
        for i, d in enumerate(p['items'][:sample_n], 1):
            L.append('--- 样本 %d ---' % i)
            for k in ('sheet', 'address', 'type', 'desc', 'rule_name', 'rule_pass',
                      'old_value', 'new_value', 'old_numfmt', 'new_numfmt', 'shift_paired_old',
                      'com_old', 'com_new', 'advanced_check', 'rule_expect', 'rule_diff_desc'):
                if k in d and d[k] not in (None, ''):
                    L.append('  %s: %s' % (k, _short(d[k], 800)))
        L.append('')
        L.append('（共 %d 条，以上展示前 %d 条）' % (p['count'], min(sample_n, p['count'])))
    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(L))
    except Exception as e:
        print('write_focus 失败: %s' % e)
    return out_path


def write_package(results, out_path):
    parts = ['======== AI 诊断包（一次粘贴给 AI） ========', '']
    for r in results:
        parts.append('########## 案例: %s ##########' % r.get('case', '?'))
        if r.get('error'):
            parts.append('⚠ 该案例运行出错: %s' % r['error'])
        sp = r.get('summary')
        try:
            if sp and os.path.isfile(sp):
                with open(sp, 'r', encoding='utf-8') as f:
                    parts.append(f.read().rstrip())
            else:
                parts.append('(无摘要)')
        except Exception as e:
            parts.append('(读取摘要失败: %s)' % e)
        parts.append('')
    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(parts))
    except Exception as e:
        print('write_package 失败: %s' % e)
    return out_path

