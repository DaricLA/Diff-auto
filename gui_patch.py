# -*- coding: utf-8 -*-
"""
规则编辑窗「数据源测试」补丁：
在数据源配置区底部加【测试】按钮 → 按当前表单配置计算命中单元格/范围
→ 弹窗显示计算结果 → 确认后通过 Excel COM 跳转并选中。
支持：offset 单格 / intersection 交叉格 / range 多格联合选中(含排除格、不连续)
      / shift 新旧文件双双跳转；排除区间 A1-A9、B2-D2、A1:A9。
不修改 main.py 任何代码，通过运行时替换类方法实现。
"""
import os, re, time
import main
from tkinter import messagebox


def _parse_rows(spec):
    rows = set()
    for part in str(spec or '').split(','):
        p = part.strip()
        if not p:
            continue
        if '-' in p:
            a, b = p.split('-', 1)
            try:
                rows.update(range(int(a), int(b) + 1))
            except Exception:
                pass
        else:
            try:
                rows.add(int(p))
            except Exception:
                pass
    return sorted(rows)


def _match_sheet(wb, sheet):
    """精确匹配；失败则按去首尾空格/忽略大小写匹配；返回真实名称或 None"""
    ns = list(wb.sheetnames)
    if sheet in ns:
        return sheet
    for s in ns:
        if str(s).strip().lower() == str(sheet).strip().lower():
            return s
    return None


def _group_addresses(addrs):
    """把单元格地址列表合并为 Excel 联合选择串，如 'B48:C49,E5'。返回 (联合串, 格数)"""
    pts = []
    for a in addrs:
        m = re.match(r'^([A-Za-z]+)(\d+)$', str(a))
        if not m:
            continue
        c = main.column_index_from_string(m.group(1).upper())
        r = int(m.group(2))
        pts.append((r, c))
    if not pts:
        return '', 0
    pts.sort()
    rows = {}
    for r, c in pts:
        rows.setdefault(r, []).append(c)
    segs = []
    for r in sorted(rows):
        cols = sorted(set(rows[r]))
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
            else:
                segs.append((r, start, prev))
                start = prev = c
        segs.append((r, start, prev))
    segs.sort()
    ranges = []
    i = 0
    n = len(segs)
    while i < n:
        r, c1, c2 = segs[i]
        j = i
        while (j + 1 < n and segs[j+1][1] == c1 and segs[j+1][2] == c2
               and segs[j+1][0] == segs[j][0] + 1):
            j += 1
        top = segs[i][0]
        bottom = segs[j][0]
        a1 = main.cell_address(c1, top)
        a2 = main.cell_address(c2, bottom)
        ranges.append(a1 if a1 == a2 else '%s:%s' % (a1, a2))
        i = j + 1
    return ','.join(ranges), len(pts)


def _parse_exclude(self, text):
    """增强版排除解析：支持 单格A5 / 相对[1,0] / 区间A1-A9、B2-D2、A1:A9（自动展开）"""
    text = (text or '').replace('，', ',')
    if not text.strip():
        return []
    parts = [p.strip() for p in text.split(',') if p.strip()]
    exclude = []
    for p in parts:
        if re.match(r'^[A-Za-z]+\d+$', p):
            exclude.append(p.upper())
        elif re.match(r'^\[\d+,\d+\]$', p):
            inner = p[1:-1].split(',')
            exclude.append([int(inner[0]), int(inner[1])])
        elif re.match(r'^[A-Za-z]+\d+[-:][A-Za-z]+\d+$', p):
            m = re.match(r'^([A-Za-z]+)(\d+)[-:]([A-Za-z]+)(\d+)$', p)
            if m:
                c1_s, r1_s, c2_s, r2_s = m.groups()
                c1 = main.column_index_from_string(c1_s.upper())
                c2 = main.column_index_from_string(c2_s.upper())
                r1, r2 = int(r1_s), int(r2_s)
                cnt = 0
                if c1 == c2:
                    col = c1_s.upper()
                    for r in range(min(r1, r2), max(r1, r2) + 1):
                        exclude.append('%s%d' % (col, r)); cnt += 1
                        if cnt > 5000:
                            break
                elif r1 == r2:
                    row = r1
                    for c in range(min(c1, c2), max(c1, c2) + 1):
                        exclude.append('%s%d' % (main.get_column_letter(c), row)); cnt += 1
                        if cnt > 5000:
                            break
                else:
                    for r in range(min(r1, r2), max(r1, r2) + 1):
                        for c in range(min(c1, c2), max(c1, c2) + 1):
                            exclude.append('%s%d' % (main.get_column_letter(c), r)); cnt += 1
                            if cnt > 5000:
                                break
    return exclude


def _ds_assemble(self):
    """与 on_ok 完全一致的 data_source 组装（去掉检查项部分）"""
    si = self.search_in_var.get().strip().upper().replace('$', '')
    try:
        ds_si = si if main.DataLocator._parse_area(si) else 'all'
    except Exception:
        ds_si = 'all'
    ds = {'name': self.rule_name_var.get(), 'sheet': self.sheet_var.get(),
          'anchor': {'text': self.anchor_text_var.get()}, 'search_in': ds_si,
          'mode': self.mode_var.get()}
    mode = self.mode_var.get()
    if mode == 'offset':
        ds['target'] = {'row_offset': self._to_int(self.offset_row_var), 'col_offset': self._to_int(self.offset_col_var)}
    elif mode == 'intersection':
        ds['row_anchor'] = {'text': self.row_anchor_text_var.get().strip(), 'search_in': ds_si}
        ds['col_anchor'] = {'text': self.col_anchor_text_var.get().strip(), 'search_in': ds_si}
    elif mode == 'range':
        ds['target'] = {'row_offset': self._to_int(self.range_row_offset_var),
                        'col_offset': self._to_int(self.range_col_offset_var),
                        'row_count': self._to_int(self.range_row_count_var, 1) or 1,
                        'col_count': self._to_int(self.range_col_count_var, 1) or 1,
                        'exclude': _parse_exclude(self, self.range_exclude_var.get())}
    elif mode == 'shift':
        ds['header_target'] = {'row_offset': self._to_int(self.sh_ro_var),
                               'col_offset': self._to_int(self.sh_co_var),
                               'row_count': max(1, self._to_int(self.sh_rc_var, 1)),
                               'col_count': max(1, self._to_int(self.sh_cc_var, 20))}
        ds['rows'] = self.sh_rows_var.get().strip()
        ds['shift_offset'] = self._to_int(self.sh_offset_var)
    return ds


def _jump(self, file_path, sheet_name, addr):
    """复用主程序跳转逻辑：连接Excel→激活工作簿→滚动并选中（addr 可为联合多区，如 'B48:C49,E5'）"""
    try:
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            import win32com.client as win32com
        except Exception:
            return False, 'win32com 不可用'
        try:
            excel = win32com.GetActiveObject("Excel.Application")
        except Exception:
            try:
                os.startfile(file_path)
            except Exception:
                return False, '无法打开文件（请确认路径有效）'
            excel = None
            deadline = time.time() + 40
            while time.time() < deadline:
                time.sleep(1)
                try:
                    excel = win32com.GetActiveObject("Excel.Application")
                    break
                except Exception:
                    pass
            if excel is None:
                return False, 'Excel 启动超时'
        wb = None
        try:
            for w in excel.Workbooks:
                if main.normalize_path(w.FullName) == main.normalize_path(file_path):
                    wb = w
                    break
        except Exception:
            pass
        if wb is None:
            try:
                wb = excel.Workbooks.Open(file_path, ReadOnly=True)
            except Exception as e:
                return False, '打开工作簿失败: %s' % e
        wb.Activate()
        try:
            ws = wb.Worksheets(sheet_name)
        except Exception as e:
            return False, '工作表不存在: %s' % e
        ws.Activate()
        first = addr.replace(' ', '').split(',')[0]
        if ':' in first:
            first = first.split(':')[0]
        col = ''.join(ch for ch in first if ch.isalpha())
        row = ''.join(ch for ch in first if ch.isdigit())
        if not col or not row:
            return False, '无效地址: %s' % addr
        app = wb.Application
        app.ActiveWindow.ScrollRow = int(row)
        app.ActiveWindow.ScrollColumn = main.column_index_from_string(col)
        ws.Range(addr).Select()
        return True, None
    except Exception as e:
        return False, str(e)


def _ds_test(self):
    import openpyxl
    try:
        mode = self.mode_var.get() or 'offset'
        sheet = self.sheet_var.get()          # 保留原始名称（可能含尾部空格）
        if not sheet:
            messagebox.showwarning('测试', '请先选择 Sheet'); return
        new_path = self.new_path
        if not new_path or not os.path.isfile(new_path):
            messagebox.showwarning('测试', '待检文件不存在，请先选择文件'); return
        ds = _ds_assemble(self)
        self.config(cursor='watch')
        try:
            wb = openpyxl.load_workbook(new_path, data_only=False)
            real_sheet = _match_sheet(wb, sheet)
            if real_sheet is None:
                messagebox.showwarning('测试', 'Sheet 不存在（当前文件实际名称: %s）' % '、'.join(wb.sheetnames[:15]))
                return
            sheet = real_sheet
            loc = main.DataLocator()
            jump_addrs = None   # 列表: [(文件路径, 地址)]
            if mode == 'shift':
                old_path = self.old_path
                if not old_path or not os.path.isfile(old_path):
                    messagebox.showwarning('测试', 'shift 模式需要旧版文件，请先选择'); return
                owb = openpyxl.load_workbook(old_path, data_only=False)
                old_sheet = _match_sheet(owb, sheet)
                if old_sheet is None:
                    messagebox.showwarning('测试', '旧版文件中无该 Sheet: %s' % sheet); return
                ows = owb[old_sheet]
                nws = wb[sheet]
                hdr_t = ds.get('header_target', {})
                o_ac = loc._merge_search_in(ds.get('anchor', {}), ds.get('search_in', ''))
                o_loc = loc._range_cfg(ows, o_ac, hdr_t, '标题行范围(旧)')
                n_loc = loc._range_cfg(nws, o_ac, hdr_t, '标题行范围(新)')
                errs = [x for x in (o_loc.get('error'), n_loc.get('error')) if x]
                if errs:
                    messagebox.showwarning('测试', '定位失败: %s' % '；'.join(errs)); return
                shift_offset = ds.get('shift_offset', 0)
                pairs = []
                c1 = o_loc.get('c1', o_loc['start'][1])
                c2 = o_loc.get('c2', o_loc['start'][1] + hdr_t.get('col_count', 1) - 1)
                r1 = o_loc.get('r1', o_loc['start'][0])
                r2 = o_loc.get('r2', o_loc['start'][0] + hdr_t.get('row_count', 1) - 1)
                for c in range(c1, c2 + 1):
                    has = False
                    for r in range(r1, r2 + 1):
                        v = ows.cell(r, c).value
                        if v is not None and str(v).strip() != '':
                            has = True; break
                    if has:
                        nc = c + shift_offset
                        if nc >= 1:
                            pairs.append((c, nc))
                if not pairs:
                    messagebox.showinfo('测试', '旧报告标题区无数据列，未生成配对'); return
                rows = _parse_rows(ds.get('rows', ''))
                demo_row = rows[0] if rows else n_loc.get('r1', 1)
                first_oc = pairs[0][0]
                first_nc = pairs[0][1]
                old_addr = main.cell_address(first_oc, demo_row)
                new_addr = main.cell_address(first_nc, demo_row)
                lines = ['【shift 配对结果】共 %d 列配对（固定偏移 %+d）' % (len(pairs), shift_offset), '']
                if len(pairs) <= 40:
                    for oc, nc in pairs:
                        lines.append('  %s列 → %s列' % (main.get_column_letter(oc), main.get_column_letter(nc)))
                else:
                    for oc, nc in pairs[:20]:
                        lines.append('  %s列 → %s列' % (main.get_column_letter(oc), main.get_column_letter(nc)))
                    lines.append('  … 共%d列' % len(pairs))
                lines.append('')
                lines.append('垂直范围: %s | 示例目标: 旧%s!%s 新%s!%s' % (
                    ds.get('rows') or '(未填→从头起)', old_sheet, old_addr, sheet, new_addr))
                msg = '\n'.join(lines)
                jump_addrs = [(old_path, old_sheet, old_addr), (new_path, sheet, new_addr)]
            else:
                loc.rules = [ds]
                res = loc.locate_all(wb).get(ds.get('name', ''))
                if not isinstance(res, dict):
                    messagebox.showwarning('测试', '计算失败（规则未命中定位逻辑）'); return
                if 'error' in res:
                    messagebox.showwarning('测试', '定位失败: %s' % res['error']); return
                if mode == 'range':
                    addrs = res.get('addresses') or ([res.get('address')] if res.get('address') else [])
                    union, cnt = _group_addresses(addrs)
                    if cnt > 2000 or len(union.split(',')) > 300:
                        b1 = main.cell_address(res.get('c1', 1), res.get('r1', 1))
                        b2 = main.cell_address(res.get('c2', 1), res.get('r2', 1))
                        jump_addr = '%s:%s' % (b1, b2)
                        msg = '【range 命中】\n命中 %d 格（数量多，已选中整体范围）\n%s!%s:%s' % (cnt, sheet, b1, b2)
                    else:
                        jump_addr = union
                        msg = '【range 命中】\n选中 %d 格: %s!%s' % (cnt, sheet, union)
                    jump_addrs = [(new_path, sheet, jump_addr)]
                else:
                    addr = res.get('address')
                    msg = '【%s 命中】\n单元格: %s!%s\n值: %s' % (mode, sheet, addr, res.get('value'))
                    jump_addrs = [(new_path, sheet, addr)]
            if not jump_addrs:
                messagebox.showinfo('测试', msg); return
            if messagebox.askyesno('测试结果', msg + '\n\n是否跳转并选中？'):
                errs = []
                for fp, sh, ad in jump_addrs:
                    ok, err = _jump(self, fp, sh, ad)
                    if not ok:
                        errs.append('%s: %s' % (main.normalize_path(fp), err))
                if errs:
                    messagebox.showwarning('跳转失败', '\n'.join(errs))
                # 成功时不弹多余窗口
        finally:
            self.config(cursor='')
    except Exception as e:
        messagebox.showerror('测试异常', str(e))


def apply():
    _orig_build_ui = main.RuleEditorDialog._build_ui

    def _patched_build_ui(self):
        _orig_build_ui(self)
        try:
            btn_row = main.tb.Frame(self.param_frame.master)
            btn_row.pack(fill='x', pady=(6, 2))
            main.tb.Button(btn_row, text="测试", bootstyle="success-outline", width=8,
                           command=self._ds_test).pack(side='left')
        except Exception:
            pass
    main.RuleEditorDialog._build_ui = _patched_build_ui
    main.RuleEditorDialog._ds_test = _ds_test
    main.RuleEditorDialog._parse_exclude = _parse_exclude   # 增强排除解析，同步影响真实规则保存
