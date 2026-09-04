# -*- coding: utf-8 -*-
"""
规则编辑窗「数据源测试」补丁：
在数据源配置区底部加【测试】按钮 → 按当前表单配置计算命中单元格/范围
→ 弹窗显示计算结果 → 确认后通过 Excel COM 跳转并选中（复用主程序跳转逻辑）。
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
                        'exclude': self._parse_exclude(self.range_exclude_var.get())}
    elif mode == 'shift':
        ds['header_target'] = {'row_offset': self._to_int(self.sh_ro_var),
                               'col_offset': self._to_int(self.sh_co_var),
                               'row_count': max(1, self._to_int(self.sh_rc_var, 1)),
                               'col_count': max(1, self._to_int(self.sh_cc_var, 20))}
        ds['rows'] = self.sh_rows_var.get().strip()
        ds['shift_offset'] = self._to_int(self.sh_offset_var)
    return ds


def _jump(self, file_path, sheet_name, addr):
    """复用主程序跳转逻辑：连接Excel→激活工作簿→滚动并选中"""
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
        first = addr.split(':')[0] if ':' in addr else addr
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
        sheet = self.sheet_var.get().strip()
        if not sheet:
            messagebox.showwarning('测试', '请先选择 Sheet'); return
        new_path = self.new_path
        if not new_path or not os.path.isfile(new_path):
            messagebox.showwarning('测试', '待检文件不存在，请先选择文件'); return
        ds = _ds_assemble(self)
        self.config(cursor='watch')
        try:
            wb = openpyxl.load_workbook(new_path, data_only=False)
            if sheet not in wb.sheetnames:
                messagebox.showwarning('测试', 'Sheet 不存在: %s' % sheet); return
            loc = main.DataLocator()
            jump_addr = None
            if mode == 'shift':
                old_path = self.old_path
                if not old_path or not os.path.isfile(old_path):
                    messagebox.showwarning('测试', 'shift 模式需要旧版文件，请先选择'); return
                owb = openpyxl.load_workbook(old_path, data_only=False)
                ows = owb[sheet] if sheet in owb.sheetnames else None
                nws = wb[sheet]
                if ows is None:
                    messagebox.showwarning('测试', '旧版文件中无该 Sheet'); return
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
                first_nc = pairs[0][1]
                jump_addr = main.cell_address(first_nc, demo_row)
                lines = ['【shift 配对结果】共 %d 列配对（固定偏移 %+d）' % (len(pairs), shift_offset), '']
                if len(pairs) <= 40:
                    for oc, nc in pairs:
                        lines.append('  %s列 → %s列' % (main.get_column_letter(oc), main.get_column_letter(nc)))
                else:
                    for oc, nc in pairs[:20]:
                        lines.append('  %s列 → %s列' % (main.get_column_letter(oc), main.get_column_letter(nc)))
                    lines.append('  … 共%d列' % len(pairs))
                lines.append('')
                lines.append('垂直范围: %s | 示例目标: %s!%s' % (ds.get('rows') or '(未填→从头起)', sheet, jump_addr))
                msg = '\n'.join(lines)
            else:
                loc.rules = [ds]
                res = loc.locate_all(wb).get(ds.get('name', ''))
                if not isinstance(res, dict):
                    messagebox.showwarning('测试', '计算失败（规则未命中定位逻辑）'); return
                if 'error' in res:
                    messagebox.showwarning('测试', '定位失败: %s' % res['error']); return
                addr = res.get('address')
                addrs = res.get('addresses')
                if mode == 'range':
                    msg = '【range 命中】\n范围: %s!%s:%s（%d格）\n起点: %s!%s' % (
                        sheet, main.cell_address(res.get('c1', 1), res.get('r1', 1)),
                        main.cell_address(res.get('c2', 1), res.get('r2', 1)),
                        res.get('range_count', 0), sheet, addr)
                    jump_addr = '%s:%s' % (main.cell_address(res.get('c1', 1), res.get('r1', 1)),
                                           main.cell_address(res.get('c2', 1), res.get('r2', 1)))
                else:
                    msg = '【%s 命中】\n单元格: %s!%s\n值: %s' % (mode, sheet, addr, res.get('value'))
                    jump_addr = addr
            if not jump_addr:
                messagebox.showinfo('测试', msg); return
            if messagebox.askyesno('测试结果', msg + '\n\n是否跳转并选中该目标？'):
                ok, err = _jump(self, new_path, sheet, jump_addr)
                if ok:
                    messagebox.showinfo('跳转', '已跳转并选中: %s!%s' % (sheet, jump_addr))
                else:
                    messagebox.showwarning('跳转失败', err)
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
