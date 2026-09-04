# STRUCTURE 主程序结构索引（main.py v3.99，约3923行）

> 用途：给新对话的 AI 定位代码用。先看此索引，需要细节时再粘贴对应代码片段。
> 维护：每次 AI 改动 main.py 后，顺手更新本文件的行号与说明。
> 2026-09-04 校准：v3.99 新增 shift_scope（L1492 起 45 行）+ shift 内联块替换（-30 行），
> 净 +15 行；L1492 之前行号不变，L1492 之后行号 = 旧行号 + 15。

## 模块级函数（main.py）
- format_numfmt_readable(fmt) L68   数字格式码→中文可读描述
- excel_tint(hex6, tint) L74        ECMA-376 主题色 tint 变换（RGB→HLS→RGB）
- rgb_channel_close(h1, h2) L88     颜色逐通道差<=1（tint舍入噪声容忍）
- rgb_euclidean(h1, h2) L97         RGB欧氏距离（调色板吸附容差，<=60 视为同色）
- dedup_rich_text_diffs(diffs) L136 富文本重复差异去重
- merge_adjacent_diffs(diffs) L140  相邻同类差异合并
- color_signature(color) L162       颜色存储签名比较
- cell_address(col,row) L737        列号行号→地址字符串
- normalize_path(p) L738             路径归一化
- get_sheet_names_fast(path) L741   快速读 sheet 名（流式）
- formula_text(value) L770          公式解析为文本
- normalize_formula(f) L780         公式归一化
- _fmt_duration(secs) L785          耗时格式化
- parse_rich_text_from_xlsx(path) L813  解析富文本
- compare_rich_text_runs(r1, r2) L859  富文本 run 对比

## 核心类
- WorkbookStyleCache L196    样式缓存：主题色/自定义数字格式/字体等价表 font_equiv
- DataLocator L880           高级审核定位器（anchor/search_in/range 查找）
- PluginManager L1033        数据检查插件（均值偏差/参数锁定/范围检查）
- CheckProject L1347         检查项目集；from_dict(json) 加载；to_dict() 导出
- OpenpyxlComparer L1354     对比引擎（核心）
  · run() L1426 → _load_workbooks L1442 → _run_diff_mode L1462
  · shift_scope() L1492      【v3.99 提取】shift 列配对+垂直范围（引擎与测试按钮共用）
  · _apply_rule_filter L1537 → _build_diff_desc L1680（差异描述生成）
  · 属性：diffs / stats / sheet_diffs / old_wb_ref / new_wb_ref / old_cache / new_cache
- ExcelCOMVerifier L2423     Excel COM 显示层采集器
  · collect_style_data(diffs,...) L2449 写入 diff['com_style']
  · _read_cell_style L2522 / _connect_excel L2620 / _ensure_workbooks L2638
- ComCheckDialog L2704       高级审核确认窗（Excel激活警告+检查COM通道，测试按钮复用）
- RuleEditorDialog L3101     规则编辑窗（测试按钮由 gui_patch.py 注入；on_ok 组装 ds）
- DiffViewer L3264           原版 GUI（start_compare L3637 / _finalize_diag L3543 / export_diag L3577）
  · jump_to_excel L3869      跳转并选中（支持联合多区，如 B48:C49,E5）

## 差异类型名（判定/断言用）
内容变化 / 公式变化 / 富文本变化 / 字体变化 / 填充变化 / 边框变化 / 对齐变化 /
数字格式变化 / 合并新增 / 合并删除 / 行高变化 / 列宽变化 / 图片新增 / 图片变动 /
图片尺寸变化 / 条件格式新增 / 条件格式删除 / 条件格式修改 / 条件格式变化 /
单元格新增 / 单元格删除 / 高级检查

## 自动化小文件（AI 创建的独立文件）
- cli.py → auto_run.exe：main() 读 scenarios.json → OpenpyxlComparer + ExcelCOMVerifier + ai_summary
- ai_summary.py：write_summary / write_focus / write_package / group_patterns
- regression.py：CASES 列表（14个案例）+ run_one() + main()

## 关键约定
- 值变化类差异的类型名 = 「内容变化」
- 测试机运行 auto_run 前先手动打开 Excel（关激活窗）
- shift 规则：垂直范围必须合法非空，否则引擎跳过；old列+shift_offset=新列
- 模式 ID（P001...）由 ai_summary 动态生成，按数量降序
- 版本号在摘要头部（version.txt，Actions 注入，如 ac36b3c）
