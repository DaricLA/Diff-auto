# STRUCTURE 主程序结构索引（main.py v3.99，约3942行）

> 用途：给新对话的 AI 定位代码用。先看此索引，需要细节时再粘贴对应代码片段。
> 维护：每次 AI 改动 main.py 后，顺手更新本文件的行号与说明。
> 2026-09-06 校准：拍板豁免已入 main.py（字体别名内置组 L209-215、行高/列宽默认值豁免
> _compare_row_col_dimensions L1989+）；诊断块已删；相对 9-04 基准行号已刷新（本文为准）。

## 模块级函数（main.py）
- format_numfmt_readable(fmt) L68   数字格式码→中文可读描述
- excel_tint(hex6, tint) L74        ECMA-376 主题色 tint 变换（RGB→HLS→RGB）
- rgb_channel_close(h1, h2) L88     颜色逐通道差<=1（tint舍入噪声容忍）
- rgb_euclidean(h1, h2) L97         RGB欧氏距离（调色板吸附容差，<=60 视为同色）
- dedup_rich_text_diffs(diffs) L136 富文本重复差异去重
- merge_adjacent_diffs(diffs) L140  相邻同类差异合并
- color_signature(color) L162       颜色存储签名比较
- cell_address(col,row) L744        列号行号→地址字符串
- normalize_path(p) L745            路径归一化
- get_sheet_names_fast(path) L748   快速读 sheet 名（流式）
- formula_text(value) L777          公式解析为文本
- normalize_formula(f) L787         公式归一化
- _fmt_duration(secs) L792          耗时格式化
- parse_rich_text_from_xlsx(path) L820  解析富文本
- compare_rich_text_runs(r1, r2) L866  富文本 run 对比

## 核心类
- WorkbookStyleCache L196    样式缓存：主题色/自定义数字格式/字体等价表 font_equiv
                              · font_equiv 内置别名组 L209-215（微软雅黑/Microsoft YaHei/UI、
                                宋体/SimSun/NSimSun/新宋体、黑体/SimHei、楷体/KaiTi/楷体_GB2312、
                                仿宋/FangSong/仿宋_GB2312、等线/DengXian/等线Light/DengXian Light）
                              · _font_names_equivalent L358
- DataLocator L887           高级审核定位器（anchor/search_in/range 查找；_range_cfg L968）
- PluginManager L1040        数据检查插件（均值偏差/参数锁定/范围检查）
- CheckItemConfig L1333 / CheckRule L1345 / CheckProject L1354   规则数据模型
- OpenpyxlComparer L1361     对比引擎（核心）
  · __init__ L1362（check_options/check_project/stop_event/mode/color_tolerance）
  · run() L1433 → _load_workbooks L1449 → _run_diff_mode L1469
  · shift_scope() L1499      【v3.99 提取】shift 列配对+垂直范围（引擎与测试按钮共用）
  · _apply_rule_filter L1544 → 规则映射/豁免判定（range+shift）/COM 样式复核
  · _compare_worksheet L1807（value/formula 开关门控 + 数组公式区域 L1842 门控）
  · _get_cell_diff L1900（value 全局开关已消费）
  · _compare_row_col_dimensions L1989（行高=默认15/列宽=默认8.43 豁免）
  · 属性：diffs / stats / sheet_diffs / old_wb_ref / new_wb_ref / old_cache / new_cache
- ExcelCOMVerifier L2443     Excel COM 显示层采集器
  · collect_style_data(diffs,...) L2469 写入 diff['com_style']
  · _read_cell_style / _connect_excel / _ensure_workbooks
- CheckOptionsDialog L2697  常规检测设置窗
- ComCheckDialog L2724       高级审核确认窗（Excel激活警告+检查COM通道，测试按钮复用）
- CheckProjectDialog L2794   检查项目集配置窗
- RuleEditorDialog L3121     规则编辑窗（测试按钮由 gui_patch.py 注入；on_ok 组装 ds）
- DiffViewer L3284           GUI 主窗口（start_compare / _finalize_diag / export_diag）
  · jump_to_excel L3889      跳转并选中（支持联合多区，如 B48:C49,E5）

## 差异类型名（判定/断言用）
内容变化 / 公式变化 / 富文本变化 / 字体变化 / 填充变化 / 边框变化 / 对齐变化 /
数字格式变化 / 合并新增 / 合并删除 / 行高变化 / 列宽变化 / 图片新增 / 图片变动 /
图片尺寸变化 / 条件格式新增 / 条件格式删除 / 条件格式修改 / 条件格式变化 /
单元格新增 / 单元格删除 / 高级检查

## 自动化小文件（AI 创建的独立文件）
- cli.py → auto_run.exe：main() 读 scenarios.json → OpenpyxlComparer + ExcelCOMVerifier + ai_summary
- ai_summary.py：write_summary / write_focus / write_package / group_patterns
- regression.py：金标准 14 案例 + L2 全场景（B1单项/B2干扰/B3全局交叉/B4 COM分歧/B5主题专项），
  单文件当前约1154行；断言语义 DIFF/PRESENT=存在即可、SAME=无差OK或有差须豁免、NODIFF=必须无差；
  suite() 汇总全部用例（149个L2）；输出 RULES_VERDICT.txt + L2_REPORT.txt
- build.yml：构建3个exe（main.py 改动需全量重建）；已清 l2_scenarios 相关

## 关键约定
- 值变化类差异的类型名 = 「内容变化」
- 测试机运行前先手动打开 Excel 关激活窗；跑 regression 前还要关闭 Excel 里打开的 l2_* 文件（否则 Permission denied）
- shift 规则：垂直范围必须合法非空，否则引擎跳过；old列+shift_offset=新列
- 模式 ID（P001...）由 ai_summary 动态生成，按数量降序
- 版本号在摘要头部（version.txt，Actions 注入，如 ac36b3c）
- main.py VERSION 常量仍为 v3.98（功能 v3.99），待 FEEDBACK 阶段统一升版
