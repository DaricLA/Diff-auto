# STRUCTURE 主程序结构索引（main.py v3.99，约4035行）

> 用途：给新对话的 AI 定位代码用。先看此索引，需要细节时再粘贴对应代码片段。
> 维护：每次 AI 改动 main.py 后，顺手更新本文件的行号与说明。
> 2026-09-06 校准①：拍板豁免已入 main.py（字体别名内置组 L209、行高/列宽默认值豁免
> _compare_row_col_dimensions L2015+）；诊断块已删；相对 9-04 基准行号已刷新（本文为准）。
> 2026-09-06 校准②：UI 优化入 main.py（心跳/进度平滑）；行号整体刷新。
> 2026-09-06 校准③：进度条改 Canvas 自绘（整条同色随进度橙→绿渐变+呼吸，0~100单调平滑）。

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
- _prog_color(frac,breath) L792     【新】进度条颜色：橙#FF9800→绿#198754随进度渐变+呼吸亮度±14%
- _rr(cv,x0,y0,x1,y1,r,**kw) L799   【新】canvas圆角矩形
- _fmt_duration(secs) L804          耗时格式化
- parse_rich_text_from_xlsx(path) L820  解析富文本
- compare_rich_text_runs(r1, r2) L866  富文本 run 对比

## 核心类
- WorkbookStyleCache L196    样式缓存：主题色/自定义数字格式/字体等价表 font_equiv
                              · font_equiv 内置别名组 L209（微软雅黑/Microsoft YaHei/UI、
                                宋体/SimSun/NSimSun/新宋体、黑体/SimHei、楷体/KaiTi/楷体_GB2312、
                                仿宋/FangSong/仿宋_GB2312、等线/DengXian/等线Light/DengXian Light）
                              · _font_names_equivalent L358
- DataLocator L899           高级审核定位器（anchor/search_in/range 查找；_range_cfg L968）
- PluginManager L1052        数据检查插件（均值偏差/参数锁定/范围检查）
- CheckItemConfig L1345 / CheckRule L1357 / CheckProject L1366   规则数据模型
- OpenpyxlComparer L1373     对比引擎（核心）
  · __init__ L1374（check_options/check_project/stop_event/mode/color_tolerance；progress 回调）
  · _with_heartbeat L1432   心跳（⏳ 已耗时 Ns 原地刷新；pulse 开关控制进度模式）
  · run() L1449 → _load_workbooks L1465 → _run_diff_mode L1490（进度：加载5→22/对比25→70/
    插件72/高级74/生成报告76/对比完成78——100 由 GUI 收尾发）
  · _proc_pct L1485          跨 sheet 全局进度换算（逐行进度单调不回退）
  · shift_scope() L1525      【v3.99 提取】shift 列配对+垂直范围
  · _apply_rule_filter L1570 → 规则映射/豁免判定（range+shift）/COM 样式复核
  · _compare_worksheet L1833（value/formula 开关门控 + 数组公式区域门控）
  · _get_cell_diff L1926（value 全局开关已消费）
  · _compare_row_col_dimensions L2015（行高=默认15/列宽=默认8.43 豁免）
  · 属性：diffs / stats / sheet_diffs / old_wb_ref / new_wb_ref / old_cache / new_cache
- ExcelCOMVerifier L2469     Excel COM 显示层采集器
  · collect_style_data L2495 写入 diff['com_style']（progress_fn 由 GUI 传映射回调）
  · _read_cell_style / _connect_excel / _ensure_workbooks
- CheckOptionsDialog L2723  常规检测设置窗
- ComCheckDialog L2750       高级审核确认窗（Excel激活警告+检查COM通道，测试按钮复用）
- CheckProjectDialog L2820   检查项目集配置窗
- RuleEditorDialog L3147     规则编辑窗（测试按钮由 gui_patch.py 注入；on_ok 组装 ds）
- DiffViewer L3310           GUI 主窗口
  · _prog_cv 进度条创建 L3338  Canvas 自绘（高33，圆角；整条同色+呼吸）
  · log L3552                日志（心跳行原地刷新+顶替，无残影）
  · update_progress L3568    进度回调（单调 target=_prog_target 只前进）
  · _prog_start_anim L3575 / _prog_tick L3578   30ms插值动画（推进+呼吸相位+蠕动+完成闪烁8拍）
  · _draw_progress L3599     重绘（槽#e9ecef + 填充颜色：呼吸/闪烁/定格三态），Configure 自动重绘
  · _prog_reset L3611        检查开始归零+启动呼吸
  · _prog_creep L3615        【新】慢阶段时间锚定蠕动（规则过滤 93→98 缓慢推进）
  · _gui_heartbeat L3619     GUI 侧心跳工具（COM采集/规则过滤阶段用）
  · set_progress_mode L3630  空实现（保留 API；不再有模式动画）
  · _finalize_diag L3644 / export_diag L3678
  · start_compare L3738（_prog_reset；COM 采集进度经 _com_map 映射 80→93；规则过滤 93→98 蠕动）
  · on_comparison_finished L3812（停止呼吸 + update_progress(100)）
  · on_tree_select L3948     差异详情着色（规则命中：描述浅灰/规则加粗/→红粗）
  · _insert_detail_line L3973
  · jump_to_excel L3993      跳转并选中（支持联合多区，如 B48:C49,E5）

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
- main.py VERSION = "v3.99"（L14，已确认，2026-09-06）
- 进度条（Canvas 自绘）：整条颜色=橙#FF9800↔绿#198754 循环呼吸（相位 0→1→0 往返，周期约1.6s）
  （进度语义由条长度表达）；完成时光晕脉冲提醒（内部稳定绿 + 三色发光圈 #d6fbe9→#a8fbdc→#7cffc9
  扩散2次，0.6s/脉冲共1.2s）后定格 #198754；呼吸仅最终完成才停；进度单调 0→100
- UI 阶段进度映射（按耗时权重）：加载5→22 / 对比25→70（跨sheet连续，45%） / 插件72 /
  高级检查74 / 生成报告76 / 对比完成78 / COM采集80→93（13%，每sheet更新） /
  规则过滤93→98（蠕动+5%，速度近似0.15%/s） / 完成100
