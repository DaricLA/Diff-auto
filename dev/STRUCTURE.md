# STRUCTURE 主程序结构索引（main.py v4.0，约4332行）

> 用途：给新对话的 AI 定位代码用。先看此索引，需要细节时再粘贴对应代码片段。
> 维护：每次 AI 改动 main.py 后，顺手更新本文件的行号与说明。
> 2026-09-06 校准①：拍板豁免已入 main.py（字体别名内置组 L209、行高/列宽默认值豁免
> _compare_row_col_dimensions L2181+）；诊断块已删；相对 9-04 基准行号已刷新（本文为准）。
> 2026-09-06 校准②~④：UI 优化（心跳/进度平滑/Canvas 自绘/双循环）+ 性能修复 + 诊断清理。
> 诊断导出（_finalize_diag / export_diag / 导出诊断按钮）已废弃删除（2026-09-06 用户拍板）。
> 2026-09-06 下半场（列宽/进度条/趋势/文件名/白底）行号已刷新——4332行版本，本文为准；
> 各"上半场"小节行号同批校准。

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
- _read_formula_cache(path) L829    【新】流式读公式缓存值{sheet:{ref:值}}（不二次加载工作簿，152MB→0.07s）
- _formula_cache_lookup L867         【新】快读 dict / data_only workbook 双形态取值（兼容回退）
- _prog_color(frac,breath) L878     进度条颜色：橙#FF9800↔绿#198754 呼吸插值（breath 0=橙1=绿）
- _sqref_boxes(sq) L885              【新】sqref→坐标区间（纯列/纯行/多区域，绕开 MultiCellRange 慢API）
- _boxes_contains(boxes,address) L907  【新】坐标区间判定（10万次0.075s）
- _mix_hex(c1,c2,k) L919             颜色混合（96→100 末段呼吸色过渡绿）
- _rr(cv,x0,y0,x1,y1,r,**kw) L924   canvas 圆角矩形
- _fmt_duration(secs) L930           耗时格式化
- parse_rich_text_from_xlsx(path) L956  解析富文本
- compare_rich_text_runs(r1, r2) L1002  富文本 run 对比

## 核心类
- WorkbookStyleCache L196    样式缓存：主题色/自定义数字格式/字体等价表 font_equiv
                              · font_equiv 内置别名组 L209（微软雅黑/Microsoft YaHei/UI、
                                宋体/SimSun/NSimSun/新宋体、黑体/SimHei、楷体/KaiTi/楷体_GB2312、
                                仿宋/FangSong/仿宋_GB2312、等线/DengXian/等线Light/DengXian Light）
                              · _font_names_equivalent L358
- DataLocator L1024          高级审核定位器（anchor/search_in/range 查找）
  · _find_anchor L1062       性能：max_row/col 提出循环（openpyxl 全表计算属性热点）
  · _range_cfg L1096（原968，随行号整体偏移）
- PluginManager L1179        数据检查插件（均值偏差/参数锁定/范围检查）
- CheckItemConfig L1498 / CheckRule L1510 / CheckProject L1521   规则数据模型
- OpenpyxlComparer L1526     对比引擎（核心）
  · __init__ L1527（check_options/check_project/stop_event/mode/color_tolerance；progress 回调）
  · _with_heartbeat L1585   心跳（⏳ 已耗时 Ns 原地刷新；pulse 开关控制进度模式）
  · run() L1602 → _load_workbooks L1618 → _run_diff_mode L1643
  · _proc_pct L1638          跨 sheet 全局进度换算（15+10*frac，对比 15→25）
  · shift_scope() L1678      【v3.99 提取】shift 列配对+垂直范围
  · _apply_rule_filter L1723 → 规则映射/豁免判定（range+shift）/COM 样式复核
    · 性能：每200条diff报进度(35→98)；_get_cfs_for_cell 坐标区间缓存
  · _compare_worksheet L1994（value/formula 开关门控 + 数组公式区域门控）
  · _get_cell_diff L2087（value 全局开关已消费）
  · _compare_row_col_dimensions L2176（行高=默认15/列宽=默认8.43 豁免）
  · 属性：diffs / stats / sheet_diffs / old_wb_ref / new_wb_ref / old_cache / new_cache
- ExcelCOMVerifier L2644     Excel COM 显示层采集器
  · collect_style_data L2675 写入 diff['com_style']（progress_fn 由 GUI 传映射回调）
  · _read_cell_style / _connect_excel / _ensure_workbooks
- CheckOptionsDialog L2898  常规检测设置窗
- ComCheckDialog L2925       高级审核确认窗（Excel激活警告+检查COM通道，测试按钮复用）
- CheckProjectDialog L2995   检查项目集配置窗
- RuleEditorDialog L3406     规则编辑窗（测试按钮由 gui_patch.py 注入；on_ok 组装 ds）
- DiffViewer L3569           GUI 主窗口
  · _prog_cv 进度条创建 L3587  Canvas 自绘（高16，圆角；整条同色+呼吸）
  · log L3842                日志（心跳行原地刷新+顶替，无残影）
  · update_progress L3858    进度回调（单调 target=_prog_target 只前进）
  · _prog_tick L3894         30ms插值动画（推进+蠕动+完成迸发计数）
  · _breath_tick L3891       独立呼吸链 sin(真实时间×2π/1.6s)，与进度无关
  · _draw_progress L3906     重绘（槽#e9ecef + 填充三态：呼吸色/96-100过渡/完成绿；
                             完成动画=方案E3 中心迸发：白芯从中心向两端铺满15帧(0.45s)→整条泛白回落18帧(0.54s)）
  · _prog_reset L3942 / _prog_creep L3949      检查开始归零+呼吸链启动；慢阶段蠕动1%/s
  · set_progress_mode L3964  空实现（保留 API）
  · start_compare L4027（规则过滤真实进度+分阶段耗时日志；COM 采集经 _com_map 映射 28→35）
  · on_comparison_finished L4098（停止呼吸 + update_progress(100)）
  · on_tree_select L4234     差异详情着色（规则命中：描述浅灰/规则加粗/→红粗）
  · _insert_detail_line L4259
  · jump_to_excel L4279      跳转并选中（支持联合多区，如 B48:C49,E5）
  · _auto_layout_columns L3647  列宽显式管理：hold 锁定满宽 Sheet 基准(动态吃满剩余,下限340)；窗口一小于
                                hold+320 立即压缩「位置」「类型」列(Sheet 冻结不动)，位置/类型压没后
                                Sheet 才压(下限80)；收起列固定60（Tk 原生布局做不到此顺序）

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
- main.py VERSION = "v4.0"（L14，2026-09-06 升版：列宽/进度条/趋势/文件名/白底大版本）
- 进度条（Canvas 自绘，高16px）：颜色=橙#FF9800↔绿#18BC9C（控件同色）恒定循环呼吸（sin 真实时间 1.6s，
  双循环独立计时，不受进度/事件延迟影响）；仅 96→100 从呼吸色过渡到绿；完成动画=方案E3 中心迸发：
  白芯从中心向两端铺满（15帧0.45s，白绿混合 #18BC9C↔#ffffff step 0.68→0.09）→ 整条泛白回落
  （18帧0.54s ease-in-out，幅度0.62→0）后定格 #18BC9C；呼吸/动画链仅最终完成才停（keep-alive 异常不断链）
- UI 阶段进度映射（按真实耗时比例，接近匀速）：
  加载4→15(13%) / 对比15→25(10%，跨sheet连续) / 插件25 高级26 报告27 对比完成28 /
  COM采集28→35(7%) / 规则过滤35→98(63%，每200条diff回报) / 完成100
- 性能（真实数据验证 152MB+32规则 全流程7.4s，过滤4.5s）：
  _get_cfs_for_cell sqref→坐标区间缓存 / _find_anchor max_row/col提循环 / 蠕动1%/s /
  公式缓存值快读（_read_formula_cache 0.07s，替代 data_only 二次加载 152MB 全量）
- 诊断导出已废弃删除（_finalize_diag / export_diag / 导出诊断按钮，2026-09-06 用户拍板）

## 2026-09-06 下半场（列宽/进度条/趋势配置/白底，均用户拍板）
- 列宽显式管理 _auto_layout_columns（DiffViewer L3569 类内，L3647）：hold 锁定满宽 Sheet 基准，
  窗口小于 hold+320 立即压缩位置/类型（Sheet 冻结）→ 位置/类型压没后 Sheet 才压（下限80）；
  收起列恒定 60；trend_dlg 截图确认按钮固定底部
- 进度条：_prog_cv 高 16px；完成动画=方案E3 中心迸发（白芯从中心铺满15帧0.45s + 泛白回落18帧0.54s，
  _mix_hex 绿↔白插值，flash 计数 40→33）
- 日志分段着色 _insert_summary_line L3826：需人工复核=红粗 log_red_bold；已豁免=黑粗 log_bold
  （不再蓝色）
- 全局白色背景：DiffViewer.__init__ 中 style 配置 TFrame/TLabelframe/TLabel background='#ffffff'
  + root.configure(bg='#ffffff')
- DataTrendConfigDialog（L3230，680x900 白底滚动）：历史数据=锚点+行/列偏移+行数/列数矩形
  （空单元格忽略、数值0正常统计）；管制阈值/规格CPK 每项双模式互斥——行/列偏移(从表格抓取)
  或 直接数值，填行列→数值禁用，填数值→行列禁用，全空=不检查；UCL/LCL 特殊模式同样双模式；α 仍手输
- DataTrendEngine L1324：_collect_hist 矩形取值(空忽略/0计入)返回 (vals,起点)；
  _thr() 统一取值——dict{row_offset,col_offset}读单元格(相对历史锚点)/dict{value:x}/数字/字符串=数值模式；
  旧 mean_range/stddev_range（列表）继续兼容；t/F检验+α、UCL/LCL(σ倍数自动或特殊双模式)
- FileNameCheckConfigDialog（L3089，680x780 白底滚动）：模板检查+分段检查两个独立开关同时生效。
  模板检查=纯格式检查(不读表格)：模板按 _ 分段，空段/*=任意，文字段=必须相等，段数须一致；
  分段检查=起始/结束字符位(1=第1字符含端点)截取比对该段报告单元格值，点解析按 _ 自动拆分并算位置，
  编辑行可选段回填/自定义 start/end/sheet/cell；field_mappings 已废弃(保留空列表兼容)
- FileNameCheckEngine L1216：模板检查(纯格式) + 分段检查(start/end 优先，无则 index 兼容旧配置)；
  template_enabled/segment_enabled 开关(旧配置缺省=True)
- RuleEditorDialog L3406：编辑规则(含高级引擎配置入口)；SpecialReminderConfigDialog L3206

