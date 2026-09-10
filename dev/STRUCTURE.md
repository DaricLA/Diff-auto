# STRUCTURE 主程序结构索引（main.py v4.4，5054行）

> 用途：给新对话的 AI 定位代码用。先看此索引，需要细节时再粘贴对应代码片段。
> 维护：每次 AI 改动 main.py 后，顺手更新本文件的行号与说明。
> 历史校准：2026-09-06 校准①~④（豁免内置组 L223、行高/列宽豁免 L2371、UI优化/性能修复/诊断清理）；
> 2026-09-08 v4.2（引擎分层 + 四项修复 + 三处补丁）；2026-09-10 v4.3（B版功能合并：统一滚轮 + 文件名 v5；
> 三补丁回补；清理多余空行后全文件行号刷新——以本文为准）；
> 2026-09-10 v4.3.1（文件名引擎 _src_path 静默失效修复）、v4.3.2（配置确定后子引擎开关自动打开）、
> v4.4（文件名配置 UI 改造：位标尺/详情条/保存前清单；FileNameCheckConfigDialog 重写）
> ⚠️ v4.4 起 FileNameCheckConfigDialog 之后的所有行号 = 旧行号 + 384~388（全文件 4669→5054 行），以本文为准。

## 模块级函数（main.py）
- format_numfmt_readable(fmt) L68   数字格式码→中文可读描述
- excel_tint(hex6, tint) L74        ECMA-376 主题色 tint 变换（RGB→HLS→RGB）
- rgb_channel_close(h1, h2) L88     颜色逐通道差<=1（tint舍入噪声容忍）
- rgb_euclidean(h1, h2) L97         RGB欧氏距离（调色板吸附容差，<=60 视为同色）
- _grid_segments(cell_set) L109     4连通单元格集合→矩形块（相邻差异合并用）
- dedup_rich_text_diffs(diffs) L136 富文本重复差异去重
- merge_adjacent_diffs(diffs) L140  相邻同类差异合并
- color_signature(color) L162       颜色存储签名比较
- center_window(win, parent=None) L172  窗口居中（父窗口几何不可用时回退屏幕居中）
- _chat_wheel(dialog, event) L187   【v4.3】统一滚轮：只滚光标所在对话框内带 _chat_scroll_cv 标记的 Canvas
- cell_address(col,row) L758        列号行号→地址字符串
- normalize_path(p) L759            路径归一化
- get_sheet_names_fast(path) L762   快速读 sheet 名（流式）
- formula_text(value) L791          公式解析为文本
- normalize_formula(f) L801         公式归一化
- _read_formula_cache(path) L843    流式读公式缓存值{sheet:{ref:值}}（不二次加载工作簿，152MB→0.07s）
- _formula_cache_lookup L881        快读 dict / data_only workbook 双形态取值（兼容回退）
- _prog_color(frac,breath) L892     进度条颜色：橙#FF9800↔绿#18BC9C 呼吸插值（breath 0=橙1=绿）
- _sqref_boxes(sq) L899              sqref→坐标区间（纯列/纯行/多区域，绕开 MultiCellRange 慢API）
- _boxes_contains(boxes,address) L921 坐标区间判定（10万次0.075s）
- _mix_hex(c1,c2,k) L933            颜色混合（96→100 末段呼吸色过渡绿、完成迸发用）
- _rr(cv,x0,y0,x1,y1,r,**kw) L938   canvas 圆角矩形
- _fmt_duration(secs) L943          耗时格式化
- parse_rich_text_from_xlsx(path) L971  解析富文本
- compare_rich_text_runs(r1, r2) L1017  富文本 run 对比
validate_filename_rule(rule,cur,tpl) L3299  【v4.4】文件名规则校验纯函数（可单测）
                              返回 (level,msg)：ok/info（配置有效但当前样本会报警）/warn（配置可疑）/error；
                              文件名侧复用 FileNameCheckEngine._parse_date_seg；单元格侧只校验"有没有填"，不打开文件

## 核心类
- WorkbookStyleCache L210    样式缓存：主题色/自定义数字格式/字体等价表 font_equiv
                              · font_equiv 内置别名组 L223（微软雅黑/Microsoft YaHei/UI、
                                宋体/SimSun/NSimSun/新宋体、黑体/SimHei、楷体/KaiTi/楷体_GB2312、
                                仿宋/FangSong/仿宋_GB2312、等线/DengXian/等线Light/DengXian Light）
                              · _font_names_equivalent L372
- DataLocator L1038          高级审核定位器（anchor/search_in/range 查找）
  · _find_anchor L1076       性能：max_row/col 提出循环（openpyxl 全表计算属性热点）
  · _range_cfg L1121
- PluginManager L1193        数据检查插件（均值偏差/参数锁定/范围检查）
- 高级引擎（AdvancedEngine L1225 派生；注册表 ADVANCED_ENGINE_REGISTRY L1555）
  · FileNameCheckEngine L1230  【v4.3 文件名v5】rules[] 逐条执行：文本=严格相等 / 日期=拆解对比
                               （yyMMdd→固定20xx、yyyyMMdd；单元格 datetime/序列号/7种文字格式）；
                               规则级容错（起止非法、缺Sheet/单元格→跳过+日志）；告警 advanced_check=True
    · _parse_date_seg L1302 / _cell_to_date L1312
  · SpecialReminderEngine L1334  特殊提醒（trigger: always/nonempty/gt/lt；shift 防御性跳过）
  · DataTrendEngine L1383      数据趋势
    · _collect_hist L1422      矩形取值（空忽略/0计入，返回 (vals,起点)）
    · _thr L1446               阈值取值：行列抓取/数值/旧格式兼容；t/F检验+α；UCL/LCL σ倍数或特殊双模式
- CheckItemConfig L1557 / AdvancedEngineConfig L1563 / CheckRule L1569 / CheckProject L1578
                              规则数据模型（AdvancedEngineConfig = 高级引擎配置）
- OpenpyxlComparer L1585      对比引擎（核心）
  · __init__ L1586（check_options/check_project/stop_event/mode/color_tolerance；progress 回调）
  · _with_heartbeat L1644    心跳（⏳ 已耗时 Ns 原地刷新；pulse 开关控制进度模式）
  · run() L1661 → _load_workbooks L1677 → _run_diff_mode L1702（→ _run_advanced_engines L1717）
  · _proc_pct L1697           跨 sheet 全局进度换算（15+10*frac，对比 15→25）
  · shift_scope() L1737       shift 列配对+垂直范围
  · _build_waiver_map L1782   【数据限界豁免】每规则预处理阈值（数值/从表格抓取；锚点=规则数据源 anchor）
  · _waiver_lim L1814 / _waiver_judge L1836 / _waiver_fail_desc L1853
                              阈值取值&判定&未豁免描述（【v4.3 回补】阈值无效/全空/非数值→未豁免）
  · _reminder_cfg L1863 / _reminder_hit L1870  提醒配置/判定（工具方法；规则分支已不参与判定）
  · _apply_rule_filter L1880 → 规则映射/豁免判定（范围+shift）/COM 样式复核
    · 入口跳过高级告警（advanced_check 不参与规则匹配与豁免）；性能：每200条diff报进度(35→98)、
      _get_cfs_for_cell 坐标区间缓存
    · 【v4.3】豁免分支：命中 limit_waiver 的规则先判豁免——限内/无条件→记入 pass 并跳过常规检查；
      超限/抓取失败/两侧全空/非数值→直接未豁免（防漏）；提醒不参与判定
    · 规则执行逐条日志：跳過=红、其餘=绿
  · _format_pair L2088 【v4.2】带 sheet 参数
  · _compare_worksheet L2189（value/formula 开关门控 + 数组公式区域门控）
  · _get_cell_diff L2282（value 全局开关已消费）
  · _compare_row_col_dimensions L2371（行高=默认15/列宽=默认8.43 豁免）
  · _resolve_cell_value L2682 【v4.2】shift/描述统一取缓存计算值（公式格不显公式原文）
  · 属性：diffs / stats / sheet_diffs / old_wb_ref / new_wb_ref / old_cache / new_cache
- ExcelCOMVerifier L2843     Excel COM 显示层采集器
  · collect_style_data L2869 写入 diff['com_style']（progress_fn 由 GUI 传映射回调）
  · _read_cell_style L2942 / _connect_excel / _ensure_workbooks
- CheckOptionsDialog L3097  常规检测设置窗
- ComCheckDialog L3124       高级审核确认窗（Excel激活警告+检查COM通道）
- ProjectListDialog L3161    配置列表窗（加载检查项目）
- CheckProjectDialog L3194   检查项目集配置窗
- FileNameCheckConfigDialog L3353  【v4.4 重写】1120x870：文件名区 grid 三列（标签等长）
                              + **位标尺**（Canvas：当前(N)/模板(N) 两行+位号行，蓝框圈当前规则区间，
                                两行同时圈、文字真实 bbox、tag_lower 垫底、80位/行按窗宽自适应、
                                行距取字体实测 linespace 兼容高DPI）
                              + **详情条**（_refresh_detail L3611：第x~y位 │ 取到内容 │ 结论，红/绿）
                              + 规则表（起始/结束/目标/模式/Sheet/单元格；列宽已加宽）
                              + 保存前可疑清单（_suspects L3786 / _confirm_suspects L3809）
                              + 模式目标联动（_on_target_pick/_on_mode_pick：冲突时自动让路+详情条说明）
                              + _draw_ruler_inner L3641 / _seg_len L3626
                              旧「模板/分段双开关」与 field_mappings 仍废弃（旧配置不兼容，需重配）
- SpecialReminderConfigDialog L3837  特殊提醒配置窗
- DataTrendConfigDialog L3861  数据趋势配置窗（680x900 白底滚动；阈值双模式互斥）
- LimitWaiverConfigDialog L4039 数据限界豁免配置窗（无滚动自适应高度+按钮固定底部；无条件豁免冻结输入；
                              UCL/LCL 行列抓取或数值 双模式互斥）
- RuleEditorDialog L4118     规则编辑窗（测试按钮由 gui_patch.py 注入；on_ok L4260 组装 ds；
                              _open_engine_config L4176；高级检查区含 文件名一致性/特殊提醒/数据趋势/数据限界豁免）
- DiffViewer L4286           GUI 主窗口（以下类内行号 = 旧行号+388）
  · _prog_cv 进度条创建 L4322  Canvas 自绘（高16，圆角；整条同色+呼吸）
  · _auto_layout_columns L4364  列宽显式管理（类内）：hold 锁定满宽 Sheet 基准(动态吃满剩余,下限340)；
                                窗口小于 hold+320 立即压缩「位置」「类型」列(Sheet 冻结)，压没后 Sheet 才压(下限80)
  · _log_tag L4535            日志着色：异常/识别=红、成功=绿、耗时节点=黑粗（「失败 0 条」不算失败）
  · _insert_summary_line L4548 统计行分段着色（需人工复核=红粗 log_red_bold；已豁免=黑粗 log_bold）
  · log L4564                 日志（心跳行原地刷新+顶替，无残影）
  · update_progress L4580     进度回调（单调 target=_prog_target 只前进）
  · _breath_tick L4594        独立呼吸链 sin(真实时间×2π/1.6s)，与进度无关
  · _prog_tick L4601          30ms插值动画（推进+蠕动+完成迸发计数）
  · _draw_progress L4628      重绘（槽#e9ecef + 填充三态：呼吸色/96-100过渡/完成绿；
                              完成动画=方案E3 中心迸发：白芯15帧→泛白回落18帧→定格#18BC9C）
  · _prog_reset L4664 / _prog_creep L4671      检查开始归零+呼吸链启动；慢阶段蠕动1%/s
  · set_progress_mode L4686   空实现（保留 API）
  · start_compare L4749（规则过滤真实进度+分阶段耗时日志；COM 采集经 _com_map 映射 28→35；
                              L4792 调用 _apply_rule_filter）
  · on_comparison_finished L4820（停止呼吸 + update_progress(100)）
  · populate_tree L4841 / _populate_tree_inner L4862   差异树构建（需人工复核/已豁免/数据检查分组）
  · on_tree_select L4956      差异详情着色（规则命中：描述浅灰/规则加粗/箭头红粗；高级检查提醒行）
  · _insert_detail_line L4983
  · jump_to_excel L5003       跳转并选中（支持联合多区，如 B48:C49,E5）
  · stop_compare L5050        停止检查

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
  （待补：豁免限内/超限/无条件、shift 缓存值、高级告警不参与豁免 → 14→20+）
- build.yml：构建3个exe（main.py 改动需全量重建）；已清 l2_scenarios 相关

## _selftest/ 测试台（v4.4 新建，在开发机本机跑，不参与 exe 构建）
- 环境：开发机 Python 3.13.14（F:\Program\PYTHON\python.exe）+ openpyxl/ttkbootstrap/pywin32/lxml，
  Excel COM 可用（Excel 12.0）；全部在 `excel diff2\_selftest\` 下
- `run_tests.py`            引擎级自测 14 条（基础值变化/特殊提醒/豁免限内/豁免超限/文件名链路）→ 14/14
- `test_filename.py`        文件名全场景夹具矩阵 32 条（文本/日期7种格式/序列号/容错/组合）→ 32/32
- `test_filename_validate.py` 校验纯函数 + 对话框冒烟 42 条（位标尺/蓝框/对齐/两行不重叠/80位/模式联动）
                            支持 `DIFF_DIALOG_TARGET=main` 直接验已并入 main.py 的正式版 → 42/42
- `test_engine_toggle.py`   子引擎开关联动 10 条（配置确定→开关自动打开；取消→不动；on_ok 落盘）→ 10/10
- `proto_dialog_live.py` + `运行原型.bat`   改造原型（已并入 main.py，保留作对照）
- `_merge_v44.py`           原型→main.py 拼接脚本（带断言）；`symbol_lines.py` 印符号当前行号；
  `show_project_json.py`     检查项目 JSON 样例；`main_backup_v4.3.2.py` 合并前备份

## 关键约定
- 值变化类差异的类型名 = 「内容变化」
- 测试机运行前先手动打开 Excel 关激活窗；跑 regression 前还要关闭 Excel 里打开的 l2_* 文件（否则 Permission denied）
- shift 规则：垂直范围必须合法非空，否则引擎跳过；old列+shift_offset=新列
- 模式 ID（P001...）由 ai_summary 动态生成，按数量降序
- 版本号在摘要头部（version.txt，Actions 注入，如 ac36b3c）
- main.py VERSION = "v4.4"（L14，2026-09-10：文件名配置 UI 改造；v4.3.2=配置确定后开关自动打开；
  v4.3.1=文件名引擎 _src_path 修复；v4.3=B版功能合并+三补丁回补；
  v4.2=引擎分层四修复；v4.1=数据限界豁免；v4.0=列宽/进度条/趋势/文件名/白底大版本）
- 【v4.3】鼠标滚轮 _chat_wheel L187：只滚光标所在对话框内带 _chat_scroll_cv 标记的 Canvas
  （文件名/数据趋势/规则编辑器三窗接入；旧 _ds_wheel 方案已删）
- 【v4.3】文件名一致性 v5：规则列表 rules[]（文本=严格相等；日期=yyMMdd→固定20xx、yyyyMMdd）；
  旧配置（template_enabled/segments）不兼容，需重配
- ⚠️ 版本防混淆：两份 4561 行旧档辨识法——搜 `_waiver_fail_desc`（旧稿 0 处 / 定稿 3 处）
- 进度条（Canvas 自绘，高16px）：颜色=橙#FF9800↔绿#18BC9C 恒定循环呼吸（sin 真实时间 1.6s，
  双循环独立计时，不受进度/事件延迟影响）；仅 96→100 从呼吸色过渡到绿；完成动画=方案E3 中心迸发：
  白芯从中心向两端铺满（15帧0.45s）→ 整条泛白回落（18帧0.54s）后定格 #18BC9C；
  呼吸/动画链仅最终完成才停（keep-alive 异常不断链）
- UI 阶段进度映射（按真实耗时比例，接近匀速）：
  加载4→15(13%) / 对比15→25(10%，跨sheet连续) / 插件25 高级26 报告27 对比完成28 /
  COM采集28→35(7%) / 规则过滤35→98(63%，每200条diff回报) / 完成100
- 性能（真实数据验证 152MB+32规则）：_get_cfs_for_cell 坐标区间缓存 / _find_anchor max_row/col提循环 /
  蠕动1%/s / 公式缓存值快读（_read_formula_cache 0.07s，替代 data_only 二次加载）
- 诊断导出已废弃删除（_finalize_diag / export_diag / 导出诊断按钮，2026-09-06 用户拍板）

## 2026-09-06 下半场（列宽/进度条/趋势配置/白底，均用户拍板）
- 列宽显式管理 _auto_layout_columns（DiffViewer L3898 类内，L3976）：hold 锁定满宽 Sheet 基准，
  窗口小于 hold+320 立即压缩位置/类型（Sheet 冻结）→ 位置/类型压没后 Sheet 才压（下限80）；
  收起列恒定 60
- 进度条：_prog_cv 高 16px；完成动画=方案E3 中心迸发（_mix_hex 绿↔白插值，flash 计数 33）
- 日志分段着色 _insert_summary_line L4160：需人工复核=红粗 log_red_bold；已豁免=黑粗 log_bold
- 全局白色背景：DiffViewer.__init__ style 配置 TFrame/TLabelframe/TLabel background='#ffffff' + root 白
- DataTrendConfigDialog L3477（680x900 白底滚动）：历史数据=锚点+行/列偏移+行数/列数矩形
  （空单元格忽略、数值0正常统计）；管制阈值/规格CPK 每项双模式互斥——行/列偏移(从表格抓取) 或 直接数值，
  填行列→数值禁用，填数值→行列禁用，全空=不检查；UCL/LCL 特殊模式同样双模式；α 仍手输
- FileNameCheckConfigDialog L3288（720x780）：【v4.3 升级为 v5，见上方「核心类」】
- SpecialReminderConfigDialog L3453；RuleEditorDialog L3734
- LimitWaiverConfigDialog L3655【数据限界豁免】：无条件豁免开关（开启冻结全部输入→保存时 ucl/lcl 清空）
  + UCL/LCL 双模式互斥（行/列=相对规则数据源锚点抓取 或 直接数值），一侧全空只查另一侧，两侧全空不做范围豁免；
  窗口无滚动自适应高度，按钮 pack 底部必然可见；测试按钮由 gui_patch.py 注入

## 2026-09-08 v4.2 / 2026-09-10 v4.3（引擎分层 + B版合并）
- 引擎分层模型（用户拍板）：提醒型（特殊提醒/文件名/趋势）=只产告警diff（橙色斜体、永远「需人工复核」，
  不参与规则豁免）；豁免型（数据限界豁免）=唯一参与判定；基础判定=规则检查项；多规则聚合=AND
- 四項修复（v4.2）：①高级告警跳过规则匹配（_apply_rule_filter L1880 内 `if d.get('advanced_check'): continue`）
  ②shift/描述统一写缓存计算值（_resolve_cell_value L2682）③规则逐条日志（跳過=红、其餘=绿）
  ④_log_tag L4535 着色语义（异常/识别=红、成功=绿、耗时节点=黑粗）
- 三处补丁（v4.3 回补）：_waiver_fail_desc L1853；_apply_rule_filter 两分支 else
  「超限/抓取失败/两侧全空/非数值→直接未豁免」（防漏）；提醒分支已删除（提醒职责分离）
  自检：`_waiver_fail_desc` 3处、`→ 需人工确认` 0处
- B版合并（v4.3）：①统一滚轮 _chat_wheel L187 ②文件名一致性 v5（对话框 L3288 + 引擎 L1230；
  ⚠️ 旧配置不兼容需重配）③【从旧版读取】按钮（对话框内）
- 【v4.1 数据限界豁免 引擎侧】豁免不进 ADVANCED_ENGINE_REGISTRY（_run_advanced_engines L1717 自动跳过，
  不产生高级检查告警）；判定在 _apply_rule_filter（L1880）过滤期执行：_build_waiver_map（L1782）每规则
  预处理阈值（数值/抓取，锚点=规则数据源 anchor；抓取失败/锚点缺失→该侧视为不满足→不豁免），
  命中规则数据源的 diff 先判豁免：限内/无条件→记 pass 并跳过该规则常规检查（多规则 AND）；
  豁免描述如「数据限界豁免: 值 52 在管制限内（≥ LCL 40 且 ≤ UCL 60）」；
  豁免引擎在 shift 模式禁止（UI 拦截，与数据趋势/特殊提醒一致）

## 2026-09-10 v4.3.1 / v4.3.2 / v4.4（文件名引擎修复 + 开关联动 + 文件名配置 UI 改造）
### v4.3.1 修复（真 bug，已加断言）
- 现象：文件名一致性检查在真实流程中**静默失效**（永远 0 告警），日志仍打「[高级检查] filename_check 完成」
- 根因：FileNameCheckEngine（L1239）从 `new_wb._src_path` 或 `new_wb.filename` 取文件名，但 `_src_path`
  全项目只在 L1239 被**读**、从不写入；openpyxl Workbook 也没有 `filename` 属性 → fname 恒为空 → 直接 return []
- 修复（_load_workbooks，写在 L1681 同一行末尾，避免全文件行号位移）：
  `; old_wb._src_path=self.old_path; new_wb._src_path=self.new_path`

### v4.3.2 行为变更（用户拍板）
- `RuleEditorDialog._open_engine_config`（L4176）：配置对话框点「确定」后，该子引擎开关若为关则**自动打开**
  （防止「配了但开关没开 → on_ok L4260 静默丢弃」）；点「取消」不改动开关与配置
- 四个子引擎统一：filename_check / special_reminder / data_trend / limit_waiver
- 断言：`_selftest/test_engine_toggle.py` 10/10（含 on_ok 落盘验证）

### v4.4 文件名配置 UI 改造（FileNameCheckConfigDialog L3353 重写；引擎判定逻辑零改动）
- 窗口 720x780 → **1120x870**（高 DPI 下 80 位/行的位标尺需 ~1040px）；文件名区改 **grid 三列**，
  标签等长「当前文件名: / 模板文件名:」不再错位；两个按钮同宽；规则表列宽加宽
- **位标尺**（Canvas 手绘，_draw_ruler L3632 / _draw_ruler_inner L3641 / _seg_len L3626）：
  · 同时显示「位号 / 当前(N) / 模板(N)」三行；蓝框圈当前编辑/选中规则的区间（**两行同时圈**）
  · 蓝框用文字**真实 bbox** 画并 tag_lower 垫底（完整包住字高，不遮字、不再偏上）
  · 每行默认 80 位、按窗宽自适应；**行距取字体实测 linespace**（兼容高 DPI，不写死像素）
  · 已删「其他规则灰线」；点表格行即可预览任意规则；Canvas 加 `<Configure>` 重绘（带 busy 保护）
- **详情条**（表格下方，_refresh_detail L3611）：实时显示「第x~y位（n 位）│ 取到内容 │ 结论」，红/绿着色；
  取代早期「给表格加两列」方案（列宽会横向滚动）
- **保存前可疑清单**（_suspects L3786 / _confirm_suspects L3809）：点确定时列出
  ①日期段不可识别 ②结束位超出文件名长度 ③模板名为空 ④模板已填但规则表为空；
  可「返回修改」或「仍要保存」（不阻断）；**「模版段不一致」属检查正常产出，不计入清单**
  → 为此给校验函数加了 `info` 级别（有效但当前样本会报警，仅详情条显示）
- **模式/目标联动**：日期模式选「模版」→ 模式自动回「文本」（不再静默把目标弹回）；选「日期」时目标自动
  回「待测文件」；两种自动切换均在详情条尾部给出说明文字
- **校验纯函数** `validate_filename_rule`（L3299，模块级）：文件名侧复用引擎 `_parse_date_seg`；
  单元格侧只校验"有没有填"，**不打开文件**（避免 152MB 卡住配置界面）；模版目标严格逐字符比较（与引擎一致）
- 坑位备忘（已修并加断言）：
  · trace 回调会传 (name,index,mode) 三参 → 必须用 `lambda *a:` 包，否则名/表框打字时标尺不刷新
  · 高 DPI（本机 Tk scaling 2.0，等宽字符 12px、行高 26px）：像素不能写死，须用 `font.metrics('linespace')`
- 断言：`_selftest/test_filename_validate.py` 42/42（含 `DIFF_DIALOG_TARGET=main` 直接验正式版）
