# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 里程碑（2026-09-06 全部达成 ✅）
✅ 金标准回归 14/14
✅ L2 全场景 149/149 全绿（B1 4/4、B2 115/115、B3 10/10、B4 12/12、B5 8/8）
✅ 业务拍板全部落地并验证：行高=默认15/列宽=默认8.43豁免、字体别名归一、主题色按显示色判定
✅ 主要修复：CheckProject参数错位、value开关消费、数组公式门控、B4/B5工作簿混用、
   主题色索引(accent2)、断言语义体系(DIFF=存在/SAME=豁免/NODIFF=无差)、图片tempfile、诊断已删

## 仓库与文件（GitHub: Diff-auto, Public）
- main.py        v3.98码/v3.99功能（3942行；STRUCTURE 已校准；VERSION 待升 v3.99）
- gui_main.py / gui_patch.py   测试按钮 v2.2 已验证（未动）
- cli.py / ai_summary.py       auto_run.exe（FEEDBACK 阶段再动）
- regression.py  单文件（约1157行）：金标准14 + L2 149用例，输出 RULES_VERDICT.txt + L2_REPORT.txt
- build.yml      已清 l2_scenarios；main.py 改动需全量重建3个exe
- dev/PROGRESS.md / dev/STRUCTURE.md

## 下一步（全部完成，进入日常使用）
✅ VERSION 升 v3.99（已验证，报告显示版本 v3.99）
✅ COM 观察标签改为引擎实际结果（已验证：行高15/列宽8.43 显示"程序: 无差"）
✅ 开发闭环完成：引擎+测试+规则+豁免+GOLD/FEEDBACK 流程全部就绪
- 日常使用：业务文件跑 auto_run/GUI；发现问题在 FEEDBACK.txt 填一行 → AI 自动分析

## 2026-09-06 第二批：UI/进度/日志优化（已改 main.py，待回归验证）
- 差异详情着色：规则命中时 描述=浅灰、规则名=加粗、结果行 → 红粗（on_tree_select L3910）
- 进度条：①全局单调防回退（update_progress _prog_last）；②跨 sheet 连续进度（_proc_pct）；
  ③indeterminate 改伪动画大块 15%~70% 循环（原生块宽无法配置，块宽>150px）；
  ④流程全覆盖：加载5-20 / 对比25-80 / 插件85 / 高级检查87（新增）/ 规则过滤92 / 报告95 / 100
- ⏳ 心跳补齐：逐sheet对比、插件、高级检查、COM采集、规则过滤（引擎 _with_heartbeat pulse 参数；
  GUI _gui_heartbeat）；心跳行原地刷新并被完成行顶替（log 残影 bug 已修）
- COM 阶段改动画+心跳，不再把进度从100拉回0；verifier 不传 progress 回调
- 判定逻辑零改动（纯 UI/进度/日志层）；STRUCTURE 行号已刷新（main.py 4009行）
- 本机验证（Python 3.13.15）：py_compile ✓ / import main ✓ / 引擎全链路冒烟（进度单调 5→100）✓
  / _with_heartbeat pulse 分支 ✓ / _proc_pct 全局进度 ✓ / update_progress 单调 ✓ / 伪动画循环与恢复 ✓ / log 心跳顶替无残留 ✓
- 修复2个bug：① _prog_last 跨次检查不重置导致第二次进度卡死 → start_compare 重置归零；
  ② 伪动画起始值 clamp（避免从满格直接跳回15）→ 按当前进度值起步（15~70范围）
- ✅ 回归验证通过（2026-09-06 测试机）：金标准 14/14 + L2 149/149 全绿（B1 4/4、B2 115/115、
  B3 10/10、B4 12/12、B5 8/8），无硬断言失败；COM 观察点均为已知观察项
  （tint/主题色内部色号差异等，判定不受影响）；行高15/列宽8.43 显示"程序: 无差"（豁免生效）
- 状态：UI/进度/日志优化已完成并入库，进入日常使用

## 2026-09-06 第三批：进度条 Canvas 自绘重做（已改 main.py，待回归+GUI实机看效果）
- 用户拍板：方案D 整条同色 橙#FF9800→绿#198754 随进度渐变；呼吸=亮度±14%(amp0.28,周期1.58s)；
  高度33px(与原ttkbootstrap一致)；纯色无条纹；完成时呼吸停止固化纯绿
- 技术验证：ttkbootstrap flatly 主题下 style.configure(background) 不生效（实验证明）→
  改为 Canvas 自绘（_prog_color L792 / _rr L799 / _prog_cv L3338 / _prog_tick L3578）
- 0~100单调平滑：update_progress 仅单调 target；30ms插值动画消除阶梯；删伪动画/_pulse_*
- 全流程进度映射（不回退）：加载5→20/对比25→80/插件85/高级87/报告95/对比完成96/
  COM采集97→99(verifier进度经_com_map映射)/规则过滤99/完成100（GUI收尾发）
- 本机验证：py_compile ✓ / _prog_color 色值断言 ✓ / Tk真实窗口动画测试（单调、平滑、呼吸
  启停、_draw_progress）✓ / 引擎冒烟：进度 5→15→20→52→80→95→96 单调 ✓
- 判定逻辑零改动；regression.py 内嵌副本不受影响；STRUCTURE 行号已刷新（4035行）

## 2026-09-06 第四批：进度分配重排 + 呼吸全程不断（已改 main.py，待回归+实机确认）
- 用户反馈：①慢阶段%占比太少（COM 只2%、过滤1%）→ 按耗时权重重排映射：
  加载5→22 / 对比25→70(45%) / 插件72 / 高级74 / 报告76 / 对比完成78 /
  COM采集80→93(13%，每sheet更新) / 规则过滤93→98(蠕动0.15%/s) / 完成100
- 用户反馈：②进度卡住时呼吸不能停 → _prog_tick 循环条件改为 breathe 常真（breathe=True 恒循环），
  仅 on_comparison_finished（最终完成）才停；新增 _prog_creep 慢阶段时间锚定蠕动
- 本机验证：py_compile ✓ / 引擎冒烟 5→15→22→47→70→76→78 单调 ✓ /
  Tk 实测：进度停住 anim 仍运行 ✓ creep 推进不超上限 ✓ 完成才停 ✓ 卡住时呼吸颜色持续变化 ✓
- 判定逻辑零改动；STRUCTURE 行号已刷新（4047行）

## 2026-09-06 第五批：呼吸改版（橙↔绿循环）+ 完成闪烁提醒（已实现，待回归+实机确认）
- 用户拍板：①未满进度时颜色改为 橙#FF9800↔绿#198754 全跨度循环呼吸（相位0→1→0往返约1.6s，
  明显可见；进度语义由条长度表达）；②到达终点高亮闪烁提醒=方案1脉冲闪烁（亮#7cffc9↔暗#0d5a3c
  0.3s/拍×4次共2.4s，之后定格#198754）——原闪烁色与基绿对比太小不可见，已优化参数
- 实现：_prog_color 改为按 breath 插值（frac 不再影响颜色）；_prog_tick 加入闪烁序列
  （_prog_flash 8拍 + _prog_flash_wait 节拍器）；_draw_progress 三态着色
- 本机验证：py_compile ✓ / 呼吸端点断言（0=#ff9800、1=#198754）✓ / 相位往返循环 ✓ /
  闪烁8拍后动画停止 ✓ / 定格色 #198754（canvas fills 实测）✓
- 判定逻辑零改动；STRUCTURE 已同步

## 2026-09-06 第六批：完成提醒改为方案B 光晕脉冲（已实现，待回归+实机确认）
- 用户拍板：完成提醒=方案B（macOS 风光晕脉冲）：内部稳定绿，外围三色发光圈
  （#d6fbe9→#a8fbdc→#7cffc9）扩散 2 次（0.6s/脉冲，30ms/帧共40帧=1.2s）后定格 #198754
- 实现要点：canvas polygon 不支持 width 描边（TclError 被吞导致早期不可见）→ 改用 3 层
  填充色内缩模拟渐隐扩散；_prog_tick 光晕计数 _prog_flash 0..40
- 本机验证：py_compile ✓ / 中间帧 fills=[浅3层+槽+绿] ✓ / 40帧后动画停 ✓ / 定格无光晕 ✓
- 判定逻辑零改动；STRUCTURE 已同步

## 关键约定（易踩坑）
- 非程序员：代码修改必须给【行号+原文+替换文本】，代码块内无行号（手机无法搜索）
- main.py 改动需用户拍板且零行为变化+回归验证
- 测试机：离线/无权限/无Python；跑前先手动打开 Excel 关激活弹窗；
         跑 regression 前须关闭 Excel 中打开的 l2_* 文件（否则 Permission denied）
- 回传小文件：AI_PACKAGE.txt / RULES_VERDICT.txt / L2_REPORT.txt；DIAG_FULL.json 留测试机
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格
- shift：垂直范围必须合法非空否则跳过；old列+shift_offset=新列
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 换对话交接（下一步就做）
1. 用户确认开始 FEEDBACK.txt / GOLD.md 阶段 → AI 起草模板 → 用户填写/确认 → 入库
2. 贴回 PROGRESS + STRUCTURE（均为最新）
