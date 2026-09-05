# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 仓库与文件（GitHub: Diff-auto, Public）
- main.py        v3.98码/v3.99功能（含 shift_scope、拍板豁免；约3942行，STRUCTURE 已校准）
- gui_main.py    启动入口（import main → gui_patch apply）
- gui_patch.py   测试按钮 v2.2 已验证（引擎同源：offset/intersection/range→locate_all 双侧；
                 shift→shift_scope；跳转复用 jump_to_excel；文件未开→询问打开→ComCheckDialog 确认；
                 主窗口经 _VIEWER 全局登记获取）
- cli.py / ai_summary.py   auto_run.exe（暂未动；FEEDBACK 阶段再改）
- regression.py  单文件（约1154行）：金标准14 + L2 全场景 B1-B5（149 用例）
                 输出 RULES_VERDICT.txt + L2_REPORT.txt；诊断区已删
- build.yml      已清 l2_scenarios 相关；构建3个exe（main.py 改动需全量重建）
- dev/PROGRESS.md / dev/STRUCTURE.md

## 已验证 ✔
- 金标准回归 14/14（稳定）
- L2 全场景 149/149 全绿（2026-09-06：B1 4/4、B2 115/115、B3 10/10、B4 12/12、B5 8/8）
- 测试按钮 v2.2 全流程（引擎计算→COM确认→跳转）
- 规则系统（range/shift）+ 全局开关 + COM 对照框架全部按预期工作

## 当前进度（L2 全绿，进入收尾）
- [x] L2 全部通过：B1-B5 全场景 149/149
- [x] 业务拍板落地（代码已改，待最终验证）：
      · 行高=默认15 / 列宽=默认8.43 → 豁免不报（_compare_row_col_dimensions）
      · 字体别名归一（font_equiv 内置组：微软雅黑/Microsoft YaHei、宋体/SimSun、黑体/SimHei、
        楷体/KaiTi、仿宋/FangSong、等线/DengXian 等）
      · 主题色：按显示色解析判定（theme5≠4472C4 为真差异，COM 证实）→ 不做等价处理
- [ ] regression.py 4 条断言同步（B4-4/B4-7/B4-8/B5-5 → NODIFF）→ 重建 → 测试机跑 → 预期仍 149/149
- [ ] FEEDBACK.txt 统一反馈单 → GOLD.md 预期差异登记（豁免/等价规则登记为正式预期）
- [ ] COM 观察点标签修正（观察行"程序:报差/无差"改为引擎实际结果）
- [ ] main.py VERSION 升 v3.99

## 关键约定（易踩坑）
- 非程序员：代码修改必须给【行号+原文+替换文本】，代码块内无行号（手机无法搜索，已强记忆）
- main.py 改动需用户拍板且零行为变化+回归验证
- 测试机：离线/无权限/无Python/只能跑免安装exe；跑前先手动打开 Excel 关激活弹窗；
         跑 regression 前须关闭 Excel 中打开的 l2_* 文件（否则 Permission denied）
- 回传小文件：AI_PACKAGE.txt/RULES_VERDICT.txt/L2_REPORT.txt；DIAG_FULL.json 留测试机
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格（精确匹配不 strip）
- shift：垂直范围必须合法非空否则跳过；old列+shift_offset=新列
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 换对话交接（下一步就做）
1. 确认 regression.py 4 条断言已改（B4-4/B4-7/B4-8/B5-5 → NODIFF）→ Actions 重建 → 测试机跑
2. 回传 RULES_VERDICT.txt + L2_REPORT.txt（预期 14/14 + 149/149）→ AI 确认后进入 FEEDBACK/GOLD 阶段
3. 贴回 PROGRESS + STRUCTURE + 报告 → AI 分析
