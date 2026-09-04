# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 仓库与文件（GitHub: Diff-auto，Public）
- main.py       原程序 v3.98（GUI；OpenpyxlComparer/ExcelCOMVerifier/CheckProject/DataLocator/RuleEditorDialog）→ excel_diff_gui.exe，勿动
- gui_main.py   GUI 启动入口（加载 main + gui_patch 补丁），编译入口已改为它
- gui_patch.py  【测试按钮补丁】规则编辑窗数据源区测试按钮；offset/intersection/range/shift 计算+跳转；排除区间增强；新旧双跳；多格联合选中（最新版待替换）
- cli.py        auto_run.exe：读 scenarios.json，批量对比+COM采集+摘要（待加 FEEDBACK）
- ai_summary.py 摘要生成：模式归并/类型概览/COM判定/focus_pattern（待加 FEEDBACK）
- regression.py regression.exe：14案例金标准回归（待加 shift/range/豁免案例）
- scenarios.json 运行清单（测试机记事本编辑）
- .github/workflows/build.yml  增量打包；gui 过滤名单含 gui_main.py/gui_patch.py
- requirements.txt / dev/PROGRESS.md / dev/STRUCTURE.md

## 已完成
- [x] Actions 增量构建成功；测试机 3 exe 全链路跑通
- [x] 回归 14/14；COM 采集 12/12
- [x] 第一次真实数据诊断：4278 差异≈真实批量更新（Oven+6天已确认=预期差异；Summary/Raw Data 真实）
- [x] 测试按钮补丁已实现；四模式基本跳转正常；已修 Sheet 尾部空格被 strip 的 bug
- [x] 四点增强已确认并出码：shift新旧双跳 / range多格联合选中(>2000兜底) / 排除区间A1-A9,B2-D2,A1:A9 / 去掉跳转成功弹窗

## 当前待办（下一轮）
1. 替换仓库 gui_patch.py 为最新版 → Actions 重建 → 下载 excel_diff_gui → 测试机验收四点
2. 验收通过后：L2 防漏报抽查（cli.py+ai_summary.py）
3. FEEDBACK.txt 统一反馈单（A-D自动段 + E-H人工填写区）
4. regression 增加 shift/range/豁免金标准案例
5. GOLD.md 预期差异清单机制（靠每次拍板积累）

## 关键约定（易踩坑）
- 非程序员：所有代码改动按「行号+原文」提供；文件尽量整文件替换
- 测试机离线/无权限/不能装软件/无Python；只能跑免安装exe
- 跑 auto_run 前先手动打开 Excel 关激活窗，否则 COM 失败
- 只回传小文件给AI：AI_PACKAGE.txt + RULES_VERDICT.txt（以后 FEEDBACK.txt）；DIAG_FULL.json 留测试机
- 深挖：scenarios.json 加 "focus_pattern": "P002" 重跑
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格，匹配用「精确+去空格兜底」
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 诊断提示词模板
【分析】对比工具自动运行结果：
===== A. 金标准回归 =====（RULES_VERDICT.txt）
===== B. 真实数据摘要 =====（AI_PACKAGE.txt）
【任务】1)判断回归是否通过；2)逐条说明每个模式真伪；3)给出下一步
