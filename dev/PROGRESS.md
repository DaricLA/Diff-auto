# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 仓库与文件（GitHub: Diff-auto，Public）
- main.py        原程序 v3.98 → 已获准小改：v3.99 新增 OpenpyxlComparer.shift_scope(旧wb,新wb,rule,log_fn)
                 （提取 shift 引擎逻辑：列配对+垂直范围，_apply_rule_filter 改为调用它，零行为变化）
- gui_main.py    GUI 启动入口（import main → gui_patch apply → 启动）
- gui_patch.py   【测试按钮补丁 v2.2 已验证】不再自写范围计算，全部指向引擎：
                 * offset/intersection/range → DataLocator.locate_all() 旧/新双侧
                 * shift → OpenpyxlComparer.shift_scope()
                 * 点测试→计算→文件未开则询问并打开(40s等待界面不卡)→主界面同款 ComCheckDialog
                   （检查COM通道→确认开启）→自动跳转（旧→新，只选数据区=shift垂直行×配对列）
                 * 跳转复用主程序 jump_to_excel（支持联合多区）；成功无弹窗，失败才提示
                 * 主窗口通过补丁登记 _VIEWER（DiffViewer.__init__ 打补丁）获取
                 * 保留：排除区间解析 A1-A9/B2-D2/A1:A9（同时用于真实规则保存）
- cli.py         auto_run.exe：读 scenarios.json，批量对比+COM采集+摘要（待加 FEEDBACK/L2抽样）
- ai_summary.py  摘要生成：模式归并/类型概览/COM判定/focus_pattern（待加 FEEDBACK/L2抽样）
- regression.py  regression.exe：14案例金标准回归（待加 shift/range/豁免案例）
- scenarios.json 运行清单（测试机记事本编辑）
- .github/workflows/build.yml  增量打包（main.py / gui_patch.py 变化会重建 GUI）
- requirements.txt / dev/PROGRESS.md / dev/STRUCTURE.md

## 已完成
- [x] Actions 增量构建成功；测试机 3 exe 全链路跑通
- [x] 回归 14/14；COM 采集 12/12
- [x] 第一次真实数据诊断：4278 差异≈真实批量更新（Oven+6天已确认=预期差异；Summary/Raw Data 真实）
- [x] 测试按钮全流程验证通过（v2.2）：
      - 引擎实测范围正确（如 offset 双侧 H4）
      - 文件未开：询问→打开（40s，不卡界面）→ ComCheckDialog 确认→自动跳转
      - shift 数据区联合选中（垂直行×配对列）；range 排除区间/多格联合选中
- [x] main.py v3.99 shift_scope 提取（引擎与测试按钮共用同一函数，单一事实源）

## 当前待办（下一轮）
1. L2 防漏报抽查（cli.py+ai_summary.py）：抽样模块——每sheet 10格+边界5格+程序判定"相同"10格，
   用 COM 读显示层对照程序判定 → 产出漏报/误报候选
2. FEEDBACK.txt 统一反馈单：自动段（摘要/COM/回归）+ 人工填写段（E抽查/F新异常/G拍板/H GUI验收）
3. regression.py 增加 shift 分组迁移/range/豁免 golden cases（14→20+）
4. GOLD.md 预期差异登记机制（Oven 批量更新=预期差异 先登记）

## 关键约定（易踩坑）
- 非程序员：所有代码改动按「行号+原文」提供；文件尽量整文件替换；缩进用整块选中复制
- main.py 允许小改，但必须先经用户拍板，改动需零行为变化+回归验证
- 测试机离线/无权限/不能装软件/无Python；只能跑免安装exe
- 跑 auto_run 前先手动打开 Excel 关激活弹窗（COM 需通道正常）
- 只回传小文件给AI：AI_PACKAGE.txt + RULES_VERDICT.txt（以后 FEEDBACK.txt）；DIAG_FULL.json 留测试机
- 深挖：scenarios.json 加 "focus_pattern": "P002" 重跑
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格，GUI 下拉抓取后精确匹配，不 strip
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 诊断提示词模板
【分析】对比工具自动运行结果：
===== A. 金标准回归 =====（RULES_VERDICT.txt）
===== B. 真实数据摘要 =====（AI_PACKAGE.txt）
【任务】1)判断回归是否通过；2)逐条说明每个模式真伪；3)给出下一步
