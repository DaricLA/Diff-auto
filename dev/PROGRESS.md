# PROGRESS 开发进度卡（新对话先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层数据 vs Excel 显示层差异）+ 自动化开发闭环。
场景：测试机离线/无权限/仅能跑免安装exe，开发在手机（AI+GitHub），编译靠 GitHub Actions，U盘传数据。

## 仓库与文件（GitHub: Diff-auto，Public）
- main.py          原程序 v3.98（GUI；含 OpenpyxlComparer / ExcelCOMVerifier / CheckProject）→ 编译 excel_diff_gui.exe，勿动
- cli.py           命令行入口 → auto_run.exe：读 scenarios.json，批量对比+COM采集+摘要，双击运行
- ai_summary.py    摘要生成：模式归并(P001...)、COM显示层判定、类型概览、focus_pattern 深挖
- regression.py    金标准回归 → regression.exe：14个案例验证检测逻辑，双击运行
- scenarios.json   运行清单（测试机上用记事本编辑，相对路径 data/old.xlsx 等）
- .github/workflows/build.yml  云端编译出 3 个 exe（PyInstaller）
- requirements.txt 依赖清单

## 已完成
- [x] regression.py 两处笔误修正（'值变化'→'内容变化'；v12 合并用例重写）→ 回归 14/14 通过
- [x] ai_summary.py 新增【类型概览】段（类型|处数|模式数|代表位置），构建 08fcc59 已生效
- [x] 第二次诊断（版本 08fcc59）：4278 处差异
  · 内容变化 4277 处 = 真实差异（Oven profile 整批+6天；Summary/Raw Data 数值变化）
  · 格式类 98 处 = 待深挖判定真伪（条件格式48/图片尺寸34/边框4/对齐3/数字格式3/图片新增2/图片变动2/填充2/公式1）
- [x] COM 采集成功 12/12（覆盖边框、对齐、数字格式、填充全部 12 条样式类差异）

## 当前待办（下一轮）
1. 用户确认：Oven profile 整体+6天是否预期新数据
2. 深挖格式类判定真伪：P051 条件格式 → P059 图片尺寸 → P061 边框 → P063/P064 W18 → P168 填充 → P167 公式
3. 小修（下轮代码）：① RULES_VERDICT 版本号显示"?"（main.VERSION 缺失，改读 version.txt）；
   ② 摘要"总单元格: 0"（stats.total_cells 未传入）
4. AI 重点诊断格式类差异（条件格式/图片/边框等）是真差异还是「底层 vs 显示层」误报
5. 待用户确认：a) Oven profile 整体+6天是否预期新数据；b) 日常主要关注哪类差异（决定是否配高级审核过滤噪音）

## 关键约定（易踩坑）
- 测试机离线、不能装软件、只能跑免安装exe；无 Python
- 跑 auto_run 前必须先手动打开 Excel 关激活窗，否则 COM 失败（接收者已拒绝直接呼叫）
- 只传小文件给 AI：AI_PACKAGE.txt + RULES_VERDICT.txt；DIAG_FULL.json（几MB）留测试机
- 深挖模式：scenarios.json 加 "focus_pattern": "P002" 重跑 → 生成 FOCUS_P002.txt
- 差异类型名：值变化实际叫「内容变化」；Excel 未激活时 COM 需手动开 Excel
- 版本号由 Actions 写入 version.txt（如 ac36b3c），摘要头部可见

## 诊断提示词模板
【分析】对比工具自动运行结果：
===== A. 金标准回归 =====
（RULES_VERDICT.txt）
===== B. 真实数据摘要 =====
（AI_PACKAGE.txt）
【任务】1)判断回归是否通过；2)逐条说明每个模式真伪；3)给出下一步
