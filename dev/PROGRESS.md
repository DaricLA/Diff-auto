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

## 下一步（顺序）
1. FEEDBACK.txt 反馈单（用户→开发问题/建议收集；含 COM 观察标签修正：观察行"程序:报差/无差"
   应取引擎实际结果而非设计值）
2. GOLD.md 预期差异登记表（把豁免/等价规则登记为正式预期：行高15、列宽8.43、字体别名、
   主题色真差异、浮点尾差、公式大小写等价、索引色等价等）
3. main.py VERSION 升 v3.99；更新 PROGRESS/STRUCTURE
4. 开放后续：cli/auto_run 联动新引擎能力、业务场景回归扩展

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
