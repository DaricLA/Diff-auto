# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 仓库与文件（GitHub: Diff-auto, Public）
- main.py        v3.99（含 shift_scope 提取，_apply_rule_filter 调用它；零行为变化，已稳定）
- gui_main.py    启动入口（import main → gui_patch apply）
- gui_patch.py   测试按钮 v2.2 已验证（引擎同源：offset/intersection/range→locate_all 双侧；
                 shift→shift_scope；跳转复用 jump_to_excel；文件未开→询问打开→ComCheckDialog 确认；
                 主窗口经 _VIEWER 全局登记获取）
- cli.py / ai_summary.py   auto_run.exe（暂未动；后续 FEEDBACK 阶段再改）
- regression.py  【当前主战场】v3 合并单文件（约1120行）：金标准14案例 + L2 全场景
                 B1单项矩阵/B2干扰矩阵/B3全局交叉/B4 COM分歧12例/B5主题专项8例
                 suite() 汇总 ~150 case；输出 RULES_VERDICT.txt + L2_REPORT.txt
- l2_scenarios.py 【已删除】曾并入 regression.py；如仓库仍有请删除
- .github/workflows/build.yml  增量打包；regression 编译必须不含 --hidden-import l2_scenarios
- dev/PROGRESS.md / dev/STRUCTURE.md

## 已验证 ✔
- 金标准回归 14/14（v3 单文件下仍全部通过；v13 公式变化实际报"内容变化"单测通过）
- 测试按钮 v2.2 全流程（引擎计算→COM确认→跳转）
- main.py v3.99 shift_scope 提取（回归无回退）

## 当前进度（L2 全场景开发中）
- [x] L2 框架完成：regression.py 合并单文件（解决 pyinstaller 不收集 l2_scenarios 的问题）
- [x] 排查出 7 处 bug 修复方案（已给精确行号，基于用户当前 1120 行版 regression.py）：
      1) 3 处图片 BytesIO → 临时 PNG 文件（打包环境"closed file"崩溃）
      2) _b2_full13_asserts：A12（列宽行）无单元级差异 → 断言需跳过/NODIFF；A13 无 PIL 时同理
      3) 两两组合 t1=='col_width' 时交换（否则 A2 无 diff 断言误报）
      4) shift 断言：B12 改 SAME（配对列宽无差→豁免）；A13 改 DIFF（配对图片有差→不豁免）
- [ ] 应用 7 处修复 → 重建 → 测试机跑 → 传回 RULES_VERDICT.txt + L2_REPORT.txt
- [ ] 分析 L2 报告：硬断言❌=引擎真bug（修）；观察点/FP/FN候选=用户拍板
- [ ] 后续：FEEDBACK.txt 统一反馈单 → GOLD.md 预期差异登记 → regression 金标准扩展

## 关键约定（易踩坑）
- 非程序员：代码修改必须给【行号+原文+替换文本】，代码块内无行号（手机无法搜索，已强记忆）
- main.py 改动需用户拍板且零行为变化+回归验证
- 测试机：离线/无权限/无Python/只能跑免安装exe；跑测试前先手动打开 Excel 关激活弹窗
- 回传小文件：AI_PACKAGE.txt/RULES_VERDICT.txt/L2_REPORT.txt；DIAG_FULL.json 留测试机
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格（精确匹配不 strip）
- shift：垂直范围必须合法非空否则跳过；old列+shift_offset=新列
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 换对话交接（下一步就做）
1. 用户把仓库 regression.py 更新为"已应用 7 处修复"的最新版 → 确认 build.yml 已删
   --hidden-import 与 l2_scenarios.py 名单行 → 删除仓库 l2_scenarios.py
2. Actions 重建 → 下载 regression → 测试机运行 → 传回两个 txt
3. 贴回 PROGRESS + STRUCTURE + 两个报告 → AI 分析
