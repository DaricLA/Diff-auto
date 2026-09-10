# PROGRESS 开发进度卡（新对话/新设备先粘贴本文件）

## 项目
Excel 差异审计工具（openpyxl 底层 vs Excel 显示层）+ 自动化开发闭环。
场景：测试机离线/无权限/仅免安装exe；开发在手机或电脑（AI+GitHub）；GitHub Actions 云端编译；U盘传数据。

## 最新状态（2026-09-10 v4.4：文件名配置 UI 改造 + 引擎修复 + 开关联动；本机离线自测全绿，待实机验证）
- **main.py 5054 行，L14 VERSION = "v4.4"**；cli.py / ai_summary.py / regression.py / build.yml 未动
- **v4.3.1（真 bug 修复）**：文件名一致性检查在真实流程中**静默失效**（永远 0 告警，日志仍打完成）。
  根因：引擎从 `new_wb._src_path`/`new_wb.filename` 取文件名，但 `_src_path` 全项目只读从不写入，
  openpyxl Workbook 也无 `filename` 属性。修复：`_load_workbooks` 里加 `old_wb._src_path=self.old_path;
  new_wb._src_path=self.new_path`（写在 L1681 同一行末尾，O全文件行号不变）。
  ⚠️ **上线后第一次文件名检查会真的开始报警，这是预期行为变更，别当成新 bug**
- **v4.3.2（行为变更，用户拍板）**：RuleEditorDialog._open_engine_config（L4176）配置对话框点「确定」后，
  该子引擎开关若为关则**自动打开**（防止「配了但开关没开 → on_ok 静默丢弃」）；点「取消」不动。四个子引擎统一
- **v4.4（UI 改造，FileNameCheckConfigDialog L3353 重写；引擎判定逻辑零改动）**：
  · 窗口 720x780 → **1120x870**（高 DPI 下 80 位/行的位标尺需 ~1040px）；文件名区 grid 三列、标签等长不再错位
  · **位标尺**（Canvas）：位号/当前(N)/模板(N) 三行，蓝框圈当前规则区间（两行同时圈、真实 bbox、垫底）、
    80 位/行自适应、行距取字体实测 linespace（兼容高 DPI）；已删「其他规则灰线」
  · **详情条**（表格下方）：第x~y位 │ 取到内容 │ 结论，红/绿实时刷新（取代「给表格加列」方案）
  · **保存前可疑清单**：①日期段不可识别 ②结束位超长 ③模板名为空 ④模板已填但规则表空；不阻断；
    「模版段不一致」属正常产出，不计入清单
  · 模式/目标联动：日期模式选「模版」→ 模式自动回「文本」（不再静默弹回，详情条给说明）；反之亦然
  · 新增模块级纯函数 `validate_filename_rule`（L3299）；文件名侧复用引擎 _parse_date_seg；单元格侧不打开文件
- **本机验证状态**：引擎级离线自测 **全绿**（14/14、32/32、42/42、10/10，stderr 干净，py_compile 过）；
  对话框已实测可构造/可映射/窗口可见而且几何正确。**尚未做**：GUI 全流程手点验证、
  Excel COM 全流程、152MB 冒烟、测试机实机
- ⚠️ **行号已变**：v4.4 起 FileNameCheckConfigDialog 之后的所有符号 = 旧行号 + 384~388（全文 4669→5054），
  以 STRUCTURE.md 为准；需查符号当前行号可跑 `_selftest\symbol_lines.py`

## 上一版状态（2026-09-10 v4.3：B 版功能合并 + 三处补丁回补）
- **v4.3 = v4.2 定稿 + B版(mian-new.py)两大功能合并 + 三处补丁回补**（main.py 4669 行，L14 VERSION="v4.3"）
  - ① 统一滚轮 `_chat_wheel`（L187）：滚轮只滚光标所在对话框的滚动区（文件名/数据趋势/规则编辑器
    三窗已接入；删除旧 `_ds_wheel` 方案）
  - ② 文件名一致性 v5：规则列表 rules[]（起始/结束/目标`模版|待测文件`/模式`文本|日期`/Sheet/单元格）
    + 日期模式（文件名 yyMMdd→固定20xx / yyyyMMdd；单元格 datetime/Excel序列号/7种文字格式）
    + 对话框 720x780 +【从旧版读取】
    ⚠️ 旧配置（template_enabled/segments）不兼容，需重配规则
  - ③ 三处补丁回补（此前合并误用旧稿导致缺失，已修复）：`_waiver_fail_desc`（L1853）
    + 两分支 else「超限/抓取失败/两侧全空/非数值→**直接未豁免**」（防漏）
    + 提醒分支删除（提醒职责分离）→ 回到已拍板的 **方案F** 语义
- 本机验证状态：**待跑** —— py_compile → 小表3场景单测（提醒命中→橙字+规则判定照常 /
  豁免限内→救场 / 豁免超限→未豁免行）→ 文件名 v5 文本/日期模式测试 → 152MB 冒烟 → 回報
- 上传后待做：构建 3 exe → 测试机实机验证（三窗口滚轮、文件名 v5 配置、四修复场景）
- 当前文件状态：main.py 4669 行（已清理多余空行）；cli.py / ai_summary.py / regression.py / build.yml 未动
- ⚠️ 版本防混淆：存在两份 4561 行旧档（未含补丁的「旧稿」vs 含补丁的「定稿」）——
  辨识法：搜 `_waiver_fail_desc`（旧稿 0 处 / 定稿 3 处）

## v4.2 纪要（2026-09-08：引擎分层 + 四项修复）
- 引擎分层模型（用户拍板）：提醒型（特殊提醒/文件名/趋势）=只产告警diff（橙色斜体、永远「需人工复核」，
  不参与规则豁免）；豁免型（数据限界豁免）=唯一参与判定；基础判定=规则检查项；多规则聚合=AND。
- 四项修复：①高级告警跳过规则匹配 ②shift/描述统一写缓存计算值 ③规则执行逐条日志（跳過=红、其餘=绿）
  ④日志着色语义（_log_tag：异常/识别=红、成功=绿、耗时节点=黑粗）
- 三处补丁：_waiver_fail_desc + 两分支 else 防漏 + 提醒分支删除（=方案F）——v4.2 曾核对通过，
  但本次合并一度丢失（底稿用了旧稿），v4.3 已回补

## v4.1 纪要（2026-09-08：数据限界豁免）
- 高级检查区第4项【数据限界豁免】（豁免型）：命中规则数据源单元格满足条件→直接豁免（不再跑常规检查项）
- 两种模式：☑无条件豁免（冻结输入）/ 数据范围 LCL≤v≤UCL（含边界；只填一侧只查一侧；两侧全空→不豁免防漏报）
- UCL/LCL 值来源双模式（行/列=相对规则锚点偏移抓取 或 直接数值）；shift 模式禁用（UI 拦截）
- 实现位置：LimitWaiverConfigDialog（v4.3 行号 L3655）+ 过滤期豁免逻辑（_build_waiver_map L1782 /
  _waiver_lim L1814 / _waiver_judge L1836 / _apply_rule_filter L1880 豁免分支）；豁免引擎不进 REGISTRY

## 里程碑（2026-09-06 达成 ✅）
✅ 金标准回归 14/14；L2 全场景 149/149 全绿（B1 4/4、B2 115/115、B3 10/10、B4 12/12、B5 8/8）
✅ 业务拍板落地验证：行高15/列宽8.43 豁免、字体别名归一、主题色按显示色判定
✅ 主要修复：CheckProject参数错位、value开关消费、数组公式门控、B4/B5工作簿混用、主题色索引(accent2)、
   断言语义体系(DIFF=存在/SAME=豁免/NODIFF=无差)、图片tempfile；诊断导出已删
（2026-09-06 各批次明细：列宽管理/进度条双循环/白底/趋势与文件名配置/性能修复——
  已并入 v4.0/v4.1 版本内容，明细从略）

## 仓库与文件（GitHub: Diff-auto, Public）
- main.py        v4.4（5054行；2026-09-10 = 文件名配置 UI 改造 + v4.3.2 开关联动 + v4.3.1 引擎修复；
                    前版：v4.3=B版功能合并+三补丁回补、v4.2=引擎分层四修复、v4.1=数据限界豁免、
                    v4.0=列宽/进度条/趋势/文件名/白底）
- gui_main.py / gui_patch.py   测试按钮 v2.2 已验证（未动）
- cli.py / ai_summary.py       auto_run.exe（FEEDBACK 阶段再动）
- regression.py  单文件（约1154行）：金标准14 + L2 149用例，输出 RULES_VERDICT.txt + L2_REPORT.txt
                 （待补：豁免限内/超限/无条件、shift 缓存值、高级告警不参与豁免 → 14→20+）
- build.yml      已清 l2_scenarios；main.py 改动需全量重建3个exe
- dev/PROGRESS.md / dev/STRUCTURE.md   （2026-09-10 已更新至 v4.4）
- _selftest/（本机测试台，不参与 exe 构建）：run_tests.py / test_filename.py /
  test_filename_validate.py / test_engine_toggle.py / proto_dialog_live.py + 运行原型.bat /
  _merge_v44.py / symbol_lines.py / show_project_json.py / main_backup_v4.3.2.py

## 下一步（按顺序）
1. 上传 main.py（v4.4）→ 触发 Actions **全量重建 3 个 exe**
2. 测试机实机验证：① 三窗口滚轮 ② 文件名一致性（**首次会真报警，预期**）
   ③ 文件名配置新界面（位标尺/详情条/保存前清单/模式目标联动）④ 四修复场景
3. （可选）本机跑 Excel COM 全流程 + 152MB 冒烟
4. regression.py 补用例（14→20+）；考虑把 _selftest 的引擎用例（14+32 条）并入金标准
5. FEEDBACK.txt 统一反馈单（AI 起草模板 → 用户填写/确认 → 入库）
6. L2 防漏报抽查

## 关键约定（易踩坑）
- 非程序员：代码修改必须给【行号+原文+替换文本】，代码块内无行号（手机无法搜索）
- main.py 改动需用户拍板且零行为变化+回归验证；行为变更必升版本
- 测试机：离线/无权限/无Python；跑前先手动打开 Excel 关激活弹窗；
         跑 regression 前须关闭 Excel 中打开的 l2_* 文件（否则 Permission denied）
- 回传小文件：AI_PACKAGE.txt / RULES_VERDICT.txt / L2_REPORT.txt；DIAG_FULL.json 留测试机
- 差异类型名：值变化=「内容变化」；Sheet 名可能含尾部空格
- shift：垂直范围必须合法非空否则跳过；old列+shift_offset=新列
- 文件名 v5 旧配置不兼容（需重配）；日期模式固定 20xx 补全
- 【v4.3.2】配置子引擎时点「确定」= 自动开开关（不会再白填）；配置对话框点「取消」不影响开关
- 【v4.4 高DPI】像素不能写死：本机 Tk scaling 2.0（等宽字符 12px、行高 26px），
  标尺行距必须用 `font.metrics('linespace')`；蓝框用文字真实 bbox 画再 tag_lower
- 【v4.4】tkinter trace 回调会传 (name,index,mode) 三参，绑定必须用 `lambda *a:` 包一层，
  否则打字时回调静默报错、界面不刷新
- 【v4.4】子引擎配置存在项目 JSON 的 rules[].advanced_engines[]，落盘路径 = `<exe目录>\<项目名>.json`；
  U 盘拷 exe 时**要连同名 json 一起拷**才有规则
- 业务结论：Oven profile 整批+6天 = 预期差异（已拍板）

## 换对话交接（下一步就做）
1. 上传 main.py（v4.4）构建 3 exe → 测试机实机验证（重点：文件名检查首次真报警）
2. 贴回 PROGRESS + STRUCTURE（均为最新），需要符号行号就跑 `_selftest\symbol_lines.py`
