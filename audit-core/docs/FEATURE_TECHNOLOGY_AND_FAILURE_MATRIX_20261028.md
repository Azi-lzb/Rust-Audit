# Rust-Audit 按功能的技术链路、回退与失败策略清单

资料核对日期：2026-10-28。覆盖九项正式功能，86个实际步骤、104个失败场景、31组证据索引。此文用于业务决定失败处理范围；没有更改产品策略。

## 先读这几条

- **没有 Calamine 普通值读取失败→OOXML 普通值读取的通用回退。** 普通数据使用 excel-reader/Calamine。命名区域、条件格式、批注、关系、工作表范围等使用专用 ZIP/XML 读取，是并行职责或补读。只有重算结果校验的现代错误 token 兼容读取在限定错误下修改临时副本后再用 Calamine 重读。
- **Rust 当前生产修改只有 Umya，openpyxl 是候选兼容适配器。** excel_route::production_modifier 只登记合格 Umya。框架有写前拒绝后换后端机制，但只有一个合格生产后端时不会自动换 openpyxl。写后失败或验证失败也明确停止回退。
- **旧 Python S1-F1 的 OOXML→openpyxl 外部表复制回退不等于 Rust 现状。** 旧 Python 的纯数据 OOXML 快路遇公式转 openpyxl；当前 Rust 直接用 Umya ImportRange 保留外部公式，不以‘外部表有公式’拒绝整条链。新补的非 A1 起点和跨表依赖闭包用例仍未运行。该问题属于 S1-F1。
- **Windows 实际为 PowerShell Excel COM，并非 pywin32。** sidecar 使用 powershell.exe / New-Object -ComObject Excel.Application / CalculateFullRebuild / SaveAs。你要求 pywin32 COM 或 LO；Windows 桥接技术仍需定向调整。切换 pywin32 不会自动增加 Excel 的公式兼容能力。
- **auto 是按平台选择，失败后不会 COM→LO 或 LO→COM 自动换引擎。** Windows auto 选 Microsoft Excel COM，Linux auto 选 soffice。缺失即明确失败。没有成功 LO 真重算证据，也没有跨引擎回退成功证据。
- **条件格式 XML 存的是规则，不是‘本轮已触发’的审核结论。** 项目读取 OOXML CF 规则/AppliesTo，再由 Rust 有限求值器结合单元格存储值判断 cellIs/expression。色阶/图标集/数据条、未支持表达式与上下文必须带诊断，不能用‘规则存在’替代触发或将其当通过。
- **S1 CLI 和 GUI 的职责及扫描口径存在差异。** GUI S1-F1 完整复制→真重算→抽取，并对重算副本扫描校验区域。s1-audit-directory 接收已复制审核副本的 manifest，可显式重算或明确使用存储结果，当前校验区域扫描使用 source_file；它不是 GUI 复制入口的完全替代。需用源缓存与副本不同的样本确认预期。
- **旧验证记录要按范围拆开。** 12:28前历史回归435通过/0失败/28忽略，不覆盖15:26 S1/S5修复；旧 S3-F3计划测试未写回，Windows能力样本仍是独立副本，Linux报告才记录备份后原地更新。旧 S5输出比对未覆盖非空CF结果。

**当前整体判断：** 九项功能已有相应实现，但不能把所有细分路径、所有回退和最新源码统一标成测试成功。当前最明确的缺口为 Windows 仍用 PowerShell 调用 Excel COM、未接生产 openpyxl 回退、无成功 LO 重算证据、S1/S5 新修复未执行测试。

**本轮没有编译、运行程序或执行测试。** 现有二进制约12:18生成，最新15:26 S1/S5修复未编译执行；下面的“R/H”都是限定范围历史证据。测试通过表示实现与既定预期一致，可能包括正确报告业务异常/取数失败；不是输入数据没有异常。

[交互清单源码](C:/Users/AZI/.cursor/projects/1791953125140/canvases/audit-technology-fallback-failure-policy.canvas.tsx)已生成，设计上可以按功能、证据和关键词筛选，并记录业务决定。其语法和基于简化 React 声明的组件类型已静态检查；**尚未验证它在当前 Codex 宿主中的渲染和状态保存**，也不会自动更改产品代码或配置。本文及[结构化JSON](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/docs/FEATURE_TECHNOLOGY_AND_FAILURE_MATRIX_20261028.json)包含完整清单，可以直接审阅。

## 快照与范围

- 仓库：C:/Users/AZI/Desktop/VibeCode/Rust-Audit
- 分支：main
- HEAD：fd7c982f0358a74e2eb4f9cb74d022f88a1dfafe
- 快照时间：2026-10-28 16:10 +08:00；21:20 再次核对这 61 个文件的 SHA-256 均未变化
- 包含此前多轮未提交修改；HEAD 不代表当前源码。报告任务只新增说明和分析附件。
- 本轮未编译、未运行程序、未执行测试。15:26 S1/S5 修复晚于现有二进制及12:28历史验证，仍未执行验证。
- dist/open-data-audit.exe 与 lo-convert.exe 修改时间约12:18，不能代表15:26后的源码。
- 61个相关源码/记录文件的SHA256已写在配套JSON的sourceSnapshot，只作本报告对应版本定位。

- 技术判断基于当前 Rust 真实调用代码；只读取这九项功能所需的旧 Python 参考与历史证据。没有执行全项目审计。
- 逐项区分主路径、元数据补读、有限兼容重试、业务范围兜底、失败回滚、替代引擎。它们不能互称为回退成功。
- 所有主路径和回退测试栏都说明证据范围；历史成功不等于当前未提交源码成功，负向用例正确报错不等于回退输出成功。
- 本报告不修改任何业务失败策略。每条待决策项的用户决定默认为‘待决定’，建议只是供业务判断。
- 继续执行只代表允许产生带缺项说明的结果；缺数据、规则不支持、取数失败不能改记为审核通过。
- 源文件安全、公式与引用、关键结构、条件格式审核判断属于业务正确性；普通字体、边框差异不扩展范围。

## 如何读“测试成功”和“回退”

- **R 历史报告可核**：找到机器可读历史报告，所述断言/计数成立；只证明该版本、平台和样本范围。
- **H 历史记录声称通过**：有修改记录或汇总日志；缺少该细分步骤独立成功凭证，或不能定位到当前源码。
- **U 当前修改未运行**：当前源码/新增测试存在，但未编译执行；静态检查不算通过。
- **N 未找到成功证据**：没有定位到覆盖这个路径及触发条件的成功证据。
- **X 存在失败或未完成**：报告明确失败、不完整、未验证，不能算通过。
- **NA 不适用**：当前生产路径没有该回退，或该步骤不涉及公式重算。

所有步骤的“当前版本”均未在本轮运行。某步骤引用历史R只证明列出的平台/样本/断言；没有该步骤独立证据的标H/N。框架mock测试不证明真实后备引擎可用；缺少引擎时正确报错也不证明另一引擎回退成功。NA表示当前没有生产回退，不能写成“回退通过”。

回退分为：限定错误后重读、业务模板/范围兜底、编码兼容、合格替代引擎、文件事务回滚。元数据补读、可选输入未执行、跳过文件/规则，不是同一个意思。

## 技术库实际分工

### Calamine 0.37.0

- 用途：值、存储公式结果、公式文本、合并信息；按入口支持 XLS/XLSX 等
- 当前路径：生产主读；大多数错误直接返回调用方。read_value_columns 限制返回列，不保证完全流式或不解压整张表。
- 回退：不接通用 OOXML/openpyxl 数据读取回退；仅 sidecar 校验 opt-in 现代错误 token 临时归一后重读。
- 边界：不编辑、不计算公式；data_only/存储值是上次保存值。

### rust_xlsxwriter 0.101.0

- 用途：S1结果/汇总、S2结果、S3F1/F2/F4/冲突报告、S5结果等新工作簿
- 当前路径：Excel-Writer 薄封装 re-export；生成新文件，部分入口 Calamine 独立读回。
- 回退：无生产替代 writer；写失败返回错误。
- 边界：不是原工作簿编辑器，不计算公式；新结果通常由 Rust 先算好再写值。

### Umya 3.1.1

- 用途：S1F1 模板/外部区域 ImportRange；S3F3 SetCellValue；S3F5 表单结构与值修改
- 当前路径：ExcelEditPlan→ExcelModifier→Umya；目前唯一登记的生产编辑后端，限定Windows/Linux x86_64。
- 回退：框架能在 UnsupportedBeforeWrite 后试下一个合格后端；现生产无第二个。
- 边界：有对象预检；删除列且存在定义名称拒绝；普通公式支持不代表共享/数组/未知对象均支持。

### Direct OOXML（ZIP/XML）

- 用途：名称/CF/批注/关系/范围元数据；公式缓存清理、重算属性补丁、包与保留项验证
- 当前路径：按对象专用读/补丁，不能称通用编辑/读取后备引擎。
- 回退：可做限定错误的临时副本兼容；不篡改源文件，也不把公式文本当计算值。
- 边界：.xls BIFF不适用；跨工作簿整表复制涉及关系和引用，不能只拷 sheet.xml。

### 隔离 Python openpyxl worker

- 用途：已有工作簿修改的兼容候选；无审核规则计算
- 当前路径：源码及测试候选存在；当前 production_modifier 没有注册。
- 回退：不是已经可用/已通过的 Umya 后备。旧 S1 候选报告有失败与样式表缺陷。
- 边界：不作为 Calamine 读取回退、不运行 Audit 业务、不算公式；离线内置运行时，不能依赖用户系统 Python。

### Microsoft Excel / COM

- 用途：S1F1 可选外部计算能力；执行模板和外部公式真实重算
- 当前路径：Windows sidecar 当前 PowerShell COM；历史合成 Excel16重算E2E通过。
- 回退：无运行时切到 LO 的链路；按你的要求建议改为隔离 pywin32 COM。
- 边界：要有真实 Excel COM；pywin32只是桥，安装pywin32不能替代安装Excel；当前无子进程超时/CalculationState等待保护。

### LibreOffice Calc / soffice

- 用途：S1F1 Linux 外部重算；sidecar 也有显式 lo 选择
- 当前路径：headless、独立用户profile、临时输入/输出、重算后验证。
- 回退：缺失明确失败；已见缺失引擎负向记录，未见成功LO重算/跨引擎回退证据。
- 边界：Linux核心功能不强制安装LO；启用公式重算须提供可用离线组件/外部安装来源并在目标平台验证。

### Rust 业务计算 / Decimal / 表达式解析

- 用途：S2、S3、S5 配置规则、配对、单位换算；CF有限求值
- 当前路径：这些业务规则不需要Excel重算；S5不注入模板审核公式。
- 回退：缺上期/外部/指标一般按具体规则报错；不是把Excel公式交给Rust通用求值器。
- 边界：有限DSL/表达式只支持已实现语法；S1Excel任意公式仍交Excel/LO。


## 九项功能概览

### S1-F1 模板审核 / 汇总核查表校验

- 真实入口：GUI：S1 F1→start_dag_flow→start_s1_f1_workflow→execute_s1_f1_workflow；CLI s1-audit-directory 处理已复制副本manifest，可显式重算，另有诊断命令。
- 输入：源XLSX目录/勾选文件、模板目录/选择、1.0映射/历史配置、可选外部XLSX、输出目录及计算sidecar。
- 输出：每个通过预检文件的独立重算审核副本 + 合流、历史分类后的13列审核结果。
- 实现/验证状态：已有完整GUI主链的历史合成Excel COM通过记录；新外部公式依赖/非A1坐标修复未运行。
- 当前停止策略：预检不匹配跳过文件；正式执行的复制/重算/抽取/CF不完整任一出错，GUI整批停止，run目录guard清理。
- 尚缺/限制：pywin32桥未实现；LO成功重算未证；openpyxl生产回退未接；当前外部公式修复未运行；CLI/GUI校验区域扫描对象有差异。
- 证据：[E02](#e02)、[E03](#e03)、[E05](#e05)、[E06](#e06)、[E07](#e07)

### S1-F2 汇总校验结果说明 / 命名区域汇总

- 真实入口：GUI：S1 F2→start_dag_flow→run_s1_region_summary_worker；CLI s1-region-summary。s1-merge-workbook-tables是首行表头合并辅助入口，并非命名区汇总的等价替代。
- 输入：已选源XLSX、模板及1.0汇总映射；任意行/固定行命名区、表头区域、附加字段；可选历史。
- 输出：新建多sheet汇总XLSX，业务列+附加字段+历史字段+来源定位，不修改源文件。
- 实现/验证状态：历史主汇总和‘兜底模板’合成回退有成功报告；真实模板全覆盖和当前GUI未验收。
- 当前停止策略：表结构预检可跳过文件；正式读取失败、公式缓存无法确认、配置/输出验证错误会停止本次汇总；缺某sheet/空白数据按规则跳过并计数。
- 尚缺/限制：没有Calamine/Writer替代引擎；输入含公式时只用保存缓存，未重算。真实批次旧partial记录不能算全部通过。
- 证据：[E08](#e08)、[E09](#e09)、[E07](#e07)

### S2-F1 执行报表规则

- 真实入口：GUI S2 F1→doCreatePeriodAudit→start_period_compare_v2；CLI s2。
- 输入：本期/上期报送目录、S2配置、可选外部集中系统XLSX、输出路径。
- 输出：新建XLSX三sheet：环比规则16列、外部核对规则16列、校验规则12列；进度/质量提示。
- 实现/验证状态：本上期真实子集、外部非时序真实子集及合成规则历史通过；当前GUI与异常输入策略需人工确认。
- 当前停止策略：不识别文件名跳过；已识别工作簿读取失败、已选外部读取失败、全局配置/输出错误停止本次。个别取数/规则错误按现规则记录诊断。
- 尚缺/限制：没有库级fallback；外部实际时序布局全覆盖未证；重复key首条保留的业务策略需用户判断。
- 证据：[E12](#e12)、[E13](#e13)、[E14](#e14)

### S3-F1 两期数据→比较文件

- 真实入口：GUI runCentralComparison→start_central_comparison；CLI s3-execution-compare。
- 输入：本期和上期大集中CSV/XLSX/XLS；3.0通用配置+3.1执行比较配置；输出/可选日志路径。
- 输出：按数据期命名的比较sheet，21个业务字段+人工‘输出说明’共22列；可选执行日志。
- 实现/验证状态：真实子集parity通过，但报告自标migration_in_progress；当前总体人工验收未完成。
- 当前停止策略：文件/表头/数据期/配置失败停止；部分异常行/表达式问题记录日志或按规则跳过；CSV存在编码兼容回退。
- 尚缺/限制：编码替换、重复key后条覆盖、缺上期/非数值、表达式错误是否可继续需业务定政策；小浮点尾差需按业务精度验收。
- 证据：[E15](#e15)、[E30](#e30)

### S3-F2 比较文件→说明文件

- 真实入口：GUI runCentralExplanationExport→start_central_explanation_export；CLI s3-export-explanations。
- 输入：F1比较XLSX（人工‘输出说明’标记）、可选3.0机构地区映射配置、输出目录。
- 输出：新建指标说明XLSX：说明文件8个业务字段+生成信息sheet。
- 实现/验证状态：历史真实比较簿2610选中行输出通过；与F3是不同流程。
- 当前停止策略：缺必需表头/读取失败停止；人工未选中记录正常不导出；无选中可产生0行输出，需用户区分空结果与失败。
- 尚缺/限制：手工更改表头/标记/机构映射歧义等负向策略待决定；没有替代reader/writer。
- 证据：[E16](#e16)、[E30](#e30)

### S3-F3 机构说明→说明文件（先备份、原地更新）

- 真实入口：GUI runCentralExplanationImport→start_central_explanation_feedback_import；CLI s3-import-explanation-feedback；只读预览s3-plan-explanation-feedback-import。
- 输入：待更新目标说明XLSX+一个/多个机构反馈XLSX/只读XLSM；输出计划/冲突报告位置。
- 输出：先生成目标旁字节备份，再原地更新目标‘变动原因’列；计划JSON/冲突XLSX/备份路径与回滚信息。
- 实现/验证状态：计划/冲突历史合成通过；Linux能力报告记录原地写回；Windows能力旧样本为独立副本，不能当原地事务验收。
- 当前停止策略：匹配冲突/未匹配行跳过且记录；合法行继续写；备份/编辑/输入变化/替换失败停止并保护或恢复目标。
- 尚缺/限制：Windows原地更新/锁文件/备份失败/恢复失败的当前版本实测未证；旧to_copy保留但不是当前生产契约；无openpyxl生产回退。
- 证据：[E17](#e17)、[E03](#e03)、[E04](#e04)

### S3-F4 本期数值核对

- 真实入口：GUI runCentralCrossPeriod→start_central_cross_period；CLI s3-check-current-values。
- 输入：本期CSV/XLSX/XLS，3.0+3.2；CLI支持可选上期，当前GUI主要本期输入；输出与可选日志。
- 输出：对比结果sheet17列：日期、机构地区、校验编码/名称/类型、备注、左右值/差异/绝对值/幅度、是否说明/说明内容/公式；可选校验日志。
- 实现/验证状态：真实183配置/163启用规则214结果行3日志行历史比对通过；不是所有规则均正常执行。
- 当前停止策略：输入/日期/配置读取错误停止；前提或左右取数失败记录日志后跳过当前规则；校验公式错误可带错误内容输出行；不回退Excel。
- 尚缺/限制：规则日志是否必须阻断、跨期能力GUI/CLI范围、工作簿异常类型、日志与主结果同时提交需业务验收。
- 证据：[E18](#e18)

### S3-F5 比较文件→金融表单

- 真实入口：GUI runCentralForms→start_central_indicator_split；CLI s3-generate-financial-forms。
- 输入：F1比较XLSX、3.3含金融表模板/报表清单/转表设置、3.0通用警戒配置、输出根目录。
- 输出：指标比较拆分_时间戳目录，每机构/地区一份基于3.3副本的金融表单XLSX，保持关键结构。
- 实现/验证状态：Umya编辑样本和历史合成CLI/API通过；真实3.3模板仍须验收。仓库3.3原本是占位，不能代替真实配置。
- 当前停止策略：任一机构编辑失败则整批暂存目录清理，不发布部分；所有机构成功才发布目录。比较表头不完全一致目前仅warning并固定列读取。
- 尚缺/限制：含定义名称且需删列会拒绝；表头错位可能按固定列误读；无openpyxl生产回退；真实43条启用表单配置未整体验收。3.3分析文件单位忽略，接收F1已按3.0目标单位的值。
- 证据：[E03](#e03)、[E04](#e04)、[E19](#e19)

### S5-F1 配置模式汇总核查表校验

- 真实入口：GUI S5 F1→start_s5→s5_web::run；CLI s5-audit-run。GUI同模板/同业务/同报告期的一批本期、可选上期按机构匹配，单外部物理簿；CLI支持显式多逻辑源。
- 输入：S5独立配置、已选本期工作簿、模板/模板目录、可选上期目录/报告期、可选外部XLSX、输出目录。
- 输出：汇总审核结果XLSX（核查表校验结果13列、条件格式触发13列、必要时质量提示）、只读审计副本、运行总账JSON、结构比较信息。
- 实现/验证状态：历史DSL/配置/V_PRE/区域投影/旧输出子集通过；最新GUI批次、owner/hash/外部身份/CF范围/缓存公式扫描修复全部未运行。
- 当前停止策略：全局路径/配置/模板/期间/来源合同失败阻断；文件身份或绑定错误可记录失败并继续其他；规则失败保留；CF不支持标PARTIAL；全部文件失败为FAILED，部分为COMPLETED_WITH_ISSUES；输入变化使成功结论失效。
- 尚缺/限制：最新7项定向测试未运行；非空CF输出、批次多文件归属与副本、只读快照新契约无当前成功证据。S5按Rust配置直接算，不迁入S1公式注入。
- 证据：[E20](#e20)、[E21](#e21)、[E22](#e22)、[E23](#e23)、[E24](#e24)、[E25](#e25)、[E26](#e26)、[E27](#e27)、[E28](#e28)、[E00](#e00)


## 每一步的技术、回退、测试与失败行为

### S1-F1 模板审核 / 汇总核查表校验

#### S1F1.01 选择文件、配置和输出位置

- 输入/前置：源目录、模板目录、勾选列表、配置、外部文件、sidecar。
- **当前技术**：Rust+现有HTML/Web API；std::fs路径规范化。
- 实际处理：限定输入/输出身份、过滤临时文件和已有产物；配置/目录先检查。
- **主路径测试：H**。历史合成CLI链及API输入检查记录；没有当前Windows GUI逐按钮验收凭证。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：路径/权限/sidecar路径非法：开始前阻断；无合格文件：阻断。
- 可优化：把环境检测和实际能力试启动分别展示；读取路径失败保留系统错误而非只显示未匹配。
- 需业务判断的场景：[G01](#g01)、[G02](#g02)、[G03](#g03)、[S1F1-E01](#s1f1-e01)
- 源码：[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[discovery.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/discovery.rs)

#### S1F1.02 模板目录及工作表匹配

- 输入/前置：源/模板工作表目录。
- **当前技术**：专用OOXML workbook.xml+relationships；Rust集合评分/目录发现。
- 实际处理：这里读取sheet名直接选OOXML元数据适配器；模板评分/覆盖和歧义判断，不是Calamine失败后回退。
- **主路径测试：R**。历史真实子集含模板建议匹配；当前GUI多模板/坏包需验收。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E05](#e05)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：未匹配、歧义、模板无效：预检记原因/排除；s1_template部分sheet名读取错误unwrap_or_default会变空集合。
- 可优化：记录‘XML无法解析’与‘业务不匹配’不同原因；模板目录元数据一次缓存。
- 需业务判断的场景：[S1F1-E02](#s1f1-e02)、[G04](#g04)
- 源码：[s1_template.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_template.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S1F1.03 读取业务配置、命名区域及表结构

- 输入/前置：模板、映射表、源表头/结构区域。
- **当前技术**：Calamine普通值+OOXML定义名称/批注/范围；Rust映射与结构比较。
- 实际处理：解析校验区域、条件格式区域、表结构区域及模板规则；按既定表结构检查源文件。
- **主路径测试：R**。历史配置/CF报告覆盖所列区域和结构；任意错误名称/损坏XML没有成功兼容证据。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E05](#e05)、[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：结构不匹配/无法确认：预检跳过该文件；全跳过则不开始；共享模板/配置加载失败阻断批次。
- 可优化：复用解析后的名称和网格；明确工作簿级/工作表级名称作用域，不能取错同名区域。
- 需业务判断的场景：[S1F1-E03](#s1f1-e03)、[G04](#g04)、[G05](#g05)
- 源码：[s1_template_rules.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_template_rules.rs)、[s1_template_inspection.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_template_inspection.rs)、[s1_structure_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_structure_compare.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S1F1.04 读取源值、模板公式和必要批注

- 输入/前置：已通过预检的源XLSX、模板XLSX。
- **当前技术**：excel-reader / Calamine值、公式、合并；OOXML元数据。
- 实际处理：获得复制范围及公式文本；必须能确认公式存在情况和必要结构。
- **主路径测试：R**。Umya能力样本前置读取及普通模板公式已覆盖；最新外部公式案例未运行。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E03](#e03)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：Calamine/元数据失败直接错误并经GUI停止整批；无Calamine→openpyxl/OOXML值读取回退。
- 可优化：按所需sheet/列/范围读并复用；未知现代错误token不能悄悄当空值。
- 需业务判断的场景：[G04](#g04)、[G05](#g05)、[S1F1-E04](#s1f1-e04)
- 源码：[s1_formula_copy.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_formula_copy.rs)、[lib.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-reader/src/lib.rs)

#### S1F1.05 选择外部工作表并补齐公式依赖

- 输入/前置：可选外部XLSX；模板公式跨表引用。
- **当前技术**：Rust公式引用扫描+Calamine；OOXML目标sheet目录。
- 实际处理：静态选模板引用sheet，递归外部公式引用形成依赖闭包；外部范围从原始起点复制到相同起点；公式不是拒绝条件。
- **主路径测试：U**。历史E03外部样本为普通lookup值；最新C4公式→另一sheet B1及坐标保护测试已写未运行。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E03](#e03)、[E00](#e00)。
- **回退类别/是否存在**：业务范围兜底；已有此限定分支。
- 回退触发：模板/外部引用含INDIRECT/OFFSET等不能静态判定的sheet依赖。
- 回退执行：扩大为导入全部外部sheet；仍检查名称冲突，不换编辑库。
- **回退测试：N**。存在falls_back_to_all_external_sheets_when_template_references_are_dynamic测试源码；未找到这一分支独立执行成功凭证；新闭包修复未运行。
- **当前失败处理**：外部必需但未选、缺引用sheet、与源sheet重名/大小写冲突：错误，GUI整批停止。
- 可优化：保持直接Umya带公式导入；若以后增加OOXML快路，只用于已核定纯数据契约，公式在写前路由到可保留公式的编辑器。
- 需业务判断的场景：[S1F1-E05](#s1f1-e05)、[S1F1-E06](#s1f1-e06)、[S1F1-E07](#s1f1-e07)
- 源码：[s1_formula_copy.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_formula_copy.rs)

#### S1F1.06 建立独立副本、清旧缓存和编辑计划

- 输入/前置：原源文件；模板/外部区域；新的输出路径。
- **当前技术**：std::fs临时副本+OOXML缓存补丁+ExcelEditPlan。
- 实际处理：源保持只读；修改临时副本并设置重算属性；保留关键公式与结构，禁止覆盖输入/已有输出。
- **主路径测试：R**。历史合成输入hash、缓存/重算flags及独立输出保护覆盖。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)、[E03](#e03)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：共享/数组/带特殊属性公式在缓存清理环节可能明确拒绝；空间/权限/同路径错误停止整批。
- 可优化：对共享/数组公式列为独立能力，不用普通公式用例替代证明；一次准备编辑批次。
- 需业务判断的场景：[S1F1-E08](#s1f1-e08)、[G02](#g02)、[G06](#g06)
- 源码：[s1_formula_copy.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_formula_copy.rs)、[lib.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/ooxml-patch/src/lib.rs)、[plan.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-editor/src/plan.rs)

#### S1F1.07 复制外部区域、模板校验范围及公式

- 输入/前置：ExcelEditPlan的CreateSheet/ImportRange/SetComment等操作。
- **当前技术**：ExcelModifier→唯一生产Umya 3.1.1；OOXML预检/保留校验。
- 实际处理：复制值、公式和必要批注；同坐标保留原公式对象，不做无谓翻译；跨坐标相对/绝对引用按编辑器翻译；不追求普通样式一致。
- **主路径测试：U**。旧模板公式与基础修改R；最新外部带公式/依赖/非A1起点仍U。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E03](#e03)、[E04](#e04)、[E00](#e00)。
- **回退类别/是否存在**：仅框架机制，生产无第二后端；当前没有实际可用的生产替代引擎/路径。
- 回退触发：后端UnsupportedBeforeWrite且候选文件仍空，且另有覆盖整批契约的Passed后端。
- 回退执行：框架会用原件新副本试下一个；当前生产只注册Umya，因此没有实际openpyxl回退。写后失败禁止回退。
- **回退测试：NA**。框架mock回退测试源码存在；这不证明实际Umya→openpyxl成功。旧openpyxl报告E10/E11反而未通过。
- **当前失败处理**：对象不受支持、公式翻译/包验证失败或无合格平台后端：整批停止；不自动试openpyxl。
- 可优化：先修具体ImportRange保留与引用问题；不要为有公式恢复纯数据拒绝，也不要用按值重建替代保留公式。
- 需业务判断的场景：[G07](#g07)、[G08](#g08)、[S1F1-E09](#s1f1-e09)
- 源码：[excel_route.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/excel_route.rs)、[umya_backend.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-editor/src/umya_backend.rs)、[s1_formula_copy.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_formula_copy.rs)

#### S1F1.08 真实公式重算并保存

- 输入/前置：已注入公式的独立副本。
- **当前技术**：Windows当前PowerShell→Excel.Application COM；Linux soffice/LO；独立sidecar。
- 实际处理：清旧缓存后的文件调用真正计算引擎，CalculateFullRebuild/保存或LO headless转换保存，获得新缓存。
- **主路径测试：H**。Windows Excel16合成E2E记录通过；Linux仅见缺少soffice的正确报错。pywin32路径当前不存在。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)、[E29](#e29)。
- **回退类别/是否存在**：按平台选引擎，不是运行时回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：auto按目标OS分派，Windows=com、Linux=lo。
- 回退执行：当前选择失败就报错；不会自动切换另一引擎，不使用旧缓存/公式文本冒充。
- **回退测试：N**。没有实际COM→LO/LO→COM成功回退证据；LO成功重算、pywin32桥都未证。
- **当前失败处理**：Excel未装/COM不可激活、PowerShell受限、LO不存在/超时/无输出：GUI整批停止。当前Command.output无明确超时。
- 可优化：按要求改独立pywin32 COM桥；加CalculationState完成等待、硬超时/进程回收、引擎版本回执；LO单独做真实成功样本。
- 需业务判断的场景：[S1F1-E10](#s1f1-e10)、[S1F1-E11](#s1f1-e11)、[S1F1-E12](#s1f1-e12)
- 源码：[main.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/sidecar/lo-convert/src/main.rs)、[s1_recalculation.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_recalculation.rs)

#### S1F1.09 验证重算公式/缓存并提交审核副本

- 输入/前置：计算前后XLSX；引擎JSON回执。
- **当前技术**：OOXML公式缓存计数+Calamine比较；限定错误token归一兼容；std::fs hard_link发布。
- 实际处理：确认引擎/输出/公式文本及保存结果；新文件无覆盖提交。缓存存在验证不等于每个复杂函数语义都正确。
- **主路径测试：H**。历史COM普通错误样本验证成功；现代token兼容测试源码存在但本轮未执行。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)。
- **回退类别/是否存在**：有限读取兼容重试；已有此限定分支。
- 回退触发：仅Calamine错误含Unsupported cell error value，且XLSX/XLSM确实发现未知literal错误token。
- 回退执行：临时副本把未知错误token保留为文本，再用Calamine读取；原件不变；只sidecar校验调用。
- **回退测试：N**。unknown_error_tokens测试源码存在；未定位到该分支单独成功运行回执，不能标回退测试成功。
- **当前失败处理**：回执引擎错误、公式改变、缓存不满足验证、发布hard_link失败：整批失败；无overwrite rename后备。
- 可优化：兼容读取若要复用于后续抽取，先确定统一‘错误仍是异常’契约并测试；不要因sidecar可读就假设后续也可读。
- 需业务判断的场景：[S1F1-E04](#s1f1-e04)、[S1F1-E13](#s1f1-e13)、[G06](#g06)
- 源码：[main.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/sidecar/lo-convert/src/main.rs)、[s1_recalculation.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_recalculation.rs)、[lib.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-reader/src/lib.rs)

#### S1F1.10 提取模板规则公式审核异常

- 输入/前置：重算审核副本；模板规则、定位单元格、原源路径、机构/期间元数据。
- **当前技术**：Calamine read_value_columns+Rust过滤/错误归一/Issue构造。
- 实际处理：读取真实保存结果，按配置识别异常，保留位置/公式/数值/说明和原源与审核副本链接。
- **主路径测试：R**。历史真实子集字段比对通过；合成重算后DIV0/REF整链有记录。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E06](#e06)、[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：读取不支持错误token或缺目标sheet/规则坐标错误：错误，GUI整批停止；这里未调用sidecar的现代错误token兼容读取。
- 可优化：统一错误token透传能力；公式文本与结果值分栏，不能将文本当计算结果。
- 需业务判断的场景：[S1F1-E04](#s1f1-e04)、[S1F1-E14](#s1f1-e14)
- 源码：[s1_formula_audit.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_formula_audit.rs)

#### S1F1.11 补扫校验区域中的公式错误

- 输入/前置：模板校验区域公式坐标；被扫描文件已存缓存。
- **当前技术**：Calamine+OOXML名称；S1ValidationFormulaScanPlan。
- 实际处理：补充规则抽取之外错误并去重。GUI扫重算副本；当前s1-audit-directory扫source_file，需业务确认不同口径。
- **主路径测试：U**。历史GUI/CLI合成覆盖有限；当前源缓存与重算副本差异分支无单独成功证据。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)、[E00](#e00)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：GUI加载/扫描失败会整批停止；CLI部分扫描错误构造成失败Issue保留，处理政策不同。
- 可优化：增加‘源无错误、重算副本有错’与反例的定向验收；明确GUI/CLI是否应同口径后再改。
- 需业务判断的场景：[S1F1-E15](#s1f1-e15)、[S1F1-E14](#s1f1-e14)
- 源码：[s1_validation_formula_scan.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_validation_formula_scan.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[open-data-audit.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/bin/open-data-audit.rs)

#### S1F1.12 读取并判断实际触发的条件格式

- 输入/前置：重算副本CF规则、AppliesTo、扫描区域、单元格缓存、结构/批注。
- **当前技术**：OOXML规则元数据+Calamine值+Rust有限CF表达式求值。
- 实际处理：支持已实现cellIs/expression子集，处理相对/绝对引用和优先级等已实现上下文；输出触发单元格/规则/数值/说明，不以填充颜色判断。
- **主路径测试：R**。历史CF真实子集与合成B2触发；不是Excel完整CF实现。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E05](#e05)、[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：存在不支持规则、求值错误、指标未解析或批注上下文未确认→completed_with_diagnostics；GUI要求completed，否则停止整批结果。
- 可优化：先提供逐规则诊断并让业务决定必审/提示；计算引擎重算不等于RustCF求值器支持全部规则。
- 需业务判断的场景：[S1F1-E16](#s1f1-e16)、[G09](#g09)
- 源码：[conditional_format.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/conditional_format.rs)、[s1_conditional_audit.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_conditional_audit.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S1F1.13 合流公式/CF异常并处理历史

- 输入/前置：每文件两类Issue；可选历史配置/历史审核记录。
- **当前技术**：Rust稳定合流/去重/历史匹配；Calamine ISO日期读。
- 实际处理：按源顺序整合，分类新增/延续/解决，承接人工说明；不回写原输入。
- **主路径测试：R**。历史分类字段报告通过；GUI配置/历史同簿复杂真实文件仍需验收。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E07](#e07)、[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：历史选择后读取失败、键/格式不合法：GUI返回失败；未选历史走空历史初始化属于可选功能，不是读失败回退。
- 可优化：把‘无历史’和‘历史无法读取’区分；重复Issue key和机构归属清晰提示。
- 需业务判断的场景：[S1F1-E17](#s1f1-e17)
- 源码：[s1_issue_chain.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_issue_chain.rs)、[s1_history.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_history.rs)

#### S1F1.14 输出13列结果及GUI路径

- 输入/前置：合流历史分类Issue、审核副本列表、引擎回执。
- **当前技术**：rust_xlsxwriter新建+RustWeb结果状态；输出目录guard。
- 实际处理：写既定结果与链接/计数，GUI显示输出路径及使用计算引擎；成功后保留run目录。
- **主路径测试：H**。旧Excel合成整链结果可重开、13列与3条异常记录；当前批次GUI人工操作未验收。 当前版本：当前工作区未重新编译或执行该步骤；历史证据不自动覆盖最新改动。 证据：[E02](#e02)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：该步骤失败
- 回退执行：按调用层返回错误，不换读取/修改库。
- **回退测试：NA**。当前没有可测的替代生产路径。
- **当前失败处理**：写结果失败/磁盘不足/被占用：整批失败并清理run目录；不是保留部分成功文件供继续。
- 可优化：需要用户决定是否允许部分成果发布；失败清单和已处理计数应同时保留，避免只留通用错误。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[S1F1-E18](#s1f1-e18)
- 源码：[s1_result_workbook.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_result_workbook.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S1-F2 汇总校验结果说明 / 命名区域汇总

#### S1F2.01 选择来源、模板与汇总模式

- 输入/前置：源目录/勾选、模板、映射/历史配置。
- **当前技术**：Rust fs + Web/CLI参数。
- 实际处理：选择任意行或固定行汇总功能，排除配置/模板/已有产物并规范化路径。
- **主路径测试：R**。历史合成CLI调用/输入保护；GUI逐按钮没有独立通过凭证。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：非法路径、模板无效、无来源→阻断本次汇总。
- 可优化：统一GUI/CLI已选择文件清单并在输出中列出实际参与来源。
- 需业务判断的场景：[G01](#g01)、[G02](#g02)、[G03](#g03)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S1F2.02 读模板值、名称、结构和附加字段

- 输入/前置：模板XLSX/1.0映射。
- **当前技术**：Calamine+OOXML定义名称/worksheet extents+Rust合并表头处理。
- 实际处理：解析数据区、表头区、附加字段及兜底模板配置，验证区域及输出列数。
- **主路径测试：R**。历史任意行/固定行、元数据字段合成通过。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：名称缺失/非法/不连续等不符既定契约→阻断；不能确认公式/合并状态也报错。
- 可优化：一次读模板并缓存范围；模板错误不能以空表头继续。
- 需业务判断的场景：[G04](#g04)、[G05](#g05)、[S1F2-E01](#s1f2-e01)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)、[lib.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-reader/src/lib.rs)

#### S1F2.03 表结构预检筛选文件

- 输入/前置：源表头/结构区域与模板。
- **当前技术**：复用S1结构比较；Calamine+OOXML。
- 实际处理：允许可用文件继续，记录不通过及无法比较的原因；全不通过报无来源。
- **主路径测试：R**。历史合成主链；手工样例另列header bad，未重新执行。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：预检失败跳过当前工作簿；全跳过阻断。
- 可优化：将跳过明细固定写入结果/日志，避免只显示最终行数。
- 需业务判断的场景：[S1F2-E02](#s1f2-e02)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)、[s1_structure_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_structure_compare.rs)

#### S1F2.04 正式读取源工作簿及公式缓存

- 输入/前置：通过预检的源XLSX。
- **当前技术**：Calamine WorkbookDump+OOXML实际范围。
- 实际处理：需要能确认公式文本/合并状态；有公式的取值区域读取已存缓存，禁止以公式字符串顶替值。
- **主路径测试：R**。历史合成/历史追加子集；真实批次另有partial，不能泛化。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)、[E09](#e09)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：正式read_workbook/extents失败或必要公式缓存不可确认→?返回，停止整个汇总。
- 可优化：如果业务允许单文件继续，需要显式隔离正式读取错误并标部分汇总；不偷偷回退到值重建。
- 需业务判断的场景：[G04](#g04)、[S1F2-E03](#s1f2-e03)、[S1F2-E04](#s1f2-e04)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)

#### S1F2.05 匹配同名sheet、执行普通/兜底汇总

- 输入/前置：源各sheet、模板对应数据区/表头；兜底模板区域。
- **当前技术**：Rust命名区解析/网格取值/表头合并/行筛选。
- 实际处理：普通同名sheet汇总；未匹配sheet且配置兜底模板时，把固定/任意行范围与局部字段投影到源sheet；同名结果合并。
- **主路径测试：R**。主汇总9行；fallback_template明确passed，任意行/固定行各4行，数值及表名差异0。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)。
- **回退类别/是否存在**：业务模板兜底；已有此限定分支。
- 回退触发：源sheet未匹配正常模板，但配置了有效‘兜底模板’数据区和表头区域。
- 回退执行：按兜底区域汇总源sheet；不换reader/writer。
- **回退测试：R**。E08已核fallback_template.status=passed，任意行与固定行两个分支通过；历史限定合成样本，当前未重新跑。
- **当前失败处理**：无对应sheet且无兜底→该sheet不参与并计数；空白行/无效候选按规则略过；无有效命名区/表头错误阻断。
- 可优化：这项可明确写‘历史回退测试通过’；超出范围/大表仍按契约校验，不能无模板猜表头。
- 需业务判断的场景：[S1F2-E05](#s1f2-e05)、[S1F2-E06](#s1f2-e06)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)

#### S1F2.06 历史匹配和说明字段追加

- 输入/前置：汇总行、配置/历史审核记录。
- **当前技术**：Calamine ISO日期读取+Rust历史键/字段整理。
- 实际处理：匹配既定机构、指标/位置等历史键，追加人工说明及历史字段；空附加元数据按规则留空并计数。
- **主路径测试：R**。历史命名区/历史追加合成报告passed（列宽差异另列）。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E09](#e09)、[E07](#e07)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：已选历史无法读→阻断；找不到匹配历史正常留空；不能把历史读失败当没有历史。
- 可优化：明确重复历史记录择取规则、日期文本/数值语义和来源。
- 需业务判断的场景：[S1F1-E17](#s1f1-e17)、[S1F2-E07](#s1f2-e07)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)、[s1_history.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_history.rs)

#### S1F2.07 新建结果并独立读回

- 输入/前置：分组数据、表头、来源字段及历史字段。
- **当前技术**：rust_xlsxwriter新建→Calamine validate_summary_workbook。
- 实际处理：逐组生成结果sheet/筛选/来源列，写临时文件，独立校验表名、维度和内容，再提交。
- **主路径测试：R**。历史输出字段和结果读回通过；普通列宽差异不升级业务范围。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)、[E09](#e09)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：XLSX行列超限、写盘、读回不符→删除临时文件并终止，无writer回退。
- 可优化：大汇总可评估constant_memory顺序写，先覆盖来源/历史列和读回验证。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G10](#g10)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)

#### S1F2.08 提交文件和GUI结果路径

- 输入/前置：已验证临时XLSX、目标文件。
- **当前技术**：std::fs rename + Web状态输出。
- 实际处理：成功后展示来源数/跳过数/结果sheet/行数及输出路径。
- **主路径测试：H**。CLI输出有证据；当前GUI交互完整人工验收未证。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E08](#e08)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：提交rename/目标占用失败→清理临时文件并终止；无另一目录自动继续。
- 可优化：若允许重试另存，必须用新路径且保留同一已验证产物，不重新计算制造不一致。
- 需业务判断的场景：[G06](#g06)
- 源码：[s1_region_summary.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_region_summary.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S2-F1 执行报表规则

#### S2F1.01 加载配置与输入检查

- 输入/前置：S2配置、两期目录、可选外部、输出。
- **当前技术**：Web/CLI+Calamine配置读取+Rust配置验证。
- 实际处理：解析机构/指标参照、环比、外部核对和校验规则、单位等既定字段。
- **主路径测试：R**。历史真实服务路径覆盖配置；本轮未执行配置检查程序。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E13](#e13)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：配置表头/单位/表达式定义不可用→阻断；未选择外部只跳过外部分支，不是读失败回退。
- 可优化：配置一次加载，GUI检查与执行共用同份解析快照。
- 需业务判断的场景：[G05](#g05)、[S2F1-E01](#s2f1-e01)
- 源码：[s2_config.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_config.rs)、[s2_config_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_config_check.rs)、[s2_service.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_service.rs)

#### S2F1.02 发现并读取两期报表

- 输入/前置：本/上期目录内.xls/.xlsx，前三列指标数据。
- **当前技术**：std::fs排序+Calamine read_value_columns(3)。
- 实际处理：按文件/表名识别机构、表单、日期；从第4行读取指标码、名称和数值，保留路径和原始行。
- **主路径测试：R**。两期真实子集有输入保护和输出比对；旧导入合成报告也存在。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：文件名不识别→warning跳过；已识别文件Calamine失败→整个服务返回错误；空目录阻断。
- 可优化：按文件统计‘未识别’与‘打不开’，可考虑业务授权后的坏文件隔离；不能无提示少算。
- 需业务判断的场景：[G04](#g04)、[S2F1-E02](#s2f1-e02)
- 源码：[s2_import.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_import.rs)

#### S2F1.03 机构/指标映射与单位归一

- 输入/前置：导入记录、机构/指标/表单配置。
- **当前技术**：Rust映射/DecimalTuple金额转换；ISO/原值语义。
- 实际处理：源身份优先，配置补缺；核对表单归属，保留异常文本/空值的既定语义并记录质量问题。
- **主路径测试：R**。历史合成单位与真实导入子集；不是任意Excel显示格式解析。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E14](#e14)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：机构/表单不一致、非数值、精度损失按warning/质量检查分支；不可转换业务项不等于0或通过。
- 可优化：统一Decimal到输出精度边界并保留源单位与归一值。
- 需业务判断的场景：[S2F1-E03](#s2f1-e03)、[G11](#g11)
- 源码：[s2_import.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_import.rs)、[s2_config.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_config.rs)

#### S2F1.04 逻辑key去重与数据质量检查

- 输入/前置：两期机构+表单/指标记录。
- **当前技术**：Rust有序容器+质量诊断。
- 实际处理：相同逻辑key按现实现保留第一条，记录冲突/重复；构建两期索引。
- **主路径测试：R**。历史样本契约覆盖部分重复/质量路径；任意跨文件冲突需定向样例。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E14](#e14)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：现为先到先用并提示，继续后续审核；用户需决定值不同时是否应阻断该机构或整批。
- 可优化：区分相同值重复与不同值冲突；输出所采用源路径和被忽略来源。
- 需业务判断的场景：[S2F1-E04](#s2f1-e04)
- 源码：[s2_import.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_import.rs)、[s2_validation.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_validation.rs)

#### S2F1.05 两期比较与环比规则

- 输入/前置：本期/上期索引、阈值和规则。
- **当前技术**：Rust直接数值计算与period/rule engine。
- 实际处理：配对、增减/环比/空值与零分母按原规则，输出触发/质量/取数问题。
- **主路径测试：R**。历史真实本上期结果一致；合成规则报告通过。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E14](#e14)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：缺上期/缺指标、异常数值按具体规则诊断/跳过，非外部Excel计算失败；不能记成未触发通过。
- 可优化：错误结果与未触发结果分开计数；保留缺值与0的区别。
- 需业务判断的场景：[S2F1-E05](#s2f1-e05)
- 源码：[s2_period.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_period.rs)、[s2_analysis.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_analysis.rs)、[s2_rule_engine.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_rule_engine.rs)

#### S2F1.06 校验表达式与规则计算

- 输入/前置：S2规则、指标字典和期间上下文。
- **当前技术**：Rust s2_expression / s2_rule_engine 有限语法。
- 实际处理：按既定前提/比较/规则逻辑取数并计算，生成规则结果；不注入Excel公式。
- **主路径测试：R**。合成语法和真实配置子集历史通过；无任意Excel函数承诺。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E14](#e14)、[E12](#e12)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：语法错误、除零/取数失败按配置校验/规则诊断路径；没有交COM/LO再算的回退。
- 可优化：缓存编译后的表达式；启动前列出不支持函数及具体rule id。
- 需业务判断的场景：[S2F1-E06](#s2f1-e06)、[G12](#g12)
- 源码：[s2_expression.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_expression.rs)、[s2_rule_engine.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_rule_engine.rs)

#### S2F1.07 读可选外部集中系统数据

- 输入/前置：选中的外部XLSX、集中系统数据sheet、机构参照/旧参照表、静态/时序区域。
- **当前技术**：Calamine read_value_columns_with_iso_dates + Rust外部布局解析。
- 实际处理：按外部配置定位机构、指标、数据期和单位，建立外部索引；只支持已确认XLSX映射。
- **主路径测试：R**。真实外部非时序子集通过；实际时序全覆盖未证明。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E13](#e13)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：未选外部→显式不执行；选择但损坏/缺关键sheet/格式非XLSX→整个服务失败；无OOXML/openpyxl读取回退。
- 可优化：验证所选外部布局并预览参与机构/数据期；时序未找到与未装引擎不要混淆。
- 需业务判断的场景：[S2F1-E07](#s2f1-e07)、[G04](#g04)
- 源码：[s2_external_import.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_external_import.rs)

#### S2F1.08 外部核对规则

- 输入/前置：本期规范值、外部索引、机构映射与规则。
- **当前技术**：Rust s2_external+规则引擎。
- 实际处理：做单位、机构/指标/日期映射并计算差值，输出16列外部核对结果。
- **主路径测试：R**。历史所选外部真实数据工作簿差异0；缺映射/歧义覆盖有限。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E13](#e13)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：外部缺机构/指标或多值歧义按现规则诊断，不能用第一个值假装匹配。
- 可优化：在结果保留外部sheet/坐标、单位、映射来源和失败粒度。
- 需业务判断的场景：[S2F1-E08](#s2f1-e08)
- 源码：[s2_external.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_external.rs)、[s2_rule_engine.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_rule_engine.rs)

#### S2F1.09 写三表并独立读回

- 输入/前置：环比/外部/校验结果、诊断。
- **当前技术**：rust_xlsxwriter→Calamine独立验证。
- 实际处理：新建三sheet（16/16/12列），按结果类型写数值/文本并读回核对。
- **主路径测试：R**。历史真实无外部/有外部子集输出通过；不是现有工作簿保真修改。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E13](#e13)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：写失败/行数超限/读回不符→停止输出；无备用writer或Umya重建。
- 可优化：大结果顺序写与共享格式复用；先量内存，不因库切换重复实现。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G10](#g10)
- 源码：[s2_exporter.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_exporter.rs)

#### S2F1.10 进度与输出路径

- 输入/前置：服务结果计数、warnings、output_path。
- **当前技术**：Rust服务回调+Web/API/CLI JSON。
- 实际处理：展示耗时、引擎阶段、三类结果和输出路径；业务触发异常与软件失败区分。
- **主路径测试：H**。历史CLI/service通过；当前真实Windows GUI人工验收未证。 当前版本：当前工作区未重新执行本步骤；历史证据只限所述范围。 证据：[E12](#e12)、[E13](#e13)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或按明确业务跳过；不切换读取/写入库。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：前序服务错误终止当前功能；普通质量提示可继续，UI不能把结果文件存在当全部规则已审。
- 可优化：将规则失败数、输入跳过数和异常触发数分开。
- 需业务判断的场景：[G13](#g13)
- 源码：[s2_service.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s2_service.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S3-F1 两期数据→比较文件

#### S3F1.01 输入与配置读取

- 输入/前置：两期文件、3.0/3.1工作簿、输出路径。
- **当前技术**：Rust Web/CLI参数+Calamine读取配置。
- 实际处理：3.0运行参数/默认源与目标单位、警戒、例外；3.1指标参照、动作、5段/8段兼容表达式。
- **主路径测试：R**。历史真实配置与CLI链子集通过。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：配置路径/字段/单位无法解释→停止本功能；GUI/CLI输入数量能力有差别。
- 可优化：解析后的3.0/3.1快照绑定到本次run，写入配置hash/版本。
- 需业务判断的场景：[G05](#g05)、[S3F1-E01](#s3f1-e01)
- 源码：[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S3F1.02 CSV编码和记录读取

- 输入/前置：CSV原始字节、14列既定表头。
- **当前技术**：Rust csv+encoding_rs::GBK+UTF8解码。
- 实际处理：继承旧流程解码顺序：严格GBK→UTF8/BOM→GBK替换解码；兼容既定换行与CSV引号。
- **主路径测试：R**。真实CSV子集通过；三种解码分支未分别定位成功凭证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：编码兼容回退；已有此限定分支。
- 回退触发：GBK严格解码失败，再尝试UTF8；两者都失败时GBK替换。
- 回退执行：只改变字节解码方式，仍走同一业务14列解析；不是Excel读取回退。
- **回退测试：N**。编码分支测试源码存在，历史真实CSV总结果通过；没有三分支各自执行成功的独立凭证，替换后业务数据正确也未证。
- **当前失败处理**：无法按14列表头识别则停止；末级替换可能产生乱码但继续解析。
- 可优化：回执记录最终编码和替换字符数；用户决定出现替换是否阻断，不能把‘能解码’等同内容正确。
- 需业务判断的场景：[S3F1-E02](#s3f1-e02)、[S3F1-E03](#s3f1-e03)
- 源码：[central_csv.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/central_csv.rs)

#### S3F1.03 工作簿读取与统一输入记录

- 输入/前置：XLSX/XLS工作表14列值、日期/数值。
- **当前技术**：Calamine ISO日期适配+同一central normalization。
- 实际处理：工作簿与CSV汇合到CentralDataset；识别数据日期、机构/地区/业务/指标/属性/币种/频度/批次。
- **主路径测试：H**。历史工作簿入口用例存在，F4有专门工作簿报告；不把F4覆盖直接当F1全部格式通过。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)、[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：损坏/加密/未知类型、14列不符、同文件日期不一致→停止；异常数据行按ImportIssue处理。
- 可优化：保留源行和数据类型，避免代码前导0丢失；限制所需列。
- 需业务判断的场景：[G04](#g04)、[S3F1-E04](#s3f1-e04)、[S3F1-E05](#s3f1-e05)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)、[central_csv.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/central_csv.rs)

#### S3F1.04 规范化、单位、重复key处理

- 输入/前置：两期CentralDataset+3.0单位与例外。
- **当前技术**：Rust数值换算/字典索引；当前CentralValue用f64。
- 实际处理：按八维业务key建索引；旧字典语义重复key后条覆盖；按默认单位/指标例外换算。
- **主路径测试：R**。历史真实子集通过；全部重复/代码碰撞/尾差场景未证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：异常文本保留/记录，部分无效行略过；重复key后值覆盖继续；未知单位全局错误停止。
- 可优化：给覆盖行为输出来源与不同值冲突；单位和舍入可复用现有Decimal边界，不能直接更改旧精度规则。
- 需业务判断的场景：[S3F1-E06](#s3f1-e06)、[G11](#g11)
- 源码：[central_csv.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/central_csv.rs)、[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F1.05 两期配对、增减和环比

- 输入/前置：本期/上期八维索引、单位精度。
- **当前技术**：Rust build_base_comparison。
- 实际处理：构造两期比较、仅本期/仅上期记录、增减与零分母等既定分支；不需要Office。
- **主路径测试：R**。历史真实数据和1行合成烟测；单行不是所有边界。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)、[E30](#e30)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：缺一时期/异常值按业务默认分支输出或提示；期间级不合法停止。
- 可优化：逐项标记缺失状态，不能把缺数据隐式0与真实0混合统计。
- 需业务判断的场景：[S3F1-E07](#s3f1-e07)
- 源码：[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F1.06 环比警戒和动作规则

- 输入/前置：3.0警戒、3.1动作规则、比较记录。
- **当前技术**：Rust apply_action_rules+阈值/范围匹配。
- 实际处理：生成规则提示、是否说明等业务字段；不自动替代人工‘输出说明’勾选。
- **主路径测试：R**。历史动作/真实比较子集。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：配置作用域/规则非法可停止加载；单行条件不命中正常继续，执行诊断按日志分支。
- 可优化：规则id/触发阈值/采用单位写入日志便于业务追溯。
- 需业务判断的场景：[G12](#g12)、[S3F1-E08](#s3f1-e08)
- 源码：[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F1.07 5段/8段表达式校验

- 输入/前置：比较记录、3.1表达式及机构/地区范围。
- **当前技术**：Rust有限表达式与scope匹配。
- 实际处理：计算前提、左右表达式和校验条件，输出异常及日志；非Excel公式通用计算。
- **主路径测试：R**。历史真实配置表达式子集；不支持语法的全量覆盖未证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：表达式/取数错误按具体报告与日志处理；不会回退到Python eval/Excel/LO。
- 可优化：显式统计未执行规则数；配置检查与实际运行共享解析结果。
- 需业务判断的场景：[G12](#g12)、[S3F1-E08](#s3f1-e08)
- 源码：[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)、[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)

#### S3F1.08 新建22列比较文件与可选日志

- 输入/前置：比较结果、数据期、执行日志。
- **当前技术**：rust_xlsxwriter+Calamine读回验证。
- 实际处理：保持21业务列和人工输出说明列，附执行日志文件（若请求）；不改输入。
- **主路径测试：R**。历史真实结果结构与源保护；旧浮点尾差需按样本限制解释。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E15](#e15)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：写盘/读回不符/目标路径问题停止；日志若在结果后写失败可能已有结果文件，需要避免假整体成功。
- 可优化：结果和可选日志共同暂存再提交，产物清单注明主报告与日志状态。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G10](#g10)、[S3F1-E09](#s3f1-e09)
- 源码：[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F1.09 输出路径与下游交接

- 输入/前置：比较文件、日志与计数。
- **当前技术**：Web/API/CLI JSON+既有HTML。
- 实际处理：F2按人工输出说明筛选；F5使用同一比较结果；F3不直接回写本比较簿。
- **主路径测试：H**。历史CLI/API烟测有记录；当前实际GUI五项连续操作未验收。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E30](#e30)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：界面状态不替代结果有效性；旧编号或默认路径错误可能选错下游输入。
- 可优化：在文件输出元信息保留feature_id、输入hash和产物类型，避免凭文件名误用。
- 需业务判断的场景：[G13](#g13)、[S3-E01](#s3-e01)
- 源码：[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S3-F2 比较文件→说明文件

#### S3F2.01 读取比较文件和机构地区配置

- 输入/前置：比较XLSX、可选3.0。
- **当前技术**：Calamine普通值+Rust列识别/参照加载。
- 实际处理：查找比较结果表及必需列，加载已选择的机构地区参照。
- **主路径测试：R**。历史真实2610行场景输入/配置未变。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：缺列/没有合格sheet/已选配置读失败→停止；未选配置使用既有输入字段，不等于读错后降级。
- 可优化：输出实际采用的列映射和参照来源。
- 需业务判断的场景：[G04](#g04)、[G05](#g05)、[S3F2-E01](#s3f2-e01)
- 源码：[s3_explanation_export.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_export.rs)

#### S3F2.02 按人工输出说明筛选

- 输入/前置：22列中的输出说明及人工真值标记。
- **当前技术**：Rust真值/文本兼容规则。
- 实际处理：只导出人工标记行；保留旧Python接受的既定真值，不用规则‘是否说明’自动决定。
- **主路径测试：R**。历史Python/Rust各选2610行；另E30为0行烟测。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：未选择行正常跳过；全部未选择当前可完成0行说明，不自动替用户选行。
- 可优化：在执行前/结果中显示选中数；是否0行阻断由用户决定。
- 需业务判断的场景：[S3F2-E02](#s3f2-e02)
- 源码：[s3_explanation_export.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_export.rs)

#### S3F2.03 机构/地区映射及多地区拆分

- 输入/前置：机构代码名称/地区、所选记录。
- **当前技术**：Rust参照字典、代码正规化/多地区组织。
- 实际处理：按原业务规则展示机构地区及多地区机构标识，不重写F1原件。
- **主路径测试：R**。历史真实所选机构/地区映射子集。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：缺映射沿既有字段/规则处理；存在歧义不能默认为任意另一机构，具体需带诊断验收。
- 可优化：映射缺失/歧义数量独立统计，不以名称相似猜代码。
- 需业务判断的场景：[S3F2-E03](#s3f2-e03)
- 源码：[s3_explanation_export.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_export.rs)

#### S3F2.04 整理8个说明字段

- 输入/前置：机构/指标、数据值、上期、增减、原因等。
- **当前技术**：Rust结果投影/数值与代码类型处理。
- 实际处理：形成说明文件既定8列：机构代码、机构名称、指标代码、指标名称、数据值、上期值、增减额、变动原因。
- **主路径测试：R**。2610行字段比对通过；名称/金额极端类型未全部覆盖。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：格式异常/非法数值按现读取语义；不得为填满表格把异常文本静默变0。
- 可优化：为F3四字段匹配保留可重复的数值/代码规范，不只保留显示格式。
- 需业务判断的场景：[G11](#g11)、[S3F3-E02](#s3f3-e02)
- 源码：[s3_explanation_export.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_export.rs)

#### S3F2.05 输出新说明文件及生成信息

- 输入/前置：说明行、批次与来源。
- **当前技术**：rust_xlsxwriter新建+独立Calamine读回。
- 实际处理：写说明文件和生成信息两sheet，保护比较/配置输入。
- **主路径测试：R**。历史非空2610行和0行产物路径均有报告，不能互相替代范围。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)、[E30](#e30)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：写失败/读回不符→停止；无Umya/openpyxl新建后备。
- 可优化：新建报表顺序写与共用格式；原件安全无需引入Python。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G10](#g10)
- 源码：[s3_explanation_export.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_export.rs)

#### S3F2.06 交给机构填写并供F3导入

- 输入/前置：生成的说明文件。
- **当前技术**：现有GUI结果路径、文件系统。
- 实际处理：用户人工分发/填写原因，再选择F3目标和机构反馈；本程序不会自动发送他人。
- **主路径测试：H**。文件生成有证据；人工填写再回导整个真实流程未验收。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E16](#e16)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：机构改了匹配字段可能F3未匹配；只改变文件名不改变四字段键。
- 可优化：生成信息提示四个匹配字段应保留，并在F3预览未匹配原因。
- 需业务判断的场景：[S3-E01](#s3-e01)、[S3F3-E02](#s3f3-e02)
- 源码：[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S3-F3 机构说明→说明文件（先备份、原地更新）

#### S3F3.01 读取目标和机构反馈

- 输入/前置：目标.xlsx；反馈.xlsx/.xlsm；输出计划/冲突位置。
- **当前技术**：Calamine read_value_columns_with_iso_dates+fs路径/格式检查。
- 实际处理：目标只允许可编辑XLSX；反馈XLSM只读；识别说明sheet和必需表头，hash保护所有输入。
- **主路径测试：R**。历史计划合成读入通过；特殊类型/真实机构改表头仍待验收。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E17](#e17)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：目标XLSM/反馈格式不支持、缺机构代码等表头、任一已选文件读失败→停止本次；无读取回退。
- 可优化：允许坏反馈隔离需要另设业务策略，不能默默不导入。
- 需业务判断的场景：[G04](#g04)、[G05](#g05)、[S3F3-E01](#s3f3-e01)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.02 建立四字段匹配键

- 输入/前置：目标/反馈机构代码、指标代码、数据值、增减额；原因说明。
- **当前技术**：Rust规范化+HashMap/BTreeMap。
- 实际处理：严格按四字段身份匹配；机构/指标名称不是key。清理既定代码引号/.0，金额按已实现规范；空key/空原因不导入。
- **主路径测试：R**。历史计划单元格与冲突字段比对通过。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E17](#e17)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：金额改动/四字段不同→未匹配跳过；名称变化并不能补救键不一致；不能猜测同机构同指标就写入。
- 可优化：预览具体哪个key字段不一致；规范化必须与F2输出一致。
- 需业务判断的场景：[S3F3-E02](#s3f3-e02)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.03 去重反馈、冲突隔离及导入计划

- 输入/前置：同key反馈原因集合、目标匹配行。
- **当前技术**：Rust业务合并+ExcelEditPlan(SetCellValue)。
- 实际处理：同原因多份合并；多个原因或多个目标行→冲突；未匹配计数；合法唯一项计划写入‘变动原因’列。
- **主路径测试：R**。历史2条冲突、计划坐标/写入值已比对；不证明实际修改。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E17](#e17)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：冲突行跳过且出报告；合法行继续，现政策不因某行冲突整批停止。
- 可优化：决定目标已有非空原因被新唯一原因替换是否要二次冲突策略；报告保留旧/新值。
- 需业务判断的场景：[S3F3-E03](#s3f3-e03)、[S3F3-E04](#s3f3-e04)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.04 输出只读计划和冲突报告

- 输入/前置：ExcelEditPlan、冲突明细。
- **当前技术**：Serde JSON+rust_xlsxwriter新报告。
- 实际处理：生成导入计划/冲突XLSX，保持生产目标尚未变动；计划步骤本身不执行编辑。
- **主路径测试：R**。旧合成计划status=passed，但excel_modifier_execution=not run。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E17](#e17)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：计划/冲突文件已存在或写不出→停止，不能把计划成功显示为导入成功。
- 可优化：GUI分别列‘匹配完成’和‘原地写回完成’，精确给出跳过数。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G13](#g13)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.05 暂存编辑并校验来源未变化

- 输入/前置：目标快照、合法SetCellValue批次。
- **当前技术**：ExcelModifier→Umya；SHA256+独立结果验证。
- 实际处理：在独立暂存候选执行所有合法‘变动原因’写入，验证关键保留项，校验计划到提交间目标/反馈hash未变。
- **主路径测试：R**。Windows旧独立副本能力样本通过；Linux原地业务使用的低层编辑通过；当前Windows完整事务未证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E03](#e03)、[E04](#e04)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：无后端、对象不支持、写后验证错、hash变化→停止且不替换原目标。
- 可优化：同一事务一个open/save；来源变化不能用重试编辑掩盖。
- 需业务判断的场景：[G07](#g07)、[G08](#g08)、[G14](#g14)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)、[excel_route.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/excel_route.rs)、[umya_backend.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-editor/src/umya_backend.rs)

#### S3F3.06 创建目标旁备份并核验

- 输入/前置：目标原始字节、权限。
- **当前技术**：Rust fs create_new/copy/sync+SHA256。
- 实际处理：在替换前创建‘导入前备份’新路径并校验与原目标字节一致，避免覆盖旧备份。
- **主路径测试：R**。Linux能力报告明确字节一致备份；Windows原地备份链缺独立成功报告。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E04](#e04)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：备份目录无权限/空间不足/校验不符→停止，不改原目标；没有‘不备份继续’回退。
- 可优化：保留备份路径、hash、生成时间；备份不是业务选择的可忽略项（用户已明确要求）。
- 需业务判断的场景：[S3F3-E05](#s3f3-e05)、[G02](#g02)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.07 原地替换目标与失败回滚

- 输入/前置：验证候选、目标、同目录rollback暂存、备份。
- **当前技术**：Rust rename事务：target→rollback，candidate→target；失败还原。
- 实际处理：成功替换仍保留旁备份；安装候选失败尽力把rollback改回目标；恢复也失败时明确恢复路径。
- **主路径测试：H**。Linux记录说明包含回滚契约测试；JSON记录原地成功但没有故障注入全过程。Windows当前锁文件/恢复失败未证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E04](#e04)、[E02](#e02)。
- **回退类别/是否存在**：文件事务回滚；已有此限定分支。
- 回退触发：已把原目标移到rollback后，新候选安装失败。
- 回退执行：把rollback恢复到原路径；保留备份。不切换Excel引擎，也不是重新导入成功。
- **回退测试：H**。failed_candidate_install_restores_the_target_and_retains_its_backup测试源码存在，E02历史记录称Linux回滚覆盖；未找到Windows当前版本独立故障回执。
- **当前失败处理**：target被锁→替换前失败；候选安装失败→尝试恢复；恢复失败→明确停止并保留备份/rollback供人工恢复。
- 可优化：增加Windows目标锁、候选rename失败、还原失败定向故障记录；避免把回滚成功记成反馈导入成功。
- 需业务判断的场景：[S3F3-E06](#s3f3-e06)、[S3F3-E07](#s3f3-e07)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)

#### S3F3.08 备份/冲突/写回状态交付

- 输入/前置：更新目标、备份、计划和冲突报告、清理警告。
- **当前技术**：Web/CLI报告+fs清理。
- 实际处理：明确目标原地更新的路径与条数；rollback清理失败仅提示，备份继续保留；旧_to_copy函数保留但不作生产入口。
- **主路径测试：H**。Linux原地报告有路径；当前WindowsGUI逐步验收未证。 当前版本：未在当前工作区重新执行；不把历史子集当当前完整验收。 证据：[E04](#e04)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确错误/规则日志返回，不切换读取或写入引擎。
- **回退测试：NA**。没有实际生产替代路径。
- **当前失败处理**：导入已成功但清理rollback失败→warning；不得因此重试把原因重复写或删除备份。
- 可优化：输出状态拆成导入、备份、回滚、清理四项，便于人工判断。
- 需业务判断的场景：[G13](#g13)、[S3F3-E08](#s3f3-e08)
- 源码：[s3_explanation_feedback.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_explanation_feedback.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

### S3-F4 本期数值核对

#### S3F4.01 读3.0/3.2和输入参数

- 输入/前置：本期（CLI可选上期）、通用配置/本期数值核对配置。
- **当前技术**：Calamine配置+Rust Web/CLI。
- 实际处理：加载运行单位/例外、校验规则、启用范围/前提/左右表达式/校验公式、输出控制。
- **主路径测试：R**。历史真实163启用规则；GUI当前主要本期，CLI前后期扩展应分开验收。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：配置缺表/字段/单位不支持、必须输入不存在→停止；不是F1/F2的替代编号。
- 可优化：配置检查和执行使用同一规则筛选快照。
- 需业务判断的场景：[G05](#g05)、[S3F4-E01](#s3f4-e01)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)

#### S3F4.02 读CSV或工作簿并规范化

- 输入/前置：14列本期数据；可选上期。
- **当前技术**：CSV+encoding_rs或Calamine ISO日期适配；central_csv。
- 实际处理：复用S3输入契约、单位换算、机构/地区/指标索引；保留ImportIssue。
- **主路径测试：R**。真实CSV结果通过；工作簿报告另有样本限制，不证明所有Excel类型。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：CSV编码兼容；已有此限定分支。
- 回退触发：CSV严格GBK失败/UTF8失败。
- 回退执行：与F1相同GBK→UTF8→GBK替换；工作簿不走此分支。
- **回退测试：N**。历史整条输入子集通过，但未定位每种编码回退的单独成功证明。
- **当前失败处理**：表头不符/同文件日期不一致、DurationIso等不支持数据类型→停止；异常行记录后略过。
- 可优化：返回采用编码/替换计数、无效行及实际参与记录数。
- 需业务判断的场景：[S3F1-E02](#s3f1-e02)、[S3F1-E03](#s3f1-e03)、[S3F1-E04](#s3f1-e04)、[S3F1-E05](#s3f1-e05)、[S3F4-E02](#s3f4-e02)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)、[central_csv.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/central_csv.rs)

#### S3F4.03 按机构/地区/规则范围筛选

- 输入/前置：规则作用域、数据期、机构地区索引。
- **当前技术**：Rust正则/范围匹配与规则启用筛选。
- 实际处理：仅适用规则执行；前提为假正常略过；无数据或配置关闭与执行失败分别统计。
- **主路径测试：R**。历史作用域/规则输出子集；模式极端/旧Python正则差异需按配置验收。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：不支持/错误scope模式不能假装匹配全部；当前日志/配置错误按实现路径报告。
- 可优化：为每条规则给出未适用、前提假、取数失败三种原因。
- 需业务判断的场景：[G12](#g12)、[S3F4-E03](#s3f4-e03)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)

#### S3F4.04 前提、左右表达式取数和计算

- 输入/前置：指标索引、表达式、Thd等变量。
- **当前技术**：Rust calculate_expression /有限解析器。
- 实际处理：先算前提，再左右值；计算增减、绝对差、幅度；近零差按现epsilon归零，双0按旧规则不出行。
- **主路径测试：R**。历史214行/3日志通过，日志意味着部分规则问题按既定逻辑保留。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：前提/左右表达式失败→写日志后continue当前规则；不调用COM/LO补算，也不当通过。
- 可优化：将‘规则跳过因错误’计入不完整状态，不能只看214行结果。
- 需业务判断的场景：[S3F4-E04](#s3f4-e04)、[G11](#g11)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)

#### S3F4.05 校验公式、反转及成功抑制

- 输入/前置：left/right、校验公式、逻辑/反转/不出成功结果标记。
- **当前技术**：Rust evaluate_expression+业务分类。
- 实际处理：计算有限公式/真假，结合反转与抑制决定出行，区分逻辑/数值校验；公式错误写入诊断和结果。
- **主路径测试：R**。真实规则子集通过；全部语法/多错误组合未证。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：公式解析/除零等→记录formula_error/日志，部分仍输出带错误行；不能把None解释为校验成功。
- 可优化：沿用现DSL而非增加Excel公式引擎；前置校验能发现的语法问题提前定位。
- 需业务判断的场景：[S3F4-E05](#s3f4-e05)、[G12](#g12)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)

#### S3F4.06 输出17列结果、日志与读回

- 输入/前置：核对结果/日志、输出文件。
- **当前技术**：rust_xlsxwriter+Calamine独立验证。
- 实际处理：新建对比结果17列和可选日志文件，保留公式文本用于解释，数值由Rust已计算。
- **主路径测试：R**。历史结果214行、日志3行、差异0；不是Office公式重算。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：保存/读回失败停止；如主结果已存后日志失败，可能留下可用主文件，应明确状态。
- 可优化：两个文件同批暂存/统一产物清单；大文件顺序写评估。
- 需业务判断的场景：[G02](#g02)、[G06](#g06)、[G10](#g10)、[S3F1-E09](#s3f1-e09)
- 源码：[s3_current_value_check.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_current_value_check.rs)

#### S3F4.07 GUI/CLI完成信息

- 输入/前置：行数、日志数、输出与耗时。
- **当前技术**：现有HTML+Web/CLI结果。
- 实际处理：展示本期核对结果与日志路径，避免因输出不为空称所有规则已核对。
- **主路径测试：H**。CLI历史真实结果有凭证；当前GUI真实输入人工验收未证。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E18](#e18)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：前序错误停止；含规则错误的输出可能仍完成，需要用户选择整体部分完成/阻断政策。
- 可优化：在GUI显式列执行失败/跳过规则数量。
- 需业务判断的场景：[G13](#g13)、[S3F4-E06](#s3f4-e06)
- 源码：[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S3-F5 比较文件→金融表单

#### S3F5.01 读比较结果

- 输入/前置：F1比较XLSX的对比结果sheet。
- **当前技术**：Calamine read_value_columns_with_iso_dates。
- 实际处理：至少15列；读取前15个既定业务位置、按机构代码+地区分组；相同指标映射key当前先条保留。
- **主路径测试：H**。合成F5CLI/API有记录；真实表头变动/重复key未证。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E19](#e19)。
- **回退类别/是否存在**：位置兼容兜底；已有此限定分支。
- 回退触发：前15列表头与标准不完全一致，但列数足够。
- 回退执行：警告后按固定列位置取值；不是安全列重排或reader切换。
- **回退测试：N**。未找到表头错位情况下兜底正确输出的独立成功证据；位置改变可能误读，需业务决定是否应阻断。
- **当前失败处理**：无sheet/列少于15→阻断；表头不一致当前warning后按旧固定位置继续。
- 可优化：优先核定按列名解析/严格模板契约；仅改显示名与真正列移位应区分。
- 需业务判断的场景：[G04](#g04)、[S3F5-E01](#s3f5-e01)、[S3F5-E02](#s3f5-e02)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)

#### S3F5.02 读3.3报表清单、模板与设置

- 输入/前置：3.3XLSX中的报表清单、转表设置、启用金融表sheet。
- **当前技术**：Calamine值+OOXML实际extents+Rust配置解析。
- 实际处理：加载表单启用、频度/批次、映射、空行隐藏/空表删除。3.3‘分析文件数据单位’按用户确认忽略；F5承接F1按3.0默认目标单位生成的值，不进行另一次独立转换。
- **主路径测试：H**。合成配置/表单编辑通过；仓库占位配置不能证明真实业务43条启用配置。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E19](#e19)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：缺报表清单/无有效启用sheet/模板结构不符→阻断；不能把空占位当正常无数据。
- 可优化：执行前展示启用表单数与单位来源；检查比较结果是否由相同3.0目标单位产生，当前未见全面单位校验。
- 需业务判断的场景：[G05](#g05)、[S3F5-E03](#s3f5-e03)、[S3F5-E04](#s3f5-e04)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)、[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F5.03 读3.0警戒并建立机构/指标计划

- 输入/前置：3.0环比预警区间、比较值、3.3映射。
- **当前技术**：Rust配对/指标映射/值投影。
- 实际处理：按业务类/指标/属性/币种/频度/批次定位金融表格本上期值；准备增减/警戒等数值和标色，不运行模板Excel公式。
- **主路径测试：H**。合成F5逻辑/编辑子集；真实复杂映射未全证。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E19](#e19)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：缺映射或重复指标按现warning/缺值分支；不同机构同输出名提前阻断；没有Python业务回退。
- 可优化：每机构报告映射成功/未找到/重复冲突数量，避免表单空白而显示成功。
- 需业务判断的场景：[S3F5-E02](#s3f5-e02)、[S3F5-E05](#s3f5-e05)、[G11](#g11)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)、[s3_execution_compare.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_execution_compare.rs)

#### S3F5.04 规划删列/删表/隐藏行及写值

- 输入/前置：3.3模板header markers、机构数据、启用表单。
- **当前技术**：Rust S3IndicatorSplitEdit→ExcelEditPlan。
- 实际处理：按既定规则删除占位/无关列与空表、清理配置sheet、填文字/数值/数字格式/警戒填充、隐藏空行。
- **主路径测试：R**。能力样本验证列移位、公式/合并/批注/CF/DV/尺寸/页面等；不等于任意真实模板。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E03](#e03)、[E04](#e04)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：删完无有效表单/维度或引用不可确认→阻断；未知表单对象不许悄悄丢弃。
- 可优化：保持一个机构一整批编辑；普通字体边框不影响范围，CF判断相关保留仍必要。
- 需业务判断的场景：[S3F5-E06](#s3f5-e06)、[G07](#g07)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)

#### S3F5.05 文件对象预检和编辑路由

- 输入/前置：模板profile、操作/保留清单、平台。
- **当前技术**：OOXML profile+ExcelModifier能力筛选→Umya。
- 实际处理：生产仅合格Windows/Linux x86_64 Umya；保证本次整批操作与保留项都受支持。
- **主路径测试：R**。受限能力样本通过；不是所有对象全保真。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E03](#e03)、[E04](#e04)。
- **回退类别/是否存在**：框架可换后端但生产无第二后端；当前没有实际可用的生产替代引擎/路径。
- 回退触发：写前不支持且另有同契约合格后端。
- 回退执行：当前没有生产openpyxl候选可用；失败结束本次机构及整批。
- **回退测试：NA**。没有真实Umya→openpyxl F5回退成功证据；直接openpyxl删列也不自动维护所有依赖。
- **当前失败处理**：DeleteColumn且存在defined names明确写前拒绝；图表/透视/VBA/externalLinks等未覆盖对象拒绝，当前无后备。
- 可优化：先用真实3.3样本确认是否触发definedNames限制；不能删掉名称来绕过，也不能把普通openpyxl当万能后备。
- 需业务判断的场景：[S3F5-E07](#s3f5-e07)、[G07](#g07)、[G08](#g08)
- 源码：[excel_route.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/excel_route.rs)、[umya_backend.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-editor/src/umya_backend.rs)

#### S3F5.06 每机构暂存编辑和验证

- 输入/前置：模板副本、各机构ExcelEditPlan。
- **当前技术**：ExcelModifier/Umya+OOXML/Calamine保留及值验证。
- 实际处理：一机构一输出，原模板/比较/配置只读；写完校验需要保留对象和目标单元格。
- **主路径测试：R**。历史合成编辑与输入保护、CLI/API通过；当前真实全量表单待验收。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E03](#e03)、[E04](#e04)、[E19](#e19)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：任一机构编辑/读回失败→函数返回，暂存guard清理整个批次；不继续其他机构并发布部分。
- 可优化：用户若选‘隔离机构’需增加可靠失败清单和部分状态；当前不能描述成已支持。
- 需业务判断的场景：[S3F5-E08](#s3f5-e08)、[G02](#g02)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)、[umya_backend.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-editor/src/umya_backend.rs)

#### S3F5.07 全部成功后提交目录

- 输入/前置：全部已验证机构XLSX、输出根目录。
- **当前技术**：Rust同根staging目录+rename发布。
- 实际处理：全部完成才发布指标比较拆分目录；命名冲突/不安全文件名提前检查。
- **主路径测试：H**。合成成功发布/失败路径有历史记录；磁盘/占用/中断故障无独立证据。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E19](#e19)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：提交失败→不报告成功并尽力清理暂存；目标已存在不覆盖。
- 可优化：发布回执列每个文件hash/机构和业务状态；失败也保留不含敏感值的清单。
- 需业务判断的场景：[S3F5-E08](#s3f5-e08)、[G06](#g06)
- 源码：[s3_indicator_split.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s3_indicator_split.rs)

#### S3F5.08 GUI/CLI交付文件目录

- 输入/前置：各机构文件、警告、匹配计数和编辑器信息。
- **当前技术**：Web/API/CLI结果+现有HTML。
- 实际处理：展示机构数/表单数/输出目录及警告；与旧错误编号对应但保留有效业务实现。
- **主路径测试：H**。历史合成HTTP/API/CLI通过；用户真实配置逐机构人工验收未证。 当前版本：当前工作区未重新编译/运行；历史证据不等于当前完整验收。 证据：[E19](#e19)。
- **回退类别/是否存在**：无生产回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：返回错误或明确规则日志，不换业务/读写引擎。
- **回退测试：NA**。无实际生产替代路径。
- **当前失败处理**：输出存在但映射警告/表单空白仍可能不完整，GUI不能统一说业务通过。
- 可优化：提供每机构产物摘要便于抽查；不扩展普通样式逐项对拍。
- 需业务判断的场景：[G13](#g13)、[S3F5-E09](#s3f5-e09)
- 源码：[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

### S5-F1 配置模式汇总核查表校验

#### S5F1.01 选择本/上期、模板、配置与外部输入

- 输入/前置：GUI源目录/勾选，模板和业务/报告期，上期可选，单外部。
- **当前技术**：现有HTML+Rust Web/CLI/fs。
- 实际处理：校验物理文件与格式；GUI约束同模板同业务同期间，CLI多源须显式绑定；检查配置与识别由用户触发。
- **主路径测试：U**。旧CLI单源合成completed；新GUI批次路径未运行，不能写GUI验收通过。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E30](#e30)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：任一已选路径不存在/全局配置不合法/上期选了但未填期间→开始前整批阻断。
- 可优化：保留勾选清单与排除原因；不要让‘配置检查按钮通过’代替本次输入预检。
- 需业务判断的场景：[G01](#g01)、[G03](#g03)、[S5F1-E01](#s5f1-e01)
- 源码：[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[web_bridge.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/web_bridge.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

#### S5F1.02 模板匹配、业务和数据期识别

- 输入/前置：源sheet集合、模板目录、S5业务/期间来源配置。
- **当前技术**：专用OOXML sheet目录+Rust模板/业务/期间识别；需要时Calamine值。
- 实际处理：先识别模板再确定业务和唯一报告期；显式模板优先；解析期别不能猜多个冲突日期。
- **主路径测试：H**。历史模板投影/单源链证据；当前同一批混期/混模板未实测。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E24](#e24)、[E30](#e30)。
- **回退类别/是否存在**：显式优先/业务识别兼容；已有此限定分支。
- 回退触发：显式指定模板或目录存在唯一匹配；不存在可用数值区域时另见第9步。
- 回退执行：显式模板优先于目录候选，不是读值失败回退。
- **回退测试：R**。E24合成CLI含显式模板优先/同分歧义拒绝；当前GUI批次未验证。
- **当前失败处理**：模板歧义/业务不唯一/无报告期或多期冲突→整批阻断；某些底层sheet读取错误转空集合，应提示真实错误。
- 可优化：模板元数据一次缓存；损坏包与未匹配原因分开。
- 需业务判断的场景：[S5F1-E02](#s5f1-e02)、[G04](#g04)
- 源码：[s5_template_matching.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_template_matching.rs)、[s5_business.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_business.rs)、[s5_period_recognition.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_period_recognition.rs)

#### S5F1.03 读S5配置并选本期生效版本

- 输入/前置：独立S5配置、业务、报告期、规则版本/启用/有效期。
- **当前技术**：Calamine→Rust Configuration/DSL+select_rules_for_period。
- 实际处理：预检与运行使用同一生效规则选择；已停用或过期版本不应要求来源；显式source覆盖默认来源。
- **主路径测试：U**。旧配置加载和运行时通过；最新同版本筛选/显式source预检测试已写未跑。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E20](#e20)、[E21](#e21)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：全局配置错误/无业务启用规则/来源合同有歧义→整批停止；具体规则错误按运行结果保留。
- 可优化：缓存DSL解析和活跃规则，但缓存键必须含配置hash、业务、期间/版本。
- 需业务判断的场景：[G05](#g05)、[S5F1-E03](#s5f1-e03)、[G12](#g12)
- 源码：[s5_config.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_config.rs)、[s5_config_dsl.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_config_dsl.rs)、[s5_service.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_service.rs)、[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)

#### S5F1.04 规范化逻辑来源与预检外部能力

- 输入/前置：当前/上期/外部逻辑source id，规则显式source。
- **当前技术**：Rust来源合同/优先级与业务绑定。
- 实际处理：GUI只承载一个当前逻辑入口和一个外部物理簿；启用规则需外部但未选时明确失败；不错误地用失效规则扩充来源要求。
- **主路径测试：U**。旧CLI多源歧义负向对拍通过；新GUI预检选择修复未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E28](#e28)、[E00](#e00)。
- **回退类别/是否存在**：来源绑定兼容规则；已有此限定分支。
- 回退触发：取数调用未显式source，且默认/合同/业务候选能唯一确定。
- 回退执行：按现有明确优先规则选择；歧义报错，不静默择一。
- **回退测试：N**。E28只证明旧歧义正确报错，不是替代来源成功；未找到该来源兜底成功的单独证据，最新显式source覆盖逻辑U。
- **当前失败处理**：多个必要逻辑来源无法由GUI入口表达/外部必需未选→整批阻断；不能随便取第一个来源。
- 可优化：将‘歧义负向通过’与‘替代来源成功’单列，减少错误迁移。
- 需业务判断的场景：[S5F1-E04](#s5f1-e04)
- 源码：[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[s5_retrieval.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_retrieval.rs)、[s5_service.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_service.rs)

#### S5F1.05 识别机构身份和绑定上期

- 输入/前置：本期/上期banner/来源合同、机构参照/外部catalog、代码名称。
- **当前技术**：Calamine ScanBook+Rust身份校验；复用机构参照映射。
- 实际处理：本期必须有可信机构代码；上期按同机构唯一匹配；名称别名只按显式映射，不猜企业。
- **主路径测试：U**。旧V_PRE CLI通过；新按机构批次、严格外部身份/invalid flags测试未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E22](#e22)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：本期身份冲突/无映射→当前文件失败继续其他；同机构多份上期→当前文件失败；没有上期→依赖V_PRE规则失败而仅本期规则可继续；上期识别问题记录批次。
- 可优化：分别列出身份失败、上期缺失、上期歧义；用户决定是否应让纯本期规则在上期歧义时继续。
- 需业务判断的场景：[S5F1-E05](#s5f1-e05)、[S5F1-E06](#s5f1-e06)、[S5F1-E07](#s5f1-e07)
- 源码：[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[s5_external_layout.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_external_layout.rs)

#### S5F1.06 捕获全部物理输入hash

- 输入/前置：本/上期每份、外部、模板、配置。
- **当前技术**：SHA256+Rust fs。
- 实际处理：记录执行前hash，物理路径参与身份；同逻辑source id的多文件不能互相覆盖清单；后续与源及副本逐份对比。
- **主路径测试：U**。旧单源哈希变化降级通过；新多物理路径键/副本碰撞修复未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E27](#e27)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：执行前某输入hash读取失败→全局返回；执行中变化→最终成功规则降级，结果不能作审核结论。
- 可优化：每物理文件只读一次哈希或流式复制同时hash并最终复核；一致性优先于省一次IO。
- 需业务判断的场景：[G14](#g14)、[S5F1-E08](#s5f1-e08)
- 源码：[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[s5_outputs.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_outputs.rs)

#### S5F1.07 读取本/上期存储值及结构

- 输入/前置：绑定的源XLSX/XLSM。
- **当前技术**：Calamine+OOXML名字/extents/数字lexeme+Rust合并单元格映射。
- 实际处理：ScanBook保留缓存值、公式位置、合并区域、存储类型和整数词法，供检索/结构比对；没有源公式重算。
- **主路径测试：R**。历史ScanBook/规则取数子集通过；新批次主链未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E21](#e21)、[E22](#e22)、[E25](#e25)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：损坏包/现代未知错误token、关键公式/合并状态不可确认→读取/绑定错误；在文件循环内通常标该文件失败继续；执行前hash失败则整批。
- 可优化：重用同文件ScanBook，别在身份/检索/CF各重复全簿读；不能引入openpyxl读业务。
- 需业务判断的场景：[G04](#g04)、[G09](#g09)、[S5F1-E09](#s5f1-e09)
- 源码：[s5_scan.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_scan.rs)、[lib.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/crates/excel-reader/src/lib.rs)

#### S5F1.08 外部数据布局、机构目录和日期定位

- 输入/前置：选定外部簿、集中数据/机构参照、静态/时序布局。
- **当前技术**：Calamine+Rust s5_external_layout；catalog映射-only模式。
- 实际处理：外部取数按逻辑来源/机构代码或已映射别名/指标/期别/单位，拒绝未映射身份或非法标记；不要求catalog-only解析出全部数据布局。
- **主路径测试：U**。旧多外部歧义负向通过；新strict catalog/无映射/标记修复测试未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E28](#e28)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：外部映射/布局异常在当前机构身份/检索阶段形成文件或规则失败，继续其他；共享损坏外部可能让全部文件逐一失败，最终FAILED。
- 可优化：共享外部全局结构先解析一次，失败可提前列受影响机构；缺单机构映射按文件，不混成所有规则失败。
- 需业务判断的场景：[S5F1-E10](#s5f1-e10)、[S5F1-E11](#s5f1-e11)
- 源码：[s5_external_layout.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_external_layout.rs)、[s5_retrieval.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_retrieval.rs)

#### S5F1.09 确定数值区域与行列/结构范围

- 输入/前置：源命名区、显式模板、目录候选、数值布局。
- **当前技术**：OOXML定义名称+Rust NumericScopeResolver/版式投影。
- 实际处理：数值区域优先源/已匹配模板；投影匹配处理动态行列/布局；找不到可用数值范围时保留旧整sheet检索兼容行为并提示。
- **主路径测试：R**。真实4区域坐标子集+合成模板选择；无模板377行契约匹配（含失败结果），非全业务通过。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E24](#e24)、[E25](#e25)。
- **回退类别/是否存在**：数值范围兼容兜底；已有此限定分支。
- 回退触发：未获得可用数值命名区域/模板投影，且路径允许旧兼容范围。
- 回退执行：保留整sheet搜索并质量提示；不是Calamine读取失败时换库。
- **回退测试：R**。E24合成兼容歧义结果、E25无模板运行契约有历史报告；不证明整sheet搜索对任意新布局正确或业务通过。
- **当前失败处理**：缺唯一布局/投影不符可返回STRUCTURE_ERROR；兼容无区域路径扩大检索范围，可能出现多值歧义，不能当精确匹配。
- 可优化：业务可把正式审核限定有效数值区，把整sheet标‘需复核’；这会改政策，当前不自动实施。
- 需业务判断的场景：[S5F1-E12](#s5f1-e12)、[S5F1-E13](#s5f1-e13)
- 源码：[s5_numeric_scope.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_numeric_scope.rs)、[s5_layout.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_layout.rs)、[s5_retrieval.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_retrieval.rs)

#### S5F1.10 V_CUR/V_PRE/V_EXT取数与定位

- 输入/前置：规则参数、单位/过滤、数值区、机构、报告期。
- **当前技术**：Rust RetrievalEngine+Decimal值/检索轨迹。
- 实际处理：按数据归属、逻辑source优先级和行列标签定位；返回数值/单位/真实坐标/源hash轨迹，区别空、0、缺失、歧义。
- **主路径测试：R**。旧单绑定/V_PRE/区域投影/歧义负向有限子集；不是当前批次完整通过。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E21](#e21)、[E22](#e22)、[E24](#e24)、[E28](#e28)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：找不到/多候选/单位不兼容/无上期等→RETRIEVAL_FAILED或STRUCTURE_ERROR，当前规则保留失败继续其他规则。
- 可优化：索引行列标签、缓存重复查询但键含source/机构/期间/区域/单位；缓存错误也要有来源。
- 需业务判断的场景：[S5F1-E07](#s5f1-e07)、[S5F1-E11](#s5f1-e11)、[S5F1-E13](#s5f1-e13)、[G11](#g11)
- 源码：[s5_retrieval.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_retrieval.rs)、[s5_layout.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_layout.rs)、[s5_external_layout.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_external_layout.rs)

#### S5F1.11 按配置直接算规则与生成结论

- 输入/前置：取数、变量字典、前提/比较式、启用版本。
- **当前技术**：Rust s5_service/s5_eval/s5_config_dsl+Decimal/已有业务表达式。
- 实际处理：不复制模板公式；执行配置规则，输出PASSED/TRIGGERED/CONFIG_ERROR/RETRIEVAL_FAILED/CALCULATION_FAILED等；记录原物理源归属。
- **主路径测试：U**。旧80项运行时比对0差异含3错误；新owner/生效版本/显式来源联动未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E21](#e21)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：规则解析/计算/取数失败保留对应失败结果并继续；严重运行函数错误记当前文件失败继续其他；不是交Excel再算回退。
- 可优化：不为S5引入Office通用计算；编译规则一次、执行多机构，同时保留每项实际执行状态。
- 需业务判断的场景：[G12](#g12)、[S5F1-E14](#s5f1-e14)
- 源码：[s5_service.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_service.rs)、[s5_eval.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_eval.rs)、[s5_config_dsl.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_config_dsl.rs)

#### S5F1.12 只读补扫模板校验区域公式错误

- 输入/前置：模板校验区域公式坐标、本期源已保存缓存。
- **当前技术**：复用S1ValidationFormulaScanPlan+Rust S5结果投影。
- 实际处理：发现#REF!/#DIV0等源缓存错误/缺sheet，作为带机构/源路径的S5失败项；不注入/不重算模板公式。GUI逐本期文件；CLI仅明确模板时扫primary。
- **主路径测试：U**。新cached_validation_errors_are_projected_as_owned_s5_failures_without_recalculation测试已写未运行；无主路径成功证据。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：模板读取/扫描失败、缺sheet→失败Outcome并继续；无显式模板的CLI可SKIPPED，不得宣称完成模板公式审核。
- 可优化：将‘Rust规则审核完成’和‘源保存公式错误扫描’分别计数；如果需要刷新源Excel缓存，应另开显式计算能力，不能假称当前已刷新。
- 需业务判断的场景：[S5F1-E15](#s5f1-e15)、[S5F1-E16](#s5f1-e16)
- 源码：[s5_validation_formula_scan.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_validation_formula_scan.rs)、[s1_validation_formula_scan.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s1_validation_formula_scan.rs)、[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[open-data-audit.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/bin/open-data-audit.rs)

#### S5F1.13 条件格式区域读取、求值与归属

- 输入/前置：模板条件格式区域（若配置）、源CF及保存值、机构与报表项目。
- **当前技术**：OOXML规则+Calamine值+复用Rust有限CF evaluator。
- 实际处理：模板限定扫描范围，规则和值仍来自源；无模板指定范围使用源AppliesTo；输出TRIGGERED/UNSUPPORTED/ERROR真实判断，不读红色就算异常。
- **主路径测试：U**。底层S1有限CF历史通过；旧S5输出CF为空；新模板范围与报表归属测试未运行，非空S5CF整链无成功证据。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E05](#e05)、[E23](#e23)、[E00](#e00)。
- **回退类别/是否存在**：扫描范围默认；已有此限定分支。
- 回退触发：未提供模板条件格式区域。
- 回退执行：按源规则AppliesTo扫描；不是读取失败回退，存在模板但读取坏不能假装没配置。
- **回退测试：N**。底层CF有旧用例；新S5区域桥接/批次非空输出没有实际运行证据。
- **当前失败处理**：读/评估异常被记录PARTIAL、保留诊断；其他规则/文件继续；不把不支持当无触发。
- 可优化：将未支持CF列为用户可选阻断/需复核策略，精确到规则/机构；加非空多机构写出样本。
- 需业务判断的场景：[G09](#g09)、[S5F1-E17](#s5f1-e17)
- 源码：[s5_conditional_format.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_conditional_format.rs)、[conditional_format.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/conditional_format.rs)、[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)

#### S5F1.14 结构快照比较与质量提示

- 输入/前置：本机构结构观察、已有基线、数值检索警告。
- **当前技术**：Rust JSON快照/哈希+PendingSnapshots。
- 实际处理：GUI批次比较既有基线，finish_read_only，不把本次机构或无证据区域自动覆盖基线；保留物理源/机构owner。CLI旧单次仍可持久化被接受基线。
- **主路径测试：U**。旧结构快照契约历史完成；新只读/owner/无证据不污染测试未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E26](#e26)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：结构变动形成诊断/结构错误；读取目录/key/写总账问题可能导致输出失败；read_only比较仍可能创建.schema-key，不等于完全零写文件。
- 可优化：明确首次无基线是‘未比较’而非通过；基线接纳应单独显式操作，不随批次自动刷新。
- 需业务判断的场景：[S5F1-E18](#s5f1-e18)、[S5F1-E19](#s5f1-e19)
- 源码：[s5_schema_snapshots.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_schema_snapshots.rs)、[s5_outputs.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_outputs.rs)、[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)

#### S5F1.15 制作安全审计副本与逐文件完整性校验

- 输入/前置：全部实际参与源/模板/配置、执行前hash。
- **当前技术**：Rust fs create_new唯一命名+SHA256复核；不经Excel编辑。
- 实际处理：不改原件；原名清洗冲突用唯一副本名；每物理文件前/后/副本hash分别核验并绑定owner；失真时降级结论。
- **主路径测试：U**。旧单源输入变化降级通过；新多文件同source id、副本名碰撞、逐物理hash保护未运行。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E27](#e27)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：空间/权限/副本失败→输出失败；源改变/副本不一致→INPUT_CHANGED_DURING_RUN，成功结论降级，CF标INVALIDATED_INPUT_CHANGED。
- 可优化：已创建失败产物保留清晰状态；不要输出可误认为正式完成的孤立XLSX。
- 需业务判断的场景：[G02](#g02)、[G14](#g14)、[S5F1-E20](#s5f1-e20)
- 源码：[s5_outputs.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_outputs.rs)

#### S5F1.16 写13列结果/CF/质量表及运行总账

- 输入/前置：规则/公式扫描/CF结果、批次统计/源归属/副本链接。
- **当前技术**：rust_xlsxwriter新XLSX+serde JSON ledger+Web状态。
- 实际处理：只将触发与失败行写主结果，CF独立13列含状态；保留全规则计数。全部处理0→FAILED；有文件/规则失败→COMPLETED_WITH_ISSUES；仅CF限制→COMPLETED_WITH_CONDITIONAL_FORMAT_LIMITATIONS；输入变化优先失效。
- **主路径测试：U**。旧输出diff0但CF/质量为空；旧单源CLIcompleted；新批次混合/全失败/非空CF/owner联动无运行证据。 当前版本：15:26相关修复没有编译或运行；旧单源/CLI证据不能证明当前GUI批次。 证据：[E23](#e23)、[E30](#e30)、[E00](#e00)。
- **回退类别/是否存在**：无生产引擎回退；当前没有实际可用的生产替代引擎/路径。
- 回退触发：当前步骤失败
- 回退执行：按明确配置/文件/规则/输出失败处理，不改用Python业务或Excel通用计算。
- **回退测试：NA**。无实际生产替代引擎可标通过。
- **当前失败处理**：写XLSX/JSON/路径错误→返回失败，可能有已生成副本；CF受限/规则失败可输出部分结果但不能当全部通过。当前不承诺每份新输出都有统一独立Calamine读回验收。
- 可优化：产物暂存+总账最后提交；明确业务触发数量与软件失败数量。需要只对关键输出增加读回，避免广泛样式审计。
- 需业务判断的场景：[G06](#g06)、[G10](#g10)、[G13](#g13)、[S5F1-E21](#s5f1-e21)
- 源码：[s5_outputs.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_outputs.rs)、[s5_web.rs](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/src/s5_web.rs)、[index.html](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/frontend/web/index.html)

## 104个失败场景：当前行为与待决定策略

“建议”不等于已经实施。“停止本功能整批”仅指本次功能调用，不关闭应用，也不阻断其他独立功能。停止文件可以是机构/工作簿；停止规则可以是区域/指标。继续仅能生成明确标注缺项的结果，不能把未审项算通过。

可选决定：STOP_ALL（停止本功能整批）；STOP_FILE（停止当前文件/机构，继续其他）；STOP_RULE（停止当前规则/区域，继续其他）；CONTINUE_DIAGNOSTIC（继续并列明缺项）；MANUAL_REVIEW（仅给需人工复核结果）；RETRY_SAFE（保护原件后安全重试）；DIAGNOSTIC_ONLY（结果失效，只保留诊断）。

P0表示建议优先判断的正确性或结果失效风险，P1为其余待决策略，不是本轮确认的软件缺陷等级。很多条目是应按设计明确拒绝的能力边界。

<a id="g01"></a>

### G01 已选输入不存在、路径指向目录/临时产物、输出与输入相同

- 优先级/涉及功能：P1；S1-F1、S1-F2、S5-F1；步骤S1F1.01、S1F2.01、S5F1.01。
- **当前实际行为**：入口路径/身份检查返回错误；不同入口校验覆盖不完全一致。
- 业务影响：无法确定实际被审核文件或可能覆盖原件。
- **建议供判断**：停止本功能整批（STOP_ALL）。开始前拦截并给具体路径；不能静默换文件/改输出覆盖源。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g02"></a>

### G02 没有权限、输入占用、输出/临时盘空间不足、网络盘断连

- 优先级/涉及功能：P1；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1；步骤S1F1.01、S1F1.06、S1F1.14、S1F2.01、S1F2.07、S2F1.09、S3F1.08、S3F2.05、S3F3.04、S3F3.06、S3F4.06、S3F5.06、S5F1.15。
- **当前实际行为**：通常IO错误返回；S1/F5暂存guard清理，S5可能已留下副本/部分输出。
- 业务影响：读取不全或产物不完整；备份和源安全受影响。
- **建议供判断**：停止本功能整批（STOP_ALL）。共享配置/输出故障阻断批次；只影响单源的读取可按文件隔离，前提是明确缺项。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g03"></a>

### G03 没有选文件/所有预检来源被跳过/批次0可执行项

- 优先级/涉及功能：P1；S1-F1、S1-F2、S5-F1；步骤S1F1.01、S1F2.01、S5F1.01。
- **当前实际行为**：大多入口开始前返回无来源；S5执行后全失败可生成FAILED总账。
- 业务影响：0条异常可能被误读成全部通过。
- **建议供判断**：停止本功能整批（STOP_ALL）。保留失败/跳过清单；不出‘审核通过’。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g04"></a>

### G04 XLSX ZIP损坏、XML/关系异常、文件加密、伪扩展名、Calamine不支持类型

- 优先级/涉及功能：P1；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F5、S5-F1；步骤S1F1.02、S1F1.03、S1F1.04、S1F2.02、S1F2.04、S2F1.02、S2F1.07、S3F1.03、S3F2.01、S3F3.01、S3F5.01、S5F1.02、S5F1.07。
- **当前实际行为**：没有通用Calamine→OOXML/openpyxl值读取回退；按所在阶段返回错误或预检跳过。
- 业务影响：所有依赖值/坐标都不可信；元数据补读不能补救值读取。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。共享模板/配置坏应整批停止；仅一份源坏可隔离并标部分未审。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g05"></a>

### G05 共享配置或模板缺表、关键表头/定义名称错误、作用域不明

- 优先级/涉及功能：P1；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1；步骤S1F1.03、S1F1.04、S1F2.02、S2F1.01、S3F1.01、S3F2.01、S3F3.01、S3F4.01、S3F5.02、S5F1.03。
- **当前实际行为**：配置/模板加载/规则规划报错；部分旧兼容路径可warning继续，应逐功能确认。
- 业务影响：可能影响所有文件/规则，无法得出可信统一结论。
- **建议供判断**：停止本功能整批（STOP_ALL）。配置/模板级错误优先阻断受影响功能；不能把错表头当空规则。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g06"></a>

### G06 目标已存在、被Excel占用、原子提交/硬链接/rename失败

- 优先级/涉及功能：P1；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1；步骤S1F1.06、S1F1.09、S1F1.14、S1F2.07、S1F2.08、S2F1.09、S3F1.08、S3F2.05、S3F3.04、S3F4.06、S3F5.07、S5F1.16。
- **当前实际行为**：按具体输出函数报错且不覆盖；有的已保存文件会保留，有的清理暂存。
- 业务影响：有内容不代表已成功发布；重试可能覆盖或重复。
- **建议供判断**：保护原件后安全重试（RETRY_SAFE）。只可在保留同一已验证候选且新输出路径的前提下安全重试；失败应有明确产物清单。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g07"></a>

### G07 文件含未支持图表/透视/VBA/ActiveX/嵌入/外链/未知drawing等，或平台无Passed编辑器

- 优先级/涉及功能：P0；S1-F1、S3-F3、S3-F5；步骤S1F1.07、S3F3.05、S3F5.04、S3F5.05。
- **当前实际行为**：编辑前profile/能力筛选拒绝；当前生产只有Umya，无openpyxl后备。
- 业务影响：继续编辑可能丢关系/公式/关键对象。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。未覆盖对象先保持原件，只出诊断；新后备须覆盖整批同契约后才可启用。
- 故障/回退测试证据：R/H：E03/E04有限保留样本；特定未知对象拒绝为现源码规定，未证明所有负向输入。
- **用户决定：待决定。**

<a id="g08"></a>

### G08 编辑已写入后报错、结果验证失败、后端冒称写前拒绝但已写出字节

- 优先级/涉及功能：P0；S1-F1、S3-F3、S3-F5；步骤S1F1.07、S3F3.05、S3F5.05。
- **当前实际行为**：ExcelModifier停止回退；不发布无效候选。
- 业务影响：再用另一引擎可能掩盖/放大损坏，不能从坏候选继续。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。丢弃候选，保留原件/诊断；不能当自动后备成功。
- 故障/回退测试证据：N：有框架mock测试源码，未定位当前实际双引擎故障回执。
- **用户决定：待决定。**

<a id="g09"></a>

### G09 条件格式有不支持规则/函数、缺缓存/上下文、色阶/图标集/数据条等未实现判断

- 优先级/涉及功能：P0；S1-F1、S5-F1；步骤S1F1.12、S5F1.07、S5F1.13。
- **当前实际行为**：CF求值输出unsupported/error；S1GUI要求完整否则整批停，S5保留PARTIAL继续。
- 业务影响：未审的条件可能隐藏真实异常；颜色或规则存在不能代替本轮触发。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。可继续生成带缺项结果，但受影响范围不允许标通过；用户决定必审规则是否阻断。
- 故障/回退测试证据：R：E05有限支持子集；新S5CF桥U。
- **用户决定：待决定。**

<a id="g10"></a>

### G10 输出超过XLSX单表行列上限，或单格文本过长

- 优先级/涉及功能：P1；S1-F2、S2-F1、S3-F1、S3-F2、S3-F4、S5-F1；步骤S1F2.07、S2F1.09、S3F1.08、S3F2.05、S3F4.06、S5F1.16。
- **当前实际行为**：部分writer显式限制/库返回错误，S5CF有行数检查；不会自动分卷。
- 业务影响：可能丢行/截文本或无法保存。
- **建议供判断**：停止本功能整批（STOP_ALL）。不静默截断；用户另行决定分sheet/分卷是否符合结果契约。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g11"></a>

### G11 单位缺失/不识别、极大/非有限数值、前导0标识、金额舍入/浮点尾差

- 优先级/涉及功能：P1；S2-F1、S3-F1、S3-F2、S3-F4、S3-F5、S5-F1；步骤S2F1.03、S3F1.04、S3F2.04、S3F4.04、S3F5.03、S5F1.10。
- **当前实际行为**：不同模块采用Decimal或f64及既定兼容语义；可能规则错误/质量警告/文本保留。
- 业务影响：影响业务匹配和阈值判定；代码与金额不是同种转换。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。保留原单位/原始类型/规范值；影响全局单位则整批停止，不自行改舍入规则。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="g12"></a>

### G12 配置表达式包含不支持函数/语法、除零、正则语义差异

- 优先级/涉及功能：P1；S2-F1、S3-F1、S3-F4、S5-F1；步骤S2F1.06、S3F1.06、S3F1.07、S3F4.03、S3F4.05、S5F1.03、S5F1.11。
- **当前实际行为**：配置错误或规则日志/失败Outcome；没有Python eval或Excel补算。
- 业务影响：某规则没有实际审核，不能当未触发。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。隔离受影响规则并计入未完成；若为业务必审规则再由用户决定文件/整批阻断。
- 故障/回退测试证据：R：E14/E18/E21含有限错误案例，非全部语法。
- **用户决定：待决定。**

<a id="g13"></a>

### G13 GUI显示完成但包含文件/规则失败，或结果文件存在而总账/日志失败

- 优先级/涉及功能：P0；S2-F1、S3-F1、S3-F3、S3-F4、S3-F5、S5-F1；步骤S2F1.10、S3F1.09、S3F3.04、S3F3.08、S3F4.07、S3F5.08、S5F1.16。
- **当前实际行为**：各功能状态口径不一；S5已有部分/全失败细分，其他入口可能主要返回计数或通用完成。
- 业务影响：用户误把部分结果当完整通过。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。已有有效产物可展示，但须明确未完成范围/数量和不可作为通过结论的部分。
- 故障/回退测试证据：H/U：历史API/CLI记录，当前GUI真实交互未验收。
- **用户决定：待决定。**

<a id="g14"></a>

### G14 运行期间输入/配置/模板被另一个进程改动，或副本hash不一致

- 优先级/涉及功能：P0；S3-F3、S5-F1；步骤S3F3.05、S5F1.06、S5F1.15。
- **当前实际行为**：S3F3提交前拒绝；S5降级成功规则和CF结论；其他路径源保护不等于都有实时变化监视。
- 业务影响：一份报告混合了不同版本输入。
- **建议供判断**：结果失效，只保留诊断（DIAGNOSTIC_ONLY）。只保留带失效标记诊断；需要输入稳定后整批重跑。
- 故障/回退测试证据：R：E27旧单源降级；新多物理源保护U。
- 已确认约束：不得将变化输入产生的结果当已审核通过。
- **用户决定：待决定。**

<a id="s1f1-e01"></a>

### S1F1-E01 S1F1缺计算sidecar或引擎能力路径

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.01。
- **当前实际行为**：开始执行前拒绝，不进入模板复制/审核。
- 业务影响：F1公式结果无法刷新，不能以存储旧缓存顶替。
- **建议供判断**：停止本功能整批（STOP_ALL）。该功能需要真实计算；其他不需要计算的功能可独立使用。
- 故障/回退测试证据：H：E02/E29有环境缺失及成功记录，当前引擎前置行为为静态核对。
- **用户决定：待决定。**

<a id="s1f1-e02"></a>

### S1F1-E02 模板不匹配/多个候选/元数据读取坏转空sheet集合

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.02。
- **当前实际行为**：预检排除该文件；部分推荐函数将错误变空集合。
- 业务影响：可能误把损坏文件说成模板未匹配。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。跳过该文件并保留实际解析/匹配原因；整批仍有合格文件可继续。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="s1f1-e03"></a>

### S1F1-E03 表结构/表头不匹配或名称区域无法确认

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.03。
- **当前实际行为**：预检记录原因并跳过；全跳过则停止。
- 业务影响：套错模板会复制到错位置。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保持现有预检隔离；共享模板错误则整批阻断。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="s1f1-e04"></a>

### S1F1-E04 现代公式错误#SPILL!/#CALC!等导致Calamine不支持

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.04、S1F1.09、S1F1.10。
- **当前实际行为**：仅sidecar校验有opt-in临时归一读取；后续公式/CF/校验区域读取未统一使用。
- 业务影响：可能已完成真实重算，却在后续抽取阶段整批失败。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。应先保证错误token可无损作为异常透传；不得跳过报错单元格而宣称通过。
- 故障/回退测试证据：N：unknown_error_tokens测试源码存在但细分执行回执未定位；未跑当前链。
- **用户决定：待决定。**

<a id="s1f1-e05"></a>

### S1F1-E05 模板公式依赖外部sheet但未选外部/外部缺sheet

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.05。
- **当前实际行为**：规划/读取时报缺引用，GUI正式执行错误导致整批停止。
- 业务影响：#REF或缺数导致公式审核无效。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。缺依赖时该文件不可审；可由用户决定其他独立源是否继续。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="s1f1-e06"></a>

### S1F1-E06 外部sheet有普通公式、跨sheet依赖或UsedRange从C4等非A1开始

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.05。
- **当前实际行为**：当前Rust用Umya导入公式，新增依赖闭包/保持起点；不是有公式就拒绝。
- 业务影响：旧起点错误/漏依赖会使公式值错误；修复未执行测试。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。应支持这类正常输入；在修复完成验证前失败不能当跳过外部继续。
- 故障/回退测试证据：U：external_formula_sheets_and_their_dependencies_are_copied_without_rejecting_formulas未运行。
- **用户决定：待决定。**

<a id="s1f1-e07"></a>

### S1F1-E07 INDIRECT/OFFSET等动态sheet引用，扩大全部外部后又重名

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.05。
- **当前实际行为**：当前扩大全部sheet；与源sheet同名（大小写不敏感）拒绝。
- 业务影响：静态闭包不全或覆盖源sheet风险。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。动态公式允许扩大来源需完整检查；不能为了继续随意改sheet名且不改所有引用。
- 故障/回退测试证据：N：动态全导入/冲突测试源码存在，独立成功回执未定位。
- **用户决定：待决定。**

<a id="s1f1-e08"></a>

### S1F1-E08 输入/模板包含共享、数组、动态数组等带特殊属性公式

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.06。
- **当前实际行为**：OOXML缓存清理及能力契约可能在写前明确拒绝。
- 业务影响：普通复制/缓存清除可能破坏数组或共享主从关系。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。普通公式可用不覆盖特殊公式；不能清掉属性强行继续。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="s1f1-e09"></a>

### S1F1-E09 跨区域/跨表公式引用或必要批注/名称保存后不一致

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.07。
- **当前实际行为**：编辑或保留校验失败终止；当前无实际openpyxl自动后备。
- 业务影响：公式逻辑、异常说明或定位改变。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保留原件，只允许重新从原件生成候选且证明相同契约后再试。
- 故障/回退测试证据：R有限样本E03/E04；复杂外部公式最新U。
- **用户决定：待决定。**

<a id="s1f1-e10"></a>

### S1F1-E10 Windows无Excel/COM激活失败、PowerShell策略拦截或COM挂起

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.08。
- **当前实际行为**：当前PowerShell COM路径报错或等待，无pywin32、无明确硬超时、无LO自动后备。
- 业务影响：重算无法完成，界面可能长时间停留运行中。
- **建议供判断**：停止本功能整批（STOP_ALL）。先按要求换/补pywin32桥和生命周期保护；当前失败不读旧缓存继续。
- 故障/回退测试证据：H：旧Excel16合成成功；pywin32路径N。
- **用户决定：待决定。**

<a id="s1f1-e11"></a>

### S1F1-E11 Linux无soffice、profile不可写、LO没产生输出

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.08。
- **当前实际行为**：sidecar明确失败，GUI停止；没有引擎回退。
- 业务影响：不能完成F1真实重算。
- **建议供判断**：停止本功能整批（STOP_ALL）。该功能提示需要可用LO；其他Rust功能独立继续。
- 故障/回退测试证据：H：E29缺soffice负向记录；LO成功N。
- **用户决定：待决定。**

<a id="s1f1-e12"></a>

### S1F1-E12 Excel与LO公式语义不同、外部链接/易失函数/异步数据依赖未完成

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.08。
- **当前实际行为**：选择引擎后没有运行时跨引擎重试；公式文本/缓存检查不覆盖所有语义差异。
- 业务影响：缓存存在仍可能是错误业务值。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。只允许经业务确认的另一计算引擎从原候选重算，并比较关键结果；不能静默降级。
- 故障/回退测试证据：N：无跨引擎成功回退和真实复杂公式对照。
- **用户决定：待决定。**

<a id="s1f1-e13"></a>

### S1F1-E13 重算后公式文本改变、缓存缺失/数量不符、返回引擎错误

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.09。
- **当前实际行为**：sidecar/调用层校验失败，独立候选不发布。
- 业务影响：公式计算结果不可信。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。不能以修改公式比较规则/填0掩盖；保留诊断供修复。
- 故障/回退测试证据：H：普通Excel合成链E02；失败样本无当前独立证据。
- **用户决定：待决定。**

<a id="s1f1-e14"></a>

### S1F1-E14 规则定位sheet/坐标不存在、公式结果抽取/校验扫描失败

- 优先级/涉及功能：P1；S1-F1；步骤S1F1.10、S1F1.11。
- **当前实际行为**：GUI用?终止整批；CLI有的扫描错误转失败Issue。
- 业务影响：少提异常或扫描范围错误。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。受影响文件未完成，用户决定是否保留其他文件完整结果。
- 故障/回退测试证据：N：未找到针对这一故障及当前源码的独立执行证据。
- **用户决定：待决定。**

<a id="s1f1-e15"></a>

### S1F1-E15 GUI扫重算副本、CLI校验区域扫原源缓存导致不同结论

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.11。
- **当前实际行为**：现源码口径不同；CLI是已复制manifest入口而非GUI完整复制链。
- 业务影响：原缓存与本轮公式不同时可能漏报/多报。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。必须用反例确定既定业务口径；在确定前不宣称两入口等价。
- 故障/回退测试证据：N：未找到专门区分源/副本缓存差异的成功验收。
- **用户决定：待决定。**

<a id="s1f1-e16"></a>

### S1F1-E16 CF规则不支持或指标/批注上下文不完整

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.12。
- **当前实际行为**：S1 GUI强制CF report.status=completed，否则整批停止并不生成最终结果。
- 业务影响：部分CF不确定导致全部已算好的文件也不交付。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。建议文件级保守阻断，其他完整文件可生成部分结果；是否必须整批由业务决定。
- 故障/回退测试证据：R：支持子集E05/E02；全体CF语法不适用。
- **用户决定：待决定。**

<a id="s1f1-e17"></a>

### S1F1-E17 已选历史簿打不开/历史key重复或格式有误

- 优先级/涉及功能：P1；S1-F1、S1-F2；步骤S1F1.13、S1F2.06。
- **当前实际行为**：GUI历史读取错误导致整批失败；没选历史才用空历史。
- 业务影响：不能正确承接说明/识别延续。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。可保留本次异常并标‘历史分类未完成’，是否允许按无历史输出需业务决定，不能隐瞒。
- 故障/回退测试证据：R历史正常分类E07；具体故障N。
- **用户决定：待决定。**

<a id="s1f1-e18"></a>

### S1F1-E18 批次第N个文件失败或最终writer失败

- 优先级/涉及功能：P0；S1-F1；步骤S1F1.14。
- **当前实际行为**：正式循环?使整批结束；run目录guard删除此前审核副本/成果。
- 业务影响：其他独立文件成功成果没有交付。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。建议隔离独立文件并输出部分批次清单；当前未实现，是否修改由用户决定。
- 故障/回退测试证据：N：未定位当前多文件‘前成功后失败’业务验收。
- **用户决定：待决定。**

<a id="s1f2-e01"></a>

### S1F2-E01 数据区/表头区/附加字段命名区缺失或定义非法

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.02。
- **当前实际行为**：规划时失败，终止本次汇总。
- 业务影响：列语义不能确定，所有汇总可能错位。
- **建议供判断**：停止本功能整批（STOP_ALL）。模板/映射级错误阻断；不是临时改成第一行表头继续。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s1f2-e02"></a>

### S1F2-E02 部分文件表结构预检失败

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.03。
- **当前实际行为**：跳过该工作簿，保留原因；全跳过报错。
- 业务影响：汇总缺少部分来源。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。可继续其他文件，但必须显示未参与清单。
- 故障/回退测试证据：H/R：历史汇总与人工样例记录；本轮未重新跑坏表头。
- **用户决定：待决定。**

<a id="s1f2-e03"></a>

### S1F2-E03 预检已过但正式读源/OOXML范围失败

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.04。
- **当前实际行为**：正式读循环?返回，整个汇总失败。
- 业务影响：单坏文件使其他结果不交付。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。如果业务允许不完整汇总，可文件隔离并显式列缺项；当前未实现此改动。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s1f2-e04"></a>

### S1F2-E04 需要汇总的公式缓存缺失、无法确认公式或合并信息

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.04。
- **当前实际行为**：ensure_formula_cache_values/状态检查报错，不使用公式文本顶替。
- 业务影响：汇总值可能是空/旧值而非本轮值。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。要求用户先计算保存；若要自动重算应明确新能力/入口，不能宣称当前F2已重算。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s1f2-e05"></a>

### S1F2-E05 源sheet无同名模板但存在有效兜底模板

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.05。
- **当前实际行为**：按兜底任意行/固定行区域读取，结果可继续。
- 业务影响：范围选择依赖业务模板，不是智能猜表。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。明确标明采用兜底及来源sheet，便于抽查。
- 故障/回退测试证据：R：E08两种兜底分支各4行、内容一致、源未变，历史通过。
- **用户决定：待决定。**

<a id="s1f2-e06"></a>

### S1F2-E06 源sheet无同名模板且无兜底，或数据区全空

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.05。
- **当前实际行为**：跳过该sheet/空数据并计数。
- 业务影响：部分sheet没有参与汇总。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。若应有业务数据但缺模板，可能应阻断文件；纯空sheet可正常略过。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s1f2-e07"></a>

### S1F2-E07 附加字段/历史匹配不到或重复

- 优先级/涉及功能：P1；S1-F2；步骤S1F2.06。
- **当前实际行为**：既定缺匹配可留空并计数；非法历史读取终止。
- 业务影响：说明和机构辅助字段可能缺失。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。空值合法性由业务判断，不能把读失败当匹配不到。
- 故障/回退测试证据：R：E08/E09所列空字段/历史正常匹配；歧义负向N。
- **用户决定：待决定。**

<a id="s2f1-e01"></a>

### S2F1-E01 未选外部文件但配置存在外部规则

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.01。
- **当前实际行为**：服务按可选输入跳过外部核对分支。
- 业务影响：其他规则完成不代表外部规则完成。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。若外部核对属必审，应阻断或标该分支未执行；不能仅因未选就显示全功能通过。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s2f1-e02"></a>

### S2F1-E02 文件名不识别、或已识别文件打不开

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.02。
- **当前实际行为**：不识别→warning跳过；识别后读取错误→整个服务停止。
- 业务影响：两个相似坏文件触发不同范围停止。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保留身份不可识别/损坏两类原因，让其余机构可部分完成需明确政策。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s2f1-e03"></a>

### S2F1-E03 机构代码/名称/表单归属不一致、指标缺参照或异常金额

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.03。
- **当前实际行为**：源身份优先，配置补缺；warning/质量问题及规则失败按现逻辑。
- 业务影响：可能用错机构/指标单位。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。身份冲突阻断该机构；纯描述字段缺失可仅提示，不一律置0。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s2f1-e04"></a>

### S2F1-E04 相同逻辑key在多文件中出现不同值

- 优先级/涉及功能：P0；S2-F1；步骤S2F1.04。
- **当前实际行为**：当前保留第一条并记录质量问题。
- 业务影响：后续规则使用哪份取决于排序，可能漏异常。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。相同值重复可提示；不同值冲突建议隔离对应机构/指标，用户决定是否整批停。
- 故障/回退测试证据：R有限合成导入契约；当前真实冲突样例N。
- **用户决定：待决定。**

<a id="s2f1-e05"></a>

### S2F1-E05 上期缺机构/指标、分母0、空值与真实0混用

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.05。
- **当前实际行为**：按既定环比/质量逻辑诊断或跳过。
- 业务影响：无法算环比不等于该规则通过。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。继续本期可算规则并列环比缺项；核心环比必审可扩大阻断。
- 故障/回退测试证据：R：E12/E14有限样本。
- **用户决定：待决定。**

<a id="s2f1-e06"></a>

### S2F1-E06 S2表达式无取数/不支持函数/除零

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.06。
- **当前实际行为**：配置检查或规则执行报错/诊断，没有COM/LO补算。
- 业务影响：校验没有实际执行。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。保留rule id和失败取数，不能隐式false。
- 故障/回退测试证据：R：E14所列错误案例；非全部语法。
- **用户决定：待决定。**

<a id="s2f1-e07"></a>

### S2F1-E07 已选外部不是XLSX、缺集中系统数据/参照、静态时序布局错

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.07。
- **当前实际行为**：外部导入?返回，当前整个S2服务停止。
- 业务影响：一份外部坏可能使可算的环比结果也不交付。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。可分离外部分支并标未完成；若外部为必审则整批停止。
- 故障/回退测试证据：R：E13正常非时序；异常/时序全覆盖N。
- **用户决定：待决定。**

<a id="s2f1-e08"></a>

### S2F1-E08 外部机构/指标/期别不存在或多行冲突

- 优先级/涉及功能：P1；S2-F1；步骤S2F1.08。
- **当前实际行为**：规则取数/外部匹配诊断，不存在通用读取回退。
- 业务影响：取错外部值会改变审核结论。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。仅隔离受影响外部规则，保留坐标/候选与单位。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f1-e01"></a>

### S3F1-E01 3.0默认单位/3.1规则动作或表达式配置缺失、无效

- 优先级/涉及功能：P1；S3-F1；步骤S3F1.01。
- **当前实际行为**：配置/服务加载失败；部分空配置可能无规则可执行。
- 业务影响：所有比较/警戒尺度可能错误。
- **建议供判断**：停止本功能整批（STOP_ALL）。核心运行单位必须有效，启用规则数量应预览确认。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f1-e02"></a>

### S3F1-E02 GBK失败后UTF8读取，或UTF8字节也能被GBK无错解码

- 优先级/涉及功能：P1；S3-F1、S3-F4；步骤S3F1.02、S3F4.02。
- **当前实际行为**：解码优先GBK，然后UTF8；不是先UTF8。
- 业务影响：英文表头/文本可能混合可读但实际汉字误解码。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。用户应确认编码来源；报告记录选择路径/必要时显式编码，不仅看能否解码。
- 故障/回退测试证据：N：未定位每个编码分支独立成功凭证。
- **用户决定：待决定。**

<a id="s3f1-e03"></a>

### S3F1-E03 两种严格解码均失败、最终GBK替换产生字符损失

- 优先级/涉及功能：P0；S3-F1、S3-F4；步骤S3F1.02、S3F4.02。
- **当前实际行为**：当前以替换解码继续，再按业务表头/行解析。
- 业务影响：机构/指标代码或说明被替换会造成丢配/错配。
- **建议供判断**：停止本功能整批（STOP_ALL）。建议出现替换字符先阻断本文件；若损坏仅非关键说明字段可业务批准带警告继续。
- 故障/回退测试证据：N：替换后内容仍正确的成功证据未见。
- **用户决定：待决定。**

<a id="s3f1-e04"></a>

### S3F1-E04 CSV/Excel不满足14列表头，或数据行列数异常

- 优先级/涉及功能：P1；S3-F1、S3-F4；步骤S3F1.03、S3F4.02。
- **当前实际行为**：关键表头错误停止；异常行按ImportIssue保留/略过。
- 业务影响：某些数据行未进入后续比较。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。表头错不可猜列；坏行允许隔离必须显示数量及行号。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f1-e05"></a>

### S3F1-E05 同文件数据日期不一致、两期关系异常、代码前导0丢失

- 优先级/涉及功能：P1；S3-F1、S3-F4；步骤S3F1.03、S3F4.02。
- **当前实际行为**：同文件日期不一致明确报错；两期关系按period警告/规则；数字代码已丢0不能自动还原。
- 业务影响：期别/身份配对可能错误。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。期间/身份是全局键；不能用任意一行日期代表整个文件。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f1-e06"></a>

### S3F1-E06 八维相同key重复且不同值；拼接key无分隔歧义

- 优先级/涉及功能：P0；S3-F1；步骤S3F1.04。
- **当前实际行为**：当前CentralDataset字典后条覆盖；key为多字段字符串拼接。
- 业务影响：重复与潜在拼接碰撞改变采用的值。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。先报告重复与采用来源；应以结构化tuple优化key但需保留旧业务重复语义再验收。
- 故障/回退测试证据：N：未见真实冲突/拼接碰撞专项验收。
- **用户决定：待决定。**

<a id="s3f1-e07"></a>

### S3F1-E07 本期或上期独有记录/非数值/合法空值

- 优先级/涉及功能：P1；S3-F1；步骤S3F1.05。
- **当前实际行为**：按base comparison既定分支构造比较，可能采用缺期默认值/保留异常。
- 业务影响：当作0会影响增减和警戒。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。结果必须有缺期/异常值标记；能否参与阈值由业务判定。
- 故障/回退测试证据：R：E15有限真实比较子集。
- **用户决定：待决定。**

<a id="s3f1-e08"></a>

### S3F1-E08 动作/表达式规则错误或无法取数

- 优先级/涉及功能：P1；S3-F1；步骤S3F1.06、S3F1.07。
- **当前实际行为**：按具体分支写日志/跳过或传播配置错误；不调用旧Python求值。
- 业务影响：比较文件可能有效但部分规则未执行。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。输出未执行规则清单，是否必审由用户选。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f1-e09"></a>

### S3F1-E09 主结果已写而可选日志写失败/读回失败

- 优先级/涉及功能：P1；S3-F1、S3-F4；步骤S3F1.08、S3F4.06。
- **当前实际行为**：错误返回，前面已保存结果可能存在；无通用跨文件事务。
- 业务影响：用户看到文件但本次整体显示失败。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。如果主结果已独立验证可展示，明确日志缺失；更稳妥是统一暂存后提交。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3-e01"></a>

### S3-E01 仍按旧Rust编号选错F2/F3/F4/F5输入，或把F3当直接更新比较文件

- 优先级/涉及功能：P0；S3-F1、S3-F2；步骤S3F1.09、S3F2.06。
- **当前实际行为**：当前正式菜单已按五项；历史名字/CLI辅助命令仍可能混淆。
- 业务影响：错误文件能读但业务完全不同。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。展示产物类型与预期关键表头；F1→F2→F3、F1→F5，F4独立，不能只凭文件后缀。
- 故障/回退测试证据：H：菜单/调用源已核，真实五项接续人工验收未证。
- **用户决定：待决定。**

<a id="s3f2-e01"></a>

### S3F2-E01 比较文件手工改了关键表头/删输出说明列

- 优先级/涉及功能：P1；S3-F2；步骤S3F2.01。
- **当前实际行为**：找不到必需字段报错，无法导出。
- 业务影响：不能确定人工选择和说明数据。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。不自动退回规则‘是否说明’或猜列位置。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f2-e02"></a>

### S3F2-E02 人工输出说明全未选或标记不在兼容真值集合

- 优先级/涉及功能：P1；S3-F2；步骤S3F2.02。
- **当前实际行为**：未选行被排除，可能生成0行且completed。
- 业务影响：用户可能以为有规则要求说明就会导出。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。0行可以是正常，但应提示实际选中0并可由业务要求阻断。
- 故障/回退测试证据：R：E16非空2610；E30 0行输出，范围分开。
- **用户决定：待决定。**

<a id="s3f2-e03"></a>

### S3F2-E03 3.0机构地区映射缺失/多地区身份不一致

- 优先级/涉及功能：P1；S3-F2；步骤S3F2.03。
- **当前实际行为**：按现有输入/参照映射分支输出；读配置失败阻断。
- 业务影响：机构展示或F3身份后续不一致。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。原始代码保留；多地区派生标识由既定规则，不能靠近似名称改码。
- 故障/回退测试证据：R：E16有限真实映射，歧义负向N。
- **用户决定：待决定。**

<a id="s3f3-e01"></a>

### S3F3-E01 反馈表头被机构修改、目标XLSM、任一反馈坏文件

- 优先级/涉及功能：P1；S3-F3；步骤S3F3.01。
- **当前实际行为**：当前读入/格式检查停止整个导入计划。
- 业务影响：其他合法反馈也不导入。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。可考虑隔离反馈文件后更新其余，但先生成清单；目标不支持必须停止本次。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f3-e02"></a>

### S3F3-E02 机构/指标码、数据值或增减额被改/规范化不一致

- 优先级/涉及功能：P0；S3-F2、S3-F3；步骤S3F2.04、S3F2.06、S3F3.02。
- **当前实际行为**：四字段key不匹配计入unmatched，跳过；机构/指标名称不作key。
- 业务影响：原因说明未写入预期行。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。允许其他合法行继续，但必须给可定位未匹配明细；不能降低为名称相似匹配。
- 故障/回退测试证据：R：E17计划子集；真实金额改动专项N。
- **用户决定：待决定。**

<a id="s3f3-e03"></a>

### S3F3-E03 同key反馈不同原因，或目标同key多行

- 优先级/涉及功能：P1；S3-F3；步骤S3F3.03。
- **当前实际行为**：冲突计数与冲突报告；跳过冲突key，合法行继续。
- 业务影响：不能自动决定哪个原因写哪行。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。保留所有冲突来源，用户后续修正再导，不自动选最后一个。
- 故障/回退测试证据：R：E17计划与2条冲突已比对。
- **用户决定：待决定。**

<a id="s3f3-e04"></a>

### S3F3-E04 目标原因已有文本，新唯一反馈将其更新

- 优先级/涉及功能：P1；S3-F3；步骤S3F3.03。
- **当前实际行为**：当前合法SetCellValue计划可替换原因；与来源冲突不同。
- 业务影响：用户已有人工说明可能被覆盖。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。先备份是已确认保护；是否非空原因差异也作冲突需要用户决定，不能擅改现行为。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f3-e05"></a>

### S3F3-E05 无法创建/验证字节备份

- 优先级/涉及功能：P0；S3-F3；步骤S3F3.06。
- **当前实际行为**：执行停止，原目标不被替换。
- 业务影响：丢失可恢复原版风险。
- **建议供判断**：停止本功能整批（STOP_ALL）。保持‘无有效备份不写回’，不提供无备份继续。
- 故障/回退测试证据：R：E04 Linux备份成功；失败故障N。
- 已确认约束：用户已明确：必须原地更新目标文件并先备份。
- **用户决定：待决定。**

<a id="s3f3-e06"></a>

### S3F3-E06 目标被锁、或候选安装rename失败

- 优先级/涉及功能：P0；S3-F3；步骤S3F3.07。
- **当前实际行为**：目标首次rename失败则原件不动；已挪走时尝试rollback恢复。
- 业务影响：原路径可能暂时不存在；导入未成功。
- **建议供判断**：保护原件后安全重试（RETRY_SAFE）。只在确认已恢复原件且输入hash一致后重试；回滚成功也不算导入成功。
- 故障/回退测试证据：H：E02称Linux回滚覆盖；Windows当前独立证据N。
- **用户决定：待决定。**

<a id="s3f3-e07"></a>

### S3F3-E07 安装失败且rollback恢复也失败

- 优先级/涉及功能：P0；S3-F3；步骤S3F3.07。
- **当前实际行为**：抛带备份/rollback路径的错误，保留恢复材料。
- 业务影响：目标路径可能缺失或状态不确定。
- **建议供判断**：停止本功能整批（STOP_ALL）。立即停止本功能，必须人工确认恢复，禁止自动继续写其他目标。
- 故障/回退测试证据：N：未找到双rename故障注入的当前证据。
- **用户决定：待决定。**

<a id="s3f3-e08"></a>

### S3F3-E08 写回成功但rollback暂存文件删除失败

- 优先级/涉及功能：P1；S3-F3；步骤S3F3.08。
- **当前实际行为**：返回cleanup warning；保留正式目标与旁备份。
- 业务影响：多余副本不影响已验证结果，但不能盲删恢复材料。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。告知清理警告；无需重跑导入或删除正式备份。
- 故障/回退测试证据：N：清理故障专项未见。
- **用户决定：待决定。**

<a id="s3f4-e01"></a>

### S3F4-E01 当前GUI仅本期，CLI请求跨期但上期缺失或规则期别不一致

- 优先级/涉及功能：P1；S3-F4；步骤S3F4.01。
- **当前实际行为**：GUI/CLI参数能力不同；加载/取数按实际绑定报错。
- 业务影响：某跨期规则无输入，不等于可在本期模式执行。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。区分纯本期必需输入与跨期规则必需输入；不要把F4全部称两期比较。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f4-e02"></a>

### S3F4-E02 工作簿日期型/DurationIso值不受支持

- 优先级/涉及功能：P1；S3-F4；步骤S3F4.02。
- **当前实际行为**：DurationIso等明确报不支持，不按任意字符串/0继续。
- 业务影响：错误日期或金额会污染全部指标。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保留真实类型/位置，用户应修源值或确定兼容语义后再改。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f4-e03"></a>

### S3F4-E03 机构/地区scope模式无法表达或不匹配

- 优先级/涉及功能：P1；S3-F4；步骤S3F4.03。
- **当前实际行为**：不匹配正常不执行；解析/表达式问题走日志/配置错误。
- 业务影响：错误scope可能导致没有规则执行却输出成功。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。明确‘未适用’和‘模式错误’的不同状态。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f4-e04"></a>

### S3F4-E04 前提/左右表达式取数失败

- 优先级/涉及功能：P0；S3-F4；步骤S3F4.04。
- **当前实际行为**：写执行日志并continue当前规则，不写正常比较行。
- 业务影响：结果表看不到失败规则，可能被当0异常。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。把跳过规则数量合并到run状态；必审规则可提升到文件阻断。
- 故障/回退测试证据：R：E18有214结果+3日志，不能说无规则错误。
- **用户决定：待决定。**

<a id="s3f4-e05"></a>

### S3F4-E05 校验公式错误、反转/抑制标记有误、epsilon归零边界

- 优先级/涉及功能：P1；S3-F4；步骤S3F4.05。
- **当前实际行为**：校验公式错误可记formula_error并继续生成带错结果；双0不出行。
- 业务影响：不出行未必通过；阈值边界影响结论。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。明确输出错误说明，近零容差需业务认可，不擅改旧规则。
- 故障/回退测试证据：R：E18所列真实配置，不涵盖全部边界。
- **用户决定：待决定。**

<a id="s3f4-e06"></a>

### S3F4-E06 本期核对报告完成但有规则错误日志

- 优先级/涉及功能：P1；S3-F4；步骤S3F4.07。
- **当前实际行为**：当前结果/日志可成功输出；GUI是否清晰呈现需人工验收。
- 业务影响：completed可能只是产物完成。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。完整列出失败规则，不把‘214结果行比对成功’当业务无异常。
- 故障/回退测试证据：R：E18结果/日志一致；GUI当前N。
- **用户决定：待决定。**

<a id="s3f5-e01"></a>

### S3F5-E01 F5比较表头改名/列被重排但列数仍≥15

- 优先级/涉及功能：P0；S3-F5；步骤S3F5.01。
- **当前实际行为**：当前warning后按固定列位置读取。
- 业务影响：可把机构/指标/金额读错后生成看似正常表单。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。建议列位置改变阻断；仅允许已明确的同义表头兼容并核对实际索引。
- 故障/回退测试证据：N：未见该兼容分支正确业务输出专项证据。
- **用户决定：待决定。**

<a id="s3f5-e02"></a>

### S3F5-E02 同机构同指标映射key重复或输出机构名碰撞

- 优先级/涉及功能：P1；S3-F5；步骤S3F5.01、S3F5.03。
- **当前实际行为**：比较数据映射先条保留；不同机构清洗后输出名相同则提前报错。
- 业务影响：重复值可能被忽略；产物不能唯一归属。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。相同值可提示，冲突值阻断对应机构；命名冲突可用稳定代码区分但需保留业务可识别名。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f5-e03"></a>

### S3F5-E03 使用仓库空占位3.3，或报表清单没有启用表单

- 优先级/涉及功能：P0；S3-F5；步骤S3F5.02。
- **当前实际行为**：有效表单/映射为空，规划或执行报错/无法生成。
- 业务影响：空配置不代表没有应输出的业务。
- **建议供判断**：停止本功能整批（STOP_ALL）。必须使用真实3.3模板并预览启用数；原只读参考有43条启用记录不能靠占位验收。
- 故障/回退测试证据：H：DEVELOPMENT_GOAL已有明确限制；真实全量验收N。
- **用户决定：待决定。**

<a id="s3f5-e04"></a>

### S3F5-E04 比较簿单位与生成时3.0目标单位不同，3.3遗留单位另有值

- 优先级/涉及功能：P0；S3-F5；步骤S3F5.02。
- **当前实际行为**：3.3字段忽略；F5取比较簿既有值，无独立3.3换算路径；3.0主要提供警戒。
- 业务影响：错误来源比较簿可能带不同量纲而无全局校验。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保持用户已确认忽略3.3；应验证F1产物单位来源，不能自行再按3.3换算。
- 故障/回退测试证据：N：跨配置单位不一致专项未见。
- 已确认约束：用户已确认3.3分析文件数据单位为遗留，沿用3.0默认目标单位。
- **用户决定：待决定。**

<a id="s3f5-e05"></a>

### S3F5-E05 金融表指标/频度/批次/币种无映射，生成空白或缺值

- 优先级/涉及功能：P1；S3-F5；步骤S3F5.03。
- **当前实际行为**：按配置映射，未命中按warning/空值规则；不自动借用近似指标。
- 业务影响：文件生成成功但关键数据缺失。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。列出每表缺项；业务必填项未命中应阻断机构而不是默认为0。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s3f5-e06"></a>

### S3F5-E06 删列/删表影响已有公式/合并/CF/DV/打印范围

- 优先级/涉及功能：P1；S3-F5；步骤S3F5.04。
- **当前实际行为**：编辑能力预检和保留校验；受支持样本可处理，未知对象失败。
- 业务影响：结构错位可能改变数据和审核判断。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。不能为了少量版式差异扩迁，但关键引用/CF区域必须安全。
- 故障/回退测试证据：R：E03/E04受限结构样本，不代表真实3.3所有对象。
- **用户决定：待决定。**

<a id="s3f5-e07"></a>

### S3F5-E07 模板存在defined names且本批次要DeleteColumn

- 优先级/涉及功能：P0；S3-F5；步骤S3F5.05。
- **当前实际行为**：Umya后端在写前拒绝；生产无合格openpyxl后备。
- 业务影响：删列后名称可能失效，不能确认关键范围。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保留模板，针对真实模板确定必要依赖；openpyxl也不自动维护所有删列依赖，不能简单换库承诺解决。
- 故障/回退测试证据：N：无该类模板成功编辑或后备证据。
- **用户决定：待决定。**

<a id="s3f5-e08"></a>

### S3F5-E08 多机构批次第N个编辑失败/最终目录发布失败

- 优先级/涉及功能：P0；S3-F5；步骤S3F5.06、S3F5.07。
- **当前实际行为**：当前停止，staging目录guard清理，不发布前面成功机构。
- 业务影响：一机构问题阻断所有金融表交付。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。如允许部分完成，需先设计每机构失败清单和可靠发布；当前行为不可说成已隔离机构。
- 故障/回退测试证据：H：历史成功/失败烟测；真实多机构混合故障N。
- **用户决定：待决定。**

<a id="s3f5-e09"></a>

### S3F5-E09 金融表单保存成功但映射/空表/警戒有警告

- 优先级/涉及功能：P1；S3-F5；步骤S3F5.08。
- **当前实际行为**：当前可返回输出路径和warnings。
- 业务影响：用户可能以为每个启用表都填全。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。生成状态、业务完备性和警告分开；必填缺项由用户定阻断等级。
- 故障/回退测试证据：N：未找到该故障在当前源码上的独立执行回执。
- **用户决定：待决定。**

<a id="s5f1-e01"></a>

### S5F1-E01 批次混模板/混业务/混期，选上期未填期间

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.01。
- **当前实际行为**：全局匹配/期间/参数检查阻断。
- 业务影响：不能为多机构共享同一活跃规则和模板上下文。
- **建议供判断**：停止本功能整批（STOP_ALL）。当前GUI合同是同模板同业务同期间，不擅自拆分成多批。
- 故障/回退测试证据：U：当前批次修复未执行。
- **用户决定：待决定。**

<a id="s5f1-e02"></a>

### S5F1-E02 模板候选同分/显式模板坏包/业务或报告期不唯一

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.02。
- **当前实际行为**：全局识别失败；某些元数据底层可空集合导致未匹配。
- 业务影响：匹配错将影响全部规则和CF范围。
- **建议供判断**：停止本功能整批（STOP_ALL）。必须明确唯一模板业务期别；保留解析错误细节。
- 故障/回退测试证据：R：E24旧CLI歧义拒绝/显式优先；当前GUI U。
- **用户决定：待决定。**

<a id="s5f1-e03"></a>

### S5F1-E03 过期/禁用版本要求来源，或显式source被默认来源覆盖

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.03。
- **当前实际行为**：当前源码已统一生效规则筛选与显式source优先，修复未运行。
- 业务影响：错误地阻断输入，或从错误文件取数。
- **建议供判断**：停止本功能整批（STOP_ALL）。预检与执行必须同规则快照；生效规则错误不能用跳过掩盖。
- 故障/回退测试证据：U：active revision/explicit source测试未运行。
- **用户决定：待决定。**

<a id="s5f1-e04"></a>

### S5F1-E04 GUI一个外部物理簿无法表达启用规则引用多个逻辑外部来源

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.04。
- **当前实际行为**：明确报错，建议按来源拆配置或CLI显式多来源；不取第一个。
- 业务影响：来源选择不唯一。
- **建议供判断**：停止本功能整批（STOP_ALL）。保留当前GUI边界；多源能力要在CLI单独验收，不能把CLI扩展当GUI已有。
- 故障/回退测试证据：R：旧多外部歧义负向E28；新预检U。
- **用户决定：待决定。**

<a id="s5f1-e05"></a>

### S5F1-E05 当前机构代码缺失/冲突/外部参照无别名映射

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.05。
- **当前实际行为**：当前文件失败Outcome，继续其他文件。
- 业务影响：结果无法可靠归属机构。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。不能按文件名/近似机构名猜身份；待补映射或修banner后重跑。
- 故障/回退测试证据：U：新strict外部身份/owner修复未执行。
- **用户决定：待决定。**

<a id="s5f1-e06"></a>

### S5F1-E06 同机构匹配到多份上期

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.05。
- **当前实际行为**：当前整个本期文件标PREVIOUS_SOURCE_AMBIGUOUS失败继续其他文件。
- 业务影响：纯本期规则也没有继续执行。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。建议只阻断依赖上期规则而让纯本期继续，但需业务确认并改代码；当前尚未这么做。
- 故障/回退测试证据：U：当前GUI批次逻辑未执行。
- **用户决定：待决定。**

<a id="s5f1-e07"></a>

### S5F1-E07 某机构没有上期或V_PRE指标缺失

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.05、S5F1.10。
- **当前实际行为**：V_PRE取数失败保留结果；仅本期规则可继续；上期身份问题记批次。
- 业务影响：未做完跨期规则不能算文件完整通过。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。明确缺的是上期文件还是具体指标，用户决定跨期必审是否文件级阻断。
- 故障/回退测试证据：R：旧V_PRE正常CLI E22；新GUI缺期分支U。
- **用户决定：待决定。**

<a id="s5f1-e08"></a>

### S5F1-E08 执行前某个文件hash读取失败，或同source id多物理文件覆盖清单

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.06。
- **当前实际行为**：前者全局返回；后者已新增物理path区分，修复未执行。
- 业务影响：批次来源/副本无法逐份核对。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。共享配置/模板hash失败仍整批；单源可隔离，但不能丢hash继续出通过结论。
- 故障/回退测试证据：U：多物理hash修复与测试未运行。
- **用户决定：待决定。**

<a id="s5f1-e09"></a>

### S5F1-E09 源ScanBook打不开/公式或合并状态不可确认/未知错误token

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.07。
- **当前实际行为**：文件身份/检索绑定阶段失败时记录文件失败继续其他；global hash/识别阶段可能先阻断。
- 业务影响：不具备可靠取数数据。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。保留实际发生阶段，不能笼统说所有坏文件都能隔离。
- 故障/回退测试证据：R：旧读取子集；当前批次故障N。
- **用户决定：待决定。**

<a id="s5f1-e10"></a>

### S5F1-E10 外部catalog缺机构映射/有非法启用标记

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.08。
- **当前实际行为**：当前严格解析为失败，能落到文件/规则问题；新修复未跑。
- 业务影响：不能把机构默认匹配至错误外部行。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。影响全部机构的catalog全局错误应一次阻断/诊断，单机构缺映射按文件。
- 故障/回退测试证据：U：unmapped identity/invalid flags测试未运行。
- **用户决定：待决定。**

<a id="s5f1-e11"></a>

### S5F1-E11 外部布局错、指标/时间点不存在、源优先级歧义

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.08、S5F1.10。
- **当前实际行为**：外部绑定/取数失败，保留失败项，不选任意候选。
- 业务影响：外部审核未执行。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。其他来源规则可继续，必审外部规则可提升文件阻断。
- 故障/回退测试证据：R：旧歧义负向E28；真实时序全覆盖N。
- **用户决定：待决定。**

<a id="s5f1-e12"></a>

### S5F1-E12 没有数值命名区域/没有可用模板投影，退到整sheet检索

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.09。
- **当前实际行为**：旧兼容范围保留，带质量提示；会扩大候选。
- 业务影响：可能从表头/说明/其他表块取到同名标签。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。可作为诊断/手工复核结果，正式必审是否允许整sheet由用户决定。
- 故障/回退测试证据：R：E24/E25旧范围契约匹配，不代表所有结果业务通过。
- **用户决定：待决定。**

<a id="s5f1-e13"></a>

### S5F1-E13 动态行列变动、合并结构不符、多个同名指标值

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.09、S5F1.10。
- **当前实际行为**：返回STRUCTURE_ERROR/RETRIEVAL_FAILED，继续其他规则。
- 业务影响：有值并不意味着取到了唯一正确值。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。不采用随意首个候选；保留候选坐标及变化结构。
- 故障/回退测试证据：R：E21/E25旧有限结构错误结果；新owner U。
- **用户决定：待决定。**

<a id="s5f1-e14"></a>

### S5F1-E14 规则语法/变量/单位错误或某文件规则运行函数失败

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.11。
- **当前实际行为**：单规则失败保留；文件级异常生成RULE_EXECUTION_FAILED后继续其他。
- 业务影响：只有部分规则运行，不能用无TRIGGERED称无问题。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。必审规则失败可升级文件未完成；结果要统计全部执行状态。
- 故障/回退测试证据：R：E21 80项含3错误；最新批次U。
- **用户决定：待决定。**

<a id="s5f1-e15"></a>

### S5F1-E15 源校验区域公式缓存陈旧/缺失，但S5规则本身已能计算

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.12。
- **当前实际行为**：S5只读保存缓存，不做S1式注入/重算；新scanner仅投影已有错误/失败。
- 业务影响：不能证明源Excel公式本轮已正确算过。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。分别展示Rust规则结论和缓存错误扫描，不把后者称真重算；若需要刷新另开显式Excel/LO能力。
- 故障/回退测试证据：U：新缓存scanner测试未运行。
- **用户决定：待决定。**

<a id="s5f1-e16"></a>

### S5F1-E16 模板校验区域扫描失败、源缺sheet、CLI未给明确模板

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.12。
- **当前实际行为**：GUI失败Outcome并继续；CLI无明确模板可SKIPPED，只扫primary。
- 业务影响：模板公式错误扫描覆盖不全。
- **建议供判断**：停止当前规则/区域，继续其他（STOP_RULE）。主规则可以继续，但必须标该扫描未完成；若该补扫必审可阻断文件。
- 故障/回退测试证据：U：新模块未运行。
- **用户决定：待决定。**

<a id="s5f1-e17"></a>

### S5F1-E17 CF模板范围错误/源规则不支持/源缓存不新/无法确认上下文

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.13。
- **当前实际行为**：当前CF PARTIAL保留诊断，规则/其他文件继续；无模板区域默认AppliesTo。
- 业务影响：不确定规则不是0异常。
- **建议供判断**：仅给需人工复核结果（MANUAL_REVIEW）。按CF业务重要性决定阻断文件或输出需复核；不能把显示颜色作替代判定。
- 故障/回退测试证据：U：新桥接范围/归属未运行，旧S5输出CF为空。
- **用户决定：待决定。**

<a id="s5f1-e18"></a>

### S5F1-E18 结构基线来自另一机构/本次没有证据却写基线

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.14。
- **当前实际行为**：新GUI批次finish_read_only和owner设计避免覆盖，测试未跑；CLI旧单次持久化不同。
- 业务影响：后续结构审核基准被污染。
- **建议供判断**：停止当前文件/机构，继续其他（STOP_FILE）。无证据状态应未比较，不自动接受；按机构/模板归属核验。
- 故障/回退测试证据：U：read_only owner/no-evidence测试未运行。
- **用户决定：待决定。**

<a id="s5f1-e19"></a>

### S5F1-E19 结构快照目录/.schema-key无权限，或首次无基线

- 优先级/涉及功能：P1；S5-F1；步骤S5F1.14。
- **当前实际行为**：read_only比较仍可能创建key；目录/key错误可使输出失败；无基线需区别未比较。
- 业务影响：‘只读基线’不代表完全不写本地状态；可能影响交付。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。若业务允许，可以只缺结构比较并输出已完成规则，但不能声称结构已通过。
- 故障/回退测试证据：N：目录/只读介质故障未实测；新逻辑U。
- **用户决定：待决定。**

<a id="s5f1-e20"></a>

### S5F1-E20 不同源名清洗后相同、副本被覆盖/损坏、复制后原件又变

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.15。
- **当前实际行为**：新create_new唯一副本/逐物理hash，完整性不一致使结论降级。
- 业务影响：链接可能指错机构或错误版本。
- **建议供判断**：结果失效，只保留诊断（DIAGNOSTIC_ONLY）。安全副本没验证就不交正式结论；可保留失效诊断。
- 故障/回退测试证据：U：sanitized collision/per physical hash用例未运行。
- **用户决定：待决定。**

<a id="s5f1-e21"></a>

### S5F1-E21 多文件部分失败/全失败、CF不完整、结果XLSX成功但ledger失败

- 优先级/涉及功能：P0；S5-F1；步骤S5F1.16。
- **当前实际行为**：新状态区分FAILED/COMPLETED_WITH_ISSUES/CF_LIMITATIONS/INPUT_CHANGED；写出失败返回错误，可能已有副本。
- 业务影响：产物存在但不可理解完整性；失败机构可能错误显示通过。
- **建议供判断**：继续并列明缺项（CONTINUE_DIAGNOSTIC）。保留每物理文件归属、全状态统计和总账；全失败不能标完成审核。
- 故障/回退测试证据：U：当前混合/全失败/非空CF/ledger联动未跑。
- **用户决定：待决定。**

## 优化建议：保留现有技术库与业务实现

所有建议均未在本轮实施。库官方文档用于确认能力边界，不作为本项目测试成功证据；源码锁定版本以本报告为准。没有任何库自动保证任意工作簿和全部对象保真。

### O01 完成已写的S1外部公式与S5新修复验证

- 顺序/范围：先补正确性；S1-F1、S5-F1。
- 当前：最新修复晚于现有二进制及12:28验证；普通源、含公式外部表、非空CF、多机构owner路径不能以旧报告代替。
- 建议：先按本报告验证清单做用户授权后的定向构建/样本验收，失败后局部修复。保留现有有效Rust实现。
- 价值：比再增加一个未验证引擎更直接降低错数/错归属风险。
- 边界：本轮不执行；不扩大全项目对拍，不重构Excel-Core。
- 需要的定向验证（本轮未执行）：S1非A1公式跨表依赖；S5多源副本、规则版本/显式source、外部映射、CF范围、缓存公式扫描、快照只读。

### O02 Windows计算桥按要求改为隔离pywin32 COM

- 顺序/范围：计算能力；S1-F1。
- 当前：生产sidecar以PowerShell调用真实Excel，合成成功不是pywin32成功。
- 建议：保留现有sidecar请求/结果契约，Windows桥改用隔离pywin32 win32com.client启动独立Excel实例；只承担打开/CalculateFullRebuild/等待/保存/关闭与引擎回执，不迁入审核逻辑。
- 价值：符合技术要求，便于超时、COM异常和应用进程生命周期管理；不声称天然更快。
- 边界：pywin32不是计算引擎本身，仍需Excel；不能使用用户系统Python或联网安装作为交付前提。
- 需要的定向验证（本轮未执行）：普通公式/外部公式、计算状态、超时、Excel已开其他文件、COM失败、SaveAs失败、退出后无本任务孤儿进程、源hash不变，全部重新验证。
- 官方依据：[pywin32 COM支持](https://github.com/mhammond/pywin32)、[Excel CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild)、[Excel CalculationState](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationstate)。

### O03 LO独立成功验收与显式换引擎政策

- 顺序/范围：计算能力；S1-F1。
- 当前：Linux有缺soffice正确报错记录；没有成功LO真重算和与Excel关键结果对照。auto不是回退链。
- 建议：保留headless/独立UserInstallation和隔离输入输出；提供明确引擎来源/版本、超时/输出检查，按真实模板测关键结果。跨引擎尝试须从原始已注入候选重新开始。
- 价值：建立Linux确实可用的公式计算能力，并明确失败边界。
- 边界：不静默COM→LO；LO函数和保存语义不能假设等于Excel；核心S2/S3/S5无需装LO。
- 需要的定向验证（本轮未执行）：成功保存/缓存刷新、兼容与不兼容函数、profile不可写、路径空格/中文、进程占用/挂起、无输出、目标麒麟/UOS与离线运行时。
- 官方依据：[LibreOffice启动参数与profile要求](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)。

### O04 复用本次解析结果，减少重复解压和全簿读取

- 顺序/范围：读取优化；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1。
- 当前：同一源可能在身份/模板/公式/CF/取数中多次打开；Calamine专用返回列限制不等于底层完全流式。
- 建议：功能内部复用现有WorkbookDump/ScanBook、模板名称和结构、配置解析快照；缓存键包含规范路径与执行前hash。需要值才读值，sheet目录只用已存在OOXML元数据适配器。
- 价值：减少IO/解压/重复计算，尤其S5批次及S1模板重复读。
- 边界：不以全局无限缓存占内存，不把不同文件/机构或变化输入共享错上下文，不另写通用reader。
- 需要的定向验证（本轮未执行）：同数据结果/来源轨迹一致，输入改变缓存失效，多文件内存/耗时实测后再决定。
- 官方依据：[Calamine读取值与公式API](https://docs.rs/calamine/latest/calamine/)。

### O05 现代错误token兼容集中到明确的读值契约

- 顺序/范围：读取正确性；S1-F1、S5-F1。
- 当前：sidecar校验能临时归一未知错误token，后续抽取/CF/扫描仍可能普通Calamine失败。
- 建议：在本地excel-reader现有opt-in能力上限定扩展调用，保留错误字面值/坐标/原始类型标记；只在已识别错误触发，不对损坏文件泛化重试。
- 价值：真实Excel#SPILL/#CALC等可成为明确异常而非整链莫名读失败。
- 边界：不是Calamine失败后用OOXML重写数据；不修改源、不把错误转0/空，不把公式文本当结果。
- 需要的定向验证（本轮未执行）：普通错误/新错误/合法文本同名/公式缓存错误/损坏XML均覆盖；sidecar到最终Issue和CF诊断必须同语义。

### O06 大报表评估rust_xlsxwriter constant_memory

- 顺序/范围：新结果写出；S1-F2、S2-F1、S3-F1、S3-F2、S3-F4、S5-F1。
- 当前：新建报表主要逐行写，workspace没有启用constant_memory feature。
- 建议：仅对数据量实际造成内存问题的新报表评估该feature和顺序写sheet；设置可控磁盘临时目录，保留较小随机访问sheet原模式。
- 价值：降低大结果写出的峰值内存；不是保证所有输入更快。
- 边界：必须按行递增，已过行不能回写；后置设置样式/合并/表格特性要逐项检查。不能用于已有工作簿修改，也不算公式。
- 需要的定向验证（本轮未执行）：大行数结果值/类型/链接/CF诊断列/来源列独立读回一致，低磁盘与temp目录不可写时明确失败。
- 官方依据：[rust_xlsxwriter 0.101.0内存模式与限制](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/performance/index.html)、[公式写入不会计算结果](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.set_formula_result)。

### O07 保持整批一次打开/保存，按样本评估Umya lazy_read

- 顺序/范围：编辑优化；S1-F1、S3-F3、S3-F5。
- 当前：既有ExcelEditPlan已可把一个工作簿一批修改交给同一后端；F5每机构仍需要独立输出。
- 建议：先消除后端内同源的重复加载，复用可验证的模板读取；只有真实大簿测量显示读取瓶颈时再评估umya lazy_read。
- 价值：降低多机构模板读取成本，同时保持现有编辑路由。
- 边界：lazy_read是库提供的候选能力，不代表当前preserve清单已通过；可变工作簿不能跨机构共享导致串值。
- 需要的定向验证（本轮未执行）：普通/未触碰sheet、CF/DV/名称/打印/批注关系、公式/缓存、输入hash及机构间独立性同样验证。
- 官方依据：[Umya 3.1.1 read/lazy_read](https://docs.rs/umya-spreadsheet/3.1.1/umya_spreadsheet/)。

### O08 openpyxl只补真实编辑缺口，能力合格后再登记

- 顺序/范围：限定兼容后备；S1-F1、S3-F3、S3-F5。
- 当前：worker源码存在，但无生产注册；S1候选历史还存在重算/结构问题。
- 建议：先挑用户实际失败的明确编辑批次，以隔离内置runtime和data_only=false保留公式；核实同一保留契约后登记相应功能/平台/操作，才可作为写前后备。
- 价值：可覆盖已证实的具体现有工作簿修改缺口，避免无谓重写。
- 边界：不做读取路由/审核业务/计算；openpyxl删列不会自动维护所有外部公式/定义名称等依赖，不能承诺解决F5 definedNames限制；跨工作簿整表也不是copy_worksheet即完成。
- 需要的定向验证（本轮未执行）：真实故障触发→主引擎写前拒绝→新原件副本→worker全批→关键保留/输出验证通过→输入不变；另外测写后失败绝不回退。
- 官方依据：[openpyxl不求值公式](https://openpyxl.readthedocs.io/en/stable/simple_formulae.html)、[openpyxl插删行列不管理全部依赖](https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html)、[copy_worksheet限制和保存注意](https://openpyxl.readthedocs.io/en/stable/tutorial.html)。

### O09 继续元数据/窄补丁，不新增通用XML复制引擎

- 顺序/范围：限定OOXML优化；S1-F1、S3-F5、S5-F1。
- 当前：已有名称/范围/CF/批注/缓存等窄适配器，旧Python有纯数据OOXML快路。
- 建议：优先复用已有适配器；如果性能测量确有必要，纯数据复制快路必须能在写前识别公式/关系/合并/DV等不支持项并交同契约后端。
- 价值：保留读取元数据和小补丁的优势，避免整簿重序列化。
- 边界：当前Rust无此生产快路；不能直接拷sheet.xml或遇公式删除公式；复杂公式交Umya/经验证编辑器再交Excel/LO重算。
- 需要的定向验证（本轮未执行）：空白起点、共享字符串、样式索引、关系、名称、公式/依赖、包部件及源文件保护；无实际需求不开发。

### O10 S3 key结构化、固定位置读表与单位来源检查

- 顺序/范围：业务正确性；S3-F1、S3-F5。
- 当前：central key字符串拼接与重复后覆盖；F5比较表头不符仍固定列读取、重复先保留。
- 建议：在业务批准后将内部key改tuple而不改八维含义；区分相同值重复/不同值冲突；F5严格列契约或明确同义表头索引，验证比较簿目标单位来源。
- 价值：减少无错误提示的错配/错位，比切换Excel库更直接。
- 边界：本轮只记录，不直接改重复优先级/列兼容/单位规则。3.3遗留单位仍忽略。
- 需要的定向验证（本轮未执行）：拼接碰撞反例、重复同值/异值、列重排/改名、不同3.0目标单位生成的比较簿。

### O11 统一‘产物完成’与‘审核完整’状态表达

- 顺序/范围：结果完整性；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1。
- 当前：S1正式失败整批丢弃、F5全部成功再发、S5部分结果细分；其他功能有日志但仍可完成。
- 建议：先由用户为104场景选策略，再局部实现文件/规则隔离和受影响范围。所有结果都保留执行/未执行/失败/触发/通过计数及来源。
- 价值：避免用‘无异常行’掩盖没执行，支持按业务重要性选择停止范围。
- 边界：不自动将S1/F5改成部分发布，不将必审失败标通过；只是建议。
- 需要的定向验证（本轮未执行）：前成功后失败、全部失败、0选中、CF不支持、日志失败、结果写成功但ledger失败，UI/CLI状态一致。

### O12 新结果统一暂存、原地事务保留恢复信息

- 顺序/范围：文件事务；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1。
- 当前：F3已有备份/回滚；F5目录事务；部分新报表与日志/JSON不在同一事务。
- 建议：复用现有temp/guard模式，将主文件+必要日志/总账纳入一次产物清单提交；F3不改变先备份原地更新合同。
- 价值：减少‘显示失败但有孤立正式文件’和重复提交。
- 边界：不删除既有备份/历史成果；提交后清理失败只警告，不能重跑已成功业务。
- 需要的定向验证（本轮未执行）：IO故障、进程中断/锁、候选安装失败/恢复失败、新文件名碰撞；只在副本/临时目录故障注入。

### O13 将成功回执绑定源码、依赖、引擎和分支

- 顺序/范围：测试证据；S1-F1、S1-F2、S2-F1、S3-F1、S3-F2、S3-F3、S3-F4、S3-F5、S5-F1。
- 当前：现有报告多无当前WIP源码hash；汇总passed不一定覆盖需要的回退分支。
- 建议：每次用户授权验证时记录输入分类（不存敏感值）、源码hash/HEAD+WIP、库/引擎版本、实际路由、故障触发、重试路径、关键断言和输出hash。
- 价值：下次能直接回答‘这个主路径/回退是否真的成功’，不重复宽范围审计。
- 边界：只对本次功能定向样本，不因证据不足强制全项目对拍/样式逐项比较。
- 需要的定向验证（本轮未执行）：每个fallback必须有触发日志+新副本+实际后备引擎+关键结果及源保护，负向报错另列。

## 后续定向验证清单（仅计划，本轮未执行）

无需为读取/新建/Rust规则功能安装Office；只有需要真实Excel公式重算的步骤依赖Excel或LO。下面测试需按用户后续授权执行，不自动运行Cargo、测试套件或程序。

### V01 S1-F1

- 普通模板公式真实Excel/pywin32重算→公式与CF异常合流13列
- 外部C4起点公式引用另一sheet，再被模板引用；最终数值、位置、公式与输入hash
- 动态INDIRECT/OFFSET全外部范围兜底、源sheet冲突、缺依赖
- 现代错误token从重算校验一直到抽取/CF/扫描
- 共享/数组公式写前拒绝或已支持路径的独立验收
- GUI副本与CLI原缓存不同的反例
- 一文件成功后下一文件失败、未知CF和最终输出错误

- 主路径现有依据：E02旧合成有记录，最新外部公式修复U。
- 回退现有依据：动态范围/现代token兼容未找到独立成功回执；真实openpyxl/跨引擎回退无生产成功。

### V02 S1-F2

- 任意行/固定行、同名合并、历史追加
- 未匹配sheet触发兜底模板两种模式
- 坏表头跳过、正式读取失败、公式缓存缺失
- 输出读回与输入hash

- 主路径现有依据：E08/E09历史子集R，真实批次partial不能全称通过。
- 回退现有依据：E08业务兜底历史R；reader/writer无生产后备。

### V03 S2-F1

- 本上期正常/异常、重复同值/异值、缺期/空/零
- 外部静态正常/缺映射/不支持格式
- 真实时序外部与单位例外
- 三表16/16/12列读回，坏文件/规则错误的状态

- 主路径现有依据：E12/E13/E14历史子集R。
- 回退现有依据：未选外部是未执行分支，不是回退；无读取/写出后备。

### V04 S3-F1

- GBK、UTF8/BOM、替换解码三个分支分别回执
- 14列、日期冲突、非数值、重复key/拼接碰撞
- 单位及5段/8段表达式、仅一时期/零分母
- 22列比较文件+日志与输入保护

- 主路径现有依据：E15真实子集R；当前完整人工验收无。
- 回退现有依据：编码分支未各自证明成功，替换损失需业务策略。

### V05 S3-F2

- 人工输出说明标记兼容真值/全未选
- 真实机构/地区映射、多地区拆分
- 8列说明文件+生成信息，F3四键稳定

- 主路径现有依据：E16历史非空2610行R，E30只0行。
- 回退现有依据：无替代reader/writer；未选行不是回退。

### V06 S3-F3

- Windows唯一原因合法写回，目标旁备份字节一致，目标原路径更新
- 未匹配/多原因/目标多行冲突与已有原因覆盖
- 执行中hash变化/目标被Excel锁/备份失败
- 候选安装失败→恢复成功；安装+恢复均失败留可恢复路径
- 成功但清理失败仅warning，原备份保留

- 主路径现有依据：E17只计划R；E04 Linux原地R，Windows旧E03为副本。
- 回退现有依据：回滚是恢复失败事务不是导入成功；Linux历史H，当前Windows未证。

### V07 S3-F4

- 本期CSV/Excel及CLI可选上期
- scope/前提/左右式失败、校验公式错误、反转/抑制/双0/epsilon
- 17列结果+日志，错误规则计入未完成

- 主路径现有依据：E18历史214结果/3日志R；不代表所有规则正常。
- 回退现有依据：CSV编码分支同F1，规则错误不交Excel回退。

### V08 S3-F5

- 真实3.3启用表单完整映射、3.0目标单位来源
- 表头仅改名/实际重排、缺映射/重复key/输出名碰撞
- 模板含definedNames且删列、公式/CF/DV/合并/打印关键结构
- 多机构前成功后失败及全部成功目录发布

- 主路径现有依据：E03/E04低层R、E19合成H，真实全量未证。
- 回退现有依据：固定位置兼容未证；无生产openpyxl后备。

### V09 S5-F1

- 多机构同模板/业务/期间批次，部分/全部失败
- 物理路径hash去重、清洗名碰撞、副本链接/owner
- 生效版本+显式source，外部无映射/非法标记
- 上期按机构缺失/重复；纯本期与V_PRE失败粒度
- 模板数值区域与整sheet兼容、结构歧义
- 非空多机构CF、模板CF区域边界、PARTIAL状态
- 源校验公式错误缓存扫描及CLI无模板SKIPPED
- 批次只读快照、owner/no-evidence、不污染基线
- 写结果成功ledger失败和运行中输入变化使结论失效

- 主路径现有依据：旧E20-E28分片R，当前修复与新测试全部U。
- 回退现有依据：数值区域兼容有历史契约报告，不能当每条业务通过；CF区域默认与新桥U。

### V10 共用计算

- 隔离pywin32 COM独立成功、超时/Excel忙/SaveAs错误/进程回收
- LO真实公式成功保存和目标平台离线组件可用
- 经过业务允许后，主引擎不可用→另一引擎从原候选重算且关键结果一致
- 无引擎时明确不可用，其他不依赖计算的功能仍可执行

- 主路径现有依据：PowerShell Excel COM历史H；pywin32 N；LO成功N。
- 回退现有依据：当前根本没有跨计算引擎回退链，必须先实现并验收才可说成功。

### V11 候选openpyxl兼容

- 仅明确编辑缺口：主后端写前拒绝、候选完整契约合格后成功
- 写后失败/校验失败不可自动回退
- 公式文本/关键关系/CF/名称/源文件不变；再经真实Excel/LO计算（需要时）

- 主路径现有依据：E10/E11为旧不完整/缺陷记录。
- 回退现有依据：当前未登记生产后备，此计划不构成启用授权或成功证据。

## 历史证据索引

只引用报告状态、计数与能力范围；没有复制真实敏感工作簿的业务值。

<a id="e00"></a>

### E00 当前源码静态核对

- 等级：U。
- 文件：[20261028_152620_S5定向修复与S1F1外部公式保护.txt](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/GPT修改记录/20261028_152620_S5定向修复与S1F1外部公式保护.txt)。
- 可以证明：本次只读调用链及已有证据；上轮 S1/S5 新修复和7个测试未执行。
- 不能证明/限制：本轮没有运行任何业务程序、测试或编译；不能产生新的测试成功证明。

<a id="e01"></a>

### E01 12:28前工作区历史回归

- 等级：H。
- 文件：[workspace-test-final.log](F:/RustBuild/audit/workspace-test-final.log)。
- 可以证明：只读累计79组 test result：435 passed、0 failed、28 ignored。
- 不能证明/限制：汇总不定位每个细分路径，忽略项未执行；不覆盖15:26后的S1/S5修改。

<a id="e02"></a>

### E02 S1F1 Excel COM合成全链记录

- 等级：H。
- 文件：[20261028_122800_S1F1验收与跨平台交付收口.txt](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/GPT修改记录/20261028_122800_S1F1验收与跨平台交付收口.txt)。
- 可以证明：单独ignored集成用例1 passed；Excel16、真实计算、2条公式错误+1条CF异常、13列、输入哈希保护。
- 不能证明/限制：记录级证据；当前桥是PowerShell COM，不是pywin32；合成样本，无LO成功，无新外部公式依赖修复验收。

<a id="e03"></a>

### E03 Windows Umya三类编辑能力报告

- 等级：R。
- 文件：[umya_windows_x86_64_20261028.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/capability_evidence/umya_windows_x86_64_20261028.json)。
- 可以证明：S1F1公式/外部纯值、S3F5列移位与关键结构、S3F3独立副本写回样本。
- 不能证明/限制：不能证明S1外部带公式跨表依赖、当前S3F3原地备份事务、最新S5逻辑；有能力样本不等于任意真实文件通过。

<a id="e04"></a>

### E04 Linux Umya三类编辑能力报告

- 等级：R。
- 文件：[umya_linux_x86_64_20261028.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/capability_evidence/umya_linux_x86_64_20261028.json)。
- 可以证明：Linux x86_64有限编辑样本；S3F3记录先备份后原地更新，相关回滚测试覆盖由E02记录说明。
- 不能证明/限制：JSON写回步骤本身不含每个故障注入断言；不代表Windows原地事务已通过，不证明LO计算。

<a id="e05"></a>

### E05 S1条件格式/配置真实子集

- 等级：R。
- 文件：[s1_conditional_audit_config_real_windows_native_20261025.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_conditional_audit_config_real_windows_native_20261025.json)。
- 可以证明：status=passed；CF触发结果、扫描区域、结构映射、模板推荐等所列真实子集。
- 不能证明/限制：有限CF语法与样本，不是Excel完整显示引擎；不证明最新GUI全链/任意CF。

<a id="e06"></a>

### E06 S1公式结果提取真实子集

- 等级：R。
- 文件：[s1_formula_issue_extraction_real.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_formula_issue_extraction_real.json)。
- 可以证明：status=passed；旧存储结果到Issue字段抽取。
- 不能证明/限制：不证明本轮公式复制/重算；新错误token能否到达此读值入口需单独验收。

<a id="e07"></a>

### E07 S1历史分类

- 等级：R。
- 文件：[s1_history_classification_parity.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_history_classification_parity.json)。
- 可以证明：status=passed；新增/延续/解决、备注等所列字段。
- 不能证明/限制：历史合成契约；不等于任意历史配置、最新GUI已验收。

<a id="e08"></a>

### E08 S1F2命名区+兜底模板

- 等级：R。
- 文件：[s1_f02_named_region_summary_synthetic_parity_windows_native_20261025.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_f02_named_region_summary_synthetic_parity_windows_native_20261025.json)。
- 可以证明：主汇总9行；fallback_template=passed，任意行/固定行各4行，同名合并、兜底局部/全局元数据、输入未变。
- 不能证明/限制：列宽差异单列记录，不作业务阻断；这是业务模板兜底，不是Calamine失败后的读取回退。

<a id="e09"></a>

### E09 S1F2历史追加与真实批次限制

- 等级：R。
- 文件：[s1_f02_named_region_history_synthetic_parity_20261024.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_f02_named_region_history_synthetic_parity_20261024.json)。
- 可以证明：历史合成 status=passed_with_column_width_variance，历史追加字段比对。
- 不能证明/限制：另有真实个人批次报告status=partial，不能称全体真实模板已通过。

<a id="e10"></a>

### E10 S1旧openpyxl全链候选失败

- 等级：X。
- 文件：[s1_f01_openpyxl_semantic_parity_windows_native_20261027_v7_excel_com.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_f01_openpyxl_semantic_parity_windows_native_20261027_v7_excel_com.json)。
- 可以证明：status=comparison_incomplete；重算failed、读回失败，capability unverified、production_route_eligible=false。
- 不能证明/限制：既不是当前Umya路由，也不是可启用的成功回退；不追查普通样式差异扩大范围。

<a id="e11"></a>

### E11 S1旧openpyxl worker写通但有缺陷

- 等级：X。
- 文件：[s1_formula_copy_openpyxl_worker_e2e_windows_native_20261025.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s1_formula_copy_openpyxl_worker_e2e_windows_native_20261025.json)。
- 可以证明：status=worker_write_through_passed_with_new_style_table_defect。
- 不能证明/限制：写了文件不代表关键结构安全或全链通过；不能登记生产后备。

<a id="e12"></a>

### E12 S2本上期真实子集

- 等级：R。
- 文件：[s2_real_period_python_rust_parity_windows_native_20261027.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s2_real_period_python_rust_parity_windows_native_20261027.json)。
- 可以证明：status=passed；真实本/上期数据、三表输出、源保护所列范围。
- 不能证明/限制：external_rules_not_exercised；有业务异常不等于软件失败，测试通过也不代表数据无异常。

<a id="e13"></a>

### E13 S2外部数据真实子集

- 等级：R。
- 文件：[s2_real_period_external_python_rust_parity_windows_native_20261027.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s2_real_period_external_python_rust_parity_windows_native_20261027.json)。
- 可以证明：status=passed、summary_equal=true、输出差异0、输入8份未变；所选样本external_time_series=false。
- 不能证明/限制：不证明所有时序外部布局/缺损输入/恢复路径。

<a id="e14"></a>

### E14 S2规则合成契约

- 等级：R。
- 文件：[s2_rule_engine_parity.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s2_rule_engine_parity.json)。
- 可以证明：status=passed_synthetic_parity；表达式、规则取数与异常契约。
- 不能证明/限制：有限合成语法；不是Excel任意公式。

<a id="e15"></a>

### E15 S3F1真实子集

- 等级：R。
- 文件：[s3_f01_real_parity_windows_native_20261027.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s3_f01_real_parity_windows_native_20261027.json)。
- 可以证明：real_sample_parity=passed，overall_feature_status=migration_in_progress。
- 不能证明/限制：整体未验收；历史小浮点尾差和表达式覆盖限制仍按报告范围解释。

<a id="e16"></a>

### E16 S3F2真实比较簿导出

- 等级：R。
- 文件：[s3_f02_real_comparison_parity_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s3_f02_real_comparison_parity_windows_native_20261026.json)。
- 可以证明：status=passed；人工选择2610行、读回2610行、输入/配置未变。
- 不能证明/限制：不涵盖任意手工表头更改/目标格式；12:01 runner另一个样本selected_rows=0不能代替本证据。

<a id="e17"></a>

### E17 S3F3匹配计划/冲突报告

- 等级：R。
- 文件：[s3_f03_feedback_plan_synthetic_parity_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s3_f03_feedback_plan_synthetic_parity_windows_native_20261026.json)。
- 可以证明：status=passed；计划单元格与2条冲突报告。
- 不能证明/限制：excel_modifier_execution明确not run；这份报告不能证明执行写回、备份或回滚。

<a id="e18"></a>

### E18 S3F4真实规则及输出

- 等级：R。
- 文件：[s3_f04_python_rust_real_parity_windows_native_20261026_after_safety_v3.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s3_f04_python_rust_real_parity_windows_native_20261026_after_safety_v3.json)。
- 可以证明：status=passed；163启用规则，214行结果、3行日志、工作簿差异0，输入未变。
- 不能证明/限制：有限真实样本和当时配置；错误日志存在不能略去。

<a id="e19"></a>

### E19 S3F5编辑与CLI/API烟测记录

- 等级：H。
- 文件：[20261028_105500_生产能力路由与Release交付.txt](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/GPT修改记录/20261028_105500_生产能力路由与Release交付.txt)。
- 可以证明：S3F5合成CLI/API、输出表单、输入哈希；低层细节见E03/E04。
- 不能证明/限制：没有定位到当前真实3.3模板逐机构完整成功报告；有定义名称且删列的模板当前会拒绝。

<a id="e20"></a>

### E20 S5配置加载

- 等级：R。
- 文件：[s5_config_load_parity_windows_native_20261025.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_config_load_parity_windows_native_20261025.json)。
- 可以证明：status=completed；所列配置字段比对。
- 不能证明/限制：旧单绑定/配置子集，最新生效版本/显式source预检修复未运行。

<a id="e21"></a>

### E21 S5表达式运行时

- 等级：R。
- 文件：[s5_eval_runtime_windows_real_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_eval_runtime_windows_real_20261026.json)。
- 可以证明：80匹配、0差异，含3个错误案例。
- 不能证明/限制：有限DSL合成/真实参数子集；不证明GUI新批次/新公式扫描。

<a id="e22"></a>

### E22 S5 V_PRE CLI合成链

- 等级：R。
- 文件：[s5_vpre_cli_synthetic_parity_20261025.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_vpre_cli_synthetic_parity_20261025.json)。
- 可以证明：2条结果、当前/上期复制清单哈希一致、取数与差值比对。
- 不能证明/限制：不是GUI按机构批次绑定，也不是最新修复验证。

<a id="e23"></a>

### E23 S5旧输出子集

- 等级：R。
- 文件：[s5_outputs_real_template_dir_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_outputs_real_template_dir_windows_native_20261026.json)。
- 可以证明：status=completed、diff_count=0、inputs_unchanged=true。
- 不能证明/限制：明确CF和数据质量表为空，结构快照/哈希降级不在此报告范围；无非空CF输出成功证明。

<a id="e24"></a>

### E24 S5模板数值区域投影

- 等级：R。
- 文件：[s5_numeric_scope_template_projection_real_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_numeric_scope_template_projection_real_windows_native_20261026.json)。
- 可以证明：verified_subslice；1对真实样本4区域坐标一致，合成CLI含显式模板优先/歧义拒绝/兼容路径。
- 不能证明/限制：真实样本只证区域坐标，不证规则到输出整链；当前匹配和批次修复仍U。

<a id="e25"></a>

### E25 S5无模板数值范围兼容

- 等级：R。
- 文件：[s5_numeric_scope_real_audit_no_template_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_numeric_scope_real_audit_no_template_windows_native_20261026.json)。
- 可以证明：377行结果匹配；含207PASSED、162STRUCTURE_ERROR、1RETRIEVAL_FAILED、7TRIGGERED。
- 不能证明/限制：这是结果契约匹配，不是377条业务通过；无V_PRE/V_EXT、无输出/快照/哈希降级覆盖。

<a id="e26"></a>

### E26 S5结构快照旧契约

- 等级：R。
- 文件：[s5_structure_snapshot_real_parity_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_structure_snapshot_real_parity_windows_native_20261026.json)。
- 可以证明：status=completed，旧结构快照比较/持久化范围。
- 不能证明/限制：本轮需看的新‘批次只读、机构owner、没有证据不污染基线’不由旧报告证明。

<a id="e27"></a>

### E27 S5源文件变化降级

- 等级：R。
- 文件：[s5_input_integrity_downgrade_synthetic_parity_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_input_integrity_downgrade_synthetic_parity_windows_native_20261026.json)。
- 可以证明：passed_for_synthetic_integrity_downgrade_contract。
- 不能证明/限制：新多物理文件hash键、文件名碰撞、复制后逐份比对修复仍U。

<a id="e28"></a>

### E28 S5多外部歧义负向契约

- 等级：R。
- 文件：[s5_multiple_external_sources_ambiguous_synthetic_parity_windows_native_20261026.json](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/probe/reports/s5_multiple_external_sources_ambiguous_synthetic_parity_windows_native_20261026.json)。
- 可以证明：diff_count=0；CONFIG_ERROR=1、RETRIEVAL_FAILED=1与基线相同。
- 不能证明/限制：正确报错的负向测试，不是成功选择另一个外部文件；GUI单外部文件入口不等同CLI多源。

<a id="e29"></a>

### E29 Linux缺少LO与平台交付历史

- 等级：H。
- 文件：[20261028_122800_S1F1验收与跨平台交付收口.txt](C:/Users/AZI/Desktop/VibeCode/Rust-Audit/audit-core/GPT修改记录/20261028_122800_S1F1验收与跨平台交付收口.txt)。
- 可以证明：无soffice时明确报错；Debian glibc2.28等价容器安装/CLI/API等所列烟测。
- 不能证明/限制：没有LO成功重算；不代表麒麟V10/UOSV20实机GUI验收。

<a id="e30"></a>

### E30 12:01合成业务runner

- 等级：R。
- 文件：[acceptance-report.json](F:/RustBuild/audit/acceptance-runner-synthetic-0b813b7c867b4f979e155fccad36f151/acceptance-report.json)。
- 可以证明：S3F1 completed rows=1；S3F2 completed selected_rows=0；S5 completed triggered_count=0；所列输入hash未变。
- 不能证明/限制：F3/F4/F5均not_requested；不是九项验收完成；不能证明有选中行导出或异常批次。

## 本轮交付边界

本轮只新增本文、结构化JSON、交互清单和修改记录；没有更改业务源码、配置、菜单、失败策略或测试代码，没有更新dist，没有运行Cargo/编译/测试/应用，也没有commit/push或创建分支/Worktree。仓库既有大量未提交业务实现与修复全部保留。

建议你先决定S1F1-E18（单文件失败是否丢弃整批）、S1F1-E16/G09（不完整CF是否阻断）、S3F5-E01（表头位置兼容）、S3F5-E07（定义名称+删列）、S3F5-E08（单机构失败范围）、S2F1-E04/S3F1-E06（重复值冲突）、S5F1-E06/E12/E15/E17（缺上期/整sheet范围/旧公式缓存/CF受限），再开展针对性改动。S3F3备份和3.3单位已有用户决定，不重新提问。
