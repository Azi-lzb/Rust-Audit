# Open-Data-Audit 功能迁移清单

### 2026-11-13 后续验证：S2 数值安全与桌面部署

后续验证补充，不覆盖下方各轮历史记录：S2 表达式/规则/单位换算针对性测试分别 9/5/6 项通过；Windows 配置保留与旧版次标记处理单测各 1 项、发布预检 9 项通过。当前 Windows Release 已构建并交付全量 ZIP，内置 Python 3.12.10/Openpyxl 3.1.5，隔离导入、内存工作簿数值/公式文本读写及 worker 启动通过，来源 32 份配置/模板 SHA256 未变；不是公式重算/GUI 全业务重新验收。Windows 新版默认资源目录为 `config`，`config-real` 仅补缺，打包不覆盖实际配置；EXE 内置全量/精简版次，忽略旧 edition 标记，包不分发该标记或快捷方式脚本。Linux 延续现有版次标记，配置默认资源与实际文件分离，保存/恢复/重试/保存失败测试有 WSL 证据，真实 dpkg/麒麟仍 NOT_TESTED。内置 LO 模块已通过 Windows 类型编译与 Release 构建，实际 LO/批量/超时/进程清理及 Linux 集成仍 NOT_TESTED，不能继承旧 sidecar 通过结论。精简版 feature 的独立 Release 尚未构建验收。详细文件与证据见 GPT 修改记录 20261113 各任务记录。

### 2026-11-13 增量：独立 LO 辅助项目移除

按用户要求删除 sidecar/lo-convert 四个版本控制文件，LO 批量重算与必要输出校验迁至 src/libreoffice_worker.rs / libreoffice_batch.rs；原重算集成/吞吐测试迁至 tests/libreoffice_worker.rs。GUI/Linux 改由本程序内部 worker，Windows COM 热会话保持。独立格式转换/旧 PowerShell 辅助实现随该未接入正式 GUI 的工具项目删除；不扩展 Audit 的格式转换功能。实现及静态检查与运行验收分开：本轮未编译/测试/打包，旧证据不继承为当前通过。历史性能脚本保留，仅供旧版本证据，不再用于当前二进制验证。

> 当前产品仅保留 S1/S2/S3。S5 UI、CLI、API、业务模块和专属测试代码已按用户决定移除；下文 S5 条目均为带日期的历史开发/验收记录，不代表当前产品仍包含 S5。

> **历史证据档案（截至 2026-10-28）**：以下条目和状态是原工作区不同时间的记录，不作为当前进度或操作指令。当前功能状态和验收口径以 [独立项目开发目标](../DEVELOPMENT_GOAL.md) 及 [V2 需求](AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md) 为准；本文件保留旧证据，不自动要求全面对拍或编译。

### 2026-11-06 增量：S1-F02 全量样本与旧 Python Audit 对齐

对 `测试数据/S1F2全量` 的 129 本嵌套源工作簿，调用旧 GUI 的默认 `dag:汇总校验结果说明` 工作流与当前 Rust S1-F02 核心，以同一旧版汇总模板和历史配置运行。初次发现 Rust 与 Python 的工作表/字段顺序差异；现在 Rust 按 Python 的全局/工作表级定义名称迭代顺序和模板字段顺序处理。最终四张表、表头、全部单元格及 505 行数据完全一致；旧 Python 结构日志与 Rust 预检均为缺表继续，整本跳过 0。定向 Rust 单测和 CLI 集成测试各 6/6，Release 构建成功，132 个输入/参考文件哈希未变。完整证据、输出行数和遗留模板名称问题见 [S1-F02 全量对照报告](probe/reports/s1f2_full_sample_parity_20261106.md)。

未解决的业务问题：28 本源文件有“在线抽验结果”表，但旧版模板定义“在线校验结果”；旧 Python 与 Rust 都输出空“在线校验结果”。测试目录既有 `汇总信息.xlsx` 则含“在线抽验结果”表，结构与两端现跑结果不同。本轮保持旧 Python 的实际行为，没有擅自新增表名别名；需要业务方确认模板版本或名称映射。未运行全项目测试或 GUI 手动验收。

### 2026-11-06 增量：S1-F1 当前模板 Python 全量复跑与复制范围纠正

旧 Python 使用当前 19 份模板，对相同 129 个源工作簿、正式逐笔配置和大集中外部数据按 19 个业务批次复跑，19/19 批次完成、129/129 文件成功。19 个批次配置副本的 SHA-256 均与旧/新正式配置一致；源、模板、配置和外部数据共 168 个输入哈希前后不变。对照当前 Rust Release：Python 3,616 行、Rust 3,722 行；Python-only 0，Rust-only 106；3,611 个共有结果所有业务字段一致，另 5 个条件格式 `check_field`/规则编号只有 Python 追加 `#DIV/0!` 的元数据差异。60 个字符串形式数值显示差异经 Decimal 对比相等。106 个 Rust-only 键集合与先前完整对拍相同：95 条在模板校验区域且旧 Python 提取函数可重现 `IndexError`；11 条票据公式按用户确认的精确 `<>0` 触发，Python 副本相应缓存为空，Rust 副本缓存约为 `±3.63797880709171E-12`。完整对照和哈希见 `F:/RustBuild/audit/tmp/s1f1-python-current-template-129-20261106-run2/`。

用户澄清：模板“校验区域”外的数值、文字和公式可为规则验证时临时修改的测试数据，不应复制或提取；表结构/表头检查独立执行，条件格式审核来自源核查表。旧失败的 `存量同业借贷!F45` 是旧整簿公式保护比较报出的区域外位置；旧候选未保留实际前后公式文本，不能断言具体序列化差异。当前源、模板和 Rust 副本 F45 公式文本一致。S1 复制阶段已调整为只将导入命名区域公式差异作为阻断条件，区域外既有公式不再拦截；定向测试 1/1 通过。本轮 Rust 全批 Release 运行早于此源码调整，尚未以新代码构建后全批复验。

### 2026-11-06 增量：S1-F1 真实 129 本业务运行、Python 对照和后端选择（当前状态）

按用户澄清，S1-F1 的用户自定义校验内容以配置映射的“校验区域”命名区域为准，复制单元格值/文字、公式文本、批注（位置/作者/内容）并覆盖目标区域空白；模板校验区域外的人工规则试验值不复制、不扫描。条件格式审核来自源核查表；表结构/表头检查独立执行。当前真实 129 本使用命名区域模板，未覆盖结构化“审核规则”模板分支。

当前源码构建后的完整 Rust 运行：129/129 匹配与结构通过、129 本重算成功、0 跳过、0 待人工重算、0 失败；公式异常 3,373、条件格式异常 349、合计 3,722；生成 129 个审核副本与汇总工作簿，耗时 1,819,953 ms。157 个输入文件未改变。输出目录和机器报告见 DEVELOPMENT_GOAL.md 最新增量。129 个副本的配置校验区域与模板逐格对照无差异，覆盖 16,898 个公式与 896 个文字/常量及空格。旧 Python 输出在 21 本中少 42 个命名区域末尾公式。

旧 Python 同输入结果为 3,616 条。Python-only 0、Rust-only 106（均为公式异常），共有 3,616 条中 3,611 条完整字段一致；另外 5 条条件格式 check_field 仅有旧 Python 多出的 #DIV/0! 后缀。349 条条件格式触发及说明一致。106 条中 95 条涉及 13 份 Python 审核副本/提取失败样本，复跑旧 Python 提取函数可复现 IndexError；其余 11 条是票据模板以 <>0 精确比较时 Excel 计算的约 ±3.63797880709171E-12 浮点残值，Python 缓存为空。未添加容差或删除 Rust 发现；业务人员需确认票据模板逻辑，并抽查旧提取器失败样本。旧副本清除公式引用的静态标签/明细数据也会导致结果变化，需以源数据及校验区域规则判断业务正确性。

外部表依赖仍有需按模板核对的语义差异：Python 按模板全簿公式选择外部表，Rust 从校验区域公式递归追踪并读取选中表 UsedRange。若仅区域外公式引用辅助表，两端可能导入范围不同；本轮成功执行不等于已证明该边界完全等价。

S1-F01 默认修改路由为 OOXML Patch 优先、Umya 按能力兜底；高级设置可由用户显式选择 openpyxl 复制后端，失败不自动静默切换，写后结构/公式校验保留。Windows Release 已打入 worker 脚本，但不包含 Python/openpyxl runtime；需 PATH Python 可导入 openpyxl。Linux 仍要求随包提供隔离 runtime。129 本完整批次走默认 OOXML，不代表 openpyxl 全批验收；openpyxl 与 OOXML 的真实单本候选均有通过写后校验的样本，合并区域非锚点数据等能力边界会拒绝而不发布。

验证：excel-editor 64/64、excel-reader 24/24、audit-core 229 passed / 0 failed / 12 ignored；audit-core CLI cargo check 通过；Windows Release 构建通过。未运行全 workspace 测试、GUI 人工点击、Linux/Kylin/UOS 实机验收。分类清单保存在本机临时目录（19 类、129 本），原始核查表/模板未移动或修改；聚合对照报告见 audit-core/probe/reports/s1f1_real_sample_parity_20261106.md。
### S1/S2/S3 性能地图索引（2026-10-30）

当前可用于优化决策的功能级基线、数据规模、正确性证据和缺项汇总见 [PERFORMANCE_BASELINE.md](PERFORMANCE_BASELINE.md)；逐次记录及历史口径见 [PERFORMANCE_HISTORY.md](PERFORMANCE_HISTORY.md)，可复跑入口索引见 [benchmarks/README.md](benchmarks/README.md)。该台账按 S1-F1/F2、S2-F1、S3-F1 至 F5 建立；S5 排除。S1-F1 原生 COM 批次与 S3-F1 性能项目按当前用户目标暂结；S1-F1 条件格式异常指标定位、结构预检/条件审核共享源读取及模板结构投影读取均已定点优化。4k 结构预检+条件审核 CLI 中位数为 390 ms；20k 结构模板读取为 528→224 ms。完整业务链仍待计时。

S1-F1 源读取复用使用单个输入工作簿内的只读 `WorkbookDump`，不跨批次或重算阶段缓存。4,000 行、8,610 条异常的五轮测试中，Rust CLI 内部事件耗时（不含进程启动）从两次读取状态的 498 ms 降至 390 ms，异常摘要对拍 5/5，峰值工作集与前态相当。详细证据见本文件末尾对应增量及 `PERFORMANCE_HISTORY.md`。

台账复核发现 2026-10-30 汇总 Markdown 中部分 S2-F1、S3-F1/F2/F3 数字与逐轮 JSON 不符，基线使用 JSON 逐轮值重新计算中位数并保留不一致说明；未改写历史报告。大多数性能样本未记录输入字节数、统一阶段计时或峰值内存，未验证全单元格业务一致性的功能也明确标出。性能数据不得替代人工业务验收。

补充：S3-F2 另有一轮新的规模扩展剖析，使用合成 7,830 行比较表、2,610 条选中输出；三轮完整工作簿对拍差异均为 0，Rust CLI 墙钟中位数 359.88 ms、Python 2,919.43 ms。详情见 `tests/compare/perf_s3f2_scale_20261030.py` 与 `probe/reports/s3f2_scale_profile_20261030.json`。这不是不可得的旧真实工作簿重测，不据此改动业务实现。

### Windows 原生 COM 与 lo-convert 退出默认运行链（2026-10-30）

S1-F1 Windows GUI/CLI 现使用 Audit 自身的一次性批次工作进程直接调用原生 COM，批次内复用 Excel/WPS；工作进程隔离用于限制 COM 卡死对 GUI 的影响。缓存清理、计算属性设置、公式/缓存读回验证和不覆盖发布保留；自动模式 COM 失败时由 Audit 直接调用 LibreOffice，失败则走人工重算恢复。Office 进程以 PID 和创建时间核对并持有句柄，超时仅终止已验证的进程。Linux 仍使用现有 LibreOffice sidecar。可选的旧格式转换工具保留源码，其 Windows COM 转换及 Word 文本校验移除生产 PowerShell 调用；Windows 标准构建/打包不再要求或包含 lo-convert。打包脚本本身仍使用 PowerShell 压缩，已有 dist 未修改。旧 API 仍保留显式 sidecar 兼容入口，不代表默认 GUI/CLI 会调用它。

此次仅做静态检查和 Cargo 离线 metadata，未执行编译、测试、打包、Office/LibreOffice 或性能计时；此前 sidecar 实测不能作为新路径通过证据。必须重新验证 Windows 真工作簿重算、批次复用、超时清理、Excel/WPS/LibreOffice 分支以及可选格式转换。

### S1-F2 早期性能缺口记录（已由 2026-10-30 复测补齐）

早期综合性能驱动没有 S1-F2 计时，且合成源表未满足 Rust 结构预检，不能据此判断性能或业务等价。随后修正合成夹具并单独运行三轮：输出值与结构检查通过，列宽存在不超过约 0.54 的差异。当前结论及限制见“Windows 新旧性能复测”和 `probe/reports/perf_comparison_20261030.md`。

### Windows 新旧性能复测（2026-10-30）

在 Windows x86_64 对当前 Release Rust CLI 与只读 Python Audit 做了每项 3 轮复测，主表使用中位数。Python 计时不含解释器启动；Rust CLI 墙钟计时含约 21.8 ms 启动。负载主要为合成夹具；S2 使用相同仓库测试夹具。倍率定义为 Python/Rust。

- S1-F1 COM 重算子步骤（2 册 × 7,200 个 VLOOKUP）：Python `Application.CalculateFull` 3,547.7 ms，Rust 原生 COM worker 4,950.6 ms，0.72x。worker 两册间复用同一 Excel PID；6 个输出样本每册 7,200 个公式及 7,200 个数值缓存均核验通过，输入未变。该数据不含完整 S1-F1 业务链路，Python 旧 `Worksheets.Calculate` 中位数为 3,513.0 ms。
- S1-F2 命名区域汇总（两册、三张源表、九行输出）：Python 4,320 ms，Rust 59 ms，73.22x。三轮输出值与结构对比通过，列宽最大差约 0.54，输入未变；合成小样且 Python 使用 Excel COM，不能外推为通用加速比。
- S2-F1 完整服务和工作簿导出：40.0/52.0 ms，0.77x；验证两个输出可读、表名和维度一致，未逐格比较所有值。
- S3-F1 10k 行 1,064.1/1,432.5 ms，0.74x；S3-F2 5k 行（1,667 导出）1,853.2/234.6 ms，7.90x；S3-F3 1k 行反馈导入 934.0/980.9 ms，0.95x；S3-F4 10k 行 6,053.9/937.5 ms，6.46x；S3-F5 30 机构 × 40 指标 × 4 表 3,290.8/868.5 ms，3.79x。S3-F3 的导入行数、目标备份字节一致和反馈输入未变均检查通过；其余这些项目的本轮计时未做完整的值级结果对拍。
- S5 单文件 + 独立上期小负载：24.4/72.6 ms，0.34x。仅性能探测，没有恢复或修改 S5 开发。
- 本轮 Release 构建成功；`s1_native_com_session` 定向 COM 测试 1/1 通过。没有全量测试、真实用户样本验收或 Linux 计时。主报告、驱动和机器数据：[`perf_comparison_20261030.md`](probe/reports/perf_comparison_20261030.md)、[`perf_compare_20261030.py`](tests/compare/perf_compare_20261030.py)、[`perf_comparison_20261030.json`](probe/reports/perf_comparison_20261030.json)。S1-F2 三轮报告分别保存在 `probe/reports/s1f2_perf_run{1,2,3}_20261030.json`。

### S1-F1/S3-F1 性能优化实施状态（2026-10-30）

- **S1-F1 Windows 重算**：公式重算通道已从 PowerShell 脚本切换到 `lo-convert` sidecar 内原生 `IDispatch` COM 自动化。GUI 的 S1-F1 批处理和 S1 CLI 共用一个持久 sidecar session/Office COM 实例顺序处理工作簿；按文件保留独立的缓存清理、Excel 完整重算、重算结果结构/公式/缓存验证及不覆盖发布。超时会终止 sidecar，并且只针对由 session 新建的 Office 进程清理；之后的工作簿可重新建会话。一次只重算单个副本的入口也由原生 COM worker 执行。文件转换流程中的 PowerShell COM 调用不在本次修改范围。
- **S3-F1 写出**：日期、填色文本、数值和空白单元格的格式对象已移出逐行/逐格循环并按实际颜色与数值格式缓存复用。`validate_execution_compare_workbook` 的完整工作簿回读和逐项核验保留。
- **未验证边界**：本次没有执行 Cargo 编译、测试或性能计时；代码需要后续在授权环境编译。2026-10-29 的对比是修改前基线，不可作为本次优化后的实测结论。热会话节省的 Office 实例启动时间和 S3-F1 写出改善量均待同输入复测。
- **其他差距**：S2 小样本差几十毫秒，主要由进程启动和额外留痕占比造成，当前不建议删减校验或留痕。S5 配置审核根据用户 2026-10-30 指示暂停开发；不做性能更改。

### S2 对拍覆盖判定与 AL-10 语义分歧（2026-10-29 补充）

- **S2 三张结果表对拍覆盖判定**：既有 `s2_*_parity` 链构成对 D.2“S2 三张结果表合成对拍”的同等或更强覆盖，判定依据——① 每阶段 Python 基线（`run_s2_*_baseline.py` 直接导入真实 Python 服务 `base_audit.systems.s2_report_collection` 的 importer/period/external/rule/validation/exporter 引擎，对共享夹具生成基线）+ 对应 `s2_*_parity.rs` 深度 JSON 比对，全部在 501 回归中通过；② 导出器对拍 `s2_exporter_parity.json`（passed，differences=[]）对三张结果表工作簿（环比规则/外部核对规则/校验规则）逐格比对，Python 侧 `build_summary` 基线带源码 SHA256；③ 真实输入端到端对拍 `s2_real_period_python_rust_parity_windows_native_20261027.json` 与 `s2_real_period_external_python_rust_parity_windows_native_20261027.json`（均 passed）：真实机构报表经 Rust CLI 与 Python 服务各自产出三表工作簿并统计比对（后者 7 机构/3 表单/678 指标/1007 发现：环比 191、外部核对 226、规则 3、校验规则 15）。故未新写 `compare_s2_outputs.py`；该判定与理由作为正式记录。
- **基线新鲜度实测**：`run_s2_rule_engine_baseline.py` 从当前 Python 源重生成基线，与已提交版本**完全一致**——当前 Rust == 当前 Python（规则引擎阶段）。分析阶段基线重生成发现真实语义分歧（见 AL-10）。
- **AL-10（新增，需用户决策）**：S2 外部取数告警口径分歧——旧 Python 将外部取数告警转为软性“校验规则”发现（Q002）；**当前 Python service.py 已改为仅运行诊断、不纳入审核发现**（`validate_external_warnings` 已删除）。Rust 跟随旧口径（与已提交基线一致）。同一夹具实测 9 项发现（旧）vs 8 项（当前 Python），唯一差异 Q002“外部导入提示”。基线生成器已改双 API 兼容；当前契约基线存 `probe/reports/s2_analysis_python_baseline_current_20261029.json`（不覆盖旧契约基线）。统一方向待用户决定。
- 分析基线生成器 (`run_s2_analysis_baseline.py`) 已更新为双 API 兼容（旧 `validate_external_warnings` / 当前 service.py 契约），并在生成的基线中标注所采用契约。

### 性能差距分解（2026-10-29 补充）

`probe/reports/perf_gap_analysis_20261029.md`：原始基线将 S3-F1 差距定位为写后全量回读验证 + 循环内每格 `Format::new()`（Rust 写入阶段 1520ms 占 76%，配对计算 121ms）；S1-F1 旧路径使用 PowerShell COM。2026-10-30 已实现格式对象复用、保留全量验证，以及原生 COM + 批次热会话。此次改动未重新编译、测试或计时，轻量验证没有实施，性能收益待新一轮同输入复测；S5 配置审核按用户指示暂停，S2 不建议为几十毫秒差异削弱留痕或校验。

### 性能定点优化与复测（2026-10-30）

`perf_gap_analysis` 定位后的定点优化（均保持业务语义）：S3-F1 验证阶段去除
二次整簿解析（公式检查改轻量 XML 扫描），验证 936→417ms、总耗时 1412→894ms
（10k 行），对 Python 1032.5ms 反超至 1.02x；S1-F1 原生 COM 通道去除冗余
Calculation 模式设置（~150ms/册）、改为逐工作表 Calculate + SaveAs 后
Close(False)、DISPID 按名解析缓存 + METHOD|PROPERTYGET 组合标志
（Worksheets.Item 等成员需要），4950.6→4668.5ms。剩余 S1-F1 差距（0.77x）
经阶段级实测确认为 Excel 引擎固有成本（实例创建 ~2.3s 两端相当、Workbooks
Open 过滤器加载、逐表 Calculate 重算本身），非 Rust 实现缺陷；证据与分解见
`probe/reports/perf_comparison_20261030.md` 的优化迭代节。验证：全量回归
501 passed / 0 failed / 33 ignored；`s1_native_com_session` 真实 Excel 门通过
（7,200/7,200 缓存、进程复用、输入保护）。原生 COM worker 阶段计时
（stages_ms）与会话创建毫秒已入 worker 协议供后续复测。

### 新旧性能对比基准（2026-10-29，应用户要求）

`tests/compare/perf_compare_20261029.py`（驱动）+ `probe/reports/perf_comparison_20261029.json/.md`（结果）：同一合成输入分别经旧 Python 服务（进程内计时）与新 Rust CLI（子进程墙钟）运行，3 轮取最优，输出逐项校验。结果：S3-F4 校验公式求值 6.31x、S3-F2 说明导出 8.64x、S3-F5 表单批量生成 2.28x（Rust 快）；S3-F1 0.75x、S1-F1 重算子步骤 0.74x（口径差异，见报告解读——S3-F1 为工程优化空间，S1-F1 差异来自 Rust 侧附加的缓存清理/校验/发布，Excel 引擎本身占主导）；S5/S2 小负载下 Rust 含启动约 20–40ms 差。S1-F2 未计时（两端语义不等价）。合成负载单机 Windows 实测，不构成 Linux/麒麟结论。

### S1/S2/S3 对齐决策清单与引擎高级设置（2026-10-29）

本节登记 2026-10-29 对齐 Goal 的产出，并取代下方 2026-10-28 条目中已被后续工作推翻的三处陈述（S3-F3 输出契约、3.0–3.3 检查 API 缺口、编辑能力门控）。当前状态仍以 [DEVELOPMENT_GOAL](../DEVELOPMENT_GOAL.md) 为准。

- **对齐决策清单**：`docs/ALIGNMENT_DECISION_LIST_20261029.md/.json` 把失败策略矩阵的 104 个失败场景分为三类——96 项建议直接采纳（Rust 安全失败关闭与矩阵建议一致）、6 项需用户业务决策（G12、S1F1-E05、S1F1-E16、S1F1-E17、S1F1-E18、S3F4-E05）、2 项纯 S5 场景按用户口径移出；另列 9 条非失败策略对齐空间（AL-01…AL-09），其中 AL-04（S3-F1 数值口径：旧 Python 基准即 IEEE-754，不建议 Decimal 化）已完成评估待确认。实现策略在用户答复前不落地。
- **引擎高级设置（Rust 实现）**：设置中心“高级”页新增“Rust 引擎（高级）”卡片并持久化到状态目录 `settings.json`（`AUDIT_STATE_DIR` 或用户状态目录）。计算引擎四通道：auto（Windows=Excel COM→WPS 表格探测；Linux=LibreOffice 无头）/ 强制 Microsoft Excel（COM）/ 强制 WPS 表格（COM，ProgID 探测 ket.Application→KET.Application）/ 强制 LibreOffice 无头（经 lo-convert sidecar）；sidecar `lo-convert recalculate` 新增 `--engine excel|wps`，`com`/auto 的 COM 阶段脚本由 Excel-only 改为按序探测三 ProgID。Excel-Core 修改引擎两通道：Umya（默认，已验证）/ openpyxl（高级兼容路径）。强制引擎不可用时错误含引擎名与原因，不自动切换；S5 不接入该设置。相关模块：`app_settings.rs`（存储）、`excel_route.rs::production_modifier_with_edit_engine`（证据驱动注册）、`s1_recalculation.rs`（engine 透传与 WPS ProgID 校验）、sidecar `main.rs`。
- **openpyxl 修改引擎负面能力证据**：`capability_evidence/openpyxl_windows_x86_64_20261029.json`（available=false，本机 Python 3.12/openpyxl 3.1.5 worker 实测）——openpyxl 重写无法满足三个业务批次的生产保留契约：S1-F01 批次在“OOXML 关系改变超出允许集”被后端自校验拒绝；列尺寸被归一化（与旧 Python openpyxl 基线一致，故不声明 RowAndColumnDimensions）；S3-F03 计划契约（列尺寸/包部件）与 S3-F05（DeleteColumn）在路由层失败关闭且输入字节不变。GUI 选项据此置灰并显示原因；未来后端补齐契约后重跑 `AUDIT_OPENPYXL_CAPABILITY_REPORT` 即可解锁。注册读取嵌入证据（include_str!），仅在 available=true 且平台匹配时路由。
- **S3-F02 Python/Rust 对拍通过**：`probe/reports/s3_f02_synthetic_parity_20261029.json`——合成样本 4/4 行选中一致、输出逐项比对 differences=[]，9 项检查（选中行与值、地区/机构通配映射、表头与比较值不舍入、输出值/公式文本/有效样式/命名区域、表顺序/合并/冻结/筛选/行列尺寸、页边距、元数据内容与时间戳格式、CLI 双视图与耗时、输入保护）全部通过；Python 侧为只读 `export_selected_explanations` 基线（Python 3.12.10）。该证据属于合成样本对拍，不替代真实机构业务验收。
- **S1-F3 命名澄清（编号漂移）**：Rust 内部“S1-F3”是 S1-F01 人工重算恢复 worker（`s1_recovery.rs`，GUI 按钮“S1-F3 从审核副本汇总校验核查表”，状态机 Copying→WaitingManualRecalc→…→Completed），服务于 S1-F01 业务链的跨进程恢复；与 Python 菜单中暂缓的 S1-F3（组合工作表/联合模板）**不是同一功能**，属内部步骤编号与 Python 菜单编号撞名。暂缓的 S1-F3/F4/F5 未被开发。是否重命名待用户确认（决策清单 AL-05）。
- **取代 2026-10-28 陈述一**：S3-F3 输出契约已按用户确认改为“先备份、原地更新 + 失败回滚”（`execute_explanation_feedback_import`），旧“独立输出副本”描述仅保留于 `execute_explanation_feedback_import_to_copy`（显式请求副本的兼容入口）。
- **取代 2026-10-28 陈述二**：Web bridge 现已注册 3.0/3.1/3.2/3.3 全部四个只读配置检查 API（`check_central_common_config` / `check_central_expression_config` / `check_central_cross_formulas` / `check_central_forms_config`），`NOT_IMPLEMENTED` 缺口已闭合。
- **取代 2026-10-28 陈述三**：桌面运行层的 Excel-Core Passed 编辑能力已按 windows-x86_64 证据注册（`excel_route.rs`，Umya），S1-F01/S3-F03/S3-F05 编辑批次不再无条件失败关闭；Linux 证据另于 2026-10-28 登记为通过。
- **口径固定**：S1-F01 校验区域扫描双口径（CLI 扫源文件缓存值 / GUI 扫重算副本）由 `tests/s1_scan_scope_contract.rs` 固定，统一口径待用户决定（AL-03）。原文误标为 S1-F02，已勘误。
- **对拍器维护项**：`tests/compare/compare_s1_region_summary.py` 夹具落后于当前 CLI 表结构预检（合成样本表头 0/3 被跳过），对拍待夹具更新后重跑。

### S3 五项业务逐源核对与映射修正（2026-10-28）

以原 Python `core/frontend/web/index.html` 的五个 F1–F5 按钮、`core/src/base_audit/web_app.py` 的回调、`systems/s3_central_statistics/` 服务/导入导出模块和 `config-real/` 的 3.0–3.3 工作簿复核编号、输入、规则和输出。正式对应关系是 Python F1/F2/F3/F4/F5 → Rust `s3_execution_compare` / `s3_explanation_export` / `s3_explanation_feedback` / `s3_current_value_check` / `s3_indicator_split`；不把 Python F4/F5 改称 Rust F2/F3，也不把 F2/F3当作额外功能。详细逐项表与所需字段见 [V2 需求的 S3 对照表](AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-f1-至-f5-原-python-与-rust-业务对照)。本次只改业务范围和迁移状态文档，不改 S3 产品逻辑。

Python 的可见 GUI 均有五项操作入口。服务层另有未由 GUI 暴露的 F1“每期多文件合并”、F4“可选上期文件/`{...}` 取数”，F5 另有未由正式 GUI 调用的独立模板参数；不能把这些服务兼容参数写成 GUI 必需输入。原 F5 的正式配置把金融模板内嵌在 3.3 中；只读检查 `config-real/3.3大集中指标比较拆分_配置.xlsx` 的报表清单共有 43 条启用报表，`config/` 中空的占位清单不能作为实际配置状态。

Rust 现有实现保留并按 Python 业务重新映射。需要防止再次误编号的历史测试资产包括 `crates/excel-editor/tests/s3_f3_edit_operations.rs` 与 `s3_indicator_split.rs` 中 `AUDIT_S3_F3_*` 环境变量：检查其中操作批次为表单结构/格式编辑，归入 Python S3-F5；Python S3-F3 的当前代码对应 `s3_explanation_feedback`。测试文件名与部分环境变量仍保留旧 F3 标识；表单编辑测试函数/注释已按内容映射为 Python F5，测试源码于 2026-10-28 有补充且未运行，不能作为生产写入能力或验收证据。

另一个重要输出契约差异：Python S3-F3 在创建带时间戳备份后更新目标说明文件；当前 Rust 仅规划在独立输出副本上写回，且桌面运行层没有 Passed Excel-Core 编辑能力，存在写入时会失败并不发布结果。该差异列为 F3 后续补齐项，不在本轮改写现有实现。

配置核对还发现 3.3“转表设置”含“分析文件数据单位”，Python 配置加载器会保存该参数，但 `service.render_financial_forms` 只读取“隐藏空行”和“空表删除”，渲染器不使用该单位字段；F5 当前使用 F1 比较结果中的数值（已按 3.0 默认目标单位生成）。Rust `s3_indicator_split` 也只读取前述两个开关。该字段应作为未确认/当前非运行参数记录，不能仅凭配置表头扩写为已确认业务规则。

F5 输出的表单工作簿会保留 `报表清单`；Python 和 Rust 计划均删除 `使用说明`、`转表设置` 两张配置页，并按启用/空表规则处理报表页。F3 Python GUI 提示文字提到“指标名称”，实际匹配键只有机构代码、指标代码、数据值、增减额；名称不参与匹配。

本轮复核原 Python `config-real/` 的 3.0–3.3 工作表、关键表头和 3.3 启用状态；43 条非空报表全部启用。详细字段见 [S3 配置结构核对](AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-配置工作簿的实际工作表与关键字段只读核对)。原 `embed_financial_template_in_config` 只供把外部模板嵌入 3.3 配置，当前引用在测试中，不是 GUI 功能或 F5 运行时输入。

Rust GUI 还显示 3.0–3.3 四个配置检查按钮，但 Web bridge 只注册 3.3 检查 API；3.0/3.1/3.2 检查请求返回 `NOT_IMPLEMENTED`。这是五项业务之外的辅助 UI 缺口，后续需单独接通或在 UI 明确不可用。

### S3 编号与范围纠偏（2026-10-27）

原 Python GUI 的正式 S3 编号就是本清单中的五项：F1 两期数据→比较文件、F2 比较文件→说明文件、F3 机构说明→说明文件、F4 本期数值核对、F5 比较文件→金融表单。此前把 Python F4/F5 分别映射为 Rust 目标 F2/F3，并将 Python F2/F3 描述为额外流程，是错误的；现已在当前目标和需求文件中恢复五项原编号。Rust 已有的比较、说明导出、机构反馈计划、本期核对和指标拆分计划均保留，映射为 Python 同编号功能；其中 F3/F5 尚未闭环。

早期 Rust 手测 3.3 占位配置中表单停用，不代表原 Python 当前真实配置不可生成。只读检查原 Python `config-real/3.3大集中指标比较拆分_配置.xlsx` 显示 43 条非空报表清单均启用。此前 `s3_indicator_split.rs` 把指标拆分误标为 `S3-F03`；后续已将模块注释和诊断标签改为原 Python 编号 `S3-F05`，机构反馈导入为 S3-F03。

### S3-F05 执行链路进展（2026-10-28）

`s3_indicator_split` 新增 Excel-Core 编辑计划转换、逐机构独立 XLSX 执行、暂存目录及全批成功后目录发布；Rust + HTML GUI 已新增 3.3 XLSX 选择、金融表单输出目录、执行 API 和最近输出路径显示。执行入口通过真实 `ExcelModifier` 能力门控，当前桌面运行层没有注册覆盖 S3-F05 完整操作批次的 Passed 后端，因此实际调用会明确失败且不发布输出。没有 F5 CLI 或真实输出验收。测试源码已补但本轮未运行。

同轮补入 S3-F03 写回执行器及 GUI/Web API 多文件入口：沿用现有匹配/冲突规则，把非冲突更新应用到目标说明文件的独立输出副本，校验源文件哈希并保留冲突报告。当前桌面运行层没有 Passed S3-F03 编辑路由；有待写入变更时会失败且不发布副本，原始说明和反馈文件保持不变。现有 F3 CLI 仍只输出只读计划。新增测试源码未运行。

## 基线与状态口径

- 只读源项目：C:\Users\AZI\Desktop\VibeCode\3.open-data-audit
- 当前功能目录依据源项目 core/tests/TEST_GROUPS.md 与源项目 AGENTS.md 建立；本清单不是对全部实现细节的最终代码审计。
- 来源文档列出的 S1/S2/S3/S5 活跃业务项作为本轮清单；S4 在源项目中标为预览流程，是否纳入正式迁移仍待确认。
- S5-F02/F03/F04 在来源测试文档中有隐藏维护工具测试，但不在主要功能汇总表内；需继续沿源码确认其业务行为，不能仅凭隐藏状态视为废弃。
- workspace fixtures 只包含合成或脱敏测试数据，没有复制真实业务工作簿；真实配置和 2026-07-31 S1 文件仅从只读源目录做过一次模板推荐 smoke 对拍，报告只保留文件集合哈希和汇总结果。该 smoke 不能替代真实模板全集验收。
- 代码盘点（2026-10-23）：`audit-core/src/`、`pytools-core/src/`、`crates/` 和 `sidecar/` 没有 Python 源文件或 Python 子进程入口；workspace 当前 50 个 `.py` 文件仅位于 `probe/`、`tests/compare/`，作为旧实现基线和测试辅助工具保留，不进入产品运行路径。workspace 尚无 `.go`/`go.mod` 生产实现；`core-go/` 目前只有说明文档。
- 以下保留当时使用的“未验证、迁移中、受阻”等历史状态与局部证据；当前任务状态不沿用这些旧标签。代码存在、合成测试或过去的编译成功均不等同于当前真实业务验收。
- 独立项目分别记录实现、单元/冒烟、Windows 用户人工验收与 Linux/deb 目标验收；未执行的验证不标为通过，也不要求以全面对拍为人工交付前提。

## 历史功能清单（历史状态；S3 编号沿用 Python 来源标识）

| 来源标识 | 功能 | 需要覆盖的业务结果 | 当前状态 |
|---|---|---|---|
| S1-F01 | 汇总核查表校验 | 输入发现与筛选、模板匹配、规则执行、结果回写、历史记录及重复运行一致性 | 受阻；模板推荐局部、条件格式 Issue 提取及批量 CLI 切片已接 Rust；真实只读样本 348 条 Issue 与 Python 逐字段一致，原中文文件名 Issue identity 348/348 一致；样本按规则 AppliesTo 默认范围，尚未验证模板命名范围。新增 name_config::load_s1_conditional_mappings 和 resolve_s1_conditional_area_plan：Rust reader 读 S1 映射并生成条件格式扫描/结构区域计划，定向 evaluator 测试和 Python 合成 named_ranges 对拍通过；已新增 Rust 目录审核入口 `conditional-audit-directory`，从目录发现 XLSX、推断数据期和机构身份后调用批量 Issue 审核；真实只读样本 smoke 为 1 文件、1 条规则、29 个评估格、2 条 Issue，Python 服务对拍字段差异 0。目录入口支持 --template-workbook 或 --template-dir 自动推荐（可用 --selected-file 限定推荐样本）与可选 --mapping-config（未指定时用源项目内置映射）；推荐先使用筛选后的源工作簿，模板候选从审核输入和期间推断中排除，歧义时停止自动审核并要求指定模板。conditional-audit、conditional-audit-batch 和 conditional-audit-directory 均按模板 Defined Names 生成条件扫描/结构区域计划。合成配置/模板集成测试通过；真实命名区域对拍见 probe/reports/s1_conditional_named_area_real.json：4 个条件区域、8 个结构区域、4 条规则、602 个评估格、2 条 Issue，Python/Rust 字段差异 0。目录 CLI 另接重复 --selected-file 文件筛选（相对路径按输入目录解析，越出发现/扫描范围时拒绝），期间仅从选中源文件推断。合成筛选和期间冲突回归、真实目录对拍见 probe/reports/s1_conditional_selected_directory_real.json：候选 2/选中 1，识别 2026-07，2 条 Issue、602 格，Python/Rust 差异 0。本轮将表结构预检接入单文件、批量和目录 CLI；真实配对样本的 8 个结构区域、1,075 个固定标签 Python/Rust 结果一致（1,075/1,075），同次条件审核 2 条 Issue 全字段差异 0，见 probe/reports/s1_structure_preflight_real.json。新 `s1-audit-directory` CLI 已把已保存公式结果、条件格式 Issue 和历史归类串入 Rust 入口；148 条真实 Issue 按目标 Linux 的 UOS/native DAG handler 逐字段对拍差异 0。该 handler 的条件格式 Issue 留空机构字段；Windows/general handler 会传入机构字段，但不作为本次目标路线基线。完整业务规则/config 流程、公式复制/重算、既有历史工作簿写回及重复运行一致性仍未迁移/验收；非空历史工作簿读取/归类已有局部实现；详见下方局部迁移表，S1-F01 整体保持未验证。 本轮新增可选 `s1-audit-directory --result-workbook` Rust 输出：在现有公式/条件格式 Issue 合流及历史归类后新建 13 列审核结果 XLSX，并校验单元格后写入；真实只读样本中 299 条 Issue、299 行和 299 个审核副本链接与 Python writer 对拍差异 0，表值/类型/样式/列宽/筛选/冻结/链接均一致，输入未改变。该 299 条样本的完整 CLI 耗时 15,899 ms、峰值 RSS 20,456 KiB，结果写入阶段 5,872 ms；本次使用模板目录自动推荐，所选模板与 Python 基线一致；资源数据仅代表此样本，不设通用容量承诺。S1-F01 整体仍未验证。 同批单文件真实输入同期间首跑/回放已对拍：两次当前 Issue、历史合并及历史结果工作簿 Python/Rust 差异均 0；回放两侧读取 299 行、合并仍 299 行，299 条“再次出现”与旧 Python 一致（13 列历史输出契约未保存数据期），见 probe/reports/s1_history_repeated_run_real.json。公式复制 direct Umya 候选：main 协调修复 ZIP 目录项误判后，7/7 真实输入均生成临时候选、目标 inventory 完整且输入未变。Python/Rust 逐坐标对拍完成：4 个复制范围共 5,487 格，值/类型、公式文本、批注位置/文本/作者和超链接差异均为 0；复制样式采用完整 OpenPyXL 序列化组件语义（样式名称、数字格式、字体/填充/边框/对齐/保护，不比较样式 ID），对照模板差异为 100（样本 1）和 802（样本 2–7）。未修改区每本对比 3,047 格，直接以原始目标为基准的值/类型差异为 207–223，公式文本/样式/批注/超链接差异为 0；七本样本未修改区没有公式，公式缓存可比单元格为 0，缓存比较不解释值/类型变化；工作簿元数据差异为 0。旧 Python 跨工作簿复制基线出现无效字体组件引用（样本 1 为 1,012，源模板/目标无此问题），已与 Rust 对拍结果分开报告。每本仍有 17 个 ZIP 部件名和 19 个共同部件字节差异；不据此推断具体对象损坏，报告保留类别计数。Umya 候选对象级不匹配、生产路由仍 Unverified，未执行公式重算；详见 probe/reports/s1_formula_copy_candidate_real_batch.json。 Openpyxl worker 真实候选进一步确认：完整 S1 操作批次在写入前因 worker 未支持 SetWorkbookCalculationProperties 被拒；仅执行四项 ImportRange 的诊断批次，7/7 又因共用模板含 752 个非空公式缓存而在写入前拒绝，目标模板输入均未改变，未形成可比较输出。生产路由仍 Unverified；需 main 协调共享 worker 的计算属性批处理与源公式缓存保留能力后再复验，详见 probe/reports/s1_formula_copy_candidate_openpyxl_real_batch.json。main 落地请求 #3/#4（becd3cd 载荷基线等价豁免、e73bc1a 校验器行操作映射，audit-dev 同步 d5350f4/1d7c22c）后，业务计划按裁决对目标簿全部表声明 `MergedNonAnchorPayloads`（与 7/7 基线等价证据一致的逐表声明），openpyxl worker 首次写穿完整 60 操作计划（apply 通过、输出创建、2 个空作者兼容注释）；读回校验暴露新缺陷——worker 输出 sheet7 CJ2 样式索引 s=24 越界（cellXfs 仅 24），vertical=center 语义丢失，而 Python 基线同输入 s=24 在界内（cellXfs=38）语义保留，worker 路线不基线等价；已登记共享请求 #5（openpyxl worker 样式表越界引用），见 `probe/reports/s1_formula_copy_openpyxl_worker_e2e_windows_native_20261025.json` 与 `shared-capability-requests/20261025_openpyxl_worker_style_table_oob.md`。 main 已将计算属性批处理和模板来源缓存策略同步至 audit-dev（共享源提交 a3cd5f0；本地同步提交 c4ad6f4），对应 worker 4 项单测与 Rust→worker→Rust validator 集成测试通过。更新后的完整 5 操作批次在同批 7 本真实源工作簿上均因目标公式缓存保护而写入前拒绝：输出 0、输入/模板 7/7 未变；首本通过 Rust Calamine S1 公式结果读取到 264 个非空结果格（并非 OOXML 全工作簿缓存总数）。另由测试对拍 helper 仅统计缓存坐标：七本目标缓存分别为 688、692、686、692、686、688、983 格，测试 helper 仅记录与四个 ImportRange 坐标矩形的聚合关系（范围外计数 0）；未保存工作表名/单元格地址，不能证明精确 target_sheet + 地址覆盖或实际序列化 payload。原报告记录的 WSL2 isolated worker 在 preflight 因 formula_cache_not_preserved 拒绝，7 本无输出、输入模板不变；worker/runtime 版本与请求字段未捕获。该计数只作聚合诊断，不支持能力通过，生产路由仍 Unverified。由于无 Rust 输出，对象级保留对拍未运行，能力保持 Unverified。详见 probe/reports/s1_formula_copy_candidate_openpyxl_target_cache_guard_real.json。 另以同批 7 本既有 Python 公式结果副本经 `open-data-audit s1-audit-directory` 逐本做 Python/Rust 对拍（不包含 Rust 公式复制、公式重算或历史工作簿）：公式 Issue 317、条件格式 Issue 5、合并 322；7/7 Issue 与结果工作簿对拍差异 0，输入 7/7 未变，逐本结果累计 322 行/超链接。完整 7 本目录 CLI 完成 322 Issue，Rust 输出读回 322 行/超链接；开发者日志记录 33 个事件及实际引擎/阶段耗时。同批既有 Python 公式结果副本另由 Windows 原生 x86_64-pc-windows-msvc CLI 复验：7/7 Python/Rust 当前 Issue 与结果工作簿对拍总差异 0，322 条 Issue 写出 322 行/链接，7/7 输入未改变；完整原生目录 CLI 完成 322 条 Issue，33 个开发者日志事件，结果工作簿计数读回一致。该结果仍不覆盖 Rust 公式复制、公式重算或历史工作簿；另以同批首跑 Rust 专用历史输出簿作历史输入、按同期间回放 7 本：Python/Rust 每本均读取 322 条历史行，合并历史各 322 条，当前 322 条 Issue 全部归为“再次出现”；当前 Issue、合并历史及两类结果簿逐项差异 0，历史输出读回 322 行，7/7 比较输入未变。历史输入为本批 Rust 首跑生成的临时簿，不代表既有用户历史档案格式验收；详见 probe/reports/s1_audit_directory_history_replay_real_batch.json。S1-F01 总体仍未验证；详见 probe/reports/s1_audit_directory_existing_formula_copies_real_batch.json。 S1 个人贷款真实配置批次另完成 7 组 Windows 原生 Python/Rust 对拍（模板 895 条规则、27 条公式 Issue、逐本差异 0；使用已保存审核副本，未执行 Rust 公式复制/重算），汇总见 probe/reports/s1_personal_loan_real_config_batch_windows_native_20261024.json。  个人贷款公式复制候选复验：当前 Python 基线在合并单元格写入失败；openpyxl worker 写入前拒绝，Umya 候选相对既有副本/原输入分别发现复制范围 1,831、范围外 386 项差异，均未验收，见 probe/reports/s1_personal_loan_formula_copy_candidates_windows_20261024.json。  Windows 原生 openpyxl-worker 公式复制候选定向复验 2/7 均因 `target_merged_non_anchor_payload` 写前拒绝；每份目标工作簿 855 个非锚点载荷单元格全部位于 4 个复制区之外，涉及 31 个不相交载荷合并区；输入未变、无输出，能力仍 Unverified，见 `probe/reports/s1_formula_copy_target_merged_payload_guard_20261025.json`。 Windows 原名全批复验：7/7 原名工作簿逐本 Python/Rust 对拍差异 0，自动模板选择 7/7 与 Python 基线一致；Rust 目录 CLI 接受 7 本，生成 27 条当前 Issue、27 行结果及 27 行独立历史结果，逐本结果簿/历史结果簿差异 0。开发者日志含 38 个阶段事件及引擎/耗时；模板预检 895 条规则、4 个公式范围、8 个结构范围、4 个提取范围。源、副本、模板与配置未变，未使用文件名别名；公式复制/重算、既有历史簿输入及写回、Linux/deb 未验证，S1-F01 仍未验证，详见 `probe/reports/s1_personal_loan_original_filename_full_batch_windows_native_20261025.json`。main 的 DefinedNames scope 修复 6924e4b 安全同步为 eb50796 并推送；4 个 scope 单测及 1 个 Print_Area 测试通过。真实历史簿候选在 Windows 隔离 worker 写后进入 Rust DefinedNames 校验，仍报告 1 条 missing、2 条 added/changed；候选输出由测试清理，源与临时输入不变。main 批准的一次 metadata-only 诊断仅写临时文件、未进入仓库或聊天；S1 历史簿能力和 S1-F01 整体保持未验证，见 `probe/reports/s1_history_defined_names_scope_retry_windows_native_20261025.json`。 本轮新增 Rust `s1_validation_formula_scan::S1ValidationFormulaScanPlan` 并接入 `s1-audit-directory`：按模板 `校验区域*` 定义名称扫描缓存公式错误，并合入公式 Issue/后续历史归类。同批 7 本真实只读源文件和 1 个模板逐字段对拍差异 0；累计扫描 7,504 个公式单元格，Python/Rust 均发现 1 条源缓存错误、扫描失败 0、缺失工作表 0、新增 Issue 1 条。原始路径、文件名、公式及业务值未写入报告；详见 `probe/reports/s1_validation_formula_scan_real_parity_windows_native_20261025.json`。合成 CLI 回归覆盖范围内错误纳入、范围外错误忽略及公式字符串错误文本排除。 同一真实 7 文件批次另以 `s1-audit-directory` 同时生成 Rust 本期结果及新建历史结果 XLSX：两表各 323 条，扫描 Issue 进入当前 Issue 并保留在合并历史中，两个写入阶段日志均完成，读写前后源工作簿、审核副本、manifest 和模板大小/修改时间不变；临时输出已删除。Rust writer 路径由 CLI 成功和其输出校验覆盖，合成端到端测试使用 Calamine 读回两表并核对扫描 Issue ID/类型/值。  同一真实个人贷款批次再选取 1 个原始源/审核副本/显式非内置映射/模板组合，Windows Rust CLI 与 Python 基线对拍通过：模板规则发现差异 0；Python 2 条旧公式 Issue + 1 条配置校验公式扫描 Issue，Rust 同为 2+1，扫描 895 个校验公式单元格，当前及合并历史 3 条 Issue 差异均为 0。扫描详情中的本地模板路径按 basename 归一，其余 Issue 字段及公式文本精确比较；源、审核副本、模板和配置工作簿均未改变。详见 `probe/reports/s1_audit_directory_explicit_mapping_personal_loan_validation_scan_windows_native_20261025.json` 和 `probe/reports/s1_template_rule_discovery_explicit_mapping_personal_loan_windows_native_20261025.json`。另对用户提供目录中的 184 个 XLSX 使用原 Python 历史读取器只读扫描，未发现非空历史记录；其中含既有管理页的空簿由 Rust Calamine 与 Python 对拍为 0 行、字段差异 0。该结果只覆盖空历史输入，非空既有历史簿及写回仍未验证，见 `probe/reports/s1_history_existing_workbook_read_parity_windows_native_20261025.json`。此外，用户指定的原项目 `config-real` 目录含 29 个 XLSX；Python 历史读取器只读扫描发现 1 份非空历史簿（9 行），Rust Calamine/Python 实际读取字段差异 0、输入未变。该簿共 5 张表，管理历史页位于 index 1；当前计划按原 Python 语义需将其置首，但共享 MoveSheet 尚不可用，故完整更新仍写前拒绝/Unverified；该非空样本只验收读取，不代表写回，见 `probe/reports/s1_history_nonempty_config_real_read_parity_windows_native_20261025.json`。另以同一 9 行真实簿运行 ignored Rust 候选预检测试（1 passed）：index 1 的管理页在调用 Editor 前按预期拒绝；未生成输出，临时输入与源文件均未变。无候选输出，故 PageSettings 和其他保留项未做写后对拍；能力仍 Unverified，见 `probe/reports/s1_history_nonempty_config_real_update_preflight_windows_native_20261025.json`。|
| S1-F02 | 汇总校验结果说明 | 汇总说明、未匹配/兜底行为、历史说明与输出结构 | 已实现，待 Windows 真实验收。Rust CLI `s1-merge-workbook-tables` 已迁移“汇总_汇总表合并”切片；`s1-region-summary` 接入“任意行汇总/固定行汇总”命名区域切片，并读配置映射、表头范围和合并标题，按来源工作表写新簿。Windows Excel/Python 与 Rust 合成对拍：2 个源簿、2 张输出表、9 行；基础值/类型、表序、尺寸、数字格式、表头样式、筛选与冻结差异为 0；全局/局部单格元数据列对拍通过，缺少元数据工作表时留空。新增可选 `--history-workbook`：接入本地校验结果、业务说明、任务说明三类内置历史键/字段规则；普通存储值 XLSX 由 Calamine 读取，复合键归一化并按顺序合并重复历史值，含公式历史表明确拒绝。Windows Excel/Python 与 Rust 合成文件级对拍：本地校验结果历史追加表值/类型、表序、尺寸、样式和筛选差异为 0；数值键 `1.0` 与历史文本键 `1` 匹配，重复历史行按源顺序拼接，未匹配行留空，输入未变。自动列宽差异分别最大 3.57/4.57 字符；新增任意行/固定行“兜底模板”未匹配源工作表汇总：2 个源簿、2 张输出表、4 行；Windows Microsoft Excel/Python 与 Rust 的值/类型、表序、尺寸、格式、表头样式、筛选和冻结差异均为 0；同名工作表跨文件合并、无有效行工作表跳过，列宽最大差 4.21 字符。消费区域的源簿与历史表现支持读取已保存公式结果，缺少缓存或缓存为 Excel 错误值的消费公式仍明确拒绝且不触发重算；真实 F02 模板及 Linux/deb 未验收。S1-F02 整体仍未验证。见 `probe/reports/s1_f02_workbook_table_merge_synthetic_parity_20261024.json`、`probe/reports/s1_f02_named_region_summary_synthetic_parity_20261024.json`、`probe/reports/s1_f02_named_region_history_synthetic_parity_20261024.json`。  用户提供的 2026-07-31 个人贷款只读真实样本内置映射对拍：7 个源簿、27 张源表，输出 4 张表/44 行；Python Microsoft Excel 与 Rust CLI 的表序、尺寸、值/类型、数字格式、表头强调、筛选/冻结、公式/合并/定义名/页面设置/打印区域/超链接差异均为 0，输入未变。字体和对齐在 588 个非空单元格位置不同，54 列宽差异最大 12.92578125 字符；该切片仅数据/结构对拍通过，完整格式契约未验收，S1-F02 整体仍未验证。见 `probe/reports/s1_f02_named_region_real_personal_batch_windows_native_20261024.json`。 Windows 原生真实批次格式复核（2026-10-25）：7 个工作簿、27 张源表、4 张输出表、44 行；值/尺寸/字体/对齐/表头/筛选/冻结等本次比较项差异为 0，输入及模板 SHA-256 未变；列宽仍非完全一致，最大差 12.92578125 字符，因此 S1-F02 保持未验证。2026-10-25 按原 Python Excel COM 的 AutoFit 后 40 字符上限修正 Rust 列宽上限；Windows 合成 Excel/Python 对拍的值、类型、样式、表序、筛选、冻结等比较项通过，列宽仍有估算差异（主结果最大 6.859375 字符、历史结果最大 8.7890625 字符），因此仅记录局部修正，不提升功能状态。 2026-10-25 后续原生 Windows Rust+Excel/Python 合成对拍修正 AutoFit 参数单位：rust_xlsxwriter 的最大宽度参数为像素，40px 曾把多列限制在约 5.71 字符；现改为 285px（约 40 个 Calibri 11 字符单位）。合成主结果、历史、兜底、固定行列宽最大差分别为 3.5703125、4.5703125、3.7890625、3.7890625 字符；值/类型、数字格式、字体和对齐差异均为 0。列宽仍不完全一致，S1-F02 保持未验证；见 `probe/reports/s1_f02_named_region_summary_synthetic_parity_windows_native_20261025_autofit_cap_pixels.json`。 Windows 原生 CJK 列宽修正对拍（2026-10-25）：同批 7 个源簿/27 张源表，Python Microsoft Excel 与 Rust 输出列宽最大差由 12.92578125 降至 3.07421875 字符；值、尺寸、字体、对齐、表头、筛选、冻结差异均为 0，配置/模板/输入未变。列宽仍不精确，S1-F02 整体保持未验证；见 probe/reports/s1_f02_cjk_column_width_real_parity_windows_native_20261025.json。 按用户 2026-10-27 内容优先验收标准，字体、对齐和列宽等普通外观差异不再阻塞；既有真实样本的数据值/类型、关键结构、公式、合并、定义名等内容项差异均为 0，当前状态为“内容对拍通过，待人工验收”。本次未重跑样式或列宽对拍。 |
| S1-F03 | 组合工作表（核查表） | 多来源工作表合并、命名、顺序、公式和命名区域保留 | 未验证；待真实组合文件 |
| S1-F04 | 组合工作表（按配置） | 按配置/DAG 执行组合、输入验证和失败策略 | 未验证；待配置契约和黄金输出 |
| S1-F05 | 制作联合模板 | 多模板合并、基准模板选择、名称冲突处理和引用修正 | 未验证；待模板组合样本 |
| S2-F01 | 执行报表规则 | 报表导入、环比、外部核对、规则校验、三类结果导出 | 已实现，待 Windows 真实验收；Rust service/CLI 对拍通过，loopback Web bridge 可由原样 HTML 调用并完成 synthetic 端到端运行，目录选择已接入两期文件配对预览；真实期间模板/报表、外部核对、完整输入/导出错误矩阵、目标平台和 deb 仍未验收；配置校验子流程已对真实 config-real 配置由 Windows 原生 loopback Web API 与 Python 逐字段对拍：1 份配置，105 个指标、56 条环比策略，阻断/警告均 0、报告字段差异 0；双视图日志记录 Calamine reader 与耗时，见 probe/reports/s2_config_check_real_config_windows_native_20261025.json；Windows 原生真实批次端到端对拍：本期/上期各 21 个 XLS，使用真实 S2 配置与外部 XLSX；7 家机构、3 张表单、728 项指标、1095 条结果，Rust CLI 与 Python 服务摘要一致；输出三表 17520 格及扩展语义保真对比差异 0；44 个输入/配置 SHA-256 前后差异 0，受控验证合计 9185 ms（含哈希与对拍，不是性能基准），见 probe/reports/s2_real_service_parity_windows_native_20261026.json；S2 进度日志已补强：CLI 事件记录引擎、累计耗时和当前阶段耗时，Web 进度标签显示当前引擎及阶段已运行时间；合成成功/失败三路径通过，Windows 原生真实批次 69 个事件字段齐全，报告簿对拍 17520 格差异 0、44 个输入 SHA-256 不变，见 probe/reports/s2_progress_logging_windows_native_20261026.json；Windows 原生 loopback HTTP 真实批次日志复核：本期/上期各 21 个 XLS，7 家机构、3 张表单、728 项指标、1095 条结果，全部摘要计数与既有 Python/Rust CLI 报告一致；11 个阶段均记录开始/完成、匹配引擎及阶段耗时，24 条开发者详情日志，总计 2249 ms（单次观察，不是性能基准）；输出生成并经 Calamine 读回后删除。此 Web 输出本轮未再次做全工作簿 Python/Rust 对拍；17,520 格零差异和 44 个输入未变仍由既有 CLI 报告证明，详见 probe/reports/s2_web_stage_logging_real_windows_native_20261026.json；S2-F01 整体已实现，待 Windows 真实验收 |
| S3-F01 | 两期数据到比较文件 | 两期数据匹配、规则命中、比较计算、说明及结果输出 | 已实现，待 Windows 真实验收：Rust 读取 3.0/3.1 配置，完成八维数据配对、规则动作、五段/八段表达式及旧版 22 列结果写入；CLI `s3-execution-compare` 和原 GUI `start_central_comparison` API 已接通，GUI 合成 API 已验证后台执行、日志和输入保护。真实 2026.07/2026.08 两期 CSV Python/Rust 对拍通过：本期 2,595、上期 2,586、结果 2,610 行；两种表达式模式均为键维度和非数值/结构差异 0；精确数值对拍有 1,306 个 IEEE-754 尾差（最大绝对差 5.7e-14），按 1e-9 容差、10 位及 2 位小数显示均为 0 差异；归一化样式、列宽及工作表设置差异 0，4 条表达式触发、执行错误 0。证据 `probe/reports/s3_f01_real_parity_windows_native_20261027.json`。原 GUI 人工操作验收及 Linux/deb 尚未完成。 |
| S3-F02 | 比较文件到说明文件 | 将比较结果转换为说明文件并保持输出契约 | 已实现，待 Windows 真实验收；Rust CLI `s3-export-explanations` 使用 Calamine 读取、`rust_xlsxwriter` 新建；Python/Rust 合成 4 行及真实比较工作簿 2610 行对拍差异 0，核对全部输出单元格/公式文本/有效样式/合并/名称/尺寸/页面设置/元数据，两份输入均未变。真实报告 `probe/reports/s3_f02_real_comparison_parity_windows_native_20261026.json`；Rust 输出已按 Python 默认页边距校准。原样 Web UI 的 `start_central_explanation_export` 已接入 Rust bridge；`choose_central` 支持比较文件、3.0 通用配置及输出目录。Web 导出沿用 Python 的 `centralCommonConfig` 机构映射，配置不可用时明确失败；合成 Web API 测试核验映射结果与缺配置错误。Windows loopback HTTP 真实样本烟测报告 2610 条并记录 Calamine/`rust_xlsxwriter` 双视图日志。该烟测未单独对 Web 输出运行 Python/Rust 全工作簿对拍；既有完整对拍证据仍为上述 CLI 报告。XLSM、峰值内存预算、目标 Linux/deb 验收尚未完成，S3-F02 整体待 Windows 真实验收 |
| S3-F03 | 机构说明到说明文件 | 机构反馈读取、身份键匹配、冲突处理和说明回写 | 部分实现、整体未验证：Rust CLI `s3-plan-explanation-feedback-import` 保持只读计划和冲突清单；新执行器将非冲突计划写入独立副本并在发布前复核输入哈希。GUI 多文件选择和 Web API 已接通，但运行时没有 Passed S3-F03 Excel-Core 能力，需写值的任务会失败且不发布；真实多机构反馈、冲突报告和成功写回仍未验收。 |
| S3-F04 | 本期数值核对 | 本期数据核对规则、跨频度/批次关联和异常报告 | 已实现，待 Windows 真实验收：Rust 规则、`s3-check-current-values` CLI 和 `start_central_cross_period` Web API 已接通；输入支持 CSV/XLSX/XLS。Python GUI 仅提供本期文件，Rust Web 当前同样走本期输入。CSV 真实 Python/Rust 对拍见 probe/reports/s3_f04_python_rust_real_parity_windows_native_20261026_after_safety_v3.json，覆盖本期 4,212 条、上期 2,586 条、183 条规则并生成 214 行，结果/日志差异 0、输入未变。新增真实只读 XLSX 读取对拍覆盖 2,610 条记录、183 条规则（163 条启用）；该本期-only 样本没有结果行，工作簿/日志差异 0、输入未变；临时合成 XLS 对拍 1 条记录，差异 0。此前记录的测试仅为历史证据；完整错误矩阵、真实 XLS、能产生结果行的本期 XLSX/XLS 样本、GUI 人工操作及 Linux/deb 未验证。 |
| S3-F05 | 比较文件到金融表单 | 比较结果按 3.3 报表清单、指标及频度/批次映射为逐机构表单 | 部分实现，需要继续开发：Rust `s3_indicator_split` 已规划并转换为 Excel-Core 编辑计划，新增逐机构输出、隔离暂存和全批发布；Web API/GUI 输入、配置、输出目录与执行按钮已接通。桌面运行层没有 S3-F05 完整操作批次的 Passed Excel-Core 后端，执行被能力门控拒绝，不生成结果；无 CLI、真实输出和人工验收。真实 Python 3.3 配置有 43 条启用报表。后续须补足经验证的生产后端能力并做真实表单验收。 |
| S5-F01 | 汇总核查表校验（逐笔统计配置版） | 配置驱动读取、逐笔规则、校验结果与输出 | 已实现，待 Windows 真实验收；已完成报告期识别子切片：Rust CLI `s5-detect-period` 按源文件名、源表报告期字段、输入目录祖先顺序识别，支持原格式、月末归一、非法/多期拒绝和公式单元格忽略；20 例 Python `periods_in` 基线逐项一致，合成日期单元格 Python/Rust CLI 的报告期与来源一致；用户提供的两组选中文件各 7 份，Python/Rust 均识别 `2026-07-31` 且来源均为文件名，输入只读。新增 Rust CLI `s5-file-period-hints` 迁移 Python `file_period_hint`：文件名/目录优先、异常/多期提示、不读取工作簿；9 例 Python 基线一致，用户 S5 样本子树 120 个 XLSX/XLSM 路径与 Python 提示逐项一致（0 差异、输入只读），见 `probe/reports/s5_file_period_hint_python_baseline.json` 和 `probe/reports/s5_file_period_hint_windows_native_20261025.json`。普通/开发者日志含步骤、实际路径解析器/reader 状态与耗时。此前期间识别证据见 `probe/reports/s5_period_parser_python_baseline.json` 和 `probe/reports/s5_period_recognition_windows_native_20261025.json`。新增 Rust `s5-match-template` CLI，迁移 `recognition.match_s5_template` 的工作表集合评分、覆盖兜底、0.05 歧义阈值和同业务批次约束；用户真实 120 个 XLSX 逐文件 Python/Rust 对拍为 85 匹配、35 不匹配、0 差异，同模板/跨模板批次结果一致，见 `probe/reports/s5_template_matching_windows_real_20261025.json`。配置 DSL 解析与静态校验已移植为 `s5_config_dsl` 模块，真实配置 2,392 条启用规则 Python/Rust 逐字段对拍 0 差异，见 `probe/reports/s5_config_dsl_parity_windows_native_20261025.json`。配置工作簿加载层已移植为 `s5_config::load_configuration`，真实配置整簿加载 Python/Rust 逐字段对拍 0 差异（2,392 条规则全字段、4,215 个派生模板指标、版本字段与 sha256），见 `probe/reports/s5_config_load_parity_windows_native_20261025.json`。业务识别与逐文件分类已移植为 `s5_business::{match_s5_business, classify_s5_files}`，真实 120 文件样本子树逐行全字段对拍 0 差异（87 matched/14 partial/19 unmatched，18 个业务识别），见 `probe/reports/s5_classify_business_windows_real_20261025.json`。配置校验入口已移植为 `s5_config_check::validate_s5_config`（scope 门控、模板业务推断、加载错误映射及规则/引用计数）；当前业务范围模板/外部地址检查接入 `open-data-audit s5-config-check` CLI。V_EXT 指标头唯一性与 `sheet=` 过滤的 2 个合成 Python/Rust 用例全字段对拍 0 差异，见 `probe/reports/s5_config_check_external_addresses_synthetic_parity_windows_native_20261025.json`；all/all_enabled 全量模板目录扫描与外部地址检查已接入 CLI；config-real 真实目录 6 场景 Python/Rust 对拍 0 差异（错误集合按内容规范排序，避免 Python set 的非稳定顺序影响比较），见 `probe/reports/s5_config_check_fullscope_real_windows_native_20261026.json`；仍有 64 个未定位引用且两侧一致，不代表 S5-F01 整体通过。真实配置 3 个 scope 用例此前全字段对拍 0 差异，见 `probe/reports/s5_config_check_windows_real_20261025.json`。规则执行运行时已移植为 `s5_eval::evaluate`（双数值模式：excel_binary64 经 excel_round15 收敛；decimal 经内部 28 位有效数字上下文的数字向量十进制，算术银行家舍入、ROUND 半进位），79 个对拍用例（22 个双模式合成用例 + 真实配置逐业务抽样 57 条启用规则表达式，字典变量与 V 调用以确定性 number-1 桩求值）Python/Rust 逐字段 0 差异，见 `probe/reports/s5_eval_runtime_windows_real_20261025.json`。布局区域发现已移植为 `s5_layout`（NamedRegion/region_number/named_regions/numeric_regions/inferred_numeric_regions/paired_column_region/结构区域分组），真实单位贷款与个人贷款模板双例全字段对拍 0 差异，分别命中条件格式区域兼容分支与表结构区域交叉推导分支，见 `probe/reports/s5_layout_regions_windows_real_20261025.json`。地址扫描已移植为 `s5_scan`（锚点簿/_sequence/row_path_candidates 组限定别名/_contextual_row_hits/_starts/列投影/in_numeric_regions/scan_rule_paths 主循环，未绑定源按整表回退），并接通 `validate_s5_config` 的模板地址检查：真实配置+真实单位贷款模板端到端 4 用例全字段对拍 0 差异（含业务推断、扫描定位、错误结构与消息），见 `probe/reports/s5_config_check_windows_real_20261025.json`。V_CUR/V_PRE 取数核心已移植为 `s5_retrieval`（单绑定无契约路径：绑定期间/机构门控、指标目录查找、`_template_index_from_paths` 直索引、公式/错误/空值门、金额/比例/计数单位换算与完整 Evidence 记录），真实 config-real 配置 + 真实单位贷款审核副本 8 个 V_CUR 调用对拍 0 差异，见 `probe/reports/s5_retrieve_windows_real_20261025.json`。审核服务规则执行已移植为 `s5_service::evaluate_rule`（`_evaluate_rule` 语义：变量缓存/检索缓存与追踪、适用表达式门、审核/当前/对比/差值四表达式与类型检查、`_render_description`/`_display_value`、evidence 汇总去重、错误分类按取数码直取（STRUCTURE/CONFIG 集合判定状态、RETRIEVAL_FAILED 其余、`_error_code` 前缀仅用于无码计算错误）），真实配置+副本 6 条启用规则结果全字段对拍 0 差异，见 `probe/reports/s5_rule_evaluation_windows_real_20261025.json`。run 编排已移植为 `s5_service::run_rules`（规则按 rule_id 分组取最大修订、规则/配置生效期门（NOT_APPLICABLE 消息逐字）、适用范围展开为指标行路径（空 → CONFIG_ERROR“适用范围没有匹配到任何模板行路径”）、`@当前行` 以 scope 元组替换进 V 调用、结果行含规则元数据与 scope_path、状态计数 ledger；V_PRE/V_EXT 未绑定时报 Python 同款 SOURCE_NOT_SELECTED），真实配置+副本整业务 377 条结果行 + 状态计数全量对拍 0 差异，见 `probe/reports/s5_run_orchestration_windows_real_20261025.json`。结果输出已移植为 `s5_outputs::write_run_outputs`（单业务形态：审核副本+配置复制与 sha256 清单、审核结果.xlsx（触发+失败行、严重度错误分类、证据定位与副本超链接、precision 数字格式、`_style_table` 视觉语义：微软雅黑/表头深蓝/行高 54/冻结/筛选/横向 A3 fitToPage/标题行重复/条件格式触发空表）、运行总账.json 稳定字段），文件级对拍 0 差异：副本字节一致、工作簿值/类型/数字格式/超链接/对齐/列宽/冻结/筛选/行高/表头样式/页面设置与总账稳定字段、清单逐项一致（颜色按 6 位 RGB 归一、fitToWidth=1 与属性省略等价、分组列宽按 min/max 展开——均为 rust_xlsxwriter/openpyxl 序列化表示差），见 `probe/reports/s5_run_outputs_windows_real_20261025.json`。求值侧同步修正两处保真缺陷：`with_numeric_mode` 按规则模式收敛 Decimal→round15 f64 且 blank 保持空值种类；取数失败 evidence 补 Python `retrieve` 的 function/sheet setdefault 富化并直取错误码。产品 CLI 入口已接通：`open-data-audit s5-audit-run`（--config/--source/--period/--org-code/--org-name/--business×N/--output-dir/--run-id）串接配置加载→单绑定取数→run_rules 编排→write_run_outputs，JSON 汇总含规则数/状态计数/输出文件与触发行数；真实 config-real 配置 + 真实单位贷款审核副本端到端运行 377 条规则结果（状态计数与对拍一致），CLI 产物与 Python 基线文件级对拍 0 差异（审核副本字节一致、审核结果.xlsx 全字段、总账状态计数/运行状态/清单）。当前共享 reader 只接受 XLSX，XLSM 支持待 main 协调；单来源 V_EXT 配置契约与无契约普通静态布局识别已有合成 CLI 对拍；无契约静态布局告警、空白/非数值按 0 参与审核和输出已接通，空白/非数值/整数金额 3 例 Windows 原生 Python/Rust CLI 对拍差异 0，见 `probe/reports/s5_vext_dynamic_layout_quality_synthetic_parity_windows_native_20261025.json`；V_EXT 时间序列和当前业务范围外部地址检查已接入并通过合成对拍（真实外部工作簿未验收）；全量模板目录扫描已接入 `s5-config-check --template-dir --scope all/all_enabled` 并通过 6 个合成 Python/Rust 用例全字段对拍，真实全量模板目录仍未验收；数字区域模板投影取数已实现并接入 `s5-audit-run`，新增 `--template-path`/`--template-dir`，按 Python 规则从模板完整合并表头唯一投影并限定数值取数范围；合成 CLI 覆盖唯一匹配、无模板兼容回退、歧义拒绝、显式模板优先和默认目录，Python/Rust 目标定位一致；1 组真实只读源/模板样本的 4 个区域坐标逐项一致且输入未变，见 `probe/reports/s5_numeric_scope_template_projection_real_windows_native_20261026.json`。真实样本仅验证投影坐标，未执行完整规则审核与输出对拍；后续以用户提供的同批当前源/模板运行 Rust run_rules 与 Python AuditEngine.run，377/377 结果行和状态计数一致（Windows 原生 Python 3.12.10），证据见 probe/reports/s5_numeric_scope_real_audit_parity_windows_native_20261026.json；该服务对拍不覆盖 CLI 文件输出、结构快照、V_PRE/V_EXT；S5 CLI 输出层已增加输入/配置哈希变化降级：本期、上期、外部、布局模板或配置的快照不匹配时，将 PASSED/TRIGGERED 结论改为 CALCULATION_FAILED、清空触发值并在结果/总账记录 INPUT_CHANGED_DURING_RUN、完整性标记和前后哈希；Windows 原生 5 项 Rust 合成回归与 Python _refresh_integrity 临时文件基线 4 种情况语义一致。未注入真实并发变化，也未执行 Python/Rust 文件产物自动差异；结构快照等仍未验收，详见 probe/reports/s5_input_integrity_downgrade_synthetic_parity_windows_native_20261026.json；S5-F01 已接通 Windows GUI；合成正常/异常各 1 例通过，内容正确性及条件格式触发已验证，真实样本与 Python/Rust 逐项业务对拍仍未验证，等待用户人工验收 |
| S5-F02 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |
| S5-F03 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |
| S5-F04 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |

### S3-F04 XLSX/XLS 输入扩展（Windows，2026-10-26）

Rust load_central_workbook_dataset 通过共享 Calamine 读取 XLSX/XLS，按 Python 约定选首个 14 列表头匹配的工作表；稀疏单元格按物理行号生成问题行。CLI 历史参数 --current-csv/--previous-csv 兼容保留，但可接收 CSV、XLSX 或 XLS。原页面的 S3-F04 选择器和操作指南明确列出 CSV/XLSX/XLS，文件过滤使用 central_data；S3-F04 API 在后台线程调用同一 Rust service，输出路径先保护，日志记录实际 reader/writer 和耗时。

Windows Python/Rust 文件级对拍：只读真实 XLSX 输入 2,610 条记录，加载真实 3.0/3.2 配置共 183 条规则（163 条启用）；无上期输入，因此两侧结果行均为 0。结果工作簿、运行日志差异均为 0，输入/配置未变，Rust reader=calamine、writer=rust_xlsxwriter。临时合成 XLS 输入 1 条记录，Python/Rust 结果及日志差异为 0；xlrd 确认日期、TRUE、FALSE 使用对应 XLS 单元格类型。报告不含原始路径或单元格业务值：probe/reports/s3_f04_workbook_input_python_rust_parity_windows_native_20261026.json。

覆盖边界：真实 XLSX 当前期读取和输出契约已对拍，但本样本没有非空规则结果；XLS 只有合成对拍，尚无匹配真实 XLS 样本。XLS/XLSX 时长单元格仍显式拒绝（共享 reader 未提供足以还原 Python 时间/时长类型的格式信息）。Web API 端到端单测已运行，真实 GUI 点击尚未验证。该输入扩展不提升 S3-F04 整体状态，也不替代已有 214 行 CSV 两期结果对拍。

## 格式与共用能力矩阵

| 输入/对象 | 当前处理计划 | 验收状态 |
|---|---|---|
| xlsx | calamine 读取；OOXML 补充缺失元数据；新建写入用 rust_xlsxwriter；原位编辑候选 umya | 合成探针已有报告；业务格式与真实模板未验证 |
| xlsm | 逐项验证公式和宏相关部件；生产不执行 VBA；不得仅凭 OOXML 容器判断保真 | 未验证 |
| xls | calamine 读取能力逐项验证；需改 sheet 名等不支持操作时允许显式转换为 xlsx | 未验证 |
| xlsb | 逐项验证 calamine 值读取；其余元数据缺口按适配器能力处理 | 未验证 |
| csv | Rust `csv` + `encoding_rs` 已迁移 S3 固定 14 列导数的 CSV 解析/标准化；保留 GBK 优先、UTF-8 BOM 回退、引号字段、空记录行号、日期/数值及异常值语义 | 3 个合成 Python 对拍场景通过；仅覆盖 S3 CSV 路径，通用 CSV 契约及 XLS/XLSX 导入仍未验证 |
| et | 经用户显式选择后由 sidecar 转成 xlsx；验证 sheet 清单和数据 | 未验证 |
| 批注/条件格式/数据验证 | calamine 未提供时由 OOXML 或 umya 补；元数据缺失必须报告未验证 | 合成样例仅覆盖已有探针列出的对象 |
| 公式重算 | 默认不运行；用户显式要求本轮结果时调 sidecar 并检查缓存刷新 | sidecar 专项尚未完成 |
| 打印 | 仅迁移页面设置写入；PDF、渲染不在范围 | 页面属性逐项待验收 |

## 历史实现前门槛（已由独立项目工作流取代）

1. 对每个功能找到源代码入口、直接调用链、配置字段、输入/输出和错误行为。
2. 将需要保留的 workbook parts 和公式口径登记到独立 probe config。
3. 取得真实脱敏模板；无样本时可先做合成探针，但保持真实模板状态为未验证。
4. 旧工作区曾要求先完成共享 reader/writer/editor 对拍；现按根 `AGENTS.md` 的授权和用户人工验收工作流执行，不以全面对拍作为交付前提。
5. 每次状态变化在 audit-core/GPT修改记录/ 增加日期报告；不以单测或合成文件推断真实模板通过。

## 已对拍的局部迁移（不代表业务流程迁移）

| 来源 | Rust 实现 | 状态 | 证据与范围 |
|---|---|---|---|
| `base_audit.native.history_io.write_history_xlsx` | `audit_core::s1_history::build_s1_history_workbook_edit_plan` | 未验证（Audit 业务 EditPlan 构造已接通；没有生产 editor capability 或 CLI 写回入口） | 既有页改为 DeleteRows+InsertRows 原位替换，保留工作表身份及用户确认的 PageSettings；非首位页现以共享 `MoveSheet{new_index=0}` 置首（main 提交 bbd922d 经 cherry-pick 同步为 db5f7e1）。源 Python 会删除并在 index 0 重建历史页；Rust reader 只读核查用户 2026-07-31 样本目录中的工作簿，10 张表，受管页位于 index 6。Rust 计划现读取既有本地打印区域，并在行替换后追加 `SetPageSettings` 恢复原范围。Windows 原生真实 config-real 非空历史簿（5 表、管理页 index 1、9 行）候选（2026-10-25 晚间复验）：请求 #4（1d7c22c）落地后重跑，**删除区样式误报消失**（旧容忍分支已移除）；唯一残差为共享校验器 page_settings Vec 的顺序敏感比较——MoveSheet 合法重排后逐表 pageSetup 属性字典经 Rust/Python 双侧证明完全一致、仅顺序不同，测试以精确断言（sheet 数+逐表内容一致）后以 `passed_with_page_settings_order_gap` 继续，已登记共享请求 #6 `shared-capability-requests/20261025_validator_page_settings_order.md`。晚间复验前状态：main 已响应两项共享请求（`07196d5` 校验器 MoveSheet 感知、`7ac87d6` umya SetFreezePanes/SetAutoFilter，audit-dev 同步为 `69dc834`/`2f1a9ac`），完整 38 操作计划（MoveSheet+DeleteRows+InsertRows+SetCellValue×22+SetFreezePanes+SetAutoFilter+SetColumnWidth×13+SetPageSettings）umya apply 全链通过，引擎无关校验器除删除区内带样式空格的样式读回误报（M6/L3，已登记共享请求 #4 行操作坐标映射）外通过；审计侧读回验收全过（管理页置首、新内容就位、其余 4 表逐格一致、输入未变），见 `probe/reports/s1_history_move_sheet_real_candidate_windows_native_20261025.json`。前次 MoveSheet-only 候选与合成聚焦候选证据保留于同报告历史版本。第三项请求（载荷基线等价）仍未响应。剩余：校验器行操作映射（请求 #4）、openpyxl worker DeleteRows/InsertRows 路线、生产 capability 登记与 CLI 写回入口；当前不改生产路由，S1-F01 历史工作簿写回仍未验收。 |
| `base_audit.name_config.load_feature_mappings` 驱动的显式 S1 映射审核链路 | `audit_core::name_config::{load_s1_formula_mappings, load_s1_conditional_mappings}` + `audit_core::s1_template_rules::load_s1_template_rule_plan`，由 `open-data-audit s1-audit-directory --mapping-config` 接入 | 保留（真实显式映射配置的审核切片；S1-F01 整体未验证） | `probe/reports/s1_audit_directory_explicit_mapping_real_20261024.json`：同批 1 个真实源文件与已保存 Python 审核副本，显式传入 Audit 映射工作簿；Python/Rust 均生成 2 条公式 Issue、0 条条件格式 Issue、合并 2 条，Issue 与结果工作簿总差异 0，结果 2 行/2 链接，模板匹配、所有输入未变。使用 Calamine、namespace-aware OOXML、Rust S1 Issue/history 编排和 rust_xlsxwriter。未执行 Rust 公式复制/重算，无历史工作簿输入；该证据只覆盖此配置/样本，不验收完整 S1-F01。 |
| `base_audit.native.openpyxl_workbook.extract_issues` 的普通公式结果过滤与 Issue 组装 | `audit_core::s1_formula_audit::extract_s1_formula_issues`，由 `open-data-audit s1-formula-audit` 接入 | 保留（公式结果读取/Issue 组装局部；S1-F01 仍未验证） | `probe/reports/s1_formula_issue_extraction_real.json`：1 个真实模板与缓存审核副本，1,072 条启用规则、911 个非空且非 falsey 公式结果、297 条 Issue；Python/Rust 全字段差异 0，输入前后 SHA-256 不变。另以 2 个 OOXML 错误单元格验证公式错误 Issue；合成结果共 3 条 Issue，Python/Rust 全字段差异 0。Calamine 读取已保存值，支持结果提取区域、`keyword`/`error_only` 模式、错误值和公式结果字段解析。未执行重算，也未迁移公式副本写入、历史工作簿写回或重复运行；本期结果表新建写入另见下方 writer 行；完整 S1-F01 状态保持未验证。 |
| `base_audit.workflow.defaults` 的 `merge-audit-results` / `history-enrich` DAG 合流与历史归类（输入为已保存公式结果副本） | `audit_core::s1_issue_chain::merge_s1_issue_sources` + `audit_core::s1_history::process_s1_history`，由 `open-data-audit s1-audit-directory` CLI 接入公式 Issue、条件格式 Issue、历史读取/归类 | 未验证（S1 Issue 合流与历史分类局部；机构字段采用 UOS/native 目标语义） | `probe/reports/s1_issue_chain_real.json`：1 组真实只读源/审核副本和模板，146 条公式 Issue + 2 条条件格式 Issue；按 UOS/native Python handler 元数据口径，Python/Rust 当前 Issue 与合并历史均为 148 条、逐字段差异 0；条件格式评估 602 格；输入、审核副本和模板读前/读后未变。新入口先做表结构预检，要求公式副本清单精确覆盖通过预检的源文件，按每个源文件先公式、后条件格式合流。合成 CLI 测试验证顺序、审核副本路径重绑、空历史与非空历史工作簿；非空样例验证 1 条连续 Issue、1 条已整改 Issue、反馈/审核意见继承，以及 Calamine 引擎和读取耗时双视图日志。真实样本文件名为满足 Linux 路径组件限制使用 ASCII 别名，机构身份按别名解析；此项未验证原始长文件名身份解析。源 Python UOS/native handler 对条件格式 Issue 留空机构字段，而 Windows/general handler 会传入；真实样本对拍按 Linux 目标注册的 native handler 语义执行。只读取已保存公式结果，不执行公式复制/重算，该 Issue 链不写回源/历史工作簿；可选本期结果 XLSX 由下方独立 writer 负责；真实非空历史库、完整 GUI、重复运行及目标平台仍未验收。 |
| `base_audit.native.result_writer.write_audit_summary` | `audit_core::s1_result_workbook::write_s1_result_workbook`，由 `open-data-audit s1-audit-directory --result-workbook <path>` 在 Issue 合流/历史分类后调用 | 保留（本期结果 XLSX 新建/链接局部；S1-F01 整体未验证） | `probe/reports/s1_result_workbook_real.json`：真实样本 299 条 Issue；Python/Rust 当前与合流 Issue 全字段差异 0；结果表 299 行、299 条链接，值、类型、格式、样式、列宽、筛选、冻结和超链接目标/定位/显示/字体样式差异 0。引擎 `rust_xlsxwriter 0.101.0`，结果写入阶段 5,872 ms，完整 CLI 15,899 ms / 20,456 KiB 峰值 RSS；同次运行的自动模板选择与 Python 基线一致；输入未改变，未执行公式重算。资源数据仅适用于该样本，未定义或宣称通用容量上限；GUI 与目标 Linux 未验证。 |
| `base_audit.history.merge_history` + `classify_current_issues` + `base_audit.native.history_io.read_history_xlsx` | `audit_core::s1_history::{merge_history, classify_current_issues, process_s1_history, read_s1_history_workbook}`；`open-data-audit s1-history-classify --history-workbook` 使用共享 Calamine reader 读取历史表并替换 JSON history，`--history-sheet` 支持指定表，默认选择“核查表校验结果”，再回退“历史审核结果” | 保留（Issue 历史归类、人工意见继承和历史工作簿只读读取局部；写回与完整 S1-F01 仍未验证） | `probe/reports/s1_history_classification_parity.json`：8 条合成历史记录、4 条本期 Issue；新增/连续/再次出现/已整改分类、机构范围过滤、手工意见继承和合并结果与只读 Python 基线差异 0。`probe/reports/s1_history_workbook_parity.json`：标准/指定/旧版工作表的合成工作簿读取与 Python `read_history_xlsx`、分类结果逐字段差异 0；覆盖 2 条历史、1 条连续 Issue、1 条已整改 Issue、手工意见继承、空身份行跳过及超过初始读取列数的“数据期”列。只读源项目配置工作簿的历史表为空，Python/Rust 均读 0 条、差异 0。普通/开发者日志验证实际 reader 和耗时；Windows Python 3.12 驱动 WSL Ubuntu Rust CLI 对拍，`cargo test -p audit-core --locked --quiet` 为 69 个单元与 44 个集成测试通过，`cargo fmt --all -- --check` 通过。Windows 原生 Rust、非空真实历史库、历史表写回、麒麟/UOS、deb 和 S1-F01 完整流程仍未验证。 |
| `base_audit.discovery.recommend_template` 的按工作表集合模板推荐 | `audit_core::s1_template::{recommend_template, recommend_template_for_workbooks, recommend_from_sheet_sets}`，经 `open-data-audit template-suggest` CLI 调用并用共享 namespace-aware OOXML sheet-directory adapter 读取工作表名 | 保留（模板推荐局部） | `probe/reports/template_recommendation_parity.json`：4 个合成场景逐字段匹配只读 Python 基线，覆盖源工作表为模板子集、混合报表类型停止自动选择、候选接近时要求人工确认、无可信模板。`probe/reports/template_recommendation_real_smoke.json`：只读源 `config-real` 的 8 个模板和 2026-07-31 S1 目录中的 68 个 `.xlsx` 源文件，Rust 与 Python 的匹配状态、置信度、说明和候选一致，置信度 1；读前/读后模板和源文件集合 SHA-256 未变。目录入口已将推荐接入 `conditional-audit-directory --template-dir`，按已发现/已筛选源工作簿匹配并排除模板文件，歧义时失败关闭；本轮另接入 `s1-audit-directory --template-dir`，先按 `--selected-file` 筛选源再推荐并排除模板候选。合成测试覆盖自动/显式模式及普通/开发者日志；真实模板选择与 Python 基线一致，299 条真实 Issue 和本期结果 XLSX 对拍差异 0，详见 `probe/reports/s1_result_workbook_real.json`。真实自动推荐及条件审核见 `probe/reports/s1_template_directory_audit_real.json`：21 个模板候选中 Python/Rust 推荐一致、置信度 1；1 个工作簿、4 条规则、602 格、2 条 Issue，Issue 字段差异 0。该单次 WSL2 目录树性能观察为 Rust Release OOXML metadata adapter 中位数 1.37 秒、Python 基线 1.76 秒；同一任务使用 Calamine 的 Rust Release 中位数 3.70 秒。此数字只描述该样本和完整进程耗时，不外推到其他模板、目录、引擎任务或目标平台。仅是单个真实目录树的模板推荐；不覆盖真实配置全部模板、其他期间、S1 规则计算/条件格式、公式重算、输出回写、历史、重复运行、Windows 原生运行、麒麟/UOS 或 deb，S1-F01 仍未验证。 |
| `ConditionalFormatService(PYTHON)` 的静态条件格式审核与 Issue 生成 | `conditional_format::evaluate_xlsx_workbook_with_dump` + `s1_conditional_audit::{extract_conditional_issues, extract_conditional_issues_batch}`，由 `open-data-audit conditional-audit`、`conditional-audit-batch` 和 `conditional-audit-directory` 接入 | 保留（条件格式审核结果生成及批量合并切片；非完整 S1-F01） | `probe/reports/s1_conditional_audit_real.json`：一个只读真实工作簿包含 4 条规则、366 个评估单元格；Rust 与 Python 服务输出 348 条 Issue，按 `issue_id` 对齐后所有 Issue 字段 0 差异。本轮另以 `conditional-audit-batch` 入口对真实中文文件名工作簿复验：4 条规则、366 个评估单元格、348 条 Issue，Rust 与 Python 逐字段 0 差异，原始文件名的 Issue identity 前缀 348/348 与 Rust 归一化一致；此 348 条样本使用规则 AppliesTo 默认范围。Audit 的 `name_config::resolve_named_ranges_for_mappings` 覆盖精确名/下划线后缀、工作表过滤、地址规范化和稳定去重；新增单文件、批量和目录 CLI 的 `--template-workbook` 与可选 `--mapping-config` 接线，按模板 Defined Names 解析条件扫描区和结构区，并从目录输入与期间推断中排除模板/配置工作簿。另一真实只读样本对拍见 `probe/reports/s1_conditional_named_area_real.json`：模板 17 个 Defined Names，按映射解析 4 个条件区域和 8 个结构区域，4 条规则评估 602 格并生成 2 条 Issue；Python/Rust 全字段差异 0，未解析匹配名称 0。 真实目录选择样本对拍见 `probe/reports/s1_conditional_selected_directory_real.json`：目录候选 2 个，按 `--selected-file` 仅审核 1 个；期间只从选中项识别为 2026-07，Rust/Python 均评估 602 格、生成 2 条 Issue，全字段差异 0。Rust 通过共享 Calamine + namespace-aware OOXML reader 读取一次，并在 Rust 侧构造业务 Issue；公式只读取缓存、不重算。合成 evaluator/Issue/批量 API/批量 CLI 测试通过，批处理按文件使用 Rayon 并行、按输入顺序合并且保留每文件机构元数据；任一文件失败时整批失败且不返回部分结果。仅覆盖该样本和已实现的静态规则；未迁移完整 S1 配置校验和普通规则、历史工作簿写回及重复运行；本期结果 XLSX 新建见下方 writer 行；`s1-audit-directory --template-dir` 现按已筛选源工作簿自动推荐模板，歧义时失败关闭；合成双视图 CLI 测试及本行 299 条真实 Issue 对拍通过。Windows 原生真实样本已验证：`conditional-audit-directory` 显式选模并筛选 1 个源工作簿，Python/Rust 全字段差异 0（4 条规则、602 格、2 条 Issue），见 `probe/reports/s1_conditional_audit_windows_native_real_20261024.json`；源与模板只读。目标 Linux/deb、全量批次和完整 S1-F01 仍未验证。无模板映射时仍可用 `--structure-area` 显式提供 `check_field` 结构范围。 |
| base_audit.native.openpyxl_workbook.read_template 的普通公式模板规则发现与校验 | audit_core::s1_template_rules::load_s1_template_rule_plan，由 open-data-audit s1-template-rules CLI 接入；读取经共享 Calamine + namespace-aware OOXML reader | 保留（普通公式模板规则发现/校验局部；不代表完整 S1-F01） | probe/reports/s1_template_rule_discovery_real.json：1 个真实只读模板，Python/Rust 完整 TemplateDefinition 逐字段对拍 0 差异；1072/1072 规则启用，4 个公式命名区域、8 个结构范围、4 个结果提取范围、380 keyword / 692 error_only 规则一致；公式规则编号、坐标、说明、范围和模式均逐项一致。Rust 读取批注的公式文本，不重算或修改模板；合成命名区域/结构规则与 CLI 双日志测试通过。仅验证一个真实模板，未覆盖结构表真实模板、公式副本写入/重算、Issue 生成与结果回写；S1-F01 仍未验证。 本轮新增 `assess_s1_template_health`，接入 `s1-template-rules` CLI 输出及 `s1-audit-directory` 前置检查：汇总启用规则、可生成问题标记的规则、有效结构区域，并报告公式 `#REF!` 警告；无结果规则或无结构区域时在解析公式副本清单前停止，未重算或修改工作簿。真实模板 Python/Rust 体检字段及项目差异 0（1,072 条启用规则、380 条结果规则、8 个结构区域、0 条 `#REF!`），见 `probe/reports/s1_template_rule_discovery_real.json`；定向 Rust 单元/CLI 测试通过。Windows 原生真实模板 Python/Rust TemplateDefinition 与健康检查逐字段差异 0（1,072 条规则、4 个公式命名区域、8 个结构范围、4 个结果提取范围），见 `probe/reports/s1_template_rule_discovery_windows_native_real_20261024.json`；`serve` HTTP API 真实检查 1,072 个公式、0 个公式错误、3 组命名区域，普通/开发者日志均含实际引擎与耗时，见 `probe/reports/s1_template_inspection_web_api_windows_real_20261024.json`。该 Python 3.12.10 对拍仅作基线，不代表隔离 openpyxl worker 或 deb runtime。公式副本写入、重算和完整 S1-F01 仍未验证。 |
| `base_audit.discovery._periods_in_text` | `audit_core::discovery::periods_in_text` | 保留（局部） | `probe/reports/period_parser_parity.json`：49/49 个 Python 基线向量一致；仅覆盖纯文本期间解析，不覆盖目录扫描、`detect_period` 或任何端到端业务流程。 |
| `base_audit.discovery.source_workbooks` + `detect_period` | `audit_core::discovery::{source_workbooks, detect_period}` | 保留（局部） | `probe/reports/period_detection_parity.json`：9/9 个合成目录场景与 Python 基线一致，覆盖递归深度、输入过滤、文件名识别、目录回退及冲突输出；不覆盖真实模板、文件系统错误策略或端到端业务流程。 |
| `base_audit.period_compare.list_period_pairs` + `base_audit.report_source.report_source_warning` | `audit_core::period_pairs::{list_period_pairs, report_source_warning}` + `audit_core::report_source::report_source_identity`；工作表名经共享 `excel_reader::read_sheet_names` 使用 Calamine | 保留（前端预览局部） | `probe/reports/period_pairs_parity.json`：1 个合成目录场景、6 个配对行及 banks 提醒与只读 Python 基线逐字段一致；覆盖标准逐机构文件、reports、banks、多工作表、临时文件过滤、重复身份最后项规则和单边项目，路径按文件名归一。工作表名只取 Calamine workbook sheet directory，不读取单元格正文。没有有效 `.xls` fixture，故该格式的 sheet-name 读取仍未验证；不覆盖真实报送目录或端到端业务审核。 |
| `base_audit.models` 的 10 个 dataclass 与两个结果摘要方法 | `audit_core::models` | 保留（模型层） | `probe/reports/models_parity.json`：10 个模型类型、11 个实例、7 组默认值和结果计数/摘要文本与 Python 基线一致；不代表任何审核、预检、模板或文件写入流程已经迁移。 |
| `s3_central_statistics.csv_importer` 的 CSV 解码与数据标准化分支 | `audit_core::central_csv::parse_central_csv` | 保留（S3 导入局部） | `probe/reports/central_csv_parity.json`：3 个合成 CSV 场景逐字段与只读 Python `read_central_csv` 基线一致，覆盖 GBK、UTF-8 BOM、带逗号/双引号字段、空记录行号、日期解析、单位换算/豁免、异常值、短行补齐和重复键后值覆盖；不覆盖 XLS/XLSX reader dispatch、真实批次和端到端 S3 流程。 |
| `base_audit.template` 的模板名、批注、公式结果和规则行纯解析 helper | `audit_core::template::{normalize_template_name, parse_comment_rule_id, clean_rule_comment, parse_formula_result, parse_rule_rows}` | 保留（纯解析局部） | `probe/reports/template_parser_parity.json`：24 个 fixture case 与只读 Python 基线逐项一致（模板名 6、批注 7、公式结果 7、规则行 4）；覆盖 Unicode 15.0 十进制数字的日期/区域标记。只覆盖值级 helper，不覆盖 workbook/批注读取、ExcelSession、规则执行或审核流程；JSON 数组/对象的 Python `str()` 等价性未验证。 |
| `base_audit.name_config` 的功能映射数据类及纯值 helper | `audit_core::name_config::{FeatureMapping, split_names, is_limited, features_of_type, matches_named_range}` | 保留（纯配置 helper 局部） | `probe/reports/name_config_helpers_parity.json`：31 个 fixture case 与只读 Python 基线逐项一致（映射 4、名称拆分 5、限定标记 9、按类型筛选 3、命名区域匹配 10）。不覆盖 config.xlsx 读取/校验/保存、默认模块/流程生成或工作簿行为；Python `casefold()` 的特殊 Unicode 展开字符未验证。 |
| `systems.s2_report_collection.unit_conversion` 的金额类型标记、Decimal 解析、单位缩放与 `to_yuan` 兼容入口；`importer._normalize_fast_value` / `_decimal_number` 的单元格数值适配 | `audit_core::unit_conversion::{is_amount_data_type, as_decimal, convert_amount, to_yuan, decimal_number_from_cell, convert_amount_from_cell}`；共享类型 `excel_reader::CellValue` | 保留（金额单位 helper 与 typed-cell adapter） | `probe/reports/unit_conversion_parity.json`：62 个 fixture case 与只读 Python 基线逐项一致，包括 43 个既有文本/helper case 和 19 个 reader cell case；reader cell case 直接对拍源 importer 的 XLS/XLSX 归一化、Decimal tuple 和单位转换结果。Calamine 值的共享 DTO 保持原 `WorkbookDump` JSON 结构。S2 目录导入调用该适配器的范围另见下一行；此 helper 报告不覆盖真实工作簿或规则/周期/外部比较流程。 |
| `systems.s2_report_collection.config.{load_config, load_unit_settings}` 的 S2 运行配置工作簿解析 | `audit_core::s2_config::{load_config, load_unit_settings}`，通过共享 `excel_reader::read_value_columns_with_iso_dates` 使用 Calamine；配置可投影到现有导入/周期/规则/外部比较模型 | 保留（合成 XLSX 配置加载局部） | `probe/reports/s2_config_parity.json`：15 个合成 XLSX 配置场景逐字段/错误文本与只读 Python 基线一致，包含 3 个成功配置、12 个预期配置错误，以及独立单位设置读取；覆盖设置默认值、表头/别名、指标、机构、外部规则、环比策略/旧环比表、表达式规则、告警与错误。源 `config.py`、`models.py`、`unit_conversion.py` SHA-256 记录在报告中。配置检查器局部切片见下一行；默认工作簿生成、真实配置文件、服务编排和报告输出未迁移；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.config_check.{check_s2_config, format_config_report}` 的只读配置结构、策略区间和运行时解析检查 | `audit_core::s2_config_check::{check_s2_config, format_config_report}`；结构读取经共享 `excel_reader::read_value_columns_with_iso_dates`（包括 Calamine worksheet extent 供统计），运行时检查复用 `s2_config::load_config` | 保留（合成 XLSX 配置检查器局部） | `probe/reports/s2_config_check_parity.json`：19 个合成 XLSX 案例的结构化报告与 UI 文本均与只读 Python 基线逐项一致，含有效配置、旧审核级别兼容警告、策略空档/重复、日期值错误提示、完整综合错误和必需工作表缺失；对拍同时验证输入夹具字节不变。Python 基线记录 `config_check.py`、`config.py`、`models.py`、`unit_conversion.py` 的 SHA-256。未覆盖真实业务配置、损坏/加密工作簿的错误文案、服务阻断流程接线、默认工作簿生成、S2 报告输出、非 XLSX 格式或目标 Linux 平台；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.importer.import_period_directory` 的文件发现、来源身份、A-C 列指标行提取、机构/指标补充、金额单位归一、去重和警告 | `audit_core::s2_import::import_period_directory`，通过共享 `excel_reader::read_value_columns` 使用 Calamine | 保留（合成 XLSX 目录导入局部） | `probe/reports/s2_import_parity.json`：12 条记录、2 对重复项、3 条警告与只读 Python importer 结果逐字段相同；夹具覆盖标准名/按机构名映射、两张表单、余额/累发/百分数/个数/文字/零值、无效金额、未识别文件跳过及大于 `i64::MAX` 的整数指标代码归一。只覆盖合成 `.xlsx` 夹具；有效 BIFF `.xls`、`.xlsb`、真实工作簿、公式缓存边界、文件读取失败策略、进度回调、配置加载、周期/规则/外部比较、报告写入及服务路由均未验证，尚未接入生产默认路径。 |
| `systems.s2_report_collection.importer.import_external_workbook` 的外部工作簿解析 | `audit_core::s2_external_import::import_external_workbook`，通过共享 `excel_reader::read_value_columns_with_iso_dates` 使用纯 Rust Calamine | 保留（合成 XLSX 外部导入局部） | `probe/reports/s2_external_import_parity.json`：12 个合成输入场景与只读 OpenPyXL 基线逐字段/错误文本一致；覆盖普通/时序布局、1900/1904 日期系统、日期与数值归一、空值/非数值警告、机构参照表优先级/旧表兼容、重复日期行/重复指标/不完整表头和不支持扩展名拒绝。基线 Python 3.12.10，源 `importer.py` SHA-256 为 `EEF44740DE8EB7AB0E09DE7C074FF06F22F515CC9A1518982FA662D123A49BEC`。未覆盖真实外部工作簿、成功解析 `.xls/.xlsb/.et/.csv`、S2 服务编排/外部比较集成/报告写入、性能及目标平台；尚未接入生产默认路由，S2-F01 仍未验证。 |
| `systems.s2_report_collection.validation.{validate_import, validate_config_warnings, validate_external_warnings}` | `audit_core::s2_validation::{validate_import, validate_config_warnings, validate_external_warnings}` | 保留（质量发现转换局部） | `probe/reports/s2_validation_parity.json`：11 条 Python `AuditFinding` 逐字段对拍通过，覆盖重复项/跨表单冲突、缺少机构或表单身份、导入警告、配置警告及外部数据警告；路径仅为跨平台比较归一为文件名。发现生成已迁移，S2 规则引擎、配置工作簿加载、外部导入及报告写入未验证，不代表 S2-F01 流程完成。 |
| `systems.s2_report_collection.exporter.build_summary` 的报告统计数据生成 | `audit_core::s2_summary::build_summary` | 保留（报告统计数据局部） | `probe/reports/s2_summary_parity.json`：2 个合成案例与只读 Python `build_summary` 结果一致，覆盖唯一机构/非空表单计数、指标及发现总数、异常发现分类、空类型/空级别和空输入；源 `exporter.py` SHA-256 为 `EF99D056BA9038BC9AED8BD23D5D8D0D2813941BA54366C1136380AC70A2E02B`。只覆盖摘要聚合，不覆盖 S2 服务编排、XLSX 报告输出、真实模板、性能或目标平台；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.period_engine.PeriodComparisonEngine` 的两期比较计算 | `audit_core::s2_period::compare_periods` | 保留（规则计算局部） | `probe/reports/s2_period_parity.json`：13 个合成规则场景与只读 Python 基线逐字段一致，产生 15 条发现并覆盖 2 个预期规则配置错误；覆盖数值/文字/缺失项、百分数百分点、Decimal 金额、策略区间、严重度、计算轨迹和错误分支。基线源文件 SHA-256 记录在报告中。该切片使用已导入的 `IndicatorRecord` 和合成 JSON 配置，不覆盖配置工作簿加载、报表导入、外部比较、规则表达式、服务编排、报告写入或真实模板，S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.rule_engine.RuleEngine` 的累计不降与表达式规则计算 | `audit_core::s2_rule_engine::evaluate_rules` + `audit_core::s2_expression` | 保留（规则计算局部） | `probe/reports/s2_rule_engine_parity.json`：13 个合成案例与只读 Python 基线逐字段一致，产生 10 条规则发现和 3 条规则质量发现；覆盖累计不降、当前/上期占位符、旧八段指标码、Decimal 表达式、逻辑函数、Thd 金额单位换算、缺失/无效值诊断及来源/计算轨迹。`rule_engine.py`、`models.py`、`unit_conversion.py` 的 SHA-256 记录在报告中。仅覆盖已导入记录和内存配置，不覆盖配置工作簿加载、导入、外部比较、服务编排、报告写入或真实模板；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.service.PeriodAuditV2Service.run` 的已导入数据分析与发现汇总编排（不含文件/配置 I/O 与导出） | `audit_core::s2_analysis::analyze_imports` | 保留（内存分析编排局部） | `probe/reports/s2_analysis_parity.json`：2 个合成案例与只读 Python 服务组件顺序逐字段一致，覆盖本期首条逻辑键保留、环比/外部核对/表达式规则发现、配置/导入/外部警告、规则质量发现、发现拼接次序和摘要，以及无外部数据的空输入分支。Python 基线记录 `service.py`、`period_engine.py`、`external_engine.py`、`rule_engine.py`、`validation.py`、`exporter.py`、`models.py` 和 `importer.py` SHA-256。未覆盖配置检查/加载、工作簿导入、进度回调、XLSX 导出、真实模板、性能、目标平台或生产路由；S2-F01 仍未验证。 |
| `systems.s2_report_collection.exporter.ExcelExporter.export` 的三表 XLSX 报告写出候选 | `audit_core::s2_exporter::export_s2_report`，通过共享 `excel_writer` 新建工作簿、共享 `excel_reader` 重开校验 | 保留（合成 finding 到 XLSX 的局部候选；由 S2 service/CLI 候选调用，非生产默认） | `probe/reports/s2_exporter_parity.json`：5 个合成 finding、3 张工作表、116 个比较单元格（11 个数值格、1 个公式格），与只读 Python `ExcelExporter` 的文件级对拍 0 差异；数值相对容差 `1e-15`，公式缓存为空。对拍覆盖值/类型、公式与公式/文字区分、样式、列宽区间、冻结窗格、筛选、隐藏行和页面边距。`probe/reports/s2_exporter_python_baseline.json` 记录 Python 3.12.10、fixture SHA-256，以及源 `exporter.py` / `models.py` SHA-256。仅覆盖合成 finding；日期/时长值和不能精确转为 Excel `f64` 的整数明确拒绝。真实脱敏模板、性能、目标 Linux 或 deb 未验证；S2-F01 仍未验证。 |
| `systems.s2_report_collection.service.PeriodAuditV2Service.run` 的文件级 S2 编排（配置检查/加载、两期目录导入、期间/规则/可选外部比较、发现汇总、三表报告写出） | `audit_core::s2_service::run_s2_audit`，由 `audit-core/src/bin/open-data-audit.rs` 的 `open-data-audit s2` 接入 Rust CLI，连接已对拍的 `s2_config_check`、`s2_config`、`s2_import`、`s2_period`、`s2_rule_engine`、`s2_external_import`、`s2_external`、`s2_validation`、`s2_analysis`、`s2_exporter` | 保留（synthetic service + CLI 端到端候选；非生产默认） | `probe/reports/s2_service_parity.json`：无外部数据 15 条发现、含外部 `plain_layout.xlsx` 18 条发现；服务 API 与 CLI 的摘要/进度均和只读 Python 对拍 0 差异，CLI 生成的三表 XLSX 文件级比较 0 差异（分别比较 268/304 个单元格、36 个数值格）；阻断配置错误和进度一致，且不创建输出目录。相对容差 `1e-15`；Windows/WSL 绝对路径记法归一后比较；Python/Rust 输入哈希不变，Python 基线源哈希前后不变。夹具为合成文件，当前/上期复用相同报表内容并只调整文件名日期；外部读取进度中的引擎/耗时动态字段不逐字比较。未验证真实脱敏模板、全部输入/导出失败回退、性能、Windows 原生 Rust、麒麟/UOS/glibc 2.28 或 deb；S2-F01 整体仍未完成。 |

其他功能清单状态不因这些局部迁移对拍而改变，仍按上文逐项保持未验证/待确认。

### S1-F01 真实配置映射补充验证

用户提供的真实逐笔配置与个人贷款命名区域模板，经 Windows 原生 Python/Rust 完整规则定义和模板健康检查对拍，差异为 0（895 条规则、4 个公式复制区、8 个结构区）；模板未改变。另以同一真实配置的条件格式模板和匹配输入完成 S1 条件审核对拍：4 条规则、602 个评估格、2 条 Issue，字段差异 0，见 `probe/reports/s1_conditional_audit_config_real_windows_native_20261025.json`。该样本 Python 总耗时 281 ms，Rust `calamine+namespace_aware_ooxml` 总耗时 261 ms；仅为本样本观测，不用于路由决策。使用 `s1-audit-directory` 在真实源簿/既有审核副本上执行配置化 S1 链：当前公式 Issue 297 条、条件格式 Issue 0 条，Python/Rust 当前集及合流历史分类字段差异均 0；模板验证扫描 1,072 个公式、0 个错误，全部输入未改变。该链使用已有审核副本值，未执行公式复制或重算，报告见 `probe/reports/s1_audit_directory_config_real_windows_native_20261025.json`。模板规则报告：`probe/reports/s1_template_rule_discovery_config_real_windows_native_20261025.json`。以上证据只覆盖真实配置加载/模板规则发现、条件格式 Issue 切片和已有审核副本上的 S1 计算/分类链；公式复制、历史簿写回、公式重算及完整 S1-F01 仍未验证。

### S1-F01 配置目录入口真实对拍

使用 `config-real` 的 `2.报表采集系统` 配置备份和单位贷款命名区域模板，经 Windows 原生 `open-data-audit conditional-audit-directory` 目录入口运行 1 个根目录工作簿；Python/Rust 各 2 条条件格式 Issue，字段差异 0，模板推荐一致，双方各评估 602 个单元格。Rust 读取 2 条 S1 映射、12 个命名区域并扫描 4 个条件区域、8 个结构区域；该模板规则发现中，真实配置与内置映射的 1,072 条规则定义完全一致。Rust 982 ms、Python 375 ms 仅为该样本观测。原配置为 `.xlsx.bak`，为满足 Python 对拍读取，将相同字节复制至系统临时 `.xlsx`，测试结束已删除；真实输入只读，未写工作簿。用户批次目录递归包含 184 个 XLSX 且混有多个模板类型，整目录推荐按设计拒绝，故本证据仅覆盖根目录单簿；详见 `probe/reports/s1_conditional_audit_config_real_directory_windows_native_20261025.json`。S1-F01 公式复制/重算、既有历史簿写回和整体流程仍未验证。

### S1-F01 config-real 公式复制 Windows 原生候选复验

按用户提供的 `config-real` S1 映射和 2026-07-31 批次，选取同一真实模板族的 7 个带缓存审核副本；模板含 1,072 个公式、752 个公式缓存、4 个复制区和 8 个结构区。Windows 原生 Rust 完整 60 操作计划经隔离 OpenPyXL worker candidate-only 测试，7/7 在写前因 `target_merged_non_anchor_payload` 拒绝，未创建候选，模板和目标输入均未改变。七个目标分别有 983、688、692、686、692、688、686 个非空公式缓存，均按工作表与单元格坐标落在四个 ImportRange 内，范围外为 0。既有 `s1_formula_copy_target_merged_payload_guard_20261025.json` 对两个定向样本记录每本 855 个合并非锚载荷，全部位于复制区外（31 个载荷合并区域）；没有保留的 Python 写后工作簿可核对这些载荷是否保留，因此该项仍未验证。共享提交 `799cfe7` 的行列尺寸读回校验未被触发。OpenPyXL 证据见 `probe/reports/s1_formula_copy_candidate_openpyxl_windows_native_20261025.json`；同模板族首个带缓存副本的 Umya full-S1 候选写出并通过 Rust 读回，但对拍共 1,036 项差异（复制区值/公式各 78、样式 802、批注 74；范围外批注 4），输入未变，不可路由，见 `probe/reports/s1_formula_copy_candidate_umya_windows_native_20261025.json`。两类候选均不能作为 S1-F01 生产路由。

### S1-F01 目标合并非锚载荷基线等价取证

针对上述 `target_merged_non_anchor_payload` 阻塞的关键未验证项（Python 基线是否保留这些载荷），新增只读取证脚本 `tests/compare/probe_s1_target_payload_baseline.py`，在同一真实模板族 7 本带缓存审核副本的临时副本上执行两条写穿路径并按共享守卫同一语义清点：每本原始目标簿 855 个合并非锚载荷（4 区外 855、区内 0；37 个合并区域中 31 个含载荷）；Python 基线 `copy_formula_ranges_from_template`（openpyxl load+修改+save）写后载荷保留 0/855、新增 0；纯 openpyxl load+save 同样保留 0/855；两路输出的 37 个合并区域均不变，7/7 输入未变。模板口径交叉核对 4 区/1,072 公式/55 区内批注与既有候选报告一致。结论：既有 Python 产品基线本来就不保留这些载荷，openpyxl worker 的丢弃与基线逐格等价；守卫防御的内容损失与已验收 Python 行为完全相同。即使预检放行，共享读回校验器仍会在这些单元格报 `unmodified cell value changed`，因此解除阻塞需要 main 侧计划级 opt-in（目标侧预检按 ImportRange 目标矩形收窄 + 读回校验对未变更合并区域的非锚空置容忍），已按规范提交共享能力请求 `shared-capability-requests/20261025_openpyxl_target_merged_payload_baseline_equivalence.md`，Audit 不自行修改共享 crates。机器可读报告见 `probe/reports/s1_formula_copy_target_payload_baseline_20261025.json`；该取证不改变任何能力状态，S1-F01 仍未验证。

### S1-F01 非空历史簿真实链路烟测（跨模板）

Windows 原生 Python/Rust 对拍将 `config-real` 的非空历史簿接入 2026-07 S1 审核 CLI：7 组源簿/既有公式结果副本均由模板推荐器匹配同一当前模板；每组读取 9 行历史，7 组共 159 条当前 Issue，Python/Rust 当前 Issue、历史分类输入、当前结果簿与独立历史结果簿差异均为 0，历史输出共 222 行，所有输入均未改变。脱敏汇总见 `probe/reports/s1_f01_nonempty_history_real_batch_windows_native_20261025.json`。该历史簿未匹配当前模板（推荐 `matched=false`，confidence 0.63），且本批未观察到历史 Issue ID 重叠；因此只证明真实非空历史读取和无重叠合流路径，不证明同业务连续性/重复归类。使用已有公式结果副本，未运行 Rust 公式复制/重算；只新建独立历史结果簿，没有改写既有历史簿，既有 PageSettings 保留未验证。S1-F01 整体仍未验证。

### S1-F02 等线全角列宽真实样本校准

Windows 原生 Microsoft Excel/Python 与 Rust 对拍将全角字宽估算由 15px 校准为 14px；7 个真实源簿、27 张源表的输出表名/结构/值/类型/数字格式/字体/对齐/筛选/冻结差异均为 0，列宽最大差从 3.07421875 降至 0.7890625 字符，仍有 48 列宽不同，不能视为格式契约通过。合成任意行/固定行/历史/兜底对拍的值与样式差异为 0，列宽仍有估算差异。完整列宽精确性及 S1-F02 保持未验证；本次真实输入未变，输出位于临时目录。证据见 `probe/reports/s1_f02_dengxian_width_calibration_real_windows_native_20261025.json`。

边界修复：`CellValue::Integer` 的估算现在按完整十进制整数文本计位，避免先转 `f64` 后以科学计数表示而低估超大整数列宽；定向单测覆盖 `10^18`。本修复未重新测量上述 48 个真实列宽差，不改变 S1-F02 状态。
公式缓存兼容修复（Windows 原生）：S1-F02 仅在实际消费的表头、数据、附加元数据和历史键/字段范围读取 Calamine 返回的已保存公式结果，不启动 Excel/LO 重算；缺缓存或缓存为 Excel 错误值的公式在消费范围内失败关闭，范围外公式不再阻断汇总。旧 Python `ExcelSession` 在打开工作簿前切换为手动计算，区域汇总读取 `Value2`；历史表由 `openpyxl data_only=True` 读取，与 Rust 的缓存值语义一致。定向 Rust 单测 1 项、S1-F02 CLI 集成测试 5 项通过，覆盖源公式/历史公式缓存、源与历史输入字节不变、未缓存消费公式拒绝及范围外公式忽略；`open-data-audit s1-region-summary --help` 正常。Windows 原生临时合成簿中，Python Microsoft Excel 与 Rust CLI 的源公式输出值/结构逐格差异为 0（缓存值 2、输入未变）；内置本地校验历史公式缓存亦逐格一致（匹配 1 行、历史值“公式缓存|未通过”、输入未变）。仍未使用真实 F02 模板或进行 Linux/deb 验收；S1-F02 整体仍未验证。

### S5-F01 配置 DSL 解析与静态校验切片（真实配置对拍）

将源项目 `s1audit_configured.dsl` 的解析与静态校验半边（词法/语法分析、`parse_expression`、`parse_dictionary`、`validate_call`、`validate_node`、`validate_static_types`、`static_kind`、`validate_dictionary_dependencies`、`variables_in`、`calls_in`）忠实移植为 `audit-core/src/s5_config_dsl.rs`，错误消息与位置逐字对齐 Python；`evaluate` 运行时属后续规则执行切片。Windows 原生对拍以真实 config-real 生效配置（`5.逐笔统计系统_审核配置.xlsx`）的 2,392 条启用规则为样本，比较器按 `load_configuration` 规则循环的 DSL 段逐行复刻（字典解析→依赖检查→字典静态类型→各表达式解析→静态类型与字段种类约束→未定义变量检查），Python/Rust 逐字段比较状态、错误类别与原始消息：2,392/2,392 匹配、0 差异，全部规则两侧均通过校验（python_ok/rust_ok 2,392），0 个公式单元格拒绝。合成单测 11 例覆盖优先级、连续比较、转义与位置、V 函数别名/具名参数/元组路径、字典循环引用、静态类型错误等。已知 Unicode 边界差：Python `str.isdigit` 接受全部 Unicode 十进制数字而 Rust 端仅接受 ASCII 加上标三符（后者随后以相同"无效数字"错误拒绝）。该切片不覆盖配置工作簿加载、适用范围/业务行级检查和规则执行；S5-F01 整体仍未验证。证据见 `probe/reports/s5_config_dsl_parity_windows_native_20261025.json`。

### S5-F01 配置工作簿加载层切片（真实配置整簿对拍）

将源项目 `s1audit_configured.workbook_config.load_configuration` 完整移植为 `audit-core/src/s5_config.rs`：审核规则表头校验（ID 列三种别名、必填字段、重复表头、公式单元格拒绝）、规则行装配（修订、启用标志、数值运算模式、表达式布尔值取反、生效日期校验）、规则循环全序检查（复用 `s5_config_dsl` 的字典/表达式段、`@当前行` 与适用范围交叉校验、业务推断 `infer_rule_business`、规则描述变量扫描）、可选目录表（模板指标目录/适用范围目录/外部指标目录/数据源契约/配置版本）的解析与全部单位/精度/唯一性校验，以及无指标目录时从 `V_CUR`/`V_PRE` 调用派生模板指标（`_metric_profile` 单位推断）与 `V_EXT` 派生外部指标。读取经共享 `excel_reader`（calamine）并在读取层按公式映射拒绝公式单元格；sha256 采用 workspace 已管理的 `sha2`（Cargo.lock 仅登记成员依赖，无版本变化）。Windows 原生整簿对拍：真实 config-real 生效配置 Python/Rust 双侧加载后逐字段比较 Configuration（2,392 条规则全字段、4,215 个派生模板指标、版本字段与 sha256），差异 0；真实配置无可选目录表，仅覆盖其缺席分支。已知读取差：openpyxl 会把日期格式数字单元格转为 Python 日期而本移植读存储文本（真实配置字段均为文本，不受影响）；配置工作簿不存在/不可读两类错误消息含运行时路径，不逐字比较。该切片不含规则执行（`evaluate`）、业务识别（`match_s5_business`）与审核结果输出；S5-F01 整体仍未验证。证据见 `probe/reports/s5_config_load_parity_windows_native_20261025.json`。

### S5-F01 业务识别与逐文件分类切片（真实 120 文件对拍）

将 `recognition.match_s5_business` 与 `classify_s5_files` 移植为 `audit-core/src/s5_business.rs`：业务候选按配置规则工作表集合对源文件表集合的包含关系筛选，唯一候选直取，多候选按模板/文件名文本最长匹配收敛，冲突/缺失消息逐字对齐 Python；逐文件分类复用已对拍的表集合评分（含显式模板权威、0.05 歧义、0.90/覆盖兜底）与期间提示，输出模板/匹配状态/业务/期间/明细及计数字段，明细字符串（含「源工作表匹配 x/y 张」「模板中其他工作表」「多出」段）逐字一致。真实对拍：2026-07-31 逐笔统计系统样本子树 120 个 XLSX/XLSM、模板目录与生效配置均以字节级临时副本参与（源树只读；Python TemplateCatalog 的索引写落在临时副本旁），Python/Rust 120 行全字段差异 0，匹配状态 87 matched/14 partial/19 unmatched，业务识别覆盖 18 个业务（单位贷款 22、个人贷款 29 等）。已知读取差：Python `read_xlsx_sheet_names` 在 match_s5_business 内传播 XML 解析错误而 Rust 端将读取失败映射为空表集合（真实语料均可读）；TemplateCatalog 索引持久化未移植（对输出无影响）。规则执行与审核结果输出仍未迁移，S5-F01 整体仍未验证。证据见 `probe/reports/s5_classify_business_windows_real_20261025.json`。

### S5-F01 XLSM 源工作簿元数据匹配切片

`open-data-audit s5-match-template` 现接受 `.xlsx` 与 `.xlsm` 源文件；仅读取工作簿的 sheet-name 元数据，模板目录仍只收 `.xlsx`，与 Python `TemplateCatalog.glob("*.xlsx")` 一致。当前分支已有共享 reader 等效修复（`579155f`，与 main `3e53e7e` 的源码一致）。Windows 原生 Python 3.12.10/Rust CLI 合成对拍：1 个 `.xlsm` 扩展名 OOXML 源、1 个 `.xlsx` 模板，双方选择同一模板，输入前后 SHA-256 不变；真实用户样本子树盘点为 120 个 XLSX、0 个 XLSM。本对拍样本没有 VBA 部件，因此不验证真实 XLSM/VBA 内容；S5-F01 整体及共享能力请求 #8 仍未完成。证据见 `probe/reports/s5_template_matching_xlsm_synthetic_metadata_parity_windows_native_20261026.json`。

### S5-F01 配置校验入口切片（真实配置 3 scope 对拍）

将 `service.validate_s5_config` 的配置级分支移植为 `audit-core/src/s5_config_check.rs`：scope 门控（current/current_business/all/all_enabled 与未知范围拒绝）、`_configured_business_sheets` 轻量业务-工作表读取、`_infer_business_from_template` 模板业务推断（含文本匹配收敛与多业务拒绝消息）、`strict_current` 前置拒绝、加载错误映射（S5 配置语法或字典引用检查未通过）、规则/启用/业务计数、`_rule_references` 启用规则 V_CUR/V_PRE/V_EXT 引用提取（含 `@当前行` 动态引用计数），以及无模板/外部文件时 `not_checked`/`no_references` 两条路径与 message 组装。真实 config-real 生效配置 3 用例（current 无业务、current+业务、current_business 严格拒绝）Python/Rust 全字段（valid/complete/message/计数/businesses/business_checks/checks/warnings）差异 0。后续已接通当前业务的模板/外部地址检查，并新增 all/all_enabled 全量目录扫描：Rust 使用 `--template-dir` 读取 XLSX sheet 元数据、不写入模板目录索引，按 required sheet 匹配并对歧义安全停下；全量目录与外部地址的 6 个合成 Python/Rust 用例全字段差异 0，见 `probe/reports/s5_config_check_fullscope_synthetic_parity_windows_native_20261025.json`；config-real 真实模板目录和外部文件 6 个 Windows 原生用例按内容规范化后差异 0，19 个业务共检查 4,215 个模板地址、唯一定位 4,151 个、外部地址 43/43；64 个未定位结果两侧一致，见 `probe/reports/s5_config_check_fullscope_real_windows_native_20261026.json`。错误集合顺序按内容比较；损坏模板分支、Python TemplateCatalog 缓存语义和目标平台仍未验收；加载错误变体的 locations 在 Rust 侧为空（真实配置加载成功，未触发）。规则执行与审核结果输出仍未迁移，S5-F01 整体仍未验证。证据见 `probe/reports/s5_config_check_windows_real_20261025.json`。

### S5-F01 规则执行运行时切片（双数值模式 79 用例对拍）

将 `dsl.evaluate` 求值半边移植为 `audit-core/src/s5_eval.rs`：EvalValue 模型（number/string/boolean/tuple/blank、单位、指标类型、evidence 去重合并、is_literal、precision、数值模式）、全部运算符与内置函数（IF/AND/OR/NOT/SUM/ABS/ROUND/MIN/MAX/ISBLANK/ISNUMBER 与 V 调用常量转发）、比较类型规则、单位合并与不匹配错误、除零/类型/范围错误消息逐字对齐 Python。双数值模式：excel_binary64 经 `excel_round15`（15 位有效数字收敛）作用于数值构造、变量解析与比较两侧；decimal 经内部数字向量十进制的 28 位有效数字上下文（算术银行家舍入、ROUND 量化半进位、精确整除按理想指数截断尾零；中间工作范围溢出以显式错误拒绝）。Windows 原生对拍 79 用例：22 个合成用例（含 1/3 上下文舍入、2.675 半进位量化、单位不匹配、除零、字符串有序比较拒绝等错误路径）+ 真实配置逐业务抽样 57 条启用规则表达式（按规则声明的数值模式，字典变量与 V 调用以确定性 number-1 桩求值），Python/Rust 状态、数值、种类、单位、指标类型、错误类别与消息逐字段差异 0。已知边界：V 取数的真实源读取属规则引擎服务切片；evidence 合并仅以空 evidence 参与跨语言对拍（合并顺序由 Rust 单测覆盖）；十进制中间工作范围溢出为显式错误（业务数值远低于该范围）。审核结果输出仍未迁移，S5-F01 整体仍未验证。证据见 `probe/reports/s5_eval_runtime_windows_real_20261025.json`。

### S5-F01 布局区域发现切片（双真实模板对拍）

将 `layout.py` 的纯命名区域半边移植为 `audit-core/src/s5_layout.rs`：`NamedRegion`/`region_number`（尾缀 `_数字(#数字)` 解析与无界默认值）、`named_regions`（工作簿+工作表作用域 defined names、前缀过滤、多区域 `#序号` 后缀、`''` 引号工作表名回退、(sheet, 区域号, min_row, min_col) 排序）、`paired_column_region` 几何配对、`inferred_numeric_regions`（行列轴交叉推导 `数字区域_推导_NNN`、误指列轴的同序号几何复用含单元格内容校验、未使用轴的 missing 列表）、`numeric_regions`（数字区域/数值区域 → 条件格式区域兼容 → 交叉推导的逐表优先级与来源串）、结构区域分组（表结构/行索引/列索引区域按 (区域号, min_col, min_row, name) 排序）。区域引用串按 Python `NamedRegion.reference` 逐字组装。Windows 原生对拍：真实 config-real 单位贷款与个人贷款命名区域模板双例，named/numeric/推断结果/推断缺失/结构分组全字段差异 0；两例分别命中「条件格式区域（兼容现有模板）」与「表结构区域交叉推导」分支，均 4 个数字区域、0 缺失。已知边界：工作簿定义名经共享 `read_defined_names` 读取（多区域名以 Python `#序号` 约定展开）；`scan_rule_paths` 序列匹配属地址扫描切片未移植。S5-F01 整体仍未验证。证据见 `probe/reports/s5_layout_regions_windows_real_20261025.json`。

### S5-F01 地址扫描切片（真实配置+模板端到端对拍）

将 `layout.py` 的序列匹配半边移植为 `audit-core/src/s5_scan.rs`：合并单元格锚点簿（`_merged_by_row` 等价）、`_sequence`（锚点变化且非空才记 token）、`row_path_candidates`（逐行 token/路径、重复路径的连续行分组、最右重复列的唯一组限定别名提案、aliases/primary 语义）、`_contextual_row_hits`（基路径去项匹配 + 同列最近唯一组标签验证）、`_starts`、`in_numeric_regions`（锚点感知）、`projected_column_region`（几何配对优先、同序号跨表几何复用需目标内容），以及 `scan_rule_paths` 主循环（结构区域或整表回退、eligible 行与数字区域过滤、行命中别名/子序列、列扫描缓存、交集定位；未绑定源的数字区域为空按 Python 兼容回退）。`validate_s5_config` 接通模板地址检查分支：`_configuration_rule_locations` 行定位映射、`_config_locations` 错误位置、缺表/无法读取/多匹配/未找到等错误消息与结构逐字对齐，外部文件与全量分支保持显式未迁移。端到端对拍：真实 config-real 配置 + 真实单位贷款模板（当前业务 4 用例含模板用例），Python/Rust 全字段（valid/complete/message/businesses/business_checks/checks 含错误结构与 locations）差异 0；模板用例 valid=true/complete=false（@当前行动态地址需适用范围核对属 Python 语义）。已知边界：数字区域模板投影（`numeric_scope.project_numeric_regions`）属 V 取数源读取切片；错误字典 locations 依赖的行映射来自共享 calamine 读取。S5-F01 整体仍未验证。证据见 `probe/reports/s5_config_check_windows_real_20261025.json`。

### S5-F01 V 取数核心切片（真实配置+审核副本 8 调用对拍）

将 `sources.SourceRepository._retrieve_template` 的无契约单绑定路径移植为 `audit-core/src/s5_retrieval.rs`：`S5Retrieval::bind` 落实期间/机构门控（PERIOD_CONTRACT_MISSING/MISMATCH、INSTITUTION_CONTRACT_MISSING），`retrieve` 复刻 `_retrieve_template` 的参数解析（2/3 参形式、`_as_path` 校验）、工作表解析、指标目录候选（按 业务+行列路径 + 工作表过滤与业务回退）、CATALOG_METRIC_NOT_FOUND/METRIC_AMBIGUOUS/METRIC_NOT_CONFIRMED/UNIT_CONTRACT_MISSING/METRIC_NOT_FOUND/SHEET_NOT_FOUND 门，`_locate_template_value` 经 `_template_index_from_paths`（即已移植的 `scan_rule_paths` 直索引，按业务+表缓存）定位交点（ROW_PATH_NOT_FOUND/COLUMN_PATH_NOT_FOUND/STRUCTURE_ERROR/METRIC_AMBIGUOUS），`value()` 的公式拒绝（SOURCE_FORMULA_UNVERIFIED）与错误值拒绝（VALUE_TYPE_INVALID），`_convert_value` 的空值门（CELL_EMPTY/allow_blank）、文本/数值/整数类型检查、金额（元/万元/亿元因子）、比例（存储尺度与别名规范化）、计数/普通数值严格性换算与 `EvalValue.number` 十进制存储（默认 decimal 模式，随求值 `_with_numeric_mode` 收敛）。Evidence 记录含 Python Evidence 全部 22 字段（含 conversion 注记；raw_type 因双读取器单元类型差异在对拍中略去）。Windows 原生对拍：真实 config-real 生效配置 + 真实单位贷款审核副本（字节级临时副本），采样 8 条启用规则 V_CUR 字典调用（正常/空值/换算路径），Python/Rust 值、单位、指标类型、precision、evidence 全字段与错误 code/message 差异 0。已知边界：绑定契约（数据源契约身份/期间校验）、V_EXT 外部取数、数字区域模板投影（`numeric_scope.project_numeric_regions`）与检索缓存/追踪未迁移；求值侧 call_value 钩子的服务级接线属 S5 审核服务切片。S5-F01 整体仍未验证。证据见 `probe/reports/s5_retrieve_windows_real_20261025.json`。

### S5-F01 审核服务规则执行切片（真实配置+副本 6 规则对拍）

将 `engine.AuditEngine._evaluate_rule` 移植为 `audit-core/src/s5_service.rs`：`RuleHooks` 以 RefCell 接线变量缓存（未用变量集合、变量栈）、检索缓存（含错误缓存复抛）与 `_retrieval_trace` 追踪行（variable/function/source_id/sheet/requested_sheet/row_path/column_path/metric/cell/raw_value/resolved_value/单位与换算/状态与错误，路径按 Python `@` 连接字符串）；求值顺序复刻——字典解析与依赖校验（失败 CONFIG_ERROR）、适用表达式门（非布尔 CONFIG_ERROR、False → NOT_APPLICABLE 携 evidence 与未用变量）、审核表达式（非布尔 TYPE_ERROR；PASSED/TRIGGERED 按 `audit_expression_inverted`）、当前/对比/差值表达式（差值非数值/空 TYPE_ERROR）、`_render_description`+`_display_value`（binary64 走 Python `.15g` 语义、Decimal 原文、布尔 是/否、空值（空值））、三值单位/precision、变量与 evidence 汇总去重；错误路径按 RetrievalFailure code 或消息前缀分类（STRUCTURE_ERROR_CODES/CONFIG_ERROR_CODES 集合、`_error_code` 大写前缀提取）并保留缓存 evidence/变量/追踪；`@当前行` 未展开时报与 Python 相同的 CONFIG_ERROR。`EvalError` 增加 code 侧信道（取数失败码穿透求值栈），`s5_eval` 构造器放宽为 pub。Windows 原生对拍：真实 config-real 配置 + 真实单位贷款审核副本（同一临时目录双份输入），前 6 条启用规则（单位贷款）Python `AuditEngine._evaluate_rule` 与 Rust `evaluate_rule` 的状态、触发、错误码/消息、描述、三值/单位/precision、变量、evidence、检索追踪全字段差异 0（状态覆盖 PASSED/TRIGGERED 实分布）。已知边界：适用范围 @当前行 批量展开、结果输出与 ledger、多规则编排/CLI 属后续切片；取数为单绑定无契约路径。S5-F01 整体仍未验证。证据见 `probe/reports/s5_rule_evaluation_windows_real_20261025.json`。

### S5-F01 run 编排切片（整业务 377 行全量对拍）

将 `engine.AuditEngine.run` 的编排半边移植为 `s5_service::{run_rules, base_outcome, parse_period_date, within_span, within_rule_period, scope_paths, status_outcome}`：规则按 rule_id 分组（业务过滤），生效期内取最大修订，范围外 NOT_APPLICABLE（“本次报告期不在规则生效范围内”）、配置版本范围外 NOT_APPLICABLE（“本次报告期不在配置版本生效范围内”）；适用范围展开为该业务+工作表+行前缀的指标行路径（排序去重，空集 CONFIG_ERROR“适用范围没有匹配到任何模板行路径”），每条路径调用 `evaluate_rule`，`@当前行` 在 call_value 中替换为 scope 元组（未展开时同款 CONFIG_ERROR）；结果行携带规则元数据（rule_id/修订/名称/业务/期间/机构/严重度/校验类型/指标/默认表/公式格）与 scope_path（“ / ” 连接），统一经 `base_outcome` 构造；末尾按状态计数。取数侧补 V_PRE/V_EXT 未绑定的 Python 同款 SOURCE_NOT_SELECTED 消息；服务错误分类修正为取数失败码直取（状态由 STRUCTURE/CONFIG 集合判定）。Windows 原生对拍：真实 config-real 配置 + 真实单位贷款审核副本，整业务运行 Python `AuditEngine.run`（businesses=[单位贷款]、schema 快照落临时目录、不持久化）与 Rust `run_rules` 的 377 条结果行（元数据/状态/触发/错误码与消息/描述/三值与单位/scope_path）与状态计数全量差异 0（真实状态分布覆盖 PASSED/TRIGGERED/RETRIEVAL_FAILED 等）。已知边界：输入哈希完整性降级、结构快照比较、数据质量告警、outputs 写出、上期/外部绑定与 V_EXT 取数未迁移。S5-F01 整体仍未验证。证据见 `probe/reports/s5_run_orchestration_windows_real_20261025.json`。

### S1-F01 openpyxl worker 写穿与样式表越界缺陷定位

main 落地请求 #3/#4 后，S1 公式复制业务计划（`s1_formula_copy`）按裁决对目标簿全部工作表声明 `PreservationRequirement::MergedNonAnchorPayloads` 豁免（基线等价依据：7/7 真实样本证明 openpyxl load+save 丢弃全部 855 个合并非锚载荷、合并区域不变；载荷实际分布于 6 张表，含 2 张非复制区隐藏表）。共享预检豁免生效：真实 config-real 单位贷款模板 + 真实单位贷款机构审核副本的完整 60 操作计划首次通过 openpyxl worker 写穿（apply 通过、输出创建、10.6s、2 个空作者兼容注释）。读回校验暴露新缺陷：sheet7（集中系统数据）CJ2 在 worker 输出中样式索引 s=24 越界（cellXfs 仅 24 项），源 xf 0 携带的 vertical=center 语义丢失；同输入 Python 基线 s=24 在界内（cellXfs=38，xf24 保留 vertical=center）。worker 路线因此不满足基线等价，读回拒绝对。已登记共享请求 #5。S1-F01 公式复制仍无可用引擎，整体仍未验证。证据见 `probe/reports/s1_formula_copy_openpyxl_worker_e2e_windows_native_20261025.json`。

### S1-F01 显式公式副本重算接线（Windows，2026-10-26）

`open-data-audit s1-audit-directory` 新增显式 `--recalculate-sidecar <lo-convert>` 与 `--recalculated-copy-dir`，对公式副本清单中的副本调用既有 sidecar `recalculate`，并在后续公式结果提取前切换到新生成的独立审核副本。Windows 仅接受报告引擎 `Excel.Application`，Linux 仅接受 `LibreOffice`；不覆盖原副本，失败不发布结果文件，重算副本目录也从源文件发现中排除。Rust 定向测试 4 passed、1 ignored；既有 `s1_audit_directory_cli` 回归 5 passed；Windows `open-data-audit s1-audit-directory --help` 与 Audit CLI 编译通过。显式合成公式侧车测试实际调用 Windows COM，但 `lo-convert` 在 PowerShell 清理阶段以 `RPC_E_DISCONNECTED` 退出 2；Rust 断言输入字节不变且失败时没有发布输出。该 sidecar 能力仍不可用/未验收，已将具体清理失败复现交 main 修复；未运行真实业务样本，S1-F01 公式复制与完整链路仍未验证。

### S5-F01 CLI 接线切片（真实入口端到端对拍）

`open-data-audit` 新增 `s5-audit-run` 子命令（clap）：`--config` 配置工作簿、`--source` 本期绑定源、`--source-id`、`--period`、`--org-code`/`--org-name`、`--business`×N（缺省全业务）、`--output-dir`、`--run-id`（缺省时间戳）、`--developer-view`。执行链：`load_configuration` → `S5Retrieval::bind`（期间/机构门控）→ `run_rules`（编排+`@当前行` 展开）→ `write_run_outputs`（副本清单+审核结果.xlsx+运行总账），JSON 汇总输出规则数/状态计数/输出文件/触发行数/耗时，`--developer-view` 输出 started/completed 进度事件。真实入口端到端：Windows 原生 CLI 直跑真实 config-real 配置 + 真实单位贷款机构审核副本，377 条规则结果完成，状态计数（PASSED 80/TRIGGERED 134/STRUCTURE_ERROR 162/RETRIEVAL_FAILED 1）与既有对拍一致；CLI 产物与 Python `AuditEngine.run(output_dir=...)` 基线文件级对拍 0 差异（审核副本集合与字节、审核结果.xlsx 全字段、总账状态计数/运行状态/清单含路径 basename 归一）。已知边界：单绑定本期源（上期/外部源与 V_EXT 取数待迁移）、无输入哈希降级与结构快照、`--business` 过滤在配置无多业务分目录形态下与 Python businesses 参数等价。S5-F01 整体仍未验证（真实入口已接通但配置校验入口/V_EXT/上期未接入 CLI 链路）。

### S5-F01 结果输出切片（真实运行文件级对拍）

将 `outputs.write_run_outputs` 的单业务形态移植为 `audit-core/src/s5_outputs.rs`：运行目录结构（run_id/审核副本/审核结果.xlsx/运行总账.json）、副本复制与 sha256 三方清单（before/after/copy、matches_snapshot）、审核结果工作簿（TRIGGERED+四类失败行、`_error_type` 严重度分类含自定义级别保留、`_location` 证据优先/合并单元格与 METRIC_AMBIGUOUS cells 候选、`_description` 候选单元格注记、副本相对超链接（撇号转义）、precision 数字格式、`_style_table` 视觉：微软雅黑 10、表头白字深蓝底居中换行行高 30、正文 top 换行行高 54、冻结 A2、筛选、12 列宽、横向 A3 fitToPage 标题行重复、J/K/L 右对齐无换行含空格、C 列文本格式、恒建空“条件格式触发”表）、运行总账稳定字段与 output_files。求值/取数侧同步修正三处真实对拍暴露的保真缺陷：`with_numeric_mode` 改按规则模式（Python `replace(value, value=raw)`）收敛 Decimal→round15 f64 且 blank 保持空值种类（此前 blank→0 与漏转换均由文件级对拍暴露）；取数失败 evidence 补 `retrieve` 异常分支的 function/sheet setdefault 富化（镜像 `SourceRepository.retrieve` except 分支）；服务错误分类直取取数错误码。Windows 原生文件级对拍：真实 config-real 配置 + 真实单位贷款审核副本，整业务 377 规则运行后双侧输出（Python openpyxl 引擎为样式基线）——审核副本字节一致；审核结果.xlsx 表清单、值/类型、数字格式、超链接（basename#sheet!cell）、列宽（分组 col 按 min/max 展开）、冻结/筛选/隐藏网格线/页面设置、行高、表头字体填充对齐、正文与 J/K/L 对齐格式差异 0；总账 run_id/status_counts/run_status 与清单（角色/顺序/snapshot 标记/三哈希）一致。颜色按 6 位 RGB 归一、fitToWidth=1 与属性省略视为等价（rust_xlsxwriter/openpyxl 序列化表示差，Excel 语义一致，已记录）。S5-F01 整体仍未验证。证据见 `probe/reports/s5_run_outputs_windows_real_20261025.json`。

## 历史 Rust 前端接入快照

- 源项目 `core/frontend/web/index.html` 和 `core/frontend/assets/app-icon.{png,ico}` 已原样复制到 `frontend/`；HTML 与源文件 SHA-256 一致。源项目只读，没有修改。
- Windows 可直接启动新增的 `open-data-audit-gui.exe`（无参数默认端口 8750，可用 `--port`/`--config` 覆盖）；它复用原 GUI 页面和同一 HTTP bridge，仅监听 `127.0.0.1`，启动后尝试打开系统默认浏览器。原 `open-data-audit serve` 也复用该自动打开行为；打开失败时打印手动访问地址。此入口已通过 Windows HTTP 冒烟，状态仍为“已实现，待 Windows 真实验收”。
- bridge 在响应时注入运行时接口，保持 HTML 文件本身不变。当前页面启动后导航至已接入的 S2。
- 当前连接 S2 配置检查、路径选择/清除、输出目录固定/跟随、后台审核、进度和结果回传；S3-F02 说明导出与 S3-F04 CSV/XLSX/XLS 本期数值核对；另有 S1 只读模板公式检查 API。S3-F04 选择器和 API 已接 XLSX/XLS，但 Web API 合成端到端不等于 GUI 人工验收。
- S1-F01 的 `choose_file("template")` 与 `inspect_template_formulas` 提供只读模板公式检查 API；Windows 原生真实 HTTP 调用记录见 `probe/reports/s1_template_inspection_web_api_windows_real_20261024.json`，不代表 S1 完整审核 GUI 已迁移。
- 其余未迁移接口返回 HTTP 501；取消和设置持久化尚未接入。页面仍展示未迁移系统入口，只有上方逐项记录的 Rust API/业务链路已接通；任何局部 API 均不代表对应 S1/S3/S4/S5 整项已验收。
- 合成 S2 输入经 HTTP bridge 完整运行并生成三表 XLSX；该结果不替代真实业务模板、Windows 原生、麒麟/UOS、glibc 2.28 与 deb 验收。

## 2026-10-25 GLM 阶段 Codex 独立审计复核

- main 公共修复：`9f0cacf` 已同步为 `fc6b121`，逐表 OOXML 快照按工作表名比较；交接报告中的 `d2da5a1` 不存在，实际修复为 `d2da501866a4aca7d2a4132b5c2ed0d2116d4f42`，已同步为 `a5e0233`。Worker 越界样式索引现在会被拒绝，未修复其样式映射。
- S1-F01 历史写回：真实 config-real 样本 5 表、9 行历史；`UmyaXlsxBackend::apply` 通过。同步 main 的空 `pageSetUpPr` 默认归一修复 `cc64e15` 后，完整计划继续失败于关系快照：工作簿新增 `sharedStrings`，管理历史表原有 9 个超链接关系消失（原 Python 会删除并重建该表），批注/VML 关系 ID 重编号。不得全局忽略关系差异；共享请求 #7 的页面设置问题已解决，剩余差异登记于 `shared-capability-requests/20261025_umya_movesheet_relationship_validation.md`。源 SHA-256 不变；历史写回仍 Unverified。
- S1-F01 公式复制：worker 样式表守卫已同步且 worker 定向回归 10 passed；真实 worker 输出仍可能因 `cellXfs` 引用越界被拒绝，Umya 路线仍有 1,036 项真实对拍差异；没有可用生产编辑引擎，整项未验收。
- S5-F01 V_PRE CLI 切片：`s5-audit-run` 新增可选 `--previous-source`、`--previous-source-id`、`--previous-period`，审核结果总账登记并校验上期副本哈希；未提供上期时保留 Python `SOURCE_NOT_SELECTED` 失败结果。合成 Windows Python/Rust 对拍核对结果表值、状态、超链接目标和当前/上期副本清单，差异 0，见 `probe/reports/s5_vpre_cli_synthetic_parity_20261025.json`。仅覆盖单个上期来源的无契约绑定路径；真实上期样本、数据源契约校验、多来源选择、V_EXT 动态布局/时间序列（时间序列后续仅有合成验证）、结构快照、数据质量告警及全配置/多业务验收仍缺，S5-F01 整体未验证。
- S5-F01：此前 Windows CLI 与 Python `AuditEngine.run` 377 条结果及输出文件对拍差异 0 的真实证据仍仅覆盖单位贷款单业务切片。新增业务筛选拒绝未知/无启用规则业务，输出 writer 在写出前后验证源与配置快照，防止空业务或输入变化生成完成态输出；定向测试通过、CLI check 通过。
- S5-F01 真实布局模板输出：修复实际模板绑定审核只把模板用于取数、却未纳入 Python `layout` 审核副本和 SHA-256 清单的差异。Rust 现在仅复制规则取数中实际选中的模板，并用模板读取时快照校验输出。Windows 原生 Python/Rust 真实模板输出对拍 377 条结果、工作簿、审核副本及清单差异 0；显式模板路径和真实模板目录自动匹配分别见 `probe/reports/s5_outputs_real_template_fixed_windows_native_20261026.json`、`probe/reports/s5_outputs_real_template_dir_windows_native_20261026.json`。两种 `s5-audit-run` CLI 调用均 COMPLETED（PASSED 352、TRIGGERED 24、RETRIEVAL_FAILED 1），副本为 current/layout/configuration 且三项快照均匹配；目录匹配选中项与 Python 的模板编号和文件名相同。显式路径运行前后源表、配置和模板哈希不变。此证据只覆盖单位贷款单业务、Windows；S5-F01 整体仍未验证，Audit #8 共享能力问题仍未解决。
- 平台：本轮仅完成 Windows 原生编译/测试及 Windows 隔离 openpyxl worker 回归；WSL Python、Linux 目标、glibc 2.28、deb、麒麟 V10/UOS V20 均未验证。

### S5-F01 V_EXT 单来源配置契约 CLI 切片（Windows 合成对拍）

`s5-audit-run` 新增可选 `--external-source`/`--external-source-id`，在现有配置的外部来源契约下绑定单个来源，验证机构/期间身份、映射工作表与字段唯一性、报表项目和指标列唯一性；拒绝公式/错误值，并把外部来源副本及 SHA-256 纳入审核总账。后续独立普通静态布局切片已为无契约来源接入动态识别；V_EXT 时间序列在该契约切片时尚未接入（后续实现见下节），无契约静态布局告警已由后续质量告警切片接入结果表和总账（见 `probe/reports/s5_vext_dynamic_layout_quality_synthetic_parity_windows_native_20261025.json`）。Windows 原生合成 XLSX 与 Python `AuditEngine.run` CLI 对拍：审核结果表值/状态、超链接目标及本期/外部副本清单哈希差异 0，见 `probe/reports/s5_vext_contract_cli_synthetic_parity_20261025.json`。对拍发现并修复 Decimal 精确除法保留小数位差异；当前回归断言覆盖 `15000000.00 / 10000 -> 1500.00` 和零差额显示。此契约切片仅有合成证据，未验收真实外部文件；多来源、时间序列和目标平台仍未验收。普通动态布局及其质量告警已有后续合成对拍，但不代表真实外部样本或 S5-F01 整体通过。Audit 共享请求 #8 未变更、仍未解决。

### S5-F01 V_EXT 无契约普通静态布局切片（Windows 合成对拍）

新增 `s5_external_layout` 并接入 `s5-audit-run` 的 V_EXT 取数：无 `数据源契约` 时，从“机构参照表/机构参照”识别机构代码、名称、报表项目列，按代码优先/名称回退定位机构；再从普通静态报表中推断“行名称/报表项目”列、最近表头行、指标列和唯一报表项目行。支持 `sheet=` 限定重复指标所在表、单位换算和映射证据，且按本次绑定来源缓存布局。首版曾对未迁移时间序列格式、布局告警和空白/非法数值失败关闭；后续质量告警切片已按 Python 行为将布局发现告警写入 `data_quality_warnings`，并将空白、非数值及动态公式文本按 0 参与审核，同时写入“数据质量提示”工作表和 `运行总账.json`；整数数值正常换算。本段记录时序格式尚未迁移；后续实现见下节。

Windows 原生 Python/Rust 合成 CLI 对拍：同一临时 XLSX、真实 `s5-audit-run` Rust 入口和原 Python `AuditEngine.run` 动态布局路径，审核状态、当前/外部/差值数值及结果描述差异 0；Rust 输出单元格保留两位小数显示格式 `0.00`。报告：`probe/reports/s5_vext_dynamic_layout_synthetic_parity_windows_native_20261025.json`。后续质量告警切片的定向 Rust CLI 集成测试 7 项通过，覆盖普通布局取数、布局/数值告警结果表与总账、整数数值、当时的时间序列拒绝行为（本轮已由时序匹配用例替换）及既有 V_PRE/契约 V_EXT 回归；Python 对应基线 6 项通过。Windows 原生同一临时合成输入 Python/Rust CLI 对拍空白、非数值与整数金额 3 例，审核状态/取值、告警总账字段、质量表数据行及页面布局差异 0，见 `probe/reports/s5_vext_dynamic_layout_quality_synthetic_parity_windows_native_20261025.json`。未使用真实外部文件；时序子切片现已完成合成验证，但多外部来源、真实外部文件、外部文件地址扫描、Linux/deb 与目标系统验收仍缺，S5-F01 整体未验证。Audit 共享能力请求 #8 仍未解决。


### S5-F01 V_EXT 按报告期匹配时序布局切片（Windows 原生合成对拍）

Rust `s5_external_layout` 识别首个同时包含唯一“数据日期”和“行名称”的前 20 行表头，读取行标签、按审核报告期精确筛选目标记录，并排除日期/报表代码等元数据列；格式化 Excel 日期经共享 `read_value_columns_with_iso_dates` 读取，公式、空白或非法日期不参与匹配并产生与 Python 一致的数据质量提示。存在 `数据源契约` 时仅在识别到该时序表后转入同一动态布局路径，普通静态契约路径保持原样。布局告警的机构字段按 Python 基线留空。

Windows 原生同一组合成 XLSX（含 prior-date 行、Excel 日期格式单元格的本期行及非法日期行）分别执行 Python `AuditEngine.run` 与 `open-data-audit s5-audit-run`；配置含外部来源契约，审核状态、当前/外部/差值、说明、告警数量/字段/结果表行一致，输入 SHA-256 前后不变。详细记录见 `probe/reports/s5_vext_time_series_synthetic_parity_windows_native_20261025.json`。Windows Python `test_s5_integration.py` 相关基线 15 passed；Rust `s5_previous_source_cli` 9 passed（包含无当前日期失败和配置契约下时序路由）。尚无真实外部时序工作簿验收；S5-F01 整体仍未验证，Audit 共享请求 #8 仍未解决，Linux/deb、glibc 2.28、麒麟 V10/UOS V20 未执行。


### S5-F01 实际输出说明数值格式修复（Windows 真实单业务对拍）

修复 `s5_service::python_g15` 对 binary64 结果说明的格式化：严格按 Python `format(value, ".15g")` 的 15 位有效数字、科学记数阈值、进位和负零语义生成描述，避免固定小数位暴露二进制尾差。比较器改用精确 Decimal 归一化，避免 float 比较掩盖数值差异；支持传入真实业务身份并直接运行 `open-data-audit s5-audit-run`，对比前后检查配置、源表和模板输入未变化。

Windows 原生真实单位贷款审核副本、config-real 配置及模板目录对拍：Python `AuditEngine.run` 与 Rust 服务结果 377/377 行一致，差异 0；两侧状态计数一致（PASSED 80、TRIGGERED 296、RETRIEVAL_FAILED 1）。真实产品 CLI 与 Python 完整输出对比差异 0，输入未变化。证据：`probe/reports/s5_real_service_parity_windows_native_20261026.json`、`probe/reports/s5_real_cli_full_output_parity_windows_native_20261026.json`。

该证据限于 Windows、单位贷款、单机构、当前来源和所选模板目录；条件格式/数据质量表为空，V_PRE/V_EXT、结构快照与哈希完整性降级未覆盖。S5-F01 整体仍未验收，生产 capability 仍为 unverified。Audit 共享能力请求 #8 保持未解决。

### S5-F01 多本期来源绑定切片（Windows 合成对拍）

`s5-audit-run` 支持重复 `--current-source SOURCE_ID=BUSINESS=FILE` 与主来源 `--source-business`；`V_CUR` 按显式逻辑来源 ID 选择，未指定来源时按 Python `SourceRepository._binding` 的业务绑定唯一选择或返回歧义/缺失错误。所有本期来源共用本次期间和机构身份门禁，附加来源进入审核副本和哈希总账。Python/Rust 合成 CLI 对拍同时覆盖显式 ID 和业务回退两条规则，结果行 2/2、输出工作簿/副本/清单差异 0，输入未变；超链接比较将 URI fragment 与 OOXML `location` 属性归一为同一目标。证据见 `probe/reports/s5_multiple_current_sources_synthetic_parity_windows_native_20261026_v3.json`。

另用两条同业务来源覆盖歧义回退：Python 与 Rust 均生成 `CONFIG_ERROR=1`、`TRIGGERED=1`，结果工作簿和审核副本/清单差异 0，输入未变；见 `probe/reports/s5_multiple_current_sources_ambiguous_synthetic_parity_windows_native_20261026.json`。该切片仍只有 Windows 合成样本证据；真实多来源业务、业务无匹配来源、跨平台及完整 S5-F01 验收未完成，生产 capability 仍为 unverified。Audit 共享能力请求 #8 保持未解决。

### S5-F01 多外部来源绑定切片（Windows 合成 Python/Rust 对拍）

`S5Retrieval` 与 `s5-audit-run` 支持多外部来源。旧 `--external-source`/`--external-source-id` 保持原默认来源行为；新增可重复 `--external-source-binding SOURCE_ID=BUSINESS=FILE` 与可选 `--default-external-source-id`。V_EXT 选择遵循 Python `_binding("external", ...)`：调用显式 `source=` 优先，其次默认来源，再按规则业务唯一匹配；仅一个来源时允许单来源回退，多个匹配返回 `SOURCE_AMBIGUOUS`，无匹配返回 `SOURCE_NOT_SELECTED`。各来源的动态布局、时序识别缓存和文件哈希独立，所有绑定来源进入审核副本与总账。

Windows 原生 `s5-audit-run` 与原 Python `AuditEngine.run` 对拍覆盖业务回退、显式 `source=`、同业务候选歧义及未匹配来源。三组完整结果/工作簿/副本/清单差异均为 0，输入未变；歧义组状态计数为 `CONFIG_ERROR=1`、`RETRIEVAL_FAILED=1`。报告：`probe/reports/s5_multiple_external_sources_business_fallback_synthetic_parity_windows_native_20261026.json`、`probe/reports/s5_multiple_external_sources_explicit_synthetic_parity_windows_native_20261026.json`、`probe/reports/s5_multiple_external_sources_ambiguous_synthetic_parity_windows_native_20261026.json`。旧单外部来源契约、动态时序布局和多本期来源回归均通过。

多外部来源选择本身目前仍只有 Windows 合成对拍。另用 config-real 真实审核配置、2026-07-31 批次的一本单位贷款审核副本、真实大集中外部工作簿及临时复制的模板目录，完成单外部来源端到端验证：`s5-config-check --scope current_business` 为 valid/complete，377 条启用规则通过，静态模板地址 532/532、外部指标表头 1/1；`s5-audit-run` CLI 与 Python 3.12.10 `AuditEngine.run` 完整输出对拍差异 0，状态计数均为 PASSED 373、TRIGGERED 3、STRUCTURE_ERROR 1，输入未变且报告不保存原始值。报告：`probe/reports/s5_vext_real_external_single_source_parity_windows_native_20261026.json`。

同一真实输入另显式将 `EXT_REAL` 设为默认外部来源，验证 `--default-external-source-id` 路径：Python/Rust 完整输出差异仍为 0，状态计数相同，输入未变且未持久化原始值；报告中的 `default_external_source` 为 `EXT_REAL`。证据见 `probe/reports/s5_vext_real_external_default_selection_parity_windows_native_20261026.json`。

个人贷款业务也使用该真实配置、另一份 2026-07-31 当前工作簿、真实个人贷款模板及同一外部工作簿，显式验证默认来源：配置校验 valid/complete，356 条启用规则通过，模板地址 475/475、外部表头 5/5；Python/Rust `s5-audit-run` 完整输出差异 0，状态计数均为 PASSED 339、TRIGGERED 12、STRUCTURE_ERROR 5，输入未变且报告不保存原始值。证据见 `probe/reports/s5_vext_personal_loan_real_external_default_parity_windows_native_20261026.json`。

同一批次将单位贷款、个人贷款两个真实当前工作簿和 `EXT_REAL` 默认外部源合并运行时，首次发现 Rust 总账中的两个布局模板条目按来源绑定顺序排列，与 Python 按规则首次使用顺序不同；业务结果状态计数一致，结果表单元格值差异为 0，差异仅在 12 个模板清单字段。`S5Retrieval` 现在按首次选用布局模板的顺序累积清单，并按路径去重。修复后同一真实组合的 Python/Rust 完整 CLI 输出差异为 0，状态计数均为 PASSED 719、TRIGGERED 8、STRUCTURE_ERROR 6，输入未变且报告不保存原始值；见 `probe/reports/s5_real_dual_business_current_external_default_parity_windows_native_20261026.json`。

上述真实组合覆盖单位贷款和个人贷款各一本当前工作簿，但共用一个物理外部工作簿；多物理外部来源真实验收、条件格式/数据质量非空路径、输入哈希变化降级、Linux/deb、glibc 2.28、麒麟 V10/UOS V20 与完整 S5-F01 仍未验收，生产 capability 仍为 unverified。结构快照已有后续独立真实对拍，见下节。Audit 共享能力请求 #8 仍未解决。

### S5-F01 来源结构快照切片（真实 Python 基线跨语言读取）

`s5_schema_snapshots` 将 Python `SourceRepository` 的路径摘要、HMAC 令牌、结构/位置对比和原子持久化接入实际 `s5-audit-run` 取数及输出总账。无命名结构区域的回退范围现在使用共享 `excel_reader::read_xlsx_worksheet_extents` 的 OOXML 单元格边界，包含带样式的空白单元格，匹配 Python `openpyxl.max_row/max_column`。Windows 合成单测覆盖 HMAC RFC 4231 向量、基线/位置变化/逻辑变化、输入变化不持久化、非对象 JSON 快照错误映射、样式空白扩展范围和总账事件写入。

Windows 原生真实对拍使用 config-real 的 5 逐笔审核配置和 2026-07-31 单位贷款当前工作簿；Python `AuditEngine` 生成 4 个隔离结构快照后，Rust 实际 `s5-audit-run` CLI 读取这些 Python 基线，4/4 均返回 `UNCHANGED`，结构事件差异 0；377 条审核结果状态计数一致（370 PASSED、6 TRIGGERED、1 RETRIEVAL_FAILED），输入哈希前后相同，报告不保存原始单元格值或本机路径。证据见 `probe/reports/s5_structure_snapshot_real_parity_windows_native_20261026.json`。验证范围仅含 Windows、单位贷款、一个本期来源及其实际回退布局；未验证 V_PRE/V_EXT、哈希并发变化、目标 Linux/deb、麒麟 V10/UOS V20，也不代表 S5-F01 整体完成；生产 capability 仍为 `unverified`，共享请求 #8 仍未解决。

### 2026-10-27 Codex 增量：S1 缓存安全门控与 S3-F01 输出切片

S1-F01 的 `s1-audit-directory` 现在要求显式选择 sidecar 强制重算或 `--allow-saved-formula-cache-diagnostics`；只读旧缓存会在 JSON 中标记 `completed_with_diagnostics` / `saved_cache_diagnostic_only`，并明确本命令未执行公式复制。公式副本清单也拒绝把源文件本身当作审核副本。Windows 原生回归：S1 单测 60 通过、5 忽略；`s1_audit_directory_cli` 6 通过。既有真实 7 文件、322 条 Issue 的缓存副本对拍仍只覆盖审核副本之后的链路；本轮未生成真实业务输出，公式复制与真实重算未闭合。sidecar 的 Windows COM 生命周期失败和 ImportRange 样式越界能力仍受 main 处理。

S3-F01 新增旧版比较结果 XLSX 的 Rust 写入切片：22 列、旧版工作表命名、数值格式、行/备注警戒色、冻结/筛选、写后读回和无覆盖发布。Windows 合成 writer 回归通过 OOXML 结构/样式索引、单元格读回与覆盖保护检查；S3-F01 单测共 11 通过、1 个真实配置测试忽略。该输出函数尚未接到 CLI/GUI，也没有真实两期数据与 Python 输出对拍；S3-F01 保持“迁移中”。

### 2026-10-27 Codex 增量：S3-F01 CLI 入口与最小冒烟

新增 `open-data-audit s3-execution-compare`，复用既有 S3-F01 Rust 服务，接收重复的本期/上期 CSV 或工作簿路径、3.0/3.1 配置与新的输出路径；普通日志和 `--developer-view` 分别记录步骤、引擎、规则计数及耗时。Windows 合成端到端冒烟由 CLI 实际生成 1 行结果 XLSX，检查开发者进度事件、输入字节未变，并确认同路径重跑拒绝覆盖且保留旧输出。Rust S3-F01 定向单测 11 passed；旧项目只读真实 3.0/3.1 配置加载测试通过（13 档环比、1711 指标参照、5082 动作、2844 表达式规则）。目前没有可用的真实两期 S3 数据，尚未与 Python 真实结果对拍，GUI 未接入；S3-F01 仍为“迁移中”。

### 2026-10-27 Codex 增量：S3-F01 GUI API 接线与合成冒烟

原有 GUI 的两期数据、3.0/3.1 配置、输出目录及表达式模式控件已接入 start_central_comparison：调用复用既有 Rust S3-F01 服务并在后台执行；路径和格式错误明确返回，输入/配置不修改，输出使用带时间戳的新文件名。合成 GUI API 测试生成 1 行 XLSX，验证结果路径、双视图引擎日志、任务结束状态和四个输入/配置字节未变；临时输出随后删除。Windows web_bridge::tests::s3_execution_compare 2 passed；Audit CLI cargo check、实际 cargo build 和生成二进制的 serve --help 通过。缺少真实 S3 两期业务文件，未做 Python/Rust 真实结果对拍或人工浏览器验收；不代表 S3-F01 整体验收通过。
### 2026-10-27 Codex 增量：S1-F02 原 GUI 入口接线

迁移状态：迁移中。Rust Web bridge 已接通原 GUI 的 `start_dag_flow("dag:汇总校验结果说明", ...)`，并补上 S1 源目录/模板目录选择、模板文件选择、历史配置选择、源文件识别与勾选、后台汇总、结果路径和双视图日志。执行复用 `s1_region_summary` Rust 汇总服务，默认输出到源目录下 `执行结果/校验结果与报送说明汇总_时间戳.xlsx`；输出不覆盖已有文件，历史配置和源文件只读。递归发现排除 `执行结果` 目录，避免重复汇总历史输出。

Windows 合成 GUI API 冒烟：从原 GUI API 选择目录、扫描源文件、选择模板/历史配置并执行，生成可读 XLSX；源文件和历史配置副本运行前后字节一致。定向测试 2 passed；既有 `s1_region_summary_cli` 对拍回归 5 passed；`open-data-audit-gui` Windows 构建和 `--help` 通过。测试只证明合成入口与已有合成核心，不代表真实历史命中、结构异常跳过、真实输出契约或 Python/Rust 真实业务对拍通过。S1-F02 整体仍为迁移中。

S1-F01 旁路复测仍受阻：合成 Excel COM 计算返回 `RPC_E_DISCONNECTED`，原因在 sidecar 已保存关闭工作簿后 finally 再次关闭；原输入保持不变且失败输出未发布。Audit 未修改 sidecar/共享 Excel-Core；已向 main 提供复现，需 main 的 COM 生命周期修复后再重测。S1-F01 真实公式复制与强制重算链仍未闭合。

### 2026-10-27 Codex 增量：S1-F02 汇总结构预检

Rust 汇总入口已接入模板“表头区域”结构比对：正确结构的文件继续汇总，错配文件跳过并在普通日志中说明文件名/原因；缺少模板工作表时沿用旧 Python 部分继续语义，fallback-only 模板不增加固定结构门槛。Windows 合成 CLI 回归 6 passed，覆盖匹配、错配跳过及部分继续；GUI S1-F02 API 回归 2 passed，Windows GUI 构建通过。此次无适用真实 S1-F02 .xlsx 源样本，未对新增结构预检做真实 Python/Rust 对拍；已有 2026-10-24/25 真实个人贷款汇总报告未覆盖该新增步骤，输出格式尚有既有列宽差异。S1-F02 仍为迁移中。

### 2026-10-27 Codex 增量：S1-F01 显式重算副本合流复验

`s1-audit-directory` 的显式 sidecar 路径现将公式审核项与本轮实际重算副本配对，修复原清单仍指向重算前副本而导致结果合流失败的问题。Windows Excel COM 合成端到端测试使用已有触发条件格式的合成工作簿，验证计算后的完整公式文本、公式审核项与条件格式审核项同入本期结果工作簿和历史结果工作簿，源文件和清单中的原审核副本字节不变；输出只写测试临时目录，测试结束后删除。`s1_recalculation::tests` 5 passed、1 ignored；`s1_audit_directory_cli` 6 passed、1 ignored；另显式运行 ignored Excel COM 端到端测试 1 passed。Rust CLI 构建及 `s1-audit-directory --help` 通过。

该验证仍仅覆盖合成模板/样本和“清单已提供审核副本”之后的链路。CLI 未在本次执行中自动把模板校验区域复制进源文件独立副本；真实配置样本、Python 完整输出逐条对拍、公式副本引擎生产能力和工作簿对象级保真仍未验证，S1-F01 整体仍为未验证。Windows 全仓 fmt 检查因其他现存文件的格式差异未通过；本次涉及的 Rust 文件单独 rustfmt 检查通过。

### 2026-10-27 Codex 增量：S1-F01 Excel COM 重算复验

main 的 `9da7878` COM 关闭生命周期修复不在 audit-dev 提交祖先中，但同等脚本变更已存在于当前 sidecar 文件，因此安全 cherry-pick 判定为空并跳过。Windows 原生 `cargo build --offline --locked -p lo-convert` 通过；显式运行 `explicit_sidecar_recalculates_synthetic_formula_copy_without_changing_input` 1 passed，合成公式 `C1=SUM(A1:B1)` 重算读回为 5，输入字节不变。该结果验证 COM 重算切片，不代表真实 S1 模板公式复制、样式保真或完整业务入口通过；S1-F01 整体仍未验证。

### 2026-10-27 Codex 增量：S1-F02 真实历史配置对拍与 S1-F01 公式复制候选复验

S1-F02 Windows 原生真实样本使用既有历史审核配置：Rust `s1-region-summary` 汇总 7 个源工作簿、27 张源工作表，共 41 行；与旧版 Python 3.12.10 / Microsoft Excel COM 输出比较，工作表名/顺序、尺寸、值/类型、数字格式、字体、对齐、表头、筛选和冻结差异为 0。自动列宽最大差 1.57421875 字符，不能记为完整格式通过。历史配置及源文件均只读，输出限临时目录。定向 Rust 单测 5/5、CLI 测试 6/6 通过；Rust 原 GUI HTTP API 使用同批真实样本再次运行成功，返回 41 行并生成可读工作簿，与 Python/Excel 基线差异仍仅为列宽；尚未做浏览器人工操作和 Linux/deb 验收。S1-F02 状态为“已实现，待 Windows 真实验收”。

S1-F01 使用当前 config-real 映射、单位贷款模板和真实只读工作簿执行 openpyxl worker 候选：Rust 生成 4 个 ImportRange、55 个 SetComment、1 个计算属性操作（共 60 项），worker 写出临时候选；Rust 样式读回在第 2 张表 H10 报 `imported_style_semantic_mismatch`，因此候选不可发布、能力仍 unverified，未执行公式重算。H10 在源/模板/output 的 XML 中均无显式单元格节点；旧 Python `copy_formula_ranges_from_template` 仅 `source.has_style` 为真才赋样式，该格实测为 false，说明该项可能是 validator 对隐式行/列样式的误判，仍需旧版输出对拍确认。此前报告的 CJ2 样式索引越界是真实 worker 缺陷；main 的 `7fce332` 已同步并增加越界 fail-closed 校验，但没有修复根因。S1-F01 整体保持“受阻”。

### 2026-10-27 Codex 增量：S1-F01 GUI 未接通反馈

Rust 原 GUI 将 S1-F01 按钮明确标为“尚未接通”；确认执行后的 API 异步错误现在显示为 toast 并刷新状态，避免服务端 `NotImplemented` 静默失败。Windows 原生 GUI 构建通过，浏览器实际点击确认显示具体未接通错误；无输入时没有启动文件处理。S1-F01 业务链继续受阻，未改变迁移验收状态。S1-F02 定向 GUI 工作流回归 2 passed。

### 2026-10-27 Codex 增量：S5 GUI 运行适配与 V2 13 列结果协议

S5 Rust GUI 适配器通过真实只读个人贷款样本运行：356 条规则、3 条触发、0 条失败，耗时 6,428ms；生成的结果簿可由共享 reader 重新打开，真实运行测试逐项断言了 `核查表校验结果` 的 13 个字段，源工作簿字节未变化。输出 writer 现按 V2 字段写入工作簿名、工作表、单元格、错误类型、校验类型/指标、说明、当前/对比/差值、规则编号及两个保留空字段；文件仍为 `审核结果.xlsx`，触发和失败行筛选、运行账本及审核副本逻辑保留。定向 writer 回归 8/8 通过，`open-data-audit-gui` Windows cargo check 通过。

该 13 列协议遵循 V2 最终需求；旧 Python S5 README 仍将同名配置 Sheet 视作空历史表，并使用不同结果布局，因此本轮不宣称与旧 Python S5 输出格式完全对拍。历史说明和审核意见目前没有 S5 配置来源，保持空白；条件格式实际触发提取、浏览器人工操作及 Linux/deb 未验证。S5-F01 整体仍未验证。真实运行结果输出在隔离临时目录并由成功测试清除；前一次断言失败的隔离临时目录仍待清理。

### 2026-10-27 Codex 增量：S1-F01 ImportRange 真实复制与 Windows COM 复验

同步 main 的 `c1a9e38` 共享 openpyxl Worker ImportRange 语义修复，Audit 分支提交为 `4297315` 并已推送 `origin/audit-dev`。main 提供的定向回归为 Python Worker 16/16、Excel-Core Rust 单测 45/45、Rust→Worker 批处理集成 1/1。Audit Windows 原生真实模板候选执行 60 项（4 个 ImportRange、55 条批注、计算属性），写出临时候选；结构/值/公式未见越界引用，5,487 个复制单元格的值/类型和公式文本差异均为 0，但 Rust 读回仍在第 4 张表 I6 报 `imported_style_semantic_mismatch`，整体报告 `comparison_incomplete`。旧 Python 对照输出存在 1,012 个无效字体组件引用；候选与 Python 对照合计样式差异 1,063，不能据此验收。Rust 候选仍为 Unverified，不接入生产路由。机器报告：`probe/reports/s1_f01_openpyxl_semantic_parity_windows_native_20261027_v7_excel_com.json`。

对拍脚本的 Windows 重算分支已改用共享 sidecar 的 Excel COM，为 Python 与 Rust 临时副本分别记录成功/失败类别；Linux 的 LibreOffice 分支未改。合成公式复制与 Excel COM 缓存对拍通过（4 组预期缓存一致、0 差异），状态为 `candidate_match_unverified`；`py -3.12 -m py_compile` 通过。真实样本 COM 复验中，旧 Python 写出簿在 `Workbooks.Open` 失败，Rust 候选被 Excel 打开计算后因公式文本变化遭 sidecar 拒绝，未发布重算输出；输入未改。结构读回 I6 与 openpyxl 实际输出仍有解释差异，及真实公式文本变化已向 main 提交复现，需共享 validator/sidecar 核实。S1-F01 公式复制、Office 重算、两类问题合并输出和 GUI 入口仍未闭合，整体保持“受阻”；未把合成对拍或候选落盘计为验收通过。

### 2026-10-27 Codex 增量：收敛迁移对拍标准与可人工测试样例

按用户本轮要求，普通字体/颜色/填充/边框/对齐、样式编号、行列尺寸和打印外观不再阻塞迁移；不重复执行全量样式对拍。继续检查值与类型、公式文本和引用、必要工作表/区域关系、审核规则及输出文件可读性。

S3-F01、S5-F01 的 Windows CLI 现已有最小人工样例，状态为“已实现，待 Windows 真实验收”。输入、配置、独立预期和实际 XLSX 输出位于 audit-core/manual_acceptance/20261027/。S3 正常/异常分别输出 0/1 条触发动作；S5 正常/异常分别输出 PASSED 1 / TRIGGERED 1、失败 0。四个结果工作簿均由共享 Rust Calamine reader 成功重开。GUI 目标编译通过，但本轮未作浏览器人工验收。

S1-F01 仍未闭合/受阻：真实候选的 5,487 个复制格值/类型和公式文本对 Python 基线差异为 0；普通样式差异不再计为阻塞。范围外报告仍有 216 个相对输入的值/类型差异（Python/Rust 候选相互差异为 0），空字符串与空白单元格语义尚未确认。真实 Excel COM 重算报告为失败（Python 打开失败，Rust 候选公式文本变化），没有可交付的重算结果。未把内容疑点或重算失败归类为外观差异，也未重跑全量格式对拍。Linux/deb 未验证。

### 2026-10-27 Codex 增量：S1-F01 按内容验收

S1 公式复制计划现使用 `RangeStyleMode::None`，继续复制值、公式文本并单独写入模板批注；普通样式不再要求共享编辑器读回一致。`compare_s1_formula_copy_candidate.py` 不再读取/比较样式组件、样式索引、行列尺寸、冻结窗格、页面设置或整包字节差异；候选状态只由业务内容、公式/缓存、批注/链接、定义名称、合并、验证/条件格式规则、工作表状态及输入未变等必要检查决定。既有真实报告仍是唯一真实候选证据，5,487 个范围格内容/公式对拍差异为 0；未重跑真实样本或样式比较。

本轮 S1 计划定向单测 7/7、Audit `cargo check` 和目标 Rust 文件 `rustfmt --check` 通过。新的合成 worker 冒烟未完成：Cargo 在重新链接时报告运行中的 `open-data-audit-gui.exe` 文件占用，故本轮没有新的 worker 端到端通过证据。S1-F01 仍未接入 GUI/生产路由；真实 COM 重算及 Python/Rust 完整输出仍未验证/受阻。

### 2026-10-27 Codex 增量：S5 条件格式接入与 Windows 手测版本

条件格式入口补充了含真实 `conditionalFormatting/cfRule` 公式的 S5 合成 XLSX 手测样例：正常值 120 对 `C4>120` 不触发，异常值 125 在 `贷款表!C4` 触发。使用 Windows Rust `s5-audit-run` 正常/异常均完成；运行总账分别记录触发 0/1、规则 1/1 supported、错误与不支持均 0，业务规则结果分别 PASSED 1 与 TRIGGERED 1。只证明此合成公式与 CLI 输出切片，不推进 S5 整体验收状态。

S5 条件格式 Rust 求值已接到 `s5-audit-run` CLI 与现有 GUI S5 适配器；真实触发、UNSUPPORTED、ERROR 均写入“条件格式触发”页及运行总账，存在无法判断项时状态为 `COMPLETED_WITH_CONDITIONAL_FORMAT_LIMITATIONS`，不会静默视作未触发或通过。Windows 定向测试：条件规则评估 1 passed，结果写表/不隐藏限制 1 passed；`cargo build --offline --locked -p audit-core --bin open-data-audit --bin open-data-audit-gui` 通过，产物在 `F:\RustBuild\audit_manual_ready_20261027\debug\`。

当前二进制实际运行现有合成样例：S3-F01 正常/异常均 COMPLETED、结果各 1 行、触发动作 0/1；S5-F01 正常/异常均 COMPLETED 且失败 0，状态分别 PASSED 1 与 TRIGGERED 1，条件格式规则数均为 0。最新实际输出在 `manual_acceptance/20261027/outputs/codex_ready_20261027_114631/`，独立预期在 `expected-results.txt`。GUI 服务在 127.0.0.1:8751 返回 HTTP 200，页面含 bridge 与 S5 入口；未代替用户点击验收。S3-F01/S5-F01 仅就列出的模拟案例待 Windows 人工验收，不代表整项或真实业务验收通过。

S1-F01 普通外观样式差异按用户新标准不再阻塞；值/类型和公式文本差异为 0 的已有候选证据保留，但真实 Excel COM 重算失败，生产复制/重算链仍受阻。未重复全量格式对拍、真实全量样本测试或 Linux/deb 验收。

### 2026-10-27 Codex 增量：S5 V_PRE 首版明确不支持

S5 规则执行钩子现会在取数前拦截 `V_PRE`，无论是否绑定上期文件都返回单条规则状态 `UNSUPPORTED`、错误码 `FUNCTION_NOT_SUPPORTED`，并在结果表显示“规则不支持”；`V_CUR` 路径不变。S5-F01 整体状态仍为未验证；这项边界处理不代表真实样本验收通过。

### 2026-10-27 Codex 增量：Windows 最小手测交付

CLI 既有隔离构建为 `F:\RustBuild\audit_manual_handoff_20261027\debug\open-data-audit.exe`；本轮 GUI 单独重编译至 `F:\RustBuild\audit_manual_handoff_20261027_final\debug\open-data-audit-gui.exe`。最终 GUI 服务运行于 `http://127.0.0.1:8753`，HTTP 200；因 `start_central_forms` 未注册，S3-F05 的旧 GUI 按钮已明确禁用并标为尚未接通。此项只做入口安全烟测，尚未进行浏览器业务点击。S2-F01 模拟输入位于 `manual_acceptance/20261027/s2_f01/`：本期/上期使用同一合成报表，独立预期的 11 个指标变动均为 0；既有 Rust 输出为 15 条结果/4 条硬性命中，配置自身会触发规则，不能当作零异常案例。带外部文件路径读入 4 条数据并产生 3 条输入提示，输出 18 条结果/7 条命中；外部比较规则实际命中 0，因此不声称完成外部不一致异常样例或新的 Python/Rust 对拍。S2 仍未验收通过。

S1-F02 已补充 Rust 生成的任意行汇总/历史富化手测样例：两个源簿分别 2、3 行，正常输出 5 行且 history_rows_matched=3；表头异常案例对 3 个源簿结构检查 2 通过/1 跳过，保留 5 行有效结果；输出已由 Calamine 重开并核对历史拼接字段，命令和独立预期见 `manual_acceptance/20261027/README.md` 与 `expected-results.txt`。此合成 CLI 样例不代表真实 S1-F02 人工验收或 Python/Rust 文件级对拍。S3-F01、S3-F02、S3-F04、S5-F01 的模拟输入、配置/模板、独立预期和 Rust 实际输出见 `manual_acceptance/20261027/`。S3-F02 新增普通筛选（1 行）和多地区机构区分（2 行）两例，CLI 实跑后由 Calamine 按预期读回；Python/Rust 文件级重对拍和 GUI 人工操作仍未执行。普通字体、填充、边框、对齐、尺寸及打印外观差异不再阻塞当前业务验收；数据/类型、公式、映射、审核结果、结构和文件可打开性仍是必要门槛。S1-F01 真实公式复制/Office 重算仍受阻，S5 V_PRE 首版仍明确不支持；Windows 浏览器业务操作和 Linux/deb 未验证。

### 2026-10-27 Codex 增量：S2-F01 外部核对手测样例

补充 manual_acceptance/20261027/s2_external_hit/，以正式 S2 CLI 运行外部规则正常/异常两例。正常外部值 1,234,567.89 元时，外部异常为 0；异常值 1,234,500.00 元时，MANUAL-EXT-001 产生 1 条硬性异常，差额 67.89 元。两份 3-sheet《报表审核清单》均通过共享 Calamine reader 重开并核对外部结果行。CLI 运行摘要分别为 17 条结果/6 条异常和 17 条结果/7 条异常。本轮只验证合成工作簿，没有重跑 Python 对拍或真实数据；状态更新为“已实现，待 Windows 真实验收”，等待人工验收。

### 2026-10-27 Codex 增量：S1-F01 合并单元格载荷保护

S1 公式复制编辑计划不再授权编辑器丢弃合并单元格非锚点上的既有值。合并范围内的数据仍按工作簿内容保护；遇到当前共享编辑器无法保留的样本时，操作应在写入前失败，不能沿用旧 Python 序列化造成的载荷丢失。`builds_one_protected_workbook_plan_for_template_ranges` 定向单测通过。现有真实样本曾记录有非锚载荷；按当前共享预检代码，这些载荷会阻断写入。本轮未重新执行真实样本，S1-F01 公式复制与 Office 重算链继续受阻。

### 2026-10-27 Codex 增量：S3-F01 五段式/八段式 CLI 表达式入口回归

新增 S3-F01 CLI 集成回归：同一合成配置同时包含旧版“表达式校验（兼容）”八段式规则和五段式规则，分别通过两个正式 CLI 模式运行正常值（100 对 100，不触发）与异常值（120 对 100，触发）。4 次执行均生成可由 Calamine 重开的结果簿；启用规则数、触发数、错误数以及“是否说明”输出单元格均符合独立数值预期。Windows `cargo test --offline --locked -p audit-core --test s3_execution_compare_cli` 2 passed。此前 CLI 集成测试只覆盖五段模式空规则，本测试补上两个表达式模式的真实入口覆盖。S3-F01 仍待真实业务对拍及 GUI 人工验收；本测试不升级整体验收状态。

### 2026-10-27 Codex 增量：S2 默认配置副本自动识别

Audit GUI 两个启动入口（`open-data-audit serve` 与 `open-data-audit-gui`）共用 S2 配置解析：先按既有目录优先级选择标准文件名；标准文件缺失时识别唯一的 `2.报表采集系统_配置 (N).xlsx` 数字副本，多个副本则明确要求通过 `--config` 或 `OPEN_DATA_AUDIT_S2_CONFIG` 选择，避免静默挑错配置。Windows 原生定向单测 4/4 通过，覆盖标准文件优先、两种括号空格写法、多个副本拒绝及排除相似非副本文件；两个 GUI/CLI 入口编译检查、`rustfmt --check` 和 `git diff --check` 通过。S2 业务审核结论及整体验收状态未变。

### 2026-10-27 Codex 增量：S2 无默认配置时保持 GUI 可启动

当标准 S2 配置与唯一数字副本均不存在时，解析器恢复旧版相对路径回退 `2.报表采集系统_配置.xlsx`，使 GUI 能启动并在状态面板提示配置缺失，不阻断其他界面功能。多个副本仍要求显式选择。定向路径测试 4/4、两个启动入口 `cargo check`、`rustfmt --check` 通过；无默认配置时的浏览器操作未验证，S2 业务验收状态未变。

### 2026-10-27 Codex 增量：S1-F01 内容复制计划门槛验证

S1 公式复制计划不再要求复制或逐项对齐普通外观样式，仍保留值/类型、公式、批注、合并、条件格式、数据验证、定义名称、关系和工作表身份保护。Windows `s1_formula_copy::tests` 7/7 通过，候选错误分类测试 1/1 通过，Python 对拍脚本 AST 解析、目标 Rust 文件 `rustfmt --check` 与 `git diff --check` 通过。真实 ImportRange 候选未重跑：既有独立目标报告中 5,487 个范围格值/类型和公式差异为 0，但输出读回仍有 OOXML 关系语义失败；重算未执行，S1-F01 生产路由仍 Unverified/受阻。

### 2026-10-28 Codex 增量：S1-F01 GUI 正式审核流程接线

GUI worker 在原有逐文件模板匹配、模板体检和表结构预检后，现继续调用 S1 公式副本计划与 Excel-Core、外部工作表导入、真实重算 sidecar、公式异常和校验区域异常提取、重算副本上的条件格式触发提取、Issue 合流、配置簿历史行读取/归类及 13 列结果写入。批次审核副本和结果放在独立新目录；失败删除本次目录、源文件不变；失败与成功都在 GUI 日志逐文件显示状态/跳过原因，成功记录实际结果路径。S1 校验区域扫描已改为读取重算副本。新增输出目录 guard 测试及 Umya 公式复制候选测试源码，均未运行。

当前生产 Excel-Core 路由列表仍为空；没有 S1-F01 全操作批次的 Passed 保留证据，故实际任务会在副本编辑前明确拒绝，尚不能生成最终结果，也没有验证本机 Excel/LibreOffice 是否可用。`web_bridge.rs` rustfmt 检查及 `git diff --check` 通过；CLI rustfmt 检查仅报告现存 import 排序差异，本次未格式化整文件。未执行 Cargo/build/check/test/run、编译、打包、真实工作簿操作或 Git push。S1-F01 状态仍为代码流程已接线但生产执行受阻，后续需经授权测试/能力验证后再做 Windows 样本和 GUI 人工验收。

### 2026-10-28 Codex 增量：S1-F01 同坐标导入保留原公式

现有 S1 模板区域复制的源/目标坐标相同，但 Excel-Core 的平移导入路径仍会对公式运行转换解析器。`UmyaXlsxBackend` 现仅在坐标确实变化时才平移公式；坐标相同则复制完整公式对象，避免无意义改写并允许原样携带结构化引用/外部引用。新增集成测试源码覆盖同坐标跨簿导入上述两类公式及源字节不变。本次只做 `rustfmt --check` 和 `git diff --check`，未运行 Cargo 或测试；该补丁尚未证明此前真实候选中的 78 个公式/值差异已修复。S1-F1 生产路由仍无 Passed 证据，实际审核依旧失败关闭。

### 2026-10-28 Goal 收口补充：S1-F1 真重算、Linux 交付与验收运行器

本段记录当前工作区在本次 Goal 收口时的实测状态；上方历史段保留其原始时点结论，不应覆盖本段之后的事实。

- **S1-F1 Excel COM 合成端到端**：`s1_directory_copies_template_formulas_recalculates_and_combines_real_issues` 在 Windows 上以 `--ignored --exact` 显式运行，1 passed。真实 Microsoft Excel COM 16.0（EXCEL.EXE 16.0.20527.20088）重算合成公式副本；独立预期及读回为 `检查表!D1=#DIV/0!`、`D2=#REF!` 各 1 条公式异常，`B2` 有 1 条实际触发的条件格式异常；合流结果为 13 列，工作簿可重开，源文件、模板和重算前审核副本哈希不变。此证据不等于真实机构文件验收。
- **Linux 生产能力与 deb**：`audit-core/tests/production_capability_verification.rs` 在 Linux x86_64 1 passed，证据写入 `audit-core/capability_evidence/umya_linux_x86_64_20261028.json`；证据覆盖 S1 编辑批次、S3-F3 备份后原地更新及 S3-F5 编辑批次，未知平台仍失败关闭。离线包为 `dist/linux-delivery/open-data-audit_0.1.0_amd64.deb`（13,498,066 字节）；musl 静态二进制，gzip 控制包，未声明系统依赖。在 Debian GLIBC 2.28 / `dpkg-deb` 1.19.8 等价容器中安装、CLI、HTTP API 和 S3-F1/F2/F5 合成 CLI 均通过；容器无 `soffice` 时 `lo-convert recalculate` 明确返回 LibreOffice 不可用。麒麟 V10/UOS V20 实机仍未验收。
- **验收运行器**：`acceptance-run` 从 JSON 清单调用既有入口并输出 JSON/Markdown 报告；独立合成运行中 S3-F1、F2 与 S5 均完成，输入文件哈希全不变。报告位于 `F:\RustBuild\audit\acceptance-runner-synthetic-0b813b7c867b4f979e155fccad36f151`。运行器的报告指标和输入变化检查测试均通过。
- **Windows Release/ZIP**：`build-release.bat` 与 `package.bat` 成功；ZIP 为 `dist/open-data-audit-windows-x64-20261028-122102.zip`。解压版本通过 `--version`、验收/S3-F1 CLI 帮助及 `get_state` API；随包 3.0–3.3 配置可发现。`package.bat` 原 LF 行尾造成命令解析异常，转换为 CRLF 后成功。
- **完整回归**：`CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -- --quiet` 汇总 435 passed、0 failed、28 ignored；另显式 Excel COM 集成用例 1 passed。所有忽略的真实样本/外部运行时用例仍保持未验收状态。
- **业务口径确认**：用户已确认 S3-F3 必须先备份并原地更新目标文件；S3-F5 的 3.3「分析文件数据单位」是历史遗留字段，沿用 3.0 默认目标单位。实现与验收运行器口径均按此记录。

剩余事项仅是用户提供真实 S1/S2/S3/S5 业务配置与输入后的人工验收，以及在麒麟 V10/UOS V20 实机安装/操作验收；不得把本段合成样本或等价容器记录描述为生产业务验收。

### 2026-10-28 Codex 增量：S5-F1 GUI 批次与 V_PRE 接线

沿用现有 Rust 配置审核、取数、规则执行、条件格式和结果写入模块，补齐 GUI 从文件选择到审核输出的批次链路：本期目录中勾选多份 XLSX/XLSM 后按同一模板/业务/报告期审核并生成一份汇总结果；可选上期目录自动扫描 XLSX/XLSM，填写较早的上期报告期后，按机构代码匹配上期工作簿并绑定到现有 `V_PRE` 取数。旧 Python S5 README 说明原面板暂未显示上期文件框，但其 Web API 与 `run_batch` 服务已有 `previous_files`/`previous_period` 输入；Rust 现把该既有服务能力接入可操作的 GUI。没有匹配上期、匹配多份或上期身份识别失败均保留明确诊断/失败信息；不会将缺失的上期值当作 0 或通过。当前规则结果、条件格式发现、结构观察、数据质量提示、选中本期/上期/外部数据及模板的哈希记录继续使用现有单份结果簿/运行总账输出。

S5 GUI 仍只提供一个外部工作簿；启用规则引用多个逻辑外部来源时会明确拒绝并提示使用现有 CLI 多来源绑定或拆分配置。为上一期身份诊断补充结果/日志字段，并补充前端路径选择测试源码；本轮未执行 Cargo build/check/test/run 或 GUI 运行，代码尚未编译及测试。S5-F1 仍待用户用真实配置、本期/上期文件和可选大集中外部文件进行 GUI 人工验收；此记录只描述代码接线，不宣称本轮功能测试通过。

### 2026-10-28 Codex 补充：S5-F1 定向静态复核与状态纠正

用户询问是否应直接人工验收后，沿 S5 GUI 批次、来源绑定、结果 writer 及其对应 Python S5 服务做定向静态复核。确认批次哈希键覆盖、无证据结果的文件归属错误、审核副本命名碰撞、外部身份参照错误被忽略及规则修订/生效期未用于来源预检等问题；另核实旧 Python 的可选模板条件格式范围、校验公式缓存错误扫描和批次不持久化结构快照路径未完整对齐。完整触发条件与位置记录在 `GPT修改记录/20261028_150010_S5F1定向静态复核与状态纠正.txt`。

当前 S5-F1 状态改为“部分实现，需要继续开发；存在已知问题”，先修复再定向验证，之后才进入真实业务人工验收。本轮只纠正当前状态与记录，未修改业务代码，未执行 Cargo、测试、Python/Rust 运行或工作簿操作。S1/S2/S3 不在本次定向复核范围；既有实现和历史证据保持原有范围。

### 2026-10-28 Codex 增量：S5 定向修复与 S1-F1 外部公式保护

S5 上节八项问题已完成源码修复：为批次结果保留所属本期文件，副本文件名碰撞时独占创建新文件，按物理路径保存重复逻辑来源哈希并复核所有最终副本；来源预检和执行共用有效规则修订选择，显式 source 不再与未使用默认值冲突；机构参照加载与业务数值布局分离，已选参照的错误明确报告；条件格式使用模板可选“条件格式区域”；GUI 逐文件及 CLI 指定模板下的主本期文件复用已有只读校验公式缓存错误扫描。GUI 批次只比较结构快照，保留机构/文件上下文，不更新共享基线；逐文件记录写入运行总账，并区分部分失败、全部失败状态。

用户补充的旧 Python “纯数据 OOXML 优先、外部表含公式时回退 openpyxl”归属 S1-F1。当前 Rust S1-F1 本已通过支持公式的 Umya ImportRange 导入外部表，本次没有把它替换成拒绝公式的纯数据导入器，也未注册未经验证的 openpyxl 生产能力。修复了外部已用区域不从 A1 起始时仍导入 A1 的坐标错误，补齐外部公式的传递工作表依赖，动态引用时保守选择全部外部工作表。新增 C4→另一表 B1、公式保留与源文件保护的回归用例源码。

新增七个定向测试函数的源码，均未运行；仅执行静态检查与 rustfmt，不执行 Cargo 编译、测试、程序运行或打包。已有 dist/ZIP/deb 未更新，历史测试结果不证明本次修复。详细文件与边界见 `GPT修改记录/20261028_152620_S5定向修复与S1F1外部公式保护.txt`。

### 2026-10-29 Codex 增量：Umya 真实 S1-F1 复制验收与 S1-F3 发布幂等

本记录只增加本日验证，不覆盖上方历史结论。S1-F1 真实样本继续使用原 Python 项目的真实输入、模板、外部表和映射配置作为只读参考，Rust 生产 Umya 路由执行外部表和模板范围导入；Windows Excel COM 可重开各阶段副本。约 210 个 `参照表!D2:D211` VLOOKUP 与跨工作表依赖保持公式文本，COM 重算值与独立 Excel 重算外部表基线一致，所有输入哈希保持不变。样本涵盖纯数据表、普通公式、跨 Sheet 公式和「参照表」/「机构参照表」共同复制。该结果验证复制、保存、重开和重算，不等于 GUI 最终审核结果人工验收。

确认 S1-F1 当前跨簿公式复制只由 Umya 承担；OOXML Patch 没有已实现且通过验证的同语义公式工作表 ImportRange，因此不登记为此操作的回退。通用 ExcelModifier 只在首选后端“写入前不支持”且后备对同一能力有有效登记时，从原始副本重试；两者均不支持则失败关闭。回退协议测试是合成路由协议验证，不虚报 SetCellValue 的生产 OOXML 支持。

修正 S1-F3 结果发布在文件创建后、任务状态落盘前中断会产生重复报告的问题：按 run/审核结果指纹确定输出身份，先持久化候选路径及摘要，再发布最终文件；重启后摘要匹配则复用，审核结论变化时输出新版本并保留旧文件。新增端到端 writer 回归验证存储重建后的结果路径/字节稳定、结论变更不覆盖。Calamine 已有未知错误值临时 OOXML 规范化回退仍仅适用于 XLSX/XLSM 特定错误 token；本日 `excel-reader` 17 个单测+两个集成目标共 19 passed，2 个集成测试均 passed；新增分类用例涵盖文件缺失、权限拒绝、不支持格式、工作表读取错误、损坏包、未知错误类型和缺失文件不走回退。`ooxml-patch` 15、S1 恢复状态 4、S1 复制计划 12、S1-F3 发布幂等 1、ExcelModifier 两个路由测试和 `lo-convert` 8 单测+1 CLI 集成均通过。批次隔离断言已修正为保留成功文件上的条件格式诊断，并通过聚焦回归。此记录写入时 full workspace 第二轮回归尚未开始；最终统计、Windows Release/入口冒烟和未解决项将在本记录后续增补。

#### 最终回归与 Windows Release 补充

2026-10-29 完成 `CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -- --quiet`：80 个测试目标，459 passed、0 failed、31 ignored。完整日志位于 `F:\RustBuild\audit\workspace-test-final-20261029.log`。首次完整运行暴露一条过时 CLI 集成断言，仍期待单个工作簿失败时丢弃整个批次；按现有逐文件隔离行为改为验证成功文件的 4 条结果保留、失败文件单独标记且批次 `completed_with_diagnostics`，定向用例和随后完整回归均通过。依赖本地真实业务文件/人工操作的 ignored 用例不计作验收。

同日 `build-release.bat` 成功构建并将 `open-data-audit.exe`、`lo-convert.exe` 复制到项目 `dist`。当前产物的 `--version`、CLI 命令帮助、重算侧车帮助通过；本轮未重打包 ZIP。`rustfmt --check`（本轮触及的 Rust 文件）及 `git diff --check` 通过。仍待用户按真实 S1-F1 模板审核和 S1-F3 人工重算恢复流程进行 GUI 验收；缓存节点存在不能证明缓存新鲜或计算结果正确。

#### 2026-10-29 增量：Excel COM 持续 xlPending 修复与公式结果复核

对前述 S1-F1 COM 合成 E2E 的 pending 失败进行定点复查。Excel 打开同时标记 `fullCalcOnLoad=1`、`forceFullCalc=1` 的公式副本时，即使执行 `CalculateFull()` 和 `CalculateFullRebuild()` 并已在内存得到 `#DIV/0!`、`#REF!` 等计算结果，`Application.CalculationState` 仍可能持续返回 `xlPending`。受控 COM 样本证实：仅移除此两项加载标记后，缓存仍已清除、COM 仍执行全量计算和依赖重建，状态会到达 `xlDone`。因此 sidecar 现在为 Windows COM 路由创建独立输入副本，将两个加载标记设为 false，仍保留原件、公式文本及无缓存状态；LibreOffice 路由继续收到全量重算标记。COM 仍在 `xlPending` 时不会保存。

保存后的复核现要求工作表顺序与公式文本保持不变、每个公式均有缓存节点，且 Calamine 能为每个公式读取带类型的结果值；公式错误和空字符串都是有效计算结果类型。这个校验确认了缓存已由外部计算引擎产出，但不独自证明数学结果正确。S1-F1 业务 E2E 另外检查独立预期的 3 条公式异常与 1 条条件格式异常。

验证：`lo-convert` 13 个单元测试及 1 个集成测试通过；S1-F1 `s1_directory_copies_template_formulas_recalculates_and_combines_real_issues` 使用 Release `lo-convert.exe` 经 Windows Excel COM 通过。完整回归为 80 个测试目标、464 passed、0 failed、31 ignored，日志 `F:\RustBuild\audit\workspace-test-final-20261029-com-pending.log`。`build-release.bat` 成功；`open-data-audit.exe --version` 与 `lo-convert.exe --help` 冒烟通过。sidecar 单文件 `rustfmt --check` 与 `git diff --check` 通过；`cargo fmt --all -- --check` 仍报告工作树其他既有文件格式化差异，本轮未格式化或覆盖无关 WIP。

本轮运行环境未提供真实 S1-F1 源文件、模板、外部表及映射配置路径，故 210 条 VLOOKUP 的真实工作簿测试未在这次 COM pending 修复后重跑；该工作簿验证有本日此前运行记录，不能替代修复后的复验。LibreOffice 真实复杂工作簿和麒麟/UOS 目标机仍待验。S1-F1 GUI 最终结果和 S1-F3 人工强制重算/重启恢复仍需用户验收，总 Goal 继续 active。

#### 2026-10-29 增量：公式结果可读性及 LibreOffice 公式规范化

补齐计算结果复核边界：S1 sidecar 重算副本和 S1-F3 人工恢复现在都要求公式单元格具有 OOXML 缓存节点，且 Calamine 对每个公式都读出带类型的结果。公式错误和空字符串都属于可消费结果；缓存节点存在但读回空值时，S1-F3 会保留 `WaitingManualRecalc` 并提示重新计算、保存、关闭。该校验证明结果可读，不单独证明数学正确。

LibreOffice 保存时会将 `TRUE`/`FALSE` 常量写为 `TRUE()`/`FALSE()`。为避免误拒真实重算，公式比较与 S1-F3 恢复指纹仅容忍这两种等价布尔字面量写法；公式单元格位置、工作表顺序/名称、跨表引用、范围以及其他公式文本仍要求一致。公式字符串中的 `"TRUE()"` 和单引号包围的工作表名 `'TRUE()'` 都不做规范化。S1-F3 保留精确指纹，另加带默认值的规范化指纹，以支持新任务并兼容旧任务清单格式。

验证：WSL x86_64 `lo-convert` 13 个单测和 1 个实际 LibreOffice CLI 集成通过；集成覆盖 VLOOKUP 结果、跨表公式、`#DIV/0!`、空字符串结果、公式位置/内容和工作表结构。S1-F3 规范化指纹定向测试通过。Windows 全工作区回归为 80 个测试目标、467 passed、0 failed、31 ignored。首次回归遇到一条并行测试临时目录名碰撞；辅助函数已改用原子序号和冲突重试，之后完整回归通过。之后又为避免把 `'TRUE()'!A1` 工作表引用误认为布尔函数增加了严格的单引号保护；该边界变更后的公式比较单测、WSL LibreOffice 集成和 Windows COM E2E 均再次通过，完整工作区计数对应边界保护前代码。

Windows `build-release.bat` 成功重建 Release 二进制；`open-data-audit.exe --version`、`lo-convert.exe --help` 通过。以新 Release `lo-convert.exe` 运行 S1-F1 `s1_directory_copies_template_formulas_recalculates_and_combines_real_issues` 的 Microsoft Excel COM 端到端测试通过：公式异常、触发的条件格式异常、13 列合流输出及副本/输入保护断言均通过。触及的 Rust 文件 `rustfmt --check` 和 `git diff --check` 通过。

当前进程仍无可用 `AUDIT_S1_REAL_*` 工作簿路径，本轮未重跑真实旧 Python S1-F1 工作簿中约 210 个 VLOOKUP 的 Umya 复制/重算样本；相关 2026-10-29 先前记录仍是已有证据，但没有覆盖之后的 COM `xlPending` 修复。真实机构工作簿、S1-F1 GUI 最终结果、S1-F3 人工保存/恢复/重启、麒麟 V10/UOS V20 与真实复杂工作簿 LibreOffice 验收仍待完成，总 Goal 保持 active。

#### 2026-10-29 增量：S1-F3 人工恢复副本输入完整性

F1 等待人工重算的副本先清空公式结果缓存，因此 F3 可要求用户重新计算；但此前恢复只核对公式文本和可读缓存，未核对非公式输入或条件格式规则。现在新清单另存语义摘要，包含非公式单元格值、条件格式的适用范围/规则内容/相对优先顺序、合并区域及 OOXML 命名区域，排除公式结果值；F3 继续检查公式指纹与缓存后，还要求摘要一致。摘要忽略条件格式样式 ID、无关默认属性及解析后等价的命名区域排序，兼容 LibreOffice 重存；数据验证不参与审核判定，不纳入摘要。由此允许重算产生公式结果缓存，同时拒绝用户在手动重算期间改动审核输入或审核规则。旧清单的新增字段为可选，历史任务仍按已有公式完整性行为恢复。

已新增合成测试：公式缓存清除前后摘要一致；编辑非公式单元格后摘要变化；恢复指纹校验拒绝输入摘要变化，并兼容缺少该字段的旧清单。WSL x86_64 已用真实 LibreOffice sidecar 重算 `synthetic_probe.xlsx` 并验证摘要兼容通过。修改后 Windows 全工作区回归共 80 个目标，470 passed、0 failed、31 ignored，日志 `F:\RustBuild\audit\workspace-test-s1f3-semantic-guard-20261029.log`；包括 S1-F3 恢复指纹/持久化/结果发布幂等单测。Windows Excel COM 的 S1-F1 合成端到端测试通过。`build-release.bat` 已重建最新 Windows Release，版本与 sidecar 帮助冒烟通过；使用新 Release sidecar 再运行 S1-F1 合成 COM 端到端通过，日志 `F:\RustBuild\audit\s1f1-com-e2e-release-semantic-guard-20261029.log`。该校验不证明缓存的新鲜度或数学正确性；仍需要真实 Excel/LibreOffice 人工强制重算保存及 GUI 验收。

#### 2026-10-29 增量：COM 修复后真实 S1-F1 Umya 复验

上方 01:40 记录所述“当前进程无真实样本路径”仅指当时运行环境。随后在旧 Python 参考项目的只读目录定位到真实单位贷款源文件、模板、S1 映射、外部数据及同机构旧 Python 审核副本；没有修改这些文件。复用 `audit-core/tests/s1_real_umya_copy.rs`，使用正式生产 Umya modifier 和最新 Windows Excel COM sidecar，在 `F:\RustBuild\audit` 运行 3 个 ignored 探针，3 passed / 0 failed，完整日志 `F:\RustBuild\audit\s1f1-real-umya-final-20261029.log`。

- 纯数据源往返：公式数为 0 时不错误调用公式重算；检查表名/顺序、公式和所有非空单元格值。该样本 Umya 将 211 个显式空字符串单元格规范化为缺省空白（6 张表），其余非空值相同。此差异已显式记录，不能描述为逐单元格序列化完全一致。
- 真实 S1-F1 复制/重算：覆盖外部「参照表」「机构参照表」共同导入、模板区域导入、普通/跨 Sheet 公式、`参照表!D2:D211` 约 210 个 VLOOKUP；公式文本和工作表依赖保留。外部 VLOOKUP 值与外部工作簿独立 COM 重算基线一致。
- 公式结果基线：模板 1,072 个公式位置在 Rust 副本中公式文本与模板一致；1,072 个可读公式结果与同机构旧 Python 审核副本逐格相同，其中 695 个为非空结果，0 个差异。Excel COM 分别打开外部表单独导入、模板区域单独导入和完整合并副本；完整副本由 COM 重算保存后可重新读取。源、模板、外部数据及映射配置的字节哈希均保持不变。
- 引擎分工：真实 S1-F1 跨簿公式 ImportRange 由 Umya 满足。OOXML Patch 尚无该操作的已验证同语义能力，因此不作为公式复制回退；生产 Umya 失败时按 Excel-Core 既有策略失败关闭。通用能力路由的“Umya 写入前拒绝→从干净副本走已验证后备”和“均不支持→失败关闭”合成测试已包含在此前 470 项 workspace 回归中。

以上验证是单个真实单位贷款样本的工作簿复制与业务结果证据，不代替用户在 GUI 中检查批次跳过原因、条件格式诊断和 13 列结果，也不覆盖所有机构文件。真实 GUI S1-F1 与 S1-F3 人工重算恢复仍待验收。

#### 2026-10-29 增量：S1-F3 恢复 worker 端到端与缓存缺失拒绝

为让 S1-F3 的恢复服务可以在隔离状态目录中验证，将 worker 内部实现分为默认用户状态目录入口和可注入状态存储入口；生产 GUI 仍调用默认目录。新增两个合成 worker 测试：

- 重建 `S1RecoveryStore` 实例后读取持久化待处理任务，执行真实 F3 模板预检、公式/条件格式提取、历史/结果合并和结果发布；确认提取到 1 条公式异常、结果表含 13 列且可由 Calamine 重开、源文件/模板/审核副本哈希保持不变、任务完成后待处理任务数为 0。
- 使用现有 OOXML 清缓存实现生成缺缓存审核副本；确认 F3 不生成结果、任务继续处于 `WAITING_MANUAL_RECALC` 并带有缺缓存诊断，副本不被修改。

两条 worker 测试 2/2 通过；完整串行工作区回归 80 个测试目标、472 passed / 0 failed / 31 ignored，日志 `F:\RustBuild\audit\workspace-test-s1f3-worker-e2e-serial-20261029.log`。默认并行全工作区回归曾有 Windows `rustc` 子进程以 `STATUS_STACK_BUFFER_OVERRUN` 退出；限制 Cargo 并发为 1 后全量回归通过。`build-release.bat` 成功，Release 主程序版本和 sidecar 帮助冒烟通过。

该合成缓存只用于验证 F3 恢复阶段，不代表 Excel/LibreOffice 真正重算。还需用户在 GUI 中由 F1 产生待恢复任务、用真实 Excel/WPS/LibreOffice 强制重算并保存、重启后在 S1-F3 恢复，核对真实审核结果；S1-F1 真实 GUI 整批输出验收也仍待完成。

#### 2026-10-29 增量：S1-F1 复制后公式与输入值校验

S1-F1 GUI 复制 worker 在 `ExcelModifier` 发布到批次临时目录后、清理公式缓存和调用外部计算引擎前，新增 S1 业务层候选校验。校验按编辑计划重建目标工作簿的预期工作表顺序、公式映射与非公式存储值：逐个导入模板/外部数据范围，要求公式单元格位置和文本完全匹配来源，引用表的非公式输入值和类型保持一致；公式缓存值单独排除，因其会在正式计算前清除。工作表名称/顺序必须精确匹配预期。唯一容忍的值表示差异是空字符串被 Umya 规范化为缺省空白，这与此前单样本观察一致；其他差异失败关闭，错误指出差异工作表/坐标，不输出单元格内容。

新增合成回归覆盖实际 `execute_s1_formula_copy_with_manual_recovery` 的待人工重算路径、模板公式丢失、工作表结构额外变化、外部「机构参照表」非公式值被改动、210 条 VLOOKUP 及公式依赖复制，并验证空字符串兼容不掩盖非空值差异。Windows `CARGO_TARGET_DIR=F:\RustBuild\audit cargo test --offline --locked -p audit-core --lib s1_formula_copy::tests:: -- --nocapture`：15 passed / 0 failed；完整日志 `F:\RustBuild\audit\s1f1-runtime-copy-validation-20261029.log`。

全工作区串行回归 `CARGO_TARGET_DIR=F:\RustBuild\audit cargo test --offline --locked --workspace -j 1 -- --quiet`：80 个目标，475 passed / 0 failed / 31 ignored；日志 `F:\RustBuild\audit\workspace-test-s1f1-runtime-copy-validation-serial-20261029.log`。`build-release.bat` 成功完成 Windows Release 并复制到项目 `dist`；`open-data-audit.exe --version` 输出 `0.1.0`，`lo-convert.exe --help` 显示 `convert` 和 `recalculate` 命令。目标 sidecar/主程序只做帮助与版本冒烟；未运行完整 COM S1 GUI E2E。

本轮运行进程未配置真实旧 Python 工作簿路径，故新增运行时断言尚未在真实业务工作簿上复验；历史真实 Umya/COM 证据见上文，不能当作本次新增断言的真实样本证据。S1-F1 GUI 业务验收及 S1-F3 人工重算恢复仍待用户验收，总 Goal 继续 active。

#### 2026-10-29 增量：Umya 空白覆盖缺陷修正与候选校验真实复验

随后从旧 Python 参考项目的只读目录恢复真实样本输入路径，并复跑现有 S1-F1 真实探针。新增候选校验首次拦截了一个实际复制缺陷：源报送工作簿中存在公式/旧值，而模板导入范围对应单元格在源模板中缺失；Umya `ImportRange` 将该缺失单元格记录为“无写入”，没有清除目标旧公式/旧值。该问题不属于校验器误报，也不能通过忽略公式差异解决。

修复限定在 Umya 后端：当跨工作簿导入显式设置 `materialize_blank_cells=true`，源范围中不存在物理单元格时，清空目标单元格的旧值/公式；不复制样式或属性时继续保留目标既有样式行为。扩展 `crates/excel-editor/tests/import_range.rs` 合成回归，在已有目标单元格中分别放置旧公式和文本，导入空白区域后确认两者均被清除。未扩展 OOXML Patch，也未改变跨工作簿公式复制的引擎路线。

真实复验通过：使用旧 Python 的真实单位贷款源、模板、S1 映射、外部表和同机构 Python 审核副本，仍只读原件；测试副本与重算输出均位于系统临时目录。三个 ignored 探针 **3 passed / 0 failed**，日志 `F:\RustBuild\audit\s1f1-real-umya-candidate-validation-20261029.log`。涵盖纯数据 Umya 往返、真实模板与外部工作表分段导入、`参照表!D2:D211` 约 210 个 VLOOKUP、`参照表`/`机构参照表` 共同导入、完整 Excel COM 真实重算；1,072 个可读模板公式结果与同机构 Python 审核副本逐格一致，989 个非空、0 差异；原始输入文件哈希未改变。

修复后全工作区串行回归 `cargo test --offline --locked --workspace -j 1 -- --quiet`：**80 个测试目标，475 passed / 0 failed / 31 ignored**，日志 `F:\RustBuild\audit\workspace-test-s1f1-blank-overwrite-serial-20261029.log`。定向 `excel-editor --test import_range` 为 1 passed。`build-release.bat` 成功，产物复制到 `dist`；主程序版本和重算 sidecar 帮助冒烟通过。真实 S1-F1 GUI 整批输出、手动 S1-F3 恢复及 Linux 目标机验收仍待人工完成，总 Goal 继续 active。

#### 2026-10-29 增量：S1-F1 条件格式异常局部隔离

条件格式读取和求值链路此前可能被一个未验证的工作表读取状态、一个不能识别的单元格值、一个坏的 `sqref` 元数据节点或单个触发值转换错误整体打断；S1-F1 的公式异常虽已提取，却可能因此不能进入该工作簿的最终汇总。本次调整为逐工作表/元数据节点/触发单元格隔离：状态未验证的表跳过并诊断；值转换失败时整表条件格式判断跳过以避免基于部分值误报；单个坏节点跳过且保留同表其他节点；单个审核异常转换失败只跳过对应触发项。若整个条件格式提取函数仍返回错误，保留该工作簿已经提取的公式异常、将其标为部分结果并记录条件格式阶段错误。公式提取、文件读取或重算完整性失败不伪装成部分成功，仍按该文件失败处理。

新增/扩展回归覆盖坏条件格式节点与有效节点并存、单表值类型不支持但其他表仍能产生触发、整个条件格式提取器失败时公式异常保留。`cargo test --offline --locked -p audit-core --lib`：265 passed / 0 failed / 10 ignored。全工作区串行回归 `cargo test --offline --locked --workspace -j 1 -- --quiet`：**80 个测试目标，478 passed / 0 failed / 31 ignored**，日志 `F:\RustBuild\audit\workspace-test-s1-partial-condition-20261029.log`。Sidecar 13 个重算单测通过；同套件中的 LibreOffice CLI 集成目标在本 Windows 环境因未安装 `soffice` 按测试定义跳过，不作为本轮实际 LibreOffice 重算证据。Linux/WSL 实际 LibreOffice 重算证据见本清单前文独立记录。定向 `rustfmt --edition 2021 --check audit-core/src/conditional_format.rs audit-core/src/s1_conditional_audit.rs audit-core/src/web_bridge.rs` 通过；全仓 `cargo fmt --all -- --check` 因工作区其他既有文件存在格式差异而未通过，本次没有对无关文件做全仓格式化。`git diff --check` 无空白错误（仅提示现有 CRLF 转换警告）。`build-release.bat` 成功重建并复制主程序与 sidecar；`open-data-audit.exe --version` 输出 `0.1.0`，`lo-convert.exe --help` 显示 `convert` 与 `recalculate`。使用该 Release sidecar 重跑 Excel COM 合成 S1-F1 端到端：1 passed / 0 failed，确认公式异常、实际触发条件格式异常及 13 列结果合流仍正常，日志 `F:\RustBuild\audit\s1f1-com-e2e-release-partial-condition-20261029.log`。

该修改不代表条件格式规则已支持完整 Excel 语义；不支持的规则仍明确给出诊断。真实机构文件、GUI 批次结果及 S1-F3 人工重算恢复仍待人工验收，总 Goal 保持 active。

#### 2026-10-29 增量：S1-F3 丢失/损坏副本可见性与混合批次恢复

恢复状态重建时，已持久化为 `WAITING_MANUAL_RECALC` 等待状态、但审核副本文件已丢失或路径无效的任务现在会转成 `FAILED`，保留“副本已丢失、需重新运行 S1-F1”的原因；`lost_copy_tasks()` 将此类记录提供给 GUI，页面显示失败原因，不会给不可恢复任务提供恢复按钮。副本仍存在但内容损坏时，worker 将任务记为 `RESUME_FAILED`、显示读取错误、不发布结果并保持审核副本字节不变；该条目仍在 S1-F3 待处理列表中，以便用户修复/替换副本后重试。恢复状态目录读取错误也继续通过页面警告显式展示。

新增回归覆盖：S1-F1 合成批次中一份副本成功完成、其余文件保持人工重算待处理，成功 Issue 进入 13 列结果、待处理任务跨状态存储重建仍可见；F3 丢失副本被标记并进入页面失败摘要；损坏副本不变、不生成结果且保留可重试状态；GUI 初始隐藏/错误和丢失副本诊断状态。相关定向测试通过。最终串行全工作区 `CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -j 1 -- --quiet` 为 **80 个测试目标，482 passed / 0 failed / 31 ignored**，日志 `F:\RustBuild\audit\workspace-test-s1-f3-lost-copy-ui-20261029.log`。工作簿真样本 ignored 探针不计入该统计；S1-F1 真实旧 Python 单位贷款样本已有单独的 3/3 通过证据，覆盖纯数据、公式/跨表/VLOOKUP、两张外部表共同导入、1,072 个重算结果与基线比较及原件哈希保护。

改动后 `build-release.bat` 成功，主程序和重算 sidecar 均已复制到 `dist`；主程序 `--version` 输出 `0.1.0`，sidecar 帮助包含 `convert`/`recalculate`。使用最新 Release `dist/lo-convert.exe` 的 Windows Excel COM S1-F1 合成 E2E 为 1 passed / 0 failed，覆盖实际公式异常、触发的条件格式异常和 13 列输出重开；源码 `rustfmt --edition 2021 --check audit-core/src/s1_recovery.rs audit-core/src/web_bridge.rs` 与 `git diff --check` 通过。该证据不等于用户在 GUI 中完成人工重算恢复验收；真实 GUI 机构批次、S1-F3 强制重算/保存/关闭/重启恢复和麒麟/UOS 目标机验收仍待用户操作。未执行 Cargo push、提交或打包。

#### 2026-10-29 增量：真实 S1-F3 Excel COM 恢复与 S1-F1 当前样本复验

新增 `web_bridge::tests::real_s1_f3_recalculates_and_resumes_a_real_s1_f1_audit_copy`，仅 Windows 编译，默认 `ignored`，从 `AUDIT_S1_REAL_*` 读取本地输入。测试使用真实 S1-F1 源、模板、外部表和映射配置，在系统临时目录生成独立审核副本；先走无 sidecar 的 F1 缓存清理路径，再由 `dist/lo-convert.exe` 经 Excel COM 将重算结果写到独立文件，复制已保存内容回临时审核副本，最后重建恢复状态存储并调用 GUI 使用的 F3 worker。这样验证真实公式工作簿经过保存后能进入恢复/结果发布链路，但不模拟人工点击 GUI 或 Excel 中的键盘操作。

真实探针：**1 passed / 0 failed**；Excel 引擎为 `Excel.Application`，公式数 1,282，提取公式异常 3 条，条件格式异常 0 条、条件格式诊断 0 条，任务 `COMPLETED`，结果工作簿通过 Calamine 重开并包含 13 列目标表头；F3 后审核副本哈希不变，源文件、模板、外部文件和映射配置哈希均不变，待处理任务归零。运行日志 `F:\RustBuild\audit\s1f3-real-resume-com-20261029.log`。

本轮 S1-F1 当前样本复验：纯数据 Umya 往返与 Excel COM 分段/整本打开探针各通过。真实公式复制/COM 重算探针在无 Python 对照输入时 **1 passed**（`F:\RustBuild\audit\s1f1-real-formula-copy-current-20261029.log`），覆盖约 210 个 VLOOKUP、工作表依赖和模板公式，独立外部表 COM 重算基线、公式文本和结果缓存可读性检查通过。将同名 Python 文件作为基线时，对照只提供 973 个公式缓存结果，未达到现有至少 1,000 个覆盖断言；该次对拍不形成通过/不通过结论，也不将阈值改低。此前另一个已对齐样本的完整 Python 基线对拍记录仍有效。另一个同名外部文件在当前映射下触发候选值一致性拒绝，S1 未进入重算；选择 `config-real` 外部数据组后真实 F1 复制探针通过。

最新串行 workspace 回归：**80 个测试目标，482 passed / 0 failed / 32 ignored**，日志 `F:\RustBuild\audit\workspace-test-s1-real-f3-20261029.log`；其中新增真实 F3 探针在常规回归中保持 ignored，以上单独执行结果不与统计重复。新增测试源码 `rustfmt --edition 2021 --check audit-core/src/web_bridge.rs` 通过；`git diff --check` 无空白错误，仅有既有 LF/CRLF 转换提示。本轮没有生产业务逻辑变更，所以没有重建 Release；没有提交、push 或执行打包。人工 GUI 真实恢复及 Linux 目标机验收仍待完成，总 Goal 保持 active。

#### 2026-10-29 增量：公式重算 sidecar 超时子进程清理

超时处理此前只保证终止和回收直接子进程；Windows 的 `taskkill /T` 在父进程已被回收后才执行，Linux 也没有进程组清理，存在 LibreOffice/引擎辅助进程遗留的风险。现在 `run_command_with_timeout` 在 Unix 启动独立进程组；超时或等待状态错误时先终止完整进程组，再 kill/wait 直接子进程。Windows 在父进程仍存活时先 `taskkill.exe /PID <pid> /T /F`，随后直接 kill/wait 作为回收兜底。未更改正常执行的引擎顺序、Umya/OOXML 路由或公式业务语义。

新增 `engine_process_timeout_terminates_descendants_too` 回归：Windows 隐藏 `cmd` 子进程和 Linux `sleep` 子进程均由真实超时流程启动，之后轮询确认子进程不再执行；保留原有直接子进程超时报错用例。验证结果：Windows `cargo test --manifest-path sidecar/lo-convert/Cargo.toml -- --nocapture` 为 14 单测通过，LibreOffice 集成入口因本机无 `soffice` 明确跳过；Ubuntu WSL 同命令在 `LO_CONVERT_REQUIRE_LO_TEST=1` 下为 14 单测通过及 1 个 LibreOffice CLI 真重算集成通过。`cargo fmt --manifest-path sidecar/lo-convert/Cargo.toml -- --check` 通过。Ubuntu 使用 `/mnt/f/RustBuild/audit/linux-x86_64` 作为 Cargo 输出目录；首次非登录 shell 调用未找到 Cargo，改用登录 shell 后成功，该次没有改动源码/配置。

Windows `build-release.bat` 已成功重建主程序与 `dist/lo-convert.exe`；构建有现存 Umya 弃用 API 和若干未使用代码警告，没有编译错误。新 Release sidecar 经 `s1_directory_copies_template_formulas_recalculates_and_combines_real_issues` Windows Excel COM 合成端到端 **1 passed / 0 failed**，确认公式异常、实际触发的条件格式异常和 13 列结果合流未回归。目标麒麟/UOS 上的 LibreOffice 进程树行为及实际引擎卡死/异常退出仍需设备验收。真实 S1-F1 Umya 公式复制、VLOOKUP、Excel COM 结果和输入哈希证据见上文对应增量；本次未改 OOXML Patch。工作区未提交、未 push，S1-F1 GUI 人工验收和总 Goal 继续 active。

#### 2026-10-29 增量：总 Goal 定向当前源码复核

为核实 Excel 引擎策略与 S1-F3 状态不只依赖旧日志，对当前源码补跑与目标直接相关的测试：

- `cargo test --offline --locked -p excel-reader -- --quiet`：19 个单元测试、2 个集成测试通过，覆盖错误分类、只对未知错误 token 的定点 OOXML 规范化回读、输入/其他值保留和非目标错误不回退。
- Excel-Core 路由定向测试 `fallback_uses_a_fresh_copy_of_the_original`、`umya_primary_rejection_falls_back_to_ooxml`、`both_umya_and_ooxml_unsupported_fail_closed`：各 1 passed；验证从原始副本开始、只在预写拒绝时回退、双方不支持时保护输入并失败关闭。
- `cargo test --offline --locked -p audit-core --lib s1_f3 -- --quiet`：5 passed、1 ignored；ignored 为需外部真实样本/引擎的探针，不计作通过。
- `cargo test --offline --locked -p audit-core --lib conditional_format -- --quiet`：11 passed，覆盖坏规则局部诊断、一个工作表读值失败时保留其他表结果等条件格式行为。

以上均使用 `CARGO_TARGET_DIR=F:\RustBuild\audit`。本轮没有提供 `AUDIT_S1_REAL_*` 输入路径，因此没有重跑真实 Python 工作簿探针；该部分继续以本清单此前记录的真实探针为证据。随后执行当前源码的串行全 workspace 回归：80 个测试目标、**483 passed / 0 failed / 32 ignored**，日志 `F:\RustBuild\audit\workspace-test-excel-engine-s1f3-current-20261029.log`。忽略项包括真实工作簿、手工 parity runner 及外部运行时探针，不计作通过。GUI 人工操作和目标机验收仍未完成，总 Goal 保持 active。

#### 2026-10-29 增量：当前 Release 的真实 S1-F1 与 S1-F3 样本复验

从只读旧 Python 项目选取单位贷款源样本，并使用匹配的 `config-real` 单位贷款模板、S1 映射和外部数据；当前 `dist/lo-convert.exe` 是本 Goal 最近一次 `build-release.bat` 产物。所有生产输入只读，测试创建的工作簿写入系统临时目录，测试前后校验源、模板、外部数据、映射的哈希。

- **S1-F1：** `cargo test --offline --locked -p audit-core --test s1_real_umya_copy -- --ignored --nocapture`，3 passed。包含纯数据 Umya 往返、模板/外部表组件导入及完整 F1 复制/重算。完整探针验证至少 1,000 个模板公式单元格文本、至少 200 个 `参照表!D2:D211` VLOOKUP 公式、共同导入的 `机构参照表`、外部表独立 Excel COM 重算基线、重算后全部公式缓存，以及四个原始文件哈希。当前未传入 Python 审核副本，所以这次没有形成全部模板公式结果的 Python 逐格对拍结论；此前另一份对齐样本的 1,072 结果对拍仍是单独历史证据。日志：`F:\RustBuild\audit\s1f1-real-configreal-source-candidate-20261029.log`。
- **安全拒绝记录：** 首次候选工作簿已含外部导入工作表同名的“集中系统数据”，F1 在写入前拒绝覆盖；同一候选的纯 Umya 往返也报告已有公式变化。未删除安全拒绝或放松候选校验。随后改用不含外部工作表冲突的实际机构报送文件，3 个探针通过。失败候选日志：`F:\RustBuild\audit\s1f1-real-current-goal-20261029.log`。
- **S1-F3：** `cargo test --offline --locked -p audit-core --lib real_s1_f3_recalculates_and_resumes_a_real_s1_f1_audit_copy -- --ignored --nocapture`，1 passed。当前 Release sidecar 通过 Excel COM 重算 1,282 个公式；用新建的持久化 store 调用恢复 worker 后，状态 `COMPLETED`，3 条公式异常，0 条条件格式异常/诊断，13 列结果表可被 Calamine 重开，待恢复队列归零。审核副本哈希和 4 个原始输入哈希不变。日志：`F:\RustBuild\audit\s1f3-real-current-goal-20261029.log`。该探针模拟“Excel 已重算并保存”的副本状态，没有模拟用户按键或真实 app 进程退出/重新启动。

真实 GUI 整批审核与用户手动重算/保存/关闭/重启恢复尚待人工验收；目标 Linux 发行版的 LibreOffice 进程行为尚无实机证据。上述探针补齐当前 Release 和真实工作簿证据，但不关闭总 Goal。

#### 2026-10-29 增量：S1-F1 真实 Umya 路由和跨工作表依赖断言

按最新引擎策略补强真实 S1-F1 探针，不修改生产编辑后端。`audit-core/tests/s1_real_umya_copy.rs` 现在在真实配置和工作簿上检查 `EditExecutionReport`：公式 ImportRange 完整批次唯一选择已验证的 `UmyaSpreadsheet`，只有一个成功尝试；并要求「参照表」D2:D211 的真实 VLOOKUP 中至少一条包含对「机构参照表」的跨 Sheet 引用，避免样本覆盖不足导致测试误通过。探针仍逐条检查复制公式文本、模板公式位置、外部表独立计算结果、工作表及文件可重开、缓存完整性，以及源文件/模板/外部数据/配置哈希。原件只读，输出位于临时目录。

验证结果：真实 S1-F1 ignored 探针 3 passed / 0 failed；测试使用当前源码构建的 Rust debug harness，公式重算使用最近一次 Release `dist/lo-convert.exe` 和 Windows Excel COM。此次没有提供 Python 结果基线，因此不把本次表述为同一运行中的完整 Python 逐公式对拍。日志：`F:\RustBuild\audit\s1f1-umya-route-cross-sheet-current-20261029.log`。`cargo test --offline --locked -p excel-editor --lib -- --quiet` 为 48 passed；公式区域/坐标两个集成测试合计 5 passed；`cargo test --offline --locked -p audit-core --lib s1_formula_copy -- --quiet` 为 15 passed。相应日志见 `F:\RustBuild\audit\excel-editor-engine-policy-current-20261029.log`、`F:\RustBuild\audit\excel-editor-import-formula-current-20261029.log`、`F:\RustBuild\audit\audit-core-s1-formula-copy-current-20261029.log`。

路由边界：Excel-Core 对目标批次要求单一后端覆盖全部操作。真实 S1-F1 包含外部公式工作表导入，OOXML Patch 尚无已实现且验证通过的同语义 ImportRange 能力，因此真实 S1 路由只注册/选中 Umya；不能用公式文本或纯数据 OOXML 路径伪装完整回退。定向单测的 Umya 预写拒绝→OOXML 干净原件回退、两者均不支持时失败关闭及写入后失败不再回退，是通用合成路由协议测试；它们没有将 S1 ImportRange 登记为 OOXML 生产能力。当前已验证的 Umya 业务能力足够，未扩写 OOXML 公式工作表复制。

本次生产逻辑未变，因此未重建 Release；Cargo 定向测试按当前授权执行。未执行 workspace 全量回归、打包、提交或 push。没有文件删除；若回滚，仅撤销该探针断言、当前目标与迁移说明和对应本轮记录。真实 GUI 批次结果、S1-F3 用户手工重算/保存/关闭/重启，以及 Kylin/UOS 目标机仍待人工验收；总 Goal 保持 active。

#### 2026-10-29 增量：Release S1-F1 验收运行器真实样本

为完成总 Goal 授权的当前 Windows Release 构建及正式业务入口复验，执行 `build-release.bat`。主程序与 `lo-convert.exe` 构建成功并复制到 `dist`，Cargo 输出目录为 `F:\RustBuild\audit`。构建无错误，保留 Umya API 弃用和少量未使用代码警告。

随后执行 `dist\open-data-audit.exe acceptance-run`，清单仅请求 S1-F1。源报送表、匹配模板、映射配置和外部表先复制至 `F:\RustBuild\audit\s1f1-release-acceptance-20261029-065200\inputs`；验收状态通过 `AUDIT_STATE_DIR` 隔离到同目录下的 state 文件夹，输出目录为全新 sibling `output`。该 CLI 运行器调用 GUI 使用的 `start_dag_flow`/S1-F1 worker，并非人工点击 GUI。原 Python 项目文件未作为写入目标。

Release 实际运行状态 `completed`，退出码 0，耗时 8,594 ms。1 个真实源文件匹配唯一单位贷款模板，结构标签 406/406、结构通过 1、跳过 0、失败 0、部分 0。生产编辑引擎为 `umya_spreadsheet`，共同导入 4 张外部工作表，准备并由 Excel COM 重算 1,282 个公式；提取公式异常 3 条、条件格式异常 0 条、总异常 3 条，待人工重算 0。清单中所有输入文件哈希均未变化。

最终 `汇总核查表校验.xlsx` 再以 Excel COM 只读打开成功：含 `本期审核结果` 工作表，实际区域 4 行×13 列，表头为预期 13 项。验收报告：`F:\RustBuild\audit\s1f1-release-acceptance-20261029-065200\output\acceptance-report.json` 和 `.md`；最终工作簿位于 `...\output\steps\S1-F01\S1-F1_审核结果_20261029_065218_524\汇总核查表校验.xlsx`。本次未提供 Python 全模板结果基线；该样本条件格式没有触发异常，触发型条件格式仍依靠独立的 Excel COM 合成 E2E 证据。

本轮使用 Cargo 执行 Release 构建，之前同一 Goal 回合也已执行真实 S1 探针及 Excel-Core/S1 定向测试；未执行全 workspace 回归、Debian 打包、Git 提交或 push。未删除文件。留在 `F:\RustBuild\audit` 的输入副本、状态目录、验收 JSON/Markdown 和结果工作簿均为本地人工检查材料。真实 GUI 选择输入/观察输出路径、S1-F3 用户手工重算保存关闭重启，以及 Kylin/UOS 目标机验收仍未完成；总 Goal 保持 active。

#### 2026-10-29 增量：Umya 优先策略复验与 Release 跨进程 S1-F3 恢复

按用户最新引擎策略，本轮先复跑真实 S1-F1 复制工作簿，不扩写 Excel-Core，也不把 OOXML Patch 标成公式 ImportRange 的替代实现。真实工作簿输入取自上一轮隔离副本 `F:\RustBuild\audit\s1f1-release-acceptance-20261029-065200\inputs`；原 Python 项目仍为只读参考。

- Excel-Core 单测：`CARGO_TARGET_DIR=F:\RustBuild\audit cargo test --offline --locked -p excel-editor --lib --quiet`，48 passed / 0 failed。
- 公式区域/坐标集成测试：`cargo test --offline --locked -p excel-editor --test import_range --test formula_copy_translation --quiet`，5 passed / 0 failed。
- S1 公式复制单测：`cargo test --offline --locked -p audit-core --lib s1_formula_copy --quiet`，15 passed / 0 failed。
- 当前源码真实 S1-F1 探针：`cargo test --offline --locked -p audit-core --test s1_real_umya_copy -- --ignored --nocapture`，3 passed / 0 failed。覆盖纯数据源往返、普通公式/跨 Sheet 公式、约 210 个 `参照表!D2:D211` VLOOKUP、共同复制的「机构参照表」、模板公式位置/文本和外部表独立 Excel COM 计算基线、工作簿重开及原始输入保护。生产编辑报告确认真实 ImportRange 批次仅通过 Umya 完成；没有 Python 结果基线，因此本次不声称完整模板结果与 Python 逐格一致。

路由结论：当前 S1-F1 业务已由真实 Umya 样本满足，公式重算/结构验证结果通过；不开发重复的 OOXML 公式工作表复制。OOXML Patch 仍只允许用于它已有且语义完全一致的定点能力；现有通用路由测试仅验证“写入前明确拒绝→从原始副本切换”和失败关闭协议，不代表 OOXML 支持 S1 ImportRange。真实 S1-F1 不注册不存在的 OOXML 后备能力。

另对 S1-F3 做了跨 Release 进程的业务恢复验证。所有输入副本、状态和结果隔离在 `F:\RustBuild\audit\s1f3-process-restart-20261029-065923\`。Release S1-F1 在刻意指定一个存在但不支持 `recalculate` 子命令的 sidecar 后，进入 `WAITING_MANUAL_RECALC` 并保存 1,282 公式审核副本；验收运行器总状态为 `failed`/退出码 2，这是预期故障注入结果，输入哈希保持不变，待恢复清单存在。随后 `lo-convert.exe recalculate --input ... --output ... --engine com` 报告 `Excel.Application` 成功、公式文本保留、1,282/1,282 缓存齐全且结果可读；将其已保存输出写回隔离审核副本以模拟用户在 Excel 中重算保存。第一进程结束后启动新的 Release `serve` 进程，`/api/list_s1_f3_tasks` 读出 `WAITING_MANUAL_RECALC` 任务；经 `/api/start_s1_f3_workflow` 恢复后状态为 `COMPLETED`，提取 3 条公式异常、0 条条件格式异常，生成可由 Excel COM 只读打开的 `本期审核结果`（4 行×13 列），待恢复任务为 0。服务进程随后停止。

本轮 Cargo 定向测试共 71 项通过（48+5+15+3）；另有 Release 跨进程 API 恢复成功。没有修改 Rust 生产代码或 Excel-Core；只更新当前目标与迁移证据。未重新执行 Release build，因为生产代码未变；本轮使用上次已构建的 Release 主程序和 sidecar。没有运行 workspace 全量测试、打包、提交或 push。未删除文件。人工通过 GUI 文件选择、手工按键重算/保存/关闭、点击 F3，以及 Kylin/UOS 实机验收仍待用户完成；总 Goal 继续 active。

#### 2026-10-29 增量：S1-F1 Release GUI 实测与输出路径状态修正

在前一项 Release worker 真实样本记录后，使用实际 Release `serve` 页面通过浏览器选择器重新选入隔离的源目录、模板目录、`mapping.xlsx`、`external.xlsx` 和单独输出目录，点击刷新并执行 F1。该次不触碰旧 Python 源目录，四个 Excel 输入都是已复制到 `F:\RustBuild\audit\s1-final-release-20261029-070522\inputs` 的本地工作副本。

最终 Release GUI 执行结果：1/1 源文件自动匹配模板；表结构 406/406；Umya 导入 4 张外部表并准备 1,282 个公式；Excel COM 完成真实重算；状态 `COMPLETED`；公式异常 3、条件格式异常 0。报告路径 `F:\RustBuild\audit\s1-final-release-20261029-070522\gui-output-ui-retest\S1-F1_审核结果_20261029_073237_123\汇总核查表校验.xlsx`。Excel COM 只读重开确认工作表 `本期审核结果`，4 行×13 列，预期 13 个字段均存在。当前样本未触发条件格式异常，本次没有提供新的 Python 逐公式结果基线；触发型条件格式由已有合成 E2E 覆盖。

运行后源报送表、模板、映射、外部数据 SHA-256 都与同一隔离目录中先前 acceptance manifest 的 `sha256_after` 一致；worker 没有修改输入。页面确认框显示的输出目录与实际结果路径一致。

GUI 复验定位并修复了此前未发现的输出状态问题：`update` API 未接受 `outputPinned`，且普通 `sync()` 会把自动跟随的 `output` 当成用户明确选择的目录。现已让后端管理输入路径变化时的跟随目录、显式输出选择时固定、取消固定时恢复跟随，清空输入/输出时同步状态；前端普通同步不再提交 `output`，只有输出目录选择动作才提交。F1 入口移除过时的“复制能力未验证”标签，执行帮助说明自动重算失败时会保留副本、原始输入不修改；确认框展示实际输出路径及固定/跟随模式。

验证：新增 `web_bridge::tests::s1_output_directory_tracks_input_until_explicitly_pinned` 通过；`rustfmt --edition 2021 --check audit-core/src/web_bridge.rs` 通过。完整 workspace 回归为 80 个测试目标、484 passed / 0 failed / 32 ignored，日志 `F:\RustBuild\audit\workspace-final-goal-ui-output-20261029.log`。最新 `build-release.bat` 成功，Release 日志 `F:\RustBuild\audit\release-final-goal-ui-output-20261029-r2.log`；版本冒烟为 `open-data-audit 0.1.0`，sidecar help 含 `recalculate`。最新 Release GUI 点击流程如上，浏览器 console 无错误。全仓 `cargo fmt --check` 因既有其他源码/样例格式差异未通过；没有全仓格式化无关文件。既有 WSL x86_64 LibreOffice 真重算证据未在本增量中重跑。没有打包、提交或 push，没有更改 OOXML Patch 能力，也未引入 Python 生产依赖。

真实业务 GUI 的当前隔离样本操作已验证；用户机构文件的审核结果解释、S1-F3 用户手动 Ctrl+Alt+F9/保存/关闭/GUI 恢复、麒麟 V10/UOS V20 目标机验收仍待人工完成。Rust-Audit 总 Goal 保持 active。

#### 2026-10-29 增量：按 Umya 优先策略复跑当前源码

按用户本轮明确的引擎策略再次核对：`S1-F1` 的生产编辑路由只注册经过平台验证的 Umya；S1 的跨簿 `ImportRange` 保留公式文本和值，复制候选在外部重算前核对工作表顺序/名称、公式位置与文本、非公式输入值。真实跨簿公式复制没有同语义且已验证的 OOXML Patch 能力，因此不注册虚假后备，也不扩写完整 OOXML 工作表复制。其他编辑只能由现有 Excel-Core 按操作能力证据选择；只有写入前拒绝才会从原件新建副本尝试下一项合格能力，写入后错误或候选验证失败均停止。S1-F1 不调用 openpyxl worker 或 Python 业务代码；Windows 公式重算仍由 PowerShell + Excel COM 执行。

本轮以 Release GUI 验收目录中的隔离样本重新执行当前源码 `s1_real_umya_copy` 忽略探针：**3 passed / 0 failed**。覆盖纯数据工作簿 Umya 往返、真实模板及外部公式导入、约 210 个 `参照表!D2:D211` VLOOKUP 与「机构参照表」依赖、候选重开及公式/值检查、Excel COM 重算，以及外部表、模板片段和完整副本由 Excel COM 打开。探针断言源工作簿、模板、外部数据和映射文件未变。本次未提供 Python 审核结果基线，故不新增 Python 逐公式对拍结论；已有 1,072 个公式结果与 Python 基线 0 差异记录仍作为独立历史证据。

当前源码定向回归：`CARGO_TARGET_DIR=F:\RustBuild\audit cargo test --offline --locked -p excel-editor --quiet` 为 71 passed / 0 failed / 5 ignored；`cargo test --offline --locked -p audit-core --lib s1_formula_copy --quiet` 为 15 passed / 0 failed；真实样本命令 `cargo test --offline --locked -p audit-core --test s1_real_umya_copy -- --ignored --nocapture` 为 3 passed / 0 failed。Excel-Core 警告包含现有 Umya 弃用 API 与未使用代码；本轮未运行 Release 构建、全 workspace 回归或打包，未新增生产代码、未删文件、未提交或 push。用户业务文件语义验收、S1-F3 用户实际手工重算/保存/关闭/GUI 恢复，以及 Kylin/UOS 实机验收仍待完成；总 Goal 保持 active。

#### 2026-10-29 增量：S1-F3、条件格式与读取回退当前源码复验

继续按总 Goal 核验人工重算恢复及其依赖：`cargo test --offline --locked -p audit-core --lib s1_f3 --quiet` 为 5 passed / 1 ignored，`... --lib s1_recovery --quiet` 为 9 passed；使用隔离 GUI 验收样本和当前 Release sidecar 的真实 S1-F3 探针为 1 passed。Excel COM 实际重算 1,282 个公式，正式恢复 worker 输出 3 条公式异常、0 条条件格式异常/诊断，状态 `COMPLETED`；恢复副本及四个输入哈希不变。此探针由 sidecar 模拟“已重算并保存”状态，未模拟用户亲自按键或在 GUI 中手工恢复。

读取与 OOXML 回归：`cargo test --offline --locked -p excel-reader --quiet` 为 19 个单测和 2 个集成测试通过，覆盖特定未知错误 token 的定点 OOXML 兼容重读及非目标错误不回退；`cargo test --offline --locked -p ooxml-patch --quiet` 为 15 passed；`audit-core --lib conditional_format` 为 11 passed，`audit-core --lib s1_conditional` 为 4 passed；另单独运行 `conditional_extraction_failure_keeps_formula_findings_as_partial_result` 与 `s1_f1_batch_publishes_completed_file_and_persists_other_file_for_manual_recalc`，各 1 passed，核实局部条件格式失败保留公式结果且独立文件继续处理。另在 Ubuntu WSL 以隔离 `CARGO_TARGET_DIR=/mnt/f/RustBuild/audit/wsl-linux` 及 `LO_CONVERT_REQUIRE_LO_TEST=1` 执行 `cargo test --offline --locked -p lo-convert -- --quiet`，14 个单测和 1 个实际 LibreOffice CLI 重算集成均通过，确认不是因 soffice 缺失而跳过。本轮没有跑 workspace 全量回归或 Release 构建。测试仍显示现有 Umya 弃用 API 和少量未使用代码警告，无失败。S1-F3 用户实际 Excel/WPS/LibreOffice 手工重算、保存、关闭后 GUI 恢复及程序重启流程，以及 Kylin/UOS 实机验收未完成；总 Goal 保持 active。

#### 2026-10-29 增量：Release S1-F1 Excel COM 合成端到端复验

按当前 Goal 再次运行 `build-release.bat`，成功生成并复制 Release 主程序和 `dist/lo-convert.exe`；`open-data-audit.exe --version` 输出 `0.1.0`，sidecar `--help` 显示 `convert` 与 `recalculate` 入口。随后执行：

`AUDIT_S1_RECALC_TEST_SIDECAR=dist/lo-convert.exe cargo test -p audit-core --test s1_audit_directory_cli --release s1_directory_copies_template_formulas_recalculates_and_combines_real_issues -- --ignored --nocapture`

结果 **1 passed / 0 failed**，由 Windows Excel COM 对隔离合成审核副本真实重算。测试逐项检查公式异常缓存（`D1=#DIV/0!`、`D2=#REF!` 及一条规则异常）、源值为 5 且实际触发的条件格式异常、合并后的 13 列结果可重开，及源文件、模板、重算前审核副本的哈希不变。测试 sidecar 来自当前 `dist` Release；Cargo 为集成测试构建的主程序来自当前源码。构建和测试输出包含既有 Umya 弃用 API、未使用导入/变量/函数警告，无编译错误。本次没有生产代码变更、没有重打包、提交或 push；此合成用例不替代用户机构文件的业务解释、S1-F3 手工恢复及 Kylin/UOS 设备验收，总 Goal 继续 active。

#### 2026-10-29 增量：Excel COM 异步连接失败关闭与 Goal 完整复验

静态复核发现 Windows Excel COM 脚本在工作簿存在连接时会调用 `CalculateUntilAsyncQueriesDone()`，但旧代码吞掉该调用异常，随后仍可能因 `CalculationState == xlDone` 而保存。现改为记录 `async_query_wait_failed` 阶段并抛出明确错误；外层失败处理会关闭工作簿、退出 Excel，并由 sidecar 按已记录 PID 清理 Excel 进程，避免把异步连接等待失败报告为成功。新增 `windows_com_does_not_save_when_async_query_wait_fails`，断言等待失败发生在任何保存之前。真实 S1 业务样本不含异步连接，因此此分支以脚本生成回归验证，没有真实外部连接工作簿集成样本。

本轮当前源码验证全部使用 `F:\RustBuild\audit`（WSL 目标位于其 `wsl-linux` 子目录）：

- `cargo fmt --manifest-path sidecar/lo-convert/Cargo.toml -- --check`：通过。
- `cargo test --offline --locked --workspace -j 1 -- --quiet`：80 个测试目标，**485 passed / 0 failed / 32 ignored**；包含 sidecar 15 个单测。日志：`F:\RustBuild\audit\workspace-excel-engine-s1f3-completion-20261029.log`。
- `cargo test --offline --locked -p lo-convert windows_com_does_not_save_when_async_query_wait_fails -- --nocapture`：1 passed。
- 真实 S1-F1：`s1_real_umya_copy` 忽略探针 **3 passed**，由当前源码 debug 测试目标执行、使用最新 Release sidecar/Excel COM。覆盖隔离真实源、模板、外部数据、映射；纯数据往返、约 210 条 VLOOKUP/跨 Sheet 引用、公式缓存与输入哈希。结果没有提供 Python 基线，因此不扩展为本轮逐公式 Python 对拍结论。
- Windows Release `build-release.bat`：成功，输出日志 `F:\RustBuild\audit\release-excel-engine-s1f3-completion-20261029.log`。新 Release sidecar 的 S1-F1 Excel COM 合成 E2E **1 passed**；新 Release 测试目标对真实 S1-F3 执行 **1 passed**，COM 重算 1,282 个公式，恢复状态 `COMPLETED`、3 条公式异常、0 条条件格式异常/诊断，审核副本和四个原始输入哈希未变。
- 在 Ubuntu WSL 执行 `CARGO_TARGET_DIR=/mnt/f/RustBuild/audit/wsl-linux LO_CONVERT_REQUIRE_LO_TEST=1 cargo test --offline --locked -p lo-convert -- --quiet`：**15 单测和 1 个实际 LibreOffice CLI 重算集成通过**。确认使用真实 `soffice`，非缺少软件时的跳过路径。
- 另以隔离 `AUDIT_STATE_DIR` 启动当前 Release GUI 服务，读取真实 `get_state` API：空任务 0、等待恢复任务 1；停止该 Release 进程并重启后任务仍为 1，`s1F3RecoveryError` 为空。测试使用隔离状态目录中的清单副本，不改原验收清单或审核工作簿。这验证跨进程持久状态与入口数据源；未点击 F3 恢复按钮，也未模拟用户手工按键。

因此，代码级 S1-F3 worker、Release 公式重算与进程重启状态读取均有当前证据。尚未进行用户在 Excel/WPS/LibreOffice 中手工强制重算、保存、关闭后亲自点击 GUI F3 的操作验收；麒麟 V10/UOS V20 设备验收也仍缺少实机证据。Goal 保持 active；未打包 ZIP/deb，未提交或 push，未修改 Python、旧 Main、Rust-Tools 或 RustHistory。

#### 2026-10-29 增量：当前 Release 的 S1-F3 GUI 恢复与完成态隐藏复验

在前述 Release 状态 API/重启探针之后，使用当前 `dist/open-data-audit.exe` 启动隔离 `AUDIT_STATE_DIR`，从真实 S1 页面检查并点击待恢复任务。任务以 `WAITING_MANUAL_RECALC` 显示，恢复面板列出工作簿并显示强制重算指引；确认后运行 S1-F3，生成核查表。当前 Release 结果目录为 `F:\RustBuild\audit\s1f3-current-release-gui-e2e-20261029\output\S1-F1_审核结果_20261029_073237_123`；结果文件为 `汇总核查表校验_恢复_d101d62801cb0aa95e3e797966aa1a70_9f619e8372e1626116a9f3344d2936c7.xlsx`。状态变为 `COMPLETED`，3 条公式异常、0 条条件格式异常，待处理数归零；审核副本 SHA-256 与其复制来源一致。没有改动原始报送文件。

首次探测使用的隔离清单文件名不符合恢复器按 `run_id` 查找文件的约定；在隔离目录中改为规范文件名后重试成功，该失败属于测试夹具问题，不是生产恢复失败。

本次 GUI 操作还暴露了所有任务完成后按钮仍显示“待处理：0”的前端缺陷：通用 `.btn` 显示规则覆盖 `hidden` 属性。现加入 `#s1F3RecoveryButton[hidden]{display:none!important}`，并在 `served_page_is_original_markup_plus_runtime_bridge` 加入资源断言；该测试在当前源码通过 1/1。基于修复后源码重新运行 `build-release.bat` 成功；新的 Release 在隔离待恢复状态显示 1 条任务，完成态中 S1-F3 按钮从可访问界面和截图消失。构建用时约 6 分钟，只有既有依赖弃用与未使用代码警告。

当前源码（包含 CSS 修复）全工作区复回归：80 个测试目标，**485 passed / 0 failed / 32 ignored**，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-current-20261029-r4.log`。最初复跑时，`sheet_index_metadata` 中两个删除工作表用例偶发读取到 `incomplete deflate stream`；两项单独运行、该测试目标并发运行及后续全工作区运行均通过。为避免时间戳目录的极低碰撞概率，测试夹具已改为 PID + 原子序号，并以独占创建确认目录唯一；修改后定向 3 项通过，最终完整回归也通过。未改生产 Excel 编辑逻辑。新增测试文件 `rustfmt --check` 通过；全仓 `cargo fmt --all -- --check` 因其他文件已有格式差异仍失败，本轮没有改格式无关文件。此前 Release GUI、Windows COM 合成 E2E、真实 Umya 样本、Ubuntu WSL LibreOffice 证据继续有效；本轮仅改测试隔离与记录，没有重建 Release、重打 ZIP、提交或 push。用户亲自强制重算后由 GUI 恢复及麒麟 V10/UOS V20 实机验收仍未完成；总 Goal 保持 active。

#### 2026-10-29 增量：Excel 引擎与 S1-F3 总 Goal 完成度核对

依据当前源码、最新全工作区日志、当前 Release 文件和已有真实样本记录，逐项核对 Goal 要求后更新 `DEVELOPMENT_GOAL.md` 的完成度清单：Calamine 定点兼容、Umya/OOXML 按能力路由、COM/LibreOffice 重算与缓存检查、进程清理、条件格式部分结果及 S1-F3 worker/UI 均有对应实现和测试；S1-F1 公式 ImportRange 只使用真实样本已验证的 Umya，没有给 OOXML Patch 或 Python worker 配置未经证实的回退。当前 Windows `dist` 未包含 Python/openpyxl runtime，S1-F1 生产路由也没有注册 Python backend。

当前 Release 实际冒烟：`dist/open-data-audit.exe --version` 输出 `0.1.0`；`dist/lo-convert.exe --help` 显示 `recalculate --engine auto|lo|com`。`open-data-audit.exe` 文件时间为 08:32，晚于嵌入 GUI CSS 的 08:25 修改；`lo-convert.exe` 为 08:12，晚于 sidecar 源码的 08:06 修改。当前只有测试隔离和文档变化，因此未重复 Release 构建。Goal 未达到完成状态：真实 Linux 机构工作簿/麒麟 V10/UOS V20 目标环境、用户亲自在 Excel/WPS/LibreOffice 强制重算保存关闭后通过 GUI 恢复，以及用户对机构审核结果的签收仍缺少直接证据。合成重算、COM 模拟人工结果、Release GUI 自动点击均不能替代这些验收。

#### 2026-10-29 增量：Ubuntu WSL LibreOffice 当前源码实测

在 `C:\Users\AZI\Desktop\VibeCode\Rust-Audit` 运行 `wsl.exe -d Ubuntu -- bash -lc "cd /mnt/c/Users/AZI/Desktop/VibeCode/Rust-Audit && CARGO_TARGET_DIR=/mnt/f/RustBuild/audit/wsl-linux LO_CONVERT_REQUIRE_LO_TEST=1 cargo test --offline --locked -p lo-convert -- --quiet"`，退出码 **0**。结果：15 个 sidecar 单测通过，`recalculate_libreoffice` CLI 集成测试 1 项通过；该集成实际启动 `soffice`，验证 VLOOKUP、跨表公式、除零错误值、空字符串结果、公式文本/缓存节点及工作表结构。它使用合成工作簿，不等同于真实 S1 机构工作簿的 LibreOffice 验收。首次捕获日志的外层状态包装有 quoting 错误；之后直接运行获得上述 Cargo 退出码 0 与测试结果。输出物全部位于 `F:\RustBuild\audit\wsl-linux`。

#### 2026-10-29 增量：S1-F1 真实 LibreOffice 错误常量序列化兼容

首次对真实 S1 外部工作簿运行 Linux 重算时，错误发生在独立外部表基线阶段，未进入 Umya 跨簿复制。保留原输入副本并用 LibreOffice 原始输出作 OOXML 检查：`参照表!D2:D211` 的 210 条 VLOOKUP 公式文本均保持一致，但 LibreOffice 增加 `aca="false"` 属性；源文件 `参照表!C211` 是常量错误值 `#REF!`，LibreOffice 将其序列化成值同为 `#REF!` 的公式。这是计算引擎序列化差异，不能把公式文本改写当成可接受回退。

新增 `ooxml_patch::restore_promoted_error_constants`，仅在原输入的目标单元格是 XML 错误常量、LibreOffice/Excel 候选在同一坐标包含精确匹配的受支持错误公式与相同缓存错误、公式属性不超出 `aca="false"` 时，从独立新副本移除该公式节点并保留错误值。转换后仍运行严格公式文本/位置、工作表、缓存节点及类型化结果校验。对于其他新增/更改公式、引用变化或错误结果不一致，不会回退或发布；OOXML Patch 没有因此获得跨簿公式复制能力。

定向回归：Windows `cargo test -p ooxml-patch promoted_error_constant_tests` 为 2 passed；Windows `cargo test -p lo-convert recalculation_tests` 为 17 passed。Ubuntu WSL 以 `LO_CONVERT_REQUIRE_LO_TEST=1 cargo test --offline --locked -p lo-convert -- --quiet` 为 17 单测 + 1 LibreOffice CLI 集成通过；新增合成集成验证 literal `#REF!` 经重算仍恢复成错误常量，同时保留 VLOOKUP、跨表公式、DIV/0、空字符串和公式结果可读性。完整真实 S1-F1 WSL LibreOffice 探针 1/1 通过，覆盖 Umya 公式/数据复制、约 210 条 VLOOKUP、跨表依赖、重算结果对照、输出可重开和四个输入文件保护；日志 `F:\RustBuild\audit\s1-real-linux-lo-probe-20261029-r3.log`。最新 Windows Release 由 `build-release.bat` 构建成功，版本冒烟和 sidecar 帮助冒烟通过；该 Release sidecar 使用 Excel COM 对真实样本探针 1/1 通过，日志 `F:\RustBuild\audit\s1-real-release-com-error-restore-20261029.log`、构建日志 `F:\RustBuild\audit\release-s1f1-error-constant-normalization-20261029.log`。修改的三个 Rust 文件 `rustfmt --check` 通过；未运行全工作区回归。用户业务验收、S1-F3 用户手工恢复及 Kylin/UOS 设备测试仍是未完成状态，Goal 保持 active。

#### 2026-10-29 增量：引擎回退契约与全工作区回归

按当前 Umya 主引擎/OOXML Patch 按能力使用的策略，复核 S1-F1 的 `ImportRange` 生产路由：真实业务复制由 Umya 执行；OOXML Patch 没有可证明与该跨簿公式复制语义一致的能力，因此不注册其为该操作的回退，Umya 失败仍失败关闭。通用 ExcelModifier 只在后端于写入前明确拒绝且另一个已登记后端具有同操作能力时，才从原始输入副本重试；写入后错误或候选校验失败不会自动回退。

定向 Excel-Core 路由测试通过：Umya 明确拒绝后从全新原始副本使用 OOXML Patch 的合成回退 1/1、双引擎均不支持时失败关闭并保护原文件 1/1、候选仅写入验证通过后发布且清理临时尝试 1/1。连同真实 S1-F1 探针，Umya 已验证纯数据、普通公式、跨 Sheet 公式、约 210 个 VLOOKUP、同时复制「参照表」与「机构参照表」、工作表/关键结构、保存重开、Excel COM/LibreOffice 重算和源文件哈希保护。无证据表明 OOXML Patch 可以复制含公式的外部工作表，不扩大该能力，也未增加 openpyxl/Python 生产回退。

Windows 当前源码全工作区回归 `CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -j 1 -- --quiet` 退出码 0：80 个测试目标，**489 passed / 0 failed / 32 ignored**；日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-final-20261029-r5.log`。既有 release build 和真实 Windows Excel COM/WSL LibreOffice 样本证据见上节。尚待用户对机构业务结论签收、亲自在 Office 中强制重算并通过 S1-F3 恢复，以及 Kylin V10/UOS V20 实机验收；不将自动化探针或合成测试计作这三项人工验收。

#### 2026-10-29 增量：当前源码真实样本 S1-F3 worker 复验

全工作区回归中的 `real_s1_f3_recalculates_and_resumes_a_real_s1_f1_audit_copy` 因其 `#[ignore]` 和所需路径未预设，没有被计入 489 个通过项。之后从 `F:\RustBuild\audit\s1f3-process-restart-20261029-065923\inputs` 取隔离真实 S1-F1 源文件、模板、外部数据和映射，显式设置 `AUDIT_S1_REAL_*`，指定当前 `dist\lo-convert.exe` 并以 `cargo test --release ... -- --ignored --nocapture` 在当前源码下运行；结果 1/1 通过、退出码 0。S1-F1 新建了 1,282 个公式的审核副本，sidecar 用 Excel COM 重算独立候选并写回隔离审核副本；重建 Store 后，真实 S1-F3 worker 成功恢复，3 条公式异常、0 条条件格式异常、0 诊断、`COMPLETED`，结果工作簿包含预期 13 列，源、模板、外部表和映射哈希保持不变。日志 `F:\RustBuild\audit\s1-f3-real-worker-current-20261029.log`。再次运行当前 `dist` 冒烟，`open-data-audit.exe --version` 与 `lo-convert.exe --help` 均退出码 0；两个 Release 文件生成时间晚于相关源码修改时间。此处不是人工重算证明：用户本人用 Excel/WPS/LibreOffice 强制重算、保存关闭后通过 GUI 恢复仍未执行；机构业务签收和 Kylin V10/UOS V20 实机验收也未完成。

### S1-F1 COM 最终性能与批量稳定性验收（2026-10-30）

- 使用最终 Release 二进制在 Windows x64 对合成 1,200 行工作簿批测，每册 6,000 个 VLOOKUP 公式、7,200 个非空缓存值；Python DispatchEx 与 Rust 原生 COM worker 交错运行。Python/Rust 批次中位数：2 册 7,570.60/7,171.91 ms（Rust 快 5.27%）、10 册 12,246.33/12,139.69 ms（快 0.87%）、33 册 27,006.11/27,138.94 ms（慢 0.49%）。各批只创建一个专用 Excel；COM 激活中位数 2,524/2,391 ms，Rust worker spawn 和 ready IPC 约 43 ms/批。
- 端到端计时等待专属 Excel.EXE 实际退出；Python/Rust 进程退出等待中位数为 2,738/2,741 ms。此前约 0.2 秒只是 COM 释放，并不等于 Excel 进程已关闭。阶段与总时长逐批核对，未归因差额中位数 3.06 ms、最大绝对值 63.81 ms。最终 A/B 二进制 SHA-256 与最后 Release 构建匹配。
- 每个引擎 9 个批次成功输出 96 册，损坏输入失败后仍继续处理；公式结果配对差异为 0。96 对输出逐册核验 6,000 个公式结果和 7,200 个缓存值；输出公式/结果哈希一致，原始 SHA-256 不变。15 个基线 Excel PID 保留；18 个批次自有 Excel 和 worker 均正常退出，没有强制终止或 worker 孤儿。最终系统快照另有一个非本 harness 所有的隐藏 Excel PID 12264，随后查询时已自行退出；期间还观察到短暂的其他隐藏 Office COM 进程，未清理也不宣称由本轮创建。
- 正常 COM 会话测试与受控 35 秒 parent-timeout 清理测试分别 1/1 通过。timeout 用例是在公式重算成功后人为延迟 worker 响应，验证 worker/准确归属的 Excel 清理、源文件保护和不发布输出；并非 Excel 内部 COM 永久挂起模拟。
- 当前源码全 workspace 回归：82 个测试目标、501 passed / 0 failed / 34 ignored；Release open-data-audit 构建成功；Rust 修改文件 rustfmt --check 与 Python harness py_compile 通过。ignored 项新增 1 个受控 COM timeout 测试，已单独实际运行。没有运行 push、打包或 Linux 实测。
- **结论：S1-F1 Windows COM 性能优化已完成。** Excel 冷启动是两端共同主成本；Rust 约 43 ms/批的 worker 启动/ready 开销不值得牺牲进程隔离。性能数据视为平价；合成验收不能替代真实机构输入、GUI 完整链路或业务审核结论人工验收。报告与原始数据：probe/reports/s1f1_com_final_acceptance_20261030.md/.json；批测脚本 tests/compare/s1f1_final_acceptance_20261030.py；全 workspace 日志 F:\RustBuild\audit\s1f1_final_workspace_regression_20261030.log。

### 2026-10-30 增量：S2-F1 外部规则性能剖析

在只读参考 Python Audit 上调用 `PeriodAuditV2Service`，Rust 使用既有 Release CLI；输入来自隔离的 S2-F1 外部规则人工验收样本，正常和异常值各预热一次并交错运行三轮。输入包括两期源工作簿（各 6,033 bytes、2 张表）、6 表配置（8,656 bytes）及 3 表外部数据（6,164/6,161 bytes）。输出两端均为 3 张分类工作表、292 个比较单元格。

Python 业务服务中位数 41.92/42.19 ms；Rust CLI 墙钟 50.26/51.56 ms，CLI 内部阶段约 24 ms。Python/Rust 倍率依次为 CLI 口径 0.83×/0.82×，服务口径 1.75×/1.76×。Rust 各阶段计时显示配置 3 ms、外部簿读取 1 ms、导出 2 ms、保存/回读验证 15 ms；Python 配置检查与读取约 29 ms、外部簿读取约 3–4 ms、保存约 4 ms。独立 CLI 的启动开销抹平了服务内优势；保存验证虽较长，但总服务仍快，不能据此删减验证。

两种情形各 3/3 轮摘要与输出逐格/结构对拍通过；无业务差异，只有每个输出中 4 个 Windows 文件路径的 `\\?\` 前缀表示差异。外部值、0/1,233,333.39 差值及硬性异常级别符合样本预期；源文件哈希及旧 Python 核心文件哈希均未改变。单独进程采样峰值工作集 Python 40.69/40.71 MB、Rust 13.09/12.38 MB（含进程启动）。

详细计时和负载见 `audit-core/probe/reports/s2f1_external_profile_20261030.md/.json`，复跑脚本为 `audit-core/tests/compare/perf_s2f1_external_20261030.py`。该结果未触发 S2 代码修改；S2-F1 当前已具备一个有外部数据读取和规则输出对拍的定点性能基线，仍不代表 21+21 份生产报表的 A/B。

### 2026-10-30 增量：S3-F4 本期核对值级性能复测

复用本期值核对历史性能基准的 10,000 行 CSV 和 3.0/3.2 配置，含两条本期核对规则。Python `run_cross_period_check` 服务 7,071.61 ms；Rust CLI 墙钟 989.30 ms，服务内部 945 ms（Python/Rust 7.15×/7.48×）。Rust 配置、输入读取、计算、写出/验证中位数分别 1、84、157、702 ms。两边各输出 20,000 条结果，工作簿 20,001×17；预热及三轮逐格/公式/关键结构比较每轮 340,017 个格，差异均为 0。输出约 1.39 MB/1.01 MB。输入与 Python 服务源文件哈希均未变化。

单独进程树采样峰值工作集 Python 181.42 MB、Rust CLI 94.49 MB；两端各一进程。写出验证为最大 Rust 阶段（约 74%），但当前实现整体约 1 秒且比 Python 快逾 7 倍，没有看到值得优先承担实现复杂度的改进点。

数据和限制见 `audit-core/probe/reports/s3f4_parity_profile_20261030.md/.json`，复跑入口 `audit-core/tests/compare/perf_s3f4_parity_20261030.py`。该结果是合成配置的完整输出对照，不替代机构真实数据验收。

### 2026-10-30 增量：S3-F5 有效渲染基准纠错与复测

检查此前性能夹具时发现，2026-10-29/30 两个 S3-F5 生成器均用 `AF001`…`AF004` 作为表单工作表名。实际业务按 `A1/A2` 标识推断人民币/外币业务类，因此这些负载没有匹配业务记录，Rust 报告为 `matched_indicators=0`，所有结果表为空；此前只比较输出文件数量，速度数据不具业务代表性。机器数据保留，但在性能历史中标记无效。修正当前 `perf_compare_20261030.py` 配置生成器，并为聚焦剖析器加入“每表单指标行必须命中”和“回填数值格必须非零”断言；没有修改 S3-F5 产品代码。

有效样本为 30 家机构 × 40 指标 × 4 张 A1 表，共 1,200 行比较输入与 4,800 个表单指标行。Python/Rust 服务中位数 3,509.17/1,510 ms；Rust CLI 墙钟 1,538.36 ms，Python/Rust 分别快 2.32×（服务）及 2.28×（CLI）。Rust 读表 36 ms、规划 51 ms、机构文件编辑中位数每份 45 ms、编辑阶段合计 1,449 ms。每轮生成同名 30 册、6 张表/册；三轮逐单元格值、公式、关键结构和隐藏行对拍为 0 差异（每轮 35,460 格、19,200 数值格）。输入配置哈希未变。

单独进程树采样峰值工作集约 Python 475.34 MB（13 个进程）、Rust 24.25 MB（1 个进程）；S3-F5 Rust 运行仍约 2.3× 快，故本轮未并行化单册写入。详细报告 `audit-core/probe/reports/s3f5_parity_profile_20261030.md/.json`，复跑驱动 `audit-core/tests/compare/perf_s3f5_parity_20261030.py`。当前两个 benchmark 脚本已 py_compile 且实际通过三轮运行和输出对拍；未执行 Cargo。

### 2026-10-30 增量：S3-F3 原地计划内存传递优化

检查 S3-F3 原地写回计时时发现，执行链会将已构造且已校验的 ExcelEditPlan 序列化到暂存 JSON、同步磁盘，随后立即读回；此计划文件不属于用户输出。`audit-core/src/s3_explanation_feedback.rs` 现将公共计划 API 与内部执行规划拆开：显式计划导出仍持久化原 JSON；原地更新及保留的副本入口从内存取得同一份计划并交给 Excel-Core，省去无用磁盘往返。匹配、冲突、输入哈希保护、备份、原子发布及工作簿修改语义未变。旧记录中 S3-F3 小规模略慢的结论被本次同负载优化后结果取代，原值仍保留在历史比较表中。

同一业务数据为 1,000 条唯一反馈、无冲突。Rust 内部中位数由 946 ms 降至 203 ms（78.54%），CLI 墙钟由 975.25 ms 降至 235.54 ms（75.85%）；优化后 Python/Rust CLI 为 912.62/235.54 ms。三轮完整值、公式计数、关键结构和输入保护比较均为 0 差异；只改 H3:H1002；备份与原目标 SHA-256 一致。优化前后 XLSX 容器大小因生成元数据各相差 1 byte，数据布局与业务数据相同。峰值工作集约 38.86/21.73 MB。

定向验证通过：`audit-core` 库测试 7/7（含执行冲突清单及目标保护用例）、S3-F3 计划 CLI 集成测试 3/3；Release `open-data-audit` 构建成功，二进制 SHA-256 `d21e355a8ec05129e81ccdda008f257a98702b64537010c24b773433b416db66`。基准驱动语法及优化前后 JSON 解析检查通过。未运行全工作区回归；`cargo fmt --check --all` 因多处既有格式差异失败，未格式化无关文件。详细数据见 `audit-core/probe/reports/s3f3_parity_profile_20261030.md/.json` 与 `audit-core/probe/reports/s3f3_parity_optimized_20261030.md/.json`；复跑脚本为 `audit-core/tests/compare/perf_s3f3_parity_20261030.py`。

当前无冲突与全冲突 1,000 键合成样本均已记录；实际机构规模及 GUI 常驻服务负载仍缺。变更留在 `main` 工作区，HEAD `7517a9c`，未提交、未 push；S1/S2/S3 持续性能目标保持进行中。

### 2026-10-30 增量：S3-F3 全冲突输出剖析与阶段可观测性

为覆盖原无冲突样本未经过的冲突分支，新增 1,000 匹配键、两份反馈原因相互冲突的 Python/Rust A/B。三轮均输出 1,000 条冲突、0 导入、0 未匹配；冲突清单及目标值/结构对拍 0 差异。目标原字节未改，备份与目标 SHA-256 一致，反馈输入哈希不变。Python 服务中位数 1,332.49 ms、Rust CLI 墙钟 149.89 ms（约 8.89×），Rust 内部 120 ms；冲突清单写出为 18–21 ms。进程树峰值工作集约 Python 41.04 MB、Rust 16.39 MB。该全冲突比例仅是压力测试，不代表生产数据分布。

`S3ExplanationFeedbackExecutionReport` 增加 `conflict_write_time_ms`，并由现有 CLI JSON 序列化输出，使冲突表写出阶段可以从执行结果观测。S3-F3 定向库测试 7/7、计划 CLI 集成测试 3/3 通过，Release 构建成功（SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`）。未跑全工作区回归；性能目标仍在进行中。报告、逐轮数据和脚本见 `audit-core/probe/reports/s3f3_conflict_profile_20261030.md/.json`、`audit-core/tests/compare/perf_s3f3_conflict_profile_20261030.py`。

### 2026-10-30 增量：S1-F1 条件格式指标定位优化

旧条件格式微基准 `s1_conditional_profile_20261030` 没有把模板结构区域传给任一实现，因此其中 1,343/246 ms（5.46×）未覆盖生产异常字段 `check_field` 的指标定位。本次保留该历史报告并在性能基线中标明此范围，不把轻量路径结果当作生产结构区域性能。

新专项使用 4,000 行、32,008 个非空值、4 条条件格式规则、16,000 个扫描格及 `S1Conditional!A1:H1` 结构区域，五轮交错得到 8,610 条异常。旧 Rust 每个触发项都从结构起点重复扫描同列全部上方行，异常组装中位数达 15,523 ms。现按工作表/列分组并单次累积有序去重的上方指标标签，保留左侧标签、合并锚点、值过滤、字段格式和原触发顺序。没有改 Python 或审核规则。

同一驱动的优化前/后 Rust 异常组装为 15,523/57 ms；结构预检+条件格式 Rust developer event 耗时（不含启动）为 16,044/498 ms。优化后 Python/Rust 为 19,264.7/498 ms（38.68×，两端均为内部事件计时（不含进程启动），倍率只用于该专项参考）。Rust 采样峰值工作集约 86.61/86.56 MB。五轮每轮异常数均为 8,610，规范化字段摘要 5/5 对拍一致；结构标签 8/8，源/模板哈希不变。

定向库测试 `cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture` 为 5/5 通过；Release 构建成功，二进制 SHA-256 `d7c49818f3905a5f31da7a351cfd8d44e5bc20fb3cfe1e8b76992bba6b536e56`。没有运行全 workspace、真实机构 GUI 或 Linux 测试。完整链路耗时仍未知；结构预检和条件审核分别重复读取源工作簿约 99/97 ms，是否复用整簿数据待完整多文件链路内存测量后再决定。报告、前后逐轮数据与驱动见 `PERFORMANCE_BASELINE.md`、`PERFORMANCE_HISTORY.md`、`probe/reports/s1_indicator_resolution_optimization_20261030.md` 和 `tests/compare/perf_s1_preflight_reads_20261030.py`。
### 2026-10-30 增量：S1-F1 结构预检和条件审核复用源读取

在 S1-F1 单工作簿作用域内，结构预检和条件格式异常提取现在复用一次读取所得的 `WorkbookDump` 与读取诊断；工作簿投影在后续公式副本/重算开始前释放。保留既有文件路径 API，生产目录和单文件 CLI 路径接入复用实现；新增复用路径结果一致性测试。没有跨多个输入文件驻留内存缓存。

相同生成器的 4,000 行结构区域样本进行了预热 1 轮、交错 5 轮：Rust Release CLI 中位数由前态 498 ms 降到 390 ms（省 108 ms，21.7%），源文件解析由两次约 99+97 ms 降为一次 93 ms。每轮 8,610 条异常摘要与 Python 完全相同、结构标签 8/8 匹配，源及模板运行前后哈希不变；Rust 峰值工作集 85.949 MB，与前态 86.562 MB 相当。前后生成 XLSX 容器相差 1 byte，因此仅声明相同生成器、值和结构，不声称前后字节哈希相同。

定向 S1 conditional、structure compare、conditional format 测试及 CLI 编译目标均通过，Release 构建成功。完整 GUI 审核链、真实机构输入、全 workspace 回归、Linux 目标机仍未在本轮验证；不替代人工业务验收。详见 `PERFORMANCE_BASELINE.md`、`PERFORMANCE_HISTORY.md`、`probe/reports/s1_source_read_reuse_optimization_20261030.md`、`probe/reports/s1_preflight_reads_20261030_shared_read_after.json` 与 `tests/compare/perf_s1_preflight_reads_20261030.py`。

### 2026-10-30 增量：S1-F1/F2 模板推荐元数据批次复用

S1-F1 worker、S1 源文件识别界面和 S1-F2 多文件自动模板识别原先在每个源文件循环中重复扫描模板目录、读取每个模板的 OOXML Sheet 目录。现在每次多文件操作只加载一次模板档案；源文件仍独立读取并逐本匹配，单文件公开推荐 API 保持兼容。Release-profile 合成对照为 21 源簿 × 5 模板簿（源簿各 8 Sheet、模板簿各 9 Sheet，输入总计 229,831 bytes，0 个数据格），模板推荐阶段中位数由 30.2814 ms 降至 5.9081 ms，约省 24.3733 ms / 5.13×；21/21 完整匹配建议完全一致。S1-F1 三文件 GUI 合成路径测试 1/1 通过，Release `open-data-audit` 构建成功，SHA-256 `7be5fdece2ae26e931feb796daa7075395ab3fa44591cc242ba0813beba96792`。本机没有可用的自动重算引擎，GUI 测试验证的是人工重算降级路径；没有全 workspace 或 Linux 回归。源码 HEAD `7517a9c`、工作区未提交，未 push。详见 `PERFORMANCE_BASELINE.md`、`PERFORMANCE_HISTORY.md`、`probe/reports/s1_template_profile_reuse_optimization_20261030.md/.json` 和 `tests/compare/perf_s1_template_profile_reuse_20261030.py`。

### 2026-10-30 增量：S1-F1 模板结构预检投影读取

结构预检此前用完整 `read_workbook` 构造所有单元格 JSON/BTreeMap，并解析不参与结构比对的批注、条件格式和数据验证元数据。现新增专用 Calamine 结构投影，仅保留映射区域中的值/公式、全部 Sheet 名及合并区域；所有 Sheet 的 Calamine 值读取仍执行，以维持失败行为。源结构坐标落入合并区域时，投影也读取左上锚点。Calamine 遇未知错误值时继续走现有按错误单元格 OOXML 临时规范化及重试路径；条件格式审核仍使用原完整读取器。

20,000 行 Rust-only 夹具由同一生成器生成，源/模板各 716,663 bytes、1 Sheet、160,008 非空格、4 条条件格式规则、80,000 扫描格。前后容器哈希不同，故不声称输入逐字节一致。模板预检读取中位数 528→224 ms，减少 304 ms / 57.6%；条件审核命令内部 2,165→1,654 ms、CLI 3,254.905→2,645.195 ms，但总时长变化不能全归因于本次变更。43,051 条异常逐轮规范化摘要与旧/新三轮完全相同，结构标签 8/8，输入运行前后未变。峰值工作集 385.672→385.824 MiB，基本不变，后续条件审核全量读取主导。

定向结果：`excel-reader` 全量/投影值、公式、Sheet 和合并区对拍、合并锚点、未知错误值回退共 3/3 通过；S1 结构比较库测试 7/7 通过；Release 构建通过，SHA-256 `a2ac5bc3fd16223a3c18976f9ed6042e941af6df17fc2fbd84ad5fe54b82f370`。一次未限定并行的 Cargo 命令因 Windows 页面文件不足失败；随后使用单任务并仅编译库测试通过。未运行全工作区回归、真实业务文件、GUI 完整链或 Linux/Kylin/UOS。详见 `PERFORMANCE_HISTORY.md` 与 `probe/reports/s1_structure_projection_optimization_20261030.md/.json`。

### 2026-10-30 Codex 增量：S1-F1 源结构预检投影同输入性能对照

新增 `tests/s1_structure_projection_source_perf.rs` 作为忽略式、显式运行的性能测试，以及 `tests/compare/perf_s1_structure_source_projection_20261030.py` 复跑驱动。Release 同进程、同输入交替对照完整 `read_workbook`+结构比对与 Calamine 投影路径，20,000 行数据、160,008 非空格；预热 1 次后 5 轮中位数为 501.532→237.579 ms（2.111× 子步骤加速）。两路径每轮匹配状态、分数、标签数、缺失表及诊断完全相同，均为 8/8 标签命中；源/模板各自 SHA-256 前后不变。此为 Rust 内部路径对照，不是 Python/Rust 或完整 S1-F1 性能结果；没有独立峰值 RSS、真实文件、GUI、全工作区或 Linux 回归数据。未修改产品业务代码。完整原始数据和测试边界见 `probe/reports/s1_structure_source_projection_20261030.md/.json`；未提交、未 push。

### 2026-10-30 Codex 增量：S3-F2 峰值工作集专测

为补齐 S3-F2 7,830 行扩展负载的内存数据，新增只采进程树峰值工作集的 Python/Rust 对照，未重复耗时基准。三轮分别执行全新 Python 进程和 Rust Release CLI，5 ms 间隔采样；峰值中位数为 43.688/44.340 MB、最大值 43.945/44.383 MB，进程树各观察到 1 个进程。每轮 2,610 行完整输出对拍均为 0 差异，输入 SHA-256 不变。差距约 0.65 MB，不构成优化理由。采样可能遗漏短于 5 ms 的尖峰，夹具为合成数据。驱动及证据见 `tests/compare/perf_s3f2_peak_memory_20261030.py`、`probe/reports/s3f2_peak_memory_20261030.json/.md`。本次未修改产品代码，也未运行 Cargo 测试/构建、未提交、未 push。

### 2026-10-30 Codex 增量：S1-F2 峰值工作集专测

新增 `tests/compare/perf_s1f2_peak_memory_20261030.py`，仅采集 S1-F2 进程树峰值内存。2 本源工作簿、3 张源表、9 行输出，三轮交替执行 Python `ExcelSession`/Excel COM 与 Rust Release CLI；以 5 ms 间隔监测 worker 与本轮新启动的 Excel PID，排除运行前已有 Excel PID。Python+Excel COM/Rust 峰值中位数为 300.078/10.719 MB，最大值 301.574/10.727 MB；三轮输出值/结构对拍通过且输入哈希未变。3 个新 Excel PID 测量后均退出，未关闭用户已有实例。Python 的内存成本主要来自 Office 引擎；Rust 已约 10.7 MB，无明显内存优化空间。样本仅 5 行，采样间隔可能漏掉短峰；本轮仅测内存、未重复耗时基准、未改产品代码、未运行 Cargo 命令、未提交/推送。证据见 `probe/reports/s1f2_peak_memory_20261030.json/.md`。

### 2026-10-30 Codex 增量：S3-F1 峰值工作集专测

遵守 S3-F1 耗时性能已暂结的范围，只增加 memory-only 测量。10,000 行 current/previous CSV 各一份、14 列；3.0 与 3.1 配置各一份。Python 和 Rust 新 worker 三轮交替运行，进程树以 5 ms 间隔采样。峰值工作集中位数为 133.477/119.035 MB，最大值为 133.672/119.098 MB；每轮 220,022 个输出单元格在 1e-10 容差下对拍一致，最大浮点尾差 2.12e-22，输入 SHA-256 未变。未收集耗时，不改变约 1.02×已结项结论。比较器改为单遍流式读取；首次草测因旧比较器逐行重建工作表迭代器造成 O(n²) 耗时而中断，已停止该自建进程并修正后完成三轮。Python 源和产品业务代码均未修改；未运行 Cargo 命令、未提交、未 push。详见 `tests/compare/perf_s3f1_peak_memory_20261030.py` 与 `probe/reports/s3f1_peak_memory_20261030.json/.md`。


### 2026-10-30 Codex 增量：S1-F1 规则表批注保真与全链路复测

S1-F1 结构化模板的“审核规则”表现按 Python 整表复制语义；新增复制批注，并在真实重算结果发布前以受限 OOXML Patch 恢复源 comments XML/VML。只有源/输出批注坐标和文本相同、且 VML 仅含 Note shape 时才修补；遇到文本/位置变化或非批注 VML 则失败关闭。原始工作簿仍独立不变。Windows 合成完整 GUI 服务路径三轮通过：最终结果 52 格零差异、公式文本/缓存与独立期望相同、24 格规则表和批注文本/作者相同、Excel COM 全部可打开、输入哈希未变、新 Excel PID 均退出。S1 公式复制定点测试 17/17、OOXML 批注/VML 测试 3/3；Release 构建成功。全工作区回归、真实机构人工验收及 Linux/Kylin/UOS 未运行。性能数据和限制见 `PERFORMANCE_BASELINE.md` 与 `probe/reports/s1f1_full_chain_20261030.md/.json`；本轮未提交或 push。

### 2026-10-30 Codex 增量：S1-F1 acceptance 阶段计时

S1-F1 验收输出新增微秒阶段计时，覆盖请求/目录扫描、模板档案与匹配、模板预检、结构对比、Excel-Core 修改、真实重算、公式与条件格式异常提取、结果写出以及专用 Excel 会话退出，并记录逐文件耗时。定向 `s1_f1_` 测试 4 passed、0 failed、2 ignored；Release 构建通过。单轮诊断显示 Excel 重算 3.381 s、会话关闭 2.929 s，其他业务阶段远低于 0.2 s；测量与三轮粗复测期间有 Rust-Tools COM 基准及多个 Excel 实例并行，因此不改正式倍率基线，也不据此放松 Excel 隔离或清理。原始 clean-run JSON 因复跑使用固定路径而被覆盖；历史 Markdown 与性能历史保留 clean-run 结果/逐轮总耗时，现新增明确标记的恢复摘要和并发探针文件；驱动默认已改用唯一 UTC 时间戳输出。原始 clean-run Python 阶段日志和逐轮内存/输入哈希不再可恢复。详见 `PERFORMANCE_HISTORY.md`、`PERFORMANCE_BASELINE.md`、`probe/reports/s1f1_stage_profile_20261030.md` 与 `probe/reports/s1f1_full_chain_20261030.md`。工作区仍未提交；无 push。

### 2026-10-31 增量：GUI 导航与设置中心交互修复（未编译验证）

主界面左侧导航移除“S5 逐笔统计系统（配置版）”入口及配套导航 JS（`openS5Workbench`、`registerS5NavigationClick`、`s5EggClickTimes`）；S5 面板 HTML、CLI、API 与业务模块全部保留。设置中心五个子页接入当前真实能力：常规页递归扫描深度改为禁用说明（深度限制未接入扫描路径）；帮助与更新页在 `get_update_log`/`get_faq` 未实现时按子页显示准确静态说明；组合工作表配置页改为 S1-F3/F4 暂缓状态说明页（`renderDagFlowComposer` 三个重复定义收敛为一个静态实现）；高级页旧 Python 管线死控件（计算引擎/条件格式读取判定/六个读写器/大集中读写器/金融表单生成方式/表达式求值方式）替换为 Rust 内置管线说明，保留真实生效的 Rust 引擎卡与 3.1 表达式语法选择；文件选择方式固定为浏览器内置。`loadSettingsData`/`saveSettings` 不再调用后端不存在的 DAG 草稿接口；`loadDagStudioData` 失败回退空数据。

设置持久化：`app_settings.rs` 的 `AppSettings` 新增 `preferences` 白名单映射（29 键：目录/文件/布尔/枚举），`validate_preference`/`store_preferences`/`apply_preferences_to_state`；设置文件在 `AUDIT_STATE_DIR` 或 `%LOCALAPPDATA%/open-data-audit/settings.json`，原子 tmp+rename 写入。`web_bridge.rs` 启动时按存在性校验路径并把合法偏好叠加回 GUI 状态；`update_state` 与 `choose_history_config`/`choose_central`/`choose_period_compare` 成功后把白名单键快照写入设置文件，失败明确报错。修复了一处在持锁状态下调用会重新加锁的持久化函数导致的死锁点。

本轮修复过程说明：前一attempt 的 index.html 编辑曾误删约 2,932 行前端定义（renderGeneral/renderGuide/renderUpdates/renderAdvanced/saveSettings/openSettingsWorkbench 等丢失而调用方仍在），本轮从 HEAD `cb966f6` 恢复后以行级手术加断言方式重新实施全部修复，损坏副本备份于会话临时目录。验证方式：`node --check` 提取脚本语法通过；19 个移除标识符零残留；被调用函数与定义对齐核查通过；Rust 侧仅静态审查（导入、锁顺序、结构体构造点），未运行 cargo build/check/test（未获授权）。未提交、未 push。已知未完成：常规页“打开/设为默认/重置”等配置工作簿按钮依赖的 `open_config_file`/`reset_*`/`set_*_default` 后端接口仍未实现，点击会明确报错；帮助与更新的 LibreOffice 安装子页在 Windows 显示“仅 Linux 支持”属预期。

### 2026-10-31 增量：跨平台文件选择器（原生优先、页面内兜底，未编译验证）

GUI 文件/目录选择改为“系统原生优先，页面内目录浏览器固定兜底”。后端新增 `native_file_dialog.rs`：Windows 走 rfd 0.17 的 IFileOpenDialog（COM，`spawn_blocking` 执行，不阻塞 Tokio 工作线程）；Linux 走 ashpd 0.13 的 XDG Desktop Portal FileChooser（zbus 纯 Rust D-Bus，运行时才连接会话总线，二进制加载不依赖 libdbus/GTK；评估时确认 rfd 的 Linux 门户后端为手写 FFI 硬链接 libdbus，会造成启动期强制系统库依赖，因此 Linux 端不用 rfd 而用 ashpd）。统一契约：用户取消返回 `canceled=true`（前端不再回退）；能力不可用/超时（600 秒安全超时）返回带可读诊断的空结果（无 D-Bus 会话、门户服务缺失分别映射独立诊断），前端提示后回退页面内选择器。新增 `pick_paths` API 在 `call_api` 中于同步分发前 await 处理；过滤器复用 `browse_extensions` 模式映射；起始路径归一化（文件→父目录，不存在→系统默认并提示）；选中的路径仍全部经由既有 `choose_file`/`choose_folder`/`choose_central`/`choose_period_compare`/`choose_history_config` 等接口校验，本轮未改任何业务规则。

前端将 `usesBrowserPathPicker()` 双分支模式收敛为公共入口 `pickPathSmart`（自动=原生优先+失败回退；系统原生=失败提示后回退；浏览器内置=仅页面内），重接全部 11 个函数共 12 处选择入口（S1 pick/pickHistoryConfig/appendExtraFiles/联合模板两步、S2 pickPc、S3 pickCentral/说明反馈批量导入、S5 pickS5/维护工具输出、LibreOffice 安装源）；`filePickerMode` 设置恢复为三值（自动/系统原生/浏览器内置，默认自动），走 `update_state` 即时保存并加入 `app_settings.rs` 偏好白名单（29→30 键）跨启动恢复；高级设置页恢复实时下拉。Cargo 新增目标平台限定依赖：Windows `rfd`（default-features=false，其 win32 后端使用的 windows-sys 0.61 与项目现有钉版一致）、Linux `ashpd`（default-features=false + file_chooser/tokio，无系统库）；workspace tokio 显式补 `time` feature。

验证方式：`node --check` 通过；`usesBrowserPathPicker` 零残留；12 处入口改接逐一断言；后端多选路径校验链静态核对（S3-F03 反馈导入逐文件 .xlsx 校验保留）。未运行 cargo build/check/test（未获授权）；rfd/ashpd 尚未进入 Cargo.lock，首次构建需可解析新依赖；ashpd 0.13 与 rfd 0.17 的少量 API 签名（FileFilter::glob 所有权风格、FileDialog 消费式构建器、Request::response）以 docs.rs 0.17.2/0.13.13 文档为准，需编译确认。未提交、未 push。待人工验收：Windows 单/多文件、目录、取消、扩展名过滤；Linux Portal 可用/不可用/取消/回退；起始路径不存在、含空格或中文路径；确认离线 DEB 无新增系统依赖。

### 2026-10-31 增量：S1-F01 Linux LibreOffice 批次会话复用（WSL Ubuntu 24.04 实测；GUI 状态为当时快照）

Linux S1-F01 公式重算按 rules/13 §4.5 实现批次会话复用，替代“每本各启动一次 soffice --convert-to”。两条已实测路径：CLI/验收流（open-data-audit，重算清单在循环前已知）走 `recalculate_s1_formula_copy_batch` → sidecar 新命令 `recalculate-batch --pairs <tsv>`，单次 LibreOffice 启动顺序转换整批（`lo_batch::run_batch_convert`），逐本完成清缓存、重算标记、规范化、验证与不覆盖发布，父级对本批成功项复核 `validate_recalculated_copy` 后发布报告；批次级失败或单本缺结果时该本回退旧逐本路径，保持既有失败语义。GUI 流（web_bridge）的审核副本逐本创建并与人工恢复语义耦合，本轮保持逐本重算，但走新增的 `recalculate-session --engine lo` Linux 持久会话（`LoWarmProfileSession`）：profile 只初始化一次并排空残留进程，每本用暖 profile 转换，省去每本首启成本。Windows 原生 COM 路径与单发 sidecar 命令未改动。

方案决策记录：Rust 直控 UNO（URP）无可用绑定、C++ UNO SDK 会引入版本锁定运行时依赖，均否决；rfd 0.17 的 Linux 门户后端为手写 FFI 硬链接 libdbus，违反启动期零新增系统库约束（与本次无关，仅评估记录）。曾实现的 StarBasic 批次驱动（宏轮询任务文件）在实测中不可靠：全新 profile 首启会重写 Standard 库索引导致宏失联（改预置 Module1 后仍复现）；任务行存在时宏静默终止、无任何诊断输出（空任务则正常循环），不满足“失败必须有明确诊断”，已整体放弃并记录。因此 GUI 暖 profile 会话与 CLI 批次均使用 `--convert-to` 通道（与既有逐本路径同一转换语义：清缓存 + fullCalcOnLoad/forceFullCalc 标记）。

已知竞态与处理（WSL Ubuntu 24.04 / LibreOffice 24.2.7.2 实测）：`--terminate_after_init` 的 wrapper 进程先于 soffice.bin 退出，oosplash fork soffice.bin 的窗口在 wrapper 退出后 0.3–1.4s，宏启动会经单实例通道把 URL 转发给残留实例后秒退。处理：初始化后沉降 2 秒，再以 pkill/pgrep（仅匹配本任务 profile 参数，绝不按进程名）清理并确认排空；WSL 的 `/usr/bin/kill -0 -PGID` 对已消失的进程组仍返回 0，进程组存在性原语不可靠，故未采用。

实测（WSL Ubuntu 24.04，LibreOffice 24.2.7.2，合成 400 行公式工作簿，预热 1 轮 + 3 测量轮，输入 SHA-256 前后一致，输出可用 openpyxl 重开且表名/公式数与输入一致）：2 本逐本 3.07s / 批次 3.69s（批次含 profile 初始化与 2s 沉降固定成本，约 3 本起收益为正）；10 本 15.34s → 4.29s（省 72%）；33 本 50.58s → 6.42s（省 87%，平均每本 1.53s → 0.19s）。批次内每本阶段计时不可得（convert-to 单进程顺序转换），按批记录 convert_ms 与逐本状态。报告见 `probe/reports/s1f1_lo_batch_convert_20261031.json`，复跑入口 `tests/compare/perf_s1f1_lo_batch_session_20261031.py` 与 `tests/compare/wsl_run_perf_full.sh`。

验证与边界：Windows host 与 WSL Linux 双侧 `cargo check`/`cargo build --release` 通过；`cargo test -p lo-convert` 21/21、`audit-core --lib` 定向 s1_recalculation 5/5、s1_f1 4/4（2 ignored 为 Windows COM 前置）、served_page 1/1。未运行全 workspace 回归、真实机构输入与麒麟/UOS 目标机验证；WSL 实测不等同 Kylin/UOS 验收。未提交、未 push。


### 2026-10-31 增量：doc→docx 转换现状记录与将来接入的复用方案（未实施）

应用户要求核查 doc→docx 与 xls→xlsx 重算的调用模式差异，结论记录如下，供 D01 接入生产时使用。当前 **doc→docx 没有任何生产调用方**：audit-core 仅定义 `WordComSession`（供 lo-convert sidecar 的 Windows COM 分支引用），无 `lo-convert convert --to docx` 调用点；D01 在 rules/13 为“仅已登记方向”（unverified）；Windows 标准构建/打包不要求或包含 lo-convert。因此本次不修改代码，仅登记现状。

现状（lo-convert sidecar 内，LO 路径）：每次 doc→docx 转换 = **3 次 soffice 冷启动、各自全新 profile**——1 次 `--convert-to docx:Office Open XML Text` 转换，加 `extract_word_text_lo` 对源/输出各 1 次 `--convert-to txt:Text` 文本校验。Windows 路径：每次转换创建一个 `Word.Application` 会话（WordComSession，Drop 即退出），同样未批次复用。两者均不符合 §4.5 的批次会话要求，与 xls→xlsx 重算修复前的形态一致。

将来把 doc→docx 接入生产功能时的既定方案（工作量小，依赖已就位）：复用 `sidecar/lo-convert/src/lo_batch.rs` 的 `LoWarmProfileSession`——需小幅参数化（当前 filter/输出文件名写死 xlsx，需按目标类型选择 `docx:Office Open XML Text` 与 .docx 命名），转换与两次文本提取共用同一暖 profile，单次转换成本从 3 次冷启动降为约 1 次启动 + 3 次同进程内转换；Windows 侧 Word COM 改为批次会话（§4.5 同样要求 Word 会话按批复用）；多输入单实例批次（`run_batch_convert` 模式）同样适用于纯批量转换场景。接入时按 §4.5 与 rules/13 第 6 节对 D01 单独登记验收（当前 unverified），未验收前不得成为任何功能的默认路径。

### 2026-11-04 增量：S1-F1 Linux GUI 接入单次 LibreOffice 批次重算

纠正 2026-10-31 的 GUI 状态快照：Linux S1-F1 GUI 先为本轮所有已通过结构预检的源工作簿准备各自的公式审核副本，再通过 `recalculate_s1_formula_copy_batch` 调用 sidecar `recalculate-batch --pairs ... --engine lo`。`lo_batch::run_batch_convert` 为本批创建唯一隔离 profile，通过一次 `soffice --convert-to` 调用按序处理多本副本；不是每本工作簿单独冷启动，也不宣称存在跨文件持久 UNO 文档会话。每本输出继续独立清缓存、设置重算标记、规范化、检查公式缓存与结构，再按原有不覆盖语义发布。

LO CLI 只提供批次计时，现有 `S1RecalculationReport.elapsed_ms` 为分摊后的批次平均值，不是逐本实测耗时；单文件阶段计时不可得。

失败处理保留 S1-F1/S1-F3 语义：单本副本准备失败单独返回，其他可准备的副本仍参加本批；批次或单本重算/验证失败时，保留该本已准备的无缓存公式副本并明确标记等待人工重算，不再逐本冷启动 LO 重试。Windows 原生 COM 批次路径未改。

WSL Ubuntu + LibreOffice 24.2.7.2 实测：定向集成测试 2/2 通过；另运行一次忽略的 2/10/33 本性能探针（每本 400 行公式数据、401 个公式）。逐本 `recalculate --engine lo` 与单次批次 `recalculate-batch` 端到端耗时分别为：2 本 5,933/4,987 ms（1.19×），10 本 32,813/10,575 ms（3.10×），33 本 80,381/21,837 ms（3.68×）。三档批次报告均为一次 `soffice` 转换调用、profile 残留进程数为 0；两种路径输出的公式文本与 401 个缓存值均通过 Calamine 读取校验；原始输入字节逐本一致。一次测量仅用于验证批量复用收益，不是 Python/Rust 对照或正式三轮基线。

报告与机器数据：`probe/reports/s1f1_lo_batch_convert_20261104.md/.json`；可复跑入口为 `sidecar/lo-convert/tests/recalculate_libreoffice.rs` 中被忽略的 `manual_s1_formula_recalculation_batch_performance_2_10_33`。本轮 Linux `audit-core --lib s1_f1` 定向测试 4/4、LibreOffice sidecar 集成测试 2/2 通过；Windows S1 公式复制定向测试 17/17 通过，Windows sidecar 集成测试目标编译通过。未重复全工作区回归。未在真实机构工作簿、Linux 桌面 GUI、Kylin V10 或 UOS V20 上验收；GUI 与目标发行版人工验收仍待进行。未提交、未 push。

### 2026-10-31 增量：doc→docx 批次转换接入双引擎会话复用（移植 Rust-Tools e2973e6）

sidecar 新增 `convert-batch --pairs <tsv> --to <xlsx|docx> --engine <auto|com|lo>`，整个批次只创建一个引擎实例并顺序处理（rules/13 §4.5）。移植来源：Rust-Tools 提交 `e2973e6`（ConvertWord/ExtractWordText 请求语义、批次自有实例、进程证据），适配 Rust-Audit 的 `WordComSession`（`&self` 可复用会话 + Drop 退出）。实测（Windows 11 + Word COM 与 WSL Ubuntu 24.04 + LibreOffice 24.2.7.2 双侧）：Windows 3 本 doc→docx 同一 Word 进程 PID 贯穿整批（35772），每本 DOCX 包合法性 + 文本一致校验通过；LO 3 本单次启动 2.0s 完成转换 + 按批文本提取（源/候选各一次，取代每本两次冷启动），校验 level `docx_package_and_batched_text_content`。实现要点：`lo_batch::run_batch_convert` 参数化目标类型（xlsx/docx filter 与 staging 命名）；`run_batch_text_extract` 按批文本提取；`WordComSession` 暴露 `process_id()` 作为复用证据；`create_isolated_application` 在 Word/WPS Writer 缺少 Hwnd 属性（HRESULT 0x80020006）时退化为启动前后进程快照差集 + 可执行名匹配（winword.exe/wps.exe）确认归属，多于一个候选保守报错，不影响既有 Excel Hwnd 路径。注意：doc→docx 目前仍无 audit-core 生产调用方（D01 unverified），本能力仅以 sidecar 命令与登记提供，未成为任何功能默认路径。测试：lo-convert 21/21、audit-core --lib 284/284 三轮稳定。未提交状态随本条一并提交；未 push。

### 2026-11-06 增量：S1-F1 校验区域复制批注与 openpyxl 候选验证

S1-F1 常规模板按配置映射的“校验区域”命名范围复制值、文本、公式与批注，并覆盖目标区域中的空白位置；模板区域外内容不复制或扫描。复制阶段公式校验只比较计划导入区域，不再以源工作簿区域外公式文本差异阻断。OOXML 导入计划启用批注复制；OOXML Patch 补齐 Excel/Umya VML 命名空间识别，并在审核副本中复制批注文本、作者/位置和 Note VML 锚点。Openpyxl 输出校验按工作表身份规范化 openpyxl 重新分配的包关系 ID/批注路径，并解析条件格式的 `dxfId` 到实际 differential style、排序 `sqref` 范围后再比较：只忽略已验证语义相同的序列化变动，仍拒绝条件格式样式/规则实质变化；外部链接及其他关系继续严格比较。另针对 openpyxl 在 workbook relationships 中把目标写成 `/xl/...`、并重排未被 `workbook.xml` 引用的 styles/theme rId 的真实差异，增加按解析后的包内目标路径归一化；工作表 rId/目标仍按表身份校验，外部链接不做归一化。

单位存款真实 S1-F1 样本本轮分别通过 OOXML 与显式 openpyxl 写后验证；每份候选输出工作簿有 1,332 条公式、1,091 个批注，4 个 ImportRange 操作，源表与模板输入字节未改变。旧 Python Direct OOXML 同输入输出共 1,332 条公式、1,133 个写入单元格，保留源表现有的 1,053 个批注；比 Rust 少的 38 个模板批注属于用户明确要求复制的新增行为。旧 Python 副本在配置映射的 3,656 个校验区域单元格上与模板逐格一致（0 内容/公式差异）；Rust 候选通过相同计划范围的写后校验。另对股权投资四区域样本，默认 OOXML 候选通过（输出 574 条公式、68 个批注）；显式 openpyxl 因目标表含合并区域非锚点内容在写入前拒绝，输入未变且未发布副本。Openpyxl 验收由本机 Python 3.12.10 + openpyxl 3.1.5 测试 worker 执行；这不构成生产依赖。当前高级设置支持 openpyxl 显式试跑及 Auto 的中间回退；当前 `dist` 已包含 worker 脚本，但不含 Python/openpyxl runtime。Windows 需要 PATH Python 能导入 openpyxl；Linux 仍需提供隔离 runtime 才可启用。

与旧 Python 的核对：Python 默认 Direct OOXML 命名区域写入器不复制批注对象，但 COM Range.Copy/openpyxl 路径会复制。当前遵循用户最新要求复制批注，因此相对旧 Python 默认 Direct OOXML 是明确业务扩展，不是旧版同输出的证明。Python 外部工作表计划基于模板整簿公式选表；Rust 基于命名校验区公式并递归追踪依赖。若校验区外普通模板公式单独引用某外部表，Rust 可能少导入 Python 会复制的辅助表；这不把区域外公式变成审核规则，但仍需按实际依赖确认输出结构。

验证：`excel-editor --lib` 65/65，`audit-core --lib` 227 passed/0 failed/12 ignored，S1 公式复制定向单测 18 passed/0 failed/2 ignored；真实单位存款样本 OOXML 候选 1/1、显式 openpyxl 候选 1/1，`cargo check -p audit-core --bin open-data-audit` 通过。Openpyxl 条件格式语义归一化单测覆盖同一 dxf 内容不同索引、`sqref` 顺序变化通过，实际样式变化拒绝；关系路径单测覆盖等价 package-relative/package-absolute 路径和 styles/theme rId 变化通过，工作表目标或 rId 变化拒绝。候选测试止于重算前；当前源码的完整 129 本 COM 审核、全 workspace 回归、GUI 人工验收和 Linux/Kylin/UOS 实机均未执行。未提交、未 push。

### 2026-11-06 增量：S1-F01 修改引擎选项与票据精确比较决定

按用户确认，票据校验继续保留模板公式的精确 `<>0` 比较；程序不增加浮点容差，今后如需容差由业务人员修改公式。高级设置新增 Auto、OOXML、openpyxl、Umya 四项，适用范围仅 S1-F01。Auto 顺序为 OOXML Patch → 可用 openpyxl worker → Umya；可用性检测包括 Python、worker 文件及 `import openpyxl`。只有后端在写入前明确拒绝，才从干净的原始副本尝试下一后端；写入后失败或输出验证失败时停止，不静默回退。旧设置保存值 `umya` 表示过去的默认 OOXML→Umya 路径，读取时迁移为 Auto；显式 Umya 新值单独序列化，避免升级后改变旧设置语义。S2/S3 路由不受此设置影响。

为 13 份旧 Python 提取器 `IndexError` 样本（95 条 Rust-only 公式异常）生成了只读路径清单和逐条异常清单，源文件及重算副本均核实存在：`F:/RustBuild/audit/tmp/s1f1-corrected-inputs-129-20261106/s1f1_python_extractor_indexerror_13_paths.csv` 与同目录 `s1f1_python_extractor_indexerror_95_items.csv`。既有 129 本完整批次使用 OOXML，不作为 Auto→openpyxl→Umya 自动回退整批验证证据；人工抽查及自动回退实样验收仍待完成。本轮设置迁移测试 7/7、路由测试 5/5、高级设置桥接测试 2/2；CLI `cargo check` 通过，`build-release.bat` 通过并更新 `dist/open-data-audit.exe`。未运行全量回归或 GUI 点击验收。

### 2026-11-06 增量：S1-F01 复制验证范围及条件格式源路径纠正

公式复制写后校验仅验证从模板导入的命名“校验区域”，不再因区域外公式文本变化拒绝副本；区域外数字/文字/公式也不作为模板复制或公式异常提取范围。条件格式求值从原始源工作簿读取，以匹配旧 Python；公式异常继续从重算后的审核副本读取。定向条件格式来源路径测试 1/1 通过，Release 构建通过。新 Release 对 129 本真实 S1-F01 完整运行：129 完成，失败/跳过/待人工重算为 0，3,373 公式异常 + 349 条件格式异常，所有输入 SHA-256 未变；端到端 292,626 ms。和同口径 Python 结果比较为 3,616 对 3,722，Rust-only 106 个位置与先前完全一致，无条件格式项目增减。5 个共有条件格式字段差异、95 个 Python `IndexError` 漏项和 11 个按模板精确 `<>0` 检查的票据缓存残值继续待业务验收。报告与逐项对照位于 `F:/RustBuild/audit/tmp/s1f1-source-cf-current-129-20261106/`。未跑全 workspace 回归，未做 GUI 人工验收，未提交、未 push。


### 2026-11-06 增量：S1-F01 条件格式取值来源偏好

高级设置新增条件格式取值来源：`source`（默认，原始源工作簿）或 `audit_copy`（模板校验区域复制并重算后的审核副本）。两种模式使用同一批配置定义的条件格式扫描范围；选择只改变规则和值所在工作簿。偏好通过现有用户 JSON 白名单保存，启动 S1-F1 时记入恢复清单，S1-F3 续跑时使用原选择；旧清单缺少字段时按 `source` 处理。来源选择、设置保存/重载、S1-F1→S1-F3 恢复清单持久化测试各 1/1 通过；`build-release.bat` 成功并更新 `dist/open-data-audit.exe`。尚未实际点击 GUI；`audit_copy` 模式尚无真实工作簿对照证据。未跑全 workspace 回归，未提交、未 push。

### 2026-11-06 增量：S1-F01 条件格式字段文本与129本复验

为逐字段复现旧 Python 的条件格式指标标签，`s1_conditional_audit::display_value` 现在把 Calamine 类型化 Excel 错误转换成 Excel 显示文本；旧 Python 使用 openpyxl `data_only=True` 后会把这类缓存错误转成字符串参与 `check_field`。改动不影响触发求值、触发值或说明，但会改变五条既有 Rust 条件格式记录的 `check_field` 和依赖它生成的 `issue_id`，可能影响这五条历史记录的连续归并。新增针对 `Div0` 的单测通过 1/1；独立 Agent 只读复核确认实现边界与旧 Python 的标签提取一致。构建当前 Release 成功，`dist/open-data-audit.exe` 已更新。

使用同一组 129 本核查表、19 个实际命中模板、S1 配置与外部数据重跑：129/129 完成，跳过/失败/待人工重算均为 0；公式异常 3,373、条件格式异常 349；端到端 283,415 ms；157 个受保护输入的哈希运行前后未变，并与上一验收的输入哈希一致。编辑器为 OOXML Patch 129 本；重算引擎为 Excel 122 本、WPS 7 本。与 Python 基线相比，共有 3,616 条现在全部字段一致，Python-only 0，Rust-only 106；旧 106 个 Rust-only 位置均仍存在。95 条是旧 Python 可复现 `IndexError` 漏提取、但属于用户确认的命名校验区域规则；11 条是用户决定保留精确 `<>0` 的票据浮点残值。因此不能宣称字面全量同步，未为对齐计数而删除这些业务检查。

重新复核 7 本债券副本：无 `xl/externalLinks` 部件，`参照表!D2:D211` 每本 210 个公式均引用本地 `机构参照表`。与实际外部数据中的 `集中系统数据`、`参照表`、`机构参照表` 逐格比对，7 本公式/常量和缓存均为 0 差异；当前模板的 externalLinks 已按业务方说明修正，没有静默剥除仍然需要的依赖。

验收、逐项差异及债券引用表证据位于 `F:/RustBuild/audit/tmp/s1f1-error-label-parity-129-20261106_172554/`；此前 106 与 5 条详细列表位于 `F:/RustBuild/audit/tmp/s1f1-source-cf-current-129-20261106/S1F1_Python_Rust_106与5条差异明细_20261106.md`。定向单测 1/1，Release 构建通过；未运行全 workspace 回归，未做 GUI 人工点击或 Linux/Kylin/UOS 验收。Goal 仍未完成，等待业务方决定 106 个有效业务差异是否可作为明确例外；未提交、未 push。
### 2026-11-06：S1-F1 条件格式标签按结构文本取值

业务方确认五条共有条件格式异常仍需保留，目标单元格数值仍按规则正常判定；错误缓存 `#DIV/0!` 不参与指标标签，标签只从表结构范围文本解析。撤回把错误缓存显示文本附加到 `check_field` 的实现，定向测试 `error_valued_header_is_not_used_as_indicator_text` 通过 1/1。当前 129 本 Release 验收早于该口径修正，尚未以当前代码构建和复跑。

业务方后续模板公式统一输出 `错|校验类型|校验指标|详细描述|本期值|对比值|差值`；现有七段解析器已按位置写入这些字段。抽查旧公式 `个人贷款发生额信息!K69` 实际只返回四段，第三段为空且没有数值段，因此旧结果中的校验指标和数值字段留空是源公式内容所致。11 条票据 `±3.64E-12` 残值建议在具体模板公式设置 `1E-10` 容差；Rust 暂继续按当前模板 `<>0` 精确计算，未修改模板和通用结果筛选逻辑。

### 2026-11-06：S1-F1 票据残值与条件格式标签复核

对照 Python/Rust 汇总和对应审核副本后，11 个票据 W 单元格公式文本一致；Python 副本中对应缓存为空，Rust Excel 重算缓存为 10 个 `-3.63797880709171E-12`、1 个 `+3.63797880709171E-12`。11 个绝对残差均小于 `1E-10`。因此不能说旧 Python 将这些公式计算为零；其汇总没有列出记录与公式缓存为空相符，但 Python 链路具体未得到缓存的步骤仍未定位。容差建议只在业务模板公式中设置，Rust 不进行通用结果过滤。

五条条件格式记录的规则判断基于目标 F 单元格数值；邻近表头单元格的 `#DIV/0!` 不应影响触发判定或指标文本。`check_field` 应取结构区域标签。当前 129 本 Release 结果晚于/早于不同源码口径的构建，未用当前源码重建复验。Markdown 公式样例中的 `\\|` 是表格转义，Excel 公式文本为普通 `|`。本条为只读文件核验，未运行 Cargo、未改测试工作簿。
