# Rust Workspace Excel 读写、修改与引擎路由规范

> 2026-11-13 当前实现覆盖说明：独立 `sidecar/lo-convert` 已移除；下文该路径、独立转换器及旧候选命令属于历史记录，不再作为开发/构建指令。Windows 用 Audit 内置原生 COM；LibreOffice 重算由主程序内部 worker 承接，一批一次隔离 profile 的转换调用，并非常驻 UNO。LibreOffice 软件仍是可选外部能力。新 worker 尚未完成编译/真实环境验收。

> 状态：Rust-Audit 独立项目的 Excel 专项规范，保留 2026-10-22/23 的限定范围历史证据。范围：Excel/CSV 读取、新建、修改、转换、页面设置和具体业务所需的验证。
> 优先级：用户本轮明确要求 > workspace 根 AGENTS.md 的硬约束 > 本规范的 Excel 专项规则 > audit-core/rules/ 下的业务补充与旧项目历史规则。冲突不得静默覆盖业务行为。  
> 来源：合并 workspace 现行规则、现有 Excel 引擎盘点、Open-Data-Audit 与 py_statistics_tools 的只读规则材料，以及用户提供的建议稿。本文件是执行入口；根目录来源稿保留为输入材料，不是另一份独立规则。

## 1. 项目目标与已确认边界

本项目独立维护 Open-Data-Audit 和本地 Excel-Core。Rust 负责业务核心、读取、规则计算、新建工作簿及 ExcelEditPlan/ExcelModifier 调度；既有工作簿修改按实际能力使用已有适配器。隔离 openpyxl worker 仅限已有工作簿修改，不承接业务逻辑；其他外部引擎须明确选择并经过其适用范围的验证。旧 Python 项目只作只读行为参考。公式计算不属于 Reader、Writer 或 Modifier。

目标交付为麒麟 V10 x86_64 和统信 UOS V20 x86_64 离线 deb，兼容 glibc 2.28；核心路径不得要求用户预装 Python、Go、Office、WPS、LibreOffice 或 JVM。启用 openpyxl 引擎时可在 deb 内置隔离 Python runtime 与依赖，不得使用系统 Python、用户 PATH 或联网 pip。可选 LibreOffice 计算能力须明确独立运行时来源；Windows COM 不进入 Linux deb 或其必要依赖。ABI 和打包约束见根 AGENTS.md。

### 1.1 主路径和旁路

- 核心主路径不依赖 Office、WPS、LibreOffice、COM 或 UNO。
- 公式计算、同类别格式转换及未来 Windows COM 修改只能作为独立、显式选择的可选能力；其不可用不影响读取、新建和其他业务路径。
- 公式计算通过独立 ExcelCalculator 提供，默认关闭。用户可显式请求普通重算或完整重算；必须验证公式语义、结果和依赖关系，并检查引擎保存造成的公式、样式或结构变化。不得用旧缓存、零值或静态重建冒充重算；未安装/不可用时报告“该功能需要外部引擎”。
- 同类转换由用户明确触发，可选择经对拍的保真转换，或经说明损失范围的 Rust/Go 转换。已确认方向包括 Excel .et/.xls → .xlsx、Word .doc → .docx；.xls 可作为改工作表名等操作前的显式转换。具体方向/引擎未登记通过前仍属未验证；结果必须验证 sheet、数据和契约要求保留的对象。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。
- sidecar/worker 是独立进程，主程序启动时不检查外部引擎。用户显式触发时才检查依赖；缺失/失败不能阻塞纯 Rust 主路径，未验证输出不得报告成功。
- Microsoft Excel COM 与 WPS COM 可作为独立 Windows ExcelModifier/ExcelCalculator 后端，能力按操作分别验收。优先评估 Rust 原生 IDispatch/COM，通过专用 STA worker 调用；Excel 使用 `Excel.Application`，WPS 优先验证 `KET.Application`。不因历史方案而默认增加 pywin32 或 PowerShell 中转；只有某项能力、兼容性或测量结果明确要求时才考虑。Excel/WPS 使用不同 EngineId 和 capability 证据；不得仅凭软件已安装判定可用。必须创建本任务拥有的应用实例，不连接或关闭用户已打开的 Excel/WPS；COM 不提供 Reader/Writer，也不进入 Linux 依赖。
- LibreOffice Calc 可作为独立 ExcelCalculator/格式转换后端（Windows/Linux）；Microsoft Excel COM、WPS COM 仅限 Windows，WPS 单独验收。LibreOffice 使用其适用的 UNO/进程接口，不称为 COM。所有外部应用的进程与会话生命周期遵循 §4.5；具体操作仍须单独验收。
- PDF、纸张渲染、打印机输出和跨类别转换不在当前迁移范围。打印只写页面设置。旧 COM 业务工作流不得整体搬入 sidecar/worker 以规避 Rust 迁移。
- 可选 LibreOffice 计算能力不得假设 Linux 目标机预装；发布前须说明独立离线组件/安装包或受支持外部安装等运行时来源，并完成目标平台验证。

### 1.2 已确认模板与打印范围

已检查模板不含图表、透视表、切片器、图片、工作表保护或 VBA；这仅覆盖已检查模板，不代表未来输入没有这些对象。修改前仍须清点实际文件。

旧工作区曾记录 py_statistics_tools 的模板范围；该项目不属于 Rust-Audit 的开发或同步范围。

打印只写 orientation、fitToPage、fitToWidth、fitToHeight、清除 scale、paperSize、printArea、horizontalCentered、verticalCentered 和页脚页码。旧代码没有写 margins、printTitleRows/Columns、gridLines、headings、blackAndWhite、draft、firstPageNumber；新增属性须另立需求，不得顺手加入。

## 2. Workspace 结构与依赖边界

延续当前目录，不因接入规范重命名 crate 或复制实现：

    Rust-Audit/
    ├── Cargo.toml / AGENTS.md / rules/13_Rust_Excel读写修改与引擎路由规范.md
    ├── crates/{excel-reader,excel-writer,excel-editor,ooxml-patch}
    ├── audit-core/
    ├── sidecar/lo-convert/
    └── GPT修改记录/

- audit-core 使用本项目 crates/ 中独立维护的 Excel 原语；不得依赖 pytools-core 或其他项目路径。通用第三方依赖由本项目根 Cargo.toml 管理版本。
- crates/ 中的库只提供 Excel 原语，不含审核、时序、打印等业务概念。
- 业务层只生成与具体引擎无关的请求/ExcelEditPlan，不泄漏 calamine、umya、Excelize 等库类型。
- 一个文件的一批修改原则上由一个覆盖整批操作和保留项的已验收引擎完成，避免反复打开/保存。
- sidecar 独立构建和显式子进程调用；业务 crate 不静态依赖或导入 sidecar。
- Audit 独立维护业务 fixtures、probe 配置、报告和 deb；不通过 Git 子模块、跨项目路径或自动同步维护本地 Excel-Core。

## 3. 每项 Excel 功能的操作契约

一个端到端功能可能是 R01 → 规则计算 → W01。按每个阶段分别记录：

| 字段 | 要求 |
|---|---|
| feature_id / operation | 业务功能 ID 与第 5 节操作编号；多阶段分别列出。 |
| platform | OS/版本、架构、glibc、库/引擎版本；开发和目标环境分开记录。 |
| input_format / output_format | 按文件内容识别，扩展名只初筛；不靠改后缀转换。新建默认 xlsx，编辑默认原格式，格式变化须登记为转换。 |
| output_kind | new_report、edit_existing、format_convert、filesystem；不得把重建报表称为原位修改。render 当前不支持。 |
| write_pattern | append_only、random_access、edit_existing；“读后重建 sheet”不能标成追加。 |
| data_owner | rust、python_legacy、none；迁移期用来记录数据与业务计算所在侧，并测量跨语言传输成本。最终生产业务计算不依赖 Python；仅 ExcelModifier 的可选 openpyxl worker 可在隔离进程运行。 |
| value_view / formula_mode | stored 默认；displayed 需静态显示格式能力验收。公式模式为 cached、text、text_and_cache、recalculate；旧缓存不等于本轮重算。 |
| preserve / drop_allowed | 明确要保留和明确允许丢弃的对象；默认不丢弃未授权对象。 |
| allow_conversion / external_apps | 默认 false；格式转换、公式计算、Windows COM 修改分别显式授权，不因某项开启而自动启用其他外部能力。 |
| reference_app / engine_mode | none、Microsoft Excel、WPS、LibreOffice；Excel/WPS 作为独立参考结果，不假设一致。Windows COM 脚本可为读取/新建测试生成参考数据或文件，但不进入 Reader/Writer 生产路由。仅提供已实现、环境可用且 capability 已登记通过的选项；用户手动选择失败时不得静默换引擎。 |
| workload_limits | 记录覆盖的文件/sheet/行列/对象数量、耗时和内存；没有用户阈值先测，不编造硬限制。 |

修改已有工作簿时，按实际文件检查：值/类型、公式与缓存、共享/数组公式、合并、样式、行列尺寸/隐藏、工作表名称/顺序/可见性、批注及外观、定义名称、打印设置、超链接、验证、条件格式、表格、图表/图片/形状、外链、VBA/签名、关系和未知扩展。任务不涉及可标“不适用”；没有检查必须标“未验证”，不可假设对象不存在。输入文件变化后，旧清点和能力匹配失效。

## 4. 技术分工

### 4.1 读取

- 生产读取使用纯 Rust reader；普通值读取优先评估纯 Rust calamine，且必须按数据归属、文件格式、输出契约和对象验收。禁止 python-calamine、PyO3，openpyxl worker 不得用于读取路由。
- 当 `data_owner=rust`、输入为 XLSX、操作只读取普通存储值，且坐标、类型、空值和值语义与 §6.2、§6.3 登记的基准契约一致时，Calamine 是优先评估的 reader 候选；这不是该业务功能已经验收的证明。其他格式或不同契约按各自能力登记和对拍。
- Direct OOXML 是直接操作 ZIP/XML 的实现方式，不是固定库，也不是普通值读取在 Calamine 之后的通用 fallback。Calamine 读取失败时必须按实际错误报告，不得再用 Direct OOXML 掩盖或改写错误结果。
- 功能明确需要 Calamine 未提供的 OOXML 专属元数据、模板结构或特定 OOXML 部件时，单独选择已针对该对象和契约验收的专用 OOXML reader/适配器；适配器能力不满足时报告具体缺项，不把普通值读取失败转换成元数据补读。
- Python 拥有数据的迁移期流程保留该项目现有、同契约验收通过的路由。不能由本次 Rust/Python 实现比较推导 Python `python-calamine` 与 Python Direct OOXML 的 reader 排序；候选调整须对该流程的完整调用链同机测量，并计入进程、导入、序列化、传输、数据转换和校验。
- 对 xls、xlsx、xlsm、xlsb 的值、sheet 名、坐标、公式/缓存、合并等字段按格式和锁定版本逐项验收。不能从扩展名或 ZIP 容器推断完整支持。
- Calamine 不提供且业务确需的批注、条件格式、数据验证等字段，用针对这些对象并已验收的 namespace-aware OOXML reader 定点补读；umya 可用于特定字段对拍/补充，但不默认引入第二套完整 reader。
- “API 未提供”“没有此字段”“空值”“没有合并/批注”必须区分。
- CSV 用经对拍的 Rust CSV 实现，固定编码、分隔符、引号、空字段和类型解释。
- 多文件处理用 rayon 做文件级并行，不并行写同一 workbook。

### 4.2 新建工作簿

- 对 Rust 持有结果、从零生成纯数据 XLSX 的功能，优先评估 rust_xlsxwriter；只有输出契约对拍通过且峰值内存符合该功能已定义的资源预算，才可采用。它不是既有工作簿编辑器。
- 对拍值/类型、公式文本、样式、合并、名称、尺寸、页面设置及可打开性；验证真实算法需要顺序追加还是随机访问。
- 基准报告中的 Direct OOXML 写入是 Python S5/S2 专用 writer；报告没有比较 rust_xlsxwriter 与 Rust Direct OOXML writer。不得把该实现称为 Rust writer，也不得仅凭名称、压缩设置或 ZIP 行写方式推断不同 writer 等价。
- Direct OOXML 仅在存在专用输出契约、明确 OOXML 对象要求，或同契约实测证明资源收益时进入候选；constant-memory 模式、既有工作簿编辑和通用 OOXML writer 均须单独实现与验收。

### 4.3 修改已有工作簿

业务层生成 ExcelEditPlan，由 ExcelModifier 按“操作子类 + 格式 + 对象特征 + 保留项 + 平台/版本”选已验收引擎。未来 COM 后端仅修改既有工作簿；不得为统一形式而引入通用 ExcelWritePlan 或完整工作簿对象模型。

- **Rust OOXML Patch**：适合明确部件、明确位置、依赖关系可证明的窄修改；仅对已验收白名单执行，无法证明安全时拒绝。
- **Rust umya-spreadsheet**：复杂 XLSX 对象修改候选。批注、样式、合并、尺寸、公式、页面设置及未知部件须按真实模板分别验收；能打开/保存不等于保真。
- **Go Excelize worker**：条件性候选。仅当 umya/定点补丁对真实模板有大面积失败或维护成本过高，且 Excelize 对同一整批操作验证适用时才评估。Rust 以一个批量 IPC 请求传文件路径和完整计划，不逐单元格跨进程。若采用，随包交付，CGO_ENABLED=0。
- **隔离 openpyxl worker**：可选第四修改引擎。Rust 按每本 workbook 提交一份结构化 ExcelEditPlan，由内部绝对路径启动隔离 runtime 中的 Python/openpyxl。worker 仅执行 XLSX 修改操作，不做读取业务、统计、规则计算、公式重算或新建报表。对非空公式缓存、未知/未验收 package part、超出 capability 的格式/对象或操作批次，必须在写入前拒绝。openpyxl 会重写 XLSX package，必须验证未修改区域及计划要求保留的对象；行插入/删除当前 unsupported。运行时缺失可作为写入前不可用处理，只有存在另一条已验证路由才允许重新从原始输入执行。崩溃、超时、写入后失败或验证不一致不得自动切换。
- **Microsoft Excel/WPS COM**：仅为 Windows 可选修改/计算候选，优先评估 Rust 原生 IDispatch/COM 与专用 STA worker；每种应用、版本、操作组合、对象、保留契约分别验收。COM 修改不是 Reader/Writer。只有 Rust COM 在特定能力上无法满足要求且有实测依据时，才评估隔离 pywin32/PowerShell worker，并登记额外 runtime、IPC 和打包成本。
- 修改引擎不预设全局优先级，也不要求全部启用。先完成第 10 节操作清单和能力矩阵，再验证有实际价值的操作。引擎注册和合成测试不等于业务能力通过。
- umya 已覆盖且未证明 Go 有保真、性能或维护收益时，不引入 Go。

### 4.4 业务计算、公式计算与禁止项

规则、期间、校验、统计、去重、映射和审核结果由 Rust 业务逻辑计算；财务数值、Decimal、舍入、日期、空值、错误和文本前导零策略须明确并对拍。业务计算、公式文本、旧缓存读取、Excel 公式重算是不同能力，不能互相冒充。

openpyxl 仅可通过隔离 worker 用于已验收的既有 XLSX 修改。Windows Excel/WPS COM 操作优先评估 Rust 原生 IDispatch/COM 专用 worker，并遵循 §4.5 的批次复用要求；只有具体能力或实测结果要求时，才使用隔离 pywin32/PowerShell worker。Python/PowerShell COM worker 不得用于生产读取、新建写入、业务规则、统计或通用业务逻辑；若采用，须记录其 runtime、IPC 和随包成本。禁止 python-calamine、PyO3、pandas、xlrd、pyxlsb 进入生产。不得无证据重写已验收模块；不得以完整 OOXML writer 为默认目标；不得无测量串联 Rust 读取→中间文件→Go 修改→Go 写回。

`ExcelCalculator` 与 ExcelModifier 分离，候选为 LibreOffice Calc、Microsoft Excel COM 和 WPS COM；默认关闭，计算模式区分普通重算与完整重算。LibreOffice 基础公式计算的可行性已由用户实际测试确认；具体公式覆盖率、与 Excel 的公式语义一致性、保存保真和 Linux 离线运行时仍未验收。不得宣称计算结果完全等价。显式选择不可用或未经该功能验收的引擎时返回明确错误；界面仅列出已实现、可用且有匹配能力证据的选项。

### 4.5 外部表格应用会话与批量复用

- 对 Excel、WPS、LibreOffice 执行批量格式转换、公式计算或其他工作簿操作时，以一次用户可见的业务运行/任务批次为会话边界：为该批次创建一个本任务拥有的引擎实例或专用服务会话，按顺序处理批次中的工作簿；不得默认每本文件各启动、配置和退出一次应用。
- 每本工作簿按“打开 → 执行该文件操作 → 保存到契约指定的独立输出 → 校验 → 关闭工作簿”处理；工作簿关闭后继续复用同一应用会话。批次结束时只关闭本任务创建并能确认归属的实例，释放 COM/UNO 资源，并验证没有本任务遗留的孤儿进程。禁止按进程名批量结束用户的 Excel/WPS/LibreOffice。
- Windows Excel/WPS 优先使用 Rust 原生 COM/IDispatch，放入专用 STA worker，并在整个批次内复用一份私有 `Excel.Application` 或经单独验证的 WPS `KET.Application` 实例。不得为每本文件重复启动 PowerShell、Python/pywin32 和 Excel/WPS。现有可复用的 native COM worker 可作为实现参考，但其 S1-F1 重算能力与其他操作能力分别登记、分别验收。
- LibreOffice 不使用 COM。Linux/Windows 批量操作优先评估由 Rust 控制的单个隔离 LibreOffice 进程和复用的 UNO 会话，每批使用独立 profile，并逐本打开、操作、保存、关闭。对于已验证支持多输入的 `soffice --convert-to` 操作，也可以由单次调用按顺序处理整批；这属于批次级进程复用，不等于跨文件保持的 UNO 文档会话。不得把 `soffice --convert-to` 每文件启动一次作为默认批量路径。若具体操作无法安全复用会话，必须记录阻碍、同机端到端对比及进程成本，再选择有界的替代路径；不能假装已复用。
- 同一 Office/UNO 会话中的工作簿操作默认串行；不得为并行度在共享实例内并发操作。只有实测证明收益并验证应用隔离、资源占用与进程清理后，才评估多实例并行。
- 每个独立业务批次完成后默认关闭会话，不默认让 Office/LibreOffice 跨用户操作常驻。跨批次常驻仅能作为单独的服务架构评估，必须证明所有权、状态隔离、超时恢复、崩溃清理与用户 Office 共存安全，并有测量收益。
- 单本工作簿失败时，按业务规则记录失败并在会话仍健康时继续其他独立工作簿；引擎状态损坏时先关闭/重建本任务拥有的会话，仅可从未被修改的输入或干净副本重试。保存、公式、保留项或结果校验失败不得通过静默切换引擎掩盖。
- 性能记录需报告批次引擎创建次数、总耗时、成功/失败/跳过数和进程清理结果；可测时再提供每文件打开/操作/保存/关闭阶段。若批量接口只提供整批耗时，应明确逐文件阶段不可得，不得把批次总耗时均摊后冒充实测单文件耗时。比较必须使用同一输入、输出契约和验证强度。
- 用户提供的 `Rust-Tools/perf_xls_to_xlsx_com_comparison.md` 显示，在该 Windows/Excel XLS→XLSX 固定数据集上，Rust 原生 COM 的 B100 批次复用均值为 24.09 秒；逐文件创建实例的 Rust 原生 COM 中位数为 620.54 秒，批次复用中位数为 23.41 秒。该报告显示各批次复用相对逐文件启动节省约 88.0%–96.5%。这是格式转换的单机证据，不是公式计算、WPS 或 LibreOffice 的性能/兼容验收，也不能单独改变其他功能的默认引擎路由。
- 当前实现的生命周期状态须如实记录：S1-F1 Windows native COM worker 在一个重算批次内复用会话；Linux S1-F1 GUI 先准备本批公式副本，再通过一次隔离的 `recalculate-batch` / `soffice --convert-to` 调用顺序重算整批，并在结束时确认专用 profile 已排空。此路径不是持久 UNO 会话；单文件 API 与遗留 `recalculate-session` 不代表 GUI 批次会按该模式逐文件启动。

## 5. 操作编号和目标方向

编号是能力分类，不代表当前已经实现。每项须按操作子类、格式、对象和平台登记。本表中的“目标方向/候选顺序”仅指导评估，不等于已通过当前功能、格式、对象和平台验收的生产路由。

| 编号 | 操作 | 目标方向 |
|---|---|---|
| F01/F02 | 文件操作；按单元格内容命名 | 文件系统 API；读取另走 R01/R04，不重存 workbook。 |
| R01 | 读值、sheet 名 | 仅对 Rust 持有数据、普通存储值 XLSX 且坐标/类型/空值/值语义与专项契约相同的功能，优先评估 Calamine；其他数据归属、格式或契约保留各自已验收路由。Calamine 失败不以 Direct OOXML 通用回退。 |
| R02/R03 | 坐标、合并；公式/缓存、批注、样式、名称、打印元数据 | Calamine 已验收字段；OOXML 专属元数据使用对象/契约已验收的专用 reader，不作为普通值读取 fallback。 |
| R04/R05 | 显示文本或高级格式元数据 | 仅使用对应已验收静态解释器/reader；否则报告未验证或不支持。 |
| W01 | 新建 XLSX | Rust 持有结果、从零生成纯数据 XLSX 时优先评估 rust_xlsxwriter；对应输出契约通过且峰值内存满足预先定义的功能预算后，才可用于该功能。其他 writer 需有专用契约、OOXML 对象要求或同契约实测的资源收益。 |
| W02 | 修改已有文件的值、公式、样式、批注、打印属性 | ExcelEditPlan + ExcelModifier；现有已验证引擎按整批保留契约选择；未来 Microsoft Excel/WPS COM 仅为 Windows 独立候选。 |
| W03/W04 | 追加；插删行列、移动区域、sheet 改名/重排/删除 | 修改操作；检查其他 sheet、公式、名称、表格、打印区域和外链依赖。重建 sheet 不算保真追加。 |
| W05–W07 | 跨簿按值提取、复制指定样式、整表保真复制 | 每项定义 preserve；按值复制不能称为格式/整表复制。 |
| P01/P02 | 写打印页面设置；PDF/渲染 | P01 只写页面设置；P02 当前不支持。 |
| Q01/Q02 | 读旧公式缓存；按本轮输入普通/完整重算 | Q01 标记“文件已保存结果”；Q02 由独立 ExcelCalculator 默认关闭，显式选择已验收计算后端并验证缓存、公式、依赖和保存保真。 |
| C01/C02 | 按值重建为 xlsx；同类格式转换 | C01 必须明示丢弃范围；C02 显式转换并校验。 |
| D01 | Word 同类格式转换 | 仅已登记方向、显式 sidecar 或验收通过的 Rust/Go 转换器；PDF/跨类不支持。 |
| T01/T02 | CSV 读写、分组/连接/透视 | Rust CSV 与业务计算；生成 XLSX 另执行 W01。 |

## 6. 能力登记与当前实现快照

### 6.1 状态

能力按“操作子类 + 格式 + 对象特征 + 保留项 + 平台 + 版本范围”登记，状态仅为 **passed / unverified / unsupported**（通过/未验收/不支持）。代码存在、库文档、mock 或能打开简单文件，都不等于 passed。对拍对象逐项记“保留 / 丢失 / 损坏 / 不适用 / 未验证”；未测量写“未测量”。

登记字段至少包括 operation/feature_id、engine/version/适配器入口、输入输出格式、write_pattern、preserve/drop_allowed、platform、样本、报告、性能范围、默认值和状态。可维护在现有架构文档、静态表或 JSON，不新增插件框架。业务配置、夹具和报告分开。

候选顺序与生产路由必须分开登记：安装了库、基准更快、或出现在候选顺序中都不等于当前功能/格式/对象/平台已经 `passed`。实际生产路由仅能来自具体功能的验收证据和现有默认记录；通用基准不得自动改写默认值。

| 当前组件 | 现状证据 | 可声明范围 |
|---|---|---|
| excel-reader / calamine | fastpath_read_compare 对 synthetic_probe.xlsx：sheet names 1/1、values 65/65、formulas 1/1、merges 1/1、comments 1/1；批注由 OOXML 补读 | 仅上述合成样本/字段可 passed；真实模板及 XLS/XLSM/XLSB/CSV 矩阵仍 unverified。 |
| excel-writer / rust_xlsxwriter | 当前为库类型导出；新建合成工作簿 9/9 | 仅样本断言 passed；业务写入 API、低内存能力和真实版式仍 unverified。 |
| excel-editor / umya | 薄封装；原始合成修改报告 17/18，列定义保留失败；结合窄列补丁后的样本 18/18，补丁范围 3/3 | 这些结果只覆盖该合成样本和列定义操作；pytools 机构名称规范化也仅为合成候选。通用 W02 unverified，无统一路由/事务边界。 |
| ooxml-patch | 有未验证部件拒绝及窄列定义/列宽补丁 | 仅报告内定点操作可 passed；通用单元格、批注、公式、页面设置仍未验收。 |
| probe | 读取/JSON 导出和部分比较 | 修改计划与全引擎路由探针未形成。 |
| Excelize worker | 当前没有 Go module/worker/IPC | 当前 unsupported（未实现）；达到 §4.3 门槛后再立项。 |
| openpyxl worker | `crates/excel-editor/python/openpyxl_worker.py` + Rust `OpenpyxlXlsxBackend`；pytools capability dispatcher 已注册；支持有限 XLSX EditPlan 批处理 | 已实现并通过合成 worker/进程集成测试；Rust 独立读回覆盖已测样本的样式语义、显式行高列宽、填充色、页面设置和 OOXML Print_Area。真实项目模板、操作级 capability、Windows 原生 Rust、glibc 2.28 隔离 runtime、deb 打包和目标机均未验证。无生产 capability JSON、无默认路由；不支持行插删与 ImportRange 尺寸复制，非空公式缓存/未知 parts 写前拒绝。 |
| sidecar/lo-convert | 同类格式转换已接入；`recalculate` 命令已实现，含缓存清理/完整计算标记、输出校验和 Ubuntu/WSL2 LibreOffice 合成 `SUM` 测试 | 仅候选实现；用户已实际验证 LibreOffice 基础公式计算可行性，不重复泛泛可行性调研；真实业务公式、Excel/WPS 一致性、复杂依赖、Windows COM、超时、目标 Linux/glibc 2.28 和离线运行时来源均未验收；不能宣称通用公式能力或生产默认。 |
| Excel COM / WPS COM 修改 worker | 当前 ExcelModifier 无对应已验收后端；S1-F1 另有 Windows native COM 重算 worker | 修改能力尚未实现。分别登记 Windows-only Microsoft Excel 与 WPS capability；不得由 S1-F1 重算或 XLS 转换基准推导出修改能力，也不得用于 Linux 依赖。 |
| ExcelCalculator | 尚未形成统一公共计算接口 | 未来独立能力；LibreOffice Calc、Excel COM、WPS COM 为候选，按后端/公式/平台验收，不改变当前读写修改默认。 |

细节和报告见根目录 EXCEL_ENGINE_ARCHITECTURE.md、PYTHON_EXCEL_OPERATION_INVENTORY.md 及各自 probe/reports。报告数字只适用于其中样本和断言。已经通过的窄范围实现继续保留，不因本规则接入而重构。

### 6.2 限定范围的性能证据（报告时间：2026-10-23）

证据来源：[Direct OOXML 与 Rust Excel 引擎独立基准报告](../../excel-engine-benchmarks/20261022-ooxml-vs-rust/REPORT.md)（SHA-256：`D33A786C90D1942FB66DE33E3203A5D9A136414DA89D92A96B6935E04B84A596`）。测试环境为 Ubuntu/WSL2 x86_64、ext4、Release Rust；Rust 1.99.1、calamine 0.37.0、rust_xlsxwriter 0.101.0；Python 3.12.3、python-calamine 0.9.0、openpyxl 3.1.5。有效组合均为通过正确性检查后的至少五次串行测量。证据限定于报告中的 XLSX 输入、实现版本、规模和明确输出契约，不外推到 Windows、麒麟、统信、XLS 或 XLSB。

本节性能证据只能按 `(data_owner, 操作/output_kind, input_format/output_format, 输出契约, 平台, 引擎版本)` 匹配。它用于候选评估，不登记为任何业务功能已通过，也不替代真实模板、目标平台或完整调用链验收。

#### 读取证据适用范围

- **适用路由**：`data_owner=rust`、`input_format=xlsx`，读取普通存储值；只有当输出坐标、类型、空值和值语义与报告 typed digest 契约一致时，Calamine 才是该具体功能的首选 reader 候选。`pure_same_range` 中 Calamine operation 中位数为 132.100 ms、进程墙钟 138.571 ms、峰值 RSS 中位数 17.3 MiB；`pure_full_sheet` 分别为 134.064 ms、142.904 ms、18.1 MiB。对不同范围、类型解释或返回结构须重新验收。
- Python `ooxml_external_prepare` 对应实现中，same-range 为 2161.393 ms operation / 2354.164 ms process wall / 183.3 MiB RSS；full-sheet 为 2148.653 ms / 2448.143 ms / 183.3 MiB。它是本报告内的具体 Python ZIP/XML 实现，不能代表所有 Direct OOXML 实现。
- 报告**没有**比较 Python `python-calamine` 与 Python Direct OOXML 的同进程端到端性能。`mixed_source_reader` 还包含额外预加载和元数据构造，不得据此推广 Python-owned 流程的 reader 排序；Python-owned 迁移期流程维持原项目同契约验收路由，调整前须测完整调用链。
- 条件格式规则元数据是单独的 `metadata-only` 请求：`ooxml_cf_public` operation 中位数为 2.467 ms、含启动/导入/校验摘要的 process wall 为 88.579 ms；Calamine 没有对等 API，因此不作 reader 排名。批注、条件格式、验证等 OOXML 专属字段应由已针对该对象和契约验收的专用适配器提供。Direct OOXML 普通读取路径拒绝部分含合并、条件格式、数据验证或公式的样本；能力不满足时报告缺项，不作为 Calamine 失败后的普通值 fallback。

#### 新建工作簿证据适用范围

- **适用路由**：`data_owner=rust`、`output_kind=new_report`、从零生成 `output_format=xlsx` 的纯数据报表。基准数据预先生成；被测规模为 10,000、100,000、300,000 行 × 20 列，另有一行表头，并分别包含纯数据/格式化输出。operation 包含数据文件加载/模型适配、writer 和 fsync，不含数据生成；具体功能必须复用其自身输出字段、样式和保留契约做对拍并定义资源预算。
- `rust_xlsxwriter` 使用默认内存工作表模式，报告没有测试 constant-memory。30 万行 × 20 列 same-level 纯数据样本：Python `ooxml_s5` 为 17.406 s / 288.2 MiB RSS，Rust `rust_xlsxwriter` 为 8.729 s / 1376.7 MiB；same-level 格式化样本分别为 19.801 s / 354.0 MiB 与 7.129 s / 1376.7 MiB。Rust writer 的 1,376.7 MiB 约为 1.34 GiB，是该规模/模式的实测观测，不是通用阈值或任意大文件的内存承诺。delivery-default 是另一配置组，须单独比较，不能与 same-level 数据混算。
- 报告中的 `ooxml_s5` / `ooxml_s2` 是 Python 专用 writer，不是 Rust Direct OOXML writer。报告没有比较 `rust_xlsxwriter` 与 Rust Direct OOXML writer，也没有测试通用 OOXML writer、既有工作簿编辑、任意精度 Decimal、图表/宏或 constant-memory 模式。Direct OOXML 只在存在已定义的专用输出契约、明确 OOXML 对象要求，或同契约实测证明资源收益时进入候选；不同实现不能仅凭名称、压缩配置或 ZIP 写入方式视为等价。

#### 环境边界与路由决定

- 环境为 Ubuntu/WSL2 x86_64、ext4、Release Rust；报告记录 Rust 1.99.1、calamine 0.37.0、rust_xlsxwriter 0.101.0，以及 Python 3.12.3、python-calamine 0.9.0、openpyxl 3.1.5。该证据只覆盖报告中的 XLSX 样本和契约，不外推到 Windows、麒麟、统信、XLS、XLSB 或其他未测平台/格式。
- 候选顺序必须与已验收路由分开登记。库已安装、基准更快或候选顺序靠前都不等于具体业务功能、格式、对象、平台和输出契约已通过；本报告不改写现有生产默认。若评估修改既有默认，须先定义该功能可节省的绝对时间、加速比例、峰值内存上限和测试规模，再用同机、同输入、同契约的完整调用链测量决定。迁移期间如存在 Python/Rust 桥接，计入进程启动/导入、序列化、传输、数据转换、IPC 和校验成本。

### 6.3 第二份独立性能证据（报告时间：2026-10-23）

证据来源：[Open-Data-Audit Direct OOXML 与 Rust 引擎基准报告](../../3.open-data-audit-benchmark/results/benchmark_report.md)（SHA-256：`9BE7045378F1CDF2B45288D478933560DF915553CF3AD08A2D4490876552128A`；逐轮数据见同目录 `benchmark_runs.csv`，机器结果及环境见 `benchmark_results.json`）。这是独立于 §6.2 的第二轮测量：Ubuntu/WSL2 x86_64、Release、每项 5 次串行；Rust calamine 0.37.0、rust_xlsxwriter 0.100.0，Python 3.12.3、openpyxl 3.1.5。任务说明列出 `python-calamine 0.9.0`，但同目录 `benchmark_results.json` 的 `environment.python_calamine` 和基准 `README.md` 均记录为 0.3.1；本节暂按可追溯的机器结果字段登记为 0.3.1，实际运行时版本仍待基准环境复核，不把版本冲突外推为性能结论。`rust_xlsxwriter` 与 §6.2 的 0.101.0 不同，两份报告不合并、不平均，也不据此解释或修订彼此数字。


本节证据必须按 `(data_owner, 操作/output_kind, input_format/output_format, 输出契约, 平台, 引擎版本)` 匹配；下列数字只是对应样本和实现的观测，不是通用阈值、默认路由验收或所有 XLSX 文件的性能承诺。读取计时是报告所列 reader 对同一 XLSX 输入的测量，不含 Rust 写入路径中的 Python JSON 数据桥接；所有 Rust writer 的端到端数字均明确包含报告所述 Python JSON 编码、子进程启动、Rust JSON 解析及 XLSX 写入阶段，不能用 Rust writer 内部时间替代。

#### 读取证据适用范围

- **适用路由匹配键**：`data_owner=rust`、`operation=R01 ordinary_value_read`、`input_format=xlsx`、`output_format=JSON coordinate/value digest`、输出坐标/值摘要且范围与报告一致、Ubuntu/WSL2 x86_64、Release、calamine 0.37.0；对比实现是本报告的 Python Direct OOXML value provider，不代表所有 Direct OOXML reader。`synthetic_full_sheet` 为 50,000 行 × 20 列数据（含表头的报告 bounds 为 50,001 × 20）：Calamine 中位 1.043 s / 73.7 MB 峰值 RSS，Direct OOXML 中位 12.524 s / 770.1 MB 峰值 RSS；坐标/值摘要和 bounds 一致。**这些读取数字不含 Python JSON 数据桥接端到端时间**，且仅限该 reader 计时口径和合成 XLSX。（来源：报告“读取结果”与“读取一致性”。）
- `real_s2_xlsx_full_sheet` 的实际 S2 XLSX 全表读值分别为 Calamine 0.012 s / 4.0 MB 与 Direct OOXML 0.210 s / 29.8 MB；其 bounds 为 729 × 16，坐标/值摘要一致。匹配键仍限 `data_owner=rust`、`operation=R01 ordinary_value_read`、`input_format=xlsx`、`output_format=JSON coordinate/value digest`、同一普通值全表读取契约、Ubuntu/WSL2 x86_64、Release 和上述 reader 版本，不作为业务功能完整调用链证据。此项同样是 reader 计时，不含 Python JSON 数据桥接。（来源：报告“读取结果”与“读取一致性”。）
- `type_error_edge_probe` 的差异按 §9.1 处理，不因普通数据样本摘要一致而视为类型语义等价；格式/类型契约不同的读取路径必须单独验收。

#### 新建工作簿证据适用范围

- 报告的新建写入矩阵覆盖 `data_values_base_style` 与 `common_format` 两类 10,000、100,000、300,000 行 × 20 列样本；本节摘录 30 万行指标作为大样本观测，并单列真实 S2 三表导出。其余规模的逐项数据见报告“写入结果”，不能把本节摘录视为全矩阵结果。
- **适用路由匹配键**：`data_owner=Python-side payload via JSON bridge`、`operation=W01`、`output_kind=new_report`、`input_format=pre-generated in-memory data`、`output_format=xlsx`、报告对应的数据值或共同格式契约、Ubuntu/WSL2 x86_64、Release；Rust 候选为 rust_xlsxwriter 0.100.0 默认内存模式，Direct 对照为报告中的 Python S1 writer。该键标明迁移桥接边界，不等同于 Rust-owned 产品数据已验收。
- 30 万行 × 20 列 `data_values_base_style`：rust_xlsxwriter 两组 Rust 写入数字均为**含 Python JSON 桥接端到端口径**；ZIP level 1 中位时间为 16.631 s / 1,588.2 MB 峰值进程树 RSS，Python Direct 为 20.356 s / 376.6 MB；`common_format` 的 Rust 数字同为含 Python JSON 桥接端到端口径，为 14.722 s / 1,594.0 MB，Python Direct 为 19.121 s / 379.8 MB。Rust 端到端口径包含 Python JSON 编码、子进程启动、Rust JSON 解析及写入；Direct 数字是该 Python 专用 S1 writer 的端到端结果。内存数字只是这两个 30 万 × 20 样本的观测，不是通用上限或任意大文件承诺；报告未测 constant-memory。（来源：报告“写入结果”中 `data_values_base_style_300000x20` 与 `common_format_300000x20` 项。）
- **适用路由匹配键**：`data_owner=Python S2 report payload`、`operation=actual_s2_three_sheet_report`、`input_format=in-memory payload transported by JSON bridge`、三表新建 `output_kind=new_report` / `output_format=xlsx`、同报告 payload/样式/值摘要、Ubuntu/WSL2 x86_64、Release、Rust rust_xlsxwriter 0.100.0。首次结果中 Rust ZIP level 1 的 0.040 s 为**含 Python JSON 桥接端到端口径**，Python Direct S2 ZIP level 6 为 0.758 s；但 Direct 的该次总时间包含每个新 worker 首次导入 exporter，且压缩级别不同，故不可据此作稳定运行时间胜负结论。报告热路径复测为 Direct 0.076 s（0.071–0.082）、Rust 0.050 s（0.049–0.054）；该 Rust 数字同样为**含 Python JSON 桥接端到端口径**。Rust 输出仍有 17,580 格中的 23 格浮点值相差 1 ULP，严格摘要不等。（来源：报告“写入结果”及“补充复测：实际 S2 写入热路径”；1 ULP 差异见附录“实际 S2 写入：浮点末位（1 ULP）序列化差异”。）
- S2 三表输出的业务/值保留契约尚未通过严格浮点对拍；S3 是修改既有工作簿并保留模板部件，未由本报告与新建 rust_xlsxwriter 作同类写入比较。保留 S2/S3 专用路径，不由新建 writer 的性能数据推导替换结论。

#### 候选评估注记与范围

- 本轮读值数据使 Calamine 值得按 §9.2、§9.3 的完整调用链、输出契约及资源门槛继续评估；它是候选注记，不改任何既有读取默认。
- 新建写入暂不因本报告切换默认：30 万行样本中的 Rust 峰值 RSS 明显高于 Python Direct，且 S2 浮点严格值摘要未通过。S2/S3 专用实现按其不同输出/模板契约保留；任何替换仍须先定义节省时间、加速比例、内存上限及规模，并按 §9.2、§9.3 完成同机同输入完整流程对拍。
- 本证据限 Ubuntu/WSL2 x86_64、报告列出的 Release 版本、XLSX、样本规模及契约；不外推到 Windows、麒麟、统信、XLS、XLSB、ET、constant-memory 或既有工作簿通用编辑。XLS 侧测只用于 §7 的伪 `.xls` 预期拒绝和 §9.1 的类型/空值口径观察，不纳入 Direct OOXML 与 Calamine 的读取排名。

## 7. 路由、临时输出与安全回退

1. 解析契约、真实格式/对象、公式口径与外部引擎要求。
2. 按 §5 匹配操作及子类。
3. 只把适配器已实现且 passed、格式/平台/版本/保留项/资源范围匹配的后端加入候选。
4. 手动模式只执行所选合格后端；不可用或失败即停止。自动模式仅从合格候选中选择。
5. 执行到临时文件，按契约验证后才提交最终路径。

读取路由特例：R01 普通存储值读取遇到 Calamine 错误时，不得转为 Direct OOXML 通用重试；OOXML 专属字段是单独请求的元数据能力，只能由对应对象已验收的专用适配器提供。除该限制外，自动回退只允许执行前发现能力/平台不匹配，或引擎在写出前明确报告不支持该文件特性，且有另一完全合格后端；回退必须从原始输入或干净副本重跑并记录原因。输入损坏、加密、权限/占用、空间不足、输出验证差异、公式错误、缓存未刷新、超时或未知异常均停止，不换引擎掩盖问题。不得在部分写入或疑似损坏的中间文件上继续执行。覆盖原文件须有明确授权与备份/哈希保护。日志记录引擎、版本、耗时、路由/回退原因和未满足项。

格式签名无效的伪 `.xls`（包括实测到的 HTML 伪装导出，CFB/OLE 签名无效）必须在路由前以“格式不支持”显式拒绝并记录为预期拒绝；不得计为引擎失败，也不得触发备用引擎回退。依据：第二份独立基准报告附录“real_xls_side_test”复核。

## 8. OOXML 安全边界

OOXML 是补丁/缺失元数据层，不是主力完整读写库。必须使用 namespace-aware XML parser；禁止正则替换 XML 全文。

- 由 workbook.xml 和 .rels 解析真实 sheet 身份到 worksheet part 的关系；不假设 sheet1.xml。
- 维护操作涉及的 Relationship、[Content_Types].xml、sheet 身份/顺序、sharedStrings、styles、defined names 和公式引用。
- 只对明确、已验收子集修改；区分单元格类型、共享/内联字符串、样式 ID、缺失单元格、共享/数组公式及缓存。
- 批注可能关联 comments、VML、threaded comments、relationships 和 Content Types；未经专项验收不得改写或新增。
- sheet/行列/合并/名称/超链接/页面设置等变更须检查跨部件引用。宏、签名、外链、加密、未知扩展或非标准变体不因“未主动修改”就自动安全；超出白名单则拒绝。
- 未改 ZIP entry 的解压后语义按契约对拍；整个 ZIP 字节哈希可因重压缩而变化。
- 每个适配器须覆盖该操作涉及的全部保留项；未验证部件直接拒绝。先做小能力和拒绝规则，不实现完整编辑器。

## 9. 对拍、性能与默认切换

### 9.1 功能对拍

- 读取：probe 导出 JSON，对比值/类型、坐标、sheet 名/顺序、合并、公式文本/缓存、批注、条件格式、数据验证等本次字段。
- 新建：独立 reader 与 Python 基线对比值、类型、公式、样式、合并、名称、页面设置和文件可打开性。
- 修改：验证目标变更、未修改区域/工作表、公式、批注、样式、合并、名称、尺寸/隐藏、页眉页脚、页面设置、relations 和独立 reader 重开。
- 打印仅逐项断言 §1.2 页面设置；不做 PDF 渲染验收。
- 公式旁路验证公式文本、缓存刷新、旧项目计算基线；无引擎时验证报错且不阻塞其他功能。
- 使用真实格式/对象样本；不得用改扩展名的 XLSX 替代 XLS/XLSB/ET。只读源输入前后 SHA-256 不变。对象逐项记状态；未执行的目标机验收写明未执行。
- 文件能被独立 reader 重开不等于所有 Excel/WPS 保真项通过。
- **跨引擎浮点值不得要求位精确相等**：§6.3 的真实 S2 对拍中，rust_xlsxwriter 与 Python `repr` 序列化存在 1 ULP 差异（约 17,580 格中 23 格，集中于一列）。跨引擎/跨实现的浮点值按相对容差不大于 `1e-15` 比较，或先按业务约定归一为十进制字符串再比较；容差与十进制精度须写入该功能契约。该规则适用于 `excel-reader`、`excel-editor` 及其他跨引擎对拍与验收。SHA-256 逐格摘要只用于同一引擎的前后状态比较，不得用于跨引擎浮点值判等。依据：第二份独立基准报告附录“实际 S2 写入：浮点末位（1 ULP）序列化差异”。
- **读取适配器必须声明值解释口径**：Direct OOXML `CellValueProvider` 返回原始缓存值（bool 为 1/0、日期为序列号、错误值为缓存字符串）；Calamine 返回解释后的 bool/datetime/error 类型。适配器登记必须写明原始缓存或解释类型、日期/错误值及空值语义；声明前不得把另一口径作为替换读取路径的对拍基线。依据：第二份独立基准报告 `type_error_edge_probe` 及附录“取值口径差异”。
- **跨语言空单元格先归一再比较**：probe/对拍管道比较 python-calamine 与 Rust calamine 前，须把 Python `''` 和 Rust `Empty` 映射为已声明的同一空值状态，再按坐标比较；禁止将绑定层空值记号差异误报为数据差异。依据：第二份独立基准报告附录“real_xls_side_test”中的 14 × 21 样本，Python 侧共返回 294 个非 `None` 值，其中 214 个非空字符串、80 个空字符串；Rust 侧为 214 个 string + 80 个 Empty，实际非空字符串一致。

### 9.2 性能优化

1. 先固定现有实现、输入、输出契约和基线；定位读取、OOXML 补读、业务计算、新建、修改、转换、worker/IPC、校验或端到端热点。
2. 只检查当前操作及直接依赖的重复打开/保存、整表副本、逐单元格 IPC/COM、重复矩阵构造、不必要转换、宽泛异常回退和语义缺口；不扩展为无授权全仓重写。
3. 同机、同输入、同输出契约预热后至少五次，报告中位数、范围、机器/库版本、样本规模、峰值内存和输出大小；测量必须覆盖该功能完整调用链，同时提供分阶段计时，并计入 Python/Rust 桥接的进程启动/导入、序列化、数据传输、数据转换、进程/IPC、隔离 runtime 和校验成本。允许 deb 内置 openpyxl worker 专用 Python runtime；不得依赖系统 Python。
4. 写入使用同一份预生成数据；分别记录同压缩配置和实际发布配置。单次微基准或仅 Rust 内部耗时不能支持业务提速结论。
5. 考虑修改现有默认路由时，测量前先定义该功能的绝对节省时间、加速比例、峰值内存上限和测试规模，再按这些指标逐项判断；不能测后改门槛或编造比例。保留项/业务语义未通过时，不得因性能切换默认。
6. 兼容修复凭缺陷复现、目标契约和回归证据决策，不虚构性能收益。未有证据时保留已验收路径。

### 9.3 成为默认候选

旧工作区曾以对拍比例作为切换默认候选的门槛；它不作为本项目人工测试版本的交付前提。实际变更应按当前功能、格式、对象和平台保留适用证据，未经验证的路由不得宣称已验收。原位编辑须保护保留项并避免损坏；输出无法正常打开时记录问题，不把失败报告为成功。

## 10. 旧工作区操作盘点方法（按当前功能选用）

以下方法保留为历史参考；只在当前 Audit 功能需要时核对相应调用链，不要求先完成全项目清单，也不依赖旧工作区的文件。不能将历史模板或调用链记录冒充当前项目验收。

需要追溯的具体 Excel 操作可按适用范围记录：

1. 当前 Audit 菜单/入口 → 函数 → helper → 实际读写。
2. 输入输出、配置、格式、write_pattern、公式口径、保留/丢弃项、前置条件。
3. 分类为只读、新建、原位编辑、复制、格式转换或文件系统操作；编辑路径形成完整文件级 ExcelEditPlan。
4. OOXML Patch、umya、Excelize 和 openpyxl 按操作子类登记状态、版本、格式/对象、平台、入口、隔离 runtime、证据和性能。
5. Python/Rust 基线命令、JSON 比较字段、样本来源、输入哈希、错误语义及安全回退路径。
6. 未查明行为标 unverified；先选择真实存在、价值最高的操作做小范围验收。

重点确认：缺失合并信息不等于无合并；replace/rebuild 不等于追加；改 sheet/删行列需查公式、名称和其他 part；检查与执行分支分别核实；同名函数不代表两套产品行为相同。

按根 AGENTS.md 和 DEVELOPMENT_GOAL.md，以实际 Audit 功能为单位实施、静态检查并交付用户人工构建测试；编译、测试及打包需用户明确授权。遇到引擎能力缺口，在本项目记录证据和受影响操作，并按业务需要定点处理，不向旧 Main 同步。Windows 和 Linux/deb 验证分别记状态，不覆盖已有有效证据。

## 11. 既有规则冲突检查与可复用内容

本次检查了 workspace 根 AGENTS.md、根 README、Cargo/crate 结构、EXCEL_ENGINE_ARCHITECTURE.md、PYTHON_EXCEL_OPERATION_INVENTORY.md、audit-core/rules/03、07、13，并只读检查两个原始项目规则目录。

- 接入前 workspace 根没有 rules/；当前 workspace 的 03、07、13 位于 audit-core/rules/。根目录与 audit-core/rules/ 均没有 12。Open-Data-Audit 原项目存在 rules/12_Rust优化与引擎路由.md。
- Audit 原规则 12 可复用“不假设 Rust 必然更快、保留已有验收路径、按阶段测量与对拍再切默认”；S1/S5 菜单/系统编号、迁移顺序和旧 Office 行为不移入 workspace。
- py_statistics_tools 原 rules/ 可复用输入输出契约、操作入口/dispatcher 清点、错误分类、混合进程与端到端成本测量。其 openpyxl/pandas/xlrd/pyxlsb/Python-calamine 生产候选不符合本 workspace 目标，只作为旧基线证据。
- audit-core/rules/03 中 namespace-aware parser、关系、defined names、sheet 身份/顺序和保守拒绝原则兼容，继续作为 Audit 特定补充；其中广义 Python/Office 路径不得扩大本规范的 sidecar 白名单。
- audit-core/rules/07 的 pytest/S1-S5 分组适用于原 Audit Python 项目；Rust workspace Excel 后端验收按本规范与根 AGENTS.md，纯文档任务不要求运行 Python 全量 pytest。
- 根 README 原先将 umya 描述为既有工作簿修改方式，容易被理解成 workspace 默认；现已改为 ExcelEditPlan/ExcelModifier 按整批已验收能力路由。audit-core/rules/13 与本文件高度重叠；保留其 Audit 业务映射与细节，通过链接指向本文件。Audit 旧 W01 表格曾将 Direct OOXML writer 排在 rust_xlsxwriter 前；workspace 当前规则将 rust_xlsxwriter 列为 Rust-owned 纯数据新建 XLSX 的优先评估候选，并要求输出契约与内存预算通过后再采用。Direct OOXML 只在本规范 §4.2 条件下进入候选；以上均不构成全局已验收默认。Audit 规则保留其业务映射并通过链接遵循 workspace 证据边界。
- 两个原始项目和 workspace 根目录来源稿均为只读参考，本次不改源项目。规则冲突按本节和根 AGENTS.md 处理，不将旧项目业务/平台规则静默照搬。

## 12. 规则入口与留痕

- 涉及 Excel/CSV 读写、工作簿修改、格式转换、打印页面设置、后端路由或性能优化，先读本文件，再按需读根 AGENTS.md、EXCEL_ENGINE_ARCHITECTURE.md、PYTHON_EXCEL_OPERATION_INVENTORY.md 及 audit-core/rules/03、07。
- 修改前检查 Git/WIP 和现有报告，最小合并，不覆盖未提交内容。
- 每次实际变更按根 AGENTS.md 与 audit-core/rules/08_AI协作与修改记录.md 适用要求，在 GPT修改记录/ 新增记录，列明文件、目的、测试/未执行项、问题、风险和回滚。纯文档/治理不升级产品或组件版本、不写用户版更新日志。
- 不删除 Python 对拍基线、不修改两个原始 Python 项目、真实配置或业务数据，不做无关重构。

