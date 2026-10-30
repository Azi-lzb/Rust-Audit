# Rust-Audit 独立开发目标与状态

更新：2026-10-30（S1-F1 Windows 原生 COM 批量性能优化已完成最终验收；条件格式异常指标定位、结构预检/条件审核共用源读取及模板/源结构预检投影读取优化已完成；源结构预检同输入 Rust 对照为 501.532→237.579 ms，结构判断逐项一致；完整审核链三轮小型合成端到端复测：Python/Rust 业务步骤中位数 7,226.75/6,890 ms（1.049×，近似持平），结果、公式缓存、规则表批注作者和输入保护均通过；worker 墙钟中位数 32,747.97/6,965.34 ms，Python 值包含新 Excel 进程退出等待，不能解释为业务计算 4.70×；峰值工作集 502.141/297.863 MB；Rust acceptance 已补微秒阶段计时；并发负载下单轮诊断显示 Excel 重算约 3.38 秒、会话关闭约 2.93 秒，不纳入正式倍率基线；S5 配置审核仍按用户要求暂停开发。COM 批量结果见 [S1-F1 COM 最终验收报告](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.md)，源结构投影对照见 [专项报告](audit-core/probe/reports/s1_structure_source_projection_20261030.md)。）

状态只使用“已实现，待人工验收”“部分实现，需要继续开发”“尚未实现”“已明确暂缓”或“存在已知问题”。同一功能可有主状态和已知问题；代码、入口或局部测试不等于完整验收。实现、单元/冒烟、Windows 真实样本人工验收和 Linux/deb 分别记录。

## 2026-10-30 增量：S1、S2、S3 性能地图与持续优化

- **范围**：本持续目标覆盖 S1-F1/F2、S2-F1、S3-F1 至 F5。S5 已停止开发并排除在本目标外；S1-F1 原生 COM 批次与 S3-F1 的性能优化按用户指定暂时结项，不重复已完成比较。S1-F1 条件格式异常提取另发现生产结构区域指标定位瓶颈，本次已定点优化。
- **性能台账**：已建立 [性能基线](audit-core/PERFORMANCE_BASELINE.md)、[性能历史](audit-core/PERFORMANCE_HISTORY.md) 和 [benchmark 索引](audit-core/benchmarks/README.md)。基线先汇总已有 2026-10-29/30 的报告、机器数据及最终 COM 验收，不重复运行既有对比。
- **当前可用结论**：S1-F1 原生 COM 批次 2/10/33 册均与 Python 平价；S3-F1 最新优化报告为约 1.02×，两项暂结。S1-F1 带结构区域的条件格式异常流程中，指标组装由 15,523 ms 降至 57 ms；结构预检+条件审核 Rust developer event 耗时（不含进程启动）由 16,044 ms 降至 498 ms 后，进一步通过共享同一次源工作簿读取降至 390 ms（比前版再省 108 ms / 21.7%）；新状态五轮 8,610 条异常摘要完全一致，峰值工作集约 85.95 MB，与原状态相当。随后为结构模板增加专用 Calamine 投影读取，在 20k 合成结构/条件样本中模板预检读取由 528 ms 降至 224 ms（-304 ms / 57.6%）；三轮 43,051 条异常摘要与旧版完全一致，峰值内存仍由条件审核全量读取主导。S3-F3 已移除原地导入中临时 ExcelEditPlan JSON 的磁盘往返，同负载 Rust 内部中位数从 946 ms 降至 203 ms（减少 78.54%），优化后 Python/Rust CLI 墙钟为 912.62/235.54 ms，3.88×；输出、备份和输入保护对拍无差异。另有 1,000 条全冲突压力样本：Python/Rust CLI 1,332.49/149.89 ms（8.89×），冲突清单与目标工作簿值/结构无差异，Rust 冲突清单写出阶段约 18–21 ms。S1-F2、S3-F2/F4 在已测合成规模明显更快；S3-F2 新增 2,610 行输出规模剖析，Rust CLI 墙钟快 8.11×且 3/3 全输出对拍一致。S3-F4 的 10k 行、20k 输出结果在三轮 340,017 单元格对拍无差异下快 7.15×（CLI）；S3-F5 修正基准后，在 30×40×4 的有效回填负载下快约 2.28×（CLI），30 个工作簿逐格对拍无业务差异；Python/Rust 进程树峰值工作集约 475/24 MB。S2-F1 外部规则正常/异常样本在常驻服务计时中快约 1.75×，独立 CLI 计入启动后慢约 18%；三轮输出逐格与结构对拍无业务差异。S1-F1 完整链路最近三轮合成对照为 Python/Rust 业务步骤 7,226.75/6,890 ms（1.049×，近似持平），峰值工作集 502.141/297.863 MB；结果、公式缓存、规则表批注作者、Excel 打开及输入保护均 3/3 通过。Python worker 墙钟包含新建 Excel 进程退出等待，存在显著轮间波动。
- **数据质量说明**：2026-10-30 的原始 JSON 与汇总 Markdown 在 S2-F1、S3-F1/F2/F3 等行存在不一致；台账优先从逐轮 JSON 重算中位数，并单列 S3-F1 优化后逐轮数据缺失。另发现 2029/30 两份 S3-F5 旧基准表名导致 0 项业务命中，已将旧数字标为无业务代表性；当前基准生成器改用 A1 表名并设有命中数量门槛。历史报告和逐轮 JSON 保留原样，不覆盖。
- **尚缺性能证据**：历史记录有 S2 42 份 XLS 真实批次和 S3-F2 真实工作簿的阶段观察，但 S2 原始 21+21 份输入未保留，不能复跑同输入 Python/Rust A/B；旧 S3-F2 真实 2,610 行工作簿也未保留，不能解释其 1,186 ms 写出阶段对应的输入结构。S3-F2 合成放大剖析现包含阶段耗时与峰值工作集（Python/Rust 43.688/44.340 MB），没有显示内存优化空间；S1-F2 轻负载峰值工作集为 Python+Excel COM/Rust 300.078/10.719 MB。新增 Rust-only 阶段诊断记录了 S1-F2 的发现、读取、计算、写出和验证时间，但测量时检测到 21 个后台进程，明确不纳入性能基线；Python/Rust 三轮总耗时和完整输出对拍仍由既有报告支撑。S3-F1 10k 两期 CSV 峰值工作集为 Python/Rust 133.477/119.035 MB，耗时结项不变。S2 外部规则、S1-F2、S3-F1、S3-F2、S3-F4 和有效 S3-F5 专项记录了进程/进程树峰值工作集。S1-F1 完整业务链路已完成三轮合成耗时和峰值工作集测量；acceptance 现已提供请求准备、预检、复制/编辑、重算、异常提取、结果输出和 COM 会话关闭微秒计时，阶段剖析目前为并发负载下单轮诊断；其他 S1/S2/S3 功能仍有部分阶段需要同口径 A/B 记录。结构预检与条件格式现共用源工作簿的一次读取，4,000 行样本单次解析约 93 ms；尚未把这份投影跨公式副本/重算阶段保留，也未验证多文件场景的内存扩张。性能台账标注了数据规模、输入文件字节、输出正确性验证及其缺失项。
- **S1 读取证据边界**：历史脱敏真实工作簿仍记录 Python 77 ms、Rust 读取 983 ms（1,075 个固定标签）；该输入不可从当前夹具复跑，也没有文件字节数/峰值内存。本轮同输入 Rust 路径对照显示源结构预检完整读取+比对 501.532 ms、投影 237.579 ms（2.11×），五轮结果一致；模板预检另由 528 ms 降至 224 ms。两项都是合成数据的局部读取路径证据，未证明历史真实文件或完整 S1-F1 链路有相同比例提升；源投影对照未采集 RSS。
- **S1-F1 同模板批量优化**：模板/结构预检按 canonical 模板路径在单批次内缓存，避免每本源文件重复读取同一模板；模板结构读取现使用专用 Calamine 投影，在 20,000 行 Rust-only 样本中由 528 ms 降至 224 ms。三文件 GUI 合成路径均通过结构预检，因本机 Excel COM worker 未启动且无 LibreOffice，按预期保留人工重算副本。20k Python A/B 因参考路径耗时过长中止；本轮无跨引擎倍率，也未覆盖公式重算成功的完整 GUI 批次；后续三轮 S1-F1 全链路合成对照已通过，见 [报告](audit-core/probe/reports/s1f1_full_chain_20261030.md)。
- **S1-F1/F2 模板自动匹配元数据复用**：多文件 GUI/汇总入口原先逐文件扫描模板目录并重复读取每个模板的 Sheet 目录；现改为每次批次加载一次模板档案，源簿仍逐本读取并独立匹配。21 个源簿、5 个模板簿合成对照中，Release profile 中位数由 30.281 ms 降至 5.908 ms（省 24.373 ms，5.13×），21/21 建议完全一致。此项只减少匹配子步骤约 24 ms，不代表整项功能达到 5×。
- **本增量状态**：新增 S2-F1 外部规则、S3-F4 输出值级和 S3-F5 有效回填性能剖析；修正当前 S3-F5 基准生成器的 A1 表名映射，并为专用驱动添加非零命中门槛。新增 S2-F1 配置读取复用及输出轻量验证优化：配置解析 40→15 ms，21+21 合成批次保存阶段 312→98 ms、服务总耗时 659→429 ms；相关对拍通过。历史真实 21+21 文件已不可用，合成结果不替代真实文件验收。S3-F1/F5 历史数字的基准有效性限制已记录。S1/S2/S3 性能地图、现有瓶颈排序及有明确收益的优化均已完成；S1-F1、S1-F2 阶段诊断中的并发环境限制，以及缺失的历史原始输入/数据均在性能基线中标出。本性能目标按当前可复核证据结项；S5 仍暂停且不在本轮范围。

## 2026-10-30 增量：S1-F1 结构预检与条件审核复用源读取

- **修改**：S1-F1 结构预检与条件格式审核针对同一本源工作簿共用一次 `WorkbookDump` 读取；结构比较和条件格式异常提取都使用该只读投影，随后即释放，不跨文件缓存，也不跨公式副本/COM 重算阶段持有。单文件 CLI、目录审核和条件审核入口均已接入。
- **实测**：4,000 行合成文件、32,008 个非空格、4 条条件格式规则、16,000 个扫描格、8,610 条异常，预热 1 轮、交错 5 轮。Rust developer event 内部耗时（不含进程启动）从前一实现 498 ms 降至 390 ms，减少 108 ms（21.7%）；结构/条件处理共用的源读取从约 99+97 ms 降至 93 ms 一次。Python/Rust 内部计时为 19,222.839/390 ms；49.29×仅适用于该专项。前后 XLSX 容器大小差 1 byte，数据与结构由同一生成器产生。
- **正确性和内存**：五轮每轮 8,610 条异常规范化 SHA-256 多重集完全一致，结构标签 8/8 匹配；本轮源和模板测量前后哈希不变。峰值工作集为 Python/Rust 77.312/85.949 MB，Rust 前一实现约 86.562 MB；没有观测到共享读取增加峰值内存。
- **验证边界**：定向 Rust 测试及 Release 构建通过，完整 S1-F1 GUI 链路、真实机构文件、全工作区回归和 Linux 环境没有在本轮验证；总性能目标继续进行。详见 [专项报告](audit-core/probe/reports/s1_source_read_reuse_optimization_20261030.md) 和 [五轮原始数据](audit-core/probe/reports/s1_preflight_reads_20261030_shared_read_after.json)。

## 2026-10-30 增量：S1-F1 条件格式异常提取性能剖析

- **负载与口径**：在 Windows 使用同一生成工作簿，4,000 行、32,008 个非空单元格、4 条条件格式规则、16,000 个扫描单元格；预热 1 轮，随后交错运行 5 轮。这个早期条件格式微基准没有把模板的 `structure_ranges` 传给两端，故未执行生产指标定位；计时为 Python/Rust 服务内部时间，不含子进程启动或完整结果 JSON 序列化，另以独立子进程采样峰值工作集。
- **结果**：Python 中位数 1,343 ms，Rust 246 ms，约 5.46×。Rust 读取/判定/异常组装中位数为 102/94/24 ms；Python 规则规范化/判定/异常组装约 62/844/453 ms。Rust 峰值工作集约 85.53 MB，Python 约 87.94 MB。
- **正确性与决策**：五轮均生成 8,610 条异常，异常字段逐项一致、诊断数为 0，源工作簿 SHA-256 未变化。该合成规模没有发现 Rust 瓶颈，不改业务代码。不过此前真实小负载单次观测为 Python 375 ms、Rust 618 ms，其中 Rust 读表 578 ms；该输入路径未保留、无法做重复对照，不能据此直接改读表路径，需有可复测业务样本再决定。此专项只覆盖条件格式子步骤；在该专项完成时，S1-F1 完整 GUI/业务链路仍未测量，后续已补三轮全链路对照（见 `probe/reports/s1f1_full_chain_20261030.md`）；COM 最终批次性能结论不变。
- **后续剖析候选**：当时静态路径显示 S1-F1 结构预检与条件格式评估各自读取源工作簿；下一节已针对真实结构区域测得两次读表成本，并定位到指标前缀重复扫描。跨公式副本/重算阶段缓存整簿投影的内存代价仍未测量。
- **验证范围**：专用性能驱动本身完成五轮对拍；使用既有 Release CLI SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。未运行 Cargo build/test，也未做全工作区回归；总性能目标继续进行。

## 2026-10-30 增量：S1-F1 条件格式指标定位优化

- **瓶颈**：结构区域启用时，每条触发异常都会重新扫描本列从结构起点到当前行的所有前缀标签。4,000 行样本中约 8,610 个触发项让 Rust `issue_build` 中位数达到 15,523 ms；Python 参考流程也存在相同的重复前缀读取。
- **修改**：Rust 按工作表、目标列对触发位置分组，每列按行一次累计相同的有序去重文本标签，再按原触发顺序组装异常。保留左侧行标签、上方列标签、合并单元格锚点、数值过滤、去重顺序和字段拼接，不改审核规则或 Python 项目。
- **性能**：相同 4,000 行结构区域样本，Rust 异常组装中位数 15,523→57 ms（约 272.3×）；结构预检+条件格式阶段 CLI 中位数 16,044→498 ms（约 32.2×）。优化后 Python/Rust 分别为 19,264.7/498 ms（38.68×，Rust 计 CLI 墙钟、Python 不含解释器启动）。峰值工作集约 86.61→86.56 MB。
- **正确性与验证**：五轮每轮均得到 8,610 条异常；规范化异常字段 SHA-256 多重集 5/5 相同，8/8 结构标签匹配，源文件及模板 SHA-256 不变。`cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture` 通过 5/5；`cargo build -p audit-core --bin open-data-audit --release` 成功。未运行全 workspace 回归，也未执行真实机构文件或 GUI 手工验收。
- **剩余项**：结构预检和条件审核各自读取源文件约 99/97 ms。因两阶段之间会执行副本复制和 COM 重算，当前不跨阶段保留整簿数据；需在完整链路和多文件内存样本中量化后再决定。报告及前后逐轮 JSON 见 [性能基线](audit-core/PERFORMANCE_BASELINE.md) 和 [专项报告](audit-core/probe/reports/s1_indicator_resolution_optimization_20261030.md)。

## 2026-10-30 增量：S3-F3 原地写回计划传递优化

- **瓶颈和修改**：S3-F3 原地执行原先先将已构造的 ExcelEditPlan 写入临时 JSON 并同步磁盘，再立即读回交给 Excel-Core。现在执行路径保留经过校验的内存计划并直接交给 Excel-Core；显式“生成计划”入口仍按原契约持久化 JSON。匹配、冲突处理、输入哈希保护、备份、原子发布及输出规则不变。
- **优化前后**：同一 1,000 条唯一反馈、无冲突负载，Rust 内部中位数 946→203 ms（减少 78.54%），CLI 墙钟 975.25→235.54 ms（减少 75.85%）；优化后 Python/Rust CLI 墙钟为 912.62/235.54 ms。Python 与 Rust 输出差异为 0；备份 SHA-256 与原目标一致，仅 H3:H1002 发生业务值变化，反馈源保持不变。
- **验证**：定向库测试 7/7（新增执行冲突路径用例）、计划 CLI 集成测试 3/3 通过，Release CLI 构建成功。未运行全工作区回归；全仓格式检查仍因已有格式差异未通过，没有格式化无关代码。工作区未提交、未 push。
- **剩余证据**：已补 1,000 键、100% 冲突的合成压力样本；该结果不代表真实冲突率。实际机构文件规模及 GUI 常驻服务负载仍待测。详见 [性能基线](audit-core/PERFORMANCE_BASELINE.md)、[历史](audit-core/PERFORMANCE_HISTORY.md)、[无冲突报告](audit-core/probe/reports/s3f3_parity_optimized_20261030.md) 和 [冲突报告](audit-core/probe/reports/s3f3_conflict_profile_20261030.md)。

### 2026-10-30 增量：S3-F3 冲突输出剖析

- **补充负载**：对 1,000 个匹配键分别从两份反馈表提供不同原因，预期 1,000 条冲突、0 条导入。Python/Rust CLI 三轮完整输出对拍均为 0 差异；冲突明细值/结构一致，目标原字节不变、备份 SHA-256 与目标一致，反馈输入未变。
- **性能**：Python 服务中位数 1,332.49 ms，Rust CLI 墙钟 149.89 ms，约 8.89×；Rust 服务内部 120 ms。Rust 冲突清单写出 18–21 ms；独立进程树峰值工作集约 41.04/16.39 MB（Python/Rust）。100% 冲突是压力样本，不能外推为生产冲突率。
- **可观测性及验证**：执行报告现暴露 `conflict_write_time_ms`，CLI JSON 可直接记录冲突清单写出阶段。S3-F3 定向库测试 7/7 通过，计划 CLI 集成测试 3/3 通过；当前 Release 构建成功，SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。全工作区回归未运行；代码和文档未提交、未 push。

## 2026-10-30 增量：S1-F1 COM 最终性能与批量稳定性验收

- **性能优化状态：完成。** Python `DispatchEx` 与 Rust 原生 COM worker 均创建独立 Excel 实例；2/10/33 册批次均只创建一个 Excel 实例。最终 Release 二进制 A/B 中位数为 Python/Rust：2 册 7,570.60/7,171.91 ms，10 册 12,246.33/12,139.69 ms，33 册 27,006.11/27,138.94 ms；Rust 在不同负载下交替领先，差异约 0.5–5.3%，结论为实测平价。
- **差距来源：** Excel 新实例冷启动约 2.4–2.5 秒；COM 激活中位数随样本/负载变化（本轮 `DispatchEx` 2,524 ms、`CoCreateInstance` 2,391 ms）。Rust worker `spawn` 与 ready 协议残差合计约 43 ms/批。应用属性设置约 74/73 ms，批内逐册打开、计算、保存、关闭及缓存验证持平。旧端到端口径只计 COM 释放约 0.2 秒；本轮完整计入 Excel 进程退出等待，两边约 2.7 秒。
- **正确性与稳定性：** 18 个正式批次（每引擎 9 批）共验证 96 个输出/引擎、576,000 个公式结果/引擎；Python 与 Rust 的公式文本哈希、结果哈希配对差异为 0，6,000 公式/7,200 个缓存值样本全部通过。每批仅一个专属 Excel PID；18 个 worker 所有权 PID 均正常退出，Rust worker 均退出，15 个基线 Excel PID 保留。监视快照另见短暂的非本 harness 隐藏 Excel PID，之后查询时已退出；未对其采取清理。无强制终止、原输入哈希变化或错误结果。损坏工作簿单文件失败后，批次继续处理其后的有效文件。
- **超时：** 定向 Windows Excel COM 测试 1/1 通过，受控响应延迟超过 35 秒后，父进程结束 worker 和确切归属的 Excel 进程，源文件未变且未发布输出。该测试模拟计算成功后的响应延迟，不能证明 Excel 内部 COM 计算永久卡死时的所有行为。
- **回归与构建：** Release 构建成功；全 workspace 82 个测试目标合计 501 passed / 0 failed / 34 ignored，另有两项真实 Excel COM integration tests 分别单独运行通过。忽略项包含需要外部业务样本或连接的测试。最终报告与原始逐轮 JSON：[`报告`](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.md)、[`数据`](audit-core/probe/reports/s1f1_com_final_acceptance_20261030.json)。
- **尚待业务验收：** 本轮合成工作簿验证不替代机构文件上的 S1-F1 GUI 人工业务验收。S1-F1 性能优化已完成；没有发现值得以牺牲隔离或正确性换取的低成本优化。S5 仍按用户要求暂停开发。

## 2026-10-30 历史性能优化与复测记录（早期微基准，非最终批次验收口径）

- **S1-F1 Windows 原生 COM 热会话**：GUI/CLI 批次使用 Audit 自身的 COM worker，一批两册复用同一 Excel 进程；通过 Release 构建。定向合成会话测试 1/1 通过，确认进程 PID 复用、公式结果有效、源文件字节未变。没有重做完整 S1-F1 真实 GUI 验收，因此不把这项测试当作整链路人工验收。
- **S1-F1 性能范围**：早期三轮微基准测量两册、每册 7,200 个公式的 COM 重算/保存子步骤；报告中 Rust 4,950.6 ms、Python `Application.CalculateFull` 3,547.7 ms（0.72×），后续优化记录另列 Rust 4,668.5 ms 对 Python 3,595.5 ms（0.77×）。该组数字是特定热会话测量，不覆盖模板匹配、公式复制、最终审核结果输出；它与上方最终批量验收不是同一夹具和计时边界。最终验收每册为 6,000 个公式和 7,200 个非空缓存，批次 2/10/33 册均实测平价。两组证据均只涉及 COM 重算子步骤，不代表 S1-F1 整条业务链路。
- **S1-F2 性能与结果对拍**：补齐合成夹具所需模板结构后运行三轮。Python 中位数 4,320 ms，Rust 59 ms；逐项工作表、维度、单元格值及主要结构检查通过，列宽最大差约 0.54。输入文件未变；该规模仅 9 行输出且 Python 使用 Excel COM，不可泛化为语言级加速比。
- **其他已迁移功能性能复测**：S3-F1 10k 行 0.74x；S3-F2 5k/1,667 行 7.90x；S3-F3 1k 行 0.95x；S3-F4 10k 行 6.46x（后续专项逐格对拍为 0 差异）；S3-F5 30×40×4 的 3.79x 后确认旧输入没有任何业务命中，已作废并由有效 A1 表单负载替代；S2-F1 0.77x；S5 小负载 0.34x。S2 旧基准只核验输出表结构，S3-F3 核验原地备份/导入，S3-F2/F4/F5 的具体结果边界见性能历史报告。
- **构建与测试状态**：Release 构建成功；定向 `s1_native_com_session` 测试通过 1/1。没有运行全量 Cargo 测试、打包或 Linux 测试。完整性能口径、逐轮数据和限制记录于 `audit-core/probe/reports/perf_comparison_20261030.md` 与同名 JSON。
- **S5**：本轮只读取现有配置执行性能测量，没有修改或恢复 S5 业务开发。

## S1 逐笔统计审核

- **S1-F1 模板校验与审核：生产链路已打通，隔离真实样本的 Release GUI 点击验收通过，待用户按机构业务人工验收。** GUI worker 串起逐文件模板匹配、模板体检、结构不符跳过及原因记录、独立审核副本公式复制、外部数据引用导入、真实重算、公式/校验区域异常与重算副本上实际触发的条件格式异常提取、两类 Issue 合流、历史归类和 13 列结果写入。Windows Excel COM 16.0 合成 E2E 通过：独立预期为 `检查表!D1=#DIV/0!`、`D2=#REF!` 各 1 条公式异常及 `B2` 1 条实际触发的条件格式异常；输出可重开且源文件、模板、预重算副本哈希不变。2026-10-29 在 COM `xlPending` 修复后，用旧 Python S1-F1 的真实源文件、模板、外部数据、映射配置及同机构 Python 审核副本复跑 3 个真实探针，3/3 通过：约 210 个 `参照表!D2:D211` VLOOKUP 公式及依赖保留；模板 1,072 个公式文本一致，Excel COM 重算后 1,072 个存储结果与旧 Python 审核副本逐格一致（695 个非空，0 差异）；外部表独立重算基线一致；分段导入和完整副本均可由 Excel COM 打开；源文件、模板、外部表、配置哈希不变。Umya 将样本中 211 个显式空字符串单元格序列化为空白；其他非空值未变，且本样本 1,072 个公式结果仍与 Python 基线一致。2026-10-29 最新 Release GUI 使用隔离的真实源、模板、映射、外部数据实际点击执行并生成可重开的 13 列结果；该样本没有触发条件格式异常，触发型异常由合成 E2E 覆盖。该证据不替代用户对机构审核结论的业务验收。Windows 与 Linux x86_64 均已生成生产能力证据（Linux：`audit-core/capability_evidence/umya_linux_x86_64_20261028.json`），未登记平台仍失败关闭。校验区域扫描当前 GUI 使用重算副本、CLI 使用源文件缓存；统一口径待用户决定（AL-03）。
- **S1-F1 复制后运行时校验：已补齐并通过真实样本复验。** Excel-Core 生成的临时候选在缓存清理/真实重算前，按 S1 编辑计划重建预期公式与非公式输入值，比较所有工作表名称和顺序、公式单元格位置/文本及复制值；仅容忍已知零长度字符串规范化为空白，任一实质差异失败关闭。首次真实复验发现 Umya 对映射区域中缺失的源单元格未清除目标旧公式/值；已修正 Umya ImportRange 的空白覆盖语义，并由新的回归和候选校验验证。合成校验 15/15 通过，覆盖公式丢失/变化、引用表值变化、额外工作表和 210 条 VLOOKUP；随后真实旧 Python 工作簿完整探针 3/3 通过。真实 GUI 业务验收继续待办。
- **S1-F2 审核结果说明汇总：已实现，待人工验收。** CLI 与 GUI 入口已有模板/结构预检、任意行汇总和历史富化。历史记录有合成和部分真实样本证据；还需用户核对不同行数、异常跳过、历史关联及输出和原文件保护。普通列宽等外观差异不单独作为阻塞。
- **S1-F1 外部表公式复制：真实样本复制与计算通过，待 GUI 业务验收。** 旧 Python 的 Direct OOXML 优先、遇公式回退 openpyxl 属于 S1-F1。Rust 按本轮用户确认改为 Umya 主引擎、OOXML Patch 只承担已实现且语义相同的定点操作、Excel COM/LibreOffice 真重算。最新 Windows 真实样本探针覆盖纯数据源、普通公式和跨 Sheet 引用、约 210 个 VLOOKUP、同时复制「参照表」/「机构参照表」、模板公式及输入保护，并实际经过新增候选验证器。此前匹配样本已验证 1,072 个可读公式结果与同机构 Python 审核副本一致。2026-10-29 当前真实样本重跑中，纯数据往返与 Excel COM 分段/整本打开通过；公式复制/重算探针在未提供 Python 对照时通过，独立外部表重算结果、公式文本及缓存可读性检查均通过。使用同名 Python 对照时仅读到 973 个公式缓存结果，未达到探针要求的 1,000 个覆盖数，因此该次不形成 Python 结果对拍结论；此前已通过的 Python 基线对拍仍是独立历史证据。另一个 7 月目录中的同名外部文件与本次 `config-real` 映射下的预期非公式值不符，候选验证器在重算前明确拒绝该输入；未放宽校验。Umya 对缺失源单元格的导入现在清除目标旧值/公式。OOXML Patch 没有跨工作簿公式工作表复制能力，故 S1-F1 不登记该回退；Umya 对该操作失败时明确失败关闭。真实 GUI 整批审核输出仍待人工验收。
- **S1-F1 当前 Release 真实样本复验（2026-10-29）：3/3 通过。** 使用只读旧 Python 项目中的单位贷款报送样本及匹配的 `config-real` 模板、映射、外部数据，当前 `dist/lo-convert.exe` 经 Excel COM 完成重算。探针验证纯数据 Umya 往返、模板与外部表导入、至少 1,000 个模板公式位置/文本、外部 `参照表!D2:D211` 至少 200 个 VLOOKUP 公式以及其独立重算值、`机构参照表` 公式/值、完整副本公式缓存和原始四个输入文件哈希。未提供 Python 输出基线，因此此轮不宣称全部模板结果与 Python 逐格一致；另有此前 1,072 项完整基线对拍记录。首次选用的另一旧工作簿因已含同名“集中系统数据”页而被既有安全检查拒绝，且该候选独立 Umya 往返会改写其公式；未绕过校验，改用不含冲突页的报送样本后通过。日志：`F:\RustBuild\audit\s1f1-real-configreal-source-candidate-20261029.log`；被拒候选记录：`F:\RustBuild\audit\s1f1-real-current-goal-20261029.log`。GUI 整批输出仍待人工验收。
- **S1-F1 引擎路由与跨 Sheet 覆盖复验（2026-10-29）：3/3 真实探针通过。** 当前真实工作簿测试现在直接断言完整模板/外部公式导入批次由 `UmyaSpreadsheet` 唯一选中并成功执行；`参照表!D2:D211` 的真实 VLOOKUP 必须存在对「机构参照表」的跨 Sheet 引用，随后验证引用表与依赖表共同导入、独立 Excel COM 重算结果、输出重开和原始输入哈希。当前同语义公式工作表 `ImportRange` 没有 OOXML Patch 后备能力，故不尝试伪回退。`excel-editor` 单元测试 48 项通过（包括写入前拒绝后从原始副本切换到已声明 OOXML 能力、双方均不支持时失败关闭及写入后错误不回退）；公式坐标/区域集成测试 5 项通过；S1 公式复制单测 15 项通过。此次未提供 Python 审核副本作逐公式结果基线；真实 GUI 整批输出仍待人工验收。日志：`F:\RustBuild\audit\s1f1-umya-route-cross-sheet-current-20261029.log`、`F:\RustBuild\audit\excel-editor-engine-policy-current-20261029.log`、`F:\RustBuild\audit\excel-editor-import-formula-current-20261029.log`、`F:\RustBuild\audit\audit-core-s1-formula-copy-current-20261029.log`。
- **S1-F1 当前 Release 业务 worker 真实样本复验（2026-10-29）：已完成，待 GUI 人工验收。** `build-release.bat` 使用 `F:\RustBuild\audit` 成功构建并更新 `dist` 主程序与 sidecar；随后通过 Release `acceptance-run` 调用 GUI 使用的同一 `start_dag_flow`/S1-F1 worker，在隔离副本上执行一个真实单位贷款样本。结构匹配 406/406 个标签，1/1 文件通过、跳过 0、失败 0、部分 0；生产 Umya 导入 4 张外部表，处理 1,282 个公式，Excel COM 完成重算。运行器报告 `completed`，3 条公式异常、0 条条件格式异常、总计 3 条；源、模板、映射、外部数据和清单均哈希未变。正式 `汇总核查表校验.xlsx` 已由 Excel 只读打开，`本期审核结果` 为 13 列、4 行。此真实样本没有触发条件格式异常，触发型条件格式仍由既有合成 E2E 覆盖；本次未传 Python 结果基线，也没有执行真实 GUI 点击验收。报告目录：`F:\RustBuild\audit\s1f1-release-acceptance-20261029-065200\output`。
- **S1-F1 Release GUI 真实样本点击验收（2026-10-29）：已通过隔离样本。** 最新 Release `serve` 页面实际选择源目录、模板目录、`mapping.xlsx`、`external.xlsx` 和固定输出目录，刷新后自动匹配 1 个模板并点击 F1。真实源文件结构匹配 406/406；Umya 导入 4 张外部表并准备 1,282 个公式，Excel COM 完成重算；任务 `COMPLETED`，提取 3 条公式异常、0 条条件格式异常，输出 `本期审核结果` 4 行×13 列且由 Excel 只读重开成功。源文件、模板、映射、外部数据 SHA-256 与此前隔离验收基线完全一致。结果位于 `F:\RustBuild\audit\s1-final-release-20261029-070522\gui-output-ui-retest\S1-F1_审核结果_20261029_073237_123\汇总核查表校验.xlsx`。该样本未触发条件格式异常，本次也没有新的 Python 逐格结果基线；触发型条件格式由既有合成 E2E 覆盖。GUI 还修正了 F1 “复制能力未验证”的过时标签、输出跟随/固定状态 API 及确认框的实际输出路径，并防止打开其它输入选择器时误固定默认输出。新增输出模式回归 1 项通过。全工作区回归 80 个目标、484 passed / 0 failed / 32 ignored；日志 `F:\RustBuild\audit\workspace-final-goal-ui-output-20261029.log`。最新 Windows Release 构建成功；日志 `F:\RustBuild\audit\release-final-goal-ui-output-20261029-r2.log`。以上是隔离真实样本的自动化 GUI 点击验收，不等同于用户原始文件和审核规则的业务签收。
- **S1-F3 人工重算恢复：已实现，worker 测试及当前 Release GUI 恢复点击通过；待用户手工重算与业务验收。** F1 无法自动重算时持久化审核副本路径、模板/配置摘要及精确和兼容公式指纹；隐藏的 S1-F3 入口仅在存在待恢复任务时显示。用户在 Excel/WPS/LibreOffice 中强制重算、保存并关闭副本后，F3 检查公式位置/文本、全部缓存节点及每个公式结果是否可由 Calamine 读取，再复用 F1 公式/条件格式提取及结果写入，不重新读源文件或复制公式。公式保护仅容忍 LibreOffice 将 `TRUE`/`FALSE` 写成等价的 `TRUE()`/`FALSE()`；公式位置、引用、区域及其余文本必须一致。恢复状态跨程序重启保留；相同审核结论复用同一输出，输出已发布但状态尚未提交时不会再生成重复结果；审核结论变化会产生新结果文件且不覆盖旧版。2026-10-29 新增 worker 端到端合成回归：重建状态存储后从待处理任务执行 F3、提取 1 条公式异常并发布可重开的 13 列结果，输入哈希不变、任务完成后待处理数归零；另验证缓存缺失时保持 `WAITING_MANUAL_RECALC` 且不发布结果。另新增 Windows 真实样本探针：Umya 生成 1,282 个公式的 S1-F1 独立副本，Excel COM 在分离输出文件中重算后将已保存副本交由新建状态存储的 F3 worker 恢复，成功提取 3 条公式异常，生成可重开的 13 列结果并完成任务，源、模板、外部表和配置哈希不变。本轮用当前 Release 服务做了隔离进程级状态验证：空状态待处理数为 0；单个待恢复任务为 1；结束服务并重启后仍为 1，恢复读取错误为空。该验证覆盖持久任务状态/API，不执行 F3 恢复按钮，也不替代用户在 Excel/WPS/LibreOffice 手工重算、保存、关闭及 GUI 恢复验收。缓存节点只能证明缓存存在；typed 可读性证明值可被消费，仍不单独证明数学正确。
- **S1-F3 Release 跨进程恢复（2026-10-29）：真实工作簿流程通过。** 新的隔离 Release `acceptance-run` 在刻意使用不兼容 sidecar 后，将 1 个真实工作簿保留为 `WAITING_MANUAL_RECALC`，保存 1,282 个公式的审核副本及恢复清单；首次运行进程退出后，由 `lo-convert.exe recalculate --engine com` 使用 Excel COM 重算并保存副本。随后启动全新 Release `serve` 进程，GUI 使用的 `/api/list_s1_f3_tasks` 在进程重启后读到该待处理任务，`/api/start_s1_f3_workflow` 调用正式恢复 worker，完成 3 条公式异常提取、0 条条件格式异常、生成 Excel 可重开的 13 列结果并将队列归零。此项验证了 Release 进程之间的持久化与 GUI API 恢复路径；人工重算由 sidecar 模拟，未模拟用户按键或真实 GUI 点击。材料：`F:\RustBuild\audit\s1f3-process-restart-20261029-065923\`。
- **S1-F3 输入完整性校验补强：** 待恢复任务额外持久化非公式单元格值、条件格式语义、合并区域和命名区域摘要；F3 允许公式缓存更新，并按规则内容/相对优先顺序规范化 LibreOffice 的条件格式样式 ID、默认属性与命名区域序列化顺序。发现审核输入/规则变化时拒绝恢复。旧清单保留原公式校验兼容行为。合成回归覆盖缓存清除不改变摘要、改动普通输入导致摘要变化、恢复校验拒绝变化及 LibreOffice 重存语义兼容；真实 Windows Excel COM 工作簿恢复 worker 探针已通过，真实 LibreOffice 样本及 GUI 手工打开/重算/保存/重启恢复仍待人工验收。
- **S1-F3 恢复副本异常可见性：已实现。** 重启后发现待恢复清单引用的副本丢失，会持久化为 `FAILED` 并在 GUI 明确提示重新运行 S1-F1；存在但损坏的副本会保留原字节，恢复记录为 `RESUME_FAILED`，不发布审核结果并继续出现在可修复重试列表。GUI 的恢复清单读取错误也有独立提示。测试覆盖丢失/损坏副本、状态列表与页面渲染。
- **S1-F4/F5：已明确暂缓。** 不自动开发组合工作表或联合模板。

## Excel 引擎策略及验证状态

- **读取：Calamine 主读值。** OOXML 只补命名区域、条件格式、批注等需要的元数据；未实现通用 OOXML/openpyxl 读值回退。Calamine 报告特定未知 Excel 错误 token 时，才复制原 XLSX/XLSM 到临时包、只将该错误单元格缓存 token 作为文本规范化，再由 Calamine 重读；临时副本自动清理，原文件不改。当前源码 `excel-reader` 19 个单元测试和 2 个集成测试通过，覆盖纯文本/公式错误值、公式文本/普通数据保留、输入哈希、缓存检查、不安全 XML 拒绝及文件缺失/权限/格式/工作表读取/损坏分类。其他损坏、权限、格式、工作表读取错误不走此兼容回退；分类错误信息保留原始错误链。
- **修改：Umya 主引擎 + 能力限定的 OOXML Patch。** Excel-Core 按操作和能力证据选路，只有 Umya 在写入前明确报告不支持，且 OOXML Patch 对同一操作有通过证据时，才允许从干净原始副本回退；部分写入、保存失败、输出校验失败或未知错误均停止。合成路由测试覆盖“Umya 预拒绝→OOXML 干净副本回退”和“两者均不支持→失败关闭”；2026-10-29 对当前源码定向复跑两项均通过。该测试验证路由协议，不把未实现的 OOXML 单元格编辑登记为生产能力。S1-F1 的真实公式 ImportRange 已由 Umya 满足，OOXML Patch 不实现冗余公式复制，故此业务操作不配置伪回退；S1 业务层现于候选落盘后核验公式/引用、工作表顺序和复制的非公式值，再允许进入外部重算。
- **公式计算：Windows PowerShell + Excel COM；Linux LibreOffice。** Windows `auto` 顺序为 Excel COM、再 LibreOffice；Linux `auto` 仅 LibreOffice；人工重算是引擎不可用/失败时的 S1-F3 恢复路径。计算前在独立副本清缓存并请求全量重算；COM 专用副本将 `fullCalcOnLoad` 和 `forceFullCalc` 设为 false，避免 Excel 在两标记同时开启时于 `CalculateFullRebuild` 后仍持续报告 `xlPending`；COM 仍执行全量计算和依赖重建，并要求 `xlDone` 才保存。LibreOffice 路由继续使用两个全量重算标记。保存后核对工作表顺序/名称、公式单元格位置、公式文本（只允许 `TRUE`/`FALSE` 与 `TRUE()`/`FALSE()` 的等价序列化）、所有缓存节点存在，且 Calamine 能为每个公式读取带类型的结果（包括错误值和空字符串）；这证明外部引擎写出可消费的结果缓存，不单独证明数学结果正确，正确性仍由业务预期值/基线和审核规则验证。Windows Release sidecar 的 S1-F1 合成 E2E 和修复后真实样本 E2E 均通过；真实样本中 1,072 个模板公式结果与 Python 审核副本无差异。Linux LibreOffice 实际重算集成覆盖 VLOOKUP、跨表公式、错误值、空字符串和工作表结构。目标机与真实 GUI 验收仍待完成。
- **重算超时进程清理：已实现，Windows/Linux 回归通过。** Windows 超时先对 sidecar 启动的进程执行 `taskkill /T /F`，再直接 kill/wait 父进程；Linux 为引擎建立独立进程组，超时对该组发送 `SIGKILL`，再回收父进程。若等待状态读取出错也走相同清理。新增测试启动真实子进程并断言超时后子进程已退出。2026-10-29 Windows sidecar 全量结果：14 单测通过；集成入口因 Windows 无 `soffice` 按默认策略跳过。Ubuntu WSL：14 单测及实际 LibreOffice CLI 重算集成均通过，后者由 `LO_CONVERT_REQUIRE_LO_TEST=1` 强制启用。Windows Release 已重建；新 `dist/lo-convert.exe` 的 S1-F1 Excel COM 合成端到端回归 1/1 通过。测试仅覆盖进程树清理与既有小型 LibreOffice 样本；目标 Linux 发行版的进程行为仍待设备验收。
- **条件格式及批次隔离：** OOXML 读取条件格式规则，由 Rust 支持的子集求值；不支持/求值失败需带规则、范围和单元格诊断，标记部分结果。2026-10-29 加固了工作簿内部隔离：单个条件格式元数据节点格式错误只跳过该节点并保留同表其他有效规则；单元格值无法安全转换时跳过该表的条件格式判断并记录坐标诊断；单个触发单元格无法转为异常时跳过该异常并继续处理其他触发项。整个条件格式提取阶段意外失败时，该工作簿已提取出的公式异常仍会写入部分结果，文件标为 `Partial` 并记录条件格式失败原因；公式异常提取或重算校验本身失败仍按文件失败处理。批次级已验证独立文件失败不会停止其他文件。该项早期全工作区回归为 80 个测试目标、478 passed / 0 failed / 31 ignored；最新源码总回归见本文件末尾。Release 构建及主程序版本/sidecar 帮助冒烟通过；使用新 Release sidecar 的 Excel COM 合成 S1-F1 端到端为 1 passed，日志 `F:\RustBuild\audit\s1f1-com-e2e-release-partial-condition-20261029.log`。未知条件格式规则仍只形成诊断，不能声称已被完整求值。真实业务文件及 GUI 批次仍待人工验收。

## Excel 引擎与 S1-F3 Goal 完成度核对（2026-10-29）

- **Calamine 读取错误分类与定点兼容：已实现并有测试。** 普通读值仍以 Calamine 为准；只有特定未知 Excel 错误 token 会在原文件的临时 OOXML 副本中规范化后再交 Calamine 读取。损坏、权限、格式、工作表读取等错误不会被扩大为通用重试。`excel-reader` 有 19 个单测和 2 个集成测试覆盖分类、回退边界及原文件保护。
- **工作簿修改路由：已实现；S1-F1 的主路径已有真实工作簿证据。** 生产路由按平台能力证据注册 Umya；ExcelModifier 只在后端写入前明确拒绝且同语义回退能力已验证时才从原始副本切换引擎。S1-F1 真实公式 `ImportRange` 由 Umya 独立完成；约 210 个 VLOOKUP 与跨 Sheet 依赖、公式/非公式内容及输出重开均已检查。OOXML Patch 没有相同语义的公式工作表复制能力，因此该操作没有伪造的 Patch 回退，Umya 失败时失败关闭。S1-F1 生产路由不调用 Python/openpyxl；仓库保留的隔离兼容 worker 不是 Audit 业务逻辑，也不在当前 Windows `dist` 中。
- **公式重算与验证：实现和合成/真实 Windows 样本已通过；Linux 真实业务样本未验。** Windows 通过 PowerShell + Excel COM，Linux 通过 LibreOffice；保存前要求计算完成，保存后核对公式结构、缓存节点和 Calamine typed 结果。Windows COM 已覆盖真实 S1-F1 公式结果及既有 Python 审核副本对拍；2026-10-29 本轮再次在 Ubuntu WSL 实际运行 LibreOffice CLI 集成（15 单测 + 1 集成，进程退出码 0），覆盖 VLOOKUP、跨表引用、公式错误和空字符串结果。Linux 真实机构工作簿与麒麟 V10/UOS V20 设备仍未验证。
- **超时清理、条件格式与批次隔离：实现和自动化回归已通过。** Windows 清理 COM/sidecar 进程树，Linux 清理 LibreOffice 进程组；条件格式不支持项带诊断，局部失败保留可靠公式结果并允许独立工作簿继续。具体边界和失败状态见本节前文及迁移清单。
- **S1-F3 人工重算恢复：代码、worker、跨进程状态和当前 Release GUI 模拟流程已通过；用户亲自重算仍待验收。** 待处理任务持久化，入口按任务状态显隐；恢复时验证公式身份、输入语义和可读缓存，然后复用 F1 的异常提取/结果发布；跨重启、幂等、丢失/损坏副本及混合状态有回归覆盖。此前 Release GUI 恢复操作使用 Excel COM 计算的隔离副本，不能替代用户在 Excel/WPS/LibreOffice 中手工强制重算、保存、关闭后再点击恢复。
- **Release 与 Goal 状态：Windows Release 已构建并当前冒烟通过；总 Goal 仍未完成。** 当前 `dist/open-data-audit.exe --version` 为 `0.1.0`，`dist/lo-convert.exe --help` 含 COM/LibreOffice 重算入口。最新主程序 EXE 时间为 2026-10-29 08:32（晚于嵌入 GUI 资源修改）；sidecar 源码时间为 08:06、Release EXE 时间为 08:12，构建日志为 `F:\RustBuild\audit\release-excel-engine-s1f3-completion-20261029.log`。最终全工作区结果见本文件末尾。尚缺用户提供的真实机构文件/审核结论签收、用户亲自执行的 S1-F3 重算恢复，以及目标 Linux 设备验收；这些缺口不能由合成测试或隔离 Release 点击替代。

## S2 报表采集审核

- **S2-F1：已实现，待人工验收。** 配置加载与校验、报表导入、指标/机构匹配、单位转换、本期与上期比较、外部数据核对、配置规则和结果导出已串入 Rust service、CLI 和 Web API。配置表达式可处理本期/上期数据；“同期比较”应按实际配置规则核验，不能仅凭名称宣称专门的同比流程已实现。
- 历史有真实 XLS 两期及外部数据的 CLI 样本记录，也有 Web API 和合成正常/异常样例；这些不等于当前独立构建的完整人工验收。仍需核对真实机构/指标映射、三类规则及外部异常、输出和原文件保护。

## S3 大集中统计系统：按原 Python GUI 的五项编号

原 Python `config-real/` 的 3.3 报表清单有 43 条非空启用记录；Rust 仓库 `audit-core/config/` 中的 3.3 是空占位，不能替代真实配置验收。3.0–3.3 的工作表和关键字段见 [S3 配置结构核对](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-配置工作簿的实际工作表与关键字段只读核对)。

原 Python GUI 的正式范围为五项：F1 两期数据→比较文件、F2 比较文件→说明文件、F3 机构说明→说明文件、F4 本期数值核对、F5 比较文件→金融表单。此前将 F4/F5 分别压成 Rust 目标 F2/F3、并把 F2/F3 当作额外旧流程的描述有误，现按 Python 实际编号统一。数据流为 F1→F2→F3，F1→F5；F4 独立运行。F1/F4/F5 依赖的 3.0–3.3 配置实际表头及入口、输出结构均以 [V2 需求的 S3 对照表](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md#s3-f1-至-f5-原-python-与-rust-业务对照) 为准。

- **S3-F1 两期数据→比较文件：已实现，待人工验收。** Rust `s3_execution_compare`、CLI `s3-execution-compare` 与 Web API `start_central_comparison` 已接通；读取本期/上期数据、3.0 通用配置和 3.1 执行比较配置，执行配对、单位换算、环比警戒、规则动作、五段/八段表达式和软性规则颜色处理，输出 21 个业务列及独立“输出说明”列。没有单独外部数据文件输入；GUI 每期选择一个文件。Linux Debian 10 / glibc 2.28 容器以合成异常样本实跑 CLI：1 行输出、1 条动作规则触发；Windows GUI 和真实机构样本仍待人工验收。
- **S3-F2 比较文件→说明文件：已实现，待人工验收。** Rust `s3_explanation_export`、CLI `s3-export-explanations` 与 Web API `start_central_explanation_export` 已接通。以比较文件中的“输出说明”人工标记筛选行；3.0 的“机构地区参照”用于修正展示名称和按旧规则拼接可区分机构代码，生成八字段“说明文件”及“生成信息”表。现有文件级对拍证据仍待用户按真实机构业务验收。
- **S3-F3 机构说明→说明文件：生产写回链路已打通，待人工验收。** 匹配反馈、备份、暂存写入、失败回滚和目标原地更新已接通 GUI/Web API 与 `s3-import-explanation-feedback` CLI；写回前制作不覆盖的时间戳备份，候选写入失败时恢复原件并保留备份。Windows 和 Linux 能力 harness 均验证目标原地更新、备份与反馈输入哈希保护。已按用户确认对齐 Python 输出契约；runner 若接收用户目标，会先复制到验收输出目录再执行。
- **S3-F4 本期数值核对：已实现，待人工验收。** Rust `s3_current_value_check`、CLI `s3-check-current-values` 与 Web API `start_central_cross_period` 已接通，使用当前期数据、3.0 和 3.2 配置，生成规则核对结果。GUI 实际只提供本期文件；服务/CLI 另有可选上期数据参数，但这条能力没有 GUI 选择入口。CSV/XLSX/XLS 已有局部核对证据，真实边界和 GUI 操作尚待验收；时长单元格目前显式拒绝。
- **S3-F5 比较文件→金融表单：生产编辑链路已打通，CLI 已接通，待人工验收。** `s3_indicator_split` 把规划结果转换为 Excel-Core 编辑计划，按机构写入独立暂存文件，全部成功后发布 `指标比较拆分_时间戳` 目录；GUI/Web API 和 `s3-generate-financial-forms` CLI 已接通。Windows/Linux harness 验证编辑批次和 16 项保留契约；CLI 成功/失败路径及 Release 隔离冒烟已有证据。3.3“分析文件数据单位”按用户确认属于历史遗留字段，继续沿用 3.0 默认目标单位。

Rust 现有 F1、F2、F4 逻辑和入口继续保留；F3、F5 的部分实现也保留并补齐各自剩余链路，不因编号校正重写。完整逐项业务和输出对照见 [V2 需求](audit-core/AUDIT_RUST_V2_BUSINESS_REQUIREMENTS.md)，历史实现证据见 [迁移清单](audit-core/MIGRATION_INVENTORY.md)。

编号映射补充：Rust 当前 `s3_explanation_feedback` 是 Python S3-F3；`s3_indicator_split` 及表单编辑操作测试实际对应 Python S3-F5。遗留的 `s3_f3_edit_operations.rs` 文件名和 `AUDIT_S3_F3_*` 环境变量仍带旧编号；该测试函数/说明已按表单编辑操作归入 F5，并在 2026-10-28 补充了测试源码，但未运行，也不构成 F5 生产写入能力或验收证据，更不能作为 F3 业务实现证据。

S3 的辅助配置检查入口已补齐：GUI 显示 3.0–3.3 四个检查按钮，Web bridge 现注册全部四个只读检查 API（`check_central_common_config`、`check_central_expression_config`、`check_central_cross_formulas`、`check_central_forms_config`），并有成功/缺参单元测试。这些按钮仍只是辅助诊断，不作为额外 S3-F 编号或 F1–F5 已完成的证据。

## S5 配置化审核

- **S5-F1：已明确暂缓。** 用户于 2026-10-30 明确要求暂停 S5 配置审核开发。本条既有实现及历史验证记录保留；本轮没有修改 S5 业务代码。除非用户明确恢复，否则不继续 S5 功能开发或性能调整。此前真实配置/业务文件 GUI 人工验收仍未完成，不能视为验收通过。
- GUI 当前仍接受同模板/业务的本期批次、可选上期目录及一个物理外部工作簿；CLI 多来源绑定能力继续保留。CLI 指定模板时对主本期文件执行缓存错误扫描；GUI 按每份本期文件执行。2026-10-29 完整工作区回归通过；当前 Windows Release CLI/LibreOffice 侧车也已重建。GUI 实际操作及真实业务结果仍待人工验收。
- S5 保持 Rust 配置规则直接计算，不重新引入模板公式注入或 Office 规则计算。
- S5 旧测试图谱中的隐藏维护项 F02/F03/F04 尚未确认业务范围，不自动恢复开发。

## 其他范围与交付

- **S4：已明确暂缓。** 保留当前 Rust + HTML GUI；TUI 是可选未来方向，不在本轮实施。
- 项目内置 Excel-Core 源码和独立 Cargo workspace；外部第三方 crates 仍需要相应 Cargo 缓存或正常依赖获取。旧 Python、Main 和 Codex 工作树不是构建或运行路径。可选外部计算引擎/隔离 worker 只用于其明确支持的功能，不能推断所有路径都可在无外部引擎时执行。
- Windows：按 Excel 引擎 Goal 的最终复验再次运行 `build-release.bat`，成功生成主程序与更新后的重算 sidecar 并复制到 `dist`，日志 `F:\RustBuild\audit\release-excel-engine-s1f3-completion-20261029.log`；`open-data-audit.exe --version` 为 0.1.0。新 Release sidecar 的 Excel COM 合成 E2E 1/1 通过；F1 真实 Umya 探针 3/3、F3 Release 真实恢复探针 1/1 通过。ZIP `dist/open-data-audit-windows-x64-20261028-122102.zip` 仍是 2026-10-28 的包，未在本轮重打包。
- Linux：`linux-x86_64` 生产能力证据已生成，未知平台继续失败关闭。最终 musl deb 为 `dist/linux-delivery/open-data-audit_0.1.0_amd64.deb`（13,498,066 字节），在 Debian GLIBC 2.28 / dpkg-deb 1.19.8 等价容器通过安装、静态链接、CLI、Web API、S3-F1/F2/F5 合成 CLI 冒烟；无 `soffice` 时重算侧车明确报错。麒麟 V10/UOS V20 实机尚未验收。
- 验收运行器：在隔离输出中运行 S3-F1→F2 和 S5 合成样本，步骤均完成、所有源文件与配置哈希不变，报告 JSON/Markdown 生成于 `F:\RustBuild\audit\acceptance-runner-synthetic-0b813b7c867b4f979e155fccad36f151`。
- 完整回归：当前源码 `CARGO_TARGET_DIR=F:/RustBuild/audit cargo test --offline --locked --workspace -j 1 -- --quiet` 为 **80 个测试目标，489 passed / 0 failed / 32 ignored**，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-final-20261029-r5.log`。忽略项包括真实样本/人工操作及外部运行器，未计为通过。WSL x86_64 LibreOffice 重新计算套件以 `LO_CONVERT_REQUIRE_LO_TEST=1` 为 17 单测和 1 实际 CLI 重算集成通过；目标发行版设备验收仍待完成。
- 公开配置是结构占位，真实配置/样本保持本地忽略。audit-core/manual_acceptance/20261027/ 中大量 Excel 手测样本与输出当前只存在本机、被 Git 忽略；文档和部分数据被跟踪，不等于全套样本随仓库提交。

## 2026-10-29 S1-F3 Release GUI 验收补充

当前源码的 Release 重新构建成功。对隔离的 Excel COM 已重算审核副本执行了当前 Release GUI 完整恢复点击：S1 页显示 1 个待处理任务，打开恢复面板并确认后生成核查表；状态更新为 `COMPLETED`，包含 3 条公式异常、0 条条件格式异常，审核副本 SHA-256 与执行前复制源一致。结果目录：`F:\RustBuild\audit\s1f3-current-release-gui-e2e-20261029\output\S1-F1_审核结果_20261029_073237_123`。

该 Release 测试同时发现并修复了“待处理数为 0 时按钮仍可见”：`.btn` 的 `display` 样式覆盖了 HTML `hidden` 属性。加入 S1-F3 专属隐藏规则及页面断言后，定向页面测试 1 项通过；重新构建的 Release 在待处理数为 1 时显示按钮，恢复后待处理数归零且按钮从无障碍树与实际画面中消失。该 CSS 修复后的最新源码已重新执行全工作区回归：80 个测试目标、485 passed / 0 failed / 32 ignored，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-current-20261029-r4.log`。测试夹具目录隔离加固及首次两个不完整 deflate 错误的排查见迁移清单最新增量。

GUI 点击测试使用的是此前由 Excel COM 重算并保存的隔离副本；没有模拟用户亲自在 Excel/WPS/LibreOffice 按 `Ctrl+Alt+F9`。用户手工重算并对机构业务结果验收、麒麟 V10/UOS V20 设备验收仍待完成。

## 对齐空间与引擎高级设置（2026-10-29）

- **对齐决策清单**：`audit-core/docs/ALIGNMENT_DECISION_LIST_20261029.md`——矩阵 104 个失败场景分三类（96 建议直接采纳 / 6 需用户业务决策 / 2 项 S5 移出）+ 9 条对齐空间（AL-01…AL-09）。待用户决策项见清单第二部分与 AL-04/05/09。
- **引擎高级设置**：设置中心“高级”页新增“Rust 引擎（高级）”卡片。S1-F01 公式重算引擎四通道：auto（Windows=Excel COM→WPS 表格探测；Linux=LibreOffice 无头）/ 强制 Microsoft Excel（COM）/ 强制 WPS 表格（COM，ProgID 探测 ket.Application→KET.Application）/ 强制 LibreOffice 无头（经 lo-convert sidecar）；Excel-Core 修改引擎两通道：Umya（默认，已验证）/ openpyxl（高级兼容路径，证据驱动注册）。选择持久化到状态目录 settings.json 并在日志回显；强制引擎不可用即报错，不自动切换。
- **openpyxl 高级路径**：负面证据（capability_evidence/openpyxl_windows_x86_64_20261029.json，available=false）——openpyxl 重写无法满足三个业务批次的生产保留契约（OOXML 关系、列尺寸、包部件、DeleteColumn），全部批次路由失败关闭；GUI 置灰。未来后端补齐契约后重跑 AUDIT_OPENPYXL_CAPABILITY_REPORT 证据即可解锁。
- **对拍**：S3-F02 合成 Python/Rust 对拍通过（probe/reports/s3_f02_synthetic_parity_20261029.json，4/4 行一致、differences=[]、9 项检查全过）。S1-F02 对拍器夹具落后于当前表结构预检（0/3 被跳过），为对拍器维护项。S3-F1 数值口径评估：旧 Python 基准即 IEEE-754，不建议 Decimal 化（AL-04，待确认）。
- **编号澄清**：Rust 内部“S1-F3”是 S1-F01 人工重算恢复 worker（s1_recovery.rs），与 Python 菜单暂缓的 S1-F3（组合工作表）撞名；是否重命名待用户确认（AL-05）。
- S1-F01 校验区域扫描双口径（CLI 源缓存 / GUI 重算副本）已由 tests/s1_scan_scope_contract.rs 固定，统一口径待用户决定（AL-03）。此前误标为 S1-F02。

## 下一步

本轮已按用户授权完成 COM `xlPending` 定点修复、公式结果可读性/规范化校验、S1-F1 Umya 空白覆盖修复、复制后候选公式/输入值校验、真实旧 Python 工作簿复验、Release GUI 输出目录状态修复、全工作区回归及最新 Windows Release 构建；隔离真实样本的 GUI 点击、Excel COM 重算和 13 列结果重开已通过。另在当前 Release 上完成 S1-F3 GUI 任务选择、确认恢复、结果写出及完成态刷新；修复完成任务后按钮仍可见的问题。这些结果不替代用户机构文件/规则的业务验收。下一步是用户选定文件的人工验收：

1. S1-F1 以用户选定的机构源文件、模板、外部数据和配置在 GUI 执行审核；核对模板匹配/跳过原因、13 列结果与实际公式/条件格式异常，并确认源文件不变。隔离真实样本 GUI 路径已经过验证。
2. S1-F3 的 Release GUI 恢复点击流程已在隔离的已重算副本上通过：任务显示→选择→确认→结果写出→待处理数归零；仍待用户实际使用 Excel/WPS/LibreOffice 对由 S1-F1 保留的副本强制重算、保存并关闭，再通过 GUI 恢复，并检查真实业务结论、重启持久化和重复恢复幂等。
3. S3-F1 至 F5 使用真实 3.0–3.3 配置与业务数据验收菜单操作、输入/输出结构及 F1→F2→F3、F1→F5 数据流；F3 核对“先备份、再原地更新”，F5 按已确认口径沿用 3.0 默认目标单位。S2、S5 使用真实配置和输入文件核对结果及原文件保护。
4. 在麒麟 V10/UOS V20 x86_64 目标机安装 deb 并复核 GUI、配置发现、S1 缺少 Office 引擎时的提示及业务入口。

## 2026-10-29 增量：真实 S1-F1 LibreOffice 兼容与 Release 复验

真实 Linux S1-F1 探针首次运行时，在外部工作簿基线重算阶段失败，尚未进入 Umya 模板复制。对隔离副本直接执行 LibreOffice 并比较 OOXML 后确认：`参照表!D2:D211` 共 210 条 VLOOKUP 公式文本完全未变，LibreOffice 仅添加 `aca="false"` 序列化属性；原本为固定错误值的 `参照表!C211` 则被 LibreOffice 写成缓存值相同的 `#REF!` 公式。没有放宽通用公式校验。

现新增能力受限的 OOXML Patch `restore_promoted_error_constants`：只接受源工作表中确为 `t="e"` 的错误常量，并且计算输出在同一坐标仍为相同的受支持错误 token、公式属性仅为可选的 `aca="false"`、缓存错误值完全相同；仅从新的候选副本移除该单元格的公式节点并保留错误值。sidecar 随后仍要求工作表、所有原公式位置/文本、结果缓存节点和 Calamine 类型化结果严格通过现有校验。新公式、被改写公式、引用变化、错误值不同或公式属性异常均失败关闭。该补丁不是跨工作簿公式复制引擎；S1-F1 ImportRange 仍由 Umya 单独完成，不配置未经验证的 OOXML 公式复制回退。

最终证据：`ooxml-patch` 定向单测 2/2 通过；Windows `lo-convert` 重算单测 17/17 通过；Ubuntu WSL 以 `LO_CONVERT_REQUIRE_LO_TEST=1` 执行 `cargo test -p lo-convert -- --quiet`，17 单测和 1 个真实 LibreOffice CLI 集成通过。集成样本新增固定错误常量，确认输出仍保留原错误值且不变成公式。使用 WSL LibreOffice 对隔离真实 S1-F1 工作簿执行完整探针 1/1 通过：覆盖纯数据、普通/跨表公式、约 210 个 VLOOKUP、共同复制「参照表」及「机构参照表」、模板公式复制、重算后值比较、输出重开和原输入保护。Excel-Core 路由回退、双引擎不支持失败关闭、失败后从原副本重试三项定向测试均通过。随后 `build-release.bat` 成功更新 Windows `dist`，`--version`/sidecar `--help` 冒烟通过；最新 Release sidecar 经 Excel COM 对同一真实 S1-F1 探针 1/1 通过。最新 Windows 全工作区回归为 80 个测试目标、489 passed / 0 failed / 32 ignored，日志 `F:\RustBuild\audit\workspace-excel-engine-s1f3-final-20261029-r5.log`。其他日志：`F:\RustBuild\audit\s1-real-linux-lo-probe-20261029-r3.log`、`F:\RustBuild\audit\lo-convert-wsl-final-20261029.log`、`F:\RustBuild\audit\s1-real-release-com-error-restore-20261029.log`、`F:\RustBuild\audit\release-s1f1-error-constant-normalization-20261029.log`。三个修改源文件的 `rustfmt --check` 通过。用户机构业务签收、S1-F3 用户手工重算恢复和 Kylin/UOS 目标设备验收仍待完成；总 Goal 保持 active。

## 2026-10-29 增量：当前源码真实样本 S1-F3 worker 复验

当前环境没有预设 `AUDIT_S1_REAL_*` 路径，因此全工作区测试中的真实 S1-F3 案例按要求保持 ignored。随后从本机既有隔离样本目录显式提供只读源文件、模板、外部数据、映射及当前 `dist/lo-convert.exe`，在当前 Release 测试二进制中手动启用该测试，1/1 通过：当前 Umya 生成 1,282 公式的审核副本，Excel COM 将计算结果保存至分离副本，再恢复为原审核副本；新建恢复 Store 后执行 S1-F3 worker，状态为 `COMPLETED`，提取 3 条公式异常、0 条条件格式异常、0 条诊断，生成含预期 13 列标题的结果，源文件/模板/外部数据/映射哈希未变化。日志 `F:\RustBuild\audit\s1-f3-real-worker-current-20261029.log`。此后再次运行当前 `dist` 冒烟：主程序 `--version` 返回 `0.1.0`、sidecar `--help` 显示 `recalculate --engine auto|lo|com`，二者退出码均为 0；主程序/sidecar 分别为 09:46:26/09:46:50 构建，晚于当前相关源码修改时间，无需重复 Release 构建。这证明当前真实工作簿 worker 与重启后恢复路径可工作，但 Excel COM 仍是对人工重算步骤的自动模拟，不等于用户本人在 Office 按 `Ctrl+Alt+F9`、保存关闭并从 GUI 恢复；真实机构业务结论签收和 Kylin/UOS 设备验收继续待办，Goal 保持 active。
