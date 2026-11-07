# Rust-Audit S1/S2/S3 性能历史

仅追加历史。不要覆盖旧报告或把不同计时边界的数字拼成一个连续趋势。倍率统一按 Python ÷ Rust 计算；只有确认为原始多轮数据时才使用中位数。

## 2026-10-29：首轮性能对照

来源：[报告](probe/reports/perf_comparison_20261029.md)、[逐轮 JSON](probe/reports/perf_comparison_20261029.json)、[驱动](tests/compare/perf_compare_20261029.py)。每项三轮，报告当时取最优轮；Rust 计入 CLI 启动，Python 不计解释器启动。负载为合成输入，未对所有结果做完整逐格对拍。

归档提交：报告 `87fcbca`；性能驱动 `af387ef`。原始测量未在 JSON 中固定记录当时的源码 SHA，因此这两个提交用于追溯证据文件和驱动，不冒称每次计时发生时的精确二进制提交。

| 功能 | 输入规模 | Python / Rust | 比率 | 状态 |
|---|---|---:|---:|---|
| S3-F1 | 10k CSV 行 | 1,114 / 1,483 ms | 0.75× | 优化前；后由 S3-F1 结果验证路径定点优化。 |
| S3-F2 | 5k 行 / 1,667 选中 | 1,920 / 222 ms | 8.64× | 旧一轮基线；后续同功能测量使用不同运行数据，不能直接作趋势曲线。 |
| S3-F4 | 10k CSV 行 | 6,040 / 957 ms | 6.31× | 旧一轮基线。 |
| S3-F5 | 30 机构 × 40 指标 × 4 表 | 2,021 / 887 ms | 2.28× | 旧一轮基线。 |
| S2-F1 | 两份真实夹具报表 | 12.8 / 38.2 ms | 0.33× | 旧运行与后续完整服务+输出基准的计时/夹具边界未证明相同。 |
| S1-F1 COM 子步骤 | 7,200 个 VLOOKUP | 4,658 / 6,300 ms | 0.74× | 不同重算边界；Rust 当时还计入缓存清理、校验及原子发布。之后改用原生 COM 批次并完成最终 A/B。 |

**S3-F5 历史基准限制：** 此表所列 2,021/887 ms 使用与 2026-10-30 相同的 `AF001`…`AF004` 合成表名。S3-F5 会根据表单工作表名中的 `A1/A2` 推断业务类，这些表名未匹配输入人民币业务类。当前 Release 复现为 `matched_indicators=0`、输出数值格为 0；故以上历史计时不代表有回填工作的金融表单渲染。

## 2026-10-30：全功能合成/仓库夹具复测

来源：[报告](probe/reports/perf_comparison_20261030.md)、[原始逐轮 JSON](probe/reports/perf_comparison_20261030.json)、[驱动](tests/compare/perf_compare_20261030.py)。JSON 每功能有三轮，可复算中位数；Python 计时不含解释器启动，Rust CLI 子进程计时包含约 21.8 ms 启动。S1-F2 有单独三轮报告。

归档提交：汇总报告 `87fcbca`；驱动 `af387ef`。Release 二进制 SHA 与运行时源码提交没有对每个功能逐项绑定记录，不能从报告反推精确构建提交。

| 功能 | 测试数据 | Python 中位数 | Rust 中位数 | 倍率 | 性能样本正确性检查 |
|---|---|---:|---:|---:|---|
| S1-F2 | 2 源工作簿、3 张表、9 输出行 | 4,320 ms | 59 ms | 73.22× | 逐格值与结构通过，列宽差最大约 0.54。 |
| S2-F1 | 2 个唯一夹具各复制到本期/上期；1 输出簿 | 35.6 ms | 41.3 ms | 0.86× | 输出可读、表名和尺寸一致；无完整值对拍。 |
| S3-F1（优化前 raw JSON） | 两份 10k 行 CSV | 1,138.6 ms | 1,373.6 ms | 0.83× | 计时驱动没有完整值对拍。 |
| S3-F2 | 5k 行、1,667 选中 | 2,270.5 ms | 255.7 ms | 8.88× | 该性能轮没有全量值对拍。 |
| S3-F3 | 1k 行反馈；含备份和原地写入 | 1,030.4 ms | 1,055.3 ms | 0.98× | 导入数、备份字节、反馈源保护通过；未覆盖冲突键。 |
| S3-F4 | 10k 行 | 6,034.0 ms | 941.0 ms | 6.41× | 该性能轮没有全量值对拍。 |
| S3-F5 | 30 × 40 × 4；各 30 输出簿 | 1,936.1 ms | 963.4 ms | 2.01× | 输出数量相同；该性能轮没有全量值对拍。 |

**S3-F5 记录更正：** 本表和 2026-10-29 S3-F5 数据都使用了失效表名并得到空表；数字保留作为历史档案，但不再视为有效性能基准。旧生成器只验证文件数量，没有核验业务匹配量，这是基准夹具缺陷，而非产品业务缺陷。

**数据修订说明：** 原 Markdown 主表记录 S2-F1 为 40/52 ms，S3-F1 为 1,064.1/1,432.5 ms、S3-F2 为 1,853.2/234.6 ms、S3-F3 为 934.0/980.9 ms、S3-F4 为 6,053.9/937.5 ms、S3-F5 为 3,290.8/868.5 ms。上述数值与同目录 JSON 的逐轮值及所写“中位数”不相符。基线改用机器 JSON 逐轮数组重算；历史报告原样保留，避免静默改写。

同周期的非 A/B 阶段观察：S2 Web 42 份 XLS 批次总计 2,249 ms，其中跨期 734 ms、保存+回读 905 ms（[报告](probe/reports/s2_web_stage_logging_real_windows_native_20261026.json)，明确注明不是基准）；S3-F2 真实工作簿导出 2,610 行，Rust 读文件 619 ms、读配置 41 ms、写出及校验 1,186 ms、总计 1,959 ms（[报告](probe/reports/s3_f02_real_comparison_parity_windows_native_20261026.json)，Python/Rust 输出对拍通过，但没有 Python 计时）。这些测量不能用来计算加速比，提供了后续定点剖析候选。该阶段数据与实现源码可追溯到归档提交 `87fcbca`；对应 GUI/服务和导出入口分别由 `8b28677`、`af387ef` 保存。

### S3-F2 规模剖析（2026-10-30）

旧真实样本只保存了脱敏汇总，原输入文件不可复用。为验证行数本身是否能复现旧写出耗时，新增 [聚焦基准驱动](tests/compare/perf_s3f2_scale_20261030.py)，以相同于 2026-10-30 S3-F2 合成基准的列和参照结构，生成 7,830 行输入、2,610 行选中数据，并在临时目录运行三轮交错 Python/Rust。使用 Release 二进制 SHA-256 `9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4`。

| 计时 | Python | Rust CLI 墙钟 | Rust 内部 |
|---|---:|---:|---:|
| 三轮中位数 | 2,919.43 ms | 359.88 ms | 329 ms |

Rust 内部阶段中位数：比较文件读 129 ms、配置读 2 ms、写出+回读验证 146 ms。Python/Rust 输出 2,610 行逐轮比较均为 0 差异；输入 SHA-256 未变。Rust 输出工作簿约 103 KB，Python 输出约 116 KB；两边结构和值比较通过。相对 Python 的 CLI 墙钟加速比 8.11×，排除 CLI 启动后的 Rust 内部比 8.87×。

这组简单单表合成输入在同样的 2,610 条输出规模下远低于旧真实工作簿报告的 1,186 ms 写出+验证；旧输入 Sheet/有效区域结构没有保留，无法复现并定位差别。因此没有修改 S3-F2 writer。新驱动和报告记录于本轮工作区，当前未提交；运行不涉及 Cargo。

执行时仓库 HEAD 为 `7517a9c031f9ab1c9b1978dca25dfb609f9e7cac`；本次 Release 二进制由 SHA-256 精确标识。新驱动及报告目前为未提交文件，受项目 AGENTS.md 的禁止默认提交规则约束，本轮不创建新 commit。

### S2-F1 外部核对专项（2026-10-30）

使用 `manual_acceptance/20261027/s2_f01_external_rule_cases` 的正常/异常外部核对输入，各执行一次预热和三轮交错对照。Python 服务中位数为 41.92/42.19 ms；Rust Release CLI 墙钟为 50.26/51.56 ms（Python/Rust 为 0.83×/0.82×），CLI 服务进度阶段为 24 ms（Python/Rust 为 1.75×/1.76×）。独立 Rust CLI 冷启动包含约 26–28 ms；常驻 Web 服务不按每次业务请求支付此进程启动成本。Rust 保存、回读及发布约 15 ms，Python 约 3.8–3.9 ms，但 Rust 端到端服务仍较快；没有证据支持削弱验证或重写 writer。

正常与异常输入均逐轮核对摘要和 3 张输出表（292 个单元格），业务差异 0；预期 0 和 1,233,333.39 外部差值均核对通过。仅 4 个路径文本在 Rust 一侧多出 Windows 扩展路径前缀，规范化后相等。输入哈希及 Python 相关源码哈希前后相同。单独进程运行采样峰值工作集：Python 40.69/40.71 MB，Rust CLI 13.09/12.38 MB；采样进程启动包含在内。

详情：[报告](probe/reports/s2f1_external_profile_20261030.md)、[逐轮 JSON](probe/reports/s2f1_external_profile_20261030.json)、[驱动](tests/compare/perf_s2f1_external_20261030.py)。使用二进制 SHA-256 `9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4`。计时基于 2026-10-30 Windows x64。未修改 S2 业务代码。

### S3-F4 10,000 行输出对拍（2026-10-30）

复用当前 10k 行 CSV、3.0/3.2 配置及两条规则，预热一次、三轮交错运行。Python 服务中位数 7,071.61 ms；Rust CLI 墙钟 989.30 ms，内部服务 945 ms；Python/Rust 倍率分别 7.15×（CLI）与 7.48×（服务）。Rust 阶段：配置 1 ms、输入读取 84 ms、计算 157 ms、写出/验证 702 ms。

两端均产生 20,000 条结果。逐轮工作簿均为 20,001×17，值/公式与关键结构比较 340,017 个单元格，0 差异。输入配置哈希和旧 Python 相关源码哈希前后不变；单独进程树峰值工作集约 Python 181.42 MB / Rust 94.49 MB。工作簿输出分别约 1.39/1.01 MB。当前没有值得优先处理的 Rust 瓶颈；写出占 Rust 用时约 74%，但总耗时仍约 1 秒且比 Python 快 7 倍以上。

详情：[报告](probe/reports/s3f4_parity_profile_20261030.md)、[逐轮 JSON](probe/reports/s3f4_parity_profile_20261030.json)、[聚焦驱动](tests/compare/perf_s3f4_parity_20261030.py)。二进制 SHA-256 `9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4`。无产品代码修改。

### S3-F5 有效表单渲染与输出对拍（2026-10-30）

修正 S3-F5 基准生成器的表单业务类：旧 `AF001`…`AF004` 名称不含 `A1/A2`，Rust 将其归为“本外币”，与输入“人民币”不匹配。旧 A/B 仅生成了 30 个空表文件，`matched_indicators=0`，所以之前报告的 2029/30 S3-F5 时间不可用于业务性能判断。`perf_compare_20261030.py` 的生成器现使用含 `A1` 的表名；聚焦驱动检查命中指标行和数值格，防止空结果再被计作有效性能。

校正输入为 30 机构 × 40 指标 × 4 张表单（每张 A1 表各 40 指标），1,200 条比较记录、4,800 个表单指标行。三轮 Python/Rust 服务中位数 3,509.17/1,510 ms；Rust CLI 墙钟 1,538.36 ms。Python/Rust 倍率为 2.32×（服务）、2.28×（CLI 墙钟）。Rust 读表 36 ms、规划 51 ms、编辑/写入 1,449 ms；逐机构修改中位数 45 ms。没有开展并发写出等代码优化，当前实现已明显快于 Python。

三轮均生成 30 个同名工作簿，值/公式与关键结构逐项对拍无差异；每轮检查 180 张表、35,460 个单元格、19,200 个数值格和 5,160 行隐藏状态。输入哈希不变。另行测得 Python 进程树 475.34 MB / 13 进程、Rust 24.25 MB / 1 进程的采样峰值工作集。

详情：[报告](probe/reports/s3f5_parity_profile_20261030.md)、[逐轮 JSON](probe/reports/s3f5_parity_profile_20261030.json)、[聚焦驱动](tests/compare/perf_s3f5_parity_20261030.py)。有效 Release SHA-256 `9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4`。本轮只修改了测试基准数据生成，不改 S3-F5 业务实现。

### S3-F3 原地导入计划传递优化（2026-10-30）

此前完整输出基准使用 1,000 条唯一匹配反馈、无冲突，Python/Rust 中位数为 977.03/975.25 ms（Rust CLI）和 946 ms（Rust 内部）。阶段诊断显示只读计划命令的内部耗时 369 ms，其中计划 JSON 写入和同步为 291 ms；原地执行随后将该临时计划重新读回，而该暂存计划不会作为用户输出保留。

**修改：** 将规划实现拆为保留公共持久化行为的包装函数和内部执行路径。显式 `s3-plan-explanation-feedback-import` 仍完整写入、同步用户指定的 ExcelEditPlan JSON；S3-F3 原地/副本执行则在保持相同输入哈希保护、计划校验、冲突处理、备份与原子发布流程后，直接把内存中的 ExcelEditPlan 交给 Excel-Core，不再落盘再反序列化临时计划。没有改匹配键或输出业务规则。

优化后使用相同生成器、相同 1,000 条数据值及表结构交错运行三轮。Python/Rust 服务中位数 912.62/203 ms；Rust CLI 墙钟 235.54 ms。相较优化前，Rust 内部总耗时减少 743 ms（78.54%），CLI 墙钟减少 739.71 ms（75.85%）；Python 服务用时两轮分别为 977.03 与 912.62 ms，存在独立批次环境波动。优化前后生成的 XLSX 包大小相差 1 byte，行值和表结构相同；每份输出只修改 H3:H1002，备份逐字节等于原目标。

三轮 Python/Rust 输出值、公式计数与关键结构对拍均为 0 差异；1,000/1,000 行写回、0 未匹配、0 冲突；反馈源未变。独立进程峰值工作集约 Python 38.86 MB / Rust 21.73 MB。Rust 优化后阶段中位数：目标读 11 ms、反馈读 16 ms、Excel-Core 写回 125 ms、总计 203 ms，未分别归因的哈希/备份/发布剩余 51 ms。显式计划命令继续生成 JSON；其独立 CLI 墙钟受启动波动影响，不纳入主 A/B。

定向验证：`cargo test --offline --locked --target-dir F:\\RustBuild\\audit -p audit-core --lib s3_explanation_feedback::tests -- --nocapture`，7 passed；新增执行冲突清单及目标保护测试；`cargo test --offline --locked --target-dir F:\\RustBuild\\audit -p audit-core --test s3_explanation_feedback_plan -- --nocapture`，3 passed；`cargo build --offline --locked --target-dir F:\\RustBuild\\audit --release -p audit-core --bin open-data-audit` 成功。全 workspace 回归未运行。格式检查因仓库多个既有格式差异未通过，未格式化无关文件。未 commit/push。

报告：[优化前 Markdown](probe/reports/s3f3_parity_profile_20261030.md) / [优化前 JSON](probe/reports/s3f3_parity_profile_20261030.json)，[优化后 Markdown](probe/reports/s3f3_parity_optimized_20261030.md) / [优化后 JSON](probe/reports/s3f3_parity_optimized_20261030.json)，[驱动](tests/compare/perf_s3f3_parity_20261030.py)。优化前/后 Release SHA-256 分别为 `9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4` / `d21e355a8ec05129e81ccdda008f257a98702b64537010c24b773433b416db66`。

## 2026-10-30：S3-F1 定点优化

来源：同日性能报告“性能优化迭代”章节及 `s3_execution_compare.rs`。保留逐项输出验证，验证阶段将公式元数据检查从第二次整簿解析改为轻量 XML 节点扫描，并复用格式对象。报告记录验证阶段 936→417 ms、总耗时 1,412→894 ms；另报告 Python 1,032.5 ms、Rust 1,012.9 ms（约 1.02×），并将功能标记暂结。

代码追溯提交：`8b28677`；性能报告归档于 `87fcbca`。

可用的优化后逐轮机器数据没有在已登记 JSON 中找到；原 Markdown 称其为三轮中位数，但这两个数字也等于同一 JSON 中优化前各自的最快轮。故保留用户已确定的“约 1.02×、性能暂结”状态，同时标记优化后倍率证据缺少逐轮原始数据；不重跑已结项基准。

## 2026-10-30：S1-F1 COM 最终批次验收

来源：[最终报告](probe/reports/s1f1_com_final_acceptance_20261030.md)、[逐轮 JSON](probe/reports/s1f1_com_final_acceptance_20261030.json)、[验收驱动](tests/compare/s1f1_final_acceptance_20261030.py)。Python 使用 `DispatchEx`，Rust 使用原生 `CoCreateInstance` 和批次 worker；二者均创建私有 Excel 实例。

| 每批有效簿数 | Python 中位数 | Rust 中位数 | 结果 |
|---:|---:|---:|---|
| 2 | 7,570.60 ms | 7,171.91 ms | Rust 快 5.27%，含一个损坏输入后继续处理。 |
| 10 | 12,246.33 ms | 12,139.69 ms | Rust 快 0.87%。 |
| 33 | 27,006.11 ms | 27,138.94 ms | Rust 慢 0.49%。 |

最终判定为两端平价。COM 冷启动约 2.4–2.5 秒是主成本；Rust worker 启动与 ready 残差约 43 ms/批。公式结果、缓存、输入保护、批次失败隔离及 Excel/worker 清理均通过。此项只证明 COM 重算子步骤，不代表完整 S1-F1 GUI 业务链路耗时。相关代码提交：`f2eed96`（原生 COM 性能路径）及 `000dd54`（批次退出、缓存对拍验收）。

与历史表中的早期 S1-F1 COM 微基准不可合并：早期夹具为每册 7,200 个公式，包含较短的 `CalculateFull` 子步骤边界；最终验收夹具为每册 6,000 个公式和 7,200 个非空缓存，并测量 2/10/33 册批次及专属 Excel 进程退出。两者均未计模板匹配、复制及最终审核结果生成。最终批次数据用于当前 COM 性能结论，早期数据仍作为各自口径的历史证据保留。

## 版本和工作区记录

- 本台账建立时的项目 HEAD：`7517a9c031f9ab1c9b1978dca25dfb609f9e7cac`（`main`）。与性能实现直接相关的历史代码提交：S1-F1 `f2eed96` / `000dd54`；基准驱动在 `af387ef`；S3-F1 功能实现及验证代码在 `8b28677`。
- 本台账、目标进展、基准驱动与报告属于未提交改动；本轮另修改 `audit-core/src/s3_explanation_feedback.rs`。项目 AGENTS.md 明确要求未经单独请求不提交；未 commit/push。S3-F3 定向库测试 7/7、计划 CLI 集成测试 3/3 通过，Release CLI 构建成功；未跑全 workspace 回归。

### 2026-10-30 增量：S3-F3 全冲突输出剖析

新增 [冲突报告](probe/reports/s3f3_conflict_profile_20261030.md)、[逐轮 JSON](probe/reports/s3f3_conflict_profile_20261030.json) 和 [复跑驱动](tests/compare/perf_s3f3_conflict_profile_20261030.py)。合成输入含 1,000 个匹配键，两个反馈来源对每个键提供不同原因，三轮均预期 0 导入、0 未匹配、1,000 冲突。

Python 服务中位数 1,332.49 ms；Rust CLI 墙钟 149.89 ms，Python/Rust 约 8.89×；Rust 服务内部中位数 120 ms。Rust 各轮冲突报告写出 18/19/21 ms。逐轮冲突明细、目标工作簿值/结构对拍差异均为 0；目标字节哈希不变、备份与原目标哈希一致，反馈输入未变。峰值进程树工作集约 Python 41.04 MB、Rust 16.39 MB。文件容器大小不同，不进行普通字体/边框样式对拍。该 100% 冲突负载是压力样本，不能作为实际冲突率预测。

为便于生产批次定位，`S3ExplanationFeedbackExecutionReport` 新增 `conflict_write_time_ms`，由原有规划摘要传至执行报告并随 CLI JSON 输出；未改变冲突判定或写回行为。定向库测试 7/7、计划 CLI 集成测试 3/3 通过；Release 构建成功，SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。Python 驱动 AST 语法及报告 JSON 解析通过，`git diff --check` 通过（仅有 Git 的 LF/CRLF 提示）。未运行全工作区回归；`main` / HEAD `7517a9c`，未提交、未 push。

### 2026-10-30：S1-F1 条件格式异常提取五轮性能剖析

使用专用驱动生成一份 144,034 字节 XLSX：单 Sheet、4,000 数据行、8 列、32,008 个非空值、4 条受支持规则；共扫描 16,000 格并触发 8,610 项。Windows Python 3.12.10 与 Release CLI 使用同一文件，预热一轮后交错执行五轮。计时取 Python `ConditionalFormatService` 与 Rust `conditional-audit` 的内部报告时间；不计启动进程和序列化完整结果 JSON 的耗时。峰值工作集由父进程对子进程约每 5 ms 采样，测量范围包含 Python 解释器及模块。

| 指标 | Python | Rust |
|---|---:|---:|
| 服务内部中位数 | 1,343 ms | 246 ms |
| 区间 | 1,312–1,516 ms | 229–270 ms |
| 规则读取/规范化 | 读取字段报告为 0 ms；规范化中位数 62 ms | 读取中位数 102 ms |
| 条件格式求值 | 844 ms | 94 ms |
| 异常组装 | 453 ms | 24 ms |
| 采样峰值工作集 | 87.94 MB | 85.53 MB |

Python/Rust 倍率为 5.46×。五轮均为 8,610 条异常、缺失/多余/字段差异为 0、诊断为 0；输入哈希前后相同。在该合成规模上，Rust 条件格式子步骤没有性能短板，不实施优化。此前一份真实工作簿的单次记录为 366 个扫描格、348 条异常，Python 375 ms、Rust 618 ms，Rust 读取阶段 578 ms；因原始输入路径未留存且只有单轮，该观测不能与本轮合并或复现，仍提示小负载可能受 Rust 读取成本影响。独立进程墙钟仅供审计，受 Python worker JSON 输出量及进程启动影响，不作为倍率；该专项不覆盖结构预检、模板匹配、公式复制/重算、公式异常及最终结果写出，所以 S1-F1 完整业务链耗时仍未知。

静态调用路径还显示，启用模板结构预检时，源文件先由 `compare_s1_source_structure` 读取，条件格式提取稍后再调用共享工作簿 reader 读取一次；中间存在公式副本和重算阶段。完整流程可能付出重复解析成本，但目前没有同输入端到端分段计时，且缓存整簿投影会增加批量峰值内存，故只登记为下一步剖析候选，不据静态观察修改执行顺序。

复跑：[驱动](tests/compare/perf_s1_conditional_profile_20261030.py)、[逐轮 JSON](probe/reports/s1_conditional_profile_20261030.json)、[报告](probe/reports/s1_conditional_profile_20261030.md)。Rust 可执行文件 SHA-256 `74c167b5a304d357bfd14f4dab3f65f173c63e8c45a0b631bf7323891148edef`。未修改 Rust 业务代码，未运行 Cargo 编译/测试或全工作区回归；无新提交、未 push。

## 2026-10-30：S1-F1 条件格式指标定位优化

### 基准口径纠正

此前的 [S1-F1 条件格式五轮剖析](probe/reports/s1_conditional_profile_20261030.md) 调用 Python `ConditionalFormatService` 时给出的 `structure_ranges=[]`，Rust CLI 也没有模板结构区域参数。因此其 1,343/246 ms、5.46×只覆盖轻量规则求值和异常构造，没有覆盖 S1-F1 生产流程中基于模板结构区域解析 `check_field` 的重复前缀扫描；原始文件保留，本次已在性能基线中标清边界。

### 对齐生产结构区域后的结果

新驱动生成 4,000 行、32,008 个非空值、4 条规则的 XLSX，传入 `S1Conditional!A1:H1` 作为模板结构区域。扫描 16,000 格并生成 8,610 条异常；预热一轮后交错测量五轮。优化前为同一驱动在原 Release 二进制上的实测，优化后使用新 Release 二进制。XLSX 容器分别为 144,031 / 144,030 字节，值和结构由同一生成器产生；每一轮 Python/Rust 共用同一对输入。

| 指标 | 优化前 | 优化后 |
|---|---:|---:|
| Rust 条件格式异常组装中位数 | 15,523 ms | 57 ms |
| Rust 结构预检 + 条件审核 developer event 中位数（不含进程启动） | 16,044 ms | 498 ms |
| Python 结构预检 + 条件审核中位数 | 20,859.4 ms | 19,264.7 ms |
| 优化后 Python/Rust 内部计时参考倍率 | — | 38.68×（均不含外层进程启动）|
| Rust 采样峰值工作集 | 86.61 MB | 86.56 MB |

Rust 异常组装耗时约减少 99.63%，结构预检+条件审核 CLI 耗时约减少 96.90%。Python 没有修改，两轮重新生成 XLSX 的计时差异不作为 Python 优化收益。优化前后每轮均为 8,610 条异常；优化后五轮规范化异常字段 SHA-256 多重集 5/5 一致，结构标签 8/8 匹配，源/模板哈希不变。

### 实现与验证

旧 Rust 逻辑为每条触发异常独立扫描目标列从结构起始行至目标行，重复计算相同前缀。现在按工作表和目标列分组，按目标行一次累积相同的有序去重文本前缀，再在原位置组装异常；左侧标签、合并单元格锚点、数值过滤、去重顺序和输出顺序保持不变。新增乱序/重复目标、合并锚点和前缀语义对照单测。

- `cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture`：5 passed / 0 failed。
- `cargo build -p audit-core --bin open-data-audit --release`：成功，使用 `CARGO_TARGET_DIR=F:\RustBuild\audit`。
- 新 Release SHA-256：`d7c49818f3905a5f31da7a351cfd8d44e5bc20fb3cfe1e8b76992bba6b536e56`。
- 未运行 workspace 全回归、真实机构 GUI 或 Linux/Kylin/UOS 验收；未 commit/push。

复跑文件：[报告](probe/reports/s1_indicator_resolution_optimization_20261030.md)、[优化前机器数据](probe/reports/s1_preflight_reads_20261030.json)、[优化后机器数据](probe/reports/s1_preflight_reads_20261030_after.json)、[驱动](tests/compare/perf_s1_preflight_reads_20261030.py)。

结构预检和条件审核仍各解析一次源工作簿，4,000 行样本各约 99/97 ms。两步之间有副本复制和 COM 重算；本轮未跨阶段持有所有输入文件的完整工作簿投影。是否合并读取应由完整链路和多文件峰值内存数据决定，本轮只改已证实的指标定位瓶颈。此段描述的是当时状态；后续共享读取优化见本节末尾。

## 2026-10-30：S1-F1 结构预检与条件审核复用源工作簿读取

静态调用检查确认，生产 S1-F1 中结构预检使用 `compare_s1_source_structure` 读取源工作簿，条件格式审核随后再读取同一源工作簿。两步之间的公式复制/COM 重算不改变原始源，因此两项源侧检查可使用同一只读 `WorkbookDump`。现在 S1-F1 在单个输入工作簿作用域内读取一次，先完成结构匹配，再从同一投影提取条件格式异常；投影在进入后续复制/重算前释放。没有跨工作簿批次缓存，也没有把审核投影带过公式重算阶段。

复用路径新增 `compare_s1_source_structure_dump` 和 `extract_conditional_issues_from_dump`；原文件读取 API 保留并委托至相同处理逻辑。CLI 单文件条件审核、S1-F1 目录批次和条件审核入口都使用该组合路径。进度事件将读取时间归入结构预检并标记 `source_read_reused=true`，避免把重复读取成本或缓存驻留时间藏在审核计时中。新增单测对比直接读取与复用读取生成的异常、触发规则、诊断、未解析指标和批注状态。

使用与前一轮相同生成器及业务内容的 4,000 行夹具，结构区 `S1Conditional!A1:H1`、32,008 个非空格、4 条规则、16,000 个扫描格、8,610 条异常；预热 1 轮、交错 5 轮。重新生成的 XLSX 容器字节数与前状态相差 1 byte（144,030→144,031），各自每轮 A/B 使用完全相同输入；不声称前后 XLSX 二进制哈希相同。

| 指标 | 指标定位优化后、仍读取两次 | 共享一次源读取后 |
|---|---:|---:|
| Python 预检+条件审核中位数 | 19,264.677 ms | 19,222.839 ms |
| Rust CLI developer event elapsed（不含进程启动） | 498 ms | 390 ms |
| Python/Rust（参考） | 38.68× | 49.29× |
| Rust 源文件读取 | 约 99 ms + 97 ms | 93 ms 一次 |
| 采样峰值工作集 Python/Rust | 77.402/86.562 MB | 77.312/85.949 MB |

Rust 内部事件耗时减少 108 ms，约 21.7%；Python 时长不是优化对象，前后运行分别为 19,264.677/19,222.839 ms。两边都不计外层进程启动，倍率仅描述此专项，不外推到完整 S1-F1；本轮记录未保存 Rust 外层进程墙钟中位数。新旧合成文件哈希不同，夹具数据和结构由同一生成逻辑产生，报告明示容器 1 byte 差异。

五轮中每轮 Python/Rust 均产生 8,610 条异常，规范化异常 SHA-256 多重集完全相同；结构标签 8/8 匹配。新轮源与模板 SHA-256 在运行前后均不变；五轮工作集峰值与原 Rust 状态相当，没有观察到持有源投影导致的内存增长。一轮独立烟测亦通过同样对拍，Rust 峰值工作集 85.934 MB。

定向验证：`cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture`、`cargo test -p audit-core --lib s1_structure_compare::tests`、`cargo test -p audit-core --lib conditional_format::tests` 和 `cargo test -p audit-core --bin open-data-audit --no-run` 在本次代码状态运行通过；Release 构建成功，A/B 二进制 SHA-256 为 `78d0d25b238ff6190a6e42adef6c6e60a209526802964c335560dd25f0eed230`。未运行 workspace 全量回归、GUI/真实机构完整链或 Linux 目标环境；无提交、无 push。

报告与复跑入口：[优化报告](probe/reports/s1_source_read_reuse_optimization_20261030.md)、[五轮逐轮数据](probe/reports/s1_preflight_reads_20261030_shared_read_after.json)、[驱动](tests/compare/perf_s1_preflight_reads_20261030.py)。S1-F1 完整 GUI 业务链与模板复制/公式异常/结果写出仍需后续独立计时。

## 2026-10-30 增量：S2-F1 配置读取与输出验证优化

- **配置复用**：S2 配置检查和服务执行原先各自打开配置工作簿，运行路径最多三次解析同一配置。现在配置校验和运行配置从同一份 `SheetValues` 生成，保留公开校验报告接口。五轮、728 指标/7 机构配置下，Rust 配置阶段中位数 40→15 ms，服务 64→38 ms，CLI 94.37→66.44 ms；五轮摘要和 292 个输出单元格均一致，输入哈希不变。
- **输出验证**：原验证用完整 `read_workbook` 重新解析所有值和元数据，但实际只核对可重开、工作表名称/顺序、公式数。新增 Calamine 的 Sheet 名称/顺序及公式范围扫描，移除不需要的整簿值读取。合成 21 本本期+21 本上期、7 机构、3 表单、728 指标、每期 5,083 条记录的三轮对照中，保存阶段中位数 312→98 ms（减少 68.6%），Rust 服务 659→429 ms（减少 34.9%），CLI 墙钟 738.45→500.95 ms（减少 32.2%）。峰值工作集由 145,477,632 降至 107,913,216 bytes。
- **正确性和边界**：三轮摘要及 81,372 个输出单元格逐项比较均为 0 差异，输入哈希不变。输出 xlsx ZIP 容器长度有少量差异，不声称二进制一致。Python 未修改，服务耗时分别为 1,101.92/1,144.47 ms，存在独立运行自然波动。该批次由同一合成生成器建立；历史真实 21+21 工作簿已不可用，不能替代真实输入复验。
- **验证**：配置检查和服务定向 parity 各 1/1 通过；轻量工作簿验证单测 1/1 通过；Release 构建、两份 Python 驱动 `py_compile`、`git diff --check` 通过。未做全 workspace 回归、GUI 人工验收、真实历史数据复跑或 Linux/Kylin/UOS 验证。没有 commit/push。
- **材料**：[专项报告](probe/reports/s2f1_performance_optimization_20261030.md)、[配置 A/B](probe/reports/s2f1_config_read_reuse_20261030_before.json) / [后](probe/reports/s2f1_config_read_reuse_20261030_after.json)、[批次验证 A/B](probe/reports/s2f1_batch_scale_20261030.json) / [后](probe/reports/s2f1_batch_scale_20261030_after_save_validation.json)、[配置驱动](tests/compare/perf_s2f1_config_read_reuse_20261030.py)、[批次驱动](tests/compare/perf_s2f1_batch_scale_20261030.py)。

## 2026-10-30 增量：S1-F1 批次内模板预检缓存

- **瓶颈及变更**：`run_s1_f1_worker` 原来为每本源工作簿重新读取同一模板的规则、条件/结构映射、定义名称、Sheet 名和结构标签。现在在单次运行作用域内以 canonical 模板路径缓存预检结果，并通过 `Arc` 分享不可变计划；错误也按本次批次缓存。不同模板仍各自预检，源文件仍分别结构比对。
- **读取剖析**：20,000 行 Rust-only 合成工作簿（1 sheet、160,008 非空值、4 条 CF 规则、80,000 扫描格，源和模板各 716,663 bytes）预热 1 次、测量 3 次。中位数：源读取 490 ms、模板预检读取 528 ms、CF 计算 523 ms、异常组装 314 ms、CLI 墙钟 3,254.905 ms、峰值工作集 385.7 MiB。每轮 43,051 条异常摘要一致、结构标签 8/8 命中、输入哈希不变。
- **收益边界**：同模板单次预检实测 528 ms；同批次有 3/21 本共享该模板时，避免的重复读取估算为约 1.06/10.56 秒。该值为阶段成本乘以缓存命中数的推算，未测成 GUI 批次 A/B。20k Python 路径超过 150 CPU 秒仍未结束，已停止；不报告 20k 跨语言倍率。既有 4k 五轮 Python/Rust A/B不变。
- **业务验证**：缓存引用单测、S1-F1 两文件混合重算模拟及条件提取失败保留公式结果测试均通过。Windows 三文件完整 GUI 合成路径中，3/3 文件模板与结构预检通过、输入哈希不变；当前主机 COM worker 无法启动且无 LibreOffice，系统保留 3 本人工重算副本、无假结果。该次没有验证成功公式重算或最终结果工作簿写出。
- **构建和限制**：Release 构建成功，SHA-256 `6fdff754c20b2da32946a9b957772c0fcbb4a356f4a2e451be745876140768ff`。未运行全 workspace 回归、真实机构人工验收或 Linux/Kylin/UOS；模板/源整簿读取本身仍是下一候选，历史 983 ms 输入不可复测。
- **材料**：[模板缓存优化报告](probe/reports/s1_template_preflight_cache_optimization_20261030.md)、[20k Rust-only JSON](probe/reports/s1_rust_preflight_scale_20261030.json)、[复跑驱动](tests/compare/perf_s1_rust_preflight_scale_20261030.py)。

## 2026-10-30 增量：S1-F1/F2 模板推荐元数据复用

S1-F1 GUI 批次及 S1-F2 多文件自动识别原先为每本源簿调用模板推荐函数；该调用会重新扫描模板目录、逐个读取候选模板的 OOXML Sheet 目录。现在两处入口各自在本次操作开始时加载一次 `TemplateProfile` 列表，后续仍逐本读取源簿工作表目录并独立匹配。单文件 API 保持原有签名和行为；不同批次会重新加载模板档案。

Release-profile 对照在同一夹具上交替重复 5 轮，预热 1 轮：21 个源簿（各 8 Sheet、总 183,519 bytes）和 5 个模板簿（各 9 Sheet、总 46,312 bytes），无数据单元格。每个旧式路径调用对一个源簿重新加载全部模板；共享路径在相同批次只加载模板列表一次，两者均逐本读取 21 个源簿。模板推荐阶段中位数为 30.2814 ms → 5.9081 ms，减少 24.3733 ms / 80.5%，子步骤加速 5.13×。21 个源簿的推荐模板、置信度、说明和候选结果完全一致。

这项优化的绝对节省很小，但消除了批次内重复的目录扫描和 105 次模板 Sheet 目录读取中的 100 次冗余读取，影响的是 S1-F1 模板匹配及 S1-F2 多文件自动模板识别子步骤。它不表示整个 S1-F1 或 S1-F2 加速 5×；没有测量峰值内存或完整 GUI/S1-F2 导出时间。三个文件 GUI S1-F1 合成路径定向测试通过；重算引擎在当前主机不可用时仍按设计输出人工重算状态。Release 主程序构建成功，SHA-256 `7be5fdece2ae26e931feb796daa7075395ab3fa44591cc242ba0813beba96792`。

复跑入口为 `tests/compare/perf_s1_template_profile_reuse_20261030.py`，逐轮报告为 `probe/reports/s1_template_profile_reuse_20261030.json`，测试为 `tests/s1_template_profile_reuse.rs`。源码基于 HEAD `7517a9c`，工作区有未提交 WIP，本轮未提交、未 push；全 workspace 回归和 Linux/Kylin/UOS 验证未运行。全仓 `cargo fmt --all --check` 仍被仓库既有格式差异阻断；本次新增测试和 `s1_template.rs` 的定向 rustfmt 检查通过。

## 2026-10-30 增量：S1-F1 模板结构预检投影读取

结构预检原先用通用 `read_workbook` 读取整本模板并构造所有值/公式 JSON 映射，同时解析批注、条件格式和数据验证等无关元数据。现在新增结构专用 Calamine 投影读取：只保留工作表目录、结构映射区域值/公式及合并区域；所有工作表值仍由 Calamine 读取以保留原错误行为。遇到 Calamine 不支持错误 token 时沿用既有 OOXML 精确临时规范化回退。公式文本、类型化值、合并锚点仍参与原结构判断。

复用相同 20,000 行 Rust-only 生成器，预热 1 次、3 轮；源/模板各 716,663 bytes、1 张表、20,001 行×8 列、160,008 非空格、4 条条件格式规则、80,000 扫描格。优化前/后 XLSX SHA 不同（故不称逐字节同输入），但文件尺寸、工作簿形状和生成器相同。结构标签每轮 8/8，43,051 条异常的规范化 SHA-256 与旧三轮及新三轮完全相同，运行前后源/模板哈希不变。

模板结构预检读取中位数 528→224 ms，减少 304 ms/57.6%。内部命令 2,165→1,654 ms，CLI 墙钟 3,254.905→2,645.195 ms；总耗时的 18.7% 变化不能全部归因于本次读取器，未修改的源读取与条件计算也有轮间变化。峰值工作集 385.672→385.824 MiB，基本不变，受后续条件审核全量读取主导。该优化仍不构成 Python/Rust 20k 对比，Python 参考运行之前超过 150 CPU 秒后停止。

`excel-reader` 投影与全量读取值/公式/Sheet/合并区对拍、非锚点合并读取及未知错误值兼容回退共 3/3 通过；S1 结构对比库测试 7/7 通过；Release 构建成功，SHA-256 `a2ac5bc3fd16223a3c18976f9ed6042e941af6df17fc2fbd84ad5fe54b82f370`。首次未限并行的 `audit-core` 测试编译因 Windows 页面文件不足失败；只运行库测试并将构建任务设为 1 后通过，未做全 workspace 回归。未测完整 GUI/公式复制/COM/结果输出或真实机构样本；未测试 Linux/Kylin/UOS。

报告：[专项报告](probe/reports/s1_structure_projection_optimization_20261030.md)、[优化后逐轮 JSON](probe/reports/s1_structure_projection_optimization_20261030.json)、[复跑驱动](tests/compare/perf_s1_rust_preflight_scale_20261030.py)。

### 2026-10-30 增量：S1-F1 源结构预检投影路径同输入对照

为分开确认源文件路径是否也受益，新增忽略式 Release 性能测试和驱动：旧路径为完整 `read_workbook` 再执行既有结构比较；新路径为 Calamine 结构投影后执行同一结构比较。20,000 行数据加表头、1 张表、160,008 个非空格、716,662 bytes 的同一生成文件同时用作源和模板；两者 SHA-256 均为 `5faab2e54f24a465da130fcd5091be8849abc39c493f7b9012c48077e151b23d`。预热 1 次、交替运行 5 次，中位数 501.532→237.579 ms，投影路径为 2.111×，减少约 263.953 ms。

每轮均对比匹配状态、分数、命中/检查标签、缺失工作表和诊断，5/5 完全一致且为 8/8 标签命中；运行前后输入哈希不变。旧路径 501.532 ms 中读表占 501.458 ms、结构比较仅 0.054 ms。投影仍由 Calamine 遍历工作簿值；本轮未测独立 RSS、Python 路径、真实机构文件或完整 S1-F1，也没有改产品业务代码。定向命令成功编译并运行单个性能测试：`cargo test --locked --release -p audit-core --test s1_structure_projection_source_perf profile_source_structure_projection_same_input -- --ignored --nocapture --test-threads=1`。

报告、逐轮 JSON、复跑驱动及 Rust 测试见 `probe/reports/s1_structure_source_projection_20261030.md/.json`、`tests/compare/perf_s1_structure_source_projection_20261030.py` 和 `tests/s1_structure_projection_source_perf.rs`。现有未提交改动保留；本轮无 commit/push。完整 S1-F1 仍需独立计时，整个 S1/S2/S3 性能目标继续进行。

### 2026-10-30 增量：S3-F2 进程树峰值工作集专测

S3-F2 的 7,830 行扩展基准已有阶段耗时和输出对拍，但未采集峰值内存。为避免重跑已有耗时基准，本次只测内存：同一合成输入（比较文件 7,830 行×12 列、2,610 行选中；映射文件 401 行）分别由 Python 参考函数和 Rust Release CLI 执行，进程树每 5 ms 采样一次 RSS/工作集总和；每实现 3 次并交替先后。三轮峰值中位数为 Python 43.688 MB、Rust 44.340 MB，最大值分别为 43.945/44.383 MB，Rust 高约 0.652 MB（1.5%），无明显内存优化目标。

每轮均比较完整 2,610 行输出，单元格差异为 0，输入 SHA-256 前后不变；两种执行均只观察到单进程。采样可能遗漏不足 5 ms 的峰值，指标包含新进程启动，但不含输入生成和输出比较。输出二进制文件大小与此前耗时基准只差约 1 byte，输入由同一生成器重建；未声称与此前测时夹具逐字节相同。未记录或比较耗时，本次仅补齐内存维度。

复跑驱动、JSON 和报告见 `tests/compare/perf_s3f2_peak_memory_20261030.py`、`probe/reports/s3f2_peak_memory_20261030.json` 与 `probe/reports/s3f2_peak_memory_20261030.md`。Python 来源保持只读；没有修改业务代码、没有 Cargo 编译或测试、没有 commit/push。S3-F1 峰值内存及 S1-F1 完整业务链仍待补证。

### 2026-10-30 增量：S1-F2 进程树峰值工作集专测

补采 S1-F2 轻量样本的峰值工作集，不重复耗时基准：2 本源簿、3 张源表、5 行源数据、9 行输出。每轮独立运行 Python `ExcelSession`/Excel COM 与 Rust Release CLI，预热外测量 3 轮、顺序交替；每 5 ms 采样 worker 进程树并纳入本轮新启动的 Excel PID，排除运行前已存在的 Excel PID。

Python+Excel COM 峰值工作集中位数为 300.078 MB、最大值 301.574 MB；Rust 为 10.719/10.727 MB，Python 约为 Rust 的 28 倍。三轮输出值与结构比较通过，值差异为 0；列宽最大差 0.535，与既有 S1-F2 对拍记录一致。源簿和模板哈希不变；三次测试启动的 Excel PID 在结束后均不再运行，没有关闭预先存在的 Excel 实例。较大内存主要是 Excel COM 引擎开销，而 Rust 已约 10.7 MB，暂无需要继续压缩的 Rust 内存瓶颈。

这是五行合成数据的 memory-only 结果；采样可能遗漏 5 ms 以下峰值，包含 Python/Excel 启动，不作为大规模内存曲线。报告和逐轮数据见 `probe/reports/s1f2_peak_memory_20261030.md/.json`，驱动见 `tests/compare/perf_s1f2_peak_memory_20261030.py`。未修改业务代码；未运行 Cargo 构建/测试、未提交、未 push。S3-F1 峰值内存和 S1-F1 完整业务链耗时/内存仍待补证。

### 2026-10-30 增量：S3-F1 峰值工作集专测

S3-F1 耗时基准按用户要求暂结，本次只补其缺少的内存维度。采用既有生成器生成两份 10,000 行、14 列 CSV 和 3.0/3.1 配置；Python 参考服务与 Rust Release CLI 均在新 worker 中运行三轮并交替先后，每 5 ms 采样 worker 进程树 RSS/工作集。峰值中位数 Python/Rust 为 133.477/119.035 MB，最大值 133.672/119.098 MB，Rust 少约 14.442 MB / 10.8%。

每轮检查输出工作表与尺寸，并逐项比较 220,022 个单元格；1e-10 容差外差异 0，舍入到 10 位小数的差异 0，最大浮点尾差 `2.117582368135751e-22`，公式数均为 0。输入 SHA-256 前后不变。该测量没有收集或比较耗时，因此不改变用户确认的约 1.02×结项结论。采样可能漏掉小于 5 ms 的峰值；合成数据不能代表真实业务配置的内存曲线。

复跑脚本与机器数据/报告见 `tests/compare/perf_s3f1_peak_memory_20261030.py`、`probe/reports/s3f1_peak_memory_20261030.json` 和 `probe/reports/s3f1_peak_memory_20261030.md`。本次未修改业务代码、未运行 Cargo 构建/测试、未提交、未 push。剩余明显证据缺口集中在 S1-F1 完整业务链耗时与内存。

## 2026-10-30 增量：S1-F1 完整业务链三轮对照与结构纠偏

首次将 Python 原项目的 `AuditService + run_dag_flow(dag:汇总核查表校验)` 与 Rust Release GUI 服务入口做同输入完整链路对照：3 个源簿、1 个 2-sheet 模板、1 个 2-sheet 外部簿、1 个 5-sheet 配置；2 本被审核副本，含公式重算、条件格式异常和结果工作簿输出。三轮交替中位数：Python/Rust 业务步骤 10,627.8/6,983 ms（1.522×）；新 worker/CLI 墙钟 15,552.42/8,943.08 ms；峰值工作集 580.121/504.473 MB。测试夹具合计 37,755 bytes，结果 3 行。

每轮结果工作簿均比较 52 格且差异为 0；公式文本、公式缓存和独立期望值一致；两本审核副本均复制 24 格“审核规则”表，工作表名/顺序与 Python 相同，引用的外部表存在、未引用表省略。Excel COM 打开两边的所有输出副本并正常退出；所有主输入哈希不变；每轮新建 Excel 进程均已退出。三轮全部正确性及清理检查通过。

对拍还发现并修正两个真实差异：Rust 过去没有将结构化模板的“审核规则”页复制进审核副本；审核结果中 integral numeric 说明值曾显示 `.0`。现在 S1 公式复制计划复制该规则表，并按旧流程替换源文件中的同名表；结果值格式与 Python 一致。这些是业务对齐修复，不计为性能优化收益。定点 `s1_formula_copy` 测试 16/16 通过；整数显示测试 1/1 通过；Release 构建成功。未运行全 workspace 回归、真实机构工作簿或 Linux/Kylin/UOS 测试。

业务时间使用 Python `run_dag_flow` 内部计时和 Rust acceptance adapter 步骤计时；Rust 有约 100 ms 轮询粒度。Rust 阶段明细尚未进入 acceptance 报告，因此 Rust 各阶段 A/B 分解仍缺；墙钟与峰值内存单独记录。完整测试方法、逐轮耗时、结果和限制见[全链路报告](probe/reports/s1f1_full_chain_20261030.md)及[机器数据](probe/reports/s1f1_full_chain_20261030.json)，复跑驱动见 `tests/compare/perf_s1f1_full_chain_20261030.py`。本轮未提交、未 push；工作区原有修改均保留。

### 2026-10-30 增量：S1-F1 批注作者对齐与全链路基准复测

复测夹具新增“审核规则”!A1 批注，用 Python 的 openpyxl 设置文本与作者，以覆盖旧 Python `Worksheet.Copy` 会保留的批注信息。首次实跑发现 Rust 副本最终批注作者变为默认值；现由 OOXML Patch 仅在源与重算输出批注坐标/文本一致时，在最终发布前恢复该表的原始 comments XML/VML；位置、文本变化或非批注 VML 均失败关闭。这样保留作者和批注富文本，不扩大到普通样式复制。

定点 `s1_formula_copy` 为 17/17，OOXML 批注/VML 回归 3/3；重建 Windows Release 后，完整链路 1 轮烟测及 3 轮正式复测通过。三轮业务步骤中位数 Python/Rust 7,226.75/6,890 ms（1.049×，近似持平）；进程树峰值工作集中位数 502.141/297.863 MB。worker 墙钟中位数 32,747.97/6,965.34 ms，计时延伸到本轮新 Excel 进程退出，Python 路径观察到每轮 2–3 个新 PID、Rust 每轮 1 个；它们均自然退出，墙钟差异不能解释为业务计算快 4.70×。

3/3 轮结果工作簿 52 格差异为 0；公式文本、缓存和独立期望值一致；审核规则表内容/批注文本/作者、Excel COM 打开、源输入 SHA-256 不变均通过。逐轮 JSON、方法限制和 Release SHA 见 `probe/reports/s1f1_full_chain_20261030.json/.md`。这是此前全链路三轮基准的增强夹具复测，历史旧数值保留在上文；当前基线按这次包含批注的结果更新。全工作区回归、真实机构业务文件和 Linux/Kylin/UOS 性能仍未测；未提交、未 push。

### 2026-10-30 增量：S1-F1 acceptance 微秒阶段计时

- **记录范围**：验收输出新增请求准备/源目录扫描、模板档案与匹配、模板预检、源结构对比、工作流准备、每本副本编辑/验证/重算、公式与条件格式异常提取、结果合并写出、专用 Excel 会话初始化与关闭时间。阶段计时以微秒输出；Excel-Core 与 COM 报告原生只有毫秒精度的阶段仍以该精度换算。
- **测量数据**：单轮诊断 fixture 为 3 源簿、2 本通过、1 本跳过，加模板/外部/配置各 1 簿，总输入 37,755 bytes。Rust 输入扫描/其余请求准备合计 0.680 ms；模板档案加载、匹配、预检、结构对比分别 0.970/0.686/10.084/1.312 ms；两本审核副本复制/重算准备 3,632.188 ms，其中 Excel-Core 编辑/输出验证 156 ms、真实 Excel 重算 3,381 ms、其他准备/校验 95.188 ms；公式/校验与条件格式提取总 12.494 ms；合并历史并写结果 42.119 ms；Excel 会话关闭 2,929.371 ms，adapter 总计 6,660.763 ms。
- **瓶颈判断**：COM 首本重算约 3.199 秒，第二本热会话重算约 0.182 秒；退出 Excel 仍约 2.93 秒，与 2/10/33 册最终批次中约 2.74 秒的 Office 生命周期时间接近。没有低风险理由放松 Excel 专用进程隔离、等待退出或重算/缓存验证；不做代码优化。
- **计时限制**：阶段测量期间 Rust-Tools 的 XLS 转 XLSX COM 基准及多个 Excel COM 进程仍运行。单轮 Python/Rust 8,622.57/6,681 ms 与随后三轮粗计时（7,974.05/8,516 ms 中位数）均不纳入正式性能基线。两轮/三轮完整输出、公式/缓存、批注、Excel 打开、输入 SHA 和新建 Excel 清理检查通过。S1-F1 正式完整链路基线继续采用先前干净三轮中位数 7,226.75/6,890 ms（1.049×）。
- **历史文件修复**：一次诊断复跑使用了原本固定的 `s1f1_full_chain_20261030.json` 输出名，覆盖了原 clean-run JSON。当前探针另存为 `s1f1_full_chain_ambient_probe_20261030.json`；干净三轮的汇总、逐轮总耗时、内存和正确性结果仍在全链路 Markdown，另有明确标注的恢复版摘要 JSON。原始 clean-run JSON 内的阶段日志、逐轮输入哈希及逐轮内存样本无法恢复。现在 S1-F1 默认报告使用含微秒的 UTC 时间戳，S1-F1 与 S1-F2 性能驱动都以独占创建方式写报告；已存在的目标路径会被拒绝，避免再次覆盖历史证据。
- **验证**：S1-F1 定向测试 `s1_f1_` 4 passed / 0 failed / 2 ignored；Release 构建成功，SHA-256 `D9823BE57071493ACB04834F4144A3F91631EF3FDC4F576344E130AF59C7B1B8`；Python benchmark AST 语法解析成功。`rustfmt --check` 在 `web_bridge.rs` 仍报告若干仓库既有格式差异；新增代码格式已按输出调整。未跑全 workspace 回归；未 commit/push。
- **材料**：[微秒阶段报告](probe/reports/s1f1_stage_profile_20261030.md)、[单轮 JSON](probe/reports/s1f1_stage_profile_20261030.json)、[三轮并发负载探针](probe/reports/s1f1_full_chain_ambient_probe_20261030.json)、[干净三轮恢复摘要](probe/reports/s1f1_full_chain_clean_summary_reconstructed_20261030.json)。

### 2026-10-30 增量：S1-F2 Rust-only 阶段诊断

既有 S1-F2 对拍驱动已能输出 Rust 内部阶段时间，但原三轮汇总未保存这些字段。本次新增不启动 Excel COM 的 Rust-only 阶段驱动，复用同一合成夹具生成逻辑：2 本源簿（5,429/4,938 bytes）、3 张源表、5 行源数据、模板/映射配置 5,672 bytes、9 行结果。Windows Release reader/writer 为 Calamine + namespace-aware OOXML / rust_xlsxwriter 0.101.0。

三轮 Rust CLI 墙钟中位数 61.060 ms，内部总耗时中位数 22 ms；发现、模板读取、结构预检、源读取、计算、历史、写入、读回验证阶段中位数分别为 0/2/2/2/0/0/3/3 ms。输出表结构和全部单元格值跨三轮一致，输入 SHA-256 前后不变。计时为整数毫秒，小于 1 ms 的阶段显示为 0；CLI 墙钟与内部阶段总耗时边界不同，不相减解释单一阶段。

运行时检测到 21 个 Excel COM 或 Rust-Tools XLS 转换相关进程，因此该采样只用于阶段定位，JSON 明确标记 `performance_eligible_for_baseline=false`；没有据此更新 Python/Rust 倍率或改产品代码。既有全量 Python/Rust 三轮值与结构对拍、4,320/59 ms 中位数保持为正式 S1-F2 性能证据。当前诊断未显示值得继续优化的 Rust 业务阶段。

驱动、逐轮 JSON 和报告见 `tests/compare/perf_s1f2_stage_profile_20261030.py`、`probe/reports/s1f2_stage_profile_20261030.json`、`probe/reports/s1f2_stage_profile_20261030.md`。Release SHA-256 为 `d9823be57071493acb04834f4144a3f91631ef3fdc4f576344e130af59c7b1b8`；测量时 Git HEAD 为 `7517a9c`，相关工作区仍未提交（项目规则禁止未经请求提交）。Python AST 检查及输出可读性/跨轮一致性检查通过；没有 Cargo 编译/测试、没有 Git push。后续如需把阶段计时用于正式基线，应在无并发负载时重采。

## 2026-11-04：S1-F1 Linux LibreOffice GUI 批次复用

Linux GUI 将同一审核批次中可准备的公式副本集中送入一次 `recalculate-batch`。该 sidecar 调用为整批启动一次使用独立 profile 的 `soffice --convert-to`，顺序处理工作簿并验证各本缓存；批次结束确认专用 profile 无剩余进程。实现使用批量转换通道，不是常驻或跨文件 UNO 文档会话。重算失败的工作簿保留无缓存公式副本供人工重算，避免从错误批次路径退回逐本冷启动。

WSL Ubuntu / LibreOffice 24.2.7.2，单轮 400 行合成文件：2 本逐本/批次 5,933/4,987 ms（1.19×）；10 本 32,813/10,575 ms（3.10×）；33 本 80,381/21,837 ms（3.68×）。每档 `soffice_conversion_invocations=1`、profile 残留进程数为 0；输出公式文本、401 个公式缓存和输入字节保护均通过。Linux `audit-core --lib s1_f1` 4/4、LibreOffice sidecar 集成测试 2/2 通过；Windows 公式复制定向测试 17/17、sidecar 集成测试目标编译通过；忽略的性能测试手动运行通过。单轮合成结果只说明复用效应，不是 Python/Rust 或完整业务 A/B。真实 S1-F1 文件、Linux 桌面 GUI、Kylin V10/UOS V20 尚待人工验收。详见 `probe/reports/s1f1_lo_batch_convert_20261104.md/.json`。
