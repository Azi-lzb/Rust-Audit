# S1-F1 Windows COM 最终性能与批量稳定性验收

日期：2026-10-30  
结论：**S1-F1 原生 COM 批量性能优化完成；未发现值得继续实施的安全低成本优化。**

## 范围与测量口径

这是 Windows x64 上的 COM 重算 worker 批次验收，不是完整 S1-F1 GUI 业务验收。所有输入都是本地生成的合成 XLSX：每册 1,200 行，包含 6,000 个 VLOOKUP 公式和 7,200 个有值单元格。公式引用同一工作簿的“参照表”。两通道均逐工作表调用 `Calculate`、保存为 XLSX、关闭，并由独立 openpyxl 只读检查公式文本和缓存值是否符合逐格预期。

Python 基准使用 `win32com.client.DispatchEx("Excel.Application")`；Rust 使用每批一个独立 Audit worker，并由 `CoCreateInstance(CLSCTX_LOCAL_SERVER)` 新建 Excel。两边均在 ready 时核对新 Excel PID 及进程创建时间，拒绝连接验收开始前已存在的 Excel。Python COM 驱动运行在已启动的基准进程中；Rust worker 每批新建。Python 的解释器/基准宿主启动不计入批次，Rust worker 启动计入。

计时从每批 COM 初始化或 worker 启动开始，覆盖 Excel ready、逐册打开/计算/保存/关闭、公式缓存读取与核验、关闭 Office、等待新建的 Excel 进程确实退出，以及输入哈希和进程清单复核。夹具生成与输入副本准备在计时前完成；汇总 JSON 写盘不计入。Rust 的 JSON 行协议往返包含在启动、请求或关闭计时内。各阶段相互嵌套时不重复求和；每批以实际总时长与该批各计时段核对。

批量次数：2 册五轮，10 册两轮，33 册两轮；Python/Rust 轮次交错。每个通道各做一次预热但不纳入中位数。第一轮 2 册测试在两份有效工作簿之间放入一个损坏的 XLSX，验证单文件失败后继续处理。

## 端到端结果

| 每批有效工作簿 | 引擎 | 批次中位数 | 中位每册 | 结果 | Excel 实例 |
| ---: | --- | ---: | ---: | --- | ---: |
| 2 | Python | 7,570.60 ms | 3,785.30 ms | 10 成功、1 失败、0 跳过 | 5 批 / 5 实例 |
| 2 | Rust | 7,171.91 ms | 3,585.95 ms | 10 成功、1 失败、0 跳过 | 5 批 / 5 实例 |
| 10 | Python | 12,246.33 ms | 1,224.63 ms | 20 成功、0 失败、0 跳过 | 2 批 / 2 实例 |
| 10 | Rust | 12,139.69 ms | 1,213.97 ms | 20 成功、0 失败、0 跳过 | 2 批 / 2 实例 |
| 33 | Python | 27,006.11 ms | 818.37 ms | 66 成功、0 失败、0 跳过 | 2 批 / 2 实例 |
| 33 | Rust | 27,138.94 ms | 822.39 ms | 66 成功、0 失败、0 跳过 | 2 批 / 2 实例 |

Rust 相对 Python 在 2 册批次快约 5.27%，10 册快约 0.87%，33 册慢约 0.49%。方向随批次规模变化；单轮差异很小，且 Excel 冷启动有波动，结果应视为平价。所有 9 个批次/引擎都恰好创建一个专用 Excel 实例。

## COM 实例创建与 worker 额外成本

下表是 9 个正式批次的阶段中位数，预热已排除。进程启动偏移和 Excel ready 等待由进程创建时间戳与 COM 调用时间估算；Windows COM 没有给出 Excel 内部各加载阶段的独立事件时间。

| 阶段 | Python | Rust | 说明 |
| --- | ---: | ---: | --- |
| COM 初始化 | 0.01 ms | 2 ms | Python 基准宿主此前已初始化 COM；每批仍调用 `CoInitialize`。Rust 在新 worker 线程中初始化。 |
| Excel COM 激活 | `DispatchEx` 2,524 ms | `CoCreateInstance` 2,391 ms | 两种方式都创建新 Excel 进程，没有连接已有 Excel。激活时长会随 Excel 冷启动负载变化。 |
| 进程启动偏移 | 约 5 ms | 约 7 ms | 从激活开始到观察到新 Excel 进程创建的估算值。 |
| 进程创建后到 COM ready | 约 2,503 ms | 约 2,384 ms | Excel 冷启动与自身加载是主要成本。 |
| Excel 属性配置 | 74 ms | 73 ms | `AskToUpdateLinks` 分别约 64/67 ms，是设置项中的主要耗时。 |
| Rust worker `spawn` | — | 6.23 ms | OS 创建 worker 进程的调用耗时。 |
| Rust worker ready IPC/启动残差 | — | 37.17 ms | `spawn-to-ready` 减去 session 初始化耗时的估算残差。 |
| worker spawn 到 ready | — | 2,540 ms | 其中 COM session 初始化约 2,502 ms；绝大多数时间来自 Excel。 |

本轮 `DispatchEx` 与 `CoCreateInstance` 的激活中位数有约 133 ms 差异，但 2/10/33 册结果没有固定的 Python/Rust 领先方向，冷启动样本波动大于端到端差距。没有证据支持把约 2.4–2.5 秒归咎于 Rust。激活时间包含 Excel 自身启动和加载；仅凭当前 COM 数据无法再把加载拆分为插件、安全检查或其他 Excel 组件。

`Application.Calculation = -4135` 在 Python、Rust 两边的样本中都设置失败；两边均记录失败并继续显式逐表计算。不能把该设置宣称为成功启用。实际公式结果检查通过。

## 每册阶段与端到端对应

33 册批次的逐册中位数（毫秒）：

| 阶段 | Python | Rust | Rust 计时口径 |
| --- | ---: | ---: | --- |
| 打开 | 225 | 225 | `open` |
| 逐表计算 | 10.49 | 4 | `sheets_calculate` |
| 保存 | 132 | 135.5 | `save_as` |
| 关闭工作簿 | 18.19 | 16 | `close` |
| 逐格缓存核验 | 228 | 231.97 | 父进程独立核验 |
| 请求耗时 | — | 391.78 | Rust 单次请求的打开至关闭 JSON roundtrip，不含下行缓存核验 |
| 请求剩余估算 | — | 4.77 | roundtrip 减去 worker 报告的打开/计算/保存/关闭阶段 |

Rust 每册 roundtrip 加独立核验约 624 ms；Python 四个 Excel 阶段加核验约 614 ms。10 册结果相近：Python 逐册 228/10/137/19/239 ms，Rust Excel 阶段 roundtrip 384 ms、核验 236 ms。2 册数据受较小样本及一次失败注入影响，不用来解释稳定的单册处理差距。

每个实际批次的分段和总计均单独比较，未把不同批次的阶段中位数相加冒充总耗时。18 个批次的未归因差额中位数为 3.06 ms，最大绝对值为 63.81 ms；该误差只占各批次总时长的一小部分。驱动日志/最终 JSON 序列化发生在批次计时之后，不包含在端到端计时中。

此前记录的约 3.5/3.8 秒批次只计入 COM 释放阶段，没有等待 Excel.EXE 真正退出。本轮测得 `Quit`、COM 代理释放与进程退出的中位数如下，因此新总时长更高：

| 关闭阶段 | Python | Rust |
| --- | ---: | ---: |
| Excel `Quit` 调用 | 1.67 ms | 0 ms |
| COM 代理/接口释放 | 222.89 ms | 215 ms |
| COM uninitialize | 0 ms | 1 ms |
| 等待 Excel 进程退出 | 2,737.54 ms | 2,741.33 ms |
| 驱动观察到的完整关闭 | 2,960.98 ms | 2,956.63 ms |

等待 Excel 进程退出是两边共同的约 2.7 秒成本。它不会被优化成异步遗留进程；结束批次前必须确认由该批次创建的 Excel 已退出。

## 正确性、失败隔离与进程清理

- 每个成功输出保留 6,000 个预期公式文本，并读取 7,200 个非空缓存值；6,000 个公式结果逐格与独立数值预期比较。
- 每个引擎输出 96 册，合计核验 576,000 个公式结果和 691,200 个缓存单元格。两引擎逐批配对的公式文本哈希及计算结果哈希差异为 0；共 96 对，没有输出差异。
- 两个原始合成工作簿执行前后 SHA-256 相同；批次使用其独立副本作为输入，输入副本也在批次结束复核。
- 第一个 2 册批次按“有效、损坏、有效”的顺序处理。每个引擎均记录 1 个损坏文件失败，后一个有效文件仍成功；其他 10/33 册批次全部成功。
- 每批在 ready 消息中报告一个新建 Excel PID。18 个批次内的 Python/Rust 专属 Excel 实例均退出，Rust worker 均退出，没有强制终止或 worker 孤儿。验收开始时存在的 15 个 Excel PID 仍在；输出瞬间另记录到一个非本 harness 所有的隐藏 Excel PID 12264（15:14:28 创建），随后查询时已自行退出，未被终止。之后的只读进程快照还观测到另一个短暂的非本 harness Excel PID，系统中可能有并行 Office COM 活动；不将这些 PID 归属给本批 worker，也不宣称机器上不存在其他应用启动的 Excel。
- 另外分别运行两个需要真实 Excel 的 ignored 集成测试：正常批次复用测试 1/1 通过（同一 Excel PID 重算两册、正常 Quit/退出、源文件不变）；受控 parent-timeout 清理测试 1/1 通过（35 秒超时后 worker 和其确切拥有的 Excel 退出，源文件不变且没有发布输出）。超时用例在一次成功计算后人为延迟 worker 响应，不是 Excel 内部 COM 计算永久卡死的模拟。

## 构建与回归

- 最终 Release 命令：`cargo build --offline --locked --release --manifest-path audit-core\Cargo.toml --bin open-data-audit`，`CARGO_TARGET_DIR=F:\RustBuild\audit`；成功，最终构建耗时 2 分 14 秒。批测最终二进制 SHA-256 与此 Release 一致：`9112415fb15454136905aefdd694236006178c177e6f542b0cb83e0215ac7ba4`。
- 全 workspace：82 个测试目标，**501 passed / 0 failed / 34 ignored**，Cargo 退出码 0。ignored 项依赖人工提供路径或外部 Excel/业务样本。旧基线为 501/0/33；新增的 ignored 项是受控 COM parent-timeout 用例，本轮已另行实际执行并通过。
- 两个真实 Excel 集成测试已单独用 `--ignored` 执行，各 1/1 通过。四个修改的 Rust 文件 `rustfmt --check` 通过，批测 Python 驱动 `py_compile` 通过。
- 全局 `cargo fmt --all -- --check` 仍会被仓库内其他既有文件的格式差异阻断；本轮没有格式化或修改无关文件。

## 性能结论与限制

主要启动成本是 Excel 的独立实例冷启动，Python/Rust COM 激活差异在本机多轮观测中上下波动。Rust 每批 worker 进程与 ready 协议增加约 43 ms；批内逐册请求残差约 5 ms。33 册总批次约 27 秒，这些 worker 开销已很小，不值得通过复用用户已有 Excel、取消进程隔离或缩减正确性校验来换取几毫秒。

因此，**S1-F1 Windows COM 性能优化标记完成并停止继续迭代**。本结果仅覆盖合成重算工作簿和原生 COM worker；S1-F1 真实机构输入、GUI 菜单整链路及机构审核结论仍需人工验收。本轮没有新的真实机构样本证据。

## 证据文件

- 配套逐轮机器数据：[`s1f1_com_final_acceptance_20261030.json`](s1f1_com_final_acceptance_20261030.json)
- 批测驱动：[`s1f1_final_acceptance_20261030.py`](../../tests/compare/s1f1_final_acceptance_20261030.py)
- 全 workspace 日志：`F:\RustBuild\audit\s1f1_final_workspace_regression_20261030.log`
- 原始性能 JSON：`F:\RustBuild\audit\s1f1-final-acceptance-20261030-bx3jkoyl\s1f1_final_acceptance.json`
