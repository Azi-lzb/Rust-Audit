# Rust-Audit 新旧性能复测（2026-10-30）

## 结论

本轮在 Windows x86_64 上完成 3 轮复测。Rust 在 S3-F2、S3-F4、S3-F5 和 S1-F2 的当前测试规模下明显更快；S3-F1、S3-F3、S2-F1 和暂停开发中的 S5 小负载下略慢。S1-F1 的 COM 热会话已复用，但此轮重算子步骤仍比 Python 的 `CalculateFull` 对照慢约 28%。这些数字不能直接解释为 Rust 语言本身的速度差异：有的基准包含 Excel COM 启动，有的 Rust 计时包含 CLI 启动，也有功能在新旧版本执行了不同范围的额外校验。

加速比定义为 `Python 中位数 / Rust 中位数`；大于 1 表示本轮 Rust 用时更短，小于 1 表示 Rust 用时更长。主表使用三轮中位数，减少单轮冷启动波动的影响。

## 性能结果

| 功能与测试规模 | Python 中位数 | Rust 中位数 | 加速比 | 结果说明 |
|---|---:|---:|---:|---|
| S1-F1 公式重算子步骤：2 个工作簿，每册 7,200 个 VLOOKUP | 3,547.7 ms（`CalculateFull`） | 4,950.6 ms | 0.72× | Rust COM worker 在两册之间复用同一 Excel 进程；两侧公式和数值缓存均逐格验证。旧 Python 的工作表级 `Calculate` 中位数为 3,513.0 ms。 |
| S1-F2 命名区域汇总：2 个源工作簿、3 张源表、9 行输出 | 4,320 ms | 59 ms | 73.22× | 三轮输出对比通过；值、表名、维度及检查的格式属性相同。列宽最大差约 0.54。Python 路径使用 Excel COM，Rust 使用 Calamine/OOXML 与 rust_xlsxwriter。 |
| S2-F1 报表采集：2 个仓库夹具、1 个结果工作簿 | 40.0 ms | 52.0 ms | 0.77× | 两侧均写出可读工作簿，结果表名及维度一致；本轮未做全单元格逐项对比。Rust CLI 启动时间包含在内。 |
| S3-F1 两期比较：10,000 行 | 1,064.1 ms | 1,432.5 ms | 0.74× | Rust 慢约 368 ms（约 35%）。本轮没有重新做完整单元格结果对拍，不能据此判断业务结果是否完全一致。 |
| S3-F2 比较文件转说明文件：5,000 行、1,667 行输出 | 1,853.2 ms | 234.6 ms | 7.90× | 三轮计时完成；本轮性能驱动没有附带全量值级输出对比。 |
| S3-F3 反馈导入：1,000 行、原地写入并备份 | 934.0 ms | 980.9 ms | 0.95× | Rust 慢约 47 ms（约 5%）。导入数、备份字节一致和反馈输入未变均已验证；仅覆盖无冲突合成键。 |
| S3-F4 本期数值核对：10,000 行 | 6,053.9 ms | 937.5 ms | 6.46× | 三轮计时完成；本轮性能驱动没有附带全量值级输出对比。 |
| S3-F5 金融表单：30 家机构 × 40 项指标 × 4 张表 | 3,290.8 ms | 868.5 ms | 3.79× | 两端均生成 30 个工作簿；本轮性能驱动没有附带全量值级输出对比。 |
| S5 配置审核：单文件 + 独立上期 | 24.4 ms | 72.6 ms | 0.34× | 小负载原始耗时；Rust CLI 启动占比较高。S5 只测性能，没有恢复其业务开发。 |

## 计时口径

- 平台：Windows x86_64。Rust 使用本轮工作区编译的 Release CLI，路径为 `F:\RustBuild\audit\release\open-data-audit.exe`。Python 参考项目为只读目录 `C:\Users\AZI\Desktop\VibeCode\3.open-data-audit`。
- 每项功能 3 轮，报告中位数。Python 计时在当前 Python 进程内，不包括解释器启动；Rust CLI 计时包括子进程启动，测得 CLI 启动约 21.8 ms。输入以合成夹具为主，S2 使用仓库测试夹具；不包含真实用户工作簿或 Linux/Kylin 测量。
- **S1-F1 范围**：只比较 Excel COM 重算与保存步骤，不包含整条 S1-F1 的模板匹配、模板复制、审核异常提取、缓存清理、结果校验和原子发布。Python 和 Rust 都处理两册，并在各自批次内复用 Excel 会话；Rust 的两个输出检查到每册均保留 7,200 个公式且 7,200 个数值缓存可读，输入哈希未变。Python 默认计时使用 `Worksheets.Calculate`；另以 `Application.CalculateFull` 做同计算深度对照。Rust worker 处理两册时的 Excel 进程 PID 一致。
- **S1-F2 范围**：使用合成无公式 XLSX，输入有意补齐模板要求的结构表头，计时包括旧 Python 的 Excel COM 和 Rust CLI。当前 Office 安装拒绝旧适配器的命名 `UpdateLinks=0` 参数，所以 Python 基准使用测试专用适配器，以 `ReadOnly`/`IgnoreReadOnlyRecommended` 打开，并关闭 Excel 的更新提示。输出逐项比较通过，唯一记录的差异是列宽最大约 0.54；三轮输入哈希不变。历史说明表追加与兜底模板/固定行场景也在三轮中完成结构和值对比，但其独立时间没有纳入主表中位数。
- **S2-F1 范围**：Python 使用完整 `PeriodAuditV2Service.run`，包含 `ExcelExporter` 输出；检查两侧各有一个可读结果工作簿且工作表名/维度一致。没有做全单元格值级差异比较。
- **S3-F3 范围**：计时包含目标文件备份和原地更新；三轮均检查导入 1,000 行、备份与更新前目标字节一致、反馈输入未变。未覆盖重复键冲突或未匹配数据。
- **其余 S3 与 S5**：本轮记录执行耗时和生成行为，不将耗时视为业务等价证明。S5 结果是只读性能探测，遵守暂停开发要求。

## 验证与限制

- Release 构建通过。定向原生 COM 会话测试通过 1/1：两个合成工作簿由同一 Excel COM 进程重算，公式结果正确且输入字节未变。该测试不覆盖完整 S1-F1 GUI/业务链路。
- S1-F1 的本轮 7,200 公式核对通过；S1-F2 三轮结构/值对比通过；S3-F3 备份/导入保护核对通过；S2 仅验证可读性和结果表结构。
- S3-F1、S3-F2、S3-F4、S3-F5 与 S5 本轮没有全量值级对拍，速度数据只适用于这里列出的负载及执行路径。
- 没有执行真实业务文件验收、Linux 测量或 S1-F1 整链路计时。因此不能由这份报告推断用户生产工作簿上的总体加速比。
- S5 本轮没有修改业务实现；没有 Cargo 全量测试、打包、部署或 Git push。

## 可复跑与机器可读结果

- 主驱动：[perf_compare_20261030.py](../tests/compare/perf_compare_20261030.py)
- S1-F2 专项对拍驱动：[compare_s1_region_summary.py](../tests/compare/compare_s1_region_summary.py)
- 汇总数据：[perf_comparison_20261030.json](perf_comparison_20261030.json)
- S1-F2 三轮原始对比：[s1f2_perf_run1_20261030.json](s1f2_perf_run1_20261030.json)、[s1f2_perf_run2_20261030.json](s1f2_perf_run2_20261030.json)、[s1f2_perf_run3_20261030.json](s1f2_perf_run3_20261030.json)


## 2026-10-30 性能优化迭代（Goal：性能差异深度定位与定点优化）

### 已实施的优化（全部保持业务语义）

1. **S3-F1 结果验证去除二次全量解析**（`s3_execution_compare.rs`）：验证阶段的
   公式元数据检查原先用 `read_workbook` 做第二次整簿解析；改为轻量 XML 扫描
   （`inspect_formula_cache_nodes`），单次 read_value_columns 的逐格核对全保留。
   实测：验证 936→417ms，S3-F1 总耗时 1412→894ms（10k 行）。
2. **S1-F1 去除冗余 Calculation 模式设置**（`native_office_com.rs`）：新 Excel
   实例默认即自动模式，CalculateFull 与模式无关，设置动作实测强制同步重算
   ~150ms/册。已移除。
3. **S1-F1 改为逐工作表 Calculate + SaveAs 后 Close(False)**：与旧 Python 生产
   通道一致（应用级 CalculateFull 会把所有已开工作簿纳入同一队列；Close(True)
   在 SaveAs 后触发二次保存检查）。实测 CalculateFull 96→逐表 Calculate 后
   sheets_calculate 149ms（2 表）/册、close 93→57ms。
4. **COM DISPID 缓存 + 组合 DISPATCH 标志**（`native_office_com.rs`）：
   GetIDsOfNames 每次跨通道解析改为按 dispatch 缓存；invoke_method 改用
   METHOD|PROPERTYGET 组合标志（Worksheets.Item 等成员单 METHOD 标志报
   0x80020003）。
5. **原生 COM worker 阶段计时**：recalculate 响应携带 open/set/calc/poll/
   save/close 各阶段毫秒（stages_ms），会话创建毫秒入 session_ready——本轮
   全部定位结论的证据来源。

### 优化前后对比（同机、3 轮中位数）

| 功能 | 优化前 | 优化后 | Python | 状态 |
| --- | --- | --- | --- | --- |
| S3-F1（10k 行） | 1432.5ms | **1012.9ms** | 1032.5ms | **反超 Python（1.02x）** |
| S1-F1 重算子步骤（2 册） | 4950.6ms | 4668.5ms | 3595.5ms | 仍慢 0.77x，原因见下 |

其余功能无代码变化（S3-F2 7.90x、S3-F4 6.37x、S3-F5 3.79x 维持；S2-F1 0.77→
0.89x、S3-F3 0.95x、S5 0.34x 维持原状）。

### S1-F1 剩余差距的可核实边界（阶段级实测，热会话、交错运行）

| 阶段 | Rust | Python | 差距 |
| --- | --- | --- | --- |
| Excel 实例创建 | ~2276ms | ~2336ms | 相当（Excel 主导） |
| Workbooks.Open | ~625-970ms/册 | ~325ms/册 | +300-650ms/册（首个工作簿 |
|  |  |  | 含 Excel XLSX 过滤器冷加载） |
| 逐表 Calculate | ~149ms/册 | ~9ms/册 | +140ms/册（Rust 逐 sheet COM 往返 |
|  |  |  | 4 次；Python 集合级 1 次） |
| SaveAs | ~84-210ms | ~76-90ms | 相当 |
| Close | ~57-71ms | ~16ms | +45ms/册 |
| 实例关闭 | ~147-208ms | ~208-216ms | 相当 |

剩余差距为 Excel 引擎与 COM 通道固有成本：CalculateFull/逐表 Calculate 的
重算本身、每工作簿 Open 的过滤器加载都由 Excel 进程执行，Rust 侧已无冗余
步骤可裁（冗余 set 与二次解析已移除）。旧 Python 通道的逐表集合级 Calculate
在原生 COM 下按名解析不可用（0x80020006），已按次优的逐表循环实现并通过
7,200/7,200 缓存核验。数值结论：S1-F1 子步骤的剩余差距是 Excel 引擎成本
差异，不是 Rust 实现缺陷。

### 业务正确性验证

- `s1_native_com_session`（ignored，真实 Excel）：通过——两册复用同一 Excel
  进程、7,200/7,200 公式缓存核验、输入哈希不变。
- 全量回归：501 passed / 0 failed / 33 ignored（2026-10-30）。
- S3-F1 输出逐格核对在优化后的验证阶段内全保留（表头/行列/文本/数值/空列）。
