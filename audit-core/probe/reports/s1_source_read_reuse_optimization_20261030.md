# S1-F1 结构预检与条件审核共享源读取

日期：2026-10-30。范围限定为 S1-F1 的源文件结构预检和条件格式审核，不代表整个 S1-F1 业务链路性能。

## 瓶颈与修改

启用模板结构区域时，原生产路径先通过 `compare_s1_source_structure` 读取源工作簿，稍后条件格式审核又读取同一本源工作簿。两项源侧检查之间执行的副本复制和公式重算不需要改变原始源，因此结构匹配与条件格式规则求值可以共享单次只读投影。

现在由 `read_workbook_with_diagnostics` 读取一次；`compare_s1_source_structure_dump` 和 `extract_conditional_issues_from_dump` 分别使用同一 `WorkbookDump`。每本源工作簿的投影仅存在于该并行 worker 中，完成两项检查后释放，早于后续公式副本复制/重算。没有在批次间缓存，也没有将投影跨公式重算阶段保留。单文件 `conditional-audit`、S1-F1 目录审核及条件审核入口接入该路径；原有按文件路径调用的 API 保留。

开发者进度事件把这次读取时间记在结构预检，条件格式摘要标记 `source_read_reused=true` 并单独报告共享读取耗时，条件格式自己的 `read_time_ms` 为 0。新增测试比较普通读取与投影复用产生的异常、触发规则、规则诊断、读取诊断、未解析指标和批注状态。

## 复测口径

- Windows x86_64，Python 3.12.10，Rust Release CLI。Python 计模板读取、预检及条件审核服务内部时间；Rust 使用 `conditional-audit` 开始到完成的 developer event `elapsed_ms`，涵盖模板读取、共享源读取、结构预检及条件审核，不含进程启动。
- 合成源/模板各一份、单工作表、4,000 行数据、8 列、32,008 个非空单元格、4 条条件格式规则、16,000 个扫描格、8,610 条审核异常。
- 结构区域为 `S1Conditional!A1:H1`；每个测量轮 Python 和 Rust 共用同一对文件。1 次预热，交错执行 5 轮。
- Rust CLI SHA-256：`78d0d25b238ff6190a6e42adef6c6e60a209526802964c335560dd25f0eed230`。
- 优化前状态为指标定位优化已完成、结构和条件审核仍分别读取源工作簿的五轮结果；优化后重建相同逻辑负载。重新生成的 XLSX 容器大小为 144,030 与 144,031 bytes，前后哈希不同；此报告只声明两者由同一生成器生成、值和结构一致，不声称字节级完全相同。

## 结果

| 指标 | 优化前：两次源读取 | 优化后：共享一次读取 |
|---|---:|---:|
| Python 模板读取 + 结构预检 + 条件审核中位数 | 19,264.677 ms | 19,222.839 ms |
| Rust 结构预检 + 条件审核内部事件耗时（不含进程启动） | 498 ms | 390 ms |
| Python/Rust 内部计时参考倍率 | 38.68× | 49.29× |
| Rust 源工作簿解析 | 约 99 ms + 97 ms | 93 ms 一次 |
| 峰值工作集 Python/Rust | 77.402/86.562 MB | 77.312/85.949 MB |

Rust 内部事件耗时减少 108 ms，约 21.7%。Python 是只读对照，本轮没有改 Python；其 19,264.677 与 19,222.839 ms 差值视为运行波动，不作为收益。两侧计时边界略有差异（Rust 含完整模板读取而 Python 也独立计入模板读取），倍率仅适用于这个专项对照。报告不把 developer event 耗时称为进程墙钟；本轮没有保存 Rust 父进程计时的中位数。重用读取后工作集没有高于先前状态。

## 正确性和输入保护

五轮都得到 8,610 条 Python/Rust 异常；使用既有规范化方法计算的 SHA-256 多重集完全一致，缺失/多余均为 0。结构匹配为 8/8 标签。源文件与模板各自运行前后 SHA-256 相同。额外单轮烟测通过相同异常对拍，Rust 峰值工作集 85.934 MB。

本次没有覆盖公式副本复制/COM 重算/公式异常提取/最终结果写入，没有真实机构工作簿、GUI 全链路或 Linux/Kylin/UOS 验收；S1-F1 完整端到端性能仍未知。批量工作簿间不会同时持有投影，但多文件高并发峰值内存仍需在真实链路中确认。

## 验证

- `cargo test -p audit-core --lib s1_conditional_audit::tests -- --nocapture`：6 passed。
- `cargo test -p audit-core --lib s1_structure_compare::tests`：7 passed。
- `cargo test -p audit-core --lib conditional_format::tests`：10 passed。
- `cargo test -p audit-core --bin open-data-audit --no-run`：CLI 测试目标编译通过。
- Release `open-data-audit` 构建成功；A/B 使用的二进制 SHA-256 如上。
- 五轮性能驱动与一轮内存烟测均通过；全 workspace 回归未运行。
- 无 Git commit、无 push。

复跑入口及逐轮机器数据：[驱动](../../tests/compare/perf_s1_preflight_reads_20261030.py)、[五轮 JSON](s1_preflight_reads_20261030_shared_read_after.json)。
