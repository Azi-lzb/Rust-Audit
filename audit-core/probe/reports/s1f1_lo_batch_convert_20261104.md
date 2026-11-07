# S1-F1 Linux LibreOffice 批次复用验证

日期：2026-11-04  
环境：WSL Ubuntu，LibreOffice 24.2.7.2  
范围：S1-F1 公式审核副本的真实公式重算子步骤；不是完整 GUI 审核链，也不是 Python/Rust 对比。

## 方法

使用 rust_xlsxwriter 生成 2、10、33 本独立 XLSX，每本含 400 行 `=Brow*2+Arow` 公式及一条求和公式。逐本路径对每个文件单独调用 `lo-convert recalculate --engine lo`；批次路径把同一组文件交给一次 `lo-convert recalculate-batch --engine lo`。两种路径都使用相同的源文件和输出契约。

验证每本输出仍含 401 条公式、公式缓存值与独立期望值一致；对比输入文件字节确保源工作簿未变；批次 JSON 检查 `soffice_conversion_invocations == 1` 和 `profile_processes_remaining == 0`。

## 结果

| 工作簿数 | 逐本冷调用总耗时 | 单批总耗时 | 加速比 | 批次平均/本 | `soffice` 调用数 | 缓存 / 输入 / 清理 |
|---:|---:|---:|---:|---:|---:|---|
| 2 | 5,933 ms | 4,987 ms | 1.19× | 2,494 ms | 1 | 通过 / 不变 / 无残留 |
| 10 | 32,813 ms | 10,575 ms | 3.10× | 1,058 ms | 1 | 通过 / 不变 / 无残留 |
| 33 | 80,381 ms | 21,837 ms | 3.68× | 662 ms | 1 | 通过 / 不变 / 无残留 |

每个性能样本仅测量一轮，不能当作稳态中位数。批次在 2 本时已较快，但该档差值较小；10/33 本更明显体现为整批只支付一次 LibreOffice 进程与 profile 初始化成本。

`soffice --convert-to` 未提供逐工作簿阶段计时，因此本报告只记录逐本路径整批总时间和批次总时间；没有把批次时长除以文件数当成实测的单本时长。GUI 既有每文件报告字段如需显示批次估算，只能标明是均摊值，不能用于单本耗时归因。

## 实现结论与边界

- Linux S1-F1 GUI 在重算前准备可处理文件的公式副本，再对整批调用一次 `recalculate-batch`；LibreOffice 按序处理这些副本。
- 该实现是单次 `soffice --convert-to` 多输入批处理，拥有隔离 profile 并在结束时清理；它不是持久 UNO 文档会话，也不跨用户批次常驻。
- 单本准备错误不会阻止其他独立工作簿进入批次。批次或单本重算/验证失败时保留无缓存公式审核副本供 S1-F3 人工重算，不逐本冷启动重试。
- 不能据此推断真实 S1 工作簿的公式兼容性、条件格式审核结果、GUI 全链路耗时或 Kylin V10/UOS V20 性能。真实业务输入、Linux 桌面 GUI 和目标发行版仍需人工验收。

## 定向验证

- WSL Linux：`audit-core --lib s1_f1` 4/4；`lo-convert --test recalculate_libreoffice` 2/2；忽略的性能探针手动运行 1/1。
- Windows：`audit-core --lib s1_formula_copy` 17/17；`lo-convert --test recalculate_libreoffice --no-run` 编译通过。
- 未运行全 workspace 回归。`rustfmt --check` 对修改过的宽文件仍报告既有风格差异；没有为通过全文件格式检查而重排无关代码。

机器数据：[`s1f1_lo_batch_convert_20261104.json`](s1f1_lo_batch_convert_20261104.json)。复跑入口：`sidecar/lo-convert/tests/recalculate_libreoffice.rs::manual_s1_formula_recalculation_batch_performance_2_10_33`（测试被标记为 ignored，需明确传入 `--ignored`）。
