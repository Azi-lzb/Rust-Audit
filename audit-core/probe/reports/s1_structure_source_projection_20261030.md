# S1-F1 源结构预检投影路径同输入性能对照

## 结果

| 读取路径 | 中位数 |
|---|---:|
| 完整 `WorkbookDump` 读取 + 既有结构比对 | 501.532 ms |
| Calamine 结构投影读取 + 同一结构比对 | 237.579 ms |
| 子步骤加速 | 2.111×（减少 52.6%，约 263.953 ms） |

这是 S1-F1 源文件结构预检子步骤的 Rust 路径对照，不是 Python/Rust 对照，也不代表完整审核功能的总耗时。

## 负载与测量方法

- Windows x86_64，Release 集成测试；源码基于 HEAD `7517a9c`，运行时未修改输入。
- 20,000 行数据 + 1 行表头，1 张 `S1Conditional` 工作表，8 列、160,008 个非空格，4 条条件格式规则。
- 源工作簿和模板工作簿是逐字节相同的同一生成文件，各 716,662 bytes；SHA-256 均为 `5faab2e54f24a465da130fcd5091be8849abc39c493f7b9012c48077e151b23d`。
- 预热 1 轮、计时 5 轮；两种路径每轮交替先后次序，在同一 Rust 测试进程中使用同一源文件与同一模板定义。
- 基线执行 `excel_reader::read_workbook`，再调用原有 `compare_s1_source_structure_dump`；候选执行当前 `compare_s1_source_structure` 文件入口，使用 Calamine 结构投影。
- 每轮核对匹配状态、分数、命中/检查标签数、缺失工作表及诊断完全一致。5/5 轮均匹配 8/8 个标签，输入和模板 SHA-256 测量前后不变。

完整读取路径的中位数 501.532 ms 中，读表为 501.458 ms，纯结构比对为 0.054 ms；因此收益来自跳过完整读取构造和无关元数据解析，不来自改写结构匹配规则。投影读取仍会由 Calamine 遍历工作簿值，不能视为按 XML 物理随机访问。

## 验证及边界

- 驱动运行的定向命令：`cargo test --locked --release -p audit-core --test s1_structure_projection_source_perf profile_source_structure_projection_same_input -- --ignored --nocapture --test-threads=1`。
- 该命令成功编译并运行单个忽略式性能测试；每轮业务结构判断完全一致。
- 本次未运行 Python 参考路径、未采集独立峰值 RSS、未运行全工作区回归或真实机构文件。以前的模板读取投影结果仍单独记录为 528→224 ms；本报告测的是源文件侧路径，不能将两者相加为整项功能收益。
- 本轮只增加可复跑基准及记录，没有修改 S1 生产业务代码。Git 工作区包含此前已有的多项未提交改动，本轮予以保留；没有 commit 或 push。

复跑入口：[驱动](../../tests/compare/perf_s1_structure_source_projection_20261030.py)、[原始逐轮 JSON](s1_structure_source_projection_20261030.json)、[Rust 性能测试](../../tests/s1_structure_projection_source_perf.rs)。
