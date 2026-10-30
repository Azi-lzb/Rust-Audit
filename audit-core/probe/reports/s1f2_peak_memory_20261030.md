# S1-F2 进程树峰值工作集专测

## 结果

| 实现 | 峰值工作集：三轮中位数 | 三轮最大值 | 峰值进程数 |
|---|---:|---:|---:|
| Python 参考实现（Python + 独立 Excel COM） | 300.078 MB | 301.574 MB | 2 |
| Rust Release CLI | 10.719 MB | 10.727 MB | 1 |

同一轻量负载下，Python 路径峰值约为 Rust 的 28 倍，多用约 289.359 MB。主要差异来自旧 Python 入口启动 Excel COM 实例；Rust 使用 Calamine 和 Rust writer，无需启动 Excel。Rust 已约 10.7 MB，未发现值得优先继续压缩的产品内存热点。

## 负载与验证

- Windows 11 x86_64；Python 3.12.10；Rust 二进制 SHA-256：`a2ac5bc3fd16223a3c18976f9ed6042e941af6df17fc2fbd84ad5fe54b82f370`。
- 沿用 S1-F2 输出对拍驱动的同一夹具：2 本源工作簿、3 张源工作表、5 行源数据；模板 5,672 bytes，两个源簿 5,429 / 4,938 bytes；生成 2 张输出表、9 行数据。
- 预热不计入；三轮 Python/Rust 顺序交替。每次运行均启动独立 worker；Python 基线通过 `ExcelSession` 的 `DispatchEx` 创建专用 Excel 实例。采样每 5 ms 累计 worker 子树和本轮新启动的 `EXCEL.EXE` 工作集；启动前已存在的 Excel PID 从采样中排除。
- 三轮均由现有对拍函数检查输出。工作表顺序、尺寸、单元格值、数字格式、所测表头格式、筛选和冻结项均匹配；单元格值差异为 0。最大列宽差 0.535，与原 S1-F2 对拍一致，不影响业务值和结果结构。
- 模板及两份源输入 SHA-256 测量前后均一致。三轮中发现的 Excel 进程 PID 均在专测结束后退出；没有关闭预先存在的 Excel 实例。

## 限制与结论

本轮只比较峰值工作集，不记录或重报耗时；采样可能漏掉短于 5 ms 的瞬时峰值，且包含进程启动和 Excel COM 激活。样本仅 5 行业务数据，不能推断大批量 S1-F2 峰值内存随数据规模的增长曲线。Python 的高峰值主要是 COM/Excel 引擎常驻成本；在 Rust 已只有约 10.7 MB 的情况下，不改业务代码追求更低数值。

复跑入口：[内存专测驱动](../../tests/compare/perf_s1f2_peak_memory_20261030.py)，[逐轮机器数据](s1f2_peak_memory_20261030.json)，[原 S1-F2 对拍驱动](../../tests/compare/compare_s1_region_summary.py)。
