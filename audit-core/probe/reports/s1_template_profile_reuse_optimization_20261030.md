# S1-F1/F2 模板推荐元数据复用

日期：2026-10-30。范围：S1-F1 自动模板匹配、S1 多文件模板识别及 S1-F2 多文件自动模板识别中的模板 Sheet 目录推荐子步骤。此报告不表示完整功能端到端耗时。

## 瓶颈与修改

`recommend_template_for_workbooks` 每次会扫描模板目录并重新读取每个 `.xlsx` 模板的 OOXML `workbook.xml` Sheet 目录。S1-F1 worker、S1 源文件识别界面及 S1-F2 多文件模板识别都在每个源簿循环内调用该函数，因而在批次里重复加载同一组候选模板。

新增 `load_template_profiles` 与 `recommend_template_with_profiles`。多文件入口在操作开始时读取一次模板档案，然后逐本读取源簿 Sheet 目录并执行原来的纯匹配算法。S1-F1 同批次的完整模板预检 `Arc` 缓存继续独立运行；模板档案只在一次操作内复用，下一批次重新读取。单文件 API 保留原签名，匹配阈值、候选排序、置信度和说明生成逻辑未变。

## 对照结果

Release profile；同一组合成输入预热 1 轮、5 轮中位数：

| 指标 | 每本源簿重复加载模板 | 批次只加载一次 |
|---|---:|---:|
| 源簿 | 21 本，每本 8 张表 | 相同输入 |
| 模板簿 | 5 本，每本 9 张表 | 相同输入 |
| 输入总大小 | 源簿 183,519 bytes；模板 46,312 bytes | 相同输入 |
| 非空数据格 | 0（只测试 Sheet 元数据匹配） | 相同输入 |
| 模板推荐子步骤中位数 | 30.2814 ms | 5.9081 ms |
| 节省 | — | 24.3733 ms，约 80.5% / 5.13× |

21 个源簿的完整 `TemplateSuggestion`（模板路径、置信度、匹配说明和候选列表）逐项完全一致。旧路径每个源簿重读 5 个模板，共 105 次模板目录读取；复用路径读取 5 次，省去 100 次重复读取。两端都各自读取 21 个源簿一次。

绝对节省约 23 ms，适合保留为简单的批次 I/O 复用，但不应据此宣称 S1-F1 或 S1-F2 整体快 5×。样本没有单元格值、业务输出或完整 GUI/S1-F2 导出，且没有测量峰值内存。

## 接入与验证

- S1-F1 自动匹配 worker：一次加载候选模板档案；手动选择模板路径不额外读候选列表。
- S1 源文件识别入口：对整批文件共享候选模板档案。
- S1-F2 多选文件自动匹配：对所选文件共享候选模板档案。
- 单文件推荐 API：原签名保留，并委托到同一档案加载和纯推荐逻辑。
- `cargo test --locked -p audit-core --test s1_template_profile_reuse -- --nocapture`：1 passed；Release-profile 复跑 1 passed，21/21 推荐相同。
- S1-F1 三文件 GUI 合成路径（Release，忽略标记测试显式启用）：1 passed；3/3 通过模板/结构预检、输入哈希不变。此主机的 COM 重算 worker 未正常完成，LibreOffice 不可用；测试验证的是产品的人工重算降级，没有验证完整公式重算/结果输出。
- `cargo build --locked --release -p audit-core --bin open-data-audit`：成功；SHA-256 `7be5fdece2ae26e931feb796daa7075395ab3fa44591cc242ba0813beba96792`。
- `py_compile` 和本次直接编辑文件 `rustfmt --check` 通过。全仓 `cargo fmt --all --check` 因既有大量格式差异未通过；未格式化无关文件。未运行全 workspace 回归、Linux/Kylin/UOS 测试。

机器报告：`s1_template_profile_reuse_20261030.json`；测试夹具：`tests/s1_template_profile_reuse.rs`；复跑驱动：`tests/compare/perf_s1_template_profile_reuse_20261030.py`。源码基于 HEAD `7517a9c`，工作区存在未提交 WIP，本轮未提交、未 push。
