# GUI 日志规范验收

日期：2026-11-13。代码基线 main / 999c543。证据根目录：`F:/RustBuild/audit/tmp/gui-logging-20261112`。

本轮日志开发与实际 GUI 等效验收已完成。真实 129 本全部执行完成；最终回归核心 304/0/11。排除 lo-convert 的 workspace 为 525 passed / 9 failed / 24 ignored，不是全绿；剩余失败逐项保留。

## 固定候选与可复核记录

最终 Debug 候选：`candidate-final/AuditTools.exe`，SHA-256 `3c0be6b11b458b3144547a9de32c4a286ffca0b02cf354459c74062ed29e83cb`。`candidate-final.json` 保存五份 Rust 源文件哈希、基线提交和构建日志。`delivery-candidate-build.exit=0`。

实际 GUI 等效运行使用现有 serve/API/start_dag_flow 和生产 worker，不是模拟日志字符串，也不是另建 CLI 产品。每个场景保存 final-state、progress、user-log、developer.jsonl 和输入哈希；独立 server.json 保存候选 SHA、PID 和端口。样本及配置位于自有隔离目录，不覆盖真实输入。

## 真实工作流场景

`api-acceptance-final/results.json`：10/10 场景 validationErrors 为空，输入/模板哈希不变，diagnosticIncomplete=false。

- PASS 全部执行成功、无审核异常：2 本完成，0 异常，用户日志 2 条。
- PASS 执行成功、有审核异常：2 本完成、2 条公式异常，失败 0；不把审核不通过当执行失败。
- PASS 部分结构跳过：1 本完成、1 本跳过、1 条异常。
- PASS 缺少工作表：实际结构预检跳过，no_structure_passed；摘要不虚构审核完成。
- PASS 缺少必要字段：固定结构标签缺失，实际跳过；不算通过。
- PASS 公式计算成功：实际计算后提取并生成结果。
- PASS 公式计算失败：BetterOffice 实际拒绝 INFO 公式且禁止回退，2 本待人工重算，未发布结果，主要状态明确失败。不是复制失败冒充计算失败。
- PASS 内部回退后成功：BetterOffice 拒绝后实际 Excel COM 完成 2 本计算、2 条异常；用户日志 2 条不暴露引擎过程，开发日志有拒绝、worker/PID/IPC/关闭记录。
- PASS 部分文件执行失败：Windows 独占文件锁使第 2 本读取失败，1 本完成、1 本失败；日志明确文件不可读。
- PASS 同名文件定位：不同子目录同名文件，失败日志含 `机构02` 子目录。
- PASS 最终结果生成失败：自有输出文件位置被目录占位，实际发布拒绝，resultError 非空、resultPublished=false；摘要明确未生成，不能把已提取异常当已交付成功。

以上 11 个验收维度由 10 个场景交叉覆盖；计算成功也在普通成功与回退场景中验证，不额外伪算一次场景。

## S1/S2/S3 覆盖

- S1-F01：上述实际 GUI 服务场景、真实七文件和 129 本批量。
- S1-F02、S2-F01、S3-F01/F02/F04：现有 GUI API/worker 测试读取/写入真实合成工作簿，保留业务结果断言；按 AUDIT_LOGGING_TEST_CAPTURE 另存完整运行记录。最终 workspace 捕获目录为 workspace-final-captures，共 8 组记录，索引 workspace-final-captures-index.json，diagnosticIncomplete 均为 false。核心测试全部通过。
- S3-F03/F05：`api-s3-final/results.json` 两项 validationErrors 为空。F03 实际原地写入 1 行、未匹配/冲突 0，备份哈希与更新前目标相同，G2 实际更新；F05 1 机构、1 匹配指标、1 输出工作簿。非目标输入配置哈希不变。

## 真实样本及结果不变性

`after-real-seven-final/capture-summary.json`：7/7 完成；317 公式 + 5 条件格式 = 322 条异常；失败、跳过、部分和待人工重算均为 0；用户日志 2 条，输入/模板/配置/外部文件哈希不变。

`real-seven-output-parity-final.json`：修改前后结果全单元格值/公式、Sheet 名称与顺序、合并区域、定位链接（只剔除运行目录变化）一致，differences=[]。修改前普通 GUI 日志 58 条，修改后 2 条。

`statistics-output-reconciliation.json`：六个已发布结果的合成场景及真实七文件，摘要异常数与实际结果表数据行数全部一致。生产统计来自原业务报告；结果表读取仅测试时进行。

真实七文件耗时：摘要 151.3 秒，驱动单调观察 151.36 秒；729 次 API 轮询，P95 约 31ms、最大约 47ms。这是 Debug/并行编译环境的功能验收，不与旧 Release 进行速度倍率比较，也不等于浏览器 FPS 测试。

129 本批量：PASS，证据 real-129-final/capture-summary.json。129 本全部执行完成，3352 条公式异常 + 257 条条件格式异常 = 3609；失败、跳过、部分及待人工重算均为 0，结果发布成功，输入/模板/配置/外部文件哈希未变。用户日志 2 条，开发诊断 1340 条，运行 ID 一致、含最终摘要、diagnosticIncomplete=false。

statistics-bulk-output-reconciliation.json 确认摘要 3609 与实际输出数据行 3609 相等。最终摘要总耗时 27分13.7秒；驱动单调观察 1633.859秒，7814 次 API 轮询 P95 约 31ms、最大约 203ms。主要状态时间在摘要之前取值 1632.8秒，两者取时点不同，不通过拼阶段耗时凑数。这是实际长时间与批量日志验收，不是优化性能跑分；129 本没有另跑修改前整批，不能宣称其全部业务输出已经做过前后逐格对拍。前后逐格不变性证据来自上面的真实七本。

## 构建与自动化

- 修改前全量 `cargo test --workspace --offline --locked`：退出 101，在 lo-convert 编译阶段找不到 crate::conditional_format。证据 baseline-tests.log。
- 修改后同命令：同样退出 101，证据 final-workspace-tests.log/exit。未删除失败测试或修改无关模块来使其变绿。
- 中间核心测试：304 passed / 0 failed / 11 ignored，证据 core-tests-final.log、core-tests-latest-direct.log。这些早于最后小范围错误呈现修正，不冒充最终源码全绿。
- 最终源码 `cargo test --workspace --exclude lo-convert --offline --locked --no-fail-fast`，CARGO_BUILD_JOBS=1：退出 101；525 passed / 9 failed / 24 ignored（70 个测试目标，含文档测试）；其中 audit-core lib 为 304 passed / 0 failed / 11 ignored。证据 workspace-final-tests.log/exit、workspace-final-test-summary.json。
- 四份 Python 验收助手 py_compile：PASS；git diff --check：PASS。
- 独立诊断单元测试覆盖 Unicode/顺序、I/O 失败可见且后续运行继续、分段及只清理自有日志、保留数量；GUI 测试覆盖正常 detail 不进用户缓冲、启动失败计数、复制/读取失败分类及最终摘要。

早期驱动 r1/r2 无效等待、r3/r4 复制阶段阻断不算计算/发布验收。资源不足时取消的编译及构建锁等待没有执行用例，不计测试通过/失败。全部 target 缓存保留。

## 剩余测试失败：明确未修复

本轮未删除、改写或忽略这些失败。下列涉及的 CLI 主入口、业务实现、测试文件及 Excel 后端本轮均未修改；完整基线被 lo-convert 编译阻塞，所以不能把全部运行失败武断地宣称为“已证实原有”。本轮核心日志测试和实际 GUI 结果对拍通过，不等于全仓库验收通过。

- production_capability_verification（1）：Umya 复制批注后文本变化，被保护校验拒绝。
- s1_audit_directory_cli（2）：旧入口预期 4 条实得 3 条、源文件公式错误计数预期 1 实得 0。
- s1_conditional_audit_batch_cli（1）：旧技术事件数量断言不满足。
- s1_conditional_audit_directory_cli（1）：旧正常预检日志文本断言不满足。
- s1_scan_scope_contract（1）：源码字符串要求旧入口必须 plan.scan(&input.source_file)，当前入口已无此表达式。该源码在基线亦无该表达式；没有为满足断言倒退当前口径。
- s2_analysis_parity（1）：Python/Rust 规则与质量顺序对拍差异（8 与 9 条发现，含缺指标/外部提示规则）。
- s2_service_parity（1）：invalid_config 场景预期 1 条 CLI 事件而实际 0 条。
- s3_explanation_export_cli（1）：预期“说明文件、生成信息”两张 Sheet，实际仅“说明文件”。

以上共 9 个用例、8 个目标失败。需要另行确定旧测试口径或业务缺口，不在日志任务中修改审核结果、输出表或引擎校验。

## NOT_TESTED 与已知边界

- NOT_TESTED：Linux/麒麟/UOS、离线 deb、Release 打包，本轮未请求交付这些产物。
- NOT_TESTED：浏览器人工逐按钮点击、滚动手感及 FPS。界面 HTML 本轮没有修改；真实 API 捕获证明用户日志内容和状态，不能代替人工体验。
- NOT_TESTED：真实 OS 线程资源耗尽；启动失败仅用单元测试验证，不破坏系统制造故障。
- NOT_TESTED：强制磁盘耗尽及突发塞满队列的生产负载；I/O 失败/分段/保留由单元测试验证。队列或磁盘失败会标记诊断不完整，不保证异常环境零丢失。
- 原错误没有捕获 backtrace 时保存 null；保留完整错误链，不保证第三方引擎内部堆栈。没有伪造错误起点或 DAG 节点。
- 既有数字实体公式复制与部分设置 update=501 问题见专项审计，不属于日志回归，未在本任务更改业务处理。

## 交付与工作区

本轮变更为五份 Rust 源文件、四份文档、四份验收助手及一份修改记录。index.html 与已有 LO 菜单修改记录/性能报告/probe 材料不纳入本轮 commit。只做本地提交，不 push、不合并；提交前已复核五份源码与实测候选哈希一致，index.html 的初始 Git blob 561ebb41a287a924ae8b0881d97f86d2d4abe597 保持不变；证据 final-delivery-source-verification.json。提交不包含测试数据、诊断正文和其他会话资料。
