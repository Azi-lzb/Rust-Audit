# GUI 日志实现专项核对

日期：2026-11-13。仓库 Rust-Audit；分支 main；修改前 HEAD 999c543。

## 工作区保护

初始已有 index.html 修改（五连击隐藏 LO 安装入口、移除未实现求值选项），对应 20261112_192503 修改记录、s1f1_performance_20261112.md 以及 probe 实测产物均保留，不纳入日志任务提交。初始差异存于 F:/RustBuild/audit/tmp/gui-logging-20261112/preexisting.patch。没有 checkout/reset、创建分支或 worktree；未改旧 Python 或 Rust-Tools。

## 修改前问题与本轮处理

1. 用户和技术 detail 共用 GUI 环形缓冲；缺少独立持久化完整诊断。现改为先入诊断，再按用户/技术分流，detail 不进入 GUI 缓冲。
2. currentWork 已展示文件和阶段，日志仍重复预检/复制/重算/提取、引擎、成功文件、输出目录及耗时。正常阶段改入诊断，保留开始、必要异常和一次最终摘要。
3. S1 文件执行完成可能有审核异常，结果发布也可能独立失败。摘要分别报告文件状态、异常条目及结果是否发布，未生成结果时主状态明确失败；不改既有业务返回/退出状态。
4. S2 指标记录、发现条目、审核异常不是相同全集；S3 导出、反馈导入、规则比较与拆表也不能统一成通过数。按实际报告分别写 runStatistics，不从文本反向统计或补读工作簿。
5. 后台线程启动失败缺少完整摘要。八个正式功能均补充单调耗时、taskStarted=false、executionFailed=true 和 resultPublished=false，不虚构失败文件数。
6. 错误技术链直接给用户且同名文件难区分。面向用户简化已识别技术失败，保留相对路径和实际已有 Sheet 上下文，完整链与 currentWork 入诊断。

## 代码范围

- src/web_bridge.rs：统一日志分流、运行身份、真实报告统计、摘要和错误呈现；测试按持久化技术日志检查，不削弱原业务结果断言。
- src/gui_diagnostics.rs（新增）、src/lib.rs：有界异步 JSONL 诊断、分段保留、失败标记及单元测试。
- src/bo_calc.rs、src/s1_recalculation.rs：仅在既有执行节点添加计算轮次/路由/会话/IPC 计时事件；不改公式算法、引擎路由或执行顺序。
- tests/compare/gui_logging_*_20261112.py：真实 GUI API 验收、只读批量和输出对拍助手。openpyxl 仅用于本地测试夹具与结果检查，不增加产品依赖。
- 四份 AUDITTOOLS_LOGGING 文档及本轮 GPT 修改记录。

未修改 GUI HTML、布局、筛选、复制/保存/清空、运行流程、审核规则、Excel 读写语义、结果文件格式或 DAG。

## 诊断边界

事件可按运行 ID、功能和 journalSequence 关联；异步 I/O 不占 GUI 状态锁，普通/重要队列均有界。失败记录标记 incomplete，不传播为业务失败。正常日志和实际 API 响应延迟已用真实运行捕获；未以此冒称浏览器 FPS 测试。

原错误有 anyhow backtrace 时保留，否则明确 null；并非每个第三方引擎都提供堆栈。没有增加全局 panic hook、强制回溯或修改业务错误传播。不能宣称第三方内部诊断无条件完整，也不能在磁盘/队列失败后声称零丢失。

## 本轮发现但未改的既有问题

- 全 workspace 编译被 sidecar/lo-convert 中缺少 crate::conditional_format 阻塞，修改前同样失败。不借日志任务重构或修复该独立模块。
- 现有 OOXML 导入器的 unescape_xml 不解码中文数字实体；部分 openpyxl 合成公式被二次转义后正确遭复制校验拒绝，旧 dist 同样复现。验收夹具仅将等价中文实体改为 Excel 使用的 UTF-8 字面编码，不弱化校验。此缺口需要后续独立任务。
- 设置 update 接口对 boCalcEngine/boCalcFallback 返回 501。验收只在自有隔离 AUDIT_STATE_DIR、空闲期间设置配置后调用真实 GUI worker，不修改该非日志设置逻辑。

## 证据使用规则

证据根目录 F:/RustBuild/audit/tmp/gui-logging-20261112。中间 r1/r2 驱动错误、r3/r4 复制阶段阻断不算对应计算/发布场景验收；取消的编译/构建锁等待不算测试通过或失败。最终版本和结果见 AUDITTOOLS_LOGGING_ACCEPTANCE.md。真实业务文件及日志仅本地留存，不提交客户数据。
