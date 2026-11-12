# AuditTools GUI 日志与业务统计规范

适用产品：本仓库现有 Rust + HTML GUI，正式范围 S1/S2/S3。2026-11-13。

## 用户日志

正常任务保留一次开始和一次最终摘要；文件、Sheet、阶段沿用现有 currentWork/进度显示，不再重复追加正常阶段。引擎路由、重算轮数、缓存处理、修改后端、回退、COM/worker、XML/ZIP 和阶段计时进入独立开发诊断。即使开启原有明细设置，也不把技术事件重新放入用户日志。

保留真实业务警告、跳过、部分未完成和执行错误。错误尽量包含来源相对路径以及已有的真实 Sheet 上下文；不增加读取以补造定位。不同目录同名文件保留目录区别。面向用户简述原因，完整错误链进入诊断。审核异常不等于程序失败，结构跳过不等于通过。

最终摘要包括实际业务统计、是否交付结果、必要输出位置和 Instant 单调计时的端到端耗时。任务全部跳过、启动失败、计算未完成或结果发布失败时不得宣称审核成功。未知的文件完成情况明确为“未确认”，不得猜测失败数量。

现有布局、级别筛选、保存、复制、清空、滚动、2000 条用户日志缓冲、进度及任务操作均保持。保存/复制仍操作现有用户日志；诊断独立持久化，不新增交互。

## 各功能统计契约

- S1-F01：来自最终 file_results、formula_issue_count、conditional_issue_count、total_issue_count 和 result_path/result_error。分别报告完成、部分、待人工重算、失败和预检跳过文件，以及公式/条件格式异常。COMPLETED 只表示该文件审核执行完成，不表示没有业务异常。现有报告没有可靠的检查全集，因此不显示推算的检查/通过数。提取成功但最终结果发布失败，必须明确未交付。
- S1-F02：源文件、结构通过/部分继续/跳过、输出 Sheet 及数据行。属于汇总统计，不冒称审核通过。
- S2-F01：机构、报表、导入本期指标记录、发现条目、审核异常。summary.indicators 不是全部已检查数，total_findings - abnormal_findings 也不是通过数。
- S3-F01：本期/上期输入记录、比较行、动作与表达式规则触发、表达式执行错误。触发次数不是唯一异常记录数，不能相加虚构总量。
- S3-F02：实际导出的说明行数。
- S3-F03：反馈文件、实际导入行、未匹配行、冲突行，保留原地更新备份和必要冲突路径。未匹配不是已通过。
- S3-F04：本期输入记录、配置/启用规则、报告结果行。结果行不冒称通过数。
- S3-F05：比较行、机构、匹配指标、生成工作簿，保留业务警告。不是规则审核通过统计。

计数直接来自业务报告，禁止从日志反向统计、额外读簿统计或把多阶段次数当唯一记录数。重试/切换引擎成功只按最终结果统计一次。不修改既有业务报告、任务返回状态或退出语义。

## 开发诊断

每次任务建立 runId。GUI 状态的 runId、runFeature、diagnosticLogPath、runStatistics、lastRunSummary 和 diagnosticIncomplete 可关联运行。事件保留功能、现有文件/Sheet/阶段上下文、路由、计时、已有错误链、最终结构化统计；S1 增补实际计算轮次及 COM/IPC 生命周期事件。不存在的 DAG 节点或计数不伪造。

trace_error 保留 anyhow 完整错误链；仅当原错误已经捕获 backtrace 时保存原始堆栈，否则 backtrace 为 null，不把处理点伪称错误起点。当前不保证所有第三方引擎内部堆栈或全部原始子进程输出均可获得。源单元格内容全集、数据行和公式结果全集不额外写入；已有错误文本可能包含必要公式/路径，诊断应按本地业务资料保管。

普通队列 2048、warn/error 预留队列 256，try_send 不等待磁盘。后台异步写 JSONL；诊断队列、线程或文件失败仅标记 diagnosticIncomplete，不更改审核结论。每段约 16 MiB，保留同次运行全部分段，清理时只处理本模块命名文件，保留最近 40 次运行。内存有界；单次超大运行的磁盘量仍随事件量增加，磁盘不足会标记不完整。突发超过队列容量不保证零丢失，不得在 incomplete=true 时声称完整还原。

默认目录：Windows `%LOCALAPPDATA%/audittools/diagnostics/`；Linux `$XDG_STATE_HOME/audittools/diagnostics/` 或 `~/.local/state/audittools/diagnostics/`；指定 AUDIT_STATE_DIR 时使用其 diagnostics 子目录。实际以 diagnosticLogPath 为准。文件名 gui-run-*.jsonl，分段为同名 .segment-000001.jsonl 等；合并同次运行全部段后按 journalSequence 排序，避免优先错误队列改变物理落盘顺序造成误读。

## 验证边界

自动化、实际 GUI API/worker 运行、真实样本结果对拍、Windows/Linux 和打包验证分别记录。API 捕获是同一 GUI 服务的真实执行记录，不是另一个 CLI 产品，也不是手写日志模拟。浏览器人工体验和离线 Linux 未验证时必须明示。详见 AUDITTOOLS_LOGGING_ACCEPTANCE.md。
