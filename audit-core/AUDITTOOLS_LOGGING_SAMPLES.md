# GUI 与开发诊断日志样例

日期：2026-11-13。以下摘自最终候选的实际 GUI 服务/API/worker 运行，不是模拟日志字符串。完整记录只保存在本地证据目录，不把真实业务资料纳入 Git。

证据根目录：`F:/RustBuild/audit/tmp/gui-logging-20261112`。最终候选 SHA-256：`3c0be6b11b458b3144547a9de32c4a286ffca0b02cf354459c74062ed29e83cb`。

## 全部成功，未提取到异常

运行 ID：`gui-run-20261113T003538.246018400-19736-0`。完整文件：`api-acceptance-final/passed/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:35:38] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:35:48] [info] S1-F01 最终摘要：审核完成
处理文件：2 个；完成 2 个；部分审核 0 个；待人工重算 0 个；失败 0 个；跳过 0 个
审核异常：0 条（公式 0 条；条件格式 0 条）
总耗时：9.8秒
```

## 执行成功，存在审核异常

运行 ID：`gui-run-20261113T003549.083422200-19736-1`。完整文件：`api-acceptance-final/issues/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:35:49] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:35:58] [info] S1-F01 最终摘要：审核完成（存在审核异常）
处理文件：2 个；完成 2 个；部分审核 0 个；待人工重算 0 个；失败 0 个；跳过 0 个
审核异常：2 条（公式 2 条；条件格式 0 条）
总耗时：9.4秒
```

## 部分结构跳过

运行 ID：`gui-run-20261113T003558.779231400-19736-2`。完整文件：`api-acceptance-final/skipped/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:35:58] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:35:59] [warn] 跳过：机构02_2026-07-31_核查.xlsx，表结构不匹配，已跳过：固定标签匹配率 0% (0/1)
[00:36:07] [info] S1-F01 最终摘要：审核部分完成
处理文件：2 个；完成 1 个；部分审核 0 个；待人工重算 0 个；失败 0 个；跳过 1 个
审核异常：1 条（公式 1 条；条件格式 0 条）
总耗时：8.3秒
```

## 部分文件读取失败

运行 ID：`gui-run-20261113T003634.140170900-19736-7`。完整文件：`api-acceptance-final/partial_failure/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:36:34] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:36:41] [error] 未完成：机构02_2026-07-31_核查.xlsx，文件无法读取，请检查文件格式和占用情况
[00:36:41] [info] S1-F01 最终摘要：审核部分完成
处理文件：2 个；完成 1 个；部分审核 0 个；待人工重算 0 个；失败 1 个；跳过 0 个
审核异常：1 条（公式 1 条；条件格式 0 条）
总耗时：7.6秒
```

## 同名文件定位

运行 ID：`gui-run-20261113T003642.053181700-19736-8`。完整文件：`api-acceptance-final/same_name_failure/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:36:42] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:36:48] [error] 未完成：机构02\同名_2026-07-31_核查.xlsx，文件无法读取，请检查文件格式和占用情况
[00:36:48] [info] S1-F01 最终摘要：审核部分完成
处理文件：2 个；完成 1 个；部分审核 0 个；待人工重算 0 个；失败 1 个；跳过 0 个
审核异常：1 条（公式 1 条；条件格式 0 条）
总耗时：6.6秒
```

## 公式计算未完成

运行 ID：`gui-run-20261113T003607.854501900-19736-5`。完整文件：`api-acceptance-final/calculation_failed/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:36:07] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:36:13] [error] 未完成：机构01_2026-07-31_核查.xlsx，公式计算未完成，相关审核尚未完成；请检查可用的计算程序或查看诊断记录
[00:36:13] [error] 未完成：机构02_2026-07-31_核查.xlsx，公式计算未完成，相关审核尚未完成；请检查可用的计算程序或查看诊断记录
[00:36:13] [info] 自动公式重算不可用或失败，已保留待计算的审核副本。请用 Excel 打开审核副本：按 Ctrl+Alt+F9 强制重算，等待计算完成后按 Ctrl+S 保存，即可在 Excel 中查看「校验区域」结果。请勿修改原始报表文件。
[00:36:13] [info] S1-F01 最终摘要：审核失败（未生成结果）
处理文件：2 个；完成 0 个；部分审核 0 个；待人工重算 2 个；失败 0 个；跳过 0 个
审核异常：0 条（公式 0 条；条件格式 0 条）
总耗时：5.6秒
```

## 内部回退后成功

运行 ID：`gui-run-20261113T003613.811623400-19736-6`。完整文件：`api-acceptance-final/engine_fallback/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:36:13] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:36:33] [info] S1-F01 最终摘要：审核完成（存在审核异常）
处理文件：2 个；完成 2 个；部分审核 0 个；待人工重算 0 个；失败 0 个；跳过 0 个
审核异常：2 条（公式 2 条；条件格式 0 条）
总耗时：20.1秒
```

## 最终结果发布失败

运行 ID：`gui-run-20261113T003648.850489600-19736-9`。完整文件：`api-acceptance-final/output_failure/user-log.txt`、`final-state.json`、`developer.jsonl`、`input-hashes.json`。下列摘录仅省略结果/目录长路径行。

```text
[00:36:48] [info] S1-F01 开始预检：已选 2 个源工作簿
[00:36:57] [error] 未完成：机构01_2026-07-31_核查.xlsx，结果文件未能保存，请检查输出位置和文件占用情况
[00:36:57] [error] 未完成：机构02_2026-07-31_核查.xlsx，结果文件未能保存，请检查输出位置和文件占用情况
[00:36:57] [info] S1-F01 最终摘要：审核失败（未生成结果）
处理文件：2 个；完成 0 个；部分审核 0 个；待人工重算 0 个；失败 2 个；跳过 0 个
审核异常：2 条（公式 2 条；条件格式 0 条）
总耗时：8.3秒
```

## 开发诊断与同一运行关联

同一 engine_fallback 运行的实际技术事件摘录；仅展示 runId/feature/stage/fields 等定位字段，完整错误及路径见上述 JSONL。用户日志没有显示这些正常回退过程。

```json
[
  {
    "feature": "S1-F01",
    "fields": {
      "file": "\\\\?\\F:\\RustBuild\\audit\\tmp\\gui-logging-20261112\\api-acceptance-final\\engine_fallback\\output\\S1-F1_审核结果_20261113_003613_811\\.audit-s1-f1-19736-7\\formula-copy-cacheless.xlsx",
      "reason": "1 个公式格求值为 #NAME?（模板使用了 betteroffice 未实现的函数）"
    },
    "journalSequence": 166,
    "runId": "gui-run-20261113T003613.811623400-19736-6",
    "stage": "betteroffice_refused",
    "time": "2026-11-13T00:36:19.559035400+08:00"
  },
  {
    "feature": "S1-F01",
    "fields": {
      "engineRequest": "auto",
      "nativeOffice": true
    },
    "journalSequence": 167,
    "runId": "gui-run-20261113T003613.811623400-19736-6",
    "stage": "office_worker_start",
    "time": "2026-11-13T00:36:20.104438400+08:00"
  },
  {
    "feature": "S1-F01",
    "fields": {
      "comAvailable": true,
      "engine": "Excel.Application",
      "officePid": 28080,
      "startupTimings": null,
      "workerPid": 24380
    },
    "journalSequence": 169,
    "runId": "gui-run-20261113T003613.811623400-19736-6",
    "stage": "office_worker_ready",
    "time": "2026-11-13T00:36:23.456537+08:00"
  },
  {
    "feature": "S1-F01",
    "fields": {
      "file": "\\\\?\\F:\\RustBuild\\audit\\tmp\\gui-logging-20261112\\api-acceptance-final\\engine_fallback\\output\\S1-F1_审核结果_20261113_003613_811\\.audit-s1-f1-19736-8\\formula-copy-cacheless.xlsx",
      "reason": "1 个公式格求值为 #NAME?（模板使用了 betteroffice 未实现的函数）"
    },
    "journalSequence": 177,
    "runId": "gui-run-20261113T003613.811623400-19736-6",
    "stage": "betteroffice_refused",
    "time": "2026-11-13T00:36:29.473819900+08:00"
  },
  {
    "feature": "S1-F01",
    "fields": {
      "engine": "Excel.Application",
      "officePid": 28080,
      "workerPid": 24380
    },
    "journalSequence": 184,
    "runId": "gui-run-20261113T003613.811623400-19736-6",
    "stage": "office_session_shutdown",
    "time": "2026-11-13T00:36:30.306860700+08:00"
  }
]
```

按同一 runId 收集全部分段，依 journalSequence 排序。最终统计另存于 runStatistics 和摘要诊断事件，不从以上文本反向统计。

## 真实七本工作簿

`after-real-seven-final/capture-summary.json`：7 本完成；317 条公式异常、5 条条件格式异常、合计 322；执行失败/跳过/待人工重算均为 0。GUI 用户日志 2 条，诊断未标记丢失。修改前普通日志 58 条。`real-seven-output-parity-final.json` 确认结果单元格、Sheet/合并区域和定位链接（仅运行目录变化）一致，输入哈希未变化。

Debug 实际总耗时 151.3 秒，驱动观察 151.36 秒；API 轮询 P95 约 31ms。该轮用于日志和结果验收，不用于与旧 Release 做性能倍率比较。

## 其他正式功能

S1-F02、S2-F01、S3-F01/F02/F04 的现有 GUI API 单元/业务测试保留实际工作簿输出断言，并捕获用户日志、最终状态与开发诊断。最终目录和测试结果见验收文档。S3-F03/F05 的最终独立 GUI 服务运行见 `api-s3-final/results.json`：反馈导入 1 行且原地更新前备份正确；表单生成 1 机构、1 匹配指标、1 输出工作簿。


## 真实 129 本长时间批量

运行 ID：gui-run-20261113T004143.450531300-25664-0。完整记录 real-129-final。

129 本全部执行完成；3352 公式异常 + 257 条件格式异常 = 3609；失败/跳过/部分/待人工重算 0。最终摘要总耗时 27分13.7秒；用户日志 2 条，开发诊断 1340 条，输入哈希不变。统计与实际结果表 3609 行一致。日志正文和阶段变化均保留于本地，诊断未标记丢失。
