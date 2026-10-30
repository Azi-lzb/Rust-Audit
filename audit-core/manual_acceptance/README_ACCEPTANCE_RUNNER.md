# Rust-Audit 真实业务验收运行器

`open-data-audit acceptance-run` 从 JSON 清单读取本轮真实输入，按 S1、S2、S3-F1 至 F5、S5 的顺序调用现有业务入口，并在新的输出目录生成 `acceptance-report.json` 和 `acceptance-report.md`。

```powershell
dist\open-data-audit.exe acceptance-run `
  --manifest C:\AuditSamples\acceptance.json `
  --output-dir C:\AuditSamples\acceptance-run-20261028
```

清单中的相对路径以清单所在目录为基准。输出目录必须尚不存在，且不能位于任何输入文件或输入目录中；运行器拒绝覆盖已有输出。至少提供 S1、S2、S3、S5 其中一项。没有填写的功能会被省略；S3 内未提供的子功能会记为 `not_requested`。

## 清单字段

所有节都是可选的。以下为字段形状示例，路径请换成业务验收文件：

```json
{
  "s1": {
    "sourceDirectory": "S1/source",
    "templateDirectory": "S1/templates",
    "mappingConfig": "S1/config.xlsx",
    "externalData": "S1/external.xlsx",
    "sidecar": "C:/Audit/bin/lo-convert.exe"
  },
  "s2": {
    "currentDirectory": "S2/current",
    "previousDirectory": "S2/previous",
    "config": "S2/3.0大集中通用配置.xlsx",
    "external": "S2/external.xlsx"
  },
  "s3": {
    "currentData": ["S3/current.csv"],
    "previousData": ["S3/previous.csv"],
    "commonConfig": "S3/3.0大集中通用配置.xlsx",
    "executionConfig": "S3/3.1大集中执行比较配置.xlsx",
    "currentValueConfig": "S3/3.2大集中本期数值核对配置.xlsx",
    "formsConfig": "S3/3.3大集中指标比较拆分配置.xlsx",
    "comparisonWorkbook": null,
    "feedbackTargetWorkbook": "S3/机构说明.xlsx",
    "feedbackWorkbooks": ["S3/反馈.xlsx"],
    "expressionSchema": "legacy-eight-segment"
  },
  "s5": {
    "config": "S5/audit-config.xlsx",
    "source": "S5/current.xlsx",
    "period": "2026-07-31",
    "orgCode": "I001",
    "orgName": "模拟机构",
    "templatePath": "S5/template.xlsx",
    "templateDirectory": null,
    "sourceId": "CURRENT",
    "sourceBusiness": "单位贷款",
    "currentSources": [],
    "previousSource": null,
    "previousSourceId": null,
    "previousPeriod": null,
    "externalSource": null,
    "externalSourceId": null,
    "externalSourceBindings": [],
    "defaultExternalSourceId": null,
    "businesses": [],
    "runId": "acceptance-2026-07"
  }
}
```

省略 `s1`、`s2`、`s3` 或 `s5` 即不运行对应模块。S3-F1 需要本期/上期数据、3.0 和 3.1；成功后其比较文件自动传给 S3-F2 和 S3-F5。若本轮提供 `comparisonWorkbook`，它可用于独立运行 S3-F2/S3-F5。S3-F3 只有在提供反馈工作簿后才运行；目标优先取 `feedbackTargetWorkbook`，否则使用本轮 S3-F2 生成的说明文件。运行器将目标复制到验收输出目录，再调用“备份后原地更新”，因此保留输入目标原件。S3-F4 需要 `currentValueConfig`（3.2）和 `commonConfig`（3.0）。

S3-F5 的 3.3「分析文件数据单位」按已确认口径仍为历史遗留字段；继续沿用 3.0 默认目标单位，不将其用于换算。

S1-F1 使用现有 GUI 工作流服务执行模板匹配、表结构预检、公式区域复制、真实重算、公式/条件格式异常汇总与结果输出。`sidecar` 可显式指定 `lo-convert`；省略时按现有应用规则从环境变量或程序目录发现。Windows 需要可用的 Microsoft Excel COM，Linux 需要可用的 LibreOffice Calc；缺少真实重算引擎时该步骤失败，不读取公式文本充当计算值。

## 报告内容

每个实际执行的步骤记录退出状态、耗时、命令输出/错误原文、生成文件、可解析的异常/行数计数，以及本步骤前后的输入 SHA-256 是否变化。总报告同时保留各输入文件运行前后的哈希。任何业务入口失败或输入文件哈希变化都会令总状态为 `failed`，CLI 以非零代码结束，但仍保留两份报告供排查。

