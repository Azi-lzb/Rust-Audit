# Audit Rust V2 Windows 人工测试样例

本目录只含合成数据和 Rust 实际输出，可用于人工检查 S1-F02 任意行汇总/历史富化、S2-F01 基础/外部文件路径，以及 S3-F01、S3-F02、S3-F04、S5-F01 的最小正常/边界路径。S1-F02、S3、S5 样例由 `audit-core/examples/manual_acceptance_samples.rs` 生成；普通工作簿使用 Workspace Rust `excel_writer`，S5 条件格式测试输入使用 Audit 已有的测试期 Umya 依赖写入 OOXML 规则。S2 输入和配置复制自仓库内的合成测试夹具。旧 Python 项目、真实样本和配置未参与生成。

## 启动程序

GUI：

~~~powershell
& 'F:\RustBuild\audit_manual_handoff_20261027_final\debug\open-data-audit-gui.exe' --port 8753 --config 'C:\Users\AZI\Desktop\VibeCode\audit-dev\audit-core\manual_acceptance\20261027\s2_f01\s2_config.xlsx'
~~~

若要在 GUI 中直接使用新增的 S2 外部核对规则配置，可另开端口启动一个实例：

~~~powershell
& 'F:\RustBuild\audit_manual_handoff_20261027_final\debug\open-data-audit-gui.exe' --port 8754 --config (Join-Path $sample 's2_external_hit\s2_external_config.xlsx')
~~~

随后打开 <http://127.0.0.1:8754>，在 S2 页面选择本目录 `s2_external_hit\input\current`、`s2_external_hit\input\previous`，并分别选择 `external_normal.xlsx` 或 `external_mismatch.xlsx`。

CLI：

~~~powershell
$exe = 'F:\RustBuild\audit_manual_handoff_20261027\debug\open-data-audit.exe'
$sample = 'C:\Users\AZI\Desktop\VibeCode\audit-dev\audit-core\manual_acceptance\20261027'
$stamp = Get-Date -Format yyyyMMdd_HHmmss
~~~

当前提供给人工测试的 GUI 地址：<http://127.0.0.1:8753>。GUI 默认加载本目录的合成 S2 配置；进入 S2 后分别选择 `s2_f01\input\current` 与 `s2_f01\input\previous`，需要外部核对时再选 `s2_f01\external_data.xlsx`。S5 面板使用本目录的 `s5_config.xlsx`、源文件和空模板目录。若服务已关闭，请使用上方命令启动，并打开该地址。

已验证的直接业务入口是 CLI。GUI Windows 构建成功；本轮没有做浏览器人工操作，也不把 GUI 启动等同于每个功能已经从界面验收。

## S2-F01 基础与外部数据路径

`s2_f01/` 使用仓库内生成的合成报表、有效配置及外部数据。当前与上期文件含同一组 11 项指标值；命令会输出旧版三张分类审核表。配置中的业务规则会产生审核记录，因此此基础路径不是零异常结果。带外部文件的路径另外覆盖外部数据读取与输入提示，不代表已验证“外部核对规则命中”。

~~~powershell
$s2 = Join-Path $sample 's2_f01'
& $exe s2 --current-dir (Join-Path $s2 'input\current') --previous-dir (Join-Path $s2 'input\previous') --config (Join-Path $s2 's2_config.xlsx') --output-dir (Join-Path $s2 "outputs\manual_base_$stamp")
& $exe s2 --current-dir (Join-Path $s2 'input\current') --previous-dir (Join-Path $s2 'input\previous') --config (Join-Path $s2 's2_config.xlsx') --external (Join-Path $s2 'external_data.xlsx') --output-dir (Join-Path $s2 "outputs\manual_external_$stamp")
~~~

本轮生成的结果在 `s2_f01/outputs/normal/` 和 `s2_f01/outputs/external_exception/`；输入、配置、独立可核对的零变动预期及结果边界见 `expected-results.txt`。

### 外部核对命中样例

`s2_external_hit/` 补充一组 Rust CLI 可执行的外部核对正常/异常案例。输入报表沿用仓库合成夹具；本例的配置、外部工作簿和预期值由 Rust 示例生成，异常结果由独立数值计算得出。正常外部值为 `1,234,567.89` 元，预期不触发外部异常；异常外部值为 `1,234,500.00` 元，差额 `67.89` 元，预期产生 1 条硬性 `MANUAL-EXT-001` 外部核对异常。

~~~powershell
$s2hit = Join-Path $sample 's2_external_hit'
& $exe s2 --current-dir (Join-Path $s2hit 'input\current') --previous-dir (Join-Path $s2hit 'input\previous') --config (Join-Path $s2hit 's2_external_config.xlsx') --external (Join-Path $s2hit 'external_normal.xlsx') --output-dir (Join-Path $s2hit "outputs\manual_normal_$stamp")
& $exe s2 --current-dir (Join-Path $s2hit 'input\current') --previous-dir (Join-Path $s2hit 'input\previous') --config (Join-Path $s2hit 's2_external_config.xlsx') --external (Join-Path $s2hit 'external_mismatch.xlsx') --output-dir (Join-Path $s2hit "outputs\manual_mismatch_$stamp")
~~~

Rust CLI 实跑后，正常/异常报告均为 17 条结果；异常案例比正常案例多 1 条硬性外部异常。两份《报表审核清单》均由共享 Calamine reader 重开，外部核对结果值、差额和规则编号符合上述预期。未在本轮重跑 Python 对拍；请以该样例开始 Windows 人工操作验收。

## S1-F02 任意行汇总与历史富化

`s1_f02/` 使用 Rust `excel_writer` 生成的模板、历史工作簿和源文件。正常案例有两个源工作簿，分别含 2 行和 3 行明细；结构异常案例额外包含一个表头差异文件，用来确认该文件被跳过而有效文件仍继续汇总。

~~~powershell
$s1 = Join-Path $sample 's1_f02'
& $exe s1-region-summary --template-workbook (Join-Path $s1 'template.xlsx') --input-dir (Join-Path $s1 'input\normal') --history-workbook (Join-Path $s1 'history.xlsx') --output-workbook (Join-Path $s1 'outputs\normal\s1_f02_summary.xlsx') --feature-name '任意行汇总' --developer-view
& $exe s1-region-summary --template-workbook (Join-Path $s1 'template.xlsx') --input-dir (Join-Path $s1 'input\structure_exception') --history-workbook (Join-Path $s1 'history.xlsx') --output-workbook (Join-Path $s1 'outputs\structure_exception\s1_f02_summary.xlsx') --feature-name '任意行汇总' --developer-view
~~~

Rust 实跑的两份输出均为 5 行、11 列；正常案例历史匹配 3 行，其中“甲机构”规则 1 的历史触发条数为 `10|12`，分类为 `类别甲|类别乙`，审核结果为 `通过|未通过`。结构异常案例检查 3 个源工作簿，2 个通过、1 个跳过（表头 5/6 匹配），仍生成 5 行结果。两份结果均已由 Rust Calamine reader 重开验证。独立预期及边界见 `expected-results.txt`。

## S3-F01 正常与异常样例

文件：s3_common_config.xlsx、s3_feature_config.xlsx、s3_previous.csv、s3_current_normal.csv、s3_current_exception.csv。

~~~powershell
New-Item -ItemType Directory -Force (Join-Path $sample 'outputs') | Out-Null
& $exe s3-execution-compare --current-data (Join-Path $sample 's3_current_normal.csv') --previous-data (Join-Path $sample 's3_previous.csv') --common-config (Join-Path $sample 's3_common_config.xlsx') --feature-config (Join-Path $sample 's3_feature_config.xlsx') --output-workbook (Join-Path $sample "outputs\s3_normal_$stamp.xlsx") --expression-schema five-segment --developer-view
& $exe s3-execution-compare --current-data (Join-Path $sample 's3_current_exception.csv') --previous-data (Join-Path $sample 's3_previous.csv') --common-config (Join-Path $sample 's3_common_config.xlsx') --feature-config (Join-Path $sample 's3_feature_config.xlsx') --output-workbook (Join-Path $sample "outputs\s3_exception_$stamp.xlsx") --expression-schema five-segment --developer-view
~~~

本轮已实际生成 outputs/s3_normal.xlsx 和 outputs/s3_exception.xlsx；新命令使用时间戳路径，避免覆盖。

## S3-F02 比较文件到说明文件

`s3_f02/` 提供普通筛选和多地区边界两个可直接打开的比较结果工作簿，以及对应 3.0 机构地区参照配置。GUI 中打开“大集中统计系统”，选择其中一个比较结果文件和 `s3_f02_common_config.xlsx`，设置输出目录，再点“F2 比较文件 ⇨ 说明文件”。

~~~powershell
$s3f02 = Join-Path $sample 's3_f02'
& $exe s3-export-explanations --comparison-file (Join-Path $s3f02 'input\normal\comparison_normal.xlsx') --config-workbook (Join-Path $s3f02 's3_f02_common_config.xlsx') --output-workbook (Join-Path $s3f02 'outputs\normal\s3_f02_normal.xlsx') --developer-view
& $exe s3-export-explanations --comparison-file (Join-Path $s3f02 'input\multi_region\comparison_multi_region.xlsx') --config-workbook (Join-Path $s3f02 's3_f02_common_config.xlsx') --output-workbook (Join-Path $s3f02 'outputs\multi_region\s3_f02_multi_region.xlsx') --developer-view
~~~

正常输入含 1 行“是”和 1 行“否”，预期只导出 1 行；多地区输入对同机构代码提供 R001/R002 两个地区，分别用“是”和“✓”标记，预期输出机构代码 `ORG2002R001`、`ORG2002R002`，且按地区使用不同配置名称。两份 Rust 输出已生成并由 Calamine 读回核验，独立预期见 `expected-results.txt`。此合成 CLI/API 证据不等于 GUI 人工验收。

## S3-F04 本期数值核对（XLSX）

`s3_f04/` 是本轮早期字段语义诊断目录，其中的取反标识示例曾填“否”，不符合旧版“非空即启用”的规则；不要用于验收。正式样例及输出仅使用 `s3_f04_corrected/`。

`s3_f04_corrected/` 含机器生成的本期工作簿、3.0 通用配置和 3.2 核对配置。单一本期工作簿内分别有“第1批”和“第2批”值，不需要上期文件。规则为 `left > right`，仅在触发时输出，专门核验频度/批次维度未被合并。

~~~powershell
$sampleS3F04 = Join-Path $sample 's3_f04_corrected'
& $exe s3-check-current-values --current-csv (Join-Path $sampleS3F04 's3_f04_current_normal.xlsx') --config-workbook (Join-Path $sampleS3F04 's3_f04_common_config.xlsx') --config-workbook (Join-Path $sampleS3F04 's3_f04_feature_config.xlsx') --output-workbook (Join-Path $sampleS3F04 "outputs\s3_f04_normal_$stamp.xlsx") --log-workbook (Join-Path $sampleS3F04 "outputs\s3_f04_normal_log_$stamp.xlsx") --developer-view
& $exe s3-check-current-values --current-csv (Join-Path $sampleS3F04 's3_f04_current_exception.xlsx') --config-workbook (Join-Path $sampleS3F04 's3_f04_common_config.xlsx') --config-workbook (Join-Path $sampleS3F04 's3_f04_feature_config.xlsx') --output-workbook (Join-Path $sampleS3F04 "outputs\s3_f04_exception_$stamp.xlsx") --log-workbook (Join-Path $sampleS3F04 "outputs\s3_f04_exception_log_$stamp.xlsx") --developer-view
~~~

此命令历史参数名 `--current-csv` 同样接受 XLSX。已生成并读回的基准输出为 `s3_f04_normal.xlsx` 和 `s3_f04_exception.xlsx`；预期与实跑见 `expected-results.txt`。

## S5-F01 正常与异常样例

文件：s5_config.xlsx、s5_source_normal.xlsx、s5_source_exception.xlsx。配置使用动态 V_CUR 路径，当前值在 贷款表!C4；不需要上期工作簿。S5 命令要求模板目录预先存在，先创建空目录：

~~~powershell
New-Item -ItemType Directory -Force (Join-Path $sample 's5_templates') | Out-Null
$normalOut = Join-Path $sample "outputs\s5_manual_normal_$stamp"
$exceptionOut = Join-Path $sample "outputs\s5_manual_exception_$stamp"
& $exe s5-audit-run --config (Join-Path $sample 's5_config.xlsx') --source (Join-Path $sample 's5_source_normal.xlsx') --template-dir (Join-Path $sample 's5_templates') --source-id CUR --source-business '单位贷款' --period 2026-08-31 --org-code I001 --org-name '模拟机构' --business '单位贷款' --output-dir $normalOut --run-id "manual-normal-$stamp" --developer-view
& $exe s5-audit-run --config (Join-Path $sample 's5_config.xlsx') --source (Join-Path $sample 's5_source_exception.xlsx') --template-dir (Join-Path $sample 's5_templates') --source-id CUR --source-business '单位贷款' --period 2026-08-31 --org-code I001 --org-name '模拟机构' --business '单位贷款' --output-dir $exceptionOut --run-id "manual-exception-$stamp" --developer-view
~~~

本轮已实际生成并由 Rust Calamine reader 重新打开：

- outputs/s5_normal_corrected/demo-s5-normal-corrected/审核结果.xlsx
- outputs/s5_exception/demo-s5-exception/审核结果.xlsx

运行总账.json 保存在相应运行目录内。

2026-10-27 当前独立构建的最新实跑结果位于 `outputs/codex_ready_20261027_114631/`：S3-F01 为 `s3_normal.xlsx`、`s3_exception.xlsx`；S5-F01 结果分别位于 `s5_normal/manual-normal-20261027_114631/审核结果.xlsx` 和 `s5_exception/manual-exception-20261027_114631/审核结果.xlsx`。状态与独立预期见 `expected-results.txt`。

## S5-F01 条件格式公式正常与异常样例

`s5_conditional/` 是专门验证业务入口中条件格式公式计算的独立样例。本期值位于 `贷款表!C4`，规则 `C4>120` 应用到 `C4`；样式只用于让 Excel 文件包含有效的条件格式规则，不作为比较项。

```powershell
$sampleS5Cf = Join-Path $sample 's5_conditional'
$normalOut = Join-Path $sampleS5Cf 'outputs\s5_conditional_normal_ready'
$exceptionOut = Join-Path $sampleS5Cf 'outputs\s5_conditional_exception_ready'
& $exe s5-audit-run --config (Join-Path $sampleS5Cf 's5_config.xlsx') --source (Join-Path $sampleS5Cf 's5_source_normal.xlsx') --template-dir (Join-Path $sampleS5Cf 's5_templates') --source-id CUR --source-business '单位贷款' --period 2026-08-31 --org-code I001 --org-name '模拟机构' --business '单位贷款' --output-dir $normalOut --run-id 's5-cf-normal-ready-20261027' --developer-view
& $exe s5-audit-run --config (Join-Path $sampleS5Cf 's5_config.xlsx') --source (Join-Path $sampleS5Cf 's5_source_exception.xlsx') --template-dir (Join-Path $sampleS5Cf 's5_templates') --source-id CUR --source-business '单位贷款' --period 2026-08-31 --org-code I001 --org-name '模拟机构' --business '单位贷款' --output-dir $exceptionOut --run-id 's5-cf-exception-ready-20261027' --developer-view
cargo run --offline --locked -p audit-core --example manual_acceptance_samples -- --verify-s5-conditional $sampleS5Cf
```

正式手测输出是两个 `*_ready` 目录下的 `s5-cf-*-ready-20261027\审核结果.xlsx`。`outputs\s5_conditional_normal/` 和 `outputs\s5_conditional_exception/` 是由旧构建二进制生成的诊断输出，未执行条件格式阶段，请勿作为验收结果。

## 已知边界

- S1-F01 真实公式复制/Excel COM 强制重算仍未闭合，本目录不包含 S1-F01 验收样例，也不应把 S1 按钮未接通理解为已实现。
- S1-F02、S3-F02 已包含各自独立的合成样例；CLI 实跑不等于 Python/Rust 完整对拍或 Windows GUI 人工验收。
- 当前标准不再以普通字体、颜色、填充、边框、对齐、样式编号、尺寸或打印外观差异阻塞迁移；公式、值/类型、业务结果、结构和文件可读性仍需正确。
- 这里的输出仅证明所列合成样例可运行，不能代替真实业务样本或 Windows 人工验收，也未做 Linux/deb 验收。
注意：outputs/s5_normal/demo-s5-normal 是首次检查配置表达式方向时留下的试跑结果，不是最终正常样例结果。请使用 outputs/s5_normal_corrected/demo-s5-normal-corrected/审核结果.xlsx；该文件对应当前配置并已核对为 PASSED 1。
