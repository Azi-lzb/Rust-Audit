# Rust-Audit 八功能路径架构表（详细版）

> 依据：当前代码（`audit-core/src/*.rs`、`crates/excel-editor/*`）+ 实测报告。
> 耗时：`rustaudit_alignment_report.json`（2026-11-07 全量两跑）
> 与 `s3f5_engine_solo_report.json`（2026-11-08 最新）。
> 平台：Windows x86_64 / 12 核 / release。

---

## 一、S1-F01 逐笔统计系统（公式副本 + 重算 + 审核）

**总耗时 189.5 / 189.6s**（129 本核查表）

### 阶段链路（每个源文件循环执行）

| # | 阶段 | 实现 | 引擎/路径 | 说明 |
|---|---|---|---|---|
| 1 | **模板与表结构预检** | `s1_template_rules` + `s1_structure_compare` | 纯 Rust（calamine） | 模板健康度、必需工作表、命名区域、列头比对；不通过的文件剔除 |
| 2 | **数据期识别** | `detect_period_from_workbooks` | 纯 Rust | 从源工作簿识别唯一数据期；冲突则整体失败 |
| 3 | **生成审核副本** | `s1_formula_copy` | **ExcelModifier**：OOXML → openpyxl → Umya | 复制模板 + 写入公式区域 + 注释 + 定义名称；源文件只读 |
| 4 | **外部引用剥离 / open-repair** | `excel_open_repair` | 纯 Rust（ZIP 级） | 四形态 `[N]表名!` / 路径式剥离；冻结窗格修复；断链清理 |
| 5 | **真实重算** | `s1_recalculation` | **calc_engine=auto** → Windows 原生 Excel COM（批量/单本），Linux → LibreOffice sidecar | 重算后公式缓存值落地；失败保留副本转人工 |
| 6 | **公式异常提取** | `s1_formula_audit::extract_s1_formula_issues` | 纯 Rust（读重算后副本） | 3345 条基线；错误值/公式引用异常 |
| 7 | **校验区域公式扫描** | `s1_validation_formula_scan::scan` | 纯 Rust（模板预载 plan + 副本扫描） | 校验区域公式错误 token；模板读取一次复用 |
| 8 | **条件格式提取** | `s1_conditional_audit` | **源=用户选择**（默认原始源文件）；判定方式 `PYTHON`（纯 OOXML 规则求值）；可选 `COM_DISPLAY`（Excel COM → WPS COM 回退） | 257 条基线；空白格纳入开关生效 |
| 9 | **历史说明富化** | `s1_history::process_s1_history` + `s1_issue_chain` | 纯 Rust（读 1.逐笔统计系统_配置.xlsx） | 为审核发现补历史说明/问题链 |
| 10 | **汇总与发布** | `s1_summary_merge` + `s1_result_workbook` | 纯 Rust（excel-writer） | 生成审核结果工作簿 + recovery manifest |

### 引擎与开关

| 开关 | 默认 | 作用 |
|---|---|---|
| `calc_engine` | `auto` | 第 5 步重算引擎（auto=原生 COM；lo=LibreOffice） |
| `edit_engine` | `auto` | 第 3 步副本写入后端顺序 |
| `conditionalFormatEvaluator` | `PYTHON` | 第 8 步判定方式；`COM_EVALUATE` 显式未实现报错 |
| `s1ConditionalFormatSource` | `source` | 第 8 步读原始源文件还是修复后副本 |
| `conditionalFormatIncludeBlankCells` | 开 | 第 8 步空白格命中是否计入 |

### 失败策略
- 第 1-2 步不通过 → 该文件剔除 / 整体拒绝，不产生副本
- 第 3 步后端写前拒绝 → 自动换下一个后端；写后失败或验证失败 → **fail-closed**
- 第 5 步重算失败 → 保留 cache-cleared 副本，转"待人工重算"（不冒充成功）

### 对拍
**3602 条逐条一致**（3345 公式 + 257 条件格式；Rust-only 0 / Baseline-only 0 / 字段值差 0）

---

## 二、S1-F02 汇总校验结果（505 行）

**总耗时 1.0 / 1.1s**

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **输入校验与发现** | `discover_source_workbooks` | 纯 Rust | 模板 + 源目录 + 历史配置 + 输出位置校验；递归/根层发现 |
| 2 | **模板命名区域解析** | `read_defined_names(template)` | 纯 Rust（calamine） | 提取命名区域、工作表映射 |
| 3 | **源工作簿结构预检** | `compare_s1_source_structure_dump` | 纯 Rust | 逐文件 `结构预检 [i/N]`；不通过跳过 |
| 4 | **区域聚合** | `summarize_s1_named_regions_selected` | 纯 Rust | 按命名区域收集值 |
| 5 | **三段时间列** | 内置规则 | 纯 Rust | **任务说明 / 业务说明 / 本地校验结果**，各自带历史追加字段规则 |
| 6 | **宽字符对齐与写簿** | `excel-writer` | 纯 Rust（新建 Workbook） | 中文/全角字符宽度计算（用于列宽）；输出 `.xlsx` |

### 对拍
**505 行逐格 0 差异**（任务说明 130 / 业务说明 333 / 本地校验结果 45）

**说明**：不走 ExcelModifier、无编辑器依赖，因此 1s 级。

---

## 三、S2-F01 大集中核对（1087 条）

**总耗时 0.2 / 0.2s**

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **审核配置校验** | `s2_config_check` | 纯 Rust（calamine） | 阻断级错误直接终止 |
| 2 | **读取本期报表** | `read_value_columns_with_iso_dates` | 纯 Rust（calamine） | 目录批量读 |
| 3 | **读取上期报表** | 同上 | 纯 Rust | 同上 |
| 4 | **检查导入数据** | `s2_import` | 纯 Rust | 指标归一；导入/配置提示（**运行诊断，不计入 findings**） |
| 5 | **执行环比规则** | `s2_period` + `s2_rule_engine` | 纯 Rust（内存） | 本期 vs 上期 |
| 6 | **执行校验规则** | `s2_rule_engine` | 纯 Rust | 按配置检查本期 |
| 7 | **读取大集中数据** | `s2_external_import` | 纯 Rust（calamine） | label 探测列范围 1..=4（对齐源生产）；缺失按 0 兜底 |
| 8 | **执行外部核对** | `s2_external` | 纯 Rust | 容差 `centralTolerance` 默认 100 元 |
| 9 | **生成审核结果** | `s2_exporter` | 纯 Rust（excel-writer） | 三表分类写入：环比规则 729 / 外部核对 357 / 校验规则 4 |

### 对拍
**1087 行逐格 0 差异**（729 + 357 + 4）

---

## 四、S3-F01 执行比较（3868 行）

**总耗时 2.3 / 2.4s**

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **读本期数据** | `read_value_columns` | 纯 Rust（calamine，支持 CSV/.xls/.xlsx） | 2026.10 月报二批 |
| 2 | **读上期数据** | 同上 | 纯 Rust | 2026.09 |
| 3 | **读 3.0 通用配置** | `read_value_columns` | 纯 Rust | 机构/地区/币种/频度字典 |
| 4 | **读 3.1 执行比较配置** | `read_value_columns` | 纯 Rust | 指标参照 / 规则动作 / 表达式规则表 |
| 5 | **表达式解析** | 内置 | 纯 Rust | **legacy-eight-segment**（源生产默认）；五段式仅配置检查用 |
| 6 | **数据导入校验** | 内置 | 纯 Rust | 本期/上期问题存在则**停止生成**（不产出错误结果） |
| 7 | **逐规则比较** | `s3_execution_compare` | 纯 Rust | 环比规则 / 外部核对 / 校验规则；Decimal 精确路径；`_accu_active` 累计指标门控 |
| 8 | **排序与分组** | 内置 | 纯 Rust | 三级键 casefold（对齐 VBA EndCompare） |
| 9 | **写比较结果** | `excel-writer` | 纯 Rust | 动态表名 `YYYY年M月对比结果`；填充色（警戒区间） |

### 对拍
**3868 行逐格 0 差异**（含填充色分布）

---

## 五、S3-F02 说明文件导出（单表）

**总耗时 1.1 / 1.1s**

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **读比较结果** | `read_workbook` / `read_value_columns` | 纯 Rust（calamine） | 读取 S3-F01 产物 |
| 2 | **筛选"是否说明=是"行** | 内置 | 纯 Rust | 只导出需要说明的行 |
| 3 | **推导表名** | 内置 | 纯 Rust | `(说明)YYYY.MM 说明文件`（对齐源生产 `_explanation_sheet_title`） |
| 4 | **写单表** | `excel-writer` | 纯 Rust | 不再生成自加的"生成信息"诊断表 |

### 对拍
**单表逐格 0 差异**

---

## 六、S3-F03 反馈回写（变动原因）

**总耗时 0.2 / 0.2s**（5 行真实闭环）

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **读反馈文件** | `read_value_columns_with_iso_dates` | 纯 Rust（calamine） | 扩展名放开 `.xlsx/.xlsm/.xls/.xlsb/.ods`；`.xls` 伪装件（PK 签名）嗅探 + 临时副本 |
| 2 | **读目标说明工作簿** | 纯 Rust | 纯 Rust | 目标仅允许 `.xlsx`（原地回写） |
| 3 | **匹配与冲突检测** | `s3_explanation_feedback` | 纯 Rust | 按关键列匹配；冲突写冲突报告 |
| 4 | **生成回写计划** | 同上 | 纯 Rust | 逐行 `SetCellValue` |
| 5 | **执行回写** | 同上 | **ExcelModifier**：OOXML 直改 → Umya（generic 路由） | 原地更新目标工作簿 |
| 6 | **备份与发布** | 同上 | 纯 Rust | 更新前备份 |

### 对拍
**exit 0、两跑一致**；真实闭环 **imported 5 / unmatched 0 / conflicts 0**，回写后 5 行变动原因全部非空

---

## 七、S3-F04 本期数值核对（33 行）

**总耗时 1.0 / 1.0s**

| # | 阶段 | 实现 | 路径 | 说明 |
|---|---|---|---|---|
| 1 | **读本期数据** | `read_value_columns_with_iso_dates` | 纯 Rust（calamine） | |
| 2 | **读核对规则配置** | 同上 | 纯 Rust | `3.2大集中本期数值核对_配置.xlsx` |
| 3 | **逐条核对** | `s3_current_value_check` | 纯 Rust | 容差/单位换算（`unit_conversion`） |
| 4 | **写核对结果** | `excel-writer` | 纯 Rust | 33 行 |

### 对拍
**33 行逐格 0 差异**

---

## 八、S3-F05 指标比较拆分（金融表单，8 表）

**FAST 直出 1.7-2.0s ／ openpyxl 往返 4.6-5.0s**

### 阶段链路

| # | 阶段 | 实现 | 引擎/路径 | 实测 |
|---|---|---|---|---|
| 1 | **读比较结果** | `read_value_columns_with_iso_dates` | 纯 Rust（calamine） | 含在 read 133ms |
| 2 | **读 3.3 配置 + 3.0 警戒** | 同上 | 纯 Rust | 含在 read 133ms |
| 3 | **规划机构表单** | `plan_indicator_comparison_split` | 纯 Rust（内存） | **510ms** |
| 4 | **模板编译（一次）** | `s3_fast_ooxml::compile_template` | 纯 Rust（ZIP + 正则静态缓存） | **~800-890ms** |
| 5 | **并行渲染（8 线程）** | `s3_fast_ooxml_render::render_org` | 纯 Rust | **~105-117ms/表** |
| 6 | **SheetPlan 处置** | 同上 | 纯 Rust | 删内嵌配置页（使用说明/转表设置）+ 禁用表单 + 空表；公式引用迭代收敛→降级隐藏；definedNames owner 判定/localSheetId 重映射/越界丢弃；bookViews 收敛 |
| 7 | **轻量完整性校验** | `validate_core` | 纯 Rust | 核心 XML 可解析 + sheet↔rels↔part 无悬空 |
| 8 | **写 ZIP** | `write_zip` | 纯 Rust（zip crate） | **~17-29ms/表** |
| 8' | **（可选）openpyxl 往返** | `fast_roundtrip` worker | Python openpyxl（整簿加载→保存） | **~2.2s/表**；包由 openpyxl 序列化器重建 |

### 引擎与开关

| 开关 | 值 | 效果 |
|---|---|---|
| `edit_engine=auto/ooxml` | `FastRenderMode::Direct` | FAST 直出，**1.7-2.0s** |
| `edit_engine=openpyxl` | `FastRenderMode::OpenpyxlRoundtrip` | FAST 渲染 + openpyxl 往返，**4.6-5.0s** |
| `calculate`（无） | — | 不依赖 Excel/WPS/OpenOffice |

### 已下线
**EditPlan 逐格编辑路径**（曾 513-601s）——S3-F05 不再使用。

### 对拍
**8 表逐格 0 差异**；三方（FAST / openpyxl / Python 基线）**全部 0 差异**；两跑一致；输入 SHA-256 不变。

---

## 九、横向对比

### 按路径类型

| 类型 | 功能 | 特征 | 耗时区间 |
|---|---|---|---|
| **纯 Rust 读写**（无 ExcelModifier） | S1-F02、S2-F01、S3-F01、S3-F02、S3-F04 | calamine 读 + excel-writer 写 / 内存规则引擎 | **0.2-2.4s** |
| **ExcelModifier 编辑管线** | S1-F01、S3-F03 | OOXML → openpyxl → Umya 写前拒绝回退 | 0.2s（5 行）～**189s**（129 本 + COM 重算） |
| **自研 ZIP 级渲染器** | S3-F05 | FAST_OOXML（模板编译一次 + 回填 + 并行） | **1.7-2.0s** |

### 耗时构成（已实测定位的）

| 功能 | 主要耗时 | 性质 |
|---|---|---|
| S1-F01 | Excel COM 批量重算（129 本） | 外部进程，非 Rust |
| S3-F05 FAST | 模板编译 800ms + 渲染 ~110ms×8（并行） | 已优化（正则缓存） |
| S3-F05 openpyxl | 每表 2.2s openpyxl 往返 | Python openpyxl 固有 |
| 其余 5 个 | 全部 0.2-2.4s | 纯 Rust，无外部进程 |

### 对拍汇总

| 功能 | 对拍口径 | 结果 |
|---|---|---|
| S1-F01 | 逐条 issue | **3602 条一致** |
| S1-F02 | 逐格值 | **505 行 0 差异** |
| S2-F01 | 逐格值 | **1087 行 0 差异** |
| S3-F01 | 逐格值 | **3868 行 0 差异** |
| S3-F02 | 逐格值 | 单表 **0 差异** |
| S3-F03 | 闭环流程 | exit 0 / 两跑一致 / imported 5-0-0 |
| S3-F04 | 逐格值 | **33 行 0 差异** |
| S3-F05 | 逐格值 | **8 表 0 差异**（三方） |

输入安全：**367 个原始输入 SHA-256 全程不变**

---

## 十、失败策略分层

| 层 | 策略 |
|---|---|
| ExcelModifier 后端选择 | 写前拒绝 → 下一个后端；**写后失败 / 验证失败 → fail-closed** |
| S1-F01 重算 | 失败 → 保留 cache-cleared 副本，转人工重算（不冒充成功） |
| S3-F01 数据导入 | 本期/上期问题 → 停止生成，不产出错误结果 |
| S3-F05 FAST | 编译/渲染/校验失败 → 整体 fail-closed，不留半成品 |
| S3-F05 openpyxl | worker 失败 → 该机构失败；openpyxl 缺失 → 显式报错 |
| 平台无能力证据 | Umya 未登记 → 空后端列表 + 明确报错 |
| 未验证部件 | 图表/透视/ActiveX/嵌入对象/外链/VBA → 写前拒绝 |

---

## 十一、代码位置索引

| 内容 | 文件 |
|---|---|
| 路由注册（S1-F01 / S3-F05 / generic） | `audit-core/src/excel_route.rs` |
| S1-F01 副本/重算/审核 | `s1_formula_copy.rs`、`s1_recalculation.rs`、`s1_formula_audit.rs`、`s1_conditional_audit.rs`、`s1_validation_formula_scan.rs`、`s1_history.rs`、`s1_issue_chain.rs` |
| S1-F01 预检/修复 | `s1_template_rules.rs`、`s1_structure_compare.rs`、`excel_open_repair.rs` |
| S1-F02 汇总 | `s1_region_summary.rs`、`s1_summary_merge.rs` |
| S2-F01 核对 | `s2_config_check.rs`、`s2_import.rs`、`s2_period.rs`、`s2_rule_engine.rs`、`s2_external_import.rs`、`s2_external.rs`、`s2_exporter.rs` |
| S3-F01 比较 | `s3_execution_compare.rs` |
| S3-F02 导出 | `s3_explanation_export.rs` |
| S3-F03 回写 | `s3_explanation_feedback.rs` |
| S3-F04 核对 | `s3_current_value_check.rs`、`unit_conversion.rs` |
| S3-F05 FAST 渲染 | `s3_fast_ooxml.rs`、`s3_fast_ooxml_render.rs`、`s3_indicator_split.rs` |
| 编辑后端 | `crates/excel-editor/src/{ooxml_import_backend,openpyxl_backend,umya_backend}.rs` |
| openpyxl worker 协议 | `crates/excel-editor/python/openpyxl_worker.py` |
| 设置项 | `app_settings.rs` |
| 对拍驱动 | `tests/compare/rustaudit_alignment_p1.py`、`probe/verify_s3f5_engine_solo.py` |
