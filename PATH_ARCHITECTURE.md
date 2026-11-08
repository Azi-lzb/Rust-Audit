# Rust-Audit 八功能路径架构表

> 依据：`audit-core/src/excel_route.rs`、`s3_indicator_split.rs`、`s1_formula_copy.rs`、
> `s1_recalculation.rs`、`app_settings.rs`、各功能 service 模块的**当前代码**；
> 耗时取自 `probe/reports/rustaudit_alignment_report.json`（2026-11-07 全量两跑）
> 与 `probe/reports/s3f5_engine_solo_report.json`（2026-11-08 最新实测）。
> 生成时间：2026-11-08。

## 一、总览：八个功能的默认路径 / 兜底 / 对拍 / 耗时

| # | 功能 | 任务 | **默认路径** | **兜底链路** | **对拍基线** | **实测耗时** |
|---|---|---|---|---|---|---|
| 1 | **S1-F01** | 公式副本生成 + 重算 | `OOXML 直改` → `openpyxl worker` → `Umya`（三者按序，写前拒绝才回退） | Umya（唯一 `Passed` 能力证据） | Python 旧 audit（3602 条逐条） | **189.5 / 189.6s** |
| 2 | **S1-F02** | 汇总校验结果导出 | 纯 Rust `excel-writer` 新建工作簿（**不走 ExcelModifier**） | 无（无编辑器依赖） | Python 旧 audit（505 行） | **1.0 / 1.1s** |
| 3 | **S2-F01** | 大集中核对 | 纯 Rust：`calamine` 读 + 内存规则引擎（**不走 ExcelModifier**） | 无 | Python 旧 audit（1087 条） | **0.2 / 0.2s** |
| 4 | **S3-F01** | 执行比较（月报对比） | 纯 Rust：`calamine` 读 CSV/XLS + `excel-writer` 写（**不走 ExcelModifier**） | 无 | Python 旧 audit（3868 行） | **2.3 / 2.4s** |
| 5 | **S3-F02** | 说明文件导出 | 纯 Rust `excel-writer`（**不走 ExcelModifier**） | 无 | Python 旧 audit（单表） | **1.1 / 1.1s** |
| 6 | **S3-F03** | 反馈回写（变动原因） | `OOXML 直改` → `Umya`（generic 路由） | Umya | Python 旧 audit（exit 0 一致） | **0.2 / 0.2s** |
| 7 | **S3-F04** | 本期数值核对 | 纯 Rust：`excel-reader` + `excel-writer`（**不走 ExcelModifier**） | 无 | Python 旧 audit（33 行） | **1.0 / 1.0s** |
| 8 | **S3-F05** | 指标比较拆分（金融表单） | **FAST_OOXML 直出**（Rust 自研渲染器，ZIP 级） | `openpyxl 往返`（`edit_engine=openpyxl` 或需 openpyxl 独立性时） | Python 旧 audit（8 表逐格） | **1.7-2.0s**（FAST）／**4.6-5.0s**（openpyxl 往返） |

> 耗时说明：S1-F01 的 189s 是**端到端**（含 129 本核查表的公式复制 + Excel COM 重算）。
> 其余功能为单次任务实测。所有耗时为 Windows x86_64 / 12 核 / release 构建。

## 二、按路径类型归类

### A. 走 ExcelModifier 编辑管线的（3 个）

```
                     ┌─ OOXML 直改（ZIP 级，最快，写前拒绝不支持操作则回退）
ExcelModifier ──────┼─ openpyxl worker（整簿 openpyxl 语义）
                     └─ UmyaSpreadsheet（唯一 Passed 证据，完整兜底）
```

| 功能 | 注册顺序（=优先级） | 验证器 | 备注 |
|---|---|---|---|
| **S1-F01** | `OOXML` → `openpyxl`（运行时可用时）→ `Umya` | `UmyaXlsxValidator` | 三者都是候选；写前拒绝才回退，写后失败 fail-closed |
| **S3-F03** | `OOXML` → `Umya`（generic） | `UmyaXlsxValidator` | `edit_engine` 只对 S1-F01/S3-F05 生效 |
| **S3-F05** | `FAST 渲染` →（可选）`openpyxl 往返` | FAST 自带轻量校验 + 可选 worker 校验 | **已不走 EditBackend/ExcelModifier 主链路**；`Umya` 明确排除（定义名守卫拒绝） |

### B. 纯 Rust 读写、无编辑器依赖的（5 个）

| 功能 | 读 | 写 | 特点 |
|---|---|---|---|
| **S1-F02** | `excel-reader`（calamine） | `excel-writer` 新建 `Workbook` | 汇总 130/333/45 三段 |
| **S2-F01** | `calamine` | 内存规则引擎（不写 xlsx 业务结果） | 外部文件 label 探测 1..=4 列 |
| **S3-F01** | `calamine`（CSV/.xls/.xlsx） | `excel-writer` | legacy-eight-segment 表达式口径 |
| **S3-F02** | `excel-reader` | `excel-writer` | 表名单表 `(说明)YYYY.MM 说明文件` |
| **S3-F04** | `excel-reader` | `excel-writer` | 33 行核对结果 |

## 三、引擎选择（高级设置）

| 设置项 | 取值 | 影响功能 | 默认 |
|---|---|---|---|
| `calc_engine` | `auto` / `lo`（LibreOffice） | S1-F01 重算段 | **auto**（Windows → 原生 Excel COM；批量另支持 `lo`） |
| `edit_engine` | `auto` / `ooxml` / `openpyxl` | S1-F01、S3-F05 | **auto** |
| `conditionalFormatEvaluator` | `PYTHON`（默认）/ `COM_DISPLAY` / `COM_EVALUATE`（**未实现，显式报错**） | S1-F01 条件格式 | **PYTHON** |
| `s1ConditionalFormatSource` | `source`（默认）/ 修复后副本 | S1-F01 CF 提取 | **source** |
| `conditionalFormatIncludeBlankCells` | 开/关 | S1-F01 CF | **开** |
| `centralTolerance` | 0..1,000,000 元 | S2-F01 / S3-F01 | **100 元** |

## 四、对拍（与 Python 旧 audit 对齐）状态

| 功能 | 对拍方式 | 结果 | 证据 |
|---|---|---|---|
| S1-F01 | 逐条 issue 对比 | **3602 条一致**（3345 公式 + 257 CF） | `rustaudit_alignment_report.json` |
| S1-F02 | 逐格值 | **505 行 0 差异** | 同上 |
| S2-F01 | 逐格值 | **1087 行 0 差异**（729+357+4） | 同上 |
| S3-F01 | 逐格值 | **3868 行 0 差异** | 同上 |
| S3-F02 | 逐格值 | 单表 **0 差异** | 同上 |
| S3-F03 | 闭环流程 | exit 0、两跑一致（imported 5/unmatched 0/conflicts 0） | 同上 + 真实闭环验证 |
| S3-F04 | 逐格值 | **33 行 0 差异** | 同上 |
| S3-F05 | 逐格值 | **8 表 0 差异**（两跑一致）；FAST / openpyxl / Python 基线**三方逐格 0 差异** | `s3f5_engine_solo_report.json` |

**输入安全**：367 个原始输入 SHA-256 全程不变（对齐驱动每次核验）。

## 五、S3-F05 详细路径（本轮重构后）

```
s3-generate-financial-forms
  │
  ├─ 读比较结果（calamine）+ 3.3 配置 + 3.0 警戒配置
  ├─ plan_indicator_comparison_split（内存规划，510ms）
  │
  ├─ edit_engine = auto / ooxml ─────────────────────────────┐
  │     compile_template（一次，890ms）                       │
  │     ├─ 每机构并行（std::thread::scope，8 线程）            │
  │     │    render_org（纯字符串重建 sheetData，~105ms/表）    │
  │     │    write_zip（~20ms/表）                            │
  │     └─ SheetPlan：删表/降级隐藏/definedNames 归一/bookViews │
  │        + 轻量完整性校验                                    │
  │     总 1.7-2.0s                                          │
  │                                                          │
  └─ edit_engine = openpyxl ─────────────────────────────────┘
        同上 FAST 渲染
        └─ 每机构 fast_roundtrip（openpyxl 加载→保存，~2.2s/表）
           包由 openpyxl 自己的序列化器重建（与 FAST 字节实现独立）
        总 4.6-5.0s
```

**已下线**：EditPlan 逐格编辑路径（513-601s，S3-F05 不再使用）。

## 六、失败策略

| 层 | 策略 |
|---|---|
| ExcelModifier 后端选择 | 写前拒绝（`UnsupportedBeforeWrite`）→ 回退下一个；**写后失败或验证失败 → fail-closed，不回退** |
| S3-F05 FAST | 模板编译/渲染/校验任一失败 → 整体 fail-closed，不留半成品 |
| S3-F05 openpyxl 往返 | worker 失败 → 该机构失败；openpyxl 缺失 → 显式报错 |
| 平台无能力证据 | Umya 未登记 → 保持空后端列表，明确报 "没有已登记的 Passed 后端" |
| 未验证部件 | 含图表/透视表/ActiveX/嵌入对象/外部链接/VBA → 写前拒绝 |

## 七、关键代码位置

| 内容 | 文件 |
|---|---|
| 路由注册（S1-F01 / S3-F05 / generic） | `audit-core/src/excel_route.rs` |
| S3-F05 FAST 渲染 | `audit-core/src/s3_fast_ooxml.rs`、`s3_fast_ooxml_render.rs` |
| S3-F05 并行与计划构建 | `audit-core/src/s3_indicator_split.rs` |
| S1-F01 副本与重算 | `audit-core/src/s1_formula_copy.rs`、`s1_recalculation.rs` |
| S3-F03 回写计划 | `audit-core/src/s3_explanation_feedback.rs` |
| 设置项 | `audit-core/src/app_settings.rs` |
| 编辑后端实现 | `crates/excel-editor/src/{ooxml_import_backend,openpyxl_backend,umya_backend}.rs` |
| openpyxl worker 协议 | `crates/excel-editor/python/openpyxl_worker.py`（含 `fast_roundtrip` 模式） |
| 对拍驱动 | `tests/compare/rustaudit_alignment_p1.py`、`probe/verify_s3f5_engine_solo.py` |
