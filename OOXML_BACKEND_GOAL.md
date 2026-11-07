# Rust-Audit：以 OOXML 直改后端替换 Umya 编辑引擎（四阶段）

正式工作目录：C:\Users\AZI\Desktop\VibeCode\Rust-Audit
旧 Python 只读参考：C:\Users\AZI\Desktop\VibeCode\3.open-data-audit
分支 main，接手时 HEAD `99ea348`（工作区干净，已 push 到 origin）

## 背景与动机

用户反馈 Umya 编辑引擎"经常丢东西"（工作簿经 umya 整簿 load→save 序列化后，
批注/样式/图表/数据验证/定义名称等对象可能静默丢失）。现有 `UmyaXlsxValidator`
只校验工作表可打开、未知包部件、删列数据、计算属性回读、表可见性/行隐藏/数字格式，
未覆盖上述对象，因此丢失可能在发布后才发现。

旧 Python Audit 的做法印证了这一点：它对最核心的两个写入功能**刻意绕开 openpyxl**
——S1-F01 用 `DIRECT_OOXML`（ZIP+XML 直改，副本其余部件字节原样保留），
S3-F05 默认用 `FAST_OOXML` 整簿渲染，openpyxl 仅作兼容兜底。

目标：新增 OOXML 直改后端，逐阶段替换 umya，最终禁用 umya。
openpyxl 已实测不可用（`audit-core/capability_evidence/openpyxl_windows_x86_64_20261029.json`
记录 S1-F01/S3-F03/S3-F05 三个功能全部失败），本目标不涉及 openpyxl 修复。

## 现状（已核实，勿重复调研）

- `ExcelModifier` 持有 `backends: Vec<Arc<dyn EditBackend>>`，按能力路由，失败关闭不自动换引擎。
  - `crates/excel-editor/src/plan.rs:1046` `trait EditBackend { capability, apply, validate_output }`
  - `crates/excel-editor/src/plan.rs:1143` `struct ExcelModifier`
- 现成 OOXML 后端范例：`crates/excel-editor/src/ooxml_header_backend.rs`
  （`OoxmlHeaderTextBackend` + `OoxmlHeaderTextValidator`，构造时校验 capability.engine == OoxmlPatch）
- `crates/ooxml-patch` 已有可复用能力：`patch_workbook_calculation_properties`、
  `clear_formula_caches_for_recalculation`、`restore_promoted_error_constants`、
  `patch_column_dimensions`、`copy_outline_levels`、`replace_header_text`、
  `patch_external_hyperlink(s)`、`restore_legacy_comments_for_sheet`、
  `read_header_inventory`、`unverified_parts`/`reject_unverified_parts`
- `excel-editor` 依赖已含 `zip`、`roxmltree`、`ooxml-patch`、`excel-reader`
- `EngineId::OoxmlPatch` 已存在（header 后端在用）
- 注册点：`audit-core/src/excel_route.rs`
  - `EVIDENCE_REGISTRY`（windows-x86_64 / linux-x86_64，各指向 capability_evidence JSON）
  - `production_capability(platform)` 构造 `BackendCapability`（status=Passed）
  - `production_modifier(feature)` 组装 backends
  - `production_modifier_with_edit_engine(feature, edit_engine)`：`"umya"|""` → `production_modifier`；
    `"openpyxl"` → 校验 `openpyxl_evidence_status()`
- 三个功能各自的操作面（已核实）：
  - **S1-F01**（`audit-core/src/s1_formula_copy.rs`）：`CreateSheet`、`DeleteSheet`、
    `ImportRange`、`ImportRangeWithFormulaTranslation`、`SetWorkbookCalculationProperties`
  - **S3-F03**（`audit-core/src/s3_explanation_feedback.rs`）：`SetCellValue`
  - **S3-F05**（`audit-core/src/s3_indicator_split.rs`）：`DeleteColumn`、`DeleteSheet`、
    `SetCellFillColor`、`SetCellNumberFormat`、`SetCellValue`、`SetRowHidden`、`SetSheetVisibility`
- 旧 Python 可移植实现（只读参考）：
  - `core/src/base_audit/formula_region_writer.py` — `DirectOoxmlFormulaRegionWriter`
  - `core/src/base_audit/external_sheet_writer.py` — `DIRECT_OOXML` 外部表导入
  - `core/src/base_audit/systems/s3_central_statistics/fast_ooxml.py`
  - `core/src/base_audit/systems/s3_central_statistics/form_template.py`
  - `core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py`（openpyxl 版，仅参考语义）

## 阶段 1：S1-F01 公式复制的 OOXML 直改后端（优先）

新增 `crates/excel-editor/src/ooxml_import_backend.rs`，实现 `EditBackend`：

1. 覆盖操作：`CreateSheet`、`DeleteSheet`、`ImportRange`、`ImportRangeWithFormulaTranslation`、
   `SetWorkbookCalculationProperties`。任一操作不在能力集内 → `BackendFailure::unsupported_before_write`（写前失败）。
2. 移植 `formula_region_writer.py` 的核心算法：
   - 从模板直读区域每格公式/样式/合并，写入副本目标表 `sheetData` 与 `mergeCells`
   - 模板样式合并进副本 `styles.xml`（fills/fonts/borders/numFmts/cellXfs，**同步维护各组 count**）
   - **shared formula 展开**：`<f t="shared" si=..>` 按主格 + `ref` 平移展开为普通公式；
     依赖格写出空 `<f>` 会被 Excel 判损坏，这是已知坑
   - 丢弃 `calcChain`，写 `fullCalcOnLoad`（或由 `SetWorkbookCalculationProperties` 显式控制）
   - **副本其余部件字节原样保留**（这是"不丢东西"的根本）
3. 外部表导入移植 `external_sheet_writer.py` 的 `DIRECT_OOXML` 边界：只写纯值/布尔/文本到新建无样式工作表，
   遇到公式或会影响数据语义的对象**明确失败**，不静默产出不完整工作簿。
   文本统一 `inlineStr`，不迁移 sharedStrings。
4. 新增 `OoxmlImportValidator`：至少校验
   - 候选可被 `excel_reader` 重开、工作表非空
   - `inventory_complete`（未知包部件）
   - **源簿对象清单未减少**：工作表名/顺序、公式总数、合并区域数、批注数、定义名称、
     条件格式规则数、数据验证数、超链接数、表对象数、图表/图片部件数
   - 目标区域写入后的单元格值与公式文本与计划一致
5. 注册顺序：`OoxmlImportBackend` 注册在 umya **之前**，仅当计划操作全部在能力集内才被选中；
   否则回退 umya（阶段 1 期间 umya 仍保留作为兜底）。

## 阶段 2：S3-F03 说明回写

新增覆盖 `SetCellValue` 的 OOXML 直改路径（可扩展阶段 1 的后端，或独立后端）：

1. 只改目标单元格的 `<c>/<v>`（文本用 `inlineStr` 或复用 sharedStrings 的既有索引），
   不触碰样式索引、不重排 sheetData 顺序
2. 保持行列尺寸、冻结、自动筛选、条件格式、数据验证、批注原样
3. 备份语义沿用现有 Rust 实现（写前时间戳完整备份），不得弱化
4. 校验：值写入正确 + 对象清单未减少 + 输入哈希不变

## 阶段 3：S3-F05 金融表单（最重）

覆盖 `DeleteColumn`、`DeleteSheet`、`SetCellFillColor`、`SetCellNumberFormat`、
`SetCellValue`、`SetRowHidden`、`SetSheetVisibility`：

1. `DeleteColumn`：重生成目标表 `sheetData` 列平移 + 同步 `cols` 尺寸、公式引用平移、
   合并区域裁剪/平移；**不得丢列外数据**。参考 umya 后端现有 `validate_deleted_column_payloads` 的检查项
2. 其余操作按 OOXML 直改实现，复用 `patch_column_dimensions`、`copy_outline_levels`
3. 校验：模拟 30 机构 × 40 指标 × 4 表规模的合成样本，逐格/公式/关键结构对拍，
   且对象清单未减少

## 阶段 4：禁用 Umya

1. 三个功能全部通过后，从 `production_modifier` 的 backends 列表移除 `UmyaXlsxBackend`
2. 生成新的 capability evidence JSON（`audit-core/capability_evidence/ooxml_patch_*_2026MMDD.json`），
   记录 windows-x86_64 与 linux-x86_64 两个平台的实际验证结果（harness 路径、样本集 ID、逐操作结果）
3. 更新 `EVIDENCE_REGISTRY`、`production_capability`、`PRODUCTION_OPERATIONS`
4. GUI 高级设置"修改引擎"：`umya` 项改为不可选并注明"已由 OOXML 直改后端替代"；
   `app_settings.rs` 的 `EDIT_ENGINE_VALUES`/`EDIT_ENGINE_LABELS` 同步；
   旧配置里 `edit_engine: "umya"` 需向后兼容（解析为默认新后端，不报错）
5. `UmyaXlsxBackend` 源码保留（测试与历史对照用），但不再进入生产路由

## 硬性约束

- **不删测试、不放宽断言、不把 Unverified 改 Passed、不虚构验证结果**
- 每个阶段独立可交付：阶段 N 完成即全量回归 + 该功能真实合成样本冒烟通过后再进入阶段 N+1；
  **不得一次性改完三个功能再统一验证**
- 失败必须关闭：任何操作不支持 → 写前 `unsupported_before_write`，绝不产出部分改动的文件
- Umya 失败不自动切换（沿用现有规则）；OOXML 后端失败同样不静默换引擎
- 输入文件永不修改；输出走现有 staging + 不覆盖 + 输入哈希校验
- 不引入 pywin32、不改 sidecar 协议、不改 openpyxl worker、不碰 S5
- 不新增外部依赖（`zip`/`roxmltree`/`ooxml-patch` 已在依赖中）
- 每阶段写 GPT修改记录（`audit-core/GPT修改记录/`），矩阵文档只增不改口径
- 允许 commit，禁止 push 与破坏性 Git 操作；提交前全量回归必须 0 failed

## 验收标准（每阶段）

| 阶段 | 必须通过 |
|---|---|
| 1 | S1-F01 合成样本（≥3 源簿，含 shared formula、批注、条件格式、数据验证、定义名称、合并）编辑后：对象清单零减少、公式文本与缓存契约不变、Excel 可打开、输入哈希不变；`cargo test -p excel-editor` + `cargo test -p audit-core --lib` 全绿 |
| 2 | S3-F03 1,000 条反馈导入：值正确、对象清单零减少、备份与原目标 SHA-256 一致、只改目标列 |
| 3 | S3-F05 30 机构 × 40 指标 × 4 表：逐格/公式/结构对拍 0 差异、对象清单零减少、峰值内存记录 |
| 4 | 三功能在 umya 移除后仍全绿；新 evidence JSON 存在且 `available: true`；GUI 引擎下拉正确；`edit_engine: "umya"` 旧配置可正常启动 |
| 全局 | `cargo test --workspace` 0 failed；Windows Release 重建并冒烟（含 GUI 高级设置）；DEVELOPMENT_GOAL / MIGRATION_INVENTORY / 能力证据与实际一致 |

## 交付报告要求

如实列出：
1. 每个阶段实际覆盖的操作与未覆盖的操作
2. "对象清单零减少"的具体检查项清单
3. 各阶段实测数据（耗时、峰值内存、样本规模）
4. 是否所有阶段都完成；未完成部分说明阻碍
5. 编译/测试实际运行情况（跑了什么、结果、未跑什么及原因）
6. 改动文件清单、能力证据路径、Git 仓库与分支、未提交内容
