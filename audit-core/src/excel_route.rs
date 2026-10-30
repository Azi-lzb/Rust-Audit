//! Production Excel-Core route registration for Audit business edit batches.
//!
//! The desktop runtime used to construct `ExcelModifier` with an empty backend
//! list, so S1-F01 formula copies, S3-F03 explanation write-backs and S3-F05
//! financial-form edits failed closed ("Excel-Core Passed 后端注册数=0"). This
//! module registers the pure-Rust `UmyaXlsxBackend` with a `Passed` capability
//! whose evidence is the committed verification report produced by
//! `audit-core/tests/production_capability_verification.rs` on this platform.
//!
//! Scope of the declared capability, and why each boundary is where it is:
//!
//! * Operations cover exactly the three business batches above; no other
//!   operation is advertised.
//! * `reject_unverified_parts` refuses any workbook containing charts, pivot
//!   tables, ActiveX, embedded objects, external links, VBA or non-VML drawing
//!   parts before a backend is selected, so those `WorkbookFeature` values are
//!   deliberately not declared. A file that carries them fails closed with an
//!   explicit reason instead of silently losing them.
//! * `supported_features` lists the features the verification harness creates
//!   in synthetic workbooks and proves read back identical after the exact
//!   business batches run through the full `ExcelModifier` path.
//! * Evidence is registered per platform. Platforms without a registration
//!   still have no backend and keep failing closed.

use anyhow::Context;
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditBackend, EditOperationKind,
    EngineId, ExcelModifier, OpenpyxlWorkerConfig, OpenpyxlXlsxBackend, OpenpyxlXlsxValidator,
    PreservationRequirement, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFeature, WorkbookFormat,
};
use std::{collections::BTreeSet, sync::Arc};

#[derive(Clone, Copy)]
struct EvidenceRegistration {
    platform: &'static str,
    sample_id: &'static str,
    report_path: &'static str,
    adapter_version: &'static str,
    engine_version: &'static str,
}

/// Capability evidence is deliberately platform-specific. Add a platform only
/// after the production verification harness has passed there and produced the
/// referenced report.
const EVIDENCE_REGISTRY: &[EvidenceRegistration] = &[
    EvidenceRegistration {
        platform: "windows-x86_64",
        sample_id: "audit-production-capability-verification-20261028",
        report_path: "audit-core/capability_evidence/umya_windows_x86_64_20261028.json",
        adapter_version: "excel-editor 0.1.0 UmyaXlsxBackend",
        engine_version: "umya-spreadsheet 3.1.1",
    },
    EvidenceRegistration {
        platform: "linux-x86_64",
        sample_id: "audit-production-capability-verification-20261028-linux",
        report_path: "audit-core/capability_evidence/umya_linux_x86_64_20261028.json",
        adapter_version: "excel-editor 0.1.0 UmyaXlsxBackend",
        engine_version: "umya-spreadsheet 3.1.1",
    },
];

/// Operation kinds covered by the S1-F01 formula-copy batch, the S3-F03
/// explanation write-back batch and the S3-F05 financial-form edit batch.
pub const PRODUCTION_OPERATIONS: &[EditOperationKind] = &[
    EditOperationKind::CreateSheet,
    EditOperationKind::ImportRange,
    EditOperationKind::SetComment,
    EditOperationKind::SetWorkbookCalculationProperties,
    EditOperationKind::SetCellValue,
    EditOperationKind::DeleteSheet,
    EditOperationKind::DeleteColumn,
    EditOperationKind::SetCellNumberFormat,
    EditOperationKind::SetCellFillColor,
    EditOperationKind::SetRowHidden,
    EditOperationKind::SetSheetVisibility,
];

/// Preservation contracts the verification harness proves for the business
/// batches on synthetic feature-rich workbooks.
pub fn production_preserved() -> BTreeSet<PreservationRequirement> {
    [
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::UnmodifiedPackageParts,
        PreservationRequirement::SheetIdentityAndOrder,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ]
    .into_iter()
    .collect()
}

/// Workbook features the harness proves survive the business batches on
/// synthetic samples. Features blocked upstream by `reject_unverified_parts`
/// (charts, pivot tables, drawings, images, VBA, external links, ActiveX,
/// embedded objects) are intentionally absent.
pub fn production_supported_features() -> BTreeSet<WorkbookFeature> {
    [
        WorkbookFeature::Comments,
        WorkbookFeature::ConditionalFormatting,
        WorkbookFeature::DataValidations,
        WorkbookFeature::DefinedNames,
        WorkbookFeature::Formulas,
        WorkbookFeature::MergedCells,
        WorkbookFeature::Relationships,
        WorkbookFeature::HeadersFooters,
        WorkbookFeature::FreezePanes,
        WorkbookFeature::AutoFilter,
    ]
    .into_iter()
    .collect()
}

/// Runtime platform identifier, matching the desktop routing sites.
pub fn runtime_platform() -> String {
    format!("{}-{}", std::env::consts::OS, std::env::consts::ARCH)
}

/// The `Passed` capability record behind the production route. Returns `None`
/// on platforms without committed evidence; callers then keep an empty backend
/// list so routing fails closed.
pub fn production_capability(platform: &str) -> Option<BackendCapability> {
    let evidence = EVIDENCE_REGISTRY
        .iter()
        .find(|registration| registration.platform == platform)?;
    Some(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations: PRODUCTION_OPERATIONS.iter().copied().collect(),
        supported_features: production_supported_features(),
        preserved: production_preserved(),
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: evidence.sample_id.to_owned(),
            report_path: evidence.report_path.to_owned(),
            adapter_version: evidence.adapter_version.to_owned(),
            engine_version: evidence.engine_version.to_owned(),
            platform: platform.to_owned(),
        }),
    })
}

/// Builds the production Excel modifier for the current platform. On the
/// verified platform it routes XLSX edit batches to the qualified Umya
/// backend; elsewhere the modifier has no backends and every batch fails
/// closed with `NoEligibleBackend`.
pub fn production_modifier(
    feature: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    let platform = runtime_platform();
    let capability = production_capability(&platform);
    let backends = match &capability {
        Some(record) => vec![Arc::new(UmyaXlsxBackend::new(record.clone())?) as Arc<dyn EditBackend>],
        None => Vec::new(),
    };
    let modifier = ExcelModifier::new(platform, backends, Arc::new(UmyaXlsxValidator))
        .with_context(|| format!("初始化 {feature} Excel-Core 生产路由"))?;
    Ok((modifier, capability))
}

/// Registered-backend count for user-facing diagnostics.
pub fn registered_backend_count(platform: &str) -> usize {
    usize::from(production_capability(platform).is_some())
}

/// openpyxl 高级兼容路径的能力证据（构建时嵌入，随发布更新）。证据文件由
/// `tests/production_capability_verification.rs` 的 openpyxl 段生成，逐批次
/// 记录尝试与拒绝原因；`available` 为 false 时该选项保持置灰。
const OPENPYXL_EVIDENCE: &str =
    include_str!("../capability_evidence/openpyxl_windows_x86_64_20261029.json");

/// Reads the embedded openpyxl evidence and reports whether the advanced
/// compatibility option is available on the current platform.
pub fn openpyxl_evidence_status() -> (bool, String) {
    let Ok(document) = serde_json::from_str::<serde_json::Value>(OPENPYXL_EVIDENCE) else {
        return (false, "openpyxl 能力证据文件无法解析".to_owned());
    };
    let available = document
        .get("available")
        .and_then(serde_json::Value::as_bool)
        .unwrap_or(false);
    let evidence_platform = document
        .get("platform")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default();
    if available && evidence_platform != runtime_platform() {
        return (
            false,
            format!(
                "openpyxl 能力证据平台为 {evidence_platform}，与当前平台 {} 不匹配",
                runtime_platform()
            ),
        );
    }
    if available {
        return (true, "openpyxl 能力证据声明可用".to_owned());
    }
    let conclusion = document
        .get("conclusion")
        .and_then(serde_json::Value::as_str)
        .unwrap_or("能力证据未声明可用");
    (false, format!("openpyxl 修改引擎当前不可用：{conclusion}"))
}

fn edit_operation_kind_from_debug(name: &str) -> Option<EditOperationKind> {
    use EditOperationKind::*;
    Some(match name {
        "CreateSheet" => CreateSheet,
        "CreateSheetAt" => CreateSheetAt,
        "ImportRange" => ImportRange,
        "ImportRangeWithFormulaTranslation" => ImportRangeWithFormulaTranslation,
        "SetComment" => SetComment,
        "SetWorkbookCalculationProperties" => SetWorkbookCalculationProperties,
        "SetCellValue" => SetCellValue,
        "DeleteSheet" => DeleteSheet,
        "DeleteColumn" => DeleteColumn,
        "DeleteRows" => DeleteRows,
        "InsertRows" => InsertRows,
        "SetCellNumberFormat" => SetCellNumberFormat,
        "SetCellFillColor" => SetCellFillColor,
        "SetRowHidden" => SetRowHidden,
        "SetSheetVisibility" => SetSheetVisibility,
        "SetFormulaText" => SetFormulaText,
        "SetColumnWidth" => SetColumnWidth,
        "SetRowHeight" => SetRowHeight,
        "SetFreezePanes" => SetFreezePanes,
        "SetAutoFilter" => SetAutoFilter,
        "MoveSheet" => MoveSheet,
        "RenameSheet" => RenameSheet,
        "CopyRange" => CopyRange,
        "CopyRangeWithFormulaTranslation" => CopyRangeWithFormulaTranslation,
        "SetPageSettings" => SetPageSettings,
        "ReplaceHeaderText" => ReplaceHeaderText,
        _ => return None,
    })
}

fn preservation_requirement_from_name(name: &str) -> Option<PreservationRequirement> {
    use PreservationRequirement::*;
    Some(match name {
        "cell_values_and_types" => CellValuesAndTypes,
        "formula_text" => FormulaText,
        "comments_and_positions" => CommentsAndPositions,
        "styles" => Styles,
        "merged_ranges" => MergedRanges,
        "conditional_formatting" => ConditionalFormatting,
        "data_validations" => DataValidations,
        "defined_names" => DefinedNames,
        "row_and_column_dimensions" => RowAndColumnDimensions,
        "page_settings" => PageSettings,
        "header_footer_text" => HeaderFooterText,
        "relationships" => Relationships,
        "unmodified_package_parts" => UnmodifiedPackageParts,
        "sheet_identity_and_order" => SheetIdentityAndOrder,
        "freeze_panes" => FreezePanes,
        "auto_filter" => AutoFilter,
        _ => return None,
    })
}

/// Builds the openpyxl capability from the embedded evidence. Returns `None`
/// while the evidence records the engine as unavailable (today) or when the
/// evidence platform does not match; callers keep failing closed.
fn openpyxl_capability(platform: &str) -> anyhow::Result<Option<BackendCapability>> {
    let (available, _) = openpyxl_evidence_status();
    if !available {
        return Ok(None);
    }
    let document = serde_json::from_str::<serde_json::Value>(OPENPYXL_EVIDENCE)
        .context("解析 openpyxl 能力证据失败")?;
    let mut operations = std::collections::BTreeSet::new();
    for name in document
        .get("operations")
        .and_then(serde_json::Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(serde_json::Value::as_str)
    {
        operations.insert(
            edit_operation_kind_from_debug(name)
                .with_context(|| format!("openpyxl 证据含未知操作：{name}"))?,
        );
    }
    let mut preserved = std::collections::BTreeSet::new();
    for name in document
        .get("preserved")
        .and_then(serde_json::Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(serde_json::Value::as_str)
    {
        preserved.insert(
            preservation_requirement_from_name(name)
                .with_context(|| format!("openpyxl 证据含未知保留契约：{name}"))?,
        );
    }
    let capability = BackendCapability {
        engine: EngineId::OpenpyxlWorker,
        status: CapabilityStatus::Passed,
        format: WorkbookFormat::Xlsx,
        operations,
        supported_features: production_supported_features(),
        preserved,
        platform: platform.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: document
                .get("sample_set_id")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("audit-openpyxl-capability-verification")
                .to_owned(),
            report_path: "audit-core/capability_evidence/openpyxl_windows_x86_64_20261029.json"
                .to_owned(),
            adapter_version: document
                .get("adapter")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("excel-editor 0.1.0 OpenpyxlXlsxBackend")
                .to_owned(),
            engine_version: document
                .get("engine")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("openpyxl 3.1.5 (isolated worker)")
                .to_owned(),
            platform: platform.to_owned(),
        }),
    };
    Ok(Some(capability))
}

/// Production modifier honoring the GUI 高级设置 edit-engine selection.
/// `umya` keeps the verified default route; `openpyxl` only routes when the
/// embedded evidence declares the engine available on this platform, and its
/// batches are validated by the openpyxl worker's own validator.
pub fn production_modifier_with_edit_engine(
    feature: &str,
    edit_engine: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    match edit_engine {
        "umya" | "" => production_modifier(feature),
        "openpyxl" => {
            let platform = runtime_platform();
            let Some(capability) = openpyxl_capability(&platform)? else {
                let (_, reason) = openpyxl_evidence_status();
                anyhow::bail!("{feature} 的 openpyxl 修改引擎不可用：{reason}");
            };
            // Product deployments stage the openpyxl runtime and worker next
            // to the binary (packaging/openpyxl); the resolver marks the
            // registration unsupported when that layout is absent, keeping
            // the router fail-closed instead of silently switching engines.
            let binary = std::env::current_exe().context("定位当前可执行文件以解析 openpyxl 运行时")?;
            let backend = Arc::new(OpenpyxlXlsxBackend::for_product_binary(
                capability,
                binary,
                "open-data-audit",
            )?);
            let modifier = ExcelModifier::new(
                platform,
                vec![backend],
                Arc::new(OpenpyxlXlsxValidator),
            )
            .with_context(|| format!("初始化 {feature} openpyxl 修改路由"))?;
            Ok((modifier, None))
        }
        other => anyhow::bail!("未知修改引擎：{other}"),
    }
}



#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn evidence_registry_is_platform_specific_and_fails_closed_for_unknown_platforms() {
        for platform in ["windows-x86_64", "linux-x86_64"] {
            let capability = production_capability(platform)
                .expect("platform with registered production evidence");
            assert_eq!(capability.platform, platform);
            assert_eq!(
                capability.evidence.as_ref().map(|evidence| evidence.platform.as_str()),
                Some(platform)
            );
            assert_eq!(registered_backend_count(platform), 1);
        }
        assert!(production_capability("macos-aarch64").is_none());
        assert_eq!(registered_backend_count("macos-aarch64"), 0);
    }
}
