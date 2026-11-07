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
    EngineId, ExcelModifier, OOXML_IMPORT_OPERATIONS, OoxmlImportBackend, OpenpyxlXlsxBackend,
    OpenpyxlXlsxValidator, PreservationRequirement, UmyaXlsxBackend, UmyaXlsxValidator,
    WorkbookFeature, WorkbookFormat,
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

/// OOXML 直改后端的能力记录。它只登记实现已覆盖且由集成测试验证的操作；
/// 未覆盖的操作（样式复制、DeleteColumn 等）在写入前被拒绝，由路由继续尝试
/// 后续的 Umya 后端，因此 umya 仍是完整兜底路径。
pub fn ooxml_import_capability(platform: &str) -> BackendCapability {
    BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: OOXML_IMPORT_OPERATIONS.iter().copied().collect(),
        supported_features: production_supported_features(),
        preserved: production_preserved(),
        platform: platform.to_owned(),
        evidence: None,
    }
}

/// Builds the production Excel modifier for the current platform. S1-F01 uses
/// the configured default order (OOXML, available openpyxl, Umya); other
/// features retain the existing OOXML/Umya route.
pub fn production_modifier(
    feature: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    if feature == "S1-F01" {
        return s1_f01_modifier(feature, "auto");
    }
    production_modifier_generic(feature)
}

fn production_modifier_generic(
    feature: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    let platform = runtime_platform();
    let capability = production_capability(&platform);
    // 顺序即优先级：先尝试 OOXML 直改（只重写目标工作表部件，其余字节原样，
    // 不会像整簿重写那样丢失关系/列尺寸等对象）；它不支持的操作在写前拒绝，
    // 路由随后从原始输入重新执行已验收的 Umya 后端。umya 保留为完整兜底。
    let mut backends: Vec<Arc<dyn EditBackend>> = Vec::new();
    let ooxml_capability = ooxml_import_capability(&platform);
    if let Ok(backend) = OoxmlImportBackend::new(ooxml_capability) {
        backends.push(Arc::new(backend));
    }
    if let Some(record) = &capability {
        backends.push(Arc::new(UmyaXlsxBackend::new(record.clone())?));
    }
    let modifier = ExcelModifier::new(platform, backends, Arc::new(UmyaXlsxValidator))
        .with_context(|| format!("初始化 {feature} Excel-Core 生产路由"))?;
    Ok((modifier, capability))
}

/// Builds an openpyxl backend only for S1-F01. The worker remains a candidate
/// path; every resulting workbook must pass the same independent validator.
fn s1_f01_openpyxl_backend(platform: &str, feature: &str) -> anyhow::Result<Arc<dyn EditBackend>> {
    let capability = openpyxl_capability(platform, feature)?;
    let binary = std::env::current_exe().context("定位当前可执行文件以解析 openpyxl 运行时")?;
    let backend = OpenpyxlXlsxBackend::for_product_binary(capability, binary, "open-data-audit")?;
    Ok(Arc::new(backend))
}

/// S1-F01 backend choices. Auto orders candidates OOXML → openpyxl (only when
/// its interpreter, worker and module are available) → Umya. An automatic
/// fallback occurs only after an earlier backend explicitly rejects the batch
/// before writing; backend failures after writing or failed validation stop
/// the run. Each attempt still starts from an untouched input copy.
fn s1_f01_modifier(
    feature: &str,
    selection: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    let platform = runtime_platform();
    let Some(umya_capability) = production_capability(&platform) else {
        // Keep the existing fail-closed behavior on platforms without Umya
        // production evidence; the OOXML trial does not broaden platform scope.
        return production_modifier_generic(feature);
    };

    let mut backends: Vec<Arc<dyn EditBackend>> = Vec::new();
    let mut unverified = Vec::new();
    match selection {
        "auto" | "" => {
            backends.push(Arc::new(OoxmlImportBackend::new(ooxml_import_capability(
                &platform,
            ))?));
            unverified.push(EngineId::OoxmlPatch);
            if crate::app_settings::openpyxl_selection_status().0 {
                backends.push(s1_f01_openpyxl_backend(&platform, feature)?);
                unverified.push(EngineId::OpenpyxlWorker);
            }
            backends.push(Arc::new(UmyaXlsxBackend::new(umya_capability.clone())?));
        }
        "ooxml" => {
            backends.push(Arc::new(OoxmlImportBackend::new(ooxml_import_capability(
                &platform,
            ))?));
            unverified.push(EngineId::OoxmlPatch);
        }
        "openpyxl" => {
            backends.push(s1_f01_openpyxl_backend(&platform, feature)?);
            unverified.push(EngineId::OpenpyxlWorker);
        }
        "umya" => backends.push(Arc::new(UmyaXlsxBackend::new(umya_capability.clone())?)),
        other => anyhow::bail!("未知 S1-F01 修改引擎：{other}"),
    }
    let mut modifier = ExcelModifier::new(platform, backends, Arc::new(OpenpyxlXlsxValidator))
        .with_context(|| format!("初始化 {feature} Excel-Core 生产路由"))?;
    for engine in unverified {
        modifier = modifier.with_unverified_engine_opt_in(engine);
    }
    Ok((modifier, Some(umya_capability)))
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

/// Reads the embedded openpyxl evidence and reports whether its production
/// capability verification passed on the current platform.
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

/// Builds a verified capability when the embedded evidence passed. Otherwise,
/// returns an unverified capability for explicit trial selection; the router
/// requires that opt-in and still validates every candidate before publish.
fn openpyxl_capability(platform: &str, feature: &str) -> anyhow::Result<BackendCapability> {
    let (available, _) = openpyxl_evidence_status();
    if !available {
        // S1-F01 may explicitly trial the isolated worker against its real
        // ImportRange batch. This is not a general Openpyxl capability grant:
        // the candidate remains Unverified, is opt-in only, and must pass the
        // same post-write validator before ExcelModifier publishes it.
        if feature == "S1-F01" {
            return Ok(BackendCapability {
                engine: EngineId::OpenpyxlWorker,
                status: CapabilityStatus::Unverified,
                format: WorkbookFormat::Xlsx,
                operations: [
                    EditOperationKind::CreateSheet,
                    EditOperationKind::ImportRange,
                    EditOperationKind::SetComment,
                    EditOperationKind::SetWorkbookCalculationProperties,
                    EditOperationKind::DeleteSheet,
                ]
                .into_iter()
                .collect(),
                supported_features: production_supported_features(),
                preserved: [
                    PreservationRequirement::CellValuesAndTypes,
                    PreservationRequirement::FormulaText,
                    PreservationRequirement::CommentsAndPositions,
                    PreservationRequirement::MergedRanges,
                    PreservationRequirement::ConditionalFormatting,
                    PreservationRequirement::DataValidations,
                    PreservationRequirement::DefinedNames,
                    PreservationRequirement::Relationships,
                    PreservationRequirement::SheetIdentityAndOrder,
                ]
                .into_iter()
                .collect(),
                platform: platform.to_owned(),
                evidence: None,
            });
        }
        return Ok(BackendCapability {
            engine: EngineId::OpenpyxlWorker,
            status: CapabilityStatus::Unverified,
            format: WorkbookFormat::Xlsx,
            operations: BTreeSet::new(),
            supported_features: production_supported_features(),
            preserved: BTreeSet::new(),
            platform: platform.to_owned(),
            evidence: None,
        });
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
    Ok(capability)
}

/// Production modifier honoring the GUI 高级设置 edit-engine selection.
/// The selection controls only S1-F01; other features retain their existing
/// production routes regardless of this preference.
pub fn production_modifier_with_edit_engine(
    feature: &str,
    edit_engine: &str,
) -> anyhow::Result<(ExcelModifier, Option<BackendCapability>)> {
    if feature != "S1-F01" {
        return production_modifier_generic(feature);
    }
    s1_f01_modifier(feature, edit_engine)
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{EditOperation, ExcelEditPlan, WorkbookProfile};

    fn minimal_s1_f01_plan() -> ExcelEditPlan {
        ExcelEditPlan {
            request_id: "s1-f01-route-test".to_owned(),
            input: std::env::temp_dir().join("s1-f01-route-input.xlsx"),
            output: std::env::temp_dir().join("s1-f01-route-output.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: BTreeSet::new(),
            },
            operations: vec![EditOperation::CreateSheet {
                name: "审核规则".to_owned(),
            }],
            preserve: BTreeSet::new(),
            drop_allowed: Default::default(),
        }
    }

    #[test]
    fn evidence_registry_is_platform_specific_and_fails_closed_for_unknown_platforms() {
        for platform in ["windows-x86_64", "linux-x86_64"] {
            let capability = production_capability(platform)
                .expect("platform with registered production evidence");
            assert_eq!(capability.platform, platform);
            assert_eq!(
                capability
                    .evidence
                    .as_ref()
                    .map(|evidence| evidence.platform.as_str()),
                Some(platform)
            );
            assert_eq!(registered_backend_count(platform), 1);
        }
        assert!(production_capability("macos-aarch64").is_none());
        assert_eq!(registered_backend_count("macos-aarch64"), 0);
    }

    #[test]
    fn unverified_openpyxl_trial_is_scoped_to_s1_f01_import_batches() {
        let candidate = openpyxl_capability("windows-x86_64", "S1-F01")
            .expect("construct S1-F01 Openpyxl candidate");
        assert_eq!(candidate.status, CapabilityStatus::Unverified);
        assert!(
            candidate
                .operations
                .contains(&EditOperationKind::ImportRange)
        );
        assert!(
            candidate
                .preserved
                .contains(&PreservationRequirement::Relationships)
        );

        let other = openpyxl_capability("windows-x86_64", "S3-F03")
            .expect("construct non-S1 Openpyxl capability");
        assert!(other.operations.is_empty());
    }

    #[test]
    fn s1_f01_auto_route_orders_ooxml_openpyxl_umya() {
        let (modifier, capability) = production_modifier_with_edit_engine("S1-F01", "auto")
            .expect("construct S1-F01 automatic edit route");
        assert_eq!(
            capability.as_ref().map(|value| value.engine),
            Some(EngineId::UmyaSpreadsheet)
        );
        let decisions = modifier.route(&minimal_s1_f01_plan());
        let expected = if crate::app_settings::openpyxl_selection_status().0 {
            vec![
                EngineId::OoxmlPatch,
                EngineId::OpenpyxlWorker,
                EngineId::UmyaSpreadsheet,
            ]
        } else {
            vec![EngineId::OoxmlPatch, EngineId::UmyaSpreadsheet]
        };
        assert_eq!(
            decisions
                .iter()
                .map(|decision| decision.engine)
                .collect::<Vec<_>>(),
            expected
        );
        assert!(decisions.iter().all(|decision| decision.eligible));
    }

    #[test]
    fn explicit_s1_f01_engine_selection_registers_only_that_backend() {
        for (selection, expected) in [
            ("ooxml", EngineId::OoxmlPatch),
            ("openpyxl", EngineId::OpenpyxlWorker),
            ("umya", EngineId::UmyaSpreadsheet),
        ] {
            let (modifier, _) = production_modifier_with_edit_engine("S1-F01", selection)
                .expect("construct explicit S1-F01 edit route");
            let decisions = modifier.route(&minimal_s1_f01_plan());
            assert_eq!(decisions.len(), 1, "selection={selection}");
            assert_eq!(decisions[0].engine, expected, "selection={selection}");
        }
    }

    #[test]
    fn edit_engine_selection_does_not_change_s3_routes() {
        for selection in ["auto", "ooxml", "openpyxl", "umya"] {
            let (modifier, _) = production_modifier_with_edit_engine("S3-F03", selection)
                .expect("S3 route ignores the S1-F01-only preference");
            let engines = modifier
                .route(&minimal_s1_f01_plan())
                .into_iter()
                .map(|decision| decision.engine)
                .collect::<Vec<_>>();
            assert_eq!(engines, [EngineId::OoxmlPatch, EngineId::UmyaSpreadsheet]);
        }
    }
}
