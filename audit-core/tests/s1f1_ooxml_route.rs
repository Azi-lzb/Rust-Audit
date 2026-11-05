//! 验证 S1-F01 的真实编辑计划组合能被 OOXML 直改后端接受并产出正确结果，
//! 而不是回落到 Umya。这个测试直接构造 S1-F01 的完整操作序列：
//! DeleteSheet → CreateSheet → ImportRange(values+formulas) →
//! SetWorkbookCalculationProperties，并用真实 XLSX 工作簿跑通。

use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendAttemptOutcome, BackendCapability, CapabilityStatus,
    EditBackend, EditOperation, EditOperationKind, EngineId, ExcelEditPlan, OoxmlImportBackend,
    RangeStyleMode, WorkbookFormat, OOXML_IMPORT_OPERATIONS,
};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

fn temp_path(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "s1f1_ooxml_route_{}_{}_{}",
        std::process::id(),
        nonce,
        label
    ))
}

fn write_template(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet")?;
    sheet.set_name("审核规则");
    sheet.cell_mut("A1").set_value_string("规则编号");
    sheet.cell_mut("B1").set_value_string("说明");
    sheet.cell_mut("A2").set_value_string("R001");
    sheet.cell_mut("B2").set_value_string("各项存款不得为空");
    sheet.cell_mut("C2").set_formula("LEN(A2)");
    sheet.cell_mut("C3").set_formula(r#"B2&"!""#);
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

fn write_audit_copy(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .context("default worksheet")?;
        sheet.set_name("数据");
        sheet.cell_mut("A1").set_value_string("报送数据");
        sheet.cell_mut("D5").set_value_string("区域外必须保留");
    }
    // 已存在的旧“审核规则”表：S1 会先删再建。
    workbook.add_sheet(umya_spreadsheet::Worksheet::default())?;
    let names = workbook
        .get_sheet_collection()
        .iter()
        .map(|sheet| sheet.get_name().to_owned())
        .collect::<Vec<_>>();
    let new_name = names.last().cloned().context("added sheet")?;
    let second = workbook
        .get_sheet_by_name_mut(&new_name)
        .context("added sheet lookup")?;
    second.set_name("审核规则");
    second.cell_mut("A1").set_value_string("陈旧规则");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

#[test]
fn s1f1_full_operation_sequence_is_accepted_and_correct() -> Result<()> {
    let template = temp_path("template.xlsx");
    let audit_copy = temp_path("audit.xlsx");
    let output = temp_path("output.xlsx");
    write_template(&template)?;
    write_audit_copy(&audit_copy)?;

    let capability = BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: OOXML_IMPORT_OPERATIONS.iter().copied().collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "test".to_owned(),
        evidence: None,
    };
    let backend = OoxmlImportBackend::new(capability.clone())?;

    let plan = ExcelEditPlan {
        request_id: "s1f1-full-sequence".to_owned(),
        input: audit_copy.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(&audit_copy)?,
        operations: vec![
            EditOperation::DeleteSheet {
                name: "审核规则".to_owned(),
            },
            EditOperation::CreateSheet {
                name: "审核规则".to_owned(),
            },
            EditOperation::ImportRange {
                source_workbook: template.clone(),
                source_sheet: "审核规则".to_owned(),
                source_range: "A1:C3".to_owned(),
                target_sheet: "审核规则".to_owned(),
                target_start: "A1".to_owned(),
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::None,
                copy_comments: false,
                copy_merges: false,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: true,
                hide_zero: false,
            },
            EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            },
        ],
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    };

    // 1) 后端必须接受这个计划（不是靠回落）。
    backend.apply(&plan).map_err(|failure| {
        anyhow::anyhow!(
            "S1-F01 计划被 OOXML 后端拒绝（会回落到 Umya）：{}",
            failure.message
        )
    })?;
    assert!(output.is_file(), "计划被接受时必须产出工作簿");

    // 2) 结果正确：旧规则表被删除重建，新内容来自模板。
    let dump = excel_reader::read_workbook(&output).context("reopen output")?;
    let rule_sheets: Vec<_> = dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name == "审核规则")
        .collect();
    assert_eq!(rule_sheets.len(), 1, "deleted sheet must not be duplicated");
    let rule = rule_sheets[0];
    let value = |address: &str| {
        rule.values
            .get(address)
            .and_then(|entry| entry.get("value"))
            .and_then(|entry| entry.as_str())
            .map(str::to_owned)
    };
    assert_eq!(value("A1").as_deref(), Some("规则编号"));
    assert_eq!(value("A2").as_deref(), Some("R001"));
    assert_eq!(
        value("A1").as_deref(),
        Some("规则编号"),
        "stale rule content must be replaced, not merged"
    );
    assert!(
        !value("A2").map(|text| text.contains("陈旧")).unwrap_or(false),
        "the old rule row must not survive the delete+recreate cycle"
    );

    // 3) 公式文本原样导入。
    assert_eq!(
        rule.formulas.get("C2").map(String::as_str),
        Some("LEN(A2)"),
        "formula text from the template must be preserved"
    );
    assert_eq!(
        rule.formulas.get("C3").map(String::as_str),
        Some(r#"B2&"!""#),
        "formula text with escaped characters must survive round trip"
    );

    // 4) 区域外数据与另一张表保持不变。
    let data = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据")
        .context("data sheet must survive")?;
    assert_eq!(
        data.values
            .get("D5")
            .and_then(|entry| entry.get("value"))
            .and_then(|entry| entry.as_str()),
        Some("区域外必须保留")
    );

    // 5) 重算标记已写入。
    let properties = ooxml_patch::read_workbook_calculation_properties(&output)?;
    assert_eq!(properties.full_calc_on_load, Some(true));
    assert_eq!(properties.force_full_calc, Some(true));

    // 6) 输入副本未被改动。
    let original = excel_reader::read_workbook(&audit_copy)?;
    assert!(
        original.sheets.iter().any(|sheet| {
            sheet.name == "审核规则"
                && sheet
                    .values
                    .get("A1")
                    .and_then(|entry| entry.get("value"))
                    .and_then(|entry| entry.as_str())
                    == Some("陈旧规则")
        }),
        "the input audit copy must keep its original rule sheet"
    );

    // 7) 这个引擎确实是 OOXML，不走 Umya。
    assert_eq!(backend.capability().engine, EngineId::OoxmlPatch);
    let _ = BackendAttemptOutcome::Succeeded;

    for path in [&template, &audit_copy, &output] {
        let _ = fs::remove_file(path);
    }
    Ok(())
}

#[test]
fn style_mode_full_is_rejected_before_write_so_umya_can_handle_it() -> Result<()> {
    let template = temp_path("template_style.xlsx");
    let audit_copy = temp_path("audit_style.xlsx");
    let output = temp_path("output_style.xlsx");
    write_template(&template)?;
    write_audit_copy(&audit_copy)?;

    let capability = BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: OOXML_IMPORT_OPERATIONS.iter().copied().collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "test".to_owned(),
        evidence: None,
    };
    let backend = OoxmlImportBackend::new(capability)?;
    let plan = ExcelEditPlan {
        request_id: "s1f1-style-reject".to_owned(),
        input: audit_copy.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(&audit_copy)?,
        operations: vec![EditOperation::ImportRange {
            source_workbook: template.clone(),
            source_sheet: "审核规则".to_owned(),
            source_range: "A1:C3".to_owned(),
            target_sheet: "审核规则".to_owned(),
            target_start: "A1".to_owned(),
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::Full,
            copy_comments: false,
            copy_merges: true,
            copy_row_heights: true,
            copy_column_widths: true,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: true,
            hide_zero: false,
        }],
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    };

    let failure = backend
        .apply(&plan)
        .expect_err("style_mode=Full must be rejected before write");
    assert_eq!(
        failure.kind,
        excel_editor::BackendErrorKind::UnsupportedBeforeWrite,
        "rejection must happen before any write so the router can fall back to Umya"
    );
    assert!(
        !output.exists(),
        "a pre-write rejection must not leave an output behind"
    );

    for path in [&template, &audit_copy] {
        let _ = fs::remove_file(path);
    }
    Ok(())
}
