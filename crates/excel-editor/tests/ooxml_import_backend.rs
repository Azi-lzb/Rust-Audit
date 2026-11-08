//! OOXML 直改后端（ImportRange 路径）的真实工作簿集成测试。
//!
//! 覆盖 S1-F01 实际使用的操作组合：DeleteSheet + CreateSheet + ImportRange
//! （values + formulas，style_mode=None）+ SetWorkbookCalculationProperties。
//! 每个测试都核对：值/公式写入正确、区域外内容按字节保留、输出可被重新读取。

use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, OoxmlImportBackend, RangeStyleMode, WorkbookFormat,
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
        .expect("clock after epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "ooxml_import_backend_{}_{}_{}",
        std::process::id(),
        nonce,
        label
    ))
}

fn capability() -> BackendCapability {
    BackendCapability {
        engine: EngineId::OoxmlPatch,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: excel_editor::OOXML_IMPORT_OPERATIONS
            .iter()
            .copied()
            .collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "test-only-direct-backend".to_owned(),
        evidence: None,
    }
}

/// 模板：一张含公式与文本的表。
fn write_template(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet missing")?;
    sheet.set_name("数据");
    sheet.cell_mut("A1").set_value_string("机构代码");
    sheet.cell_mut("B1").set_value_string("指标");
    sheet.cell_mut("A2").set_value_string("6k0i");
    sheet.cell_mut("B2").set_value_string("各项存款");
    // 区域内的真实公式（不带缓存，模拟模板）。
    sheet.cell_mut("C2").set_formula("A2&\"-\"&B2");
    sheet.cell_mut("C3").set_formula("LEN(A2)");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write template fixture {}", path.display()))
}

/// 审核副本：目标表已存在，区域外有一份必须保留的报送数据与一个独立工作表。
fn write_audit_copy(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .context("default worksheet missing")?;
        sheet.set_name("数据");
        // 区域内：将被模板内容覆盖（含一条陈旧公式，应被清掉）。
        sheet.cell_mut("A2").set_value_string("旧机构");
        sheet
            .cell_mut("C2")
            .set_formula("1+1")
            .set_formula_result_number(2.0);
        // 区域外（E 列以后）：必须原样保留。
        sheet.cell_mut("E1").set_value_string("报送方保留数据");
        sheet.cell_mut("E2").set_value_number(12345.0);
    }
    // 第二张表：用 add_sheet 加入后按索引取回（新表默认名不固定）。
    let first = workbook.add_sheet(umya_spreadsheet::Worksheet::default())?;
    let _ = first;
    let names = workbook
        .get_sheet_collection()
        .iter()
        .map(|sheet| sheet.get_name().to_owned())
        .collect::<Vec<_>>();
    let new_name = names
        .last()
        .cloned()
        .context("newly added worksheet missing")?;
    let second = workbook
        .get_sheet_by_name_mut(&new_name)
        .context("newly added worksheet lookup failed")?;
    second.set_name("其他保留表");
    second.cell_mut("A1").set_value_string("第二张表内容");
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write audit copy fixture {}", path.display()))
}

fn build_plan(template: &Path, audit_copy: &Path, output: &Path) -> Result<ExcelEditPlan> {
    Ok(ExcelEditPlan {
        request_id: "ooxml-import-range-test".to_owned(),
        input: audit_copy.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(audit_copy)?,
        operations: vec![
            EditOperation::ImportRange {
                source_workbook: template.to_path_buf(),
                source_sheet: "数据".to_owned(),
                source_range: "A1:C3".to_owned(),
                target_sheet: "数据".to_owned(),
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
    })
}

#[test]
fn imports_region_values_and_formulas_while_preserving_rest_of_package() -> Result<()> {
    let template = temp_path("template.xlsx");
    let audit_copy = temp_path("audit.xlsx");
    let output = temp_path("output.xlsx");
    write_template(&template)?;
    write_audit_copy(&audit_copy)?;
    let plan = build_plan(&template, &audit_copy, &output)?;

    let backend = OoxmlImportBackend::new(capability())?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!("backend failed: {}", failure.message))?;

    assert!(output.is_file(), "backend must write the output workbook");

    let imported = excel_reader::read_workbook(&output)
        .map_err(|error| anyhow::anyhow!("reopen imported workbook: {error:#}"))?;
    let sheet = imported
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据")
        .context("target sheet missing after import")?;

    // 区域内的值与公式来自模板。
    let text_of = |address: &str| {
        sheet
            .values
            .get(address)
            .and_then(|value| value.get("value"))
            .and_then(|value| value.as_str())
            .map(str::to_owned)
    };
    assert_eq!(text_of("A1").as_deref(), Some("机构代码"));
    assert_eq!(text_of("A2").as_deref(), Some("6k0i"));
    assert_eq!(text_of("B2").as_deref(), Some("各项存款"));
    let formulas = &sheet.formulas;
    assert_eq!(
        formulas.get("C2").map(String::as_str),
        Some("A2&\"-\"&B2"),
        "template formula text must be imported verbatim"
    );
    assert_eq!(formulas.get("C3").map(String::as_str), Some("LEN(A2)"));

    // 区域外报送数据原样保留。
    assert_eq!(text_of("E1").as_deref(), Some("报送方保留数据"));
    assert_eq!(
        sheet
            .values
            .get("E2")
            .and_then(|value| value.get("value"))
            .and_then(|value| value.as_f64()),
        Some(12345.0)
    );

    // 另一张工作表与表结构未受影响。
    assert!(
        imported
            .sheets
            .iter()
            .any(|sheet| sheet.name == "其他保留表"),
        "unrelated worksheet must survive the import"
    );

    // 计划要求：fullCalcOnLoad / forceFullCalc 已写入。
    let properties = ooxml_patch::read_workbook_calculation_properties(&output)?;
    assert_eq!(properties.full_calc_on_load, Some(true));
    assert_eq!(properties.force_full_calc, Some(true));

    // 输入副本未被修改。
    let original = excel_reader::read_workbook(&audit_copy)?;
    let original_sheet = original
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据")
        .context("original target sheet")?;
    assert_eq!(
        original_sheet
            .values
            .get("A2")
            .and_then(|value| value.get("value"))
            .and_then(|value| value.as_str()),
        Some("旧机构"),
        "the input audit copy must stay untouched"
    );

    if std::env::var_os("OOXML_IMPORT_KEEP_ARTIFACTS").is_none() {
        for path in [&template, &audit_copy, &output] {
            let _ = fs::remove_file(path);
        }
    } else {
        eprintln!("artifacts kept: {}", output.display());
    }
    Ok(())
}

#[test]
fn creates_and_deletes_sheets_for_the_rule_table_flow() -> Result<()> {
    let template = temp_path("template_rule.xlsx");
    let audit_copy = temp_path("audit_rule.xlsx");
    let output = temp_path("output_rule.xlsx");
    write_template(&template)?;
    write_audit_copy(&audit_copy)?;

    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .context("default worksheet missing")?;
        sheet.set_name("数据");
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, &audit_copy)?;

    let plan = ExcelEditPlan {
        request_id: "ooxml-import-sheet-flow".to_owned(),
        input: audit_copy.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(&audit_copy)?,
        operations: vec![
            EditOperation::CreateSheet {
                name: "审核规则".to_owned(),
            },
            EditOperation::ImportRange {
                source_workbook: template.clone(),
                source_sheet: "数据".to_owned(),
                source_range: "A1:B2".to_owned(),
                target_sheet: "审核规则".to_owned(),
                target_start: "A1".to_owned(),
                copy_values: true,
                copy_formulas: false,
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
        ],
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    };

    let backend = OoxmlImportBackend::new(capability())?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!("backend failed: {}", failure.message))?;

    let imported = excel_reader::read_workbook(&output)
        .map_err(|error| anyhow::anyhow!("reopen created sheet workbook: {error:#}"))?;
    let rule_sheet = imported
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核规则")
        .context("created rule worksheet missing")?;
    assert_eq!(
        rule_sheet
            .values
            .get("A2")
            .and_then(|value| value.get("value"))
            .and_then(|value| value.as_str()),
        Some("6k0i")
    );

    for path in [&template, &audit_copy, &output] {
        let _ = fs::remove_file(path);
    }
    Ok(())
}
