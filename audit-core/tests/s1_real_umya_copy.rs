//! Read-only local acceptance probe for the real S1-F1 workbook copy path.
//!
//! Inputs are supplied through environment variables so private workbooks and
//! machine-specific paths never become repository fixtures. The test makes
//! only temporary copies under the system temp directory.

use anyhow::{bail, ensure, Context, Result};
use audit_core::{
    excel_route::{production_modifier, production_modifier_with_edit_engine},
    s1_formula_copy::{
        build_s1_formula_copy_edit_plan, execute_s1_formula_copy_with_manual_recovery,
        S1FormulaCopyOutcome,
    },
    s1_recalculation::recalculate_s1_formula_copy,
    s1_template_rules::load_s1_template_rule_plan,
};
use excel_editor::{open_xlsx, save_xlsx, BackendAttemptOutcome, EditOperation, EngineId};
use excel_reader::{inspect_formula_cache_nodes, read_workbook};
use std::{
    env, fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
#[ignore = "requires local read-only S1-F1 workbooks, mapping configuration, and Excel COM or LibreOffice"]
fn real_s1_f1_copy_preserves_lookup_formulas_references_and_recalculated_values() -> Result<()> {
    let source = required_path("AUDIT_S1_REAL_SOURCE")?;
    let template = required_path("AUDIT_S1_REAL_TEMPLATE")?;
    let external = required_path("AUDIT_S1_REAL_EXTERNAL")?;
    let mapping = required_path("AUDIT_S1_REAL_MAPPING")?;
    let sidecar = required_path("AUDIT_S1_REAL_SIDECAR")?;
    for path in [&source, &template, &external, &mapping, &sidecar] {
        ensure!(
            path.is_file(),
            "acceptance input is not a file: {}",
            path.display()
        );
    }

    let source_before = fs::read(&source)?;
    let template_before = fs::read(&template)?;
    let external_before = fs::read(&external)?;
    let mapping_before = fs::read(&mapping)?;
    let template_rules = load_s1_template_rule_plan(&template, Some(&mapping))?;
    let template_dump = read_workbook(&template).context("read real S1 template formulas")?;
    let temp = TestDirectory::create()?;

    // Recalculate a byte-copy of the external workbook as an independent
    // result baseline. The original reference workbook remains untouched.
    let external_copy = temp.join("external-source-copy.xlsx");
    let external_recalculated = temp.join("external-recalculated.xlsx");
    fs::copy(&external, &external_copy)?;
    let baseline_report =
        recalculate_s1_formula_copy(&sidecar, &external_copy, &external_recalculated, "auto")?;

    let external_dump = read_workbook(&external).context("read real S1 external workbook")?;
    let reference = sheet(&external_dump.sheets, "参照表")?;
    let institution = sheet(&external_dump.sheets, "机构参照表")?;
    let lookup_formulas = (2..=211)
        .map(|row| {
            reference
                .formulas
                .get(&format!("D{row}"))
                .with_context(|| format!("missing real VLOOKUP formula at 参照表!D{row}"))
        })
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        lookup_formulas
            .iter()
            .filter(|formula| formula.to_ascii_uppercase().contains("VLOOKUP"))
            .count()
            >= 200,
        "the selected real sample does not contain the expected ~210 D2:D211 VLOOKUP formulas"
    );
    ensure!(
        lookup_formulas
            .iter()
            .any(|formula| formula.contains("机构参照表") && formula.contains('!')),
        "the selected real VLOOKUP sample does not exercise its cross-sheet 机构参照表 dependency"
    );
    ensure!(
        !institution.values.is_empty() || !institution.formulas.is_empty(),
        "real 机构参照表 has no stored values or formulas"
    );

    let modifier = production_modifier_with_edit_engine("S1-F01", "umya")?.0;
    let output = temp.join("s1-audit-copy.xlsx");
    let outcome = execute_s1_formula_copy_with_manual_recovery(
        &source,
        &template,
        &output,
        Some(&mapping),
        Some(&external),
        Some(&sidecar),
        &modifier,
        "auto",
    )?;
    let (edit, recalculation) = match outcome {
        S1FormulaCopyOutcome::Recalculated(report) => (report.edit, report.recalculation),
        S1FormulaCopyOutcome::WaitingManualRecalc(report) => {
            preserve_s1_copy_for_diagnostic(&report.audit_copy)?;
            bail!(
                "real S1 formula copy needs manual recalculation: {}",
                report.reason
            )
        }
    };
    // 生产路由自 2026-11 起先尝试 OOXML 直改后端（只重写目标工作表部件，
    // 其余 package 部件按字节保留），Umya 保留为完整兜底。两者都必须产出
    // 通过保留契约的结果；此处断言实际引擎是其中之一，并记录是哪条路径。
    ensure!(
        matches!(
            edit.engine,
            EngineId::OoxmlPatch | EngineId::UmyaSpreadsheet
        ),
        "real S1-F1 ImportRange used an unexpected engine: {:?}",
        edit.engine
    );
    ensure!(
        edit.route_decisions
            .iter()
            .any(|decision| decision.engine == edit.engine && decision.eligible),
        "the engine that produced the output must be eligible in the route decisions"
    );
    let succeeded = edit
        .attempts
        .iter()
        .find(|attempt| attempt.engine == edit.engine)
        .map(|attempt| attempt.outcome == BackendAttemptOutcome::Succeeded)
        .unwrap_or(false);
    ensure!(
        succeeded,
        "the producing engine must report a succeeded attempt; attempts={:?}",
        edit.attempts
    );
    eprintln!(
        "real S1-F1 edit engine: {:?} (OOXML 直改优先，Umya 兜底)",
        edit.engine
    );
    ensure!(
        recalculation.engine == baseline_report.engine,
        "baseline and S1 copy used different calculation engines: {:?} vs {:?}",
        baseline_report.engine,
        recalculation.engine
    );

    let audit_dump = read_workbook(&output).context("reopen real S1 audit copy")?;
    let copied_reference = sheet(&audit_dump.sheets, "参照表")?;
    let copied_institution = sheet(&audit_dump.sheets, "机构参照表")?;
    let baseline_dump =
        read_workbook(&external_recalculated).context("reopen recalculated external baseline")?;
    let baseline_reference = sheet(&baseline_dump.sheets, "参照表")?;
    let baseline_institution = sheet(&baseline_dump.sheets, "机构参照表")?;
    ensure!(
        copied_reference.formulas == reference.formulas,
        "the S1-F01 edit engine changed formulas on the copied 参照表 worksheet"
    );
    ensure!(
        copied_institution.formulas == institution.formulas,
        "the S1-F01 edit engine changed formulas on the copied 机构参照表 worksheet"
    );
    ensure!(
        copied_institution.values == baseline_institution.values,
        "recalculated 机构参照表 values differ from the independently recalculated external workbook"
    );
    for row in 2..=211 {
        let address = format!("D{row}");
        ensure!(
            copied_reference.formulas.get(&address) == reference.formulas.get(&address),
            "formula changed at 参照表!{address}"
        );
        ensure!(
            copied_reference.values.get(&address) == baseline_reference.values.get(&address),
            "Excel/LibreOffice result differs from recalculated external baseline at 参照表!{address}: {:?} vs {:?}",
            copied_reference.values.get(&address),
            baseline_reference.values.get(&address)
        );
    }

    let mut formula_cells = std::collections::BTreeSet::new();
    for rule in &template_rules.definition.rules {
        let identity = (
            rule.sheet_name.clone(),
            rule.formula_cell.to_ascii_uppercase(),
        );
        if !formula_cells.insert(identity.clone()) {
            continue;
        }
        let expected = sheet(&template_dump.sheets, &identity.0)?
            .formulas
            .get(&identity.1)
            .with_context(|| {
                format!("template formula missing at {}!{}", identity.0, identity.1)
            })?;
        let actual = sheet(&audit_dump.sheets, &identity.0)?
            .formulas
            .get(&identity.1)
            .with_context(|| format!("copied formula missing at {}!{}", identity.0, identity.1))?;
        ensure!(
            actual == expected,
            "the S1-F01 edit engine changed template audit formula at {}!{}",
            identity.0,
            identity.1
        );
    }
    ensure!(
        !formula_cells.is_empty(),
        "real S1 template named 校验区域 contains no formula cells"
    );

    if let Some(baseline_path) = env::var_os("AUDIT_S1_REAL_PYTHON_BASELINE").map(PathBuf::from) {
        ensure!(
            baseline_path.is_file(),
            "Python baseline workbook is not a file"
        );
        let baseline_dump = read_workbook(&baseline_path)
            .context("read old Python S1-F1 result workbook as a read-only baseline")?;
        let mut compared_formula_results = 0usize;
        let mut compared_nonempty_formula_results = 0usize;
        let mut result_mismatches = Vec::new();
        for (sheet_name, address) in &formula_cells {
            let baseline_sheet = sheet(&baseline_dump.sheets, sheet_name)?;
            let rust_sheet = sheet(&audit_dump.sheets, sheet_name)?;
            let Some(expected_result) = baseline_sheet.values.get(address) else {
                continue;
            };
            compared_formula_results += 1;
            compared_nonempty_formula_results += usize::from(!is_empty_text(expected_result));
            if !equivalent_formula_result(Some(expected_result), rust_sheet.values.get(address)) {
                result_mismatches.push(format!("{sheet_name}!{address}"));
            }
        }
        ensure!(
            compared_formula_results >= 1_000,
            "Python baseline exposed only {compared_formula_results} template formula result cells"
        );
        ensure!(
            result_mismatches.is_empty(),
            "S1-F01/Excel results differ from old Python S1-F1 baseline in {} 校验区域 formula cells: {:?}",
            result_mismatches.len(),
            result_mismatches.iter().take(20).collect::<Vec<_>>()
        );
        eprintln!(
            "real S1 formula result parity: formula_result_cells={compared_formula_results}, nonempty={compared_nonempty_formula_results}, mismatches=0"
        );
    }

    let cache_inventory = inspect_formula_cache_nodes(&output)?;
    ensure!(
        cache_inventory.formula_cell_count == cache_inventory.cached_formula_cell_count
            && cache_inventory.missing_cache_cells.is_empty(),
        "recalculated audit copy has missing formula caches: {:?}",
        cache_inventory.missing_cache_cells.first()
    );
    ensure!(
        fs::read(&source)? == source_before,
        "source workbook changed"
    );
    ensure!(
        fs::read(&template)? == template_before,
        "template workbook changed"
    );
    ensure!(
        fs::read(&external)? == external_before,
        "external workbook changed"
    );
    ensure!(
        fs::read(&mapping)? == mapping_before,
        "mapping configuration changed"
    );
    Ok(())
}

#[test]
#[ignore = "requires a local read-only S1 source workbook and Microsoft Excel COM or LibreOffice"]
fn real_source_survives_umya_roundtrip_with_calculation_when_needed() -> Result<()> {
    let source = required_path("AUDIT_S1_REAL_SOURCE")?;
    ensure!(source.is_file(), "source workbook is not a file");

    let before = read_workbook(&source).context("read source before Umya roundtrip")?;
    let formula_count = before
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    let temp = TestDirectory::create()?;
    let roundtrip = temp.join("umya-source-roundtrip.xlsx");
    let recalculated = temp.join("umya-source-roundtrip-recalculated.xlsx");
    let workbook = open_xlsx(&source).context("open source with Umya")?;
    save_xlsx(&workbook, &roundtrip).context("save Umya source roundtrip")?;
    drop(workbook);

    // Real S1 data inputs can be pure-data workbooks. Do not ask an external
    // formula engine to recalculate a workbook with no formulas; the separate
    // template-copy probe verifies formula calculation on the merged copy.
    let after = if formula_count == 0 {
        read_workbook(&roundtrip).context("reopen pure-data Umya roundtrip")?
    } else {
        let sidecar = required_path("AUDIT_S1_REAL_SIDECAR")?;
        ensure!(sidecar.is_file(), "recalculation sidecar is not a file");
        recalculate_s1_formula_copy(&sidecar, &roundtrip, &recalculated, "auto")
            .context("open and recalculate Umya roundtrip in Excel/LibreOffice")?;
        read_workbook(&recalculated).context("reopen recalculated Umya roundtrip")?
    };
    ensure!(
        before.sheets.len() == after.sheets.len(),
        "Umya roundtrip changed source worksheet count"
    );
    for (source_sheet, output_sheet) in before.sheets.iter().zip(&after.sheets) {
        ensure!(
            source_sheet.name == output_sheet.name,
            "Umya roundtrip changed source worksheet order or names"
        );
        ensure!(
            source_sheet.formulas == output_sheet.formulas,
            "Umya roundtrip changed formulas in {:?}",
            source_sheet.name
        );
        let source_values = source_sheet
            .values
            .iter()
            .filter(|(address, value)| {
                !source_sheet.formulas.contains_key(*address) && !is_empty_text(value)
            })
            .collect::<Vec<_>>();
        let output_values = output_sheet
            .values
            .iter()
            .filter(|(address, value)| {
                !output_sheet.formulas.contains_key(*address) && !is_empty_text(value)
            })
            .collect::<Vec<_>>();
        let empty_text_omitted = source_sheet
            .values
            .iter()
            .filter(|(address, value)| {
                !source_sheet.formulas.contains_key(*address)
                    && is_empty_text(value)
                    && !output_sheet.values.contains_key(*address)
            })
            .count();
        let empty_text_added = output_sheet
            .values
            .iter()
            .filter(|(address, value)| {
                !output_sheet.formulas.contains_key(*address)
                    && is_empty_text(value)
                    && !source_sheet.values.contains_key(*address)
            })
            .count();
        let value_addresses = source_sheet
            .values
            .keys()
            .chain(output_sheet.values.keys())
            .collect::<std::collections::BTreeSet<_>>();
        let value_differences = value_addresses
            .into_iter()
            .filter(|address| {
                !source_sheet.formulas.contains_key(*address)
                    && !equivalent_cell_content(
                        source_sheet.values.get(*address),
                        output_sheet.values.get(*address),
                    )
            })
            .take(20)
            .map(|address| {
                let cell_type = |value: Option<&serde_json::Value>| -> String {
                    let Some(value) = value else {
                        return "missing".to_owned();
                    };
                    let kind = value
                        .get("type")
                        .and_then(serde_json::Value::as_str)
                        .unwrap_or("unknown");
                    if kind == "string" {
                        let length = value
                            .get("value")
                            .and_then(serde_json::Value::as_str)
                            .map(|text| text.chars().count())
                            .unwrap_or_default();
                        format!("string-{length}-chars")
                    } else {
                        kind.to_owned()
                    }
                };
                format!(
                    "{}:{}->{}",
                    address,
                    cell_type(source_sheet.values.get(address)),
                    cell_type(output_sheet.values.get(address))
                )
            })
            .collect::<Vec<_>>();
        ensure!(
            source_values == output_values,
            "Umya roundtrip changed non-formula values in {:?} ({} input values, {} output values; type-only differences {:?})",
            source_sheet.name,
            source_values.len(),
            output_values.len(),
            value_differences
        );
        if empty_text_omitted > 0 || empty_text_added > 0 {
            eprintln!(
                "Umya empty-text cell normalization in {:?}: omitted={empty_text_omitted}, added={empty_text_added}; non-empty values unchanged",
                source_sheet.name
            );
        }
    }
    Ok(())
}

#[test]
#[ignore = "requires local read-only S1 workbooks and Microsoft Excel COM"]
#[cfg(windows)]
fn real_s1_external_import_and_template_copy_components_open_in_excel() -> Result<()> {
    let source = required_path("AUDIT_S1_REAL_SOURCE")?;
    let template = required_path("AUDIT_S1_REAL_TEMPLATE")?;
    let external = required_path("AUDIT_S1_REAL_EXTERNAL")?;
    let mapping = required_path("AUDIT_S1_REAL_MAPPING")?;
    let temp = TestDirectory::create()?;
    let plan = build_s1_formula_copy_edit_plan(
        &source,
        &template,
        temp.join("plan-placeholder.xlsx"),
        Some(&mapping),
        Some(&external),
    )?;
    let external_sheets = plan
        .external_plan
        .as_ref()
        .context("real template did not select any external worksheets")?
        .sheets
        .iter()
        .map(|sheet| sheet.sheet_name.clone())
        .collect::<std::collections::BTreeSet<_>>();
    let source_template = fs::canonicalize(&template)?;
    let source_external = fs::canonicalize(&external)?;
    let modifier = production_modifier("S1-F01 real Umya component open test")?.0;

    let mut external_only = plan.edit_plan.clone();
    external_only.request_id = "audit-s1-external-only-open-probe".to_owned();
    external_only.output = temp.join("external-only.xlsx");
    external_only
        .operations
        .retain(|operation| match operation {
            EditOperation::CreateSheet { name } => external_sheets.contains(name),
            EditOperation::ImportRange {
                source_workbook, ..
            } => source_workbook == &source_external,
            EditOperation::SetWorkbookCalculationProperties { .. } => true,
            _ => false,
        });
    modifier.execute(&external_only)?;
    assert_excel_com_opens(&external_only.output)?;

    let mut template_ranges_only = plan.edit_plan.clone();
    template_ranges_only.request_id = "audit-s1-template-ranges-only-open-probe".to_owned();
    template_ranges_only.output = temp.join("template-ranges-only.xlsx");
    template_ranges_only
        .operations
        .retain(|operation| match operation {
            EditOperation::ImportRange {
                source_workbook, ..
            } => source_workbook == &source_template,
            EditOperation::SetWorkbookCalculationProperties { .. } => true,
            _ => false,
        });
    modifier.execute(&template_ranges_only)?;
    assert_excel_com_opens(&template_ranges_only.output)?;

    let mut full_copy_without_template_comments = plan.edit_plan;
    full_copy_without_template_comments.request_id =
        "audit-s1-full-copy-without-template-comment-rewrite".to_owned();
    full_copy_without_template_comments.output =
        temp.join("full-copy-without-template-comments.xlsx");
    ensure!(
        !full_copy_without_template_comments
            .operations
            .iter()
            .any(|operation| matches!(operation, EditOperation::SetComment { .. })),
        "S1 formula-copy plan must not rewrite template comments into the audit copy"
    );
    modifier.execute(&full_copy_without_template_comments)?;
    assert_excel_com_opens(&full_copy_without_template_comments.output)?;
    Ok(())
}

#[cfg(windows)]
fn assert_excel_com_opens(path: &Path) -> Result<()> {
    let quoted_path = path.to_string_lossy().replace('\'', "''");
    let script = format!(
        "$ErrorActionPreference='Stop'; $app=$null; $book=$null; try {{ $app=New-Object -ComObject Excel.Application; $app.Visible=$false; $app.DisplayAlerts=$false; $book=$app.Workbooks.Open('{}',0,$true); Write-Output ('opened '+$book.Worksheets.Count+' sheets') }} finally {{ if ($null -ne $book) {{ $book.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}",
        quoted_path
    );
    let result = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
        ])
        .arg(script)
        .output()
        .context("start Excel COM workbook-open probe")?;
    ensure!(
        result.status.success(),
        "Excel COM could not open {}: {}",
        path.display(),
        String::from_utf8_lossy(&result.stderr).trim()
    );
    Ok(())
}

fn required_path(name: &str) -> Result<PathBuf> {
    env::var_os(name)
        .map(PathBuf::from)
        .with_context(|| format!("{name} must point to a local read-only acceptance input"))
}

fn sheet<'a>(
    sheets: &'a [excel_reader::SheetDump],
    name: &str,
) -> Result<&'a excel_reader::SheetDump> {
    sheets
        .iter()
        .find(|sheet| sheet.name == name)
        .with_context(|| format!("missing worksheet {name:?}"))
}

fn preserve_s1_copy_for_diagnostic(output: &Path) -> Result<()> {
    let Some(destination) = env::var_os("AUDIT_S1_REAL_CAPTURE_COPY") else {
        return Ok(());
    };
    ensure!(output.is_file(), "manual-recalculation copy is missing");
    let destination = PathBuf::from(destination);
    if let Some(parent) = destination.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::copy(output, &destination).with_context(|| {
        format!(
            "preserve the isolated real S1 copy for engine diagnostics at {}",
            destination.display()
        )
    })?;
    eprintln!(
        "captured isolated real S1 manual-recalculation copy at {}",
        destination.display()
    );
    Ok(())
}

fn is_empty_text(value: &serde_json::Value) -> bool {
    value.get("type").and_then(serde_json::Value::as_str) == Some("string")
        && value.get("value").and_then(serde_json::Value::as_str) == Some("")
}

fn equivalent_cell_content(
    expected: Option<&serde_json::Value>,
    actual: Option<&serde_json::Value>,
) -> bool {
    match (expected, actual) {
        (Some(expected), None) if is_empty_text(expected) => true,
        (None, Some(actual)) if is_empty_text(actual) => true,
        _ => expected == actual,
    }
}

fn equivalent_formula_result(
    expected: Option<&serde_json::Value>,
    actual: Option<&serde_json::Value>,
) -> bool {
    let Some(expected) = expected else {
        return actual.is_none();
    };
    let Some(actual) = actual else {
        return false;
    };
    let numeric = |value: &serde_json::Value| {
        matches!(
            value.get("type").and_then(serde_json::Value::as_str),
            Some("integer" | "number")
        )
        .then(|| value.get("value").and_then(serde_json::Value::as_f64))
        .flatten()
    };
    match (numeric(expected), numeric(actual)) {
        (Some(expected), Some(actual)) => expected == actual,
        _ => expected == actual,
    }
}

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn create() -> Result<Self> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = env::temp_dir().join(format!(
            "audit_s1_real_umya_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir(&path)?;
        Ok(Self(path))
    }
}

impl std::ops::Deref for TestDirectory {
    type Target = Path;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        if env::var_os("AUDIT_KEEP_S1_REAL_TEST_FILES").is_some() {
            eprintln!(
                "preserving local S1 real-test files at {}",
                self.0.display()
            );
            return;
        }
        let _ = fs::remove_dir_all(&self.0);
    }
}
