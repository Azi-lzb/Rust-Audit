//! Audit executable's private LibreOffice recalculation worker.
//! Migrated from the former standalone helper; Office software remains optional.
//! One batch uses one isolated soffice invocation, not a persistent UNO session.
use anyhow::{bail, ensure, Context, Result};
use excel_reader::WorkbookDump;
use serde_json::{json, Value};
use std::{
    collections::BTreeMap,
    env, fs,
    path::{Path, PathBuf},
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};
#[path = "libreoffice_batch.rs"]
mod lo_batch;

pub const WORKER_COMMAND: &str = "__audit-lo-worker";

/// Intercept before normal GUI/CLI parsing. Does not load GUI configuration.
pub fn maybe_run_worker() -> Option<Result<()>> {
    let mut args = env::args().skip(1);
    if args.next().as_deref() != Some(WORKER_COMMAND) {
        return None;
    }
    Some((|| {
        let mode = args.next().context("LibreOffice worker 缺少操作")?;
        let values = parse_options(args.collect())?;
        let engine = values.get("engine").map(String::as_str).unwrap_or("lo");
        ensure!(
            matches!(engine, "lo" | "auto"),
            "LibreOffice worker 只接受 lo/auto"
        );
        let report = match mode.as_str() {
            "recalculate" => {
                let input = PathBuf::from(values.get("input").context("--input is required")?);
                let output = PathBuf::from(values.get("output").context("--output is required")?);
                let mut report = execute_batch(vec![(input, output)])?;
                let mut book = report["books"]
                    .as_array_mut()
                    .context("worker 缺少结果")?
                    .remove(0);
                ensure!(
                    book["status"] == "recalculated",
                    "LibreOffice 重算失败：{}",
                    book["error"]
                );
                book["soffice_conversion_invocations"] = json!(1);
                book
            }
            "recalculate-batch" => {
                let path = PathBuf::from(values.get("pairs").context("--pairs is required")?);
                let content = fs::read_to_string(&path).context("读取批次重算清单")?;
                let mut pairs = Vec::new();
                for line in content.lines().filter(|line| !line.trim().is_empty()) {
                    let mut fields = line.split('\t');
                    let input = fields.next().context("batch pair missing input")?;
                    let output = fields.next().context("batch pair missing output")?;
                    ensure!(fields.next().is_none(), "batch pair has extra fields");
                    pairs.push((PathBuf::from(input), PathBuf::from(output)));
                }
                execute_batch(pairs)?
            }
            _ => bail!("不支持的 LibreOffice worker 操作：{mode}"),
        };
        println!("{}", serde_json::to_string(&report)?);
        Ok(())
    })())
}
fn parse_options(arguments: Vec<String>) -> Result<BTreeMap<String, String>> {
    let mut values = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(
            matches!(
                key.as_str(),
                "--input" | "--output" | "--pairs" | "--engine"
            ),
            "unknown worker option {key}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        ensure!(
            values
                .insert(key.trim_start_matches('-').to_owned(), value)
                .is_none(),
            "duplicate worker option {key}"
        );
    }
    Ok(values)
}

fn execute_batch(declared: Vec<(PathBuf, PathBuf)>) -> Result<Value> {
    ensure!(!declared.is_empty(), "batch pairs file is empty");
    let mut destinations = std::collections::BTreeSet::new();
    for (input, output) in &declared {
        ensure!(
            input.is_file(),
            "batch input workbook does not exist: {}",
            input.display()
        );
        ensure!(
            input
                .extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
                && output
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
            "LibreOffice 重算仅支持 xlsx"
        );
        if let Some(parent) = output.parent().filter(|p| !p.as_os_str().is_empty()) {
            fs::create_dir_all(parent)?;
        }
        let destination = fs::canonicalize(
            output
                .parent()
                .filter(|p| !p.as_os_str().is_empty())
                .unwrap_or(Path::new(".")),
        )?
        .join(output.file_name().context("输出路径缺少文件名")?);
        ensure!(destinations.insert(destination), "批次输出路径重复");
        ensure!(
            !output.exists(),
            "refusing to overwrite existing batch output: {}",
            output.display()
        );
    }

    let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
    let input_dir = guard.0.join("batch-input");
    let output_dir = guard.0.join("batch-output");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&output_dir)?;
    let mut prepared_pairs = Vec::with_capacity(declared.len());
    let mut cleared_counts = Vec::with_capacity(declared.len());
    for (index, (input, _)) in declared.iter().enumerate() {
        let canonical = fs::canonicalize(input)
            .with_context(|| format!("resolve batch input {}", input.display()))?;
        let cacheless = input_dir.join(format!("book-{index:04}-cacheless.xlsx"));
        let prepared = input_dir.join(format!("book-{index:04}.xlsx"));
        let cleared =
            ooxml_patch::clear_formula_caches_for_recalculation(&canonical, &cacheless)
                .with_context(|| format!("clear formula caches for batch book {}", index + 1))?;
        ooxml_patch::patch_workbook_calculation_properties(
            &cacheless,
            &prepared,
            ooxml_patch::WorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            },
        )
        .with_context(|| format!("prepare batch book {}", index + 1))?;
        cleared_counts.push(cleared.cache_values_removed);
        prepared_pairs.push((prepared, output_dir.join(format!("book-{index:04}.xlsx"))));
    }
    let batch_timeout = recalculation_timeout()
        .saturating_mul(declared.len() as u32)
        .max(recalculation_timeout());
    let convert_started = Instant::now();
    let (engine_name, stages) = lo_batch::run_batch_convert(
        &prepared_pairs,
        guard.0.as_path(),
        lo_batch::BatchTarget::Xlsx,
        batch_timeout,
    )
    .context("batch LibreOffice conversion")?;
    let convert_ms = convert_started.elapsed().as_millis() as u64;

    let mut books = Vec::with_capacity(declared.len());
    let mut failures = 0usize;
    for (index, (input, output)) in declared.iter().enumerate() {
        let candidate = stages
            .get(index)
            .and_then(|stage| stage.get("batch_output"))
            .and_then(serde_json::Value::as_str)
            .map(PathBuf::from);
        let Some(candidate) = candidate else {
            failures += 1;
            books.push(serde_json::json!({"index": index + 1, "status": "failed", "error": "batch did not produce output"}));
            continue;
        };
        let attempt = (|| -> Result<serde_json::Value> {
            require_engine_output(&candidate)?;
            let (validated_candidate, restored_error_constants) =
                normalize_promoted_error_constants(input, &candidate, &{
                    let dir = guard.0.join(format!("validation-{index}"));
                    fs::create_dir_all(&dir)?;
                    dir
                })?;
            let mut validation = validate_recalculated_workbook(input, &validated_candidate)?;
            let mut validation_json: serde_json::Value = serde_json::from_str(&validation)?;
            validation_json["promoted_error_constants_restored"] =
                serde_json::json!(restored_error_constants);
            validation = serde_json::to_string(&validation_json)?;
            publish_recalculated_output(&validated_candidate, output)?;
            Ok(serde_json::from_str(&validation)?)
        })();
        match attempt {
            Ok(validation) => {
                books.push(serde_json::json!({
                    "index": index + 1,
                    "status": "recalculated",
                    "engine": engine_name,
                    "input": input.display().to_string(),
                    "output": output.display().to_string(),
                    "validation": validation,
                    "formula_cache_values_cleared_before_engine": cleared_counts[index],
                }));
            }
            Err(error) => {
                failures += 1;
                books.push(serde_json::json!({
                    "index": index + 1,
                    "status": "failed",
                    "input": input.display().to_string(),
                    "error": format!("{error:#}"),
                }));
            }
        }
    }
    let result = serde_json::json!({
        "status": if failures == 0 { "recalculated" } else { "partial" },
        "engine": engine_name,
        "soffice_conversion_invocations": 1,
        "profile_processes_remaining": 0,
        "convert_ms": convert_ms,
        "books": books,
    });
    Ok(result)
}

fn recalculation_timeout() -> Duration {
    let seconds = env::var("AUDIT_RECALC_TIMEOUT_SECONDS")
        .ok()
        .and_then(|value| value.parse::<u64>().ok())
        .unwrap_or(300)
        .clamp(5, 3600);
    Duration::from_secs(seconds)
}

fn file_uri(path: &Path) -> Result<String> {
    ensure!(
        path.is_absolute(),
        "LibreOffice profile path must be absolute"
    );
    let path = path.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    let path = if path.starts_with('/') {
        path
    } else {
        format!("/{path}")
    };
    let mut encoded = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b':' | b'-' | b'.' | b'_' | b'~') {
            encoded.push(byte as char);
        } else {
            encoded.push_str(&format!("%{byte:02X}"));
        }
    }
    Ok(format!("file://{encoded}"))
}

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
    Ok(format!("-env:UserInstallation={}", file_uri(profile)?))
}

fn require_engine_output(path: &Path) -> Result<()> {
    ensure!(
        path.is_file() && fs::metadata(path)?.len() > 0,
        "calculation engine did not create its expected non-empty output workbook: {}",
        path.display()
    );
    Ok(())
}

fn publish_recalculated_output(candidate: &Path, output: &Path) -> Result<()> {
    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let publish_staging = TempDirGuard(unique_temp_dir(output_parent)?);
    let staged = publish_staging.0.join("recalculated.xlsx");
    fs::copy(candidate, &staged).with_context(|| {
        format!(
            "stage recalculated output for publication: {}",
            output.display()
        )
    })?;
    require_engine_output(&staged)?;
    // The staged file is beside the requested output, so a hard link publishes
    // atomically and fails if another process occupied the destination.
    fs::hard_link(&staged, output).with_context(|| {
        format!(
            "publish recalculated workbook without overwriting {}",
            output.display()
        )
    })?;
    Ok(())
}
fn validate_recalculated_workbook(input: &Path, output: &Path) -> Result<String> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .context("read source workbook before recalculation")?
        .workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(output)
        .context("read recalculated workbook")?
        .workbook;
    let formula_count = source
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    ensure!(
        formula_count > 0,
        "workbook contains no formulas to recalculate"
    );
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &result),
        "recalculated workbook changed worksheet structure, formula locations, or formula text beyond LibreOffice boolean-literal normalization"
    );
    let result_value_coverage = excel_reader::formula_result_value_coverage(&result);
    let result_value_count = result_value_coverage.typed_result_cell_count;
    ensure!(
        result_value_coverage.formula_cell_count == formula_count
            && result_value_count == formula_count
            && result_value_coverage.missing_result_cells.is_empty(),
        "recalculated workbook has formula/result value count mismatch: {formula_count} formulas, {result_value_count} typed cached results, missing={:?}",
        result_value_coverage.missing_result_cells.first()
    );
    let cache_count = formula_cache_cell_count(output)?;
    ensure!(cache_count == formula_count, "recalculated workbook has formula/cache count mismatch: {formula_count} formulas, {cache_count} cached formula cells");
    Ok(format!("{{\"formula_text_preserved\":true,\"formula_normalization\":\"Only TRUE/FALSE and TRUE()/FALSE() boolean literal spellings are treated as equivalent; matching literal error cells promoted by LibreOffice are restored before validation\",\"formula_count\":{formula_count},\"formula_result_values\":{result_value_count},\"formula_cache_cells\":{cache_count},\"promoted_error_constants_restored\":0,\"cache_nodes_complete\":true,\"result_values_readable\":true,\"scope\":\"fresh caches produced by an external calculation engine; result correctness requires business-rule or expected-value comparison\"}}"))
}

fn normalize_promoted_error_constants(
    input: &Path,
    candidate: &Path,
    attempt_dir: &Path,
) -> Result<(PathBuf, usize)> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .context("read source workbook for formula normalization")?
        .workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(candidate)
        .context("read engine output for formula normalization")?
        .workbook;
    if excel_reader::formula_texts_equivalent_after_recalculation(&source, &result) {
        return Ok((candidate.to_path_buf(), 0));
    }

    let promoted = promoted_error_constants(&source, &result)?;
    ensure!(
        !promoted.is_empty(),
        "recalculated workbook changed worksheet structure, formula locations, or formula text beyond supported normalization"
    );
    let output = attempt_dir.join("recalculated-error-constants-restored.xlsx");
    let report =
        ooxml_patch::restore_promoted_error_constants(input, candidate, &output, &promoted)
            .context("restore LibreOffice-promoted literal error cells")?;
    let restored = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&output)
        .context("read back normalized engine output")?
        .workbook;
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &restored),
        "recalculated workbook still changes formula structure after restoring literal error cells"
    );
    Ok((output, report.restored_cells.len()))
}

fn promoted_error_constants(
    source: &WorkbookDump,
    result: &WorkbookDump,
) -> Result<Vec<ooxml_patch::PromotedErrorConstant>> {
    ensure!(
        source.sheets.len() == result.sheets.len(),
        "recalculated workbook changed worksheet count"
    );
    let mut promoted = Vec::new();
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        ensure!(
            before.name == after.name,
            "recalculated workbook changed worksheet order or names"
        );
        for (cell, formula) in &before.formulas {
            let result_formula = after.formulas.get(cell).with_context(|| {
                format!(
                    "recalculated workbook removed formula at {}!{cell}",
                    before.name
                )
            })?;
            ensure!(
                excel_reader::canonicalize_formula_boolean_literals(formula)
                    == excel_reader::canonicalize_formula_boolean_literals(result_formula),
                "recalculated workbook changed formula at {}!{cell}",
                before.name
            );
        }
        for (cell, formula) in &after.formulas {
            if before.formulas.contains_key(cell) {
                continue;
            }
            ensure!(
                is_excel_error_literal(formula),
                "recalculated workbook added an unsupported formula at {}!{cell}",
                before.name
            );
            let source_value = before.values.get(cell).with_context(|| {
                format!(
                    "new formula at {}!{cell} has no matching source error value",
                    before.name
                )
            })?;
            let result_value = after.values.get(cell).with_context(|| {
                format!(
                    "new formula at {}!{cell} has no readable cached error value",
                    after.name
                )
            })?;
            ensure!(
                source_value.get("type").and_then(serde_json::Value::as_str) == Some("error")
                    && result_value.get("type").and_then(serde_json::Value::as_str)
                        == Some("error")
                    && source_value == result_value,
                "new formula at {}!{cell} does not preserve a source literal error value",
                before.name
            );
            promoted.push(ooxml_patch::PromotedErrorConstant {
                sheet_name: before.name.clone(),
                cell_reference: cell.clone(),
                error_token: formula.clone(),
            });
        }
    }
    Ok(promoted)
}

fn is_excel_error_literal(formula: &str) -> bool {
    matches!(
        formula,
        "#NULL!"
            | "#DIV/0!"
            | "#VALUE!"
            | "#REF!"
            | "#NAME?"
            | "#NUM!"
            | "#N/A"
            | "#SPILL!"
            | "#CALC!"
            | "#FIELD!"
            | "#GETTING_DATA"
    )
}

fn formula_cache_cell_count(path: &Path) -> Result<usize> {
    let inventory = excel_reader::inspect_formula_cache_nodes(path)?;
    if let Some(missing) = inventory.missing_cache_cells.first() {
        bail!(
            "formula cell {:?} on sheet {:?} has no cached value node after recalculation",
            missing.cell,
            missing.sheet_name
        );
    }
    Ok(inventory.cached_formula_cell_count)
}
fn unique_temp_dir(parent: &Path) -> Result<PathBuf> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for i in 0..100 {
        let path = parent.join(format!(".audit-lo-{}-{nanos}-{i}", std::process::id()));
        if !path.exists() {
            fs::create_dir(&path)?;
            return Ok(path);
        }
    }
    bail!("cannot allocate conversion work directory")
}
struct TempDirGuard(PathBuf);
impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::SheetDump;
    use std::{fs::File, io::Write};
    use zip::{write::SimpleFileOptions, ZipWriter};
    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    fn fixture(with_cache: bool) -> Result<PathBuf> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = env::temp_dir().join(format!(
            "lo_recalc_fixture_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let zip = ZipWriter::new(File::create(&path)?);
        let workbook = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Formula" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let cache = if with_cache { "<v>2</v>" } else { "" };
        let sheet = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="1"><c r="A1"><f>1+1</f>{cache}</c></row></sheetData></worksheet>"#
        );
        let mut writer = zip;
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/worksheets/sheet1.xml", sheet),
        ] {
            writer.start_file(name, SimpleFileOptions::default())?;
            writer.write_all(content.as_bytes())?;
        }
        writer.finish()?;
        Ok(path)
    }
    fn error_promotion_workbook(
        formulas: BTreeMap<String, String>,
        error_value: &str,
        include_error_formula: bool,
    ) -> WorkbookDump {
        let values = BTreeMap::from([(
            "C211".to_owned(),
            serde_json::json!({"type":"error","value":error_value}),
        )]);
        let mut formulas = formulas;
        if include_error_formula {
            formulas.insert("C211".to_owned(), "#REF!".to_owned());
        }
        WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "参照表".to_owned(),
                values,
                formulas,
                formula_status: "ok".to_owned(),
                merges: Vec::new(),
                merge_status: "ok".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "ok".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "ok".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "ok".to_owned(),
            }],
        }
    }
    #[test]
    fn promotes_only_an_unchanged_literal_error_cell_for_ooxml_restoration() -> Result<()> {
        let source = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "Ref",
            false,
        );
        let result = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE())".to_owned())]),
            "Ref",
            true,
        );

        let promotions = promoted_error_constants(&source, &result)?;
        assert_eq!(promotions.len(), 1);
        assert_eq!(promotions[0].sheet_name, "参照表");
        assert_eq!(promotions[0].cell_reference, "C211");
        assert_eq!(promotions[0].error_token, "#REF!");
        Ok(())
    }
    #[test]
    fn rejects_changed_formula_or_error_value_during_normalization() {
        let source = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "Ref",
            false,
        );
        let changed_formula = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Other!A:B,2,FALSE)".to_owned())]),
            "Ref",
            true,
        );
        assert!(promoted_error_constants(&source, &changed_formula)
            .expect_err("changed formula reference must be rejected")
            .to_string()
            .contains("changed formula"));

        let changed_error = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "NA",
            true,
        );
        assert!(promoted_error_constants(&source, &changed_error)
            .expect_err("different typed error must be rejected")
            .to_string()
            .contains("does not preserve a source literal error value"));
    }
    #[test]
    fn empty_string_formula_cache_remains_a_readable_result() -> Result<()> {
        let mut workbook = excel_writer::Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_formula(0, 0, "=IF(TRUE,\"\",\"x\")")?;
        sheet.set_formula_result(0, 0, "");
        let path = fixture(false)?;
        workbook.save(&path)?;
        let read = excel_reader::read_workbook(&path)?;
        assert_eq!(
            read.sheets[0].values.get("A1"),
            Some(&serde_json::json!({"type": "string", "value": ""}))
        );
        fs::remove_file(path)?;
        Ok(())
    }
    #[test]
    fn recalc_validation_rejects_changed_formula_text() -> Result<()> {
        let root = env::temp_dir().join(format!(
            "lo_recalc_formula_change_{}_{}",
            std::process::id(),
            SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
        ));
        fs::create_dir_all(&root)?;
        let input = root.join("input.xlsx");
        let output = root.join("output.xlsx");

        let mut source = excel_writer::Workbook::new();
        let source_sheet = source.add_worksheet();
        source_sheet.write_number(0, 0, 2)?;
        source_sheet.write_formula(0, 1, "=A1+1")?;
        source_sheet.set_formula_result(0, 1, "3");
        source.save(&input)?;

        let mut changed = excel_writer::Workbook::new();
        let changed_sheet = changed.add_worksheet();
        changed_sheet.write_number(0, 0, 2)?;
        changed_sheet.write_formula(0, 1, "=A1+2")?;
        changed_sheet.set_formula_result(0, 1, "4");
        changed.save(&output)?;

        let error = validate_recalculated_workbook(&input, &output)
            .expect_err("formula text changes must fail recalculation validation");
        assert!(error.to_string().contains("changed worksheet structure"));
        fs::remove_dir_all(root)?;
        Ok(())
    }
    #[test]
    fn formula_cache_counter_requires_a_value_node_for_every_formula() -> Result<()> {
        let valid = fixture(true)?;
        assert_eq!(formula_cache_cell_count(&valid)?, 1);
        fs::remove_file(valid)?;

        let invalid = fixture(false)?;
        let error =
            formula_cache_cell_count(&invalid).expect_err("formula without cache must fail");
        assert!(error.to_string().contains("has no cached value node"));
        fs::remove_file(invalid)?;
        Ok(())
    }
    #[test]
    fn missing_engine_output_is_rejected_even_when_request_copy_exists() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root =
            env::temp_dir().join(format!("lo_recalc_output_{}_{}", std::process::id(), nonce));
        let input_dir = root.join("input");
        let output_dir = root.join("engine-output");
        fs::create_dir_all(&input_dir)?;
        fs::create_dir_all(&output_dir)?;
        let request_copy = input_dir.join("recalculation-request.xlsx");
        fs::write(&request_copy, b"input copy")?;
        let expected_output = output_dir.join("recalculation-request.xlsx");

        let error = require_engine_output(&expected_output)
            .expect_err("missing engine output must not select the input copy");
        assert!(error.to_string().contains("did not create"));
        assert!(request_copy.is_file());
        assert!(!expected_output.exists());

        fs::remove_dir_all(root)?;
        Ok(())
    }
    #[test]
    fn recalculated_output_publication_preserves_existing_destination() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = env::temp_dir().join(format!(
            "lo_recalc_publish_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root)?;
        let candidate = root.join("candidate.xlsx");
        let output = root.join("output.xlsx");
        fs::write(&candidate, b"validated workbook bytes")?;
        fs::write(&output, b"pre-existing destination")?;

        let error = publish_recalculated_output(&candidate, &output)
            .expect_err("publication must not overwrite a pre-existing output");
        assert!(format!("{error:#}").contains("without overwriting"));
        assert_eq!(fs::read(&output)?, b"pre-existing destination");

        fs::remove_file(&output)?;
        publish_recalculated_output(&candidate, &output)?;
        assert_eq!(fs::read(&output)?, b"validated workbook bytes");
        fs::remove_dir_all(root)?;
        Ok(())
    }
    #[test]
    fn libreoffice_profile_path_is_encoded_as_a_file_uri() -> Result<()> {
        let root = env::temp_dir().join("lo profile # 中文");
        let uri = file_uri(&root)?;
        assert!(uri.starts_with("file://"));
        assert!(uri.contains("%20"));
        assert!(uri.contains("%23"));
        assert!(!uri.contains(' '));
        assert!(!uri.contains('中'));
        Ok(())
    }
    #[test]
    fn libreoffice_user_installation_argument_creates_profile_and_uses_file_uri() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let profile = env::temp_dir().join(format!(
            "lo_profile_arg_{}_{}_# space",
            std::process::id(),
            nonce
        ));
        assert!(!profile.exists());

        let argument = libreoffice_user_installation_arg(&profile)?;

        assert!(profile.is_dir());
        assert!(argument.starts_with("-env:UserInstallation=file://"));
        assert!(argument.contains("%20"));
        assert!(argument.contains("%23"));
        fs::remove_dir_all(profile)?;
        Ok(())
    }

    #[test]
    fn worker_refuses_existing_output_before_starting_office() -> Result<()> {
        let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
        let input = guard.0.join("original.xlsx");
        let output = guard.0.join("existing.xlsx");
        fs::write(&input, b"source")?;
        fs::write(&output, b"keep")?;
        let error = execute_batch(vec![(input.clone(), output.clone())]).unwrap_err();
        ensure!(error.to_string().contains("refusing to overwrite"));
        ensure!(fs::read(&input)? == b"source" && fs::read(&output)? == b"keep");
        Ok(())
    }
    #[test]
    fn worker_refuses_duplicate_outputs_before_starting_office() -> Result<()> {
        let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
        let input = guard.0.join("original.xlsx");
        let output = guard.0.join("new.xlsx");
        fs::write(&input, b"source")?;
        let error = execute_batch(vec![
            (input.clone(), output.clone()),
            (input.clone(), output.clone()),
        ])
        .unwrap_err();
        ensure!(error.to_string().contains("输出路径重复"));
        ensure!(!output.exists() && fs::read(&input)? == b"source");
        Ok(())
    }
}
