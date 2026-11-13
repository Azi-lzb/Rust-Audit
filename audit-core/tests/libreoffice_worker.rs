use anyhow::{ensure, Context, Result};
use roxmltree::Document;
use excel_writer::Workbook;
use std::{
    env, fs,
    io::{Read, Write},
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[test]
fn cli_recalculates_formula_and_publishes_validated_xlsx() -> Result<()> {
    if !libreoffice_available() {
        if env::var_os("AUDIT_REQUIRE_LO_TEST").is_some() {
            anyhow::bail!("AUDIT_REQUIRE_LO_TEST is set but soffice is unavailable");
        }
        eprintln!("skipping LibreOffice CLI integration: soffice is unavailable");
        return Ok(());
    }

    let root = temp_dir()?;
    fs::create_dir_all(&root)?;
    let input = root.join("formula-input.xlsx");
    let output = root.join("formula-output.xlsx");
    let mut workbook = Workbook::new();
    let data = workbook.add_worksheet();
    data.set_name("Data")?;
    data.write_string(0, 0, "A")?;
    data.write_number(0, 1, 11)?;
    data.write_string(1, 0, "B")?;
    data.write_number(1, 1, 22)?;
    data.write_string(2, 0, "#REF!")?;
    let calc = workbook.add_worksheet();
    calc.set_name("Calc")?;
    calc.write_string(0, 0, "B")?;
    calc.write_formula(0, 1, "=VLOOKUP(A1,Data!$A$1:$B$2,2,FALSE)")?;
    calc.write_formula(0, 2, "=Data!B1+2")?;
    calc.write_formula(0, 3, "=1/0")?;
    calc.write_formula(0, 4, "=IF(TRUE,\"\",\"x\")")?;
    workbook.save(&input)?;
    write_literal_error_constant(&input, "A3", "#REF!")?;
    let input_dump = excel_reader::read_workbook(&input)?;
    let source_error = input_dump.sheets[0].values["A3"].clone();

    let result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
        .arg(audit_core::libreoffice_worker::WORKER_COMMAND)
        .args([
            "recalculate",
            "--input",
            input.to_str().context("input path is not UTF-8")?,
            "--output",
            output.to_str().context("output path is not UTF-8")?,
            "--engine",
            "lo",
        ])
        .output()?;
    ensure!(
        result.status.success(),
        "recalculate CLI failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&result.stdout)?;
    ensure!(report["status"] == "recalculated");
    ensure!(report["engine"] == "LibreOffice");
    ensure!(report["formula_cache_values_cleared_before_engine"] == 4);
    ensure!(report["validation"]["formula_text_preserved"] == true);
    ensure!(report["validation"]["cache_nodes_complete"] == true);
    ensure!(report["validation"]["result_values_readable"] == true);
    ensure!(report["validation"]["promoted_error_constants_restored"] == 1);
    ensure!(output.is_file());

    let recalculated = excel_reader::read_workbook(&output)?;
    let data = recalculated
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Data")
        .context("recalculated Data sheet missing")?;
    let calc = recalculated
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Calc")
        .context("recalculated Calc sheet missing")?;
    ensure!(
        calc.formulas.len() == 4,
        "expected all four formulas to remain"
    );
    ensure!(
        calc.formulas["B1"].contains("VLOOKUP")
            && calc.formulas["C1"].contains("Data")
            && calc.formulas["D1"].contains("1/0")
            && calc.formulas["E1"].contains("IF"),
        "formula text or cross-sheet references changed: {:?}",
        calc.formulas
    );
    ensure!(
        data.values["B1"]["value"].as_f64() == Some(11.0)
            && data.values["B2"]["value"].as_f64() == Some(22.0),
        "recalculation changed source data values"
    );
    ensure!(
        !data.formulas.contains_key("A3") && data.values["A3"] == source_error,
        "LibreOffice error-constant normalization changed Data!A3"
    );
    ensure!(
        calc.values["B1"]["value"].as_f64() == Some(22.0),
        "VLOOKUP result was not recalculated: {}",
        calc.values["B1"]
    );
    ensure!(
        calc.values["C1"]["value"].as_f64() == Some(13.0),
        "cross-sheet formula result was not recalculated: {}",
        calc.values["C1"]
    );
    ensure!(
        calc.values["D1"]["type"] == "error",
        "formula error was not returned as a typed error: {}",
        calc.values["D1"]
    );
    ensure!(
        calc.values["E1"] == serde_json::json!({"type": "string", "value": ""}),
        "empty-string formula result was not preserved: {}",
        calc.values["E1"]
    );
    let coverage = excel_reader::formula_result_value_coverage(&recalculated);
    ensure!(coverage.formula_cell_count == 4 && coverage.typed_result_cell_count == 4);
    ensure!(coverage.missing_result_cells.is_empty());

    let mut zip = ZipArchive::new(fs::File::open(&output)?)?;
    let workbook_xml = {
        let mut workbook_xml = String::new();
        zip.by_name("xl/workbook.xml")?
            .read_to_string(&mut workbook_xml)?;
        workbook_xml
    };
    let document = Document::parse(&workbook_xml)?;
    ensure!(
        document
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "sheet")
            .map(|node| node.attribute("name").unwrap_or_default())
            .collect::<Vec<_>>()
            == vec!["Data", "Calc"],
        "recalculated workbook sheet structure changed"
    );

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn cli_recalculates_a_workbook_batch_with_one_lo_conversion_invocation() -> Result<()> {
    if !libreoffice_available() {
        if env::var_os("AUDIT_REQUIRE_LO_TEST").is_some() {
            anyhow::bail!("AUDIT_REQUIRE_LO_TEST is set but soffice is unavailable");
        }
        eprintln!("skipping LibreOffice batch integration: soffice is unavailable");
        return Ok(());
    }

    let root = temp_dir()?;
    fs::create_dir_all(&root)?;
    let mut pairs = Vec::new();
    let mut source_bytes = Vec::new();
    for index in 0..2 {
        let input = root.join(format!("batch-input-{index}.xlsx"));
        let output = root.join(format!("batch-output-{index}.xlsx"));
        let mut workbook = Workbook::new();
        let data = workbook.add_worksheet();
        data.set_name("Data")?;
        data.write_number(0, 0, 10.0 + index as f64)?;
        let calc = workbook.add_worksheet();
        calc.set_name("Calc")?;
        calc.write_formula(0, 0, "=Data!A1*2")?;
        workbook.save(&input)?;
        source_bytes.push(fs::read(&input)?);
        pairs.push((input, output));
    }
    let pairs_file = root.join("recalculation-pairs.tsv");
    let mut pairs_text = String::new();
    for (input, output) in &pairs {
        pairs_text.push_str(&format!("{}\t{}\n", input.display(), output.display()));
    }
    fs::write(&pairs_file, pairs_text)?;

    let result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
        .arg(audit_core::libreoffice_worker::WORKER_COMMAND)
        .args([
            "recalculate-batch",
            "--pairs",
            pairs_file.to_str().context("pairs path is not UTF-8")?,
            "--engine",
            "lo",
        ])
        .output()?;
    ensure!(
        result.status.success(),
        "batch recalculation CLI failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&result.stdout)?;
    ensure!(report["status"] == "recalculated");
    ensure!(report["engine"] == "LibreOffice");
    ensure!(report["soffice_conversion_invocations"] == 1);
    ensure!(report["profile_processes_remaining"] == 0);
    ensure!(report["books"]
        .as_array()
        .is_some_and(|books| books.len() == 2));

    for (index, (input, output)) in pairs.iter().enumerate() {
        ensure!(
            report["books"][index]["status"] == "recalculated",
            "batch book {} failed: {}",
            index + 1,
            report["books"][index]
        );
        ensure!(
            fs::read(input)? == source_bytes[index],
            "batch modified an input workbook"
        );
        let recalculated = excel_reader::read_workbook(output)?;
        let calc = recalculated
            .sheets
            .iter()
            .find(|sheet| sheet.name == "Calc")
            .context("recalculated Calc sheet missing")?;
        ensure!(calc.formulas.contains_key("A1"));
        ensure!(
            calc.values["A1"]["value"].as_f64() == Some(20.0 + index as f64 * 2.0),
            "batch formula result was not recalculated: {}",
            calc.values["A1"]
        );
    }

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
#[ignore = "manual WSL LibreOffice throughput check for 2/10/33 S1-F1 workbooks"]
fn manual_s1_formula_recalculation_batch_performance_2_10_33() -> Result<()> {
    ensure!(
        libreoffice_available(),
        "LibreOffice is required for this manual performance check"
    );
    let root = temp_dir()?;
    fs::create_dir_all(&root)?;
    let mut results = Vec::new();

    for book_count in [2usize, 10, 33] {
        let input_dir = root.join(format!("inputs-{book_count}"));
        let per_book_dir = root.join(format!("per-book-{book_count}"));
        let batch_dir = root.join(format!("batch-{book_count}"));
        fs::create_dir_all(&input_dir)?;
        fs::create_dir_all(&per_book_dir)?;
        fs::create_dir_all(&batch_dir)?;
        let mut pairs = Vec::with_capacity(book_count);
        let mut source_bytes = Vec::with_capacity(book_count);
        for index in 0..book_count {
            let input = input_dir.join(format!("book-{index:04}.xlsx"));
            build_performance_workbook(&input, 400, index + 1)?;
            source_bytes.push(fs::read(&input)?);
            pairs.push((
                input,
                per_book_dir.join(format!("book-{index:04}-recalculated.xlsx")),
            ));
        }

        let per_book_started = std::time::Instant::now();
        for (input, output) in &pairs {
            let result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
                .arg(audit_core::libreoffice_worker::WORKER_COMMAND)
                .args([
                    "recalculate",
                    "--input",
                    input.to_str().context("input path is not UTF-8")?,
                    "--output",
                    output.to_str().context("output path is not UTF-8")?,
                    "--engine",
                    "lo",
                ])
                .output()?;
            ensure!(
                result.status.success(),
                "per-book recalculation failed: {}",
                String::from_utf8_lossy(&result.stderr)
            );
        }
        let per_book_elapsed_ms = per_book_started.elapsed().as_millis() as u64;

        let batch_pairs = pairs
            .iter()
            .map(|(input, _)| -> Result<_> {
                Ok((
                    input.clone(),
                    batch_dir.join(
                        input
                            .file_name()
                            .context("input workbook name is missing")?,
                    ),
                ))
            })
            .collect::<Result<Vec<_>>>()?;
        let pairs_file = root.join(format!("pairs-{book_count}.tsv"));
        let mut pairs_text = String::new();
        for (input, output) in &batch_pairs {
            pairs_text.push_str(&format!("{}\t{}\n", input.display(), output.display()));
        }
        fs::write(&pairs_file, pairs_text)?;
        let batch_started = std::time::Instant::now();
        let batch_result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
            .arg(audit_core::libreoffice_worker::WORKER_COMMAND)
            .args([
                "recalculate-batch",
                "--pairs",
                pairs_file.to_str().context("pairs path is not UTF-8")?,
                "--engine",
                "lo",
            ])
            .output()?;
        let batch_elapsed_ms = batch_started.elapsed().as_millis() as u64;
        ensure!(
            batch_result.status.success(),
            "batch recalculation failed: {}",
            String::from_utf8_lossy(&batch_result.stderr)
        );
        let batch_report: serde_json::Value = serde_json::from_slice(&batch_result.stdout)?;
        ensure!(batch_report["status"] == "recalculated");
        ensure!(batch_report["soffice_conversion_invocations"] == 1);
        ensure!(batch_report["profile_processes_remaining"] == 0);

        for (index, ((input, per_book_output), (_, batch_output))) in
            pairs.iter().zip(&batch_pairs).enumerate()
        {
            ensure!(
                batch_report["books"][index]["status"] == "recalculated",
                "batch workbook {} failed: {}",
                index + 1,
                batch_report["books"][index]
            );
            ensure!(
                fs::read(input)? == source_bytes[index],
                "batch run modified an input workbook"
            );
            verify_performance_workbook(per_book_output, index + 1)?;
            verify_performance_workbook(batch_output, index + 1)?;
        }

        let result = serde_json::json!({
            "books": book_count,
            "rows_per_book": 400,
            "per_book_total_ms": per_book_elapsed_ms,
            "batch_total_ms": batch_elapsed_ms,
            "batch_average_per_book_ms": batch_elapsed_ms / book_count as u64,
            "soffice_conversion_invocations": batch_report["soffice_conversion_invocations"],
            "profile_processes_remaining": batch_report["profile_processes_remaining"],
            "input_unchanged": true,
            "outputs_formula_and_cache_validation": "passed"
        });
        println!("S1_LO_BATCH_PERF {result}");
        results.push(result);
    }

    if let Some(path) = env::var_os("S1_LO_BATCH_PERF_REPORT") {
        fs::write(
            PathBuf::from(path),
            serde_json::to_vec_pretty(&serde_json::json!({
                "rounds": 1,
                "workbook_rows": 400,
                "results": results,
            }))?,
        )?;
    }
    fs::remove_dir_all(root)?;
    Ok(())
}

fn build_performance_workbook(path: &Path, rows: usize, seed: usize) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Data")?;
    sheet.write_string(0, 0, "序号")?;
    sheet.write_string(0, 1, "数值")?;
    sheet.write_string(0, 2, "公式")?;
    for row in 2..=rows + 1 {
        sheet.write_number((row - 1) as u32, 0, (row - 1) as f64)?;
        sheet.write_number((row - 1) as u32, 1, ((row * 7 + seed) % 1000) as f64)?;
        let formula = format!("=B{row}*2+A{row}");
        sheet.write_formula((row - 1) as u32, 2, formula.as_str())?;
    }
    let total_row = rows + 3;
    sheet.write_string((total_row - 1) as u32, 0, "合计")?;
    let total_formula = format!("=SUM(C2:C{})", rows + 1);
    sheet.write_formula((total_row - 1) as u32, 2, total_formula.as_str())?;
    workbook.save(path)?;
    Ok(())
}

fn verify_performance_workbook(path: &Path, seed: usize) -> Result<()> {
    let workbook = excel_reader::read_workbook(path)?;
    ensure!(workbook.sheets.len() == 1 && workbook.sheets[0].name == "Data");
    let sheet = &workbook.sheets[0];
    ensure!(sheet.formulas.len() == 401);
    for row in 2..=401 {
        let expected = (row * 7 + seed) % 1000 * 2 + (row - 1);
        ensure!(
            sheet.values[&format!("C{row}")]["value"].as_f64() == Some(expected as f64),
            "formula cache mismatch at C{row}: {}",
            sheet.values[&format!("C{row}")]
        );
    }
    let coverage = excel_reader::formula_result_value_coverage(&workbook);
    ensure!(
        coverage.formula_cell_count == 401
            && coverage.typed_result_cell_count == 401
            && coverage.missing_result_cells.is_empty(),
        "formula cache coverage is incomplete: {coverage:?}"
    );
    Ok(())
}

fn write_literal_error_constant(path: &Path, cell_reference: &str, token: &str) -> Result<()> {
    let replacement_path = path.with_file_name("formula-input-literal-error.xlsx");
    ensure!(
        !replacement_path.exists(),
        "literal-error staging file exists"
    );
    let mut source_zip = ZipArchive::new(fs::File::open(path)?)?;
    let mut writer = ZipWriter::new(fs::File::create(&replacement_path)?);
    for index in 0..source_zip.len() {
        let mut entry = source_zip.by_index(index)?;
        let name = entry.name().to_owned();
        let mut options = SimpleFileOptions::default().compression_method(entry.compression());
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        writer.start_file(&name, options)?;
        if name == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            entry.read_to_string(&mut xml)?;
            let document = Document::parse(&xml)?;
            let cell = document
                .descendants()
                .find(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(MAIN_NS)
                        && node.tag_name().name() == "c"
                        && node.attribute("r") == Some(cell_reference)
                })
                .context("literal-error cell is missing")?;
            xml.replace_range(
                cell.range(),
                &format!(r#"<c r="{cell_reference}" t="e"><v>{token}</v></c>"#),
            );
            writer.write_all(xml.as_bytes())?;
        } else {
            std::io::copy(&mut entry, &mut writer)?;
        }
    }
    writer.finish()?;
    drop(source_zip);
    fs::copy(&replacement_path, path)?;
    fs::remove_file(replacement_path)?;
    Ok(())
}

fn libreoffice_available() -> bool {
    Command::new("soffice")
        .arg("--version")
        .output()
        .is_ok_and(|result| result.status.success())
}

fn temp_dir() -> Result<PathBuf> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    Ok(env::temp_dir().join(format!(
        "audit_lo_integration_{}_{}",
        std::process::id(),
        nonce
    )))
}

#[test]
fn private_lo_worker_rejects_invalid_request_without_starting_gui() -> Result<()> {
    let output = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
        .arg(audit_core::libreoffice_worker::WORKER_COMMAND)
        .args(["recalculate", "--invalid", "value"])
        .output()?;
    ensure!(!output.status.success());
    ensure!(output.stdout.is_empty());
    ensure!(String::from_utf8_lossy(&output.stderr).contains("unknown worker option"));
    Ok(())
}
