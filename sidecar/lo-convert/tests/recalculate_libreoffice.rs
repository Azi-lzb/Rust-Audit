use anyhow::{ensure, Context, Result};
use roxmltree::Document;
use rust_xlsxwriter::Workbook;
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
        if env::var_os("LO_CONVERT_REQUIRE_LO_TEST").is_some() {
            anyhow::bail!("LO_CONVERT_REQUIRE_LO_TEST is set but soffice is unavailable");
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

    let result = Command::new(env!("CARGO_BIN_EXE_lo-convert"))
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
        "lo_convert_integration_{}_{}",
        std::process::id(),
        nonce
    )))
}
