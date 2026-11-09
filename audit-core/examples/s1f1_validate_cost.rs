//! 量化 S1-F01 候选校验里各次"整簿解析"的成本。
use std::path::PathBuf;
use std::time::Instant;

fn main() -> anyhow::Result<()> {
    let path = PathBuf::from(std::env::args().nth(1).expect("workbook"));
    let t = Instant::now();
    let wb = excel_reader::read_workbook(&path)?;
    let formulas: usize = wb.sheets.iter().map(|s| s.formulas.len()).sum();
    println!(
        "{:>6} ms | read_workbook (calamine)      | {} 表 / {formulas} 公式",
        t.elapsed().as_millis(),
        wb.sheets.len()
    );

    let t = Instant::now();
    let profile = excel_editor::inspect_xlsx_profile(&path)?;
    println!(
        "{:>6} ms | inspect_xlsx_profile          | complete={}",
        t.elapsed().as_millis(),
        profile.inventory_complete
    );

    let _ = &profile;
    Ok(())
}
