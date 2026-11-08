//! 临时诊断：用生产 reader 导出指定单元格的公式文本，比较源副本与 COM 输出。
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let path = PathBuf::from(args.get(1).expect("usage: <xlsx> [sheet] [cells...]"));
    let sheet_filter = args.get(2).cloned();
    let cells: Vec<String> = args.iter().skip(3).cloned().collect();
    let report = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&path)?;
    for sheet in &report.workbook.sheets {
        if let Some(filter) = &sheet_filter {
            if &sheet.name != filter {
                continue;
            }
        }
        println!(
            "== sheet {} ({} formulas)",
            sheet.name,
            sheet.formulas.len()
        );
        if cells.is_empty() {
            for (cell, formula) in sheet.formulas.iter().take(5) {
                println!("   {cell} = {formula}");
            }
        } else {
            for cell in &cells {
                println!("   {cell} = {:?}", sheet.formulas.get(cell));
            }
        }
    }
    Ok(())
}
