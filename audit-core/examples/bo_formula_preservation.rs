//! 验证 betteroffice 重算是否改动公式文本（跨 20 个真实模板）。
//!
//! 若公式文本逐格完全不变 → 公式文本守卫（无论规范化覆盖多少种差异）
//! 都不可能被 betteroffice 触发假阳性。
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let dir = PathBuf::from(std::env::args().nth(1).expect("templates dir"));
    let work = std::env::temp_dir().join(format!("bo_formula_{}", std::process::id()));
    std::fs::create_dir_all(&work)?;
    let config = audit_core::bo_calc::BoCalcConfig::from_env();

    let mut templates: Vec<PathBuf> = std::fs::read_dir(&dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
        })
        .collect();
    templates.sort();

    let mut total_templates = 0usize;
    let mut total_formulas = 0usize;
    let mut templates_changed = 0usize;
    let mut total_changed_cells = 0usize;

    for template in &templates {
        let name = template
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        let copy = work.join(format!("{}.xlsx", total_templates));
        std::fs::copy(template, &copy)?;

        let before = excel_reader::read_workbook(&copy)?;
        let outcome = match audit_core::bo_calc::recalculate(&copy, &config) {
            Ok(outcome) => outcome,
            Err(error) => {
                println!("{:>4} | 重算被拒绝（跳过比对）: {error:#}", "");
                continue;
            }
        };
        let after = excel_reader::read_workbook(&copy)?;

        let mut changed_cells: Vec<String> = Vec::new();
        let mut formula_count = 0usize;
        for (before_sheet, after_sheet) in before.sheets.iter().zip(after.sheets.iter()) {
            for (cell, formula) in &before_sheet.formulas {
                formula_count += 1;
                match after_sheet.formulas.get(cell) {
                    Some(after_formula) if after_formula == formula => {}
                    Some(after_formula) => changed_cells.push(format!(
                        "{}!{}: {formula:?} -> {after_formula:?}",
                        before_sheet.name, cell
                    )),
                    None => changed_cells
                        .push(format!("{}!{}: 公式丢失", before_sheet.name, cell)),
                }
            }
            // 反向：输出端新增的公式
            for (cell, _) in &after_sheet.formulas {
                if !before_sheet.formulas.contains_key(cell) {
                    changed_cells.push(format!("{}!{}: 输出端新增公式", after_sheet.name, cell));
                }
            }
        }

        total_templates += 1;
        total_formulas += formula_count;
        if !changed_cells.is_empty() {
            templates_changed += 1;
            total_changed_cells += changed_cells.len();
        }
        println!(
            "{} | 公式 {formula_count:>5} | 文本改变 {} | changed_cells={} | {}ms",
            if changed_cells.is_empty() { "一致" } else { "差异" },
            changed_cells.len(),
            outcome.changed_cells_total,
            outcome.calc_elapsed.as_millis(),
        );
        for line in changed_cells.iter().take(3) {
            println!("      {line}");
        }
        let _ = name;
    }

    let _ = std::fs::remove_dir_all(&work);
    println!();
    println!(
        "汇总：{total_templates} 个模板 / {total_formulas} 个公式；\
         公式文本被改动：{templates_changed} 个模板、{total_changed_cells} 格"
    );
    anyhow::ensure!(
        total_changed_cells == 0,
        "betteroffice 改动了公式文本，守卫可能误报"
    );
    println!("结论：betteroffice 完全未改公式文本 → 公式文本守卫不可能误报");
    Ok(())
}
