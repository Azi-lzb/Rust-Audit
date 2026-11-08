//! betteroffice 拒绝分支与保真探针。
//!
//! 用法：bo_edge_probe <file.xlsx> [--fallback none]
//! 直接调用 bo_calc::recalculate（不经路由），打印结果或拒绝原因。
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let source = PathBuf::from(args.get(1).expect("usage: <file.xlsx>"));
    if args.iter().any(|arg| arg == "--fallback-none") {
        std::env::set_var("PYTOOLS_BO_CALC_FALLBACK", "none");
    }
    let config = audit_core::bo_calc::BoCalcConfig::from_env();
    eprintln!(
        "config: enabled={} rounds={} max_visits={:?} strip={} fallback(com={}, lo={})",
        config.enabled,
        config.rounds,
        config.max_visits,
        config.strip_style_cells,
        config.fallback_com,
        config.fallback_libreoffice
    );
    let work = std::env::temp_dir().join(format!(
        "bo_edge_{}",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos()
    ));
    std::fs::create_dir_all(&work)?;
    let copy = work.join("copy.xlsx");
    std::fs::copy(&source, &copy)?;
    match audit_core::bo_calc::recalculate(&copy, &config) {
        Ok(outcome) => {
            println!(
                "OK engine={} rounds={} changed={} limited={} errors={} name_errors={} stripped={} \
                 open={:?} calc={:?} save={:?}",
                outcome.application,
                outcome.rounds_used,
                outcome.changed_cells_total,
                outcome.limited_cells,
                outcome.error_cells,
                outcome.name_error_cells,
                outcome.style_cells_stripped,
                outcome.open_elapsed,
                outcome.calc_elapsed,
                outcome.save_elapsed
            );
            // 回读校验：确认输出可被 reader 打开且公式/值数量不减
            let before = excel_reader::read_workbook(&source)?;
            let after = excel_reader::read_workbook(&copy)?;
            let before_formulas: usize = before.sheets.iter().map(|s| s.formulas.len()).sum();
            let after_formulas: usize = after.sheets.iter().map(|s| s.formulas.len()).sum();
            println!(
                "readback: sheets {}->{} formulas {}->{}",
                before.sheets.len(),
                after.sheets.len(),
                before_formulas,
                after_formulas
            );
            anyhow::ensure!(
                before.sheets.len() == after.sheets.len(),
                "重算改变了工作表数量"
            );
            anyhow::ensure!(before_formulas == after_formulas, "重算改变了公式数量");
        }
        Err(error) => {
            println!("REFUSED {error:#}");
        }
    }
    let _ = std::fs::remove_dir_all(&work);
    Ok(())
}
