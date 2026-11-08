fn main() -> anyhow::Result<()> {
    let src = std::path::PathBuf::from(std::env::args().nth(1).expect("xlsx"));
    let bytes = std::fs::read(&src)?;
    let mut workbook = betteroffice_xlsx::Workbook::open(&bytes)
        .map_err(|e| anyhow::anyhow!("open: {e}"))?;
    let options = betteroffice_xlsx::CalculationOptions { now_serial: Some(45_000.0) };
    for round in 1..=5 {
        let result = workbook.recalculate_all(options);
        println!(
            "round {round}: changed={} limited={} cycle={} budget_exhausted={}",
            result.changed.len(),
            result.limited_cells.len(),
            result.cycle_cells.len(),
            result.budget_exhausted
        );
    }
    Ok(())
}
