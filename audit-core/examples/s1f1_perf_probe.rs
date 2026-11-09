//! 测量 betteroffice 在这批真实审核副本上的重算耗时。
//!
//! 用法：s1f1_perf_probe <workbooks-dir>
//! 每个文件复制到临时目录后直接调 bo_calc::recalculate，逐个打印耗时与结果。
use std::path::PathBuf;
use std::time::Instant;

fn main() -> anyhow::Result<()> {
    let dir = PathBuf::from(std::env::args().nth(1).expect("workbooks dir"));
    let work = std::env::temp_dir().join(format!("s1_perf_{}", std::process::id()));
    std::fs::create_dir_all(&work)?;
    let config = audit_core::bo_calc::BoCalcConfig::from_app();
    println!(
        "config: enabled={} rounds={} max_visits={:?} strip={}",
        config.enabled, config.rounds, config.max_visits, config.strip_style_cells
    );

    let mut files: Vec<PathBuf> = std::fs::read_dir(&dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
        })
        .collect();
    files.sort();

    let mut total = std::time::Duration::ZERO;
    for (index, source) in files.iter().enumerate() {
        let copy = work.join(format!("{index}.xlsx"));
        std::fs::copy(source, &copy)?;
        let started = Instant::now();
        match audit_core::bo_calc::recalculate(&copy, &config) {
            Ok(outcome) => {
                let elapsed = started.elapsed();
                total += elapsed;
                println!(
                    "OK   {:>6}ms | changed={:>6} | open={:?} calc={:?} save={:?} | {}",
                    elapsed.as_millis(),
                    outcome.changed_cells_total,
                    outcome.open_elapsed.as_millis(),
                    outcome.calc_elapsed.as_millis(),
                    outcome.save_elapsed.as_millis(),
                    source.file_name().unwrap_or_default().to_string_lossy()
                );
            }
            Err(error) => {
                let elapsed = started.elapsed();
                total += elapsed;
                println!(
                    "FAIL {:>6}ms | {:#} | {}",
                    elapsed.as_millis(),
                    error,
                    source.file_name().unwrap_or_default().to_string_lossy()
                );
            }
        }
    }
    println!("\n合计 {}ms / {} 个文件", total.as_millis(), files.len());
    let _ = std::fs::remove_dir_all(&work);
    Ok(())
}
