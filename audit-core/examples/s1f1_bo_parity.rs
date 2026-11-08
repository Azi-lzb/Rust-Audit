//! S1-F01 校验区域重算的引擎对拍：betteroffice vs COM/LibreOffice。
//!
//! 每个模板复制一份工作副本，只走「重算 → 扫描校验区域」两段——正好隔离
//! 「计算引擎」这一个变量（不做公式复制等与引擎无关的步骤）。
//!
//! 用法：
//! ```text
//! s1f1_bo_parity <templates_dir> <work_dir> <route:auto|off>
//! s1f1_bo_parity --compare <a.json> <b.json>
//! ```

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::time::Instant;

use anyhow::{bail, Context, Result};
use audit_core::s1_recalculation::S1RecalculationSession;
use audit_core::s1_validation_formula_scan::S1ValidationFormulaScanPlan;
use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize)]
struct TemplateResult {
    template: String,
    ok: bool,
    engine: String,
    recalc_elapsed_ms: u128,
    scan_status: String,
    hits: usize,
    cells: Vec<CellHit>,
    error: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
struct CellHit {
    sheet: String,
    cell: String,
    value: String,
    formula: String,
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    if args.len() >= 2 && args[1] == "--compare" {
        return compare(
            args.get(2).context("缺少左路由 JSON")?,
            args.get(3).context("缺少右路由 JSON")?,
        );
    }
    if args.len() < 4 {
        bail!("用法: s1f1_bo_parity <templates_dir> <work_dir> <route:auto|off> | --compare <a.json> <b.json>");
    }
    run(
        PathBuf::from(&args[1]),
        PathBuf::from(&args[2]),
        args[3].clone(),
    )
}

fn run(templates_dir: PathBuf, work_dir: PathBuf, route: String) -> Result<()> {
    match route.as_str() {
        "auto" => std::env::remove_var("PYTOOLS_BO_CALC"),
        "off" => std::env::set_var("PYTOOLS_BO_CALC", "off"),
        other => bail!("未知路由：{other}（auto|off）"),
    }
    let mut templates: Vec<PathBuf> = std::fs::read_dir(&templates_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
        })
        .collect();
    templates.sort();
    eprintln!("发现 {} 个模板（route={route}）", templates.len());
    std::fs::create_dir_all(&work_dir)?;

    // 一个会话复用于整批（COM 路径的既有语义）；betteroffice 是进程内的。
    // COM 会话 worker 是主程序的私有子命令（__audit-native-com-session）；
    // 本 example 自身没有该入口，显式指定产品 exe（AUDIT_EXE 覆盖）。
    // Linux 没有原生 COM：不构造会话，走 recalculate_s1_formula_copy（LO sidecar）。
    #[cfg(windows)]
    let mut session = {
        let audit_exe = std::env::var_os("AUDIT_EXE")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("dist/open-data-audit.exe"));
        let audit_exe = std::fs::canonicalize(&audit_exe)
            .with_context(|| format!("定位 Audit 主程序（COM worker）：{}", audit_exe.display()))?;
        Some(S1RecalculationSession::new_native_with_executable(
            &audit_exe, "auto",
        )?)
    };
    #[cfg(not(windows))]
    let mut session: Option<S1RecalculationSession> = None;

    let mut results = Vec::new();
    for template in &templates {
        let name = template
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| template.display().to_string());
        eprintln!("== {name}");
        let result = run_one(&mut session, template, &work_dir, &name);
        match result {
            Ok(result) => {
                eprintln!(
                    "   {} | engine={} | recalc={}ms | hits={}",
                    result.scan_status, result.engine, result.recalc_elapsed_ms, result.hits
                );
                results.push(result);
            }
            Err(error) => {
                let message = format!("{error:#}");
                eprintln!("   失败：{}", message.chars().take(200).collect::<String>());
                results.push(TemplateResult {
                    template: name,
                    ok: false,
                    engine: String::new(),
                    recalc_elapsed_ms: 0,
                    scan_status: "FAILED".to_owned(),
                    hits: 0,
                    cells: Vec::new(),
                    error: Some(message),
                });
            }
        }
    }

    std::fs::create_dir_all(&work_dir)?;
    let output = work_dir.join(format!("{route}.json"));
    std::fs::write(&output, serde_json::to_vec_pretty(&results)?)?;
    eprintln!("结果：{}", output.display());
    Ok(())
}

fn run_one(
    session: &mut Option<S1RecalculationSession>,
    template: &Path,
    work_dir: &Path,
    name: &str,
) -> Result<TemplateResult> {
    let stem = name.trim_end_matches(".xlsx");
    let source_copy = work_dir.join(format!("{stem}.copy.xlsx"));
    let output_copy = work_dir.join(format!("{stem}.recalculated.xlsx"));
    let _ = std::fs::remove_file(&output_copy);
    std::fs::copy(template, &source_copy)
        .with_context(|| format!("复制模板为工作副本：{}", source_copy.display()))?;

    let started = Instant::now();
    let report = match session {
        Some(session) => session.recalculate(&source_copy, &output_copy)?,
        None => audit_core::s1_recalculation::recalculate_s1_formula_copy(
            std::env::var_os("AUDIT_SIDECAR")
                .map(PathBuf::from)
                .as_deref(),
            &source_copy,
            &output_copy,
            "auto",
        )?,
    };
    let recalc_elapsed_ms = started.elapsed().as_millis();

    let plan = S1ValidationFormulaScanPlan::load(&[template.to_path_buf()])?;
    let scan = plan.scan(&output_copy)?;
    let mut cells: Vec<CellHit> = scan
        .errors
        .iter()
        .map(|error| CellHit {
            sheet: error.sheet.clone(),
            cell: error.cell.clone(),
            value: error.value.clone(),
            formula: error.formula.clone(),
        })
        .collect();
    cells.sort();

    Ok(TemplateResult {
        template: name.to_owned(),
        ok: true,
        engine: report.engine,
        recalc_elapsed_ms,
        scan_status: scan.status.to_owned(),
        hits: cells.len(),
        cells,
        error: None,
    })
}

fn compare(left_path: &str, right_path: &str) -> Result<()> {
    let left: Vec<TemplateResult> = serde_json::from_str(&std::fs::read_to_string(left_path)?)?;
    let right: Vec<TemplateResult> = serde_json::from_str(&std::fs::read_to_string(right_path)?)?;
    let right_by_name: BTreeMap<&str, &TemplateResult> = right
        .iter()
        .map(|result| (result.template.as_str(), result))
        .collect();

    let (mut fully_aligned, mut float_only, mut substantive) = (0usize, 0usize, 0usize);
    let (mut both_failed, mut mismatch) = (0usize, 0usize);
    let mut rows = Vec::new();

    for left_result in &left {
        let Some(right_result) = right_by_name.get(left_result.template.as_str()) else {
            continue;
        };
        if !left_result.ok && !right_result.ok {
            both_failed += 1;
            rows.push(format!(
                "{:<52} | 双方计算前拦截 | {}",
                left_result.template,
                shorten(left_result.error.as_deref().unwrap_or(""), 46)
            ));
            continue;
        }
        if left_result.ok != right_result.ok {
            mismatch += 1;
            rows.push(format!(
                "{:<52} | 失败态不一致 | BO_ok={} COM_ok={}",
                left_result.template, left_result.ok, right_result.ok
            ));
            continue;
        }
        let mut left_keys: Vec<(&str, &str)> = left_result
            .cells
            .iter()
            .map(|cell| (cell.sheet.as_str(), cell.cell.as_str()))
            .collect();
        let mut right_keys: Vec<(&str, &str)> = right_result
            .cells
            .iter()
            .map(|cell| (cell.sheet.as_str(), cell.cell.as_str()))
            .collect();
        left_keys.sort();
        right_keys.sort();

        let count_match = left_result.hits == right_result.hits;
        let set_match = left_keys == right_keys;
        let mut float_diffs = 0usize;
        let mut real_diffs: Vec<String> = Vec::new();
        if set_match {
            for (l, r) in left_result.cells.iter().zip(right_result.cells.iter()) {
                if l.value != r.value {
                    if float_text_equivalent(&l.value, &r.value) {
                        float_diffs += 1;
                    } else {
                        real_diffs.push(format!(
                            "{}!{} value: BO={} COM={}",
                            l.sheet, l.cell, l.value, r.value
                        ));
                    }
                }
                if l.formula != r.formula {
                    real_diffs.push(format!("{}!{} formula differs", l.sheet, l.cell));
                }
            }
        } else {
            real_diffs.push(format!(
                "cell set differs: BO={} COM={}",
                left_result.hits, right_result.hits
            ));
        }

        if count_match && set_match && real_diffs.is_empty() {
            if float_diffs > 0 {
                float_only += 1;
            } else {
                fully_aligned += 1;
            }
            rows.push(format!(
                "{:<52} | hits {:>4} | subst 0 | float {:>3} | BO {:>7}ms | COM {:>7}ms | {}",
                left_result.template,
                left_result.hits,
                float_diffs,
                left_result.recalc_elapsed_ms,
                right_result.recalc_elapsed_ms,
                engine_tag(&left_result.engine)
            ));
        } else {
            substantive += 1;
            rows.push(format!(
                "{:<52} | subst {} | BO={} COM={} | {}",
                left_result.template,
                real_diffs.len(),
                left_result.hits,
                right_result.hits,
                shorten(&real_diffs[0], 70)
            ));
        }
    }

    println!("{}", rows.join("\n"));
    println!();
    println!("完全对齐（含仅浮点）：{}", fully_aligned + float_only);
    println!("  逐字段完全相同：{fully_aligned}");
    println!("  仅浮点文本差异：{float_only}");
    println!("双方计算前拦截：{both_failed}");
    println!("实质差异：{substantive}");
    println!("失败态不一致：{mismatch}");
    let total = left.len();
    let pass = substantive == 0 && mismatch == 0;
    println!(
        "==> {}/{} 通过（实质差异 0 且失败态一致）",
        total.saturating_sub(substantive + mismatch),
        total
    );
    if !pass {
        bail!("对拍未通过：{substantive} 个实质差异、{mismatch} 个失败态不一致");
    }
    Ok(())
}

fn engine_tag(engine: &str) -> &str {
    if engine.contains("betteroffice") {
        "betteroffice"
    } else if engine.contains("回退") {
        "fallback"
    } else {
        "com/lo"
    }
}

fn shorten(text: &str, limit: usize) -> String {
    text.chars().take(limit).collect()
}

/// 15 位 vs 17 位浮点文本等价判定：解析为 f64 后相对误差 < 1e-14。
fn float_text_equivalent(left: &str, right: &str) -> bool {
    let (Ok(left), Ok(right)) = (left.trim().parse::<f64>(), right.trim().parse::<f64>()) else {
        return false;
    };
    if left == right {
        return true;
    }
    let scale = left.abs().max(right.abs()).max(1.0);
    (left - right).abs() / scale < 1e-14
}
