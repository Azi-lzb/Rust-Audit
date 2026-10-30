//! Windows native Office COM worker and direct LibreOffice fallback for S1-F1.
//!
//! The worker is the Audit executable itself, started once per batch. It calls
//! Office through native COM and stays isolated so a hung COM call can be
//! terminated without hanging the GUI. No PowerShell or lo-convert executable
//! is involved in this Windows route.

use crate::native_office_com::ExcelComSession;
use anyhow::{bail, ensure, Context, Result};
use excel_reader::WorkbookDump;
use serde_json::Value;
use std::{
    env, fs,
    io::{BufRead, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
    thread,
    time::{Duration, Instant},
};

/// Handle the private worker command before the GUI or CLI parses arguments.
pub fn maybe_run_native_com_worker() -> Option<Result<()>> {
    let mut arguments = env::args().skip(1);
    if arguments.next().as_deref() != Some("__audit-native-com-session") {
        return None;
    }
    let engine = arguments.next().unwrap_or_else(|| "com".to_owned());
    Some(run_native_com_worker(&engine))
}

fn run_native_com_worker(engine: &str) -> Result<()> {
    let names: &[&str] = match engine {
        "excel" => &["Excel.Application"],
        "wps" => &["ket.Application", "KET.Application"],
        "com" | "auto" => &["Excel.Application", "ket.Application", "KET.Application"],
        _ => bail!("不支持的原生 Office COM 引擎：{engine}"),
    };
    let session_create_started = Instant::now();
    let (mut session, initialization_error) = match ExcelComSession::create(names) {
        Ok(session) => (Some(session), None),
        Err(error) => (None, Some(format!("{error:#}"))),
    };
    println!(
        "{}",
        serde_json::json!({
            "status": "session_ready",
            "com_available": session.is_some(),
            "com_error": initialization_error.as_deref(),
            "engine": session.as_ref().map(|item| item.application_name()),
            "engine_process_id": session.as_ref().and_then(|item| item.process_id()),
            "engine_process_owned": session.as_ref().is_some_and(|item| item.owns_process()),
            "engine_process_created": session.as_ref().and_then(|item| item.process_created()),
            "session_create_ms": session_create_started.elapsed().as_millis() as u64,
        })
    );
    std::io::stdout().flush()?;

    for line in std::io::stdin().lock().lines() {
        let response = (|| -> Result<Option<Value>> {
            let request: Value = serde_json::from_str(&line?).context("解析原生 COM 请求")?;
            if request.get("type").and_then(Value::as_str) == Some("shutdown") {
                return Ok(None);
            }
            ensure!(
                request.get("type").and_then(Value::as_str) == Some("recalculate"),
                "原生 COM 请求类型无效"
            );
            let mode = request.get("mode").and_then(Value::as_str).unwrap_or("full");
            let input = PathBuf::from(
                request
                    .get("input")
                    .and_then(Value::as_str)
                    .context("缺少重算输入")?,
            );
            let output = PathBuf::from(
                request
                    .get("output")
                    .and_then(Value::as_str)
                    .context("缺少重算输出")?,
            );
            ensure!(input.is_file(), "原生 COM 输入工作簿不存在");
            ensure!(
                !output.exists() && input != output,
                "原生 COM 输出已存在或与输入相同"
            );
            let stage = output.with_extension("com-progress.txt");
            let session = session.as_mut().with_context(|| {
                format!(
                    "Office COM 不可用：{}",
                    initialization_error.as_deref().unwrap_or("未知原因")
                )
            })?;
            let stages =
                session.recalculate_with_mode(&input, &output, recalculation_timeout(), &stage, mode)?;
            Ok(Some(serde_json::json!({
                "status": "recalculated",
                "engine": session.application_name(),
                "engine_process_id": session.process_id(),
                "stages_ms": stages,
            })))
        })();
        match response {
            Ok(Some(value)) => println!("{value}"),
            Ok(None) => break,
            Err(error) => println!(
                "{}",
                serde_json::json!({"status":"failed","error":format!("{error:#}")})
            ),
        }
        std::io::stdout().flush()?;
    }
    // Timing evidence: the graceful dispatch-release teardown was measured at
    // ~2.3s per batch (Excel teardown blocking the STA thread on Release),
    // while the same shutdown costs the pywin32 channel ~0.16s. The worker
    // owns its Office process and the parent already tracks that PID for
    // cleanup, so exiting the process here (letting the OS drop the COM
    // references, which makes Excel self-exit asynchronously) removes the
    // stall without leaving the owned Excel process behind.
    let shutdown_started = Instant::now();
    drop(session);
    eprintln!(
        "{}",
        serde_json::json!({
            "type": "progress",
            "step": "native_com_shutdown",
            "drop_ms": shutdown_started.elapsed().as_millis() as u64,
        })
    );
    Ok(())
}

pub(crate) fn prepare_engine_input(input: &Path, directory: &Path, com: bool) -> Result<PathBuf> {
    let cacheless = directory.join("formula-cache-cleared.xlsx");
    let prepared = directory.join("recalculation-request.xlsx");
    ooxml_patch::clear_formula_caches_for_recalculation(input, &cacheless)
        .context("清除原有公式缓存")?;
    ooxml_patch::patch_workbook_calculation_properties(
        &cacheless,
        &prepared,
        ooxml_patch::WorkbookCalculationProperties {
            // Excel can remain xlPending with these flags set. COM explicitly
            // invokes CalculateFull; LibreOffice uses the load-time flags.
            full_calc_on_load: Some(!com),
            force_full_calc: Some(!com),
        },
    )
    .context("准备重算工作簿计算属性")?;
    let properties = ooxml_patch::read_workbook_calculation_properties(&prepared)?;
    ensure!(
        properties.full_calc_on_load == Some(!com) && properties.force_full_calc == Some(!com),
        "重算工作簿计算属性验证失败"
    );
    let inventory = excel_reader::inspect_formula_cache_nodes(&prepared)?;
    ensure!(inventory.formula_cell_count > 0, "工作簿没有可重算的公式");
    ensure!(
        inventory.cached_formula_cell_count == 0,
        "重算输入仍包含旧公式缓存"
    );
    Ok(prepared)
}

pub(crate) fn run_libreoffice(input: &Path, directory: &Path) -> Result<PathBuf> {
    use std::os::windows::process::CommandExt;

    let output_dir = directory.join("lo-output");
    let profile = directory.join("lo-profile");
    fs::create_dir(&output_dir)?;
    fs::create_dir(&profile)?;
    let stdout_path = directory.join("lo-stdout.txt");
    let stderr_path = directory.join("lo-stderr.txt");
    let stdout = fs::File::create(stdout_path)?;
    let stderr = fs::File::create(&stderr_path)?;
    let mut child = Command::new("soffice")
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg("--nofirststartwizard")
        .arg(format!("-env:UserInstallation={}", file_uri(&profile)?))
        .arg("--convert-to")
        .arg("xlsx:Calc MS Excel 2007 XML")
        .arg("--outdir")
        .arg(&output_dir)
        .arg(file_uri(input)?)
        .creation_flags(0x0800_0000)
        .stdout(Stdio::from(stdout))
        .stderr(Stdio::from(stderr))
        .spawn()
        .context("启动 LibreOffice 重算失败")?;
    let started = Instant::now();
    let status = loop {
        match child.try_wait() {
            Ok(Some(status)) => break status,
            Ok(None) if started.elapsed() < recalculation_timeout() => {
                thread::sleep(Duration::from_millis(100));
            }
            result => {
                let _ = Command::new("taskkill.exe")
                    .args(["/PID", &child.id().to_string(), "/T", "/F"])
                    .creation_flags(0x0800_0000)
                    .stdout(Stdio::null())
                    .stderr(Stdio::null())
                    .status();
                let _ = child.kill();
                let _ = child.wait();
                match result {
                    Err(error) => return Err(error).context("等待 LibreOffice 重算"),
                    _ => bail!(
                        "LibreOffice 重算超过 {} 秒",
                        recalculation_timeout().as_secs()
                    ),
                }
            }
        }
    };
    ensure!(
        status.success(),
        "LibreOffice 重算失败（{status}）：{}",
        fs::read_to_string(stderr_path).unwrap_or_default().trim()
    );
    let output = output_dir.join(input.file_name().context("重算输入缺少文件名")?);
    ensure!(
        output.is_file() && fs::metadata(&output)?.len() > 0,
        "LibreOffice 未生成重算结果"
    );
    Ok(output)
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
    ensure!(path.is_absolute(), "LibreOffice 路径必须是绝对路径");
    let path = path.to_string_lossy().replace('\\', "/");
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

/// Restore literal Excel errors that an engine represented as formulas, using
/// the same narrow policy as the existing recalculation sidecar.
pub(crate) fn normalize_engine_output(
    input: &Path,
    candidate: &Path,
    directory: &Path,
) -> Result<PathBuf> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)?.workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(candidate)?.workbook;
    if excel_reader::formula_texts_equivalent_after_recalculation(&source, &result) {
        return Ok(candidate.to_path_buf());
    }
    let promoted = promoted_error_constants(&source, &result)?;
    ensure!(!promoted.is_empty(), "重算改变了工作表或公式结构");
    let output = directory.join("recalculated-error-constants-restored.xlsx");
    ooxml_patch::restore_promoted_error_constants(input, candidate, &output, &promoted)?;
    let restored = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&output)?.workbook;
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &restored),
        "恢复错误常量后重算结果的公式结构仍不一致"
    );
    Ok(output)
}

fn promoted_error_constants(
    source: &WorkbookDump,
    result: &WorkbookDump,
) -> Result<Vec<ooxml_patch::PromotedErrorConstant>> {
    ensure!(
        source.sheets.len() == result.sheets.len(),
        "重算改变了工作表数量"
    );
    let mut promoted = Vec::new();
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        ensure!(before.name == after.name, "重算改变了工作表名称或顺序");
        for (cell, formula) in &before.formulas {
            let result_formula = after
                .formulas
                .get(cell)
                .with_context(|| format!("重算丢失公式 {}!{cell}", before.name))?;
            ensure!(
                excel_reader::canonicalize_formula_boolean_literals(formula)
                    == excel_reader::canonicalize_formula_boolean_literals(result_formula),
                "重算改变公式 {}!{cell}",
                before.name
            );
        }
        for (cell, formula) in &after.formulas {
            if before.formulas.contains_key(cell) {
                continue;
            }
            ensure!(
                matches!(
                    formula.as_str(),
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
                ),
                "重算新增了不受支持的公式 {}!{cell}",
                before.name
            );
            let source_value = before
                .values
                .get(cell)
                .context("新增公式没有对应的源错误值")?;
            let result_value = after.values.get(cell).context("新增公式没有可读的错误值")?;
            ensure!(
                source_value.get("type").and_then(Value::as_str) == Some("error")
                    && result_value.get("type").and_then(Value::as_str) == Some("error")
                    && source_value == result_value,
                "新增公式没有保留源错误值 {}!{cell}",
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
