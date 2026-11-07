//! Explicit Office/LibreOffice conversion sidecar. The product binary never
//! links or imports this crate; it starts this executable only for a requested
//! same-category conversion.

use anyhow::{anyhow, bail, ensure, Context, Result};
use encoding_rs::GBK;
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue as ReaderCellValue, WorkbookDump,
};
use std::{
    collections::BTreeMap,
    env, fs,
    io::{BufRead, Write},
    path::{Path, PathBuf},
    process::{Child, Command, Output, Stdio},
    thread,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

#[cfg(windows)]
#[path = "../../../audit-core/src/native_office_com.rs"]
mod excel_com;
#[cfg(not(windows))]
mod excel_com {
    pub struct ExcelComSession;
}
mod lo_batch;

#[cfg(test)]
const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
#[cfg(test)]
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
#[cfg(test)]
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn main() {
    if let Err(error) = run() {
        eprintln!("lo-convert: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let mut args = env::args().skip(1);
    let Some(command) = args.next() else {
        print_usage();
        bail!("missing command");
    };
    if command == "--help" || command == "-h" {
        print_usage();
        return Ok(());
    }
    if command == "recalculate" {
        return run_recalculate(args.collect());
    }
    if command == "recalculate-session" {
        return run_recalculate_session(args.collect());
    }
    if command == "recalculate-batch" {
        return run_recalculate_batch(args.collect());
    }
    if command == "__recalculate-com-worker" {
        return run_com_worker(args.collect());
    }
    if command == "convert-batch" {
        return run_convert_batch(args.collect());
    }
    ensure!(command == "convert", "unsupported command {command:?}");
    let mut values = BTreeMap::new();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let input = PathBuf::from(values.get("input").context("--input is required")?);
    let output = PathBuf::from(values.get("output").context("--output is required")?);
    let target = values
        .get("to")
        .context("--to is required")?
        .to_ascii_lowercase();
    let mode = values.get("mode").map(String::as_str).unwrap_or("fidelity");
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(mode == "fidelity", "only fidelity conversion is implemented; native value-only conversion is not available yet");
    ensure!(
        engine == "auto" || engine == "lo" || engine == "com",
        "--engine must be auto, lo or com"
    );
    ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    let source_ext = extension(&input)?;
    ensure!(
        same_category(&source_ext, &target),
        "conversion must stay within the same document category"
    );
    ensure!(
        (matches!(target.as_str(), "xlsx")
            && matches!(source_ext.as_str(), "csv" | "xls" | "xlsb" | "xlsm" | "et"))
            || (target == "docx" && source_ext == "doc"),
        "unsupported conversion: .{source_ext} -> .{target}"
    );
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)?;
    }
    ensure!(
        !output.exists(),
        "refusing to overwrite existing output: {}",
        output.display()
    );

    let temp_guard = TempDirGuard(unique_temp_dir(
        output.parent().unwrap_or_else(|| Path::new(".")),
    )?);
    let temp_dir = &temp_guard.0;
    let converted = temp_dir.join(format!("{}.{}", file_stem(&input), target));
    let selected = if engine == "auto" {
        if cfg!(windows) {
            "com"
        } else {
            "lo"
        }
    } else {
        engine
    };
    match selected {
        "lo" => run_libreoffice(&input, temp_dir, &target)?,
        "com" if cfg!(windows) => run_com(&input, &converted, &target, &source_ext)?,
        "com" => bail!(
            "COM conversion is only available on Windows; install LibreOffice or use --engine lo"
        ),
        _ => unreachable!(),
    }
    let candidate = if converted.exists() {
        converted
    } else {
        locate_converted(temp_dir, &target)?
    };
    ensure!(
        candidate.is_file() && fs::metadata(&candidate)?.len() > 0,
        "conversion engine did not create a non-empty output file"
    );
    let validation =
        validate_conversion(&input, &candidate, &source_ext, &target, selected, temp_dir)?;
    fs::rename(&candidate, &output)
        .with_context(|| format!("commit validated output {}", output.display()))?;
    println!("{{\"status\":\"converted\",\"engine\":\"{selected}\",\"input\":{},\"output\":{},\"validation\":{}}}", json_string(&input.display().to_string()), json_string(&output.display().to_string()), validation);
    Ok(())
}

fn run_recalculate(arguments: Vec<String>) -> Result<()> {
    run_recalculate_with_session(arguments, None, None, None)
}

fn run_recalculate_with_session(
    arguments: Vec<String>,
    mut com_session: Option<&mut excel_com::ExcelComSession>,
    com_initialization_error: Option<&str>,
    mut lo_batch: Option<&mut lo_batch::LoWarmProfileSession>,
) -> Result<()> {
    if arguments
        .iter()
        .any(|argument| argument == "--help" || argument == "-h")
    {
        eprintln!("Usage: lo-convert recalculate --input <xlsx> --output <new.xlsx> [--engine auto|lo|com]");
        return Ok(());
    }
    let mut values = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        ensure!(
            matches!(name.as_str(), "input" | "output" | "engine"),
            "unknown recalculate option --{name}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let input = PathBuf::from(values.get("input").context("--input is required")?);
    let output = PathBuf::from(values.get("output").context("--output is required")?);
    ensure!(
        input.is_file(),
        "input workbook does not exist: {}",
        input.display()
    );
    ensure!(
        input
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
        "formula recalculation currently supports .xlsx only"
    );
    ensure!(
        output
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
        "formula recalculation output must be .xlsx"
    );
    let canonical_input = fs::canonicalize(&input).context("resolve recalculation input")?;
    ensure!(
        output != input && (!output.exists() || fs::canonicalize(&output)? != canonical_input),
        "recalculation requires a distinct output path"
    );
    if let Some(parent) = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)?;
    }
    ensure!(
        !output.exists(),
        "refusing to overwrite existing recalculation output: {}",
        output.display()
    );
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(
        engine == "auto"
            || engine == "lo"
            || engine == "com"
            || engine == "excel"
            || engine == "wps",
        "--engine must be auto, lo, com, excel or wps"
    );
    let selected_engines = recalculation_engine_order(engine, cfg!(windows))?;

    // Keep paths passed to Excel COM short. The destination can
    // be nested deeply under a user's selected output folder; intermediates in
    // that folder may exceed legacy Office path limits.
    let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
    let input_dir = guard.0.join("input");
    let engine_output_dir = guard.0.join("engine-output");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&engine_output_dir)?;
    let cacheless_input = input_dir.join("formula-cache-cleared.xlsx");
    let recalc_input = input_dir.join("recalculation-request.xlsx");
    let cache_report =
        ooxml_patch::clear_formula_caches_for_recalculation(&input, &cacheless_input)
            .context("remove prior formula caches from the recalculation request copy")?;
    ooxml_patch::patch_workbook_calculation_properties(
        &cacheless_input,
        &recalc_input,
        ooxml_patch::WorkbookCalculationProperties {
            full_calc_on_load: Some(true),
            force_full_calc: Some(true),
        },
    )
    .context("prepare explicit full recalculation request")?;
    let mut attempts = Vec::new();
    let mut accepted = None;
    // 批次会话复用时记录本次会话处理阶段与实例归属；单发路径为 None。
    let mut lo_batch_stages: Option<serde_json::Value> = None;
    for selected in selected_engines {
        let attempt_dir = engine_output_dir.join(selected);
        fs::create_dir_all(&attempt_dir)?;
        let candidate = match selected {
            "lo" => attempt_dir.join("recalculation-request.xlsx"),
            "com" | "excel" | "wps" => attempt_dir.join("recalculated.xlsx"),
            _ => unreachable!("validated engine order contains an unknown engine"),
        };
        let attempt = match selected {
            "lo" => {
                if let Some(session) = lo_batch.as_deref_mut() {
                    // 批次会话复用：同一 Rust 拥有的 LibreOffice 实例顺序处理；
                    // 阶段计时由会话内的批次驱动逐本记录。
                    session
                        .recalculate(&recalc_input, &candidate, recalculation_timeout())
                        .map(|(engine_label, stages)| {
                            lo_batch_stages = Some(stages);
                            engine_label
                        })
                } else {
                    run_libreoffice_recalculate(&recalc_input, &attempt_dir)
                }
            }
            "com" | "excel" | "wps" => {
                let com_input = attempt_dir.join("com-recalculation-request.xlsx");
                prepare_com_recalculation_input(&recalc_input, &com_input).and_then(|()| {
                    if let Some(session) = com_session.as_deref_mut() {
                        #[cfg(windows)]
                        {
                            let stage = attempt_dir.join("excel-com-progress.txt");
                            session
                                .recalculate(
                                    &com_input,
                                    &candidate,
                                    recalculation_timeout(),
                                    &stage,
                                )
                                .map(|stages| {
                                    let _ = stages;
                                    session.application_name().to_owned()
                                })
                        }
                        #[cfg(not(windows))]
                        {
                            let _ = session;
                            bail!("native Excel COM recalculation is only available on Windows")
                        }
                    } else if let Some(error) = com_initialization_error {
                        bail!("Office COM session initialization failed: {error}")
                    } else {
                        run_com_recalculate(&com_input, &candidate, &attempt_dir, selected)
                    }
                })
            }
            _ => unreachable!("validated engine order contains an unknown engine"),
        }
        .and_then(|actual_engine| {
            require_engine_output(&candidate)?;
            let (validated_candidate, restored_error_constants) =
                normalize_promoted_error_constants(&input, &candidate, &attempt_dir)?;
            let mut validation = validate_recalculated_workbook(&input, &validated_candidate)?;
            let mut validation_json: serde_json::Value = serde_json::from_str(&validation)?;
            validation_json["promoted_error_constants_restored"] =
                serde_json::json!(restored_error_constants);
            validation = serde_json::to_string(&validation_json)?;
            Ok((actual_engine, validated_candidate, validation))
        });
        match attempt {
            Ok((actual_engine, validated_candidate, validation)) => {
                attempts.push(serde_json::json!({
                    "requested_engine": selected,
                    "engine": actual_engine,
                    "status": "validated",
                    "validation": serde_json::from_str::<serde_json::Value>(&validation)
                        .unwrap_or_else(|_| serde_json::json!({"detail": validation}))
                }));
                accepted = Some((actual_engine, validated_candidate, validation));
                break;
            }
            Err(error) => attempts.push(serde_json::json!({
                "requested_engine": selected,
                "status": "failed_or_invalid",
                "reason": format!("{error:#}")
            })),
        }
    }
    let (actual_engine, candidate, validation) = accepted.ok_or_else(|| {
        anyhow!(
            "no configured calculation engine produced a validated workbook; attempts={}",
            serde_json::to_string(&attempts).unwrap_or_else(|_| "[]".to_owned())
        )
    })?;
    publish_recalculated_output(&candidate, &output)?;
    let lo_batch_json = match lo_batch_stages {
        Some(ref stages) => format!(",\"lo_batch\":{}", serde_json::to_string(stages)?),
        None => String::new(),
    };
    println!("{{\"status\":\"recalculated\",\"engine\":{},\"input\":{},\"output\":{},\"formula_cache_values_cleared_before_engine\":{},\"attempts\":{},\"validation\":{}{lo_batch_json}}}", json_string(&actual_engine), json_string(&input.display().to_string()), json_string(&output.display().to_string()), cache_report.cache_values_removed, serde_json::to_string(&attempts)?, validation);
    Ok(())
}

fn run_recalculate_session(arguments: Vec<String>) -> Result<()> {
    #[cfg(not(windows))]
    {
        run_recalculate_session_lo(arguments)
    }
    #[cfg(windows)]
    {
        let mut values = BTreeMap::new();
        let mut args = arguments.into_iter();
        while let Some(key) = args.next() {
            ensure!(key.starts_with("--"), "expected an option, got {key:?}");
            let name = key.trim_start_matches('-').to_owned();
            ensure!(
                name == "engine",
                "unknown recalculation-session option --{name}"
            );
            let value = args
                .next()
                .with_context(|| format!("{key} requires a value"))?;
            values.insert(name, value);
        }
        let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
        let engines = recalculation_engine_order(engine, true)?;
        let needs_com = engines.iter().any(|candidate| *candidate != "lo");
        let (mut com_session, com_initialization_error) = if needs_com {
            match excel_com::ExcelComSession::create(com_recalculation_app_names(engine)) {
                Ok(session) => (Some(session), None),
                Err(error) => (None, Some(format!("{error:#}"))),
            }
        } else {
            (None, None)
        };
        let ready = serde_json::json!({
            "status": "session_ready",
            "engine": com_session.as_ref().map(|session| session.application_name()),
            "engine_process_id": com_session.as_ref().and_then(|session| session.process_id()),
            "engine_process_owned": com_session.as_ref().is_some_and(|session| session.owns_process()),
            "com_available": com_session.is_some(),
            "com_error": com_initialization_error.as_deref(),
        });
        println!("{}", serde_json::to_string(&ready)?);
        std::io::stdout().flush()?;

        for line in std::io::stdin().lock().lines() {
            let line = line?;
            let request: serde_json::Value = match serde_json::from_str(&line) {
                Ok(request) => request,
                Err(error) => {
                    println!(
                        "{}",
                        serde_json::json!({"status":"failed","error":format!("invalid session request: {error}")})
                    );
                    std::io::stdout().flush()?;
                    continue;
                }
            };
            if request.get("type").and_then(serde_json::Value::as_str) == Some("shutdown") {
                break;
            }
            let input = request.get("input").and_then(serde_json::Value::as_str);
            let output = request.get("output").and_then(serde_json::Value::as_str);
            let (Some(input), Some(output)) = (input, output) else {
                println!(
                    "{}",
                    serde_json::json!({"status":"failed","error":"session request requires input and output paths"})
                );
                std::io::stdout().flush()?;
                continue;
            };
            let recalculate_arguments = vec![
                "--input".to_owned(),
                input.to_owned(),
                "--output".to_owned(),
                output.to_owned(),
                "--engine".to_owned(),
                engine.to_owned(),
            ];
            if let Err(error) = run_recalculate_with_session(
                recalculate_arguments,
                com_session.as_mut(),
                com_initialization_error.as_deref(),
                None,
            ) {
                println!(
                    "{}",
                    serde_json::json!({"status":"failed","error":format!("{error:#}")})
                );
                std::io::stdout().flush()?;
            }
            std::io::stdout().flush()?;
        }
        drop(com_session);
        Ok(())
    }
}

/// Linux 的旧 request/response 重算协议：`--engine auto|lo` 复用 sidecar
/// worker 和独立 LibreOffice profile 初始化，但每个重算请求仍会单独启动
/// 一个 `soffice --convert-to` 命令。S1-F1 GUI 使用 `recalculate-batch`，
/// 避免把此 profile 复用误认为 LibreOffice 进程复用。`--engine
/// com/excel/wps` 在 Linux 不可用，协议保留该诊断供旧调用方使用。
fn run_recalculate_session_lo(arguments: Vec<String>) -> Result<()> {
    let mut values = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        ensure!(
            name == "engine",
            "unknown recalculation-session option --{name}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(
        engine == "auto"
            || engine == "lo"
            || engine == "com"
            || engine == "excel"
            || engine == "wps",
        "--engine must be auto, lo, com, excel or wps"
    );
    let (needs_lo, com_error) = match engine {
        "auto" | "lo" => (true, None),
        _ => (
            false,
            Some("COM recalculation is only available on Windows; use --engine lo"),
        ),
    };
    let mut lo_batch = if needs_lo {
        Some(lo_batch::LoWarmProfileSession::create().context("创建 LibreOffice 批次会话")?)
    } else {
        None
    };
    let ready = serde_json::json!({
        "status": "session_ready",
        "engine": if lo_batch.is_some() { serde_json::json!("LibreOffice") } else { serde_json::Value::Null },
        "engine_process_id": None::<u32>,
        "engine_process_owned": false,
        "com_available": false,
        "com_error": com_error,
        "lo_batch_session": lo_batch.is_some(),
        "lo_profile_owned": lo_batch.is_some(),
        "lo_batch_work_dir": lo_batch
            .as_ref()
            .map(|session| session.work_dir().display().to_string()),
    });
    println!("{}", serde_json::to_string(&ready)?);
    std::io::stdout().flush()?;

    for line in std::io::stdin().lock().lines() {
        let line = line?;
        let request: serde_json::Value = match serde_json::from_str(&line) {
            Ok(request) => request,
            Err(error) => {
                println!(
                    "{}",
                    serde_json::json!({"status":"failed","error":format!("invalid session request: {error}")})
                );
                std::io::stdout().flush()?;
                continue;
            }
        };
        if request.get("type").and_then(serde_json::Value::as_str) == Some("shutdown") {
            break;
        }
        let input = request.get("input").and_then(serde_json::Value::as_str);
        let output = request.get("output").and_then(serde_json::Value::as_str);
        let (Some(input), Some(output)) = (input, output) else {
            println!(
                "{}",
                serde_json::json!({"status":"failed","error":"session request requires input and output paths"})
            );
            std::io::stdout().flush()?;
            continue;
        };
        let recalculate_arguments = vec![
            "--input".to_owned(),
            input.to_owned(),
            "--output".to_owned(),
            output.to_owned(),
            "--engine".to_owned(),
            if engine == "auto" { "lo".to_owned() } else { engine.to_owned() },
        ];
        if let Err(error) = run_recalculate_with_session(
            recalculate_arguments,
            None,
            None,
            lo_batch.as_mut(),
        ) {
            println!(
                "{}",
                serde_json::json!({"status":"failed","error":format!("{error:#}")})
            );
            std::io::stdout().flush()?;
        }
        std::io::stdout().flush()?;
    }
    if let Some(session) = lo_batch {
        // 暖 profile 会话没有常驻进程（每本转换进程自然退出）；
        // 工作目录由 TempDirGuard 在会话结束时清理。
        let _ = session;
    }
    Ok(())
}

/// 批次重算：`recalculate-batch --pairs <tsv> --engine lo`。TSV 每行
/// `输入副本<TAB>期望输出`；一次 soffice 启动顺序转换全部输入，再逐本
/// 规范化、验证并发布。逐本准备（清缓存 + 重算标记）在批次内完成；单本
/// 验证失败不影响其他本的发布，批次级失败（引擎启动/超时）整体报错。
fn run_recalculate_batch(arguments: Vec<String>) -> Result<()> {
    let mut values: BTreeMap<String, String> = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        ensure!(
            matches!(name.as_str(), "pairs" | "engine"),
            "unknown recalculate-batch option --{name}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let engine = values.get("engine").map(String::as_str).unwrap_or("lo");
    ensure!(engine == "lo", "batch recalculation only supports LibreOffice (--engine lo)");
    let pairs_path = PathBuf::from(values.get("pairs").context("--pairs is required")?);
    let content = fs::read_to_string(&pairs_path)
        .with_context(|| format!("read batch pairs {}", pairs_path.display()))?;
    let mut declared = Vec::new();
    for line in content.lines() {
        if line.trim().is_empty() {
            continue;
        }
        let mut parts = line.split('\t');
        let input = parts.next().context("batch pair missing input")?;
        let output = parts.next().context("batch pair missing output")?;
        ensure!(
            parts.next().is_none(),
            "batch pair line has extra fields: {line:?}"
        );
        declared.push((PathBuf::from(input), PathBuf::from(output)));
    }
    ensure!(!declared.is_empty(), "batch pairs file is empty");
    for (input, output) in &declared {
        ensure!(
            input.is_file(),
            "batch input workbook does not exist: {}",
            input.display()
        );
        ensure!(
            !output.exists(),
            "refusing to overwrite existing batch output: {}",
            output.display()
        );
    }

    let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
    let input_dir = guard.0.join("batch-input");
    let output_dir = guard.0.join("batch-output");
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&output_dir)?;
    let mut prepared_pairs = Vec::with_capacity(declared.len());
    for (index, (input, _)) in declared.iter().enumerate() {
        let canonical = fs::canonicalize(input)
            .with_context(|| format!("resolve batch input {}", input.display()))?;
        let cacheless = input_dir.join(format!("book-{index:04}-cacheless.xlsx"));
        let prepared = input_dir.join(format!("book-{index:04}.xlsx"));
        ooxml_patch::clear_formula_caches_for_recalculation(&canonical, &cacheless)
            .with_context(|| format!("clear formula caches for batch book {}", index + 1))?;
        ooxml_patch::patch_workbook_calculation_properties(
            &cacheless,
            &prepared,
            ooxml_patch::WorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            },
        )
        .with_context(|| format!("prepare batch book {}", index + 1))?;
        prepared_pairs.push((prepared, output_dir.join(format!("book-{index:04}.xlsx"))));
    }
    let batch_timeout = recalculation_timeout()
        .saturating_mul(declared.len() as u32)
        .max(recalculation_timeout());
    let convert_started = Instant::now();
    let (engine_name, stages) =
        lo_batch::run_batch_convert(&prepared_pairs, guard.0.as_path(), lo_batch::BatchTarget::Xlsx, batch_timeout)
            .context("batch LibreOffice conversion")?;
    let convert_ms = convert_started.elapsed().as_millis() as u64;

    let mut books = Vec::with_capacity(declared.len());
    let mut failures = 0usize;
    for (index, (input, output)) in declared.iter().enumerate() {
        let candidate = stages
            .get(index)
            .and_then(|stage| stage.get("batch_output"))
            .and_then(serde_json::Value::as_str)
            .map(PathBuf::from);
        let Some(candidate) = candidate else {
            failures += 1;
            books.push(serde_json::json!({"index": index + 1, "status": "failed", "error": "batch did not produce output"}));
            continue;
        };
        let attempt = (|| -> Result<serde_json::Value> {
            require_engine_output(&candidate)?;
            let (validated_candidate, restored_error_constants) =
                normalize_promoted_error_constants(input, &candidate, &guard.0)?;
            let mut validation = validate_recalculated_workbook(input, &validated_candidate)?;
            let mut validation_json: serde_json::Value = serde_json::from_str(&validation)?;
            validation_json["promoted_error_constants_restored"] =
                serde_json::json!(restored_error_constants);
            validation = serde_json::to_string(&validation_json)?;
            publish_recalculated_output(&validated_candidate, output)?;
            Ok(serde_json::from_str(&validation)?)
        })();
        match attempt {
            Ok(validation) => {
                books.push(serde_json::json!({
                    "index": index + 1,
                    "status": "recalculated",
                    "engine": engine_name,
                    "input": input.display().to_string(),
                    "output": output.display().to_string(),
                    "validation": validation,
                }));
            }
            Err(error) => {
                failures += 1;
                books.push(serde_json::json!({
                    "index": index + 1,
                    "status": "failed",
                    "input": input.display().to_string(),
                    "error": format!("{error:#}"),
                }));
            }
        }
    }
    let result = serde_json::json!({
        "status": if failures == 0 { "recalculated" } else { "partial" },
        "engine": engine_name,
        "soffice_conversion_invocations": 1,
        "profile_processes_remaining": 0,
        "convert_ms": convert_ms,
        "books": books,
    });
    println!("{}", serde_json::to_string(&result)?);
    Ok(())
}

fn recalculation_engine_order(engine: &str, windows: bool) -> Result<Vec<&'static str>> {
    match engine {
        "auto" if windows => Ok(vec!["com", "lo"]),
        "auto" => Ok(vec!["lo"]),
        "lo" => Ok(vec!["lo"]),
        // `com` keeps its historical meaning: the Windows COM recalculation
        // stage, which now probes Excel first and then WPS spreadsheets.
        "com" if windows => Ok(vec!["com"]),
        "com" => bail!("COM recalculation is only available on Windows; use --engine lo"),
        "excel" if windows => Ok(vec!["excel"]),
        "excel" => bail!("Excel COM recalculation is only available on Windows; use --engine lo"),
        "wps" if windows => Ok(vec!["wps"]),
        "wps" => bail!("WPS COM recalculation is only available on Windows; use --engine lo"),
        _ => bail!("--engine must be auto, lo, com, excel or wps"),
    }
}

fn recalculation_timeout() -> Duration {
    let seconds = env::var("AUDIT_RECALC_TIMEOUT_SECONDS")
        .ok()
        .and_then(|value| value.parse::<u64>().ok())
        .unwrap_or(300)
        .clamp(5, 3600);
    Duration::from_secs(seconds)
}

fn run_command_with_timeout(
    command: &mut Command,
    output_dir: &Path,
    timeout: Duration,
    label: &str,
) -> Result<Output> {
    let stdout_path = output_dir.join("engine-stdout.log");
    let stderr_path = output_dir.join("engine-stderr.log");
    let stdout_file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&stdout_path)
        .with_context(|| format!("create {label} stdout log"))?;
    let stderr_file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&stderr_path)
        .with_context(|| format!("create {label} stderr log"))?;
    #[cfg(unix)]
    {
        use std::os::unix::process::CommandExt;
        // Give the engine and any helper processes it starts an isolated
        // process group so timeout cleanup can terminate the whole tree.
        command.process_group(0);
    }
    let mut child = command
        .stdout(Stdio::from(stdout_file))
        .stderr(Stdio::from(stderr_file))
        .spawn()
        .with_context(|| format!("start {label}"))?;
    let started = Instant::now();
    let status = loop {
        match child.try_wait() {
            Err(error) => {
                terminate_timed_out_command(&mut child);
                return Err(error).with_context(|| format!("wait for {label}"));
            }
            Ok(Some(status)) => break status,
            Ok(None) if started.elapsed() >= timeout => {
                terminate_timed_out_command(&mut child);
                bail!("{label} exceeded the {} second timeout", timeout.as_secs());
            }
            Ok(None) => thread::sleep(Duration::from_millis(100)),
        }
    };
    let stdout = fs::read(&stdout_path).with_context(|| format!("read {label} stdout log"))?;
    let stderr = fs::read(&stderr_path).with_context(|| format!("read {label} stderr log"))?;
    Ok(Output {
        status,
        stdout,
        stderr,
    })
}

fn terminate_timed_out_command(child: &mut Child) {
    let process_id = child.id();
    #[cfg(windows)]
    terminate_process_tree(process_id);
    #[cfg(unix)]
    terminate_process_group(process_id);
    // Keep this direct-child fallback even after the tree kill attempt; it
    // reaps the process if the platform-specific tree operation fails.
    let _ = child.kill();
    let _ = child.wait();
}

#[cfg(windows)]
fn terminate_process_tree(process_id: u32) {
    if process_id == 0 || process_id == std::process::id() {
        return;
    }
    let _ = Command::new("taskkill.exe")
        .arg("/PID")
        .arg(process_id.to_string())
        .args(["/T", "/F"])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
}

#[cfg(unix)]
fn terminate_process_group(process_id: u32) {
    if process_id == 0 || process_id == std::process::id() {
        return;
    }
    let group_id = format!("-{process_id}");
    let _ = Command::new("kill")
        .args(["-KILL", "--"])
        .arg(group_id)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
}

fn run_libreoffice_recalculate(input: &Path, temp_dir: &Path) -> Result<String> {
    let profile = temp_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let mut command = Command::new("soffice");
    command
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg("--nofirststartwizard")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg("xlsx:Calc MS Excel 2007 XML")
        .arg("--outdir")
        .arg(temp_dir)
        .arg(file_uri(input)?);
    let command = run_command_with_timeout(
        &mut command,
        temp_dir,
        recalculation_timeout(),
        "LibreOffice recalculation",
    )
    .context("LibreOffice sidecar is unavailable or did not finish recalculating")?;
    ensure!(
        command.status.success(),
        "LibreOffice recalculation failed ({}): {}",
        command.status,
        String::from_utf8_lossy(&command.stderr)
    );
    Ok("LibreOffice".to_owned())
}

fn file_uri(path: &Path) -> Result<String> {
    ensure!(
        path.is_absolute(),
        "LibreOffice profile path must be absolute"
    );
    let path = path.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
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

fn libreoffice_user_installation_arg(profile: &Path) -> Result<String> {
    fs::create_dir_all(profile).context("create LibreOffice isolated profile directory")?;
    Ok(format!("-env:UserInstallation={}", file_uri(profile)?))
}

fn require_engine_output(path: &Path) -> Result<()> {
    ensure!(
        path.is_file() && fs::metadata(path)?.len() > 0,
        "calculation engine did not create its expected non-empty output workbook: {}",
        path.display()
    );
    Ok(())
}

fn publish_recalculated_output(candidate: &Path, output: &Path) -> Result<()> {
    let output_parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let publish_staging = TempDirGuard(unique_temp_dir(output_parent)?);
    let staged = publish_staging.0.join("recalculated.xlsx");
    fs::copy(candidate, &staged).with_context(|| {
        format!(
            "stage recalculated output for publication: {}",
            output.display()
        )
    })?;
    require_engine_output(&staged)?;
    // The staged file is beside the requested output, so a hard link publishes
    // atomically and fails if another process occupied the destination.
    fs::hard_link(&staged, output).with_context(|| {
        format!(
            "publish recalculated workbook without overwriting {}",
            output.display()
        )
    })?;
    Ok(())
}

fn prepare_com_recalculation_input(input: &Path, output: &Path) -> Result<()> {
    // Excel reports xlPending indefinitely when both forceFullCalc and
    // fullCalcOnLoad are set, even after CalculateFullRebuild has produced
    // formula values. The formula caches were cleared before engine routing,
    // and COM explicitly runs both full calculation methods, so the COM-only
    // request can safely clear these load-time flags before opening the file.
    ooxml_patch::patch_workbook_calculation_properties(
        input,
        output,
        ooxml_patch::WorkbookCalculationProperties {
            full_calc_on_load: Some(false),
            force_full_calc: Some(false),
        },
    )
    .context("prepare Excel COM input without sticky full-calculation flags")?;
    let properties = ooxml_patch::read_workbook_calculation_properties(output)
        .context("verify Excel COM calculation flags")?;
    ensure!(
        properties.full_calc_on_load == Some(false) && properties.force_full_calc == Some(false),
        "Excel COM input still requests calculation on load"
    );
    let inventory = excel_reader::inspect_formula_cache_nodes(output)
        .context("verify Excel COM input formula caches are cleared")?;
    ensure!(
        inventory.formula_cell_count > 0,
        "workbook contains no formulas"
    );
    ensure!(
        inventory.cached_formula_cell_count == 0,
        "Excel COM input unexpectedly contains cached formula values"
    );
    Ok(())
}

/// COM ProgID probe lists per requested recalculation channel. `com` keeps
/// the historical auto-on-Windows semantics and now probes Excel first, then
/// WPS Spreadsheets; `excel` and `wps` are the user-forced single engines.
fn com_recalculation_app_names(mode: &str) -> &'static [&'static str] {
    match mode {
        "excel" => &["Excel.Application"],
        "wps" => &["ket.Application", "KET.Application"],
        _ => &["Excel.Application", "ket.Application", "KET.Application"],
    }
}

fn run_com_recalculate(input: &Path, output: &Path, temp_dir: &Path, mode: &str) -> Result<String> {
    let pid_file = temp_dir.join("owned-excel-process-id.txt");
    let stage_file = temp_dir.join("excel-com-progress.txt");
    let process_timeout = recalculation_timeout() + Duration::from_secs(5);
    let mut command =
        Command::new(env::current_exe().context("resolve the COM worker executable")?);
    command
        .arg("__recalculate-com-worker")
        .arg("--input")
        .arg(input)
        .arg("--output")
        .arg(output)
        .arg("--mode")
        .arg(mode)
        .arg("--pid-file")
        .arg(&pid_file)
        .arg("--stage-file")
        .arg(&stage_file);
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x0800_0000); // CREATE_NO_WINDOW
    }
    let result = match run_command_with_timeout(
        &mut command,
        temp_dir,
        process_timeout,
        "Office COM formula recalculation",
    ) {
        Ok(result) => result,
        Err(error) => {
            let stage = read_com_stage(&stage_file);
            terminate_owned_excel_process(&pid_file);
            return Err(error).with_context(|| format!("Excel COM progress stage: {stage}"));
        }
    };
    if !result.status.success() {
        let reason = String::from_utf8_lossy(&result.stderr).trim().to_owned();
        let stage = read_com_stage(&stage_file);
        terminate_owned_excel_process(&pid_file);
        bail!("spreadsheet COM recalculation failed at stage {stage}: {reason}");
    }
    let engine = String::from_utf8_lossy(&result.stdout).trim().to_owned();
    ensure!(
        !engine.is_empty(),
        "COM recalculation did not report the selected application"
    );
    let _ = fs::remove_file(pid_file);
    Ok(engine)
}

fn run_com_worker(arguments: Vec<String>) -> Result<()> {
    #[cfg(not(windows))]
    {
        let _ = arguments;
        bail!("native Office COM recalculation is only available on Windows")
    }
    #[cfg(windows)]
    {
        let mut values = BTreeMap::new();
        let mut args = arguments.into_iter();
        while let Some(key) = args.next() {
            ensure!(key.starts_with("--"), "expected an option, got {key:?}");
            let name = key.trim_start_matches('-').to_owned();
            ensure!(
                matches!(
                    name.as_str(),
                    "input" | "output" | "mode" | "pid-file" | "stage-file"
                ),
                "unknown COM worker option --{name}"
            );
            let value = args
                .next()
                .with_context(|| format!("{key} requires a value"))?;
            values.insert(name, value);
        }
        let input = PathBuf::from(values.get("input").context("--input is required")?);
        let output = PathBuf::from(values.get("output").context("--output is required")?);
        let mode = values.get("mode").map(String::as_str).unwrap_or("com");
        let pid_file = PathBuf::from(values.get("pid-file").context("--pid-file is required")?);
        let stage_file = PathBuf::from(
            values
                .get("stage-file")
                .context("--stage-file is required")?,
        );
        fs::write(&stage_file, "creating_excel_application")?;
        let mut session = excel_com::ExcelComSession::create(com_recalculation_app_names(mode))?;
        if session.owns_process() {
            if let Some(process_id) = session.process_id() {
                fs::write(&pid_file, process_id.to_string())?;
            }
        }
        session.recalculate(&input, &output, recalculation_timeout(), &stage_file)?;
        println!("{}", session.application_name());
        Ok(())
    }
}

#[cfg(test)]
fn powershell_compatible_path(path: &Path) -> PathBuf {
    #[cfg(windows)]
    {
        let value = path.to_string_lossy();
        if let Some(unc) = value.strip_prefix(r"\\?\UNC\") {
            return PathBuf::from(format!(r"\\{unc}"));
        }
        if let Some(drive_path) = value.strip_prefix(r"\\?\") {
            return PathBuf::from(drive_path);
        }
    }
    path.to_path_buf()
}

// Retain the former script template only as an offline unit-test fixture.
// Production formula recalculation uses `excel_com::ExcelComSession` directly.
#[cfg(test)]
fn com_recalculate_script(
    input: &Path,
    output: &Path,
    pid_file: &Path,
    stage_file: &Path,
    timeout_seconds: u64,
    app_names: &[&str],
) -> String {
    // Probe each requested ProgID in order; the first successful activation
    // wins and its name is reported so the caller can validate the engine.
    // This block is plain PowerShell (real braces) built at runtime and then
    // passed through the format below as a pre-rendered argument.
    let mut probe = String::new();
    for app_name in app_names {
        probe.push_str(&format!(
            "if ($null -eq $app) {{ try {{ $app=New-Object -ComObject '{}'; $appName='{}' }} catch {{ $lastError=$_.Exception.Message }} }}
",
            app_name, app_name
        ));
    }
    let (unavailable_stage, unavailable_message) = match app_names {
        ["Excel.Application"] => (
            "excel_application_unavailable",
            "Microsoft Excel COM is unavailable",
        ),
        ["ket.Application", "KET.Application"] => (
            "wps_application_unavailable",
            "WPS Spreadsheets COM is unavailable",
        ),
        _ => (
            "office_application_unavailable",
            "No supported Office/WPS COM engine is available",
        ),
    };
    probe.push_str(&format!(
        "if ($null -eq $app) {{ Set-AuditStage '{unavailable_stage}'; throw \"{unavailable_message}: $lastError\" }}
"
    ));

    format!(
        "$ErrorActionPreference='Stop';
$stagePath='{}'; function Set-AuditStage([string]$value) {{ try {{ [System.IO.File]::WriteAllText($stagePath,$value) }} catch {{ }} }}
Add-Type -Namespace Audit -Name ExcelProcess -MemberDefinition '[System.Runtime.InteropServices.DllImport(\"user32.dll\")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);' -ErrorAction SilentlyContinue;
$app=$null; $appName=''; $lastError=''; Set-AuditStage 'creating_excel_application';
{}$book=$null;
try {{ Set-AuditStage 'capturing_excel_process_id'; [uint32]$excelPid=0; [Audit.ExcelProcess]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$excelPid) | Out-Null; [System.IO.File]::WriteAllText('{}',[string]$excelPid); Set-AuditStage 'setting_excel_options'; $app.Visible=$false; $app.DisplayAlerts=$false; Set-AuditStage 'opening_workbook'; $book=$app.Workbooks.Open('{}',$false,$false); Set-AuditStage 'workbook_opened'; Set-AuditStage 'enabling_automatic_calculation'; $app.Calculation=-4105; Set-AuditStage 'calculating_full'; $app.CalculateFull(); Set-AuditStage 'full_calculation_returned'; if ($book.Connections.Count -gt 0) {{ Set-AuditStage 'waiting_for_async_queries'; try {{ $app.CalculateUntilAsyncQueriesDone() }} catch {{ Set-AuditStage 'async_query_wait_failed'; throw ('Excel async query wait failed: ' + $_.Exception.Message) }} }}; $deadline=[DateTime]::UtcNow.AddSeconds({}); $lastCalculationState=-1; $pendingFullRebuildRetried=$false; while ($true) {{ $calculationState=[int]$app.CalculationState; if ($calculationState -eq 0) {{ Set-AuditStage 'calculation_complete'; break }}; if ($calculationState -eq 2) {{ if (-not $pendingFullRebuildRetried) {{ $pendingFullRebuildRetried=$true; Set-AuditStage 'retrying_full_rebuild_after_pending'; $app.CalculateFullRebuild(); Start-Sleep -Milliseconds 200; continue }}; throw 'Excel calculation remains xlPending after full dependency rebuild; refusing to save' }}; if ($calculationState -ne $lastCalculationState) {{ Set-AuditStage ('waiting_for_calculation_state:' + $calculationState); $lastCalculationState=$calculationState }}; if ([DateTime]::UtcNow -ge $deadline) {{ throw ('Excel calculation state did not reach xlDone before timeout; last state=' + $calculationState) }}; Start-Sleep -Milliseconds 200 }}; Set-AuditStage 'saving_output'; $book.SaveAs('{}',51); Set-AuditStage 'output_saved'; $book.Close($true); $book=$null; Set-AuditStage 'workbook_closed' }} catch {{ Set-AuditStage ('error: ' + $_.Exception.Message); throw }} finally {{ if ($null -ne $book) {{ $book.Close($false) }}; if ($null -ne $app) {{ $app.Quit() }} }}
Set-AuditStage 'complete'; Write-Output $appName
",
        stage_file.display(),
        probe,
        pid_file.display(),
        input.display(),
        timeout_seconds,
        output.display()
    )
}

fn read_com_stage(path: &Path) -> String {
    fs::read_to_string(path)
        .ok()
        .map(|stage| stage.trim().to_owned())
        .filter(|stage| !stage.is_empty())
        .unwrap_or_else(|| "no stage recorded".to_owned())
}

#[cfg(windows)]
fn terminate_owned_excel_process(pid_file: &Path) {
    let Ok(contents) = fs::read_to_string(pid_file) else {
        return;
    };
    let Ok(process_id) = contents.trim().parse::<u32>() else {
        return;
    };
    terminate_process_tree(process_id);
    let _ = fs::remove_file(pid_file);
}

#[cfg(not(windows))]
fn terminate_owned_excel_process(_pid_file: &Path) {}

fn validate_recalculated_workbook(input: &Path, output: &Path) -> Result<String> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .context("read source workbook before recalculation")?
        .workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(output)
        .context("read recalculated workbook")?
        .workbook;
    let formula_count = source
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    ensure!(
        formula_count > 0,
        "workbook contains no formulas to recalculate"
    );
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &result),
        "recalculated workbook changed worksheet structure, formula locations, or formula text beyond LibreOffice boolean-literal normalization"
    );
    let result_value_coverage = excel_reader::formula_result_value_coverage(&result);
    let result_value_count = result_value_coverage.typed_result_cell_count;
    ensure!(
        result_value_coverage.formula_cell_count == formula_count
            && result_value_count == formula_count
            && result_value_coverage.missing_result_cells.is_empty(),
        "recalculated workbook has formula/result value count mismatch: {formula_count} formulas, {result_value_count} typed cached results, missing={:?}",
        result_value_coverage.missing_result_cells.first()
    );
    let cache_count = formula_cache_cell_count(output)?;
    ensure!(cache_count == formula_count, "recalculated workbook has formula/cache count mismatch: {formula_count} formulas, {cache_count} cached formula cells");
    Ok(format!("{{\"formula_text_preserved\":true,\"formula_normalization\":\"Only TRUE/FALSE and TRUE()/FALSE() boolean literal spellings are treated as equivalent; matching literal error cells promoted by LibreOffice are restored before validation\",\"formula_count\":{formula_count},\"formula_result_values\":{result_value_count},\"formula_cache_cells\":{cache_count},\"promoted_error_constants_restored\":0,\"cache_nodes_complete\":true,\"result_values_readable\":true,\"scope\":\"fresh caches produced by an external calculation engine; result correctness requires business-rule or expected-value comparison\"}}"))
}

fn normalize_promoted_error_constants(
    input: &Path,
    candidate: &Path,
    attempt_dir: &Path,
) -> Result<(PathBuf, usize)> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .context("read source workbook for formula normalization")?
        .workbook;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(candidate)
        .context("read engine output for formula normalization")?
        .workbook;
    if excel_reader::formula_texts_equivalent_after_recalculation(&source, &result) {
        return Ok((candidate.to_path_buf(), 0));
    }

    let promoted = promoted_error_constants(&source, &result)?;
    ensure!(
        !promoted.is_empty(),
        "recalculated workbook changed worksheet structure, formula locations, or formula text beyond supported normalization"
    );
    let output = attempt_dir.join("recalculated-error-constants-restored.xlsx");
    let report =
        ooxml_patch::restore_promoted_error_constants(input, candidate, &output, &promoted)
            .context("restore LibreOffice-promoted literal error cells")?;
    let restored = excel_reader::read_workbook_with_unknown_error_tokens_as_text(&output)
        .context("read back normalized engine output")?
        .workbook;
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &restored),
        "recalculated workbook still changes formula structure after restoring literal error cells"
    );
    Ok((output, report.restored_cells.len()))
}

fn promoted_error_constants(
    source: &WorkbookDump,
    result: &WorkbookDump,
) -> Result<Vec<ooxml_patch::PromotedErrorConstant>> {
    ensure!(
        source.sheets.len() == result.sheets.len(),
        "recalculated workbook changed worksheet count"
    );
    let mut promoted = Vec::new();
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        ensure!(
            before.name == after.name,
            "recalculated workbook changed worksheet order or names"
        );
        for (cell, formula) in &before.formulas {
            let result_formula = after.formulas.get(cell).with_context(|| {
                format!(
                    "recalculated workbook removed formula at {}!{cell}",
                    before.name
                )
            })?;
            ensure!(
                excel_reader::canonicalize_formula_boolean_literals(formula)
                    == excel_reader::canonicalize_formula_boolean_literals(result_formula),
                "recalculated workbook changed formula at {}!{cell}",
                before.name
            );
        }
        for (cell, formula) in &after.formulas {
            if before.formulas.contains_key(cell) {
                continue;
            }
            ensure!(
                is_excel_error_literal(formula),
                "recalculated workbook added an unsupported formula at {}!{cell}",
                before.name
            );
            let source_value = before.values.get(cell).with_context(|| {
                format!(
                    "new formula at {}!{cell} has no matching source error value",
                    before.name
                )
            })?;
            let result_value = after.values.get(cell).with_context(|| {
                format!(
                    "new formula at {}!{cell} has no readable cached error value",
                    after.name
                )
            })?;
            ensure!(
                source_value.get("type").and_then(serde_json::Value::as_str) == Some("error")
                    && result_value.get("type").and_then(serde_json::Value::as_str)
                        == Some("error")
                    && source_value == result_value,
                "new formula at {}!{cell} does not preserve a source literal error value",
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

fn is_excel_error_literal(formula: &str) -> bool {
    matches!(
        formula,
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
    )
}

fn formula_cache_cell_count(path: &Path) -> Result<usize> {
    let inventory = excel_reader::inspect_formula_cache_nodes(path)?;
    if let Some(missing) = inventory.missing_cache_cells.first() {
        bail!(
            "formula cell {:?} on sheet {:?} has no cached value node after recalculation",
            missing.cell,
            missing.sheet_name
        );
    }
    Ok(inventory.cached_formula_cell_count)
}

fn run_libreoffice(input: &Path, temp_dir: &Path, target: &str) -> Result<()> {
    let profile = temp_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let mut command = Command::new("soffice");
    command
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg(if target == "xlsx" {
            "xlsx:Calc MS Excel 2007 XML"
        } else {
            "docx:Office Open XML Text"
        })
        .arg("--outdir")
        .arg(temp_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let output = command.output().with_context(|| {
        "LibreOffice sidecar is unavailable; install soffice to use this conversion"
    })?;
    if !output.status.success() {
        bail!(
            "LibreOffice conversion failed ({}): {}",
            output.status,
            String::from_utf8_lossy(&output.stderr)
        );
    }
    Ok(())
}

fn run_com(input: &Path, output: &Path, target: &str, source_ext: &str) -> Result<()> {
    #[cfg(windows)]
    {
        if target == "docx" {
            let session = excel_com::WordComSession::create(&[
                "Word.Application", "kwps.Application", "KWPS.Application", "wps.Application",
            ])?;
            session.convert_to_docx(input, output)
        } else {
            let names: &[&str] = if source_ext == "et" {
                &["ket.Application", "KET.Application", "Excel.Application"]
            } else {
                &["Excel.Application", "ket.Application", "KET.Application"]
            };
            let mut session = excel_com::ExcelComSession::create(names)?;
            session.convert_to_xlsx(input, output)
        }
    }
    #[cfg(not(windows))]
    {
        let _ = (input, output, target, source_ext);
        bail!("native Office COM conversion is only available on Windows")
    }
}

fn validate_conversion(
    input: &Path,
    output: &Path,
    source_ext: &str,
    target: &str,
    engine: &str,
    temp_dir: &Path,
) -> Result<String> {
    if target == "docx" {
        let bytes = fs::read(output)?;
        ensure!(
            bytes.starts_with(b"PK\x03\x04"),
            "converted DOCX is not a valid OOXML ZIP package"
        );
        let (source_text, output_text) = match engine {
            "lo" => (
                extract_word_text_lo(input, &temp_dir.join("source-text"))?,
                extract_word_text_lo(output, &temp_dir.join("output-text"))?,
            ),
            "com" if cfg!(windows) => (
                extract_word_text_com(input, &temp_dir.join("source.txt"))?,
                extract_word_text_com(output, &temp_dir.join("output.txt"))?,
            ),
            _ => bail!("unsupported text-validation engine {engine:?}"),
        };
        ensure!(
            normalize_document_text(&source_text) == normalize_document_text(&output_text),
            "converted DOCX text content differs from the source document"
        );
        return Ok("{\"level\":\"docx_package_and_text_content\"}".into());
    }
    let converted = read_comparable_workbook(output)
        .with_context(|| format!("validate converted workbook {}", output.display()))?;
    ensure!(
        !converted.sheets.is_empty(),
        "converted workbook has no worksheets"
    );
    if source_ext == "csv" {
        validate_csv_conversion(input, &converted)?;
        return Ok(
            "{\"level\":\"one_worksheet_and_csv_values\",\"numeric_tolerance\":1e-15}".into(),
        );
    }
    let source_path = if source_ext == "et" {
        // Calamine dispatches by extension. A temporary extension copy lets
        // it inspect ET files that use an XLSX-compatible OOXML package while
        // leaving the original file untouched. If the package isn't readable,
        // reject the conversion because sheet/data preservation cannot be
        // checked safely.
        let probe_dir = temp_dir.join("source-probe");
        fs::create_dir_all(&probe_dir)?;
        let probe = probe_dir.join("source.xlsx");
        fs::copy(input, &probe).context("prepare read-only ET validation copy")?;
        probe
    } else {
        input.to_path_buf()
    };
    let source = read_comparable_workbook(&source_path).with_context(|| {
        format!(
            "read source workbook for preservation validation {}",
            input.display()
        )
    })?;
    compare_workbooks(&source, &converted)?;
    Ok(
        "{\"level\":\"worksheet_order_names_values_and_formula_text\",\"numeric_tolerance\":1e-15}"
            .into(),
    )
}

fn extract_word_text_lo(input: &Path, output_dir: &Path) -> Result<String> {
    fs::create_dir_all(output_dir)?;
    let profile = output_dir.join("lo-profile");
    let profile_argument = libreoffice_user_installation_arg(&profile)?;
    let command = Command::new("soffice")
        .arg("--headless")
        .arg("--nologo")
        .arg("--nodefault")
        .arg("--nolockcheck")
        .arg(profile_argument)
        .arg("--convert-to")
        .arg("txt:Text")
        .arg("--outdir")
        .arg(output_dir)
        .arg(input)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .with_context(|| "LibreOffice is required to validate converted Word text")?;
    ensure!(
        command.status.success(),
        "LibreOffice text extraction failed: {}",
        String::from_utf8_lossy(&command.stderr)
    );
    let text_path = output_dir.join(format!("{}.txt", file_stem(input)));
    fs::read_to_string(&text_path)
        .with_context(|| format!("read extracted Word text {}", text_path.display()))
}

fn extract_word_text_com(input: &Path, _text_output: &Path) -> Result<String> {
    #[cfg(windows)]
    {
        let session = excel_com::WordComSession::create(&[
            "Word.Application", "kwps.Application", "KWPS.Application", "wps.Application",
        ])?;
        session.read_text(input)
    }
    #[cfg(not(windows))]
    {
        let _ = input;
        bail!("native Word COM text validation is only available on Windows")
    }
}

fn normalize_document_text(text: &str) -> String {
    text.replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim_end_matches(['\n', '\u{000c}'])
        .to_owned()
}

fn read_comparable_workbook(path: &Path) -> Result<WorkbookDump> {
    let mut workbook = read_workbook(path)?;
    let has_date_serial = workbook.sheets.iter().any(|sheet| {
        sheet.values.values().any(|value| {
            value.get("type").and_then(serde_json::Value::as_str) == Some("datetime_serial")
        })
    });
    if !has_date_serial {
        return Ok(workbook);
    }
    let max_columns = workbook
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet
                .values
                .keys()
                .filter_map(|address| column_number(address))
        })
        .max()
        .unwrap_or(1)
        .max(1);
    for sheet in read_value_columns_with_iso_dates(path, max_columns)? {
        let Some(target) = workbook
            .sheets
            .iter_mut()
            .find(|candidate| candidate.name == sheet.name)
        else {
            continue;
        };
        for (address, value) in sheet.values {
            if let ReaderCellValue::DateTimeIso(value) = value {
                target.values.insert(
                    address,
                    serde_json::json!({"type":"datetime_iso","value":value}),
                );
            }
        }
    }
    Ok(workbook)
}

fn compare_workbooks(source: &WorkbookDump, converted: &WorkbookDump) -> Result<()> {
    let source_names = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let converted_names = converted
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    ensure!(
        source_names == converted_names,
        "conversion changed worksheet count, names or order"
    );
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&converted.sheets) {
        ensure!(
            source_sheet.formulas.len() == output_sheet.formulas.len(),
            "conversion changed formula count in worksheet {:?}",
            source_sheet.name
        );
        for (address, formula) in &source_sheet.formulas {
            let Some(output_formula) = output_sheet.formulas.get(address) else {
                bail!(
                    "conversion removed formula {}!{}",
                    source_sheet.name,
                    address
                );
            };
            ensure!(
                formula.trim_start_matches('=').trim()
                    == output_formula.trim_start_matches('=').trim(),
                "conversion changed formula text at {}!{}: {:?} -> {:?}",
                source_sheet.name,
                address,
                formula,
                output_formula
            );
        }
        let source_cells = source_sheet
            .values
            .keys()
            .filter(|address| !source_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        let output_cells = output_sheet
            .values
            .keys()
            .filter(|address| !output_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        ensure!(
            source_cells == output_cells,
            "conversion changed non-formula data cell coordinates in worksheet {:?}",
            source_sheet.name
        );
        for address in source_cells {
            let source_value = &source_sheet.values[address];
            let output_value = &output_sheet.values[address];
            ensure!(
                cell_values_equal(source_value, output_value),
                "conversion changed stored value at {}!{}: {} -> {}",
                source_sheet.name,
                address,
                source_value,
                output_value
            );
        }
    }
    Ok(())
}

fn cell_values_equal(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    let numeric = |kind: &str| matches!(kind, "number" | "integer" | "datetime_serial");
    let left_kind = cell_kind(left);
    let right_kind = cell_kind(right);
    if numeric(left_kind) && numeric(right_kind) {
        if left_kind == "integer" && right_kind == "integer" {
            return left.get("value").and_then(serde_json::Value::as_i64)
                == right.get("value").and_then(serde_json::Value::as_i64);
        }
        let Some(left) = left.get("value").and_then(serde_json::Value::as_f64) else {
            return left == right;
        };
        let Some(right) = right.get("value").and_then(serde_json::Value::as_f64) else {
            return false;
        };
        return left == right
            || (left.is_finite()
                && right.is_finite()
                && left != 0.0
                && right != 0.0
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()));
    }
    left == right
}

fn cell_kind(value: &serde_json::Value) -> &str {
    value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default()
}

fn column_number(address: &str) -> Option<usize> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    address[..split].bytes().try_fold(0usize, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase().checked_sub(b'A')? + 1) as usize)
    })
}

fn locate_converted(temp: &Path, target: &str) -> Result<PathBuf> {
    let mut entries = fs::read_dir(temp)?
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .filter(|p| {
            p.extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| e.eq_ignore_ascii_case(target))
        })
        .collect::<Vec<_>>();
    entries.sort();
    entries
        .into_iter()
        .next()
        .ok_or_else(|| anyhow!("conversion produced no .{target} output"))
}
fn unique_temp_dir(parent: &Path) -> Result<PathBuf> {
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for i in 0..100 {
        let path = parent.join(format!(".lo-convert-{}-{nanos}-{i}", std::process::id()));
        if !path.exists() {
            fs::create_dir(&path)?;
            return Ok(path);
        }
    }
    bail!("cannot allocate conversion work directory")
}
fn extension(path: &Path) -> Result<String> {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
        .ok_or_else(|| anyhow!("file has no extension: {}", path.display()))
}
fn file_stem(path: &Path) -> String {
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "converted".into())
}
fn same_category(source: &str, target: &str) -> bool {
    let excel = ["csv", "xls", "xlsx", "xlsm", "xlsb", "et"];
    let word = ["doc", "docx"];
    (excel.contains(&source) && target == "xlsx") || (word.contains(&source) && target == "docx")
}

fn validate_csv_conversion(input: &Path, converted: &WorkbookDump) -> Result<()> {
    let bytes = fs::read(input).with_context(|| format!("read CSV {}", input.display()))?;
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
    let decoded = match std::str::from_utf8(bytes) {
        Ok(text) => text.to_owned(),
        Err(_) => {
            let (text, _, had_errors) = GBK.decode(bytes);
            ensure!(
                !had_errors,
                "CSV is neither valid UTF-8 nor valid GBK; cannot validate converted values"
            );
            text.into_owned()
        }
    };
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(decoded.as_bytes());
    let rows = reader
        .records()
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse CSV for conversion validation")?;
    ensure!(
        converted.sheets.len() == 1,
        "CSV conversion must produce exactly one worksheet"
    );
    let sheet = &converted.sheets[0];
    let mut expected_cells = std::collections::BTreeSet::new();
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, raw) in row.iter().enumerate() {
            if raw.is_empty() {
                continue;
            }
            let address = format!("{}{}", csv_column_name(column_index + 1), row_index + 1);
            expected_cells.insert(address.clone());
            if let Some(formula) = sheet.formulas.get(&address) {
                ensure!(
                    raw.starts_with('=')
                        && raw.trim_start_matches('=').trim()
                            == formula.trim_start_matches('=').trim(),
                    "CSV conversion changed formula text at {}!{address}",
                    sheet.name
                );
                continue;
            }
            let value = sheet
                .values
                .get(&address)
                .with_context(|| format!("CSV conversion lost data at {}!{address}", sheet.name))?;
            ensure!(
                csv_text_matches_probe(raw, value),
                "CSV conversion changed data at {}!{address}: source={raw:?}, output={value}",
                sheet.name
            );
        }
    }
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        let non_empty = sheet
            .values
            .get(address)
            .and_then(|value| value.get("value"))
            .is_some_and(|value| {
                !value.is_null() && value.as_str().is_none_or(|text| !text.is_empty())
            });
        if non_empty || sheet.formulas.contains_key(address) {
            ensure!(
                expected_cells.contains(address),
                "CSV conversion introduced data at {}!{address}",
                sheet.name
            );
        }
    }
    Ok(())
}

fn csv_text_matches_probe(source: &str, actual: &serde_json::Value) -> bool {
    let Some(kind) = actual.get("type").and_then(serde_json::Value::as_str) else {
        return false;
    };
    let Some(value) = actual.get("value") else {
        return false;
    };
    match kind {
        "string" | "error" => value.as_str() == Some(source),
        "number" | "integer" | "datetime_serial" => {
            let Some(left) = source.trim().parse::<f64>().ok() else {
                return false;
            };
            let Some(right) = value.as_f64() else {
                return false;
            };
            left.is_finite()
                && right.is_finite()
                && (left - right).abs() <= 1e-15 * left.abs().max(right.abs()).max(1.0)
        }
        "boolean" => {
            let source = source.trim();
            match value.as_bool() {
                Some(true) => source.eq_ignore_ascii_case("true") || source == "1",
                Some(false) => source.eq_ignore_ascii_case("false") || source == "0",
                None => false,
            }
        }
        "datetime_iso" => {
            let source = source.trim().replace('/', "-").replace(' ', "T");
            let actual = value.as_str().unwrap_or_default();
            source == actual
                || (source.len() == 10 && actual.starts_with(&source))
                || chrono::NaiveDate::parse_from_str(&source, "%Y-%m-%d")
                    .ok()
                    .is_some_and(|date| actual.starts_with(&date.format("%Y-%m-%d").to_string()))
        }
        _ => false,
    }
}

fn csv_column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}
fn ps_quote(value: &str) -> String {
    value.replace('\'', "''")
}
fn json_string(value: &str) -> String {
    serde_json::to_string(value).unwrap_or_else(|_| "\"\"".into())
}
fn print_usage() {
    eprintln!("Usage: lo-convert convert --input <file> --output <file> --to <xlsx|docx> [--mode fidelity] [--engine auto|lo|com]\n       lo-convert recalculate --input <xlsx> --output <new.xlsx> [--engine auto|lo|com]");
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod recalculation_tests {
    use super::{
        com_recalculation_app_names, com_recalculate_script, file_uri, formula_cache_cell_count,
        libreoffice_user_installation_arg, prepare_com_recalculation_input,
        promoted_error_constants, publish_recalculated_output, recalculation_engine_order,
        require_engine_output, run_command_with_timeout, validate_recalculated_workbook, MAIN_NS,
        PACKAGE_REL_NS, REL_NS,
    };
    use anyhow::Result;
    use excel_reader::{SheetDump, WorkbookDump};
    use std::{
        collections::BTreeMap,
        env,
        fs::{self, File},
        io::Write,
        path::{Path, PathBuf},
        process::Command,
        time::{Duration, SystemTime, UNIX_EPOCH},
    };
    use zip::{write::SimpleFileOptions, ZipWriter};

    fn fixture(with_cache: bool) -> Result<PathBuf> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let path = env::temp_dir().join(format!(
            "lo_recalc_fixture_{}_{}.xlsx",
            std::process::id(),
            nonce
        ));
        let zip = ZipWriter::new(File::create(&path)?);
        let workbook = format!(
            r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Formula" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let cache = if with_cache { "<v>2</v>" } else { "" };
        let sheet = format!(
            r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="1"><c r="A1"><f>1+1</f>{cache}</c></row></sheetData></worksheet>"#
        );
        let mut writer = zip;
        for (name, content) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/worksheets/sheet1.xml", sheet),
        ] {
            writer.start_file(name, SimpleFileOptions::default())?;
            writer.write_all(content.as_bytes())?;
        }
        writer.finish()?;
        Ok(path)
    }

    fn error_promotion_workbook(
        formulas: BTreeMap<String, String>,
        error_value: &str,
        include_error_formula: bool,
    ) -> WorkbookDump {
        let values = BTreeMap::from([(
            "C211".to_owned(),
            serde_json::json!({"type":"error","value":error_value}),
        )]);
        let mut formulas = formulas;
        if include_error_formula {
            formulas.insert("C211".to_owned(), "#REF!".to_owned());
        }
        WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "参照表".to_owned(),
                values,
                formulas,
                formula_status: "ok".to_owned(),
                merges: Vec::new(),
                merge_status: "ok".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "ok".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "ok".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "ok".to_owned(),
            }],
        }
    }

    #[test]
    fn promotes_only_an_unchanged_literal_error_cell_for_ooxml_restoration() -> Result<()> {
        let source = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "Ref",
            false,
        );
        let result = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE())".to_owned())]),
            "Ref",
            true,
        );

        let promotions = promoted_error_constants(&source, &result)?;
        assert_eq!(promotions.len(), 1);
        assert_eq!(promotions[0].sheet_name, "参照表");
        assert_eq!(promotions[0].cell_reference, "C211");
        assert_eq!(promotions[0].error_token, "#REF!");
        Ok(())
    }

    #[test]
    fn rejects_changed_formula_or_error_value_during_normalization() {
        let source = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "Ref",
            false,
        );
        let changed_formula = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Other!A:B,2,FALSE)".to_owned())]),
            "Ref",
            true,
        );
        assert!(promoted_error_constants(&source, &changed_formula)
            .expect_err("changed formula reference must be rejected")
            .to_string()
            .contains("changed formula"));

        let changed_error = error_promotion_workbook(
            BTreeMap::from([("D2".to_owned(), "VLOOKUP(A2,Ref!A:B,2,FALSE)".to_owned())]),
            "NA",
            true,
        );
        assert!(promoted_error_constants(&source, &changed_error)
            .expect_err("different typed error must be rejected")
            .to_string()
            .contains("does not preserve a source literal error value"));
    }

    #[test]
    fn empty_string_formula_cache_remains_a_readable_result() -> Result<()> {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_formula(0, 0, "=IF(TRUE,\"\",\"x\")")?;
        sheet.set_formula_result(0, 0, "");
        let path = fixture(false)?;
        workbook.save(&path)?;
        let read = excel_reader::read_workbook(&path)?;
        assert_eq!(
            read.sheets[0].values.get("A1"),
            Some(&serde_json::json!({"type": "string", "value": ""}))
        );
        fs::remove_file(path)?;
        Ok(())
    }

    #[test]
    fn recalc_validation_rejects_changed_formula_text() -> Result<()> {
        let root = env::temp_dir().join(format!(
            "lo_recalc_formula_change_{}_{}",
            std::process::id(),
            SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
        ));
        fs::create_dir_all(&root)?;
        let input = root.join("input.xlsx");
        let output = root.join("output.xlsx");

        let mut source = rust_xlsxwriter::Workbook::new();
        let source_sheet = source.add_worksheet();
        source_sheet.write_number(0, 0, 2)?;
        source_sheet.write_formula(0, 1, "=A1+1")?;
        source_sheet.set_formula_result(0, 1, "3");
        source.save(&input)?;

        let mut changed = rust_xlsxwriter::Workbook::new();
        let changed_sheet = changed.add_worksheet();
        changed_sheet.write_number(0, 0, 2)?;
        changed_sheet.write_formula(0, 1, "=A1+2")?;
        changed_sheet.set_formula_result(0, 1, "4");
        changed.save(&output)?;

        let error = validate_recalculated_workbook(&input, &output)
            .expect_err("formula text changes must fail recalculation validation");
        assert!(error.to_string().contains("changed worksheet structure"));
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn com_input_clears_load_flags_and_preserves_uncached_formulas() -> Result<()> {
        let input = fixture(false)?;
        let output = input.with_file_name(format!(
            "lo_recalc_com_input_{}_{}.xlsx",
            std::process::id(),
            SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
        ));
        prepare_com_recalculation_input(&input, &output)?;
        let properties = ooxml_patch::read_workbook_calculation_properties(&output)?;
        assert_eq!(properties.full_calc_on_load, Some(false));
        assert_eq!(properties.force_full_calc, Some(false));
        let inventory = excel_reader::inspect_formula_cache_nodes(&output)?;
        assert_eq!(inventory.formula_cell_count, 1);
        assert_eq!(inventory.cached_formula_cell_count, 0);
        assert_eq!(inventory.missing_cache_cells[0].cell, "A1");
        fs::remove_file(input)?;
        fs::remove_file(output)?;
        Ok(())
    }

    #[test]
    fn formula_cache_counter_requires_a_value_node_for_every_formula() -> Result<()> {
        let valid = fixture(true)?;
        assert_eq!(formula_cache_cell_count(&valid)?, 1);
        fs::remove_file(valid)?;

        let invalid = fixture(false)?;
        let error =
            formula_cache_cell_count(&invalid).expect_err("formula without cache must fail");
        assert!(error.to_string().contains("has no cached value node"));
        fs::remove_file(invalid)?;
        Ok(())
    }

    #[test]
    fn excel_forced_recalculation_uses_microsoft_excel_only() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("excel"),
        );
        assert!(script.contains("New-Object -ComObject 'Excel.Application'"));
        assert!(script.contains("Microsoft Excel COM is unavailable"));
        assert!(script.contains("CalculateFull()"));
        assert!(script.contains("CalculateUntilAsyncQueriesDone"));
        assert!(script.contains("$book.Connections.Count -gt 0"));
        assert!(script.contains("CalculationState"));
        assert!(script.contains("Set-AuditStage 'opening_workbook'"));
        assert!(script.contains("$app.Calculation=-4105"));
        assert!(script.contains("Set-AuditStage 'enabling_automatic_calculation'"));
        assert!(script.contains("Set-AuditStage 'calculating_full'"));
        assert!(script.contains("waiting_for_calculation_state:' + $calculationState"));
        assert!(script.contains("$pendingFullRebuildRetried=$false"));
        assert!(script.contains(
            "Set-AuditStage 'retrying_full_rebuild_after_pending'; $app.CalculateFullRebuild()"
        ));
        assert!(script.contains(
            "Excel calculation remains xlPending after full dependency rebuild; refusing to save"
        ));
        assert!(!script.contains("calculation_pending_after_full_calc'; break"));
        assert!(script.contains("Set-AuditStage 'saving_output'"));
        assert!(script.contains("GetWindowThreadProcessId"));
        assert!(!script.contains("KET.Application"));
        assert!(!script.contains("ket.Application"));
        assert!(!script.contains("wps.Application"));
    }

    #[test]
    fn auto_com_channel_probes_excel_then_wps() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("com"),
        );
        let excel_probe = script
            .find("New-Object -ComObject 'Excel.Application'")
            .expect("auto channel must probe Excel first");
        let wps_probe = script
            .find("New-Object -ComObject 'ket.Application'")
            .expect("auto channel must fall back to WPS Spreadsheets");
        let ket_probe = script
            .find("New-Object -ComObject 'KET.Application'")
            .expect("auto channel must cover the capitalised WPS ProgID");
        assert!(excel_probe < wps_probe && wps_probe < ket_probe);
        assert!(script.contains(
            "No supported Office/WPS COM engine is available: $lastError"
        ));
        assert!(!script.contains("Microsoft Excel COM is unavailable"));
    }

    #[test]
    fn wps_forced_recalculation_uses_wps_only() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("wps"),
        );
        assert!(script.contains("New-Object -ComObject 'ket.Application'"));
        assert!(script.contains("New-Object -ComObject 'KET.Application'"));
        assert!(script.contains("WPS Spreadsheets COM is unavailable"));
        assert!(!script.contains("New-Object -ComObject 'Excel.Application'"));
        assert!(script.contains("CalculateFull()"));
        assert!(script.contains("Set-AuditStage 'saving_output'"));
    }

    #[test]
    fn recalculation_engine_order_covers_forced_channels() {
        assert_eq!(recalculation_engine_order("excel", true).unwrap(), vec!["excel"]);
        assert_eq!(recalculation_engine_order("wps", true).unwrap(), vec!["wps"]);
        assert_eq!(recalculation_engine_order("com", true).unwrap(), vec!["com"]);
        assert_eq!(
            recalculation_engine_order("auto", true).unwrap(),
            vec!["com", "lo"]
        );
        assert_eq!(recalculation_engine_order("auto", false).unwrap(), vec!["lo"]);
        for engine in ["excel", "wps", "com"] {
            let error = recalculation_engine_order(engine, false)
                .expect_err("COM-family engines are Windows-only");
            assert!(error.to_string().contains("only available on Windows"));
        }
        assert!(recalculation_engine_order("bogus", true).is_err());
    }

    #[test]
    fn windows_com_does_not_save_when_calculation_remains_pending() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("excel"),
        );
        let pending_error = script
            .find("Excel calculation remains xlPending after full dependency rebuild; refusing to save")
            .expect("persistent pending state must be an error");
        let save = script
            .find("Set-AuditStage 'saving_output'")
            .expect("script includes output save stage");
        assert!(
            pending_error < save,
            "pending-state guard must run before saving"
        );
        assert!(script.contains("if (-not $pendingFullRebuildRetried)"));
        assert!(
            script.contains("$app.CalculateFullRebuild(); Start-Sleep -Milliseconds 200; continue")
        );
    }

    #[test]
    fn windows_com_does_not_save_when_async_query_wait_fails() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("excel"),
        );
        let async_wait_failure = script
            .find("Set-AuditStage 'async_query_wait_failed'; throw ('Excel async query wait failed: ' + $_.Exception.Message)")
            .expect("async query wait errors must fail closed");
        let save = script
            .find("Set-AuditStage 'saving_output'")
            .expect("script includes output save stage");
        assert!(
            async_wait_failure < save,
            "an async query wait error must stop the script before saving"
        );
    }

    #[test]
    fn windows_com_recalculation_clears_book_after_successful_close() {
        let script = com_recalculate_script(
            Path::new(r"C:\input.xlsx"),
            Path::new(r"C:\output.xlsx"),
            Path::new(r"C:\owned-excel-pid.txt"),
            Path::new(r"C:\excel-com-progress.txt"),
            300,
            com_recalculation_app_names("excel"),
        );
        assert!(script.contains("$book.Close($true); $book=$null"));
        assert!(script.contains("if ($null -ne $book) { $book.Close($false) }"));
    }

    #[test]
    fn missing_engine_output_is_rejected_even_when_request_copy_exists() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root =
            env::temp_dir().join(format!("lo_recalc_output_{}_{}", std::process::id(), nonce));
        let input_dir = root.join("input");
        let output_dir = root.join("engine-output");
        fs::create_dir_all(&input_dir)?;
        fs::create_dir_all(&output_dir)?;
        let request_copy = input_dir.join("recalculation-request.xlsx");
        fs::write(&request_copy, b"input copy")?;
        let expected_output = output_dir.join("recalculation-request.xlsx");

        let error = require_engine_output(&expected_output)
            .expect_err("missing engine output must not select the input copy");
        assert!(error.to_string().contains("did not create"));
        assert!(request_copy.is_file());
        assert!(!expected_output.exists());

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn recalculated_output_publication_preserves_existing_destination() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = env::temp_dir().join(format!(
            "lo_recalc_publish_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root)?;
        let candidate = root.join("candidate.xlsx");
        let output = root.join("output.xlsx");
        fs::write(&candidate, b"validated workbook bytes")?;
        fs::write(&output, b"pre-existing destination")?;

        let error = publish_recalculated_output(&candidate, &output)
            .expect_err("publication must not overwrite a pre-existing output");
        assert!(format!("{error:#}").contains("without overwriting"));
        assert_eq!(fs::read(&output)?, b"pre-existing destination");

        fs::remove_file(&output)?;
        publish_recalculated_output(&candidate, &output)?;
        assert_eq!(fs::read(&output)?, b"validated workbook bytes");
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn libreoffice_profile_path_is_encoded_as_a_file_uri() -> Result<()> {
        let root = env::temp_dir().join("lo profile # 中文");
        let uri = file_uri(&root)?;
        assert!(uri.starts_with("file://"));
        assert!(uri.contains("%20"));
        assert!(uri.contains("%23"));
        assert!(!uri.contains(' '));
        assert!(!uri.contains('中'));
        Ok(())
    }

    #[test]
    fn libreoffice_user_installation_argument_creates_profile_and_uses_file_uri() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let profile = env::temp_dir().join(format!(
            "lo_profile_arg_{}_{}_# space",
            std::process::id(),
            nonce
        ));
        assert!(!profile.exists());

        let argument = libreoffice_user_installation_arg(&profile)?;

        assert!(profile.is_dir());
        assert!(argument.starts_with("-env:UserInstallation=file://"));
        assert!(argument.contains("%20"));
        assert!(argument.contains("%23"));
        fs::remove_dir_all(profile)?;
        Ok(())
    }

    #[test]
    fn automatic_calculation_uses_excel_then_libreoffice_on_windows() -> Result<()> {
        assert_eq!(recalculation_engine_order("auto", true)?, ["com", "lo"]);
        assert_eq!(recalculation_engine_order("auto", false)?, ["lo"]);
        assert_eq!(recalculation_engine_order("lo", true)?, ["lo"]);
        assert!(recalculation_engine_order("com", false).is_err());
        Ok(())
    }

    #[test]
    fn engine_process_timeout_kills_the_child_and_reports_timeout() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = env::temp_dir().join(format!(
            "lo_recalc_timeout_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root)?;
        let mut command = if cfg!(windows) {
            let mut command = Command::new("powershell.exe");
            command.args([
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "Start-Sleep -Seconds 10",
            ]);
            command
        } else {
            let mut command = Command::new("sh");
            command.args(["-c", "sleep 10"]);
            command
        };
        let error = run_command_with_timeout(
            &mut command,
            &root,
            Duration::from_millis(50),
            "synthetic calculation",
        )
        .expect_err("a sleeping engine should time out");
        assert!(format!("{error:#}").contains("timeout"));
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    #[cfg(any(windows, target_os = "linux"))]
    fn engine_process_timeout_terminates_descendants_too() -> Result<()> {
        let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = env::temp_dir().join(format!(
            "lo_recalc_timeout_tree_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root)?;
        let child_pid_file = root.join("child-pid.txt");
        let mut command = if cfg!(windows) {
            let mut command = Command::new("powershell.exe");
            command.args([
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "$child = Start-Process -FilePath $env:ComSpec -ArgumentList '/c ping -n 30 127.0.0.1 > nul' -PassThru -WindowStyle Hidden; [System.IO.File]::WriteAllText($env:AUDIT_TEST_CHILD_PID_FILE, [string]$child.Id); Start-Sleep -Seconds 30",
            ]);
            command.env("AUDIT_TEST_CHILD_PID_FILE", &child_pid_file);
            command
        } else {
            let mut command = Command::new("sh");
            command.args([
                "-c",
                "sleep 30 & child=$!; printf '%s' \"$child\" > \"$AUDIT_TEST_CHILD_PID_FILE\"; wait",
            ]);
            command.env("AUDIT_TEST_CHILD_PID_FILE", &child_pid_file);
            command
        };

        let error = run_command_with_timeout(
            &mut command,
            &root,
            Duration::from_secs(1),
            "synthetic calculation tree",
        )
        .expect_err("a sleeping engine tree should time out");
        assert!(format!("{error:#}").contains("timeout"));

        let child_pid = fs::read_to_string(&child_pid_file)?.trim().to_owned();
        assert!(
            !child_pid.is_empty() && child_pid.bytes().all(|byte| byte.is_ascii_digit()),
            "the test engine should record its descendant PID"
        );
        for _ in 0..30 {
            if !process_is_running(&child_pid) {
                fs::remove_dir_all(root)?;
                return Ok(());
            }
            std::thread::sleep(Duration::from_millis(100));
        }
        fs::remove_dir_all(&root)?;
        anyhow::bail!("timeout left descendant process {child_pid} running")
    }

    #[cfg(windows)]
    fn process_is_running(process_id: &str) -> bool {
        let filter = format!("PID eq {process_id}");
        Command::new("tasklist.exe")
            .args(["/FI", filter.as_str(), "/NH"])
            .output()
            .ok()
            .is_some_and(|output| {
                String::from_utf8_lossy(&output.stdout)
                    .lines()
                    .any(|line| line.split_whitespace().nth(1) == Some(process_id))
            })
    }

    #[cfg(target_os = "linux")]
    fn process_is_running(process_id: &str) -> bool {
        let Ok(stat) = fs::read_to_string(format!("/proc/{process_id}/stat")) else {
            return false;
        };
        // A killed orphan can briefly remain as a zombie until init reaps it;
        // it is no longer an executing engine process.
        !stat
            .rsplit_once(") ")
            .and_then(|(_, fields)| fields.split_whitespace().next())
            .is_some_and(|state| state == "Z")
    }
}

/// 批次格式转换：`convert-batch --pairs <tsv> --to <xlsx|docx> --engine <auto|com|lo>`。
/// 整个批次只创建一个引擎实例并顺序处理（rules/13 §4.5）：
/// - Windows COM：一个 Word/Excel/WPS 会话贯穿整批，docx 文本校验也通过
///   同一会话读取（移植自 Rust-Tools e2973e6 的 ConvertWord/ExtractWordText
///   请求语义），PowerShell 桥不参与；
/// - LibreOffice：一条 `--convert-to` 处理整批，docx 文本校验按批各提取
///   一次（源文本一次、候选文本一次），而非每本两次冷启动。
/// 每本仍执行既有验证：DOCX 包合法性 + 文本一致；工作簿按既有比较器。
fn run_convert_batch(arguments: Vec<String>) -> Result<()> {
    let mut values: BTreeMap<String, String> = BTreeMap::new();
    let mut args = arguments.into_iter();
    while let Some(key) = args.next() {
        ensure!(key.starts_with("--"), "expected an option, got {key:?}");
        let name = key.trim_start_matches('-').to_owned();
        ensure!(
            matches!(name.as_str(), "pairs" | "to" | "engine" | "mode"),
            "unknown convert-batch option --{name}"
        );
        let value = args
            .next()
            .with_context(|| format!("{key} requires a value"))?;
        values.insert(name, value);
    }
    let target = values.get("to").map(String::as_str).unwrap_or("");
    ensure!(target == "xlsx" || target == "docx", "--to must be xlsx or docx");
    let engine = values.get("engine").map(String::as_str).unwrap_or("auto");
    ensure!(
        engine == "auto" || engine == "com" || engine == "lo",
        "--engine must be auto, com or lo"
    );
    let mode = values.get("mode").map(String::as_str).unwrap_or("fidelity");
    ensure!(mode == "fidelity", "--mode must be fidelity");

    let pairs_path = PathBuf::from(values.get("pairs").context("--pairs is required")?);
    let content = fs::read_to_string(&pairs_path)
        .with_context(|| format!("read convert pairs {}", pairs_path.display()))?;
    let mut declared = Vec::new();
    for line in content.lines() {
        if line.trim().is_empty() {
            continue;
        }
        let mut parts = line.split('\t');
        let input = parts.next().context("convert pair missing input")?;
        let output = parts.next().context("convert pair missing output")?;
        ensure!(
            parts.next().is_none(),
            "convert pair line has extra fields: {line:?}"
        );
        declared.push((PathBuf::from(input), PathBuf::from(output)));
    }
    ensure!(!declared.is_empty(), "convert pairs file is empty");
    let source_ext = declared[0]
        .0
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    for (input, output) in &declared {
        ensure!(
            input.is_file(),
            "convert input does not exist: {}",
            input.display()
        );
        ensure!(
            !output.exists(),
            "refusing to overwrite existing convert output: {}",
            output.display()
        );
        let input_ext = input
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or("")
            .to_ascii_lowercase();
        ensure!(
            input_ext == source_ext,
            "batch inputs must share one source extension: {source_ext} vs {input_ext}"
        );
    }

    let guard = TempDirGuard(unique_temp_dir(&env::temp_dir())?);
    let use_com = cfg!(windows) && engine != "lo";

    #[cfg(windows)]
    let mut com_word: Option<excel_com::WordComSession> = None;
    #[cfg(windows)]
    let mut com_sheet: Option<excel_com::ExcelComSession> = None;
    #[cfg(not(windows))]
    {
        let _ = use_com;
    }

    let mut books = Vec::with_capacity(declared.len());
    let mut failures = 0usize;

    // ---- Windows COM：批次自有实例，逐本转换 ----
    if use_com {
        for (index, (input, _output)) in declared.iter().enumerate() {
            let canonical = fs::canonicalize(input)
                .with_context(|| format!("resolve convert input {}", input.display()))?;
            let staged = guard.0.join(format!("book-{index:04}.{}", source_ext));
            fs::copy(&canonical, &staged).with_context(|| {
                format!(
                    "stage convert input {} -> {}",
                    canonical.display(),
                    staged.display()
                )
            })?;
            let attempt = (|| -> Result<serde_json::Value> {
                #[cfg(windows)]
                {
                    let (engine_name, application, process_id) = if target == "docx" {
                        if com_word.is_none() {
                            com_word = Some(excel_com::WordComSession::create(&[
                                "Word.Application",
                                "kwps.Application",
                                "KWPS.Application",
                                "wps.Application",
                            ])?);
                        }
                        let session = com_word.as_ref().expect("Word session ready");
                        session.convert_to_docx(&staged, _output)?;
                        (
                            "com",
                            "Word/WPS Writer",
                            Some(session.process_id()),
                        )
                    } else {
                        if com_sheet.is_none() {
                            let names: &[&str] = if source_ext == "et" {
                                &["ket.Application", "KET.Application", "Excel.Application"]
                            } else {
                                &["Excel.Application", "ket.Application", "KET.Application"]
                            };
                            com_sheet = Some(excel_com::ExcelComSession::create(names)?);
                        }
                        let session = com_sheet.as_mut().expect("Excel session ready");
                        session.convert_to_xlsx(&staged, _output)?;
                        let application = session.application_name();
                        ( "com", application, session.process_id())
                    };
                    require_engine_output(_output)?;
                    let validation = validate_conversion(
                        input,
                        _output,
                        &source_ext,
                        target,
                        "com",
                        &guard.0,
                    )?;
                    Ok(serde_json::json!({
                        "index": index + 1,
                        "status": "converted",
                        "engine": engine_name,
                        "application": application,
                        "process_id": process_id,
                        "validation": serde_json::from_str::<serde_json::Value>(&validation)
                            .unwrap_or_else(|_| serde_json::json!({"detail": validation})),
                    }))
                }
                #[cfg(not(windows))]
                {
                    bail!("native Office COM conversion is only available on Windows")
                }
            })();
            match attempt {
                Ok(book) => books.push(book),
                Err(error) => {
                    failures += 1;
                    books.push(serde_json::json!({
                        "index": index + 1,
                        "status": "failed",
                        "input": input.display().to_string(),
                        "error": format!("{error:#}"),
                    }));
                }
            }
        }
        let result = serde_json::json!({
            "status": if failures == 0 { "converted" } else { "partial" },
            "engine": "com",
            "books": books,
        });
        println!("{}", serde_json::to_string(&result)?);
        return Ok(());
    }

    // ---- LibreOffice：一条 --convert-to 处理整批 + 按批文本提取 ----
    let mut staged = Vec::with_capacity(declared.len());
    for (index, (input, _)) in declared.iter().enumerate() {
        let canonical = fs::canonicalize(input)
            .with_context(|| format!("resolve convert input {}", input.display()))?;
        let staged_path = guard.0.join(format!("book-{index:04}.{}", source_ext));
        fs::copy(&canonical, &staged_path).with_context(|| {
            format!(
                "stage convert input {} -> {}",
                canonical.display(),
                staged_path.display()
            )
        })?;
        staged.push(staged_path);
    }
    let target_kind = if target == "docx" {
        lo_batch::BatchTarget::Docx
    } else {
        lo_batch::BatchTarget::Xlsx
    };
    let convert_started = Instant::now();
    let batch_timeout = recalculation_timeout()
        .saturating_mul(declared.len() as u32)
        .max(recalculation_timeout());
    let pairs_for_convert: Vec<(PathBuf, PathBuf)> = staged
        .iter()
        .enumerate()
        .map(|(index, staged_path)| (staged_path.clone(), declared[index].1.clone()))
        .collect();
    let (engine_name, batch_stages) =
        lo_batch::run_batch_convert(&pairs_for_convert, guard.0.as_path(), target_kind, batch_timeout)
            .context("batch LibreOffice conversion")?;
    let convert_ms = convert_started.elapsed().as_millis() as u64;
    // 逐本把批次输出发布到期望路径（硬链接、拒绝覆盖，同既有语义）。
    for (index, (_, output)) in declared.iter().enumerate() {
        let candidate = batch_stages
            .get(index)
            .and_then(|stage| stage.get("batch_output"))
            .and_then(serde_json::Value::as_str)
            .map(PathBuf::from)
            .with_context(|| format!("batch did not produce output for book {}", index + 1))?;
        publish_recalculated_output(&candidate, output)
            .with_context(|| format!("publish converted book {}", index + 1))?;
    }
    let docx_texts = if target == "docx" {
        let sources = lo_batch::run_batch_text_extract(
            &staged,
            guard.0.as_path(),
            "sources",
            batch_timeout,
        )
        .context("batch Word source text extraction")?;
        let candidates: Vec<PathBuf> = declared.iter().map(|(_, output)| output.clone()).collect();
        let candidate_texts = lo_batch::run_batch_text_extract(
            &candidates,
            guard.0.as_path(),
            "candidates",
            batch_timeout,
        )
        .context("batch Word candidate text extraction")?;
        Some(
            sources
                .into_iter()
                .zip(candidate_texts)
                .collect::<Vec<(String, String)>>(),
        )
    } else {
        None
    };
    for (index, (input, output)) in declared.iter().enumerate() {
        let attempt = (|| -> Result<serde_json::Value> {
            require_engine_output(output)?;
            let validation = if let Some(texts) = &docx_texts {
                let (source_text, candidate_text) = &texts[index];
                ensure!(
                    normalize_document_text(source_text)
                        == normalize_document_text(candidate_text),
                    "converted DOCX text content differs from the source document"
                );
                "{\"level\":\"docx_package_and_batched_text_content\"}".to_owned()
            } else {
                validate_conversion(input, output, &source_ext, target, "lo", &guard.0)?
            };
            Ok(serde_json::json!({
                "index": index + 1,
                "status": "converted",
                "engine": engine_name,
                "validation": serde_json::from_str::<serde_json::Value>(&validation)
                    .unwrap_or_else(|_| serde_json::json!({"detail": validation})),
            }))
        })();
        match attempt {
            Ok(book) => books.push(book),
            Err(error) => {
                failures += 1;
                books.push(serde_json::json!({
                    "index": index + 1,
                    "status": "failed",
                    "input": input.display().to_string(),
                    "error": format!("{error:#}"),
                }));
            }
        }
    }
    let result = serde_json::json!({
        "status": if failures == 0 { "converted" } else { "partial" },
        "engine": engine_name,
        "convert_ms": convert_ms,
        "books": books,
    });
    println!("{}", serde_json::to_string(&result)?);
    Ok(())
}
