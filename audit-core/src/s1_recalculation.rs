//! Explicit S1 formula-cache recalculation using a native Windows COM session
//! or the existing Linux calculation sidecar.
//!
//! The input to this module is already a separate formula audit copy. It never
//! edits the submitted source workbook or the manifest's original audit copy.

use anyhow::{bail, ensure, Context, Result};
use serde_json::Value;
use std::{
    fs,
    io::{BufRead, BufReader, Write},
    path::{Path, PathBuf},
    process::{Child, ChildStdin, Command, Stdio},
    sync::mpsc::{self, Receiver},
    thread,
    time::{Instant, SystemTime, UNIX_EPOCH},
};

#[derive(Clone, Debug)]
pub struct S1RecalculationReport {
    pub engine: String,
    pub input: PathBuf,
    pub output: PathBuf,
    pub elapsed_ms: u128,
}

/// Reuses one native Windows Office COM session across an S1-F1 batch. The
/// older explicit sidecar session remains available for compatibility.
pub struct S1RecalculationSession {
    sidecar_path: PathBuf,
    native_office: bool,
    engine_request: String,
    child: Option<Child>,
    stdin: Option<ChildStdin>,
    responses: Option<Receiver<std::result::Result<String, String>>>,
    excel_process_id: Option<u32>,
    owned_excel_identity: Option<(u32, u64)>,
    owned_excel_cleanup_succeeded: Option<bool>,
    owns_excel_process: bool,
    #[cfg(windows)]
    native_process_handle: Option<usize>,
}

impl S1RecalculationSession {
    pub fn new(sidecar_path: impl AsRef<Path>, engine_request: &str) -> Result<Self> {
        ensure!(
            S1_CALC_ENGINES.contains(&engine_request),
            "S1 公式重算引擎必须是 auto、excel、wps 或 lo：{engine_request}"
        );
        let sidecar_path = fs::canonicalize(sidecar_path.as_ref()).with_context(|| {
            format!("读取公式重算 sidecar：{}", sidecar_path.as_ref().display())
        })?;
        ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");
        Ok(Self {
            sidecar_path,
            native_office: false,
            engine_request: engine_request.to_owned(),
            child: None,
            stdin: None,
            responses: None,
            excel_process_id: None,
            owned_excel_identity: None,
            owned_excel_cleanup_succeeded: None,
            owns_excel_process: false,
            #[cfg(windows)]
            native_process_handle: None,
        })
    }

    /// Use this Audit executable's native COM worker on Windows. The worker
    /// is isolated to keep a hung Office call from blocking the GUI forever;
    /// it calls COM directly and is reused throughout the batch.
    pub fn new_native(engine_request: &str) -> Result<Self> {
        let executable = std::env::current_exe().context("定位 Audit 原生 COM 工作进程")?;
        Self::new_native_with_executable(executable, engine_request)
    }

    /// Create an isolated COM session worker from an explicit Audit
    /// executable. This keeps worker process ownership testable from an
    /// integration-test host and supports managed host processes.
    #[doc(hidden)]
    pub fn new_native_with_executable(
        executable: impl AsRef<Path>,
        engine_request: &str,
    ) -> Result<Self> {
        ensure!(cfg!(windows), "原生 Office COM 公式重算仅支持 Windows");
        ensure!(
            S1_CALC_ENGINES.contains(&engine_request),
            "S1 公式重算引擎必须是 auto、excel、wps 或 lo：{engine_request}"
        );
        let executable = fs::canonicalize(executable.as_ref()).with_context(|| {
            format!(
                "读取 Audit 原生 COM 工作进程：{}",
                executable.as_ref().display()
            )
        })?;
        ensure!(executable.is_file(), "Audit 原生 COM 工作进程不是文件");
        Ok(Self {
            sidecar_path: executable,
            native_office: true,
            engine_request: engine_request.to_owned(),
            child: None,
            stdin: None,
            responses: None,
            excel_process_id: None,
            owned_excel_identity: None,
            owned_excel_cleanup_succeeded: None,
            owns_excel_process: false,
            #[cfg(windows)]
            native_process_handle: None,
        })
    }

    /// Identity of the COM process created by the most recent session start.
    /// The creation time prevents a recycled PID from being mistaken for it.
    #[doc(hidden)]
    pub fn last_owned_excel_identity(&self) -> Option<(u32, u64)> {
        self.owned_excel_identity
    }

    #[doc(hidden)]
    pub fn last_owned_excel_cleanup_succeeded(&self) -> Option<bool> {
        self.owned_excel_cleanup_succeeded
    }

    pub fn recalculate(
        &mut self,
        input_path: impl AsRef<Path>,
        output_path: impl AsRef<Path>,
    ) -> Result<S1RecalculationReport> {
        #[cfg(windows)]
        if self.native_office {
            return self.recalculate_native(input_path.as_ref(), output_path.as_ref());
        }
        let started = Instant::now();
        let (input_path, output_path, staging) =
            prepare_recalculation_paths(input_path.as_ref(), output_path.as_ref())?;
        let candidate = staging.path().join("recalculated.xlsx");
        if let Err(error) = self.ensure_started() {
            self.stop_process();
            return self.fallback_to_libreoffice(&input_path, &output_path, error);
        }
        let request = serde_json::json!({
            "type": "recalculate",
            "input": sidecar_io_path(&input_path).display().to_string(),
            "output": sidecar_io_path(&candidate).display().to_string(),
        });
        let request_result = (|| -> Result<Value> {
            let stdin = self
                .stdin
                .as_mut()
                .context("Excel COM session stdin is unavailable")?;
            serde_json::to_writer(&mut *stdin, &request).context("写入 Excel COM session 请求")?;
            stdin
                .write_all(b"\n")
                .context("结束 Excel COM session 请求")?;
            stdin.flush().context("刷新 Excel COM session 请求")?;
            let response = self.receive_response()?;
            let report: Value =
                serde_json::from_str(&response).context("解析 Excel COM session 的重算结果")?;
            ensure!(
                report.get("status").and_then(Value::as_str) == Some("recalculated"),
                "S1 公式重算失败（持久 COM session）：{}",
                report
                    .get("error")
                    .and_then(Value::as_str)
                    .unwrap_or("sidecar 未报告 recalculated 状态")
            );
            Ok(report)
        })();
        let report = match request_result {
            Ok(report) => report,
            Err(error) => {
                self.stop_process();
                return self.fallback_to_libreoffice(&input_path, &output_path, error);
            }
        };
        finish_recalculation(&input_path, &output_path, &candidate, report, started)
    }

    fn fallback_to_libreoffice(
        &mut self,
        input_path: &Path,
        output_path: &Path,
        error: anyhow::Error,
    ) -> Result<S1RecalculationReport> {
        if self.engine_request != "auto" {
            return Err(error);
        }
        #[cfg(windows)]
        if self.native_office {
            return self
                .recalculate_native_libreoffice(input_path, output_path)
                .with_context(|| {
                    format!("Office COM 热会话失败（{error:#}），LibreOffice 回退也失败")
                });
        }
        self.engine_request = "lo".to_owned();
        recalculate_s1_formula_copy(&self.sidecar_path, input_path, output_path, "lo")
            .with_context(|| format!("Office COM 热会话失败（{error:#}），LibreOffice 回退也失败"))
    }

    fn ensure_started(&mut self) -> Result<()> {
        if self.child.is_some() {
            return Ok(());
        }
        let result = self.start_process();
        if result.is_err() && self.child.is_some() {
            self.stop_process();
        }
        result
    }

    fn start_process(&mut self) -> Result<()> {
        let mut command = Command::new(&self.sidecar_path);
        // Keep automatic engine routing under the parent session so a stalled
        // COM call can be stopped before starting the bounded LibreOffice
        // fallback. The sidecar's `auto` mode would otherwise wait through
        // both engine timeouts in one response.
        let session_engine = if self.engine_request == "auto" {
            "com"
        } else {
            self.engine_request.as_str()
        };
        if self.native_office {
            command
                .arg("__audit-native-com-session")
                .arg(session_engine);
        } else {
            command
                .arg("recalculate-session")
                .arg("--engine")
                .arg(session_engine);
        }
        command
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null());
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(0x0800_0000); // CREATE_NO_WINDOW
        }
        let mut child = command.spawn().with_context(|| {
            if self.native_office {
                "启动 Audit 原生 COM 计算工作进程".to_owned()
            } else {
                format!("启动持久公式重算 sidecar：{}", self.sidecar_path.display())
            }
        })?;
        let stdin = child
            .stdin
            .take()
            .context("公式重算工作进程 stdin 未创建")?;
        let stdout = child
            .stdout
            .take()
            .context("公式重算工作进程 stdout 未创建")?;
        let (sender, receiver) = mpsc::channel();
        thread::spawn(move || {
            let mut reader = BufReader::new(stdout);
            loop {
                let mut line = String::new();
                match reader.read_line(&mut line) {
                    Ok(0) => {
                        let _ = sender.send(Err("sidecar session closed stdout".to_owned()));
                        break;
                    }
                    Ok(_) => {
                        if sender.send(Ok(line)).is_err() {
                            break;
                        }
                    }
                    Err(error) => {
                        let _ = sender.send(Err(format!("read sidecar session stdout: {error}")));
                        break;
                    }
                }
            }
        });
        self.child = Some(child);
        self.stdin = Some(stdin);
        self.responses = Some(receiver);

        let ready_line = self
            .receive_response()
            .context("等待持久公式重算工作进程就绪")?;
        let ready: Value =
            serde_json::from_str(&ready_line).context("解析持久公式重算工作进程就绪状态")?;
        ensure!(
            ready.get("status").and_then(Value::as_str) == Some("session_ready"),
            "持久公式重算工作进程未就绪：{}",
            ready_line.trim()
        );
        self.excel_process_id = ready
            .get("engine_process_id")
            .and_then(Value::as_u64)
            .and_then(|value| u32::try_from(value).ok());
        self.owns_excel_process = ready
            .get("engine_process_owned")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        let com_available = ready
            .get("com_available")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        #[cfg(windows)]
        if self.native_office && com_available {
            let process_identity = self
                .excel_process_id
                .zip(ready.get("engine_process_created").and_then(Value::as_u64));
            self.owned_excel_identity = process_identity;
            self.native_process_handle = process_identity.and_then(|(process_id, created)| {
                crate::native_office_com::open_verified_process(process_id, created)
            });
            if !self.owns_excel_process || self.native_process_handle.is_none() {
                self.shutdown_unverified_native_worker();
                bail!("无法安全确认原生 COM 创建的 Office 进程，未开始重算");
            }
        }
        if self.engine_request == "auto" && !com_available {
            let reason = ready
                .get("com_error")
                .and_then(Value::as_str)
                .unwrap_or("sidecar 未能初始化 Office COM");
            self.stop_process();
            if self.native_office {
                bail!("Office COM 不可用：{reason}");
            }
            self.engine_request = "lo".to_owned();
            return self.start_process().with_context(|| {
                format!("Office COM 不可用（{reason}），启动 LibreOffice 会话失败")
            });
        }
        Ok(())
    }

    fn receive_response(&mut self) -> Result<String> {
        let timeout = recalculation_session_timeout();
        let response = self
            .responses
            .as_ref()
            .context("Excel COM session response reader is unavailable")?
            .recv_timeout(timeout);
        match response {
            Ok(Ok(line)) => Ok(line),
            Ok(Err(error)) => {
                self.stop_process();
                bail!("S1 持久 Excel COM session 失败：{error}")
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {
                self.stop_process();
                bail!(
                    "S1 持久 Excel COM session 超过 {} 秒无响应；已停止该 session",
                    timeout.as_secs()
                )
            }
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                self.stop_process();
                bail!("S1 持久 Excel COM session 响应通道已关闭")
            }
        }
    }

    fn stop_process(&mut self) {
        if let Some(mut child) = self.child.take() {
            let _ = child.kill();
            let _ = child.wait();
        }
        self.stdin.take();
        self.responses.take();
        self.terminate_owned_excel();
    }

    fn terminate_owned_excel(&mut self) {
        #[cfg(windows)]
        if self.native_office {
            if let Some(handle) = self.native_process_handle.take() {
                self.owned_excel_cleanup_succeeded = Some(
                    crate::native_office_com::close_verified_process(handle, true),
                );
            } else if self.owns_excel_process {
                self.owned_excel_cleanup_succeeded = Some(false);
            }
            self.excel_process_id = None;
            self.owns_excel_process = false;
            return;
        }
        if !self.owns_excel_process {
            return;
        }
        let Some(process_id) = self.excel_process_id.take() else {
            return;
        };
        #[cfg(windows)]
        {
            let _ = Command::new("taskkill.exe")
                .args(["/PID", &process_id.to_string(), "/T", "/F"])
                .stdout(Stdio::null())
                .stderr(Stdio::null())
                .status();
        }
    }
}

#[cfg(windows)]
impl S1RecalculationSession {
    fn shutdown_unverified_native_worker(&mut self) {
        if let Some(stdin) = self.stdin.as_mut() {
            let _ = serde_json::to_writer(&mut *stdin, &serde_json::json!({"type":"shutdown"}));
            let _ = stdin.write_all(b"\n");
            let _ = stdin.flush();
        }
        if let Some(mut child) = self.child.take() {
            let deadline = Instant::now() + std::time::Duration::from_secs(5);
            loop {
                match child.try_wait() {
                    Ok(Some(_)) => break,
                    Ok(None) if Instant::now() < deadline => {
                        thread::sleep(std::time::Duration::from_millis(50));
                    }
                    _ => {
                        let _ = child.kill();
                        let _ = child.wait();
                        break;
                    }
                }
            }
        }
        self.stdin.take();
        self.responses.take();
        self.excel_process_id = None;
        self.owns_excel_process = false;
    }

    fn recalculate_native(&mut self, input: &Path, output: &Path) -> Result<S1RecalculationReport> {
        if self.engine_request == "lo" {
            return self.recalculate_native_libreoffice(input, output);
        }
        let started = Instant::now();
        let (input, output, destination_staging) = prepare_recalculation_paths(input, output)?;
        // Excel receives short temporary paths; the final candidate is staged
        // beside the destination so publication remains atomic and no-overwrite.
        let engine_staging = StagingDirectory::create(&std::env::temp_dir())?;
        let prepared =
            crate::native_recalculation::prepare_engine_input(&input, engine_staging.path(), true)?;
        let engine_output = engine_staging.path().join("native-com-result.xlsx");
        let candidate = destination_staging.path().join("recalculated.xlsx");

        let attempt = (|| -> Result<Value> {
            self.ensure_started()?;
            let request = serde_json::json!({
                "type": "recalculate",
                "input": sidecar_io_path(&prepared).display().to_string(),
                "output": sidecar_io_path(&engine_output).display().to_string(),
            });
            let stdin = self.stdin.as_mut().context("原生 COM 会话输入不可用")?;
            serde_json::to_writer(&mut *stdin, &request)?;
            stdin.write_all(b"\n")?;
            stdin.flush()?;
            let response = self.receive_response()?;
            let report: Value = serde_json::from_str(&response).context("解析原生 COM 重算结果")?;
            ensure!(
                report.get("status").and_then(Value::as_str) == Some("recalculated"),
                "原生 COM 重算失败：{}",
                report
                    .get("error")
                    .and_then(Value::as_str)
                    .unwrap_or("未知错误")
            );
            ensure!(engine_output.is_file(), "原生 COM 没有生成重算工作簿");
            let normalized = crate::native_recalculation::normalize_engine_output(
                &input,
                &engine_output,
                engine_staging.path(),
            )?;
            fs::copy(normalized, &candidate).context("暂存原生 COM 重算结果")?;
            validate_recalculated_copy(&input, &candidate)?;
            Ok(report)
        })();
        let report = match attempt {
            Ok(report) => report,
            Err(error) => {
                self.stop_process();
                return self.fallback_to_libreoffice(&input, &output, error);
            }
        };
        finish_recalculation(&input, &output, &candidate, report, started)
    }

    fn recalculate_native_libreoffice(
        &mut self,
        input: &Path,
        output: &Path,
    ) -> Result<S1RecalculationReport> {
        let started = Instant::now();
        let (input, output, destination_staging) = prepare_recalculation_paths(input, output)?;
        let engine_staging = StagingDirectory::create(&std::env::temp_dir())?;
        let prepared = crate::native_recalculation::prepare_engine_input(
            &input,
            engine_staging.path(),
            false,
        )?;
        let engine_output =
            crate::native_recalculation::run_libreoffice(&prepared, engine_staging.path())?;
        let normalized = crate::native_recalculation::normalize_engine_output(
            &input,
            &engine_output,
            engine_staging.path(),
        )?;
        let candidate = destination_staging.path().join("recalculated.xlsx");
        fs::copy(normalized, &candidate).context("暂存 LibreOffice 重算结果")?;
        let report = serde_json::json!({"status":"recalculated", "engine":"LibreOffice"});
        finish_recalculation(&input, &output, &candidate, report, started)
    }
}

impl Drop for S1RecalculationSession {
    fn drop(&mut self) {
        let mut shutdown_acknowledged = !self.native_office;
        if let Some(stdin) = self.stdin.as_mut() {
            let _ = serde_json::to_writer(&mut *stdin, &serde_json::json!({"type":"shutdown"}));
            let _ = stdin.write_all(b"\n");
            let _ = stdin.flush();
            if self.native_office {
                if let Some(responses) = self.responses.as_ref() {
                    shutdown_acknowledged = responses
                        .recv_timeout(std::time::Duration::from_secs(10))
                        .ok()
                        .and_then(|line| line.ok())
                        .and_then(|line| serde_json::from_str::<Value>(&line).ok())
                        .is_some_and(|response| {
                            response.get("status").and_then(Value::as_str)
                                == Some("shutdown_complete")
                                && response
                                    .get("excel_quit_succeeded")
                                    .and_then(Value::as_bool)
                                    .unwrap_or(false)
                        });
                }
            }
        }
        let Some(mut child) = self.child.take() else {
            return;
        };
        self.stdin.take();
        self.responses.take();
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(5);
        let mut clean_shutdown = false;
        loop {
            match child.try_wait() {
                Ok(Some(status)) => {
                    clean_shutdown = status.success() && shutdown_acknowledged;
                    break;
                }
                Err(_) => break,
                Ok(None) => {}
            }
            if std::time::Instant::now() >= deadline {
                let _ = child.kill();
                let _ = child.wait();
                self.terminate_owned_excel();
                return;
            }
            thread::sleep(std::time::Duration::from_millis(50));
        }
        // A normal worker exit is not enough: confirm that Office accepted
        // Quit and that the exact process created by this session exited.
        if !clean_shutdown {
            self.terminate_owned_excel();
        } else {
            #[cfg(windows)]
            if let Some(handle) = self.native_process_handle.take() {
                if crate::native_office_com::wait_verified_process_exit(
                    handle,
                    std::time::Duration::from_secs(5),
                ) {
                    self.owned_excel_cleanup_succeeded = Some(
                        crate::native_office_com::close_verified_process(handle, false),
                    );
                } else {
                    self.owned_excel_cleanup_succeeded = Some(
                        crate::native_office_com::close_verified_process(handle, true),
                    );
                }
            }
        }
    }
}

fn recalculation_session_timeout() -> std::time::Duration {
    let seconds = std::env::var("AUDIT_RECALC_TIMEOUT_SECONDS")
        .ok()
        .and_then(|value| value.parse::<u64>().ok())
        .unwrap_or(300)
        .clamp(5, 3600);
    std::time::Duration::from_secs(seconds.saturating_add(30))
}

/// Require explicit Office/LibreOffice recalculation for a production S1 run.
/// Reading an existing formula cache remains available only as a clearly
/// identified diagnostic mode for Python/Rust parity and troubleshooting.
pub fn validate_s1_formula_cache_mode(
    sidecar_requested: bool,
    allow_saved_cache_diagnostics: bool,
) -> Result<()> {
    ensure!(
        sidecar_requested || allow_saved_cache_diagnostics,
        "S1-F01 正式审核必须强制 Office/LibreOffice 重算（Windows 可使用原生 COM）；仅作缓存诊断时显式使用 --allow-saved-formula-cache-diagnostics"
    );
    ensure!(
        !sidecar_requested || !allow_saved_cache_diagnostics,
        "--allow-saved-formula-cache-diagnostics 仅适用于不执行公式重算的诊断模式"
    );
    Ok(())
}

/// Linux 批次重算：一次 sidecar 调用（`recalculate-batch`）顺序处理整批
/// 工作簿——单次 LibreOffice 启动完成全部转换（rules/13 §4.5 的批次会话
/// 复用）。每本的公式缓存清理、重算标记、规范化、验证与不覆盖发布均在
/// sidecar 内完成；父级随后对本批成功项再执行 `validate_recalculated_copy`
/// 复核，保证检查不因批量化而减少。返回与 pairs 等长的结果数组：成功项
/// 为 Some(报告)，失败项为 None（由调用方决定逐本回退）。
pub fn recalculate_s1_formula_copy_batch(
    sidecar_path: impl AsRef<Path>,
    pairs: &[(PathBuf, PathBuf)],
    engine_request: &str,
) -> Result<Vec<Option<S1RecalculationReport>>> {
    ensure!(
        engine_request == "auto" || engine_request == "lo",
        "S1 批次重算仅支持 auto/lo（LibreOffice）：{engine_request}"
    );
    ensure!(!pairs.is_empty(), "S1 批次重算没有输入工作簿");
    let sidecar_path = fs::canonicalize(sidecar_path.as_ref())
        .with_context(|| format!("读取公式重算 sidecar：{}", sidecar_path.as_ref().display()))?;
    ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");
    for (input, output) in pairs {
        ensure!(input.is_file(), "S1 批次重算输入不是文件：{}", input.display());
        ensure!(is_xlsx(input), "S1 批次重算输入必须为 .xlsx");
        ensure!(is_xlsx(output), "S1 批次重算输出必须为 .xlsx");
        ensure!(
            !output.exists(),
            "拒绝覆盖 S1 批次重算输出：{}",
            output.display()
        );
    }
    let started = Instant::now();
    let pairs_file = unique_pairs_file(pairs)?;
    let command = Command::new(&sidecar_path)
        .arg("recalculate-batch")
        .arg("--pairs")
        .arg(&pairs_file)
        .arg("--engine")
        .arg(engine_request)
        .output()
        .with_context(|| format!("启动批次公式重算 sidecar：{}", sidecar_path.display()))?;
    let _ = fs::remove_file(&pairs_file);
    if !command.status.success() {
        bail!(
            "S1 批次公式重算失败（sidecar exit={}）：{}{}",
            command.status,
            String::from_utf8_lossy(&command.stderr).trim(),
            stdout_suffix(&command.stdout)
        );
    }
    let report: Value = serde_json::from_slice(&command.stdout)
        .context("解析批次公式重算 sidecar 的 JSON 结果")?;
    let books = report
        .get("books")
        .and_then(Value::as_array)
        .context("批次公式重算结果缺少 books 数组")?;
    anyhow::ensure!(
        books.len() == pairs.len(),
        "批次公式重算返回的本数不一致：{} / {}",
        books.len(),
        pairs.len()
    );
    let mut outcomes = Vec::with_capacity(pairs.len());
    for (index, book) in books.iter().enumerate() {
        let (input, output) = &pairs[index];
        if book.get("status").and_then(Value::as_str) != Some("recalculated") {
            outcomes.push(None);
            continue;
        }
        let engine = book
            .get("engine")
            .and_then(Value::as_str)
            .context("批次重算结果缺少引擎名")?
            .to_owned();
        validate_s1_engine(&engine)?;
        validate_recalculated_copy(input, output)?;
        outcomes.push(Some(S1RecalculationReport {
            engine,
            input: input.clone(),
            output: output.clone(),
            elapsed_ms: started.elapsed().as_millis(),
        }));
    }
    Ok(outcomes)
}

/// 把批次输入/输出对写入临时 TSV（tab 分隔；路径含 tab/换行时明确拒绝）。
fn unique_pairs_file(pairs: &[(PathBuf, PathBuf)]) -> Result<PathBuf> {
    let mut content = String::new();
    for (input, output) in pairs {
        let input = input.to_string_lossy();
        let output = output.to_string_lossy();
        ensure!(
            !input.contains(['\t', '\n', '\r']) && !output.contains(['\t', '\n', '\r']),
            "工作簿路径包含批次清单不允许的制表/换行字符"
        );
        content.push_str(&format!("{input}\t{output}\n"));
    }
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit-s1-batch-pairs-{}-{nonce}.tsv",
        std::process::id()
    ));
    fs::write(&path, content).context("写入批次重算清单")?;
    Ok(path)
}

/// Recalculate an existing S1 formula copy through the separately built
/// `lo-convert` sidecar, then publish the validated output without overwriting.
///
/// `lo-convert` selects the configured native Office COM channel on Windows
/// and LibreOffice Calc on Linux. Automatic Windows mode tries the supported
/// COM engine before LibreOffice. The sidecar validates formula/cache
/// presence; the original copied workbook remains untouched.
/// Sidecar `--engine` values accepted for S1 formula recalculation. `auto`
/// keeps the platform default order, `excel`/`wps` force one COM application,
/// `lo` forces LibreOffice headless.
pub const S1_CALC_ENGINES: [&str; 4] = ["auto", "excel", "wps", "lo"];

pub fn recalculate_s1_formula_copy(
    sidecar_path: impl AsRef<Path>,
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    engine_request: &str,
) -> Result<S1RecalculationReport> {
    ensure!(
        S1_CALC_ENGINES.contains(&engine_request),
        "S1 公式重算引擎必须是 auto、excel、wps 或 lo：{engine_request}"
    );
    let sidecar_path = fs::canonicalize(sidecar_path.as_ref())
        .with_context(|| format!("读取公式重算 sidecar：{}", sidecar_path.as_ref().display()))?;
    ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");
    let started = Instant::now();
    let (input_path, output_path, staging) =
        prepare_recalculation_paths(input_path.as_ref(), output_path.as_ref())?;
    let candidate = staging.path().join("recalculated.xlsx");
    let sidecar_input = sidecar_io_path(&input_path);
    let sidecar_output = sidecar_io_path(&candidate);
    let command = Command::new(&sidecar_path)
        .arg("recalculate")
        .arg("--input")
        .arg(&sidecar_input)
        .arg("--output")
        .arg(&sidecar_output)
        .arg("--engine")
        .arg(engine_request)
        .output()
        .with_context(|| format!("启动公式重算 sidecar：{}", sidecar_path.display()))?;
    if !command.status.success() {
        bail!(
            "S1 公式重算失败（sidecar exit={}）：{}{}",
            command.status,
            String::from_utf8_lossy(&command.stderr).trim(),
            stdout_suffix(&command.stdout)
        );
    }

    let report: Value =
        serde_json::from_slice(&command.stdout).context("解析公式重算 sidecar 的 JSON 结果")?;
    finish_recalculation(&input_path, &output_path, &candidate, report, started)
}

fn prepare_recalculation_paths(
    input_path: &Path,
    requested_output: &Path,
) -> Result<(PathBuf, PathBuf, StagingDirectory)> {
    let input_path = fs::canonicalize(input_path)
        .with_context(|| format!("读取 S1 公式审核副本：{}", input_path.display()))?;
    ensure!(input_path.is_file(), "S1 公式审核副本不是文件");
    ensure!(
        is_xlsx(&input_path),
        "S1 真实公式重算目前仅支持 .xlsx 审核副本"
    );
    ensure!(is_xlsx(requested_output), "S1 公式重算输出必须为 .xlsx");
    let output_parent = requested_output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let output_parent = fs::canonicalize(output_parent)
        .with_context(|| format!("读取 S1 重算输出目录：{}", output_parent.display()))?;
    ensure!(output_parent.is_dir(), "S1 重算输出路径的父级不是目录");
    let output_name = requested_output
        .file_name()
        .context("S1 重算输出必须包含文件名")?;
    let output_path = output_parent.join(output_name);
    ensure!(
        !output_path.exists(),
        "拒绝覆盖 S1 重算输出：{}",
        output_path.display()
    );
    ensure!(
        !same_path(&input_path, &output_path),
        "S1 公式重算必须写入独立于输入副本的新文件"
    );
    let staging = StagingDirectory::create(&output_parent)?;
    Ok((input_path, output_path, staging))
}

fn finish_recalculation(
    input_path: &Path,
    output_path: &Path,
    candidate: &Path,
    report: Value,
    started: Instant,
) -> Result<S1RecalculationReport> {
    ensure!(
        report.get("status").and_then(Value::as_str) == Some("recalculated"),
        "公式重算工作进程未报告 recalculated 状态：{}",
        report
            .get("error")
            .and_then(Value::as_str)
            .unwrap_or("未知错误")
    );
    let engine = report
        .get("engine")
        .and_then(Value::as_str)
        .context("公式重算工作进程未报告实际引擎")?
        .to_owned();
    validate_s1_engine(&engine)?;
    ensure!(
        candidate.is_file() && fs::metadata(candidate)?.len() > 0,
        "公式重算工作进程未生成有效输出工作簿"
    );
    validate_recalculated_copy(input_path, candidate)?;
    // The staging directory is inside the destination directory, so a hard
    // link publishes atomically and refuses a concurrent destination.
    fs::hard_link(candidate, output_path).with_context(|| {
        format!(
            "安全发布 S1 重算副本失败（输出不得覆盖）：{}",
            output_path.display()
        )
    })?;
    Ok(S1RecalculationReport {
        engine,
        input: input_path.to_path_buf(),
        output: output_path.to_path_buf(),
        elapsed_ms: started.elapsed().as_millis(),
    })
}

/// Independently validate formula preservation, cached-value node completeness
/// and typed cached-result readability before publishing the recalculated
/// candidate. This proves the calculation result is consumable, not that it
/// matches a business-specific expected value.
fn validate_recalculated_copy(input: &Path, output: &Path) -> Result<()> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .with_context(|| format!("重开重算前的 S1 公式副本：{}", input.display()))?;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(output)
        .with_context(|| format!("重开重算后的 S1 公式副本：{}", output.display()))?;
    let source = source.workbook;
    let result = result.workbook;
    ensure!(
        excel_reader::formula_texts_equivalent_after_recalculation(&source, &result),
        "S1 重算修改了工作表结构、公式位置或公式文本（不含 LibreOffice 对 TRUE/FALSE 字面量的等价规范化）"
    );
    let formula_count = source
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    ensure!(formula_count > 0, "S1 公式副本中没有可重算的公式");

    let cache_inventory = excel_reader::inspect_formula_cache_nodes(output)
        .context("检查 S1 重算副本中的公式缓存节点")?;
    ensure!(
        cache_inventory.formula_cell_count == formula_count,
        "S1 重算副本公式数量不一致：公式文本 {formula_count} 个，OOXML 公式单元格 {} 个",
        cache_inventory.formula_cell_count
    );
    ensure!(
        cache_inventory.cached_formula_cell_count == formula_count
            && cache_inventory.missing_cache_cells.is_empty(),
        "S1 重算副本存在缺失公式缓存节点：{}",
        cache_inventory
            .missing_cache_cells
            .first()
            .map(|location| format!("{}!{}", location.sheet_name, location.cell))
            .unwrap_or_else(|| "缓存计数不一致".to_owned())
    );
    let result_values = excel_reader::formula_result_value_coverage(&result);
    ensure!(
        result_values.formula_cell_count == formula_count
            && result_values.typed_result_cell_count == formula_count
            && result_values.missing_result_cells.is_empty(),
        "S1 重算副本有公式结果无法通过 Calamine 读取：{}",
        result_values
            .missing_result_cells
            .first()
            .map(|location| format!("{}!{}", location.sheet_name, location.cell))
            .unwrap_or_else(|| format!(
                "公式 {} 个，可读结果 {} 个",
                result_values.formula_cell_count, result_values.typed_result_cell_count
            ))
    );
    Ok(())
}

pub fn s1_recalculated_filename(index: usize, source_path: impl AsRef<Path>) -> Result<String> {
    ensure!(index > 0, "S1 重算副本序号必须从 1 开始");
    let name = source_path
        .as_ref()
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.trim().is_empty())
        .unwrap_or("audit-copy");
    let safe = name
        .chars()
        .take(80)
        .map(|character| {
            if character.is_alphanumeric() || matches!(character, '-' | '_' | '.') {
                character
            } else {
                '_'
            }
        })
        .collect::<String>();
    Ok(format!("{index:04}-{safe}-recalculated.xlsx"))
}

fn validate_s1_engine(engine: &str) -> Result<()> {
    #[cfg(windows)]
    ensure!(
        engine == "Excel.Application"
            || engine == "ket.Application"
            || engine == "KET.Application"
            || engine == "LibreOffice",
        "S1 Windows 公式重算实际引擎不受支持：“{engine}”，仅接受 Microsoft Excel COM、WPS 表格 COM 或 LibreOffice"
    );
    #[cfg(not(windows))]
    ensure!(
        engine == "LibreOffice",
        "S1 Linux 公式重算要求 LibreOffice Calc；sidecar 实际使用“{engine}”，拒绝将其结果作为审核输入"
    );
    Ok(())
}

fn is_xlsx(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}

fn same_path(left: &Path, right: &Path) -> bool {
    #[cfg(windows)]
    {
        left.to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy())
    }
    #[cfg(not(windows))]
    {
        left == right
    }
}

fn sidecar_io_path(path: &Path) -> PathBuf {
    #[cfg(windows)]
    {
        let path = path.to_string_lossy();
        if let Some(path) = path.strip_prefix(r"\\?\UNC\") {
            return PathBuf::from(format!(r"\\{path}"));
        }
        if let Some(path) = path.strip_prefix(r"\\?\") {
            return PathBuf::from(path);
        }
        PathBuf::from(path.as_ref())
    }
    #[cfg(not(windows))]
    {
        path.to_path_buf()
    }
}

fn stdout_suffix(stdout: &[u8]) -> String {
    let text = String::from_utf8_lossy(stdout).trim().to_owned();
    if text.is_empty() {
        String::new()
    } else {
        format!("；stdout：{text}")
    }
}

struct StagingDirectory(PathBuf);

impl StagingDirectory {
    fn create(parent: &Path) -> Result<Self> {
        let timestamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("读取系统时间")?
            .as_nanos();
        for attempt in 0..100 {
            let path = parent.join(format!(
                ".s1-recalculate-{}-{timestamp}-{attempt}",
                std::process::id()
            ));
            match fs::create_dir(&path) {
                Ok(()) => return Ok(Self(path)),
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建 S1 公式重算临时目录：{}", path.display()));
                }
            }
        }
        bail!("无法为 S1 公式重算创建唯一临时目录")
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for StagingDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::CellValue;
    use excel_writer::Workbook;
    use std::time::UNIX_EPOCH;

    #[test]
    fn recalculated_filename_is_safe_and_keeps_original_stem() {
        assert_eq!(
            s1_recalculated_filename(2, "个人贷款:2026.xlsx").unwrap(),
            "0002-个人贷款_2026-recalculated.xlsx"
        );
        assert!(s1_recalculated_filename(0, "sample.xlsx").is_err());
    }

    #[test]
    fn engine_must_be_supported_on_the_current_platform() {
        // The recalculation channels cover Microsoft Excel COM, WPS
        // Spreadsheets COM (native IDispatch) and LibreOffice per the
        // 高级设置 contract; unknown engines stay rejected.
        #[cfg(windows)]
        {
            assert!(validate_s1_engine("Excel.Application").is_ok());
            assert!(validate_s1_engine("KET.Application").is_ok());
            assert!(validate_s1_engine("ket.Application").is_ok());
            assert!(validate_s1_engine("LibreOffice").is_ok());
            assert!(validate_s1_engine("wps.Application").is_err());
            assert!(validate_s1_engine("pywin32").is_err());
        }
        #[cfg(not(windows))]
        {
            assert!(validate_s1_engine("LibreOffice").is_ok());
            assert!(validate_s1_engine("Excel.Application").is_err());
            assert!(validate_s1_engine("KET.Application").is_err());
        }
    }

    #[test]
    fn cached_formula_values_require_an_explicit_diagnostic_mode() {
        let error = validate_s1_formula_cache_mode(false, false)
            .expect_err("stored caches must not run as a production S1 audit");
        assert!(error.to_string().contains("强制 Office/LibreOffice 重算"));
        assert!(validate_s1_formula_cache_mode(false, true).is_ok());
        assert!(validate_s1_formula_cache_mode(true, false).is_ok());
        assert!(validate_s1_formula_cache_mode(true, true).is_err());
    }

    #[cfg(windows)]
    #[test]
    fn sidecar_paths_remove_windows_verbatim_prefix_for_com() {
        assert_eq!(
            sidecar_io_path(Path::new(r"\\?\C:\temp\audit.xlsx")),
            PathBuf::from(r"C:\temp\audit.xlsx")
        );
        assert_eq!(
            sidecar_io_path(Path::new(r"\\?\UNC\server\share\audit.xlsx")),
            PathBuf::from(r"\\server\share\audit.xlsx")
        );
    }

    #[test]
    fn refuses_to_overwrite_an_existing_output() -> Result<()> {
        let root = TestDirectory::create()?;
        let input = root.join("audit-copy.xlsx");
        let output = root.join("recalculated.xlsx");
        let sidecar = root.join("sidecar.exe");
        fs::write(&input, b"source-copy")?;
        fs::write(&sidecar, b"test-executable")?;
        fs::write(&output, b"existing-output")?;
        let input_before = fs::read(&input)?;

        let error = recalculate_s1_formula_copy(&sidecar, &input, &output, "auto")
            .expect_err("existing output must be refused");
        assert!(format!("{error:#}").contains("拒绝覆盖 S1 重算输出"));
        assert_eq!(fs::read(&input)?, input_before);
        assert_eq!(fs::read(&output)?, b"existing-output");
        Ok(())
    }

    #[test]
    #[ignore = "requires AUDIT_S1_RECALC_TEST_SIDECAR and a real supported calculation engine"]
    fn explicit_sidecar_recalculates_synthetic_formula_copy_without_changing_input() -> Result<()> {
        let sidecar = std::env::var_os("AUDIT_S1_RECALC_TEST_SIDECAR")
            .map(PathBuf::from)
            .context(
                "AUDIT_S1_RECALC_TEST_SIDECAR is required for the explicit integration test",
            )?;
        let root = TestDirectory::create()?;
        let input = root.join("synthetic-formula-copy.xlsx");
        let output = root.join("synthetic-formula-recalculated.xlsx");
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_number(0, 0, 2.0)?;
        sheet.write_number(0, 1, 3.0)?;
        sheet.write_formula(0, 2, "=SUM(A1:B1)")?;
        workbook.save(&input)?;
        let input_before = fs::read(&input)?;

        let recalculation = recalculate_s1_formula_copy(&sidecar, &input, &output, "auto");
        assert_eq!(fs::read(&input)?, input_before);
        if recalculation.is_err() {
            assert!(
                !output.exists(),
                "failed calculation must not publish an output"
            );
        }
        let report = recalculation?;
        assert_eq!(report.output, fs::canonicalize(&output)?);
        assert!(output.is_file());
        assert_eq!(fs::read(&input)?, input_before);
        let values = excel_reader::read_value_columns(&output, 3)?;
        let calculated = values[0]
            .values
            .get("C1")
            .context("recalculated formula cache missing from C1")?;
        match calculated {
            CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
                assert!((*value - 5.0).abs() < f64::EPSILON)
            }
            CellValue::Integer(value) => assert_eq!(*value, 5),
            other => panic!("unexpected recalculated C1 value: {other:?}"),
        }
        Ok(())
    }

    struct TestDirectory(PathBuf);

    impl TestDirectory {
        fn create() -> Result<Self> {
            let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let root = std::env::temp_dir().join(format!(
                "audit_s1_recalculation_{}_{}",
                std::process::id(),
                nonce
            ));
            fs::create_dir(&root)?;
            Ok(Self(root))
        }
    }

    impl std::ops::Deref for TestDirectory {
        type Target = Path;

        fn deref(&self) -> &Self::Target {
            &self.0
        }
    }

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
}
