//! Explicit S1 formula-cache recalculation using a native Windows COM session
//! or the existing Linux calculation sidecar.
//!
//! The input to this module is already a separate formula audit copy. It never
//! edits the submitted source workbook or the manifest's original audit copy.

use rayon::prelude::*;

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
    /// 非致命提示（公式文本被引擎等价改写等），由界面写进详细运行日志。
    pub warnings: Vec<String>,
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
    active_engine: Option<String>,
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
            active_engine: None,
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
            active_engine: None,
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

    /// The originally requested engine (`auto`/`excel`/`wps`/`lo`); the active
    /// engine may differ after a fallback. S1 副本修复等路由决策依赖它。
    pub fn engine_request(&self) -> &str {
        &self.engine_request
    }

    #[doc(hidden)]
    pub fn last_owned_excel_cleanup_succeeded(&self) -> Option<bool> {
        self.owned_excel_cleanup_succeeded
    }

    /// S1-F01 校验区域重算的引擎路由入口。
    ///
    /// 优先级（用户确认 2026-11-08）：**betteroffice（纯 Rust）→ COM（Windows）
    /// → LibreOffice（Linux）→ 明确报错**。betteroffice 拒绝（预算耗尽 /
    /// #NAME? / 循环引用 / 打开失败）时按 `PYTOOLS_BO_CALC_FALLBACK` 回退，
    /// 回退原因写进报告 engine 字段留痕；`FALLBACK=none` 时明确失败，绝不静默
    /// 采用受限结果。
    pub fn recalculate(
        &mut self,
        input_path: impl AsRef<Path>,
        output_path: impl AsRef<Path>,
    ) -> Result<S1RecalculationReport> {
        let bo_config = crate::bo_calc::BoCalcConfig::from_app();
        crate::gui_diagnostics::trace("calculation_route", serde_json::json!({"file":input_path.as_ref(),
            "engineRequest":self.engine_request,"betterofficeEnabled":bo_config.enabled,"rounds":bo_config.rounds,
            "fallbackCom":bo_config.fallback_com,"fallbackLibreOffice":bo_config.fallback_libreoffice}));
        let mut bo_fallback_reason: Option<String> = None;
        if bo_config.enabled {
            match self.recalculate_betteroffice(
                input_path.as_ref(),
                output_path.as_ref(),
                &bo_config,
            ) {
                Ok(report) => return Ok(report),
                Err(error) => {
                    let reason = format!("{error:#}");
                    crate::gui_diagnostics::trace("betteroffice_refused",serde_json::json!({"file":input_path.as_ref(),"reason":reason}));
                    if !self.bo_fallback_allowed(&bo_config) {
                        bail!("betteroffice 校验区域重算不可采用且未配置可用回退：{reason}");
                    }
                    bo_fallback_reason = Some(format!("betteroffice 回退（{reason}）"));
                }
            }
        }
        let mut report = self.recalculate_platform(input_path, output_path)?;
        if let Some(reason) = bo_fallback_reason {
            report.engine = format!("{}；{}", report.engine, reason);
        }
        Ok(report)
    }

    /// whether the configured fallback can serve this platform at all.
    fn bo_fallback_allowed(&self, config: &crate::bo_calc::BoCalcConfig) -> bool {
        if !config.fallback_available() {
            return false;
        }
        // 用户显式选定单一引擎（excel/wps/lo）时不额外插入 betteroffice 回退链，
        // 保持既有语义：lo 直接走 LibreOffice，excel/wps 走原生 COM。
        self.engine_request == "auto"
    }

    /// betteroffice 首选尝试：在工作副本上原地重算，成功后按既有验证器校验
    /// 并原子发布（与 COM/LO 同一条 finish_recalculation 路径）。
    fn recalculate_betteroffice(
        &mut self,
        input_path: &Path,
        output_path: &Path,
        config: &crate::bo_calc::BoCalcConfig,
    ) -> Result<S1RecalculationReport> {
        let started = Instant::now();
        let (input, output, destination_staging) =
            prepare_recalculation_paths(input_path, output_path)?;
        // 安全边界：调用方传入的 input 可能是用户可见的审核副本。规范化
        // **只允许改我们自己的临时工作副本**，绝不回写调用方文件。
        //
        // 同时必须保留一份「已规范化、未重算」的基线：后续
        // validate_recalculated_copy 拿它与重算产物比对。若拿未规范化的
        // 调用方文件当基线，shared ref 归一这类规范化会被误判成「重算改变
        // 公式」（实测直接让 ！40 / ！70 两条路由全部失败）。
        let baseline = destination_staging.path().join("recalc-baseline.xlsx");
        std::fs::copy(&input, &baseline)
            .with_context(|| format!("为 betteroffice 重算准备基线副本：{}", baseline.display()))?;
        let repair = crate::excel_open_repair::repair_audit_copy_for_excel(&baseline);
        if repair.changed {
            eprintln!(
                "S1-F01 审核副本 OOXML 规范化（betteroffice 路由）：shared_ref={} quoted={} external={} pane={}",
                repair.shared_formula_ref_fixed, repair.quoted,
                repair.external_ref_fixed, repair.pane_fixed
            );
        }
        let candidate = destination_staging.path().join("recalculated.xlsx");
        std::fs::copy(&baseline, &candidate).with_context(|| {
            format!("为 betteroffice 重算准备工作副本：{}", candidate.display())
        })?;
        let outcome = crate::bo_calc::recalculate(&candidate, config)?;
        let report = serde_json::json!({
            "status": "recalculated",
            "engine": outcome.application,
        });
        let mut report = finish_recalculation(&baseline, &output, &candidate, report, started)?;
        report.elapsed_ms = started.elapsed().as_millis();
        Ok(report)
    }

    /// 平台引擎分支（COM / LibreOffice）——也是 betteroffice 拒绝后的回退目标。
    fn recalculate_platform(
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
        // 与所有引擎路径同一口径：sidecar 读规范化基线副本，比对也用基线，
        // 调用方文件保持原样（shared ref 归一不再被误判成「改公式」）。
        let baseline = staging.path().join("recalc-baseline.xlsx");
        std::fs::copy(&input_path, &baseline)
            .with_context(|| format!("为 sidecar 重算准备基线副本：{}", baseline.display()))?;
        let _ = crate::excel_open_repair::repair_audit_copy_for_excel(&baseline);
        let candidate = staging.path().join("recalculated.xlsx");
        if let Err(error) = self.ensure_started() {
            self.stop_process();
            return self.fallback_to_libreoffice(&input_path, &output_path, error);
        }
        let request = serde_json::json!({
            "type": "recalculate",
            "input": sidecar_io_path(&baseline).display().to_string(),
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
        finish_recalculation(&baseline, &output_path, &candidate, report, started)
    }

    /// Scan actual Excel DisplayFormat fills in the named S1 conditional-format
    /// areas. The existing isolated native COM worker is reused for the whole
    /// batch; it never attaches to or closes a user-owned Office instance.
    #[cfg(windows)]
    pub fn scan_conditional_display_colors(
        &mut self,
        input_path: impl AsRef<Path>,
        scan_areas: &[String],
    ) -> Result<crate::conditional_format::ConditionalDisplayScan> {
        ensure!(
            self.native_office,
            "原生 Excel COM 显示色扫描工作进程不可用"
        );
        ensure!(
            !scan_areas.is_empty(),
            "模板没有配置条件格式扫描区域；无法按实际显示色审核"
        );
        let input_path = input_path.as_ref();
        ensure!(input_path.is_file(), "条件格式显示色扫描输入不存在");

        // If the user selected LibreOffice for formula calculation, the same
        // Audit worker still uses a separate private Excel COM application for
        // Windows DisplayFormat observation, then returns to the LO choice.
        let selected_engine = self.engine_request.clone();
        if selected_engine == "lo" {
            self.engine_request = "auto".to_owned();
        }
        let attempt = (|| -> Result<crate::conditional_format::ConditionalDisplayScan> {
            self.ensure_started()?;
            let request = serde_json::json!({
                "type": "scan_conditional_display_colors",
                "input": sidecar_io_path(input_path).display().to_string(),
                "scan_areas": scan_areas,
            });
            let stdin = self
                .stdin
                .as_mut()
                .context("原生 COM 显示色扫描输入不可用")?;
            serde_json::to_writer(&mut *stdin, &request)?;
            stdin.write_all(b"\n")?;
            stdin.flush()?;
            let response = self.receive_response()?;
            let report: Value =
                serde_json::from_str(&response).context("解析原生 COM 显示色扫描结果")?;
            ensure!(
                report.get("status").and_then(Value::as_str) == Some("display_colors_scanned"),
                "Excel COM 显示色扫描失败：{}",
                report
                    .get("error")
                    .and_then(Value::as_str)
                    .unwrap_or("worker 未报告扫描完成")
            );
            serde_json::from_value(
                report
                    .get("scan")
                    .cloned()
                    .context("Excel COM 显示色扫描响应缺少 scan")?,
            )
            .context("解析 Excel COM 显示色扫描数据")
        })();
        self.engine_request = selected_engine;
        attempt
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
        recalculate_s1_formula_copy(Some(&self.sidecar_path), input_path, output_path, "lo")
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
        crate::gui_diagnostics::trace("office_worker_start",serde_json::json!({"engineRequest":self.engine_request,"nativeOffice":self.native_office}));
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
        crate::gui_diagnostics::trace("office_worker_ready",serde_json::json!({"engine":ready["engine"],
            "workerPid":self.child.as_ref().map(std::process::Child::id),"officePid":ready["engine_process_id"],
            "comAvailable":ready["com_available"],"startupTimings":ready["startup_timings"]}));
        self.excel_process_id = ready
            .get("engine_process_id")
            .and_then(Value::as_u64)
            .and_then(|value| u32::try_from(value).ok());
        self.active_engine = ready
            .get("engine")
            .and_then(Value::as_str)
            .map(str::to_owned);
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
        crate::gui_diagnostics::trace("office_worker_ipc_wait",serde_json::json!({"engine":self.engine_request,
            "workerPid":self.child.as_ref().map(std::process::Child::id)}));
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
        self.active_engine = None;
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
        self.active_engine = None;
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
        // 安全边界：只对交给引擎的临时副本做规范化，绝不回写调用方 input；
        // 该副本同时作为重算后的比对基线（见 recalculate_betteroffice 说明）。
        let _ = crate::excel_open_repair::repair_audit_copy_for_excel(&prepared);
        let engine_output = engine_staging.path().join("native-com-result.xlsx");
        let candidate = destination_staging.path().join("recalculated.xlsx");

        let attempt = (|| -> Result<Value> {
            self.ensure_started()?;
            // WPS accepts these real S1 workbooks but can drop a source
            // formula when asked to SaveAs another XLSX. Give WPS a fresh
            // calculation-only copy and save that file in place instead.
            // Excel keeps the existing SaveAs path, and the original Audit
            // copy remains untouched in both cases.
            let save_in_place = self
                .active_engine
                .as_deref()
                .is_some_and(|engine| !engine.eq_ignore_ascii_case("Excel.Application"));
            if save_in_place {
                fs::copy(&prepared, &engine_output).context("准备 WPS 原位保存的独立重算副本")?;
            }
            let engine_input = if save_in_place {
                &engine_output
            } else {
                &prepared
            };
            let request = serde_json::json!({
                "type": "recalculate",
                "input": sidecar_io_path(engine_input).display().to_string(),
                "output": sidecar_io_path(&engine_output).display().to_string(),
                "save_in_place": save_in_place,
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
            ensure!(
                report.get("save_mode").and_then(Value::as_str)
                    == Some(if save_in_place { "in_place" } else { "save_as" }),
                "原生 COM 未按预期保存方式完成重算"
            );
            ensure!(engine_output.is_file(), "原生 COM 没有生成重算工作簿");
            // 基准必须是已规范化的引擎输入副本：用未规范化的调用方 input
            // 比对，会把 shared ref 归一误判成「重算改变公式」。
            let normalized = crate::native_recalculation::normalize_engine_output(
                &prepared,
                &engine_output,
                engine_staging.path(),
            )?;
            fs::copy(normalized, &candidate).context("暂存原生 COM 重算结果")?;
            // 比对基准用已规范化的引擎输入副本，避免把规范化误判为重算改动。
            validate_recalculated_copy(&prepared, &candidate)?;
            Ok(report)
        })();
        let report = match attempt {
            Ok(report) => report,
            Err(error) => {
                let should_retry_with_wps = self.engine_request == "auto"
                    && self.native_office
                    && self.active_engine.as_deref() == Some("Excel.Application")
                    && format!("{error:#}").contains("Office COM Workbooks.Open 调用失败");
                self.stop_process();
                let fallback = if should_retry_with_wps {
                    self.fallback_to_wps(&input, &output, error)
                } else {
                    self.fallback_to_libreoffice(&input, &output, error)
                };
                return fallback.map(|mut report| {
                    report.elapsed_ms = started.elapsed().as_millis();
                    report
                });
            }
        };
        // 比对基准用已规范化的引擎输入副本（不是未修复的调用方 input）。
        finish_recalculation(&prepared, &output, &candidate, report, started)
    }

    fn fallback_to_wps(
        &mut self,
        input_path: &Path,
        output_path: &Path,
        excel_error: anyhow::Error,
    ) -> Result<S1RecalculationReport> {
        let excel_error = format!("{excel_error:#}");
        self.engine_request = "wps".to_owned();
        match self.recalculate_native(input_path, output_path) {
            Ok(report) => {
                // This WPS fallback is for the current workbook only. Leaving
                // the session pinned to WPS would silently route every later
                // workbook through an engine that can rewrite shared formulas
                // in some templates. Close the WPS worker, then let the next
                // workbook start a fresh `auto` (Excel-first) session.
                self.resume_auto_after_wps_fallback();
                Ok(report)
            }
            Err(wps_error) => {
                self.stop_process();
                self.engine_request = "auto".to_owned();
                let combined_error = anyhow::anyhow!(
                    "Excel 打开失败（{excel_error}）；WPS 表格回退也失败（{wps_error:#}）"
                );
                self.fallback_to_libreoffice(input_path, output_path, combined_error)
            }
        }
    }

    fn resume_auto_after_wps_fallback(&mut self) {
        self.stop_process();
        self.engine_request = "auto".to_owned();
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
        // 同一口径：比对基准必须是已规范化的引擎输入副本。
        finish_recalculation(&prepared, &output, &candidate, report, started)
    }
}

impl Drop for S1RecalculationSession {
    fn drop(&mut self) {
        crate::gui_diagnostics::trace("office_session_shutdown",serde_json::json!({"engine":self.active_engine,
            "workerPid":self.child.as_ref().map(std::process::Child::id),"officePid":self.excel_process_id}));
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
/// 为 Some(报告)，失败项为 None（GUI 将保留待人工重算副本）。
pub fn recalculate_s1_formula_copy_batch(
    sidecar_path: Option<&Path>,
    pairs: &[(PathBuf, PathBuf)],
    engine_request: &str,
) -> Result<Vec<Option<S1RecalculationReport>>> {
    ensure!(
        engine_request == "auto" || engine_request == "lo",
        "S1 批次重算仅支持 auto/lo（LibreOffice）：{engine_request}"
    );
    ensure!(!pairs.is_empty(), "S1 批次重算没有输入工作簿");
    // betteroffice 优先（Windows / Linux 同一路由）：逐本先试 betteroffice，
    // 只有被拒绝的才进入 LibreOffice 批处理——批量语义（一次 soffice 调用）
    // 保持不变，本函数只负责收窄待转换集合。
    let bo_config = crate::bo_calc::BoCalcConfig::from_app();
    let mut betteroffice_outcomes: Vec<Option<S1RecalculationReport>> =
        (0..pairs.len()).map(|_| None).collect();
    let mut betteroffice_reasons: Vec<Option<String>> = (0..pairs.len()).map(|_| None).collect();
    let mut pending_pairs: Vec<(usize, PathBuf, PathBuf)> = Vec::new();
    if bo_config.enabled && engine_request == "auto" {
        // betteroffice 是进程内纯 Rust：按本并行（每本独立 Workbook，无共享
        // 可变状态）。大簿是瓶颈，并发放在外层；线程数取可用核数与本数较小者。
        let results: Vec<std::result::Result<S1RecalculationReport, String>> = pairs
            .par_iter()
            .map(|(input, output)| {
                recalculate_with_betteroffice(input, output, &bo_config)
                    .map_err(|error| format!("{error:#}"))
            })
            .collect();
        for (index, result) in results.into_iter().enumerate() {
            match result {
                Ok(report) => betteroffice_outcomes[index] = Some(report),
                Err(reason) => {
                    if !bo_config.fallback_available() {
                        return Err(anyhow::anyhow!(
                            "betteroffice 校验区域重算不可采用且未配置可用回退：{reason}"
                        ));
                    }
                    betteroffice_reasons[index] = Some(format!("betteroffice 回退（{reason}）"));
                    let (input, output) = &pairs[index];
                    pending_pairs.push((index, input.clone(), output.clone()));
                }
            }
        }
        if pending_pairs.is_empty() {
            // 全部由 betteroffice 完成：无需启动 LibreOffice。
            return Ok(betteroffice_outcomes);
        }
    } else {
        for (index, (input, output)) in pairs.iter().enumerate() {
            pending_pairs.push((index, input.clone(), output.clone()));
        }
    }
    // 仅剩被 betteroffice 拒绝的工作簿需要外部引擎。
    let refusal_summary: Vec<String> = betteroffice_reasons.iter().flatten().cloned().collect();
    let sidecar_missing = if refusal_summary.is_empty() {
        "betteroffice 校验区域重算不可用，且未配置公式重算 sidecar（Excel COM / LibreOffice）"
            .to_owned()
    } else {
        format!(
            "betteroffice 校验区域重算不可用（{}），且未配置公式重算 sidecar（Excel COM / LibreOffice）",
            refusal_summary.join("；")
        )
    };
    let sidecar_path = sidecar_path.context(sidecar_missing)?;
    let sidecar_path = fs::canonicalize(sidecar_path)
        .with_context(|| format!("读取公式重算 sidecar：{}", sidecar_path.display()))?;
    ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");
    for (input, output) in pairs {
        ensure!(
            input.is_file(),
            "S1 批次重算输入不是文件：{}",
            input.display()
        );
        ensure!(is_xlsx(input), "S1 批次重算输入必须为 .xlsx");
        ensure!(is_xlsx(output), "S1 批次重算输出必须为 .xlsx");
        ensure!(
            !output.exists(),
            "拒绝覆盖 S1 批次重算输出：{}",
            output.display()
        );
    }
    let started = Instant::now();
    // 安全边界 + 口径一致：为每个待处理工作簿建一份**规范化基线副本**，
    // sidecar 读基线、比对也用基线，调用方文件保持原样。
    let batch_staging = StagingDirectory::create(&std::env::temp_dir())?;
    let mut pending: Vec<(PathBuf, PathBuf)> = Vec::with_capacity(pending_pairs.len());
    let mut batch_baselines: Vec<PathBuf> = Vec::with_capacity(pending_pairs.len());
    for (sequence, (_, input, _)) in pending_pairs.iter().enumerate() {
        let baseline = batch_staging
            .path()
            .join(format!("baseline-{sequence}.xlsx"));
        std::fs::copy(input, &baseline)
            .with_context(|| format!("为批量重算准备基线副本：{}", input.display()))?;
        let _ = crate::excel_open_repair::repair_audit_copy_for_excel(&baseline);
        batch_baselines.push(baseline.clone());
        let (_, _, output) = &pending_pairs[sequence];
        pending.push((baseline, output.clone()));
    }
    let pairs_file = unique_pairs_file(&pending)?;
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
    let report: Value =
        serde_json::from_slice(&command.stdout).context("解析批次公式重算 sidecar 的 JSON 结果")?;
    let books = report
        .get("books")
        .and_then(Value::as_array)
        .context("批次公式重算结果缺少 books 数组")?;
    ensure!(
        report
            .get("soffice_conversion_invocations")
            .and_then(Value::as_u64)
            == Some(1),
        "LibreOffice 批次报告未确认单次转换引擎调用"
    );
    ensure!(
        report
            .get("profile_processes_remaining")
            .and_then(Value::as_u64)
            == Some(0),
        "LibreOffice 批次结束后仍有进程引用专用 profile"
    );
    anyhow::ensure!(
        books.len() == pending.len(),
        "批次公式重算返回的本数不一致：{} / {}",
        books.len(),
        pending.len()
    );
    // A single soffice invocation handles the whole batch. Allocate its wall
    // time evenly for the existing per-workbook performance ledger so summing
    // workbook entries does not multiply the same batch duration by N.
    let elapsed_per_book_ms = started.elapsed().as_millis() / pending.len() as u128;
    let mut outcomes = betteroffice_outcomes;
    for (pending_index, book) in books.iter().enumerate() {
        let original_index = pending_pairs[pending_index].0;
        let output = &pending_pairs[pending_index].2;
        if book.get("status").and_then(Value::as_str) != Some("recalculated") {
            continue;
        }
        let engine = book
            .get("engine")
            .and_then(Value::as_str)
            .context("批次重算结果缺少引擎名")?
            .to_owned();
        validate_s1_engine(&engine)?;
        // 与「已规范化基线」比对（sidecar 读的就是同一份基线），
        // 避免把 shared ref 归一误判成「重算改变公式」。
        let warnings = validate_recalculated_copy(&batch_baselines[pending_index], output)?;
        let engine = match betteroffice_reasons[original_index].as_ref() {
            Some(reason) => format!("{engine}；{reason}"),
            None => engine,
        };
        outcomes[original_index] = Some(S1RecalculationReport {
            engine,
            input: pending_pairs[pending_index].1.clone(),
            output: output.clone(),
            elapsed_ms: elapsed_per_book_ms,
            warnings,
        });
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
    sidecar_path: Option<&Path>,
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    engine_request: &str,
) -> Result<S1RecalculationReport> {
    ensure!(
        S1_CALC_ENGINES.contains(&engine_request),
        "S1 公式重算引擎必须是 auto、excel、wps 或 lo：{engine_request}"
    );
    // Excel COM 打开修复：与 GUI 主路径（execute_s1_formula_copy_*）同一口径。
    // 本函数也被 CLI/诊断入口直接调用，若不在此修复，shared formula 的 ref
    // 错位会让重算守卫误判「重算改变公式」。
    // betteroffice（纯 Rust）优先——Windows 与 Linux 同一路由：
    // betteroffice → 本函数既有 sidecar（Windows COM / Linux LibreOffice）。
    let bo_config = crate::bo_calc::BoCalcConfig::from_app();
    let mut bo_fallback_reason: Option<String> = None;
    if bo_config.enabled && engine_request == "auto" {
        match recalculate_with_betteroffice(input_path.as_ref(), output_path.as_ref(), &bo_config) {
            Ok(report) => return Ok(report),
            Err(error) => {
                let reason = format!("{error:#}");
                if !bo_config.fallback_available() {
                    bail!("betteroffice 校验区域重算不可采用且未配置可用回退：{reason}");
                }
                bo_fallback_reason = Some(format!("betteroffice 回退（{reason}）"));
            }
        }
    }
    // 至此 betteroffice 不可用；没有 sidecar 就没有任何可用引擎。
    // 把 betteroffice 的拒绝原因带进最终错误，否则用户只看到「缺 sidecar」，
    // 完全不知道真正原因（预算耗尽 / #NAME? / 循环引用 / 打不开）。
    let sidecar_missing = match &bo_fallback_reason {
        Some(reason) => format!(
            "betteroffice 校验区域重算不可用（{reason}），且未配置公式重算 sidecar（Excel COM / LibreOffice）"
        ),
        None => "betteroffice 校验区域重算不可用，且未配置公式重算 sidecar（Excel COM / LibreOffice）"
            .to_owned(),
    };
    let sidecar_path = sidecar_path.context(sidecar_missing)?;
    let sidecar_path = fs::canonicalize(sidecar_path)
        .with_context(|| format!("读取公式重算 sidecar：{}", sidecar_path.display()))?;
    ensure!(sidecar_path.is_file(), "公式重算 sidecar 不是文件");
    let started = Instant::now();
    let (input_path, output_path, staging) =
        prepare_recalculation_paths(input_path.as_ref(), output_path.as_ref())?;
    let candidate = staging.path().join("recalculated.xlsx");
    // 安全边界：sidecar 引擎读的是临时副本，规范化只改它。
    // 规范化基线单独留一份，供重算后比对（否则 ref 归一会被误判成改公式）。
    let baseline = staging.path().join("recalc-baseline.xlsx");
    std::fs::copy(&input_path, &baseline)
        .with_context(|| format!("为 sidecar 重算准备基线副本：{}", baseline.display()))?;
    let _ = crate::excel_open_repair::repair_audit_copy_for_excel(&baseline);
    std::fs::copy(&baseline, &candidate)
        .with_context(|| format!("为 sidecar 重算准备工作副本：{}", candidate.display()))?;
    let sidecar_input = sidecar_io_path(&candidate);
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
    let mut report = finish_recalculation(&baseline, &output_path, &candidate, report, started)?;
    if let Some(reason) = bo_fallback_reason {
        report.engine = format!("{}；{}", report.engine, reason);
    }
    Ok(report)
}

/// betteroffice 独立重算入口（Windows / Linux 共用同一实现）。
fn recalculate_with_betteroffice(
    input_path: &Path,
    output_path: &Path,
    config: &crate::bo_calc::BoCalcConfig,
) -> Result<S1RecalculationReport> {
    let started = Instant::now();
    let (input, output, destination_staging) =
        prepare_recalculation_paths(input_path, output_path)?;
    // 安全边界：只改临时工作副本，绝不回写调用方文件；并保留已规范化的
    // 基线用于重算后比对（见 recalculate_betteroffice 同处说明）。
    let baseline = destination_staging.path().join("recalc-baseline.xlsx");
    std::fs::copy(&input, &baseline)
        .with_context(|| format!("为 betteroffice 重算准备基线副本：{}", baseline.display()))?;
    let _ = crate::excel_open_repair::repair_audit_copy_for_excel(&baseline);
    let candidate = destination_staging.path().join("recalculated.xlsx");
    std::fs::copy(&baseline, &candidate)
        .with_context(|| format!("为 betteroffice 重算准备工作副本：{}", candidate.display()))?;
    let outcome = crate::bo_calc::recalculate(&candidate, config)?;
    let report = serde_json::json!({
        "status": "recalculated",
        "engine": outcome.application,
    });
    // 比对基准必须是「已规范化的基线」，否则 shared ref 归一会误判成改公式。
    let mut report = finish_recalculation(&baseline, &output, &candidate, report, started)?;
    report.elapsed_ms = started.elapsed().as_millis();
    Ok(report)
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
    let warnings = validate_recalculated_copy(input_path, candidate)?;
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
        warnings,
    })
}

/// Independently validate formula preservation, cached-value node completeness
/// and typed cached-result readability before publishing the recalculated
/// candidate. This proves the calculation result is consumable, not that it
/// matches a business-specific expected value.
fn validate_recalculated_copy(input: &Path, output: &Path) -> Result<Vec<String>> {
    let source = excel_reader::read_workbook_with_unknown_error_tokens_as_text(input)
        .with_context(|| format!("重开重算前的 S1 公式副本：{}", input.display()))?;
    let result = excel_reader::read_workbook_with_unknown_error_tokens_as_text(output)
        .with_context(|| format!("重开重算后的 S1 公式副本：{}", output.display()))?;
    let source = source.workbook;
    let result = result.workbook;

    // 硬失败只保留「工作表数量 / 名称 / 顺序」。
    //
    // 依据（2026-11 用户决定，进一步对齐 Rust-Tools 口径）：本流程只从
    // **校验区域**提取公式结果关键字，不追求整簿逐格一致。引擎重存会合法
    // 改写公式坐标（shared formula 展开）、规范化文本、丢弃非校验区域的
    // 冗余公式——这些都不影响提取结果，硬失败只会制造误报。
    //
    // 唯一有业务风险的场景是「校验区域某格的公式缓存丢失」：提取会把该格
    // 当成空值静默跳过，漏掉一条审核结论。这一条不再在整簿层拦截，改为在
    // 提取阶段对**实际要读的规则单元格**做定向检查并标注"部分完成"
    // （见 s1_formula_audit::extract_s1_formula_issues）。
    ensure!(
        source.sheets.len() == result.sheets.len(),
        "S1 重算改变了工作表数量：{} -> {}",
        source.sheets.len(),
        result.sheets.len()
    );
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        ensure!(
            before.name == after.name,
            "S1 重算改变了工作表名称或顺序：{:?} -> {:?}",
            before.name,
            after.name
        );
    }

    // 以下全部降级为提示：公式格坐标变化、公式数量、缓存节点完整性、
    // 结果可读性。如实写进详细运行日志，供事后核对，但不阻断发布。
    let mut warnings = formula_text_drift_warnings(&source, &result);
    warnings.extend(structure_drift_warnings(&source, &result, output));
    Ok(warnings)
}

/// 结构/缓存/可读性差异（软提示）。
///
/// 这些差异不影响「校验区域关键字提取」的正确性，故只作留痕；真正会造成
/// 漏报的"校验区域公式缓存缺失"由提取阶段定向检查并标注部分完成。
fn structure_drift_warnings(
    source: &excel_reader::WorkbookDump,
    result: &excel_reader::WorkbookDump,
    output: &Path,
) -> Vec<String> {
    const MAX_SAMPLES: usize = 5;
    let mut warnings = Vec::new();

    // 1) 公式格坐标集合（按表）。
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        let before_cells = before
            .formulas
            .keys()
            .collect::<std::collections::BTreeSet<_>>();
        let after_cells = after
            .formulas
            .keys()
            .collect::<std::collections::BTreeSet<_>>();
        if before_cells == after_cells {
            continue;
        }
        let added = after_cells.difference(&before_cells).count();
        let removed = before_cells.difference(&after_cells).count();
        warnings.push(format!(
            "S1-F01 提示：工作表 {:?} 的公式格坐标有变化（新增 {added} 个、消失 {removed} 个）；             校验区域结果提取不受影响",
            before.name
        ));
        for cell in before_cells.difference(&after_cells).take(MAX_SAMPLES) {
            warnings.push(format!("  消失：{}!{cell}", before.name));
        }
    }

    // 2) 整簿公式数量与缓存节点。
    let formula_count = source
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    if formula_count == 0 {
        warnings.push("S1-F01 提示：公式副本中没有可重算的公式".to_owned());
    }
    match excel_reader::inspect_formula_cache_nodes(output) {
        Ok(inventory) => {
            if inventory.formula_cell_count != formula_count
                || inventory.cached_formula_cell_count != formula_count
                || !inventory.missing_cache_cells.is_empty()
            {
                warnings.push(format!(
                    "S1-F01 提示：重算副本公式 {} 个、OOXML 公式单元格 {} 个、带缓存 {} 个、                     缺缓存 {} 个；校验区域如有缺失会在提取阶段标注部分完成",
                    formula_count,
                    inventory.formula_cell_count,
                    inventory.cached_formula_cell_count,
                    inventory.missing_cache_cells.len()
                ));
                for location in inventory.missing_cache_cells.iter().take(MAX_SAMPLES) {
                    warnings.push(format!(
                        "  缺缓存：{}!{}",
                        location.sheet_name, location.cell
                    ));
                }
            }
        }
        Err(error) => warnings.push(format!("S1-F01 提示：重算副本缓存节点检查失败：{error:#}")),
    }

    // 3) 结果可读性。
    let coverage = excel_reader::formula_result_value_coverage(result);
    if coverage.formula_cell_count != formula_count
        || coverage.typed_result_cell_count != formula_count
        || !coverage.missing_result_cells.is_empty()
    {
        warnings.push(format!(
            "S1-F01 提示：公式 {} 个，可读结果 {} 个；校验区域如有缺失会在提取阶段标注部分完成",
            coverage.formula_cell_count, coverage.typed_result_cell_count
        ));
        for location in coverage.missing_result_cells.iter().take(MAX_SAMPLES) {
            warnings.push(format!(
                "  无结果：{}!{}",
                location.sheet_name, location.cell
            ));
        }
    }
    warnings
}

/// 逐格公式文本差异（软警告）。
///
/// 只统计"引擎等价规范化之后仍不等价"的单元格，最多保留若干条样例，
/// 避免日志被单簿上千处改写淹没。返回空列表表示文本完全一致。
fn formula_text_drift_warnings(
    source: &excel_reader::WorkbookDump,
    result: &excel_reader::WorkbookDump,
) -> Vec<String> {
    const MAX_SAMPLES: usize = 5;
    let sheet_names = source
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let mut total = 0usize;
    let mut samples: Vec<String> = Vec::new();
    for (before, after) in source.sheets.iter().zip(&result.sheets) {
        for (cell, formula) in &before.formulas {
            let Some(result_formula) = after.formulas.get(cell) else {
                continue;
            };
            if excel_reader::formula_text_equivalent_after_recalculation(
                formula,
                result_formula,
                &sheet_names,
            ) {
                continue;
            }
            total += 1;
            if samples.len() < MAX_SAMPLES {
                samples.push(format!(
                    "{}!{cell}: {:?} -> {:?}",
                    before.name, formula, result_formula
                ));
            }
        }
    }
    if total == 0 {
        return Vec::new();
    }
    let mut warnings = vec![format!(
        "S1-F01 提示：重算后有 {total} 个公式格的文本被引擎改写（结构未变，计算结果不受影响；         引擎重存会合法规范化 shared formula/_xlfn/表名引号等书写形式）"
    )];
    let shown = samples.len();
    warnings.extend(samples);
    if total > shown {
        warnings.push(format!("…其余 {} 处略", total - shown));
    }
    warnings
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

/// betteroffice（纯 Rust 计算引擎）的报告名。vendor 版计算引擎作为首选
/// 路由后，实际引擎不再只有 Office COM / LibreOffice 三选一。
const BETTEROFFICE_ENGINE: &str = "betteroffice (pure Rust)";

fn is_betteroffice_engine(engine: &str) -> bool {
    engine.split('；').next().map(str::trim) == Some(BETTEROFFICE_ENGINE)
}

fn validate_s1_engine(engine: &str) -> Result<()> {
    if is_betteroffice_engine(engine) {
        // 回退实现串（"…；betteroffice 回退（原因）"）也归此分支：
        // 结果来自被接受的平台引擎，回退原因仅作留痕。
        return Ok(());
    }
    #[cfg(windows)]
    ensure!(
        engine == "Excel.Application"
            || engine == "ket.Application"
            || engine == "KET.Application"
            || engine == "LibreOffice"
            || engine.starts_with("Excel.Application；")
            || engine.starts_with("ket.Application；")
            || engine.starts_with("KET.Application；")
            || engine.starts_with("LibreOffice；"),
        "S1 Windows 公式重算实际引擎不受支持：“{engine}”，仅接受 betteroffice、Microsoft Excel COM、WPS 表格 COM 或 LibreOffice"
    );
    #[cfg(not(windows))]
    ensure!(
        engine == "LibreOffice" || engine.starts_with("LibreOffice；"),
        "S1 Linux 公式重算要求 LibreOffice Calc 或 betteroffice；sidecar 实际使用“{engine}”，拒绝将其结果作为审核输入"
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

    /// 规范化只允许改临时工作副本：重算不得回写调用方传入的输入文件。
    /// 用真实 ！40 模板（含 needs_repair 命中的 shared ref）构造调用。
    #[test]
    fn recalculation_never_rewrites_the_callers_input_file() {
        let Some(template) = real_template_path("！40.金融基础数据采集（委托贷款）202608.xlsx")
        else {
            return; // 测试数据不在时跳过（CI/裁剪仓库）
        };
        let state = std::env::temp_dir().join(format!(
            "audit_no_mutation_{}",
            std::time::SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock")
                .as_nanos()
        ));
        std::fs::create_dir_all(&state).unwrap();
        let input = state.join("caller-input.xlsx");
        let output = state.join("recalculated.xlsx");
        std::fs::copy(&template, &input).unwrap();
        let before = std::fs::read(&input).unwrap();

        // betteroffice 可用时直接成功；不可用时回退也会失败（无 sidecar），
        // 两种情形都只关心「input 是否被改写」，因此忽略结果。
        let _ = recalculate_s1_formula_copy(None, &input, &output, "auto");

        let after = std::fs::read(&input).unwrap();
        assert_eq!(
            before, after,
            "调用方输入文件被规范化回写（必须只改临时工作副本）"
        );
        let _ = std::fs::remove_dir_all(&state);
    }

    fn real_template_path(name: &str) -> Option<std::path::PathBuf> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).parent()?;
        let path = root
            .join("测试数据/S1F1_新旧对比_20261105_231620/business_templates")
            .join(name);
        path.is_file().then_some(path)
    }

    /// 公式文本差异只产生警告，不再阻断发布（2026-11 用户决定，对齐 Rust-Tools
    /// 口径）：引擎重存会合法规范化公式书写形式，逐字比对误报率高。结构性事故
    /// （丢表、丢公式、公式格坐标变化）仍由 validate_recalculated_copy 硬失败。
    #[test]
    fn formula_text_drift_becomes_warnings_not_errors() {
        use excel_reader::{SheetDump, WorkbookDump};

        let sheet = |name: &str, formulas: &[(&str, &str)]| SheetDump {
            name: name.to_owned(),
            values: Default::default(),
            formulas: formulas
                .iter()
                .map(|(cell, formula)| ((*cell).to_owned(), (*formula).to_owned()))
                .collect(),
            formula_status: "ok".to_owned(),
            merges: Vec::new(),
            merge_status: "ok".to_owned(),
            comments: Default::default(),
            comments_status: "ok".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "ok".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "ok".to_owned(),
        };
        let workbook = |sheets: Vec<SheetDump>| WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets,
        };

        // 文本被改写（结构不变）：应产生警告而非空列表，也不 panic。
        let source = workbook(vec![sheet("Sheet1", &[("A1", "=SUM(B1:B2)")])]);
        let rewritten = workbook(vec![sheet("Sheet1", &[("A1", "=SUM(Sheet1!B1:B2)")])]);
        let warnings = formula_text_drift_warnings(&source, &rewritten);
        assert_eq!(
            warnings.len(),
            2,
            "应有一条汇总 + 一条样例，实际 {warnings:?}"
        );
        assert!(warnings[0].contains("1 个公式格"), "{:?}", warnings[0]);
        assert!(warnings[1].contains("Sheet1!A1"), "{:?}", warnings[1]);

        // 文本完全一致：无警告。
        let same = formula_text_drift_warnings(&source, &source);
        assert!(same.is_empty(), "文本一致时不应产生警告：{same:?}");

        // 等价规范化（含全角括号等强制引号字符的表名补引号）不算漂移。
        let named = workbook(vec![sheet(
            "汇总（一）",
            &[("A1", "=SUM(汇总（一）!B1:B2)")],
        )]);
        let quoted = workbook(vec![sheet(
            "汇总（一）",
            &[("A1", "=SUM('汇总（一）'!B1:B2)")],
        )]);
        assert!(
            formula_text_drift_warnings(&named, &quoted).is_empty(),
            "表名补引号属等价规范化，不应记为漂移"
        );
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
    fn successful_wps_fallback_resets_session_to_auto_for_the_next_workbook() -> Result<()> {
        let root = TestDirectory::create()?;
        let executable = root.join("sidecar.exe");
        fs::write(&executable, b"test executable")?;
        let mut session = S1RecalculationSession::new(&executable, "auto")?;
        session.engine_request = "wps".to_owned();
        session.active_engine = Some("ket.Application".to_owned());

        session.resume_auto_after_wps_fallback();

        assert_eq!(session.engine_request, "auto");
        assert!(session.child.is_none());
        assert!(session.active_engine.is_none());
        Ok(())
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

        let error = recalculate_s1_formula_copy(Some(&sidecar), &input, &output, "auto")
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

        let recalculation = recalculate_s1_formula_copy(Some(&sidecar), &input, &output, "auto");
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
