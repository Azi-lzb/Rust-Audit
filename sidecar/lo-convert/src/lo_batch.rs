//! LibreOffice 会话复用（S1-F01 Linux 公式重算，rules/13 §4.5）。
//!
//! 两种已实测的复用形态：
//! - [`LoWarmProfileSession`]：一个会话只做一次 profile 初始化，随后逐本
//!   `--convert-to`（顺序、同步），供持久会话协议（GUI 逐本请求）使用，
//!   省去每本全新 profile 的首启初始化成本。
//! - [`run_batch_convert`]：一次 `soffice --convert-to` 同时携带批次的多个
//!   输入文件，单实例顺序转换全部工作簿，供上游已知整批清单时使用。
//!
//! 为什么不是 Rust 直接讲 UNO：LibreOffice 的远程 UNO（URP）协议没有可
//! 维护的 Rust 绑定，手写桥接协议不可验收；链接 C++ UNO SDK 则要求构建期
//! LibreOffice 开发头文件与运行期版本锁定的 UNO 库，不适配麒麟/UOS 离线
//! 目标。进程内的 StarBasic 批次驱动方案也已实测放弃：`macro:///` 在全新
//! profile 上的行为不稳定（首启会重写 Standard 库索引；处理任务行时宏静
//! 默终止、无任何诊断输出），无法满足“失败必须有明确诊断”的要求。
//!
//! 进程边界：所有 soffice 以独立进程组启动；超时只终止本任务创建的进程
//! 组，绝不按 `soffice` 进程名清理用户实例。已实测的竞态：初始化的
//! wrapper 进程（/usr/bin/soffice → oosplash）会先于 soffice.bin 退出，
//! 且 oosplash fork soffice.bin 的窗口在 wrapper 退出之后（0.3–1.4s），
//! 因此初始化后必须沉降并确认引用本批次 profile 的进程已排空。

use anyhow::{bail, ensure, Context, Result};
use serde_json::{json, Value};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{Duration, Instant};

/// 初始化 wrapper 退出后的沉降时间：覆盖 oosplash fork soffice.bin 的观测
/// 窗口（0.3–1.4s），保证排空检查能看到最终进程形态。
const PROFILE_FORK_SETTLE: Duration = Duration::from_secs(2);

/// 排空检查的超时上限。
const PROFILE_DRAIN_TIMEOUT: Duration = Duration::from_secs(30);

/// 转换进程的轮询间隔。
const POLL_INTERVAL: Duration = Duration::from_millis(100);

/// 暖 profile 会话：profile 只初始化一次，之后逐本 `--convert-to` 顺序
/// 复用。每本转换使用独立 soffice 进程（转换完自动退出），互不影响。
pub struct LoWarmProfileSession {
    work_dir: PathBuf,
    profile_argument: String,
    initialized: bool,
    _work_guard: crate::TempDirGuard,
}

impl LoWarmProfileSession {
    pub fn create() -> Result<Self> {
        let work_dir = crate::unique_temp_dir(&std::env::temp_dir())
            .context("创建 LibreOffice 会话工作目录")?;
        let profile_dir = work_dir.join("profile");
        fs::create_dir_all(&profile_dir).context("创建 LibreOffice profile 目录")?;
        let profile_argument = crate::libreoffice_user_installation_arg(&profile_dir)?;
        let _work_guard = crate::TempDirGuard(work_dir.clone());
        Ok(Self {
            work_dir,
            profile_argument,
            initialized: false,
            _work_guard,
        })
    }

    /// 会话工作目录（诊断信息所在）。
    pub fn work_dir(&self) -> &Path {
        self._work_guard.0.as_path()
    }

    /// 处理一本工作簿：`input` 为已清除公式缓存并标记重算的副本；转换以
    /// `--outdir` 落盘，调用方保证 output 与 input 同名同目录。返回
    /// (引擎名, 阶段耗时 JSON)。
    pub fn recalculate(
        &mut self,
        input: &Path,
        output: &Path,
        timeout: Duration,
    ) -> Result<(String, Value)> {
        self.ensure_initialized()?;
        let started_at = Instant::now();
        let out_dir = output
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let expected = out_dir.join(input.file_name().context("会话转换输入缺少文件名")?);
        if expected.exists() {
            // convert-to 拒绝覆盖已存在的输出；上一本残留不应污染本本。
            fs::remove_file(&expected)
                .with_context(|| format!("清理上一本会话输出失败：{}", expected.display()))?;
        }
        let mut command = Command::new("soffice");
        command
            .args([
                "--headless",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--nofirststartwizard",
                "--norestore",
            ])
            .arg(&self.profile_argument)
            .arg("--convert-to")
            .arg("xlsx:Calc MS Excel 2007 XML")
            .arg("--outdir")
            .arg(out_dir)
            .arg(input);
        {
            #[cfg(unix)]
            {
                use std::os::unix::process::CommandExt;
                // 独立进程组：超时只结束本任务创建的进程及其子进程。
                command.process_group(0);
            }
        }
        let convert_log = self.work_dir.join("session-convert.log");
        command.stdout(
            fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
        );
        command.stderr(
            fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
        );
        let convert_log = self.work_dir.join("session-convert.log");
        command.stdout(
            fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
        );
        command.stderr(
            fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
        );
        let status = run_with_timeout(&mut command, timeout, "会话内 LibreOffice 转换")?;
        ensure!(
            status.success(),
            "会话内 LibreOffice 转换失败（exit={status}）"
        );
        ensure!(
            expected.is_file() && fs::metadata(&expected)?.len() > 0,
            "会话内 LibreOffice 未生成本本的转换结果：{}",
            expected.display()
        );
        let stages = json!({
            "convert_ms": started_at.elapsed().as_millis() as u64,
            "profile_warm": true,
            "lo_profile": self.profile_argument,
        });
        Ok(("LibreOffice".to_owned(), stages))
    }

    fn ensure_initialized(&mut self) -> Result<()> {
        if self.initialized {
            return Ok(());
        }
        let status = Command::new("soffice")
            .args(["--headless", "--terminate_after_init"])
            .arg(&self.profile_argument)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status()
            .context("启动 LibreOffice profile 初始化失败")?;
        ensure!(
            status.success(),
            "LibreOffice profile 初始化失败（exit={status}）"
        );
        // wrapper 先于 soffice.bin 退出，且 fork 发生在 wrapper 退出之后；
        // 沉降后确认引用本会话 profile 的进程已排空，避免下一本通过单实例
        // 通道把 URL 转发给残留实例后立即退出。
        std::thread::sleep(PROFILE_FORK_SETTLE);
        self.drain_profile_processes(PROFILE_DRAIN_TIMEOUT)
            .with_context(|| {
                format!(
                    "LibreOffice profile 初始化实例未排空：{}",
                    self.work_dir.join("profile").display()
                )
            })?;
        self.initialized = true;
        Ok(())
    }

    fn drain_profile_processes(&self, timeout: Duration) -> Result<()> {
        let deadline = Instant::now() + timeout;
        loop {
            if no_process_uses_profile(&self.profile_argument) {
                return Ok(());
            }
            kill_profile_processes(&self.profile_argument);
            if Instant::now() >= deadline {
                bail!("引用本会话 profile 的进程在 {} 秒内未排空", timeout.as_secs());
            }
            std::thread::sleep(Duration::from_millis(100));
        }
    }
}

/// 批次转换的目标类型：决定转换 filter、staging 扩展名与校验配套。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BatchTarget {
    Xlsx,
    Docx,
}

impl BatchTarget {
    pub fn filter(self) -> &'static str {
        match self {
            BatchTarget::Xlsx => "xlsx:Calc MS Excel 2007 XML",
            BatchTarget::Docx => "docx:Office Open XML Text",
        }
    }

    pub fn staging_extension(self) -> &'static str {
        match self {
            BatchTarget::Xlsx => "xlsx",
            BatchTarget::Docx => "docx",
        }
    }

    /// 文本提取（docx 保真校验用）：整个批次各一次，而非每本两次。
    pub fn text_filter(self) -> Option<&'static str> {
        match self {
            BatchTarget::Xlsx => None,
            BatchTarget::Docx => Some("txt:Text"),
        }
    }
}

/// 批次转换：一次 soffice 启动顺序转换多个输入。`pairs` 为（已准备的输入
/// 副本, 期望输出路径）；各输入文件名必须互不相同（同一 `--outdir` 以输入
/// 文件名落盘）。`target` 决定 filter 与 staging 命名。返回每本结果；任一
/// 本缺少输出即返回 Err，由调用方决定逐本回退。
pub fn run_batch_convert(
    pairs: &[(PathBuf, PathBuf)],
    work_dir: &Path,
    target: BatchTarget,
    timeout: Duration,
) -> Result<(String, Vec<Value>)> {
    ensure!(!pairs.is_empty(), "批次转换没有输入");
    let out_dir = work_dir.join("batch-output");
    fs::create_dir_all(&out_dir).context("创建批次输出目录")?;
    // 每个输入以固定名 book-XXXX 硬链接进批次目录，输出同名落盘后可对应
    // 回各本；输出目录必须先清空同名残留。
    let mut expected = Vec::with_capacity(pairs.len());
    for (index, (input, _)) in pairs.iter().enumerate() {
        let name = format!(
            "book-{index:04}.{}",
            input
                .extension()
                .and_then(|value| value.to_str())
                .unwrap_or(target.staging_extension())
        );
        let link = out_dir.join(&name);
        let _ = fs::remove_file(&link);
        fs::hard_link(input, &link).with_context(|| {
            format!("固定批次输入 {} → {}", input.display(), link.display())
        })?;
        expected.push((out_dir.join(&name), index));
    }
    let started_at = Instant::now();
    let mut command = Command::new("soffice");
    command
        .args([
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            "--norestore",
        ])
        .arg(format!(
            "-env:UserInstallation=file://{}",
            work_dir.join("profile").display()
        ))
        .arg("--convert-to")
        .arg(target.filter())
        .arg("--outdir")
        .arg(&out_dir);
    for (input, _) in pairs {
        command.arg(input);
    }
    {
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            command.process_group(0);
        }
    }
    let convert_log = work_dir.join("batch-convert.log");
    command.stdout(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    command.stderr(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    let convert_log = work_dir.join("batch-convert.log");
    command.stdout(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    command.stderr(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    let status = run_with_timeout(&mut command, timeout, "批次 LibreOffice 转换")?;
    ensure!(status.success(), "批次 LibreOffice 转换失败（exit={status}）");
    let mut stages = Vec::with_capacity(pairs.len());
    for (output_path, index) in &expected {
        ensure!(
            output_path.is_file() && fs::metadata(output_path)?.len() > 0,
            "批次转换未生成第 {} 本的结果",
            index + 1
        );
        stages.push(json!({
            "index": index + 1,
            "status": "recalculated",
            "batch_output": output_path.display().to_string(),
        }));
    }
    let _ = started_at;
    Ok(("LibreOffice".to_owned(), stages))
}

/// 批内文本提取（docx 保真校验，移植自 Rust-Tools e2973e6）：整个批次各
/// 一次 `--convert-to txt:Text`，而非每本两次冷启动。`inputs` 为已放入
/// staging 的文档（文件名互不相同），返回与 inputs 等长的纯文本列表。
pub fn run_batch_text_extract(
    inputs: &[PathBuf],
    work_dir: &Path,
    group: &str,
    timeout: Duration,
) -> Result<Vec<String>> {
    ensure!(!inputs.is_empty(), "批次文本提取没有输入");
    let out_dir = work_dir.join(format!("text-{group}"));
    fs::create_dir_all(&out_dir).context("创建批次文本输出目录")?;
    let mut command = Command::new("soffice");
    command
        .args([
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            "--norestore",
        ])
        .arg(format!(
            "-env:UserInstallation=file://{}",
            work_dir.join("profile").display()
        ))
        .arg("--convert-to")
        .arg("txt:Text")
        .arg("--outdir")
        .arg(&out_dir);
    for input in inputs {
        command.arg(input);
    }
    {
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            command.process_group(0);
        }
    }
    let convert_log = work_dir.join(format!("text-extract-{group}.log"));
    command.stdout(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    command.stderr(
        fs::OpenOptions::new().create(true).append(true).open(&convert_log)?,
    );
    let status = run_with_timeout(&mut command, timeout, "批次文本提取")?;
    ensure!(status.success(), "批次文本提取失败（exit={status}）");
    let mut texts = Vec::with_capacity(inputs.len());
    for input in inputs {
        let stem = input.file_stem().context("批次文本提取输入缺少文件名")?;
        let text_path = out_dir.join(stem).with_extension("txt");
        ensure!(
            text_path.is_file(),
            "批次文本提取未生成 {}",
            text_path.display()
        );
        texts.push(fs::read_to_string(&text_path).with_context(|| {
            format!("read extracted Word text {}", text_path.display())
        })?);
    }
    Ok(texts)
}

/// 带超时地运行一个 soffice 命令；超时终止本任务创建的进程组。
fn run_with_timeout(
    command: &mut Command,
    timeout: Duration,
    label: &str,
) -> Result<std::process::ExitStatus> {
    let mut child = command
        .spawn()
        .with_context(|| format!("启动 {label}"))?;
    let started = Instant::now();
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return Ok(status),
            Ok(None) if started.elapsed() < timeout => {
                std::thread::sleep(POLL_INTERVAL);
            }
            Ok(None) => {
                let process_id = child.id();
                terminate_process_group(process_id, false);
                let deadline = Instant::now() + Duration::from_secs(5);
                loop {
                    match child.try_wait() {
                        Ok(Some(_)) => break,
                        Ok(None) if Instant::now() < deadline => {
                            std::thread::sleep(Duration::from_millis(50));
                        }
                        Ok(None) => {
                            terminate_process_group(process_id, true);
                        }
                        Err(_) => break,
                    }
                }
                let _ = child.wait();
                bail!("{label} 超过 {} 秒，已终止本任务创建的进程组", timeout.as_secs());
            }
            Err(error) => return Err(error).with_context(|| format!("等待 {label}")),
        }
    }
}

/// 终止一个进程组；两次尝试（TERM → KILL）。仅用于本任务创建的进程组。
#[cfg(unix)]
fn terminate_process_group(process_id: u32, force: bool) {
    if process_id == 0 || process_id == std::process::id() {
        return;
    }
    let signal = if force { "-KILL" } else { "-TERM" };
    let _ = Command::new("kill")
        .args([signal, &format!("-{process_id}")])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
}

#[cfg(not(unix))]
fn terminate_process_group(_process_id: u32, _force: bool) {}

/// 是否没有进程的命令行仍引用给定 profile 启动参数（pgrep -f 匹配）。
/// 匹配串去掉 `-env:` 前导破折号，避免被 pgrep 当作选项解析。
#[cfg(unix)]
fn no_process_uses_profile(profile_argument: &str) -> bool {
    let pattern = format!(
        "UserInstallation={}",
        profile_argument
            .split_once("UserInstallation=")
            .map(|(_, rest)| rest.to_owned())
            .unwrap_or_else(|| profile_argument.to_owned())
    );
    Command::new("pgrep")
        .args(["-f", "--", &pattern])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .map(|status| !status.success())
        .unwrap_or(true)
}

#[cfg(not(unix))]
fn no_process_uses_profile(_profile_argument: &str) -> bool {
    true
}

/// 终止仍引用给定 profile 的进程（只匹配本任务创建的 profile 参数）。
#[cfg(unix)]
fn kill_profile_processes(profile_argument: &str) {
    let pattern = format!(
        "UserInstallation={}",
        profile_argument
            .split_once("UserInstallation=")
            .map(|(_, rest)| rest.to_owned())
            .unwrap_or_else(|| profile_argument.to_owned())
    );
    let _ = Command::new("pkill")
        .args(["-f", "--", &pattern])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
}

#[cfg(not(unix))]
fn kill_profile_processes(_profile_argument: &str) {}
