//! LibreOffice isolated-profile XLSX batch implementation, retained from the former helper.
use anyhow::{bail, ensure, Context, Result};
use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    time::{Duration, Instant},
};
const PROFILE_FORK_SETTLE: Duration = Duration::from_secs(2);
const POLL_INTERVAL: Duration = Duration::from_millis(100);
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BatchTarget {
    Xlsx,
}

impl BatchTarget {
    pub fn filter(self) -> &'static str {
        match self {
            BatchTarget::Xlsx => "xlsx:Calc MS Excel 2007 XML",
        }
    }

    pub fn staging_extension(self) -> &'static str {
        match self {
            BatchTarget::Xlsx => "xlsx",
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
        fs::hard_link(input, &link)
            .with_context(|| format!("固定批次输入 {} → {}", input.display(), link.display()))?;
        expected.push((out_dir.join(&name), index));
    }
    let started_at = Instant::now();
    let profile_argument = super::libreoffice_user_installation_arg(&work_dir.join("profile"))?;
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
        .arg(&profile_argument)
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
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x0800_0000);
    }
    let convert_log = work_dir.join("batch-convert.log");
    command.stdout(
        fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&convert_log)?,
    );
    command.stderr(
        fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&convert_log)?,
    );
    let status = run_with_timeout(&mut command, timeout, "批次 LibreOffice 转换")?;
    // The soffice wrapper can exit before oosplash finishes forking
    // soffice.bin. Let the owned profile settle, then wait for its processes
    // to exit. If they outlive the batch, terminate only commands carrying
    // this unique profile argument and confirm cleanup before returning.
    drain_profile_after_batch(
        &profile_argument,
        timeout.saturating_sub(started_at.elapsed()),
    )
    .context("清理本批次拥有的 LibreOffice profile 进程")?;
    ensure!(
        status.success(),
        "批次 LibreOffice 转换失败（exit={status}）"
    );
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
    let elapsed_ms = started_at.elapsed().as_millis() as u64;
    for stage in &mut stages {
        stage["batch_convert_ms"] = json!(elapsed_ms);
        stage["soffice_conversion_invocations"] = json!(1);
        stage["profile_processes_remaining"] = json!(0);
    }
    Ok(("LibreOffice".to_owned(), stages))
}

/// 带超时地运行一个 soffice 命令；超时终止本任务创建的进程组。
fn run_with_timeout(
    command: &mut Command,
    timeout: Duration,
    label: &str,
) -> Result<std::process::ExitStatus> {
    let mut child = command.spawn().with_context(|| format!("启动 {label}"))?;
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
                bail!(
                    "{label} 超过 {} 秒，已终止本任务创建的进程组",
                    timeout.as_secs()
                );
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
fn terminate_process_group(process_id: u32, _force: bool) {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        let _ = Command::new("taskkill.exe")
            .args(["/PID", &process_id.to_string(), "/T", "/F"])
            .creation_flags(0x0800_0000)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status();
    }
}

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
        .is_ok_and(|status| status.code() == Some(1))
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

/// Wait for a batch-owned profile's LO process to exit, then clean up only
/// processes still carrying that unique profile. The grace period accounts
/// for the observed oosplash fork delay; the remaining wait is bounded by the
/// caller's batch timeout.
fn drain_profile_after_batch(profile_argument: &str, timeout: Duration) -> Result<()> {
    let settle = PROFILE_FORK_SETTLE.min(timeout);
    std::thread::sleep(settle);
    let deadline = Instant::now() + timeout.saturating_sub(settle);
    while !no_process_uses_profile(profile_argument) {
        if Instant::now() >= deadline {
            break;
        }
        std::thread::sleep(POLL_INTERVAL);
    }
    if no_process_uses_profile(profile_argument) {
        return Ok(());
    }

    kill_profile_processes(profile_argument);
    let cleanup_deadline = Instant::now() + Duration::from_secs(5);
    while !no_process_uses_profile(profile_argument) && Instant::now() < cleanup_deadline {
        std::thread::sleep(POLL_INTERVAL);
    }
    ensure!(
        no_process_uses_profile(profile_argument),
        "LibreOffice profile 进程在定向终止后仍未退出"
    );
    Ok(())
}
