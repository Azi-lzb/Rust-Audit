#![cfg(windows)]

use anyhow::{ensure, Context, Result};
use excel_reader::CellValue;
use excel_writer::Workbook;
use serde_json::{json, Value};
use std::{
    fs,
    io::{BufRead, BufReader, Write},
    os::windows::process::CommandExt,
    path::{Path, PathBuf},
    process::{Child, ChildStdin, Command, Stdio},
    sync::mpsc::{self, Receiver},
    thread,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

struct Worker {
    child: Child,
    stdin: ChildStdin,
    responses: Receiver<std::result::Result<String, String>>,
    office_handle: Option<usize>,
}

impl Worker {
    fn start() -> Result<Self> {
        let mut child = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
            .args(["__audit-native-com-session", "excel"])
            .creation_flags(0x0800_0000)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .spawn()?;
        let stdin = child.stdin.take().context("worker stdin")?;
        let stdout = child.stdout.take().context("worker stdout")?;
        let (sender, responses) = mpsc::channel();
        thread::spawn(move || {
            for line in BufReader::new(stdout).lines() {
                let item = line.map_err(|error| error.to_string());
                if sender.send(item).is_err() {
                    break;
                }
            }
        });
        Ok(Self {
            child,
            stdin,
            responses,
            office_handle: None,
        })
    }

    fn response(&mut self) -> Result<Value> {
        let line = self
            .responses
            .recv_timeout(Duration::from_secs(60))
            .context("waiting for native COM worker")?
            .map_err(anyhow::Error::msg)?;
        Ok(serde_json::from_str(&line)?)
    }

    fn request(&mut self, value: Value) -> Result<Value> {
        serde_json::to_writer(&mut self.stdin, &value)?;
        self.stdin.write_all(b"\n")?;
        self.stdin.flush()?;
        self.response()
    }
}

struct EnvironmentRestore(Vec<(String, Option<std::ffi::OsString>)>);

impl EnvironmentRestore {
    fn set(values: &[(&str, &str)]) -> Self {
        let mut previous = Vec::new();
        for (key, value) in values {
            previous.push(((*key).to_owned(), std::env::var_os(key)));
            std::env::set_var(key, value);
        }
        Self(previous)
    }
}

impl Drop for EnvironmentRestore {
    fn drop(&mut self) {
        for (key, value) in self.0.drain(..) {
            if let Some(value) = value {
                std::env::set_var(key, value);
            } else {
                std::env::remove_var(key);
            }
        }
    }
}

fn write_simple_formula_book(path: &Path) -> Result<Vec<u8>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.write_number(0, 0, 2.0)?;
    sheet.write_number(0, 1, 3.0)?;
    sheet.write_formula(0, 2, "=SUM(A1:B1)")?;
    workbook.save(path)?;
    Ok(fs::read(path)?)
}

impl Drop for Worker {
    fn drop(&mut self) {
        if self.office_handle.is_none() {
            return;
        }
        let _ = serde_json::to_writer(&mut self.stdin, &json!({"type":"shutdown"}));
        let _ = self.stdin.write_all(b"\n");
        let _ = self.stdin.flush();
        let _ = self.response();
        let deadline = Instant::now() + Duration::from_secs(10);
        let mut clean_shutdown = false;
        while Instant::now() < deadline {
            match self.child.try_wait() {
                Ok(Some(status)) => {
                    clean_shutdown = status.success();
                    break;
                }
                Ok(None) => thread::sleep(Duration::from_millis(50)),
                Err(_) => break,
            }
        }
        if !clean_shutdown {
            let _ = self.child.kill();
            let _ = self.child.wait();
        }
        if let Some(handle) = self.office_handle.take() {
            let office_exited = clean_shutdown
                && audit_core::native_office_com::wait_verified_process_exit(
                    handle,
                    Duration::from_secs(5),
                );
            audit_core::native_office_com::close_verified_process(handle, !office_exited);
        }
    }
}

impl Worker {
    fn shutdown_and_verify(&mut self) -> Result<()> {
        serde_json::to_writer(&mut self.stdin, &json!({"type":"shutdown"}))?;
        self.stdin.write_all(b"\n")?;
        self.stdin.flush()?;
        let response = self.response()?;
        ensure!(
            response["status"] == "shutdown_complete" && response["excel_quit_succeeded"] == true,
            "Excel COM session did not report a successful Quit: {response}"
        );
        let deadline = Instant::now() + Duration::from_secs(10);
        let status = loop {
            match self.child.try_wait()? {
                Some(status) => break status,
                None if Instant::now() < deadline => thread::sleep(Duration::from_millis(50)),
                None => {
                    let _ = self.child.kill();
                    let _ = self.child.wait();
                    if let Some(handle) = self.office_handle.take() {
                        audit_core::native_office_com::close_verified_process(handle, true);
                    }
                    anyhow::bail!("native COM worker did not exit after shutdown acknowledgement")
                }
            }
        };
        ensure!(status.success(), "native COM worker exited with {status}");
        let handle = self
            .office_handle
            .take()
            .context("verified Excel process handle was already closed")?;
        let exited = audit_core::native_office_com::wait_verified_process_exit(
            handle,
            Duration::from_secs(5),
        );
        audit_core::native_office_com::close_verified_process(handle, !exited);
        ensure!(
            exited,
            "Excel.EXE remained after successful worker shutdown"
        );
        Ok(())
    }
}

/// This test invokes the Audit binary itself and requires a real Excel COM
/// installation. It is intentionally ignored until explicitly authorized.
#[test]
#[ignore = "requires Windows Microsoft Excel COM"]
fn native_com_worker_reuses_one_excel_for_two_workbooks() -> Result<()> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let root: PathBuf =
        std::env::temp_dir().join(format!("audit-native-com-{}-{nonce}", std::process::id()));
    fs::create_dir(&root)?;
    let mut worker = Worker::start()?;
    let ready = worker.response()?;
    ensure!(
        ready["status"] == "session_ready" && ready["com_available"] == true,
        "Excel COM not available: {ready}"
    );
    ensure!(
        ready["engine_process_owned"] == true
            && ready["initialization_stages_ms"]["com_initialize_ms"].is_number(),
        "worker did not confirm isolated process ownership and startup timings: {ready}"
    );
    let process_id = ready["engine_process_id"]
        .as_u64()
        .context("Office process id")? as u32;
    let created = ready["engine_process_created"]
        .as_u64()
        .context("Office process creation time")?;
    worker.office_handle = Some(
        audit_core::native_office_com::open_verified_process(process_id, created)
            .context("verify Office process")?,
    );

    for (index, (left, right)) in [(2.0, 3.0), (7.0, 5.0)].into_iter().enumerate() {
        let input = root.join(format!("input-{index}.xlsx"));
        let output = root.join(format!("output-{index}.xlsx"));
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_number(0, 0, left)?;
        sheet.write_number(0, 1, right)?;
        sheet.write_formula(0, 2, "=SUM(A1:B1)")?;
        workbook.save(&input)?;
        let original = fs::read(&input)?;
        let report = worker.request(json!({"type":"recalculate", "input":input.to_string_lossy(), "output":output.to_string_lossy()}))?;
        ensure!(
            report["status"] == "recalculated",
            "recalculation failed: {report}"
        );
        ensure!(
            report["engine_process_id"].as_u64() == Some(u64::from(process_id)),
            "Office session was not reused"
        );
        ensure!(fs::read(&input)? == original, "source workbook changed");
        let values = excel_reader::read_value_columns(&output, 3)?;
        let actual = values[0].values.get("C1").context("formula result C1")?;
        match actual {
            CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
                ensure!((*value - left - right).abs() < 1e-9, "unexpected result")
            }
            CellValue::Integer(value) => ensure!(
                (*value as f64 - left - right).abs() < 1e-9,
                "unexpected result"
            ),
            _ => anyhow::bail!("formula result is not numeric: {actual:?}"),
        }
    }
    worker.shutdown_and_verify()?;
    fs::remove_dir_all(root)?;
    Ok(())
}

/// The Debug worker delays its successful response for 40 seconds. The
/// production parent timeout is 35 seconds when configured to its minimum,
/// so this exercises caller timeout cleanup without relying on a nondeterministic
/// Excel formula hang. It verifies the exact owned Excel process is gone.
#[test]
#[ignore = "requires Windows Microsoft Excel COM; exercises a controlled worker-response timeout"]
fn native_parent_timeout_terminates_worker_and_owned_excel() -> Result<()> {
    let nonce = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    let root: PathBuf = std::env::temp_dir().join(format!(
        "audit-native-com-timeout-{}-{nonce}",
        std::process::id()
    ));
    fs::create_dir(&root)?;
    let input = root.join("input.xlsx");
    let output = root.join("output.xlsx");
    let original = write_simple_formula_book(&input)?;
    let _environment = EnvironmentRestore::set(&[
        ("AUDIT_RECALC_TIMEOUT_SECONDS", "5"),
        ("AUDIT_TEST_COM_RESPONSE_DELAY_MS", "40000"),
    ]);
    let mut session =
        audit_core::s1_recalculation::S1RecalculationSession::new_native_with_executable(
            env!("CARGO_BIN_EXE_AuditTools"),
            "excel",
        )?;
    let error = session
        .recalculate(&input, &output)
        .expect_err("controlled response delay should exceed the parent timeout");
    ensure!(
        format!("{error:#}").contains("超过 35 秒无响应"),
        "unexpected timeout result: {error:#}"
    );
    let (process_id, created) = session
        .last_owned_excel_identity()
        .context("timeout cleanup did not record the created Excel identity")?;
    let cleanup_succeeded = session.last_owned_excel_cleanup_succeeded();
    if let Some(handle) = audit_core::native_office_com::open_verified_process(process_id, created)
    {
        let forced_test_cleanup =
            audit_core::native_office_com::close_verified_process(handle, true);
        anyhow::bail!("timeout cleanup left owned Excel PID {process_id} alive (parent cleanup={cleanup_succeeded:?}, test cleanup={forced_test_cleanup})");
    }
    ensure!(
        cleanup_succeeded == Some(true),
        "owned Excel cleanup was not confirmed: {cleanup_succeeded:?}"
    );
    ensure!(
        fs::read(&input)? == original,
        "timeout cleanup modified the source workbook"
    );
    ensure!(
        !output.exists(),
        "timed-out worker published an output workbook"
    );
    drop(session);
    fs::remove_dir_all(root)?;
    Ok(())
}
