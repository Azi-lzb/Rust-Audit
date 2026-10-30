#![cfg(windows)]

use anyhow::{ensure, Context, Result};
use excel_reader::CellValue;
use excel_writer::Workbook;
use serde_json::{json, Value};
use std::{
    fs,
    io::{BufRead, BufReader, Write},
    os::windows::process::CommandExt,
    path::PathBuf,
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
        let mut child = Command::new(env!("CARGO_BIN_EXE_open-data-audit"))
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

impl Drop for Worker {
    fn drop(&mut self) {
        let _ = serde_json::to_writer(&mut self.stdin, &json!({"type":"shutdown"}));
        let _ = self.stdin.write_all(b"\n");
        let _ = self.stdin.flush();
        let deadline = Instant::now() + Duration::from_secs(5);
        let mut exited = false;
        let mut clean_shutdown = false;
        while Instant::now() < deadline {
            match self.child.try_wait() {
                Ok(Some(status)) => {
                    exited = true;
                    clean_shutdown = status.success();
                    break;
                }
                Ok(None) => thread::sleep(Duration::from_millis(50)),
                Err(_) => break,
            }
        }
        if !exited {
            let _ = self.child.kill();
            let _ = self.child.wait();
        }
        if let Some(handle) = self.office_handle.take() {
            audit_core::native_office_com::close_verified_process(handle, !clean_shutdown);
        }
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
    drop(worker);
    fs::remove_dir_all(root)?;
    Ok(())
}
