//! Best-effort, bounded, asynchronous diagnostic journal for GUI runs.
//! No workbook reads and no audit decisions belong here.
use serde_json::{json, Value};
use std::{
    fs::{self, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
    sync::{
        mpsc::{self, SyncSender},
        Mutex, OnceLock,
    },
    time::Duration,
};

thread_local! {
    static CONTEXT: std::cell::RefCell<Option<Value>> = const { std::cell::RefCell::new(None) };
}
pub(crate) struct Scope(Option<Value>);
impl Drop for Scope {
    fn drop(&mut self) {
        CONTEXT.with(|slot| *slot.borrow_mut() = self.0.take());
    }
}
/// Carries only run identity through a business worker, never source cell data.
pub(crate) fn scope(state: &Value) -> Scope {
    let context = json!({"runId": state["runId"],"feature":state["runFeature"],"path":state["diagnosticLogPath"]});
    Scope(CONTEXT.with(|slot| slot.replace(Some(context))))
}
pub(crate) fn trace(stage: &str, fields: Value) {
    CONTEXT.with(|slot| {
        let context = slot.borrow();
        let Some(context) = context.as_ref() else {
            return;
        };
        let Some(path) = context["path"].as_str() else {
            return;
        };
        let queued = record(
            PathBuf::from(path),
            json!({"runId":context["runId"],"feature":context["feature"],
            "time":chrono::Local::now().to_rfc3339(),"technical":true,"level":"info",
            "stage":stage,"message":stage,"fields":fields}),
        );
        if !queued {
            let mut failed = failures().lock().unwrap_or_else(|error| error.into_inner());
            failed.insert(PathBuf::from(path));
            while failed.len() > 80 {
                failed.pop_first();
            }
        }
    });
}

pub(crate) fn trace_error(stage: &str, error: &anyhow::Error) {
    CONTEXT.with(|slot| {
        let context = slot.borrow();
        let Some(context) = context.as_ref() else {
            return;
        };
        let Some(path) = context["path"].as_str() else {
            return;
        };
        let backtrace = if error.backtrace().status() == std::backtrace::BacktraceStatus::Captured {
            Some(error.backtrace().to_string())
        } else {
            None
        };
        if !record(
            PathBuf::from(path),
            json!({"runId":context["runId"],"feature":context["feature"],
            "time":chrono::Local::now().to_rfc3339(),"level":"error","technical":true,"stage":stage,
            "message":format!("{error:#}"),"backtrace":backtrace}),
        ) {
            let mut failed = failures().lock().unwrap_or_else(|e| e.into_inner());
            failed.insert(PathBuf::from(path));
            while failed.len() > 80 {
                failed.pop_first();
            }
        }
    });
}

const QUEUE_CAPACITY: usize = 2048;
const MAX_FILE_BYTES: u64 = 16 * 1024 * 1024;
const RETAIN_RUNS: usize = 40;
struct Writer {
    normal: SyncSender<Command>,
    important: SyncSender<Command>,
}
enum Command {
    Event { path: PathBuf, event: Value },
    Flush(mpsc::Sender<()>),
}
fn sender() -> Option<&'static Writer> {
    static WRITER: OnceLock<Option<Writer>> = OnceLock::new();
    WRITER
        .get_or_init(|| {
            let (tx, rx) = mpsc::sync_channel(QUEUE_CAPACITY);
            let (important, priority) = mpsc::sync_channel(256);
            std::thread::Builder::new()
                .name("audit-diagnostics".into())
                .spawn(move || loop {
                    let command = match priority.try_recv() {
                        Ok(command) => command,
                        Err(_) => match rx.recv_timeout(Duration::from_millis(20)) {
                            Ok(command) => command,
                            Err(mpsc::RecvTimeoutError::Timeout) => continue,
                            Err(mpsc::RecvTimeoutError::Disconnected) => break,
                        },
                    };
                    match command {
                        Command::Event { path, event } => {
                            if let Err(error) = append_event(&path, &event) {
                                let mut failed =
                                    failures().lock().unwrap_or_else(|e| e.into_inner());
                                let first = failed.insert(path);
                                while failed.len() > 80 {
                                    failed.pop_first();
                                }
                                if first {
                                    eprintln!("Audit diagnostic journal unavailable: {error}");
                                }
                            }
                        }
                        Command::Flush(done) => {
                            let _ = done.send(());
                        }
                    }
                })
                .ok()
                .map(|_| Writer {
                    normal: tx,
                    important,
                })
        })
        .as_ref()
}
fn failures() -> &'static Mutex<std::collections::BTreeSet<PathBuf>> {
    static FAILED: OnceLock<Mutex<std::collections::BTreeSet<PathBuf>>> = OnceLock::new();
    FAILED.get_or_init(|| Mutex::new(std::collections::BTreeSet::new()))
}
pub(crate) fn incomplete(path: &Path) -> bool {
    failures()
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .contains(path)
}
/// False means diagnostics were not queued; caller retains that fact in run state.
pub(crate) fn record(path: PathBuf, mut event: Value) -> bool {
    static SEQUENCE: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
    event["journalSequence"] = json!(SEQUENCE.fetch_add(1, std::sync::atomic::Ordering::Relaxed));
    let important = matches!(event["level"].as_str(), Some("warn" | "error"));
    sender().is_some_and(|writer| {
        let queue = if important {
            &writer.important
        } else {
            &writer.normal
        };
        queue.try_send(Command::Event { path, event }).is_ok()
    })
}
/// Used only by non-GUI acceptance runners before their process exits.
pub(crate) fn flush() -> bool {
    let (tx, rx) = mpsc::channel();
    sender().is_some_and(|writer| {
        writer
            .important
            .try_send(Command::Flush(tx.clone()))
            .is_ok()
            && writer.normal.try_send(Command::Flush(tx)).is_ok()
            && rx.recv_timeout(Duration::from_secs(2)).is_ok()
            && rx.recv_timeout(Duration::from_secs(2)).is_ok()
    })
}
fn append_event(path: &Path, event: &Value) -> io::Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| io::Error::other("diagnostic directory missing"))?;
    fs::create_dir_all(parent)?;
    if !path.exists() {
        prune_runs(parent, path)?;
    }
    // Segment without overwriting earlier events of this run.
    let mut target = path.to_path_buf();
    let mut segment = 0u32;
    while fs::metadata(&target).is_ok_and(|meta| meta.len() >= MAX_FILE_BYTES) {
        segment += 1;
        target = path.with_file_name(format!(
            "{}.segment-{segment:06}.jsonl",
            path.file_stem().unwrap_or_default().to_string_lossy()
        ));
    }
    let mut file = OpenOptions::new().create(true).append(true).open(target)?;
    serde_json::to_writer(&mut file, event)?;
    file.write_all(b"\n")
}
fn prune_runs(directory: &Path, current: &Path) -> io::Result<()> {
    let mut files = fs::read_dir(directory)?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path != current
                && path.file_name().and_then(|n| n.to_str()).is_some_and(|n| {
                    n.starts_with("gui-run-") && n.ends_with(".jsonl") && !n.contains(".segment-")
                })
        })
        .collect::<Vec<_>>();
    files.sort();
    let excess = files.len().saturating_sub(RETAIN_RUNS - 1);
    for path in files.into_iter().take(excess) {
        fs::remove_file(&path)?;
        let prefix = format!(
            "{}.segment-",
            path.file_stem().unwrap_or_default().to_string_lossy()
        );
        for entry in fs::read_dir(directory)?.filter_map(Result::ok) {
            let candidate = entry.path();
            if candidate
                .file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with(&prefix) && name.ends_with(".jsonl"))
            {
                fs::remove_file(candidate)?;
            }
        }
    }
    Ok(())
}
pub(crate) fn event(state: &mut Value, level: &str, detail: bool, message: &str) {
    let Some(path) = state["diagnosticLogPath"].as_str().map(PathBuf::from) else {
        return;
    };
    let sequence = state["diagnosticSequence"].as_u64().unwrap_or(0) + 1;
    state["diagnosticSequence"] = json!(sequence);
    let event = json!({
        "runId": state["runId"], "sequence": sequence,
        "time": chrono::Local::now().to_rfc3339(), "feature": state["runFeature"],
        "level": level, "technical": detail, "message": message,
        "currentWork": state["currentWork"],
    });
    if !record(path, event) {
        state["diagnosticIncomplete"] = json!(true);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn work(name: &str) -> PathBuf {
        static ID: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "audit-journal-{}-{}-{name}",
            std::process::id(),
            ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        ));
        fs::create_dir_all(&path).unwrap();
        path
    }
    #[test]
    fn async_journal_keeps_order_and_unicode_outside_gui_ring() {
        let root = work("unicode");
        let path = root.join("gui-run-test.jsonl");
        for i in 0..100 {
            assert!(record(
                path.clone(),
                json!({"runId":"test", "seq":i,"file":"机构.xlsx"})
            ));
        }
        assert!(flush());
        let events: Vec<Value> = fs::read_to_string(&path)
            .unwrap()
            .lines()
            .map(|line| serde_json::from_str(line).unwrap())
            .collect();
        assert_eq!(events.len(), 100);
        for (i, event) in events.iter().enumerate() {
            assert_eq!(event["seq"], i);
            assert_eq!(event["file"], "机构.xlsx");
        }
        fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn failed_journal_is_observable_and_does_not_stop_following_run() {
        let root = work("failure");
        let blocked = root.join("not-directory");
        fs::write(&blocked, b"x").unwrap();
        let bad = blocked.join("gui-run-bad.jsonl");
        assert!(record(bad.clone(), json!({"level":"error"})));
        let good = root.join("gui-run-good.jsonl");
        assert!(record(good.clone(), json!({"runId":"good"})));
        assert!(flush());
        assert!(incomplete(&bad));
        assert!(!incomplete(&good));
        assert!(good.is_file());
        fs::remove_dir_all(root).unwrap();
    }
    #[test]
    fn journal_rotation_and_retention_touch_only_owned_files() {
        let root = work("retention");
        fs::write(root.join("keep.jsonl"), b"keep").unwrap();
        for i in 0..RETAIN_RUNS + 2 {
            append_event(
                &root.join(format!("gui-run-{i:03}.jsonl")),
                &json!({"seq":i}),
            )
            .unwrap();
        }
        let count = fs::read_dir(&root)
            .unwrap()
            .filter_map(Result::ok)
            .filter(|entry| entry.file_name().to_string_lossy().starts_with("gui-run-"))
            .count();
        assert_eq!(count, RETAIN_RUNS);
        assert_eq!(fs::read(root.join("keep.jsonl")).unwrap(), b"keep");
        let last = root.join("gui-run-999.jsonl");
        let file = fs::File::create(&last).unwrap();
        file.set_len(MAX_FILE_BYTES).unwrap();
        drop(file);
        append_event(&last, &json!({"message":"after rotation"})).unwrap();
        assert_eq!(fs::metadata(&last).unwrap().len(), MAX_FILE_BYTES);
        let segment = last.with_file_name("gui-run-999.segment-000001.jsonl");
        assert!(fs::read_to_string(&segment)
            .unwrap()
            .contains("after rotation"));
        fs::remove_dir_all(root).unwrap();
    }
}
