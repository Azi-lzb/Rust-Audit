//! S3-F03 planning and institution-feedback write-back.
//!
//! Planning remains read-only and records an ExcelEditPlan plus conflict
//! workbook. The production GUI path applies a candidate through the
//! capability-gated ExcelModifier, creates a byte-identical backup beside the
//! selected target, then updates that target in place with rollback on failure.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_editor::{
    inspect_xlsx_profile, CellValue as EditCellValue, EditExecutionReport, EditOperation,
    ExcelEditPlan, ExcelModifier, PreservationRequirement, WorkbookFormat,
};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use excel_writer::{Color, Format, Workbook};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

const REQUIRED_ID_HEADERS: [&str; 4] = ["机构代码", "机构名称", "指标代码", "指标名称"];
const DATA_VALUE_HEADER: &str = "数据值";
const CHANGE_AMOUNT_PREFIX: &str = "增减额";
const REASON_HEADER: &str = "变动原因";
const READER: &str = "calamine";
const BUSINESS_ENGINE: &str = "audit-core (Rust)";
const CONFLICT_WRITER: &str = "rust_xlsxwriter";
const CONFLICT_HEADERS: [&str; 10] = [
    "状态",
    "机构代码",
    "机构名称",
    "指标代码",
    "指标名称",
    "数据值",
    "增减额",
    "目标说明文件行号",
    "反馈文件及原因",
    "冲突原因",
];
const CONFLICT_WIDTHS: [f64; 10] = [10.0, 14.0, 26.0, 14.0, 32.0, 14.0, 14.0, 20.0, 72.0, 38.0];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);
static PLAN_COUNTER: AtomicU64 = AtomicU64::new(0);
static STAGING_COUNTER: AtomicU64 = AtomicU64::new(0);

type Identity = (String, String, String, String);

#[derive(Clone, Debug, Serialize)]
pub struct S3ExplanationFeedbackPlanSummary {
    pub status: &'static str,
    pub applied: bool,
    pub target_workbook: PathBuf,
    pub plan_file: PathBuf,
    pub candidate_output_workbook: PathBuf,
    pub imported: usize,
    pub unmatched: usize,
    pub conflicts: usize,
    pub planned_cell_updates: usize,
    pub conflict_report_path: Option<PathBuf>,
    pub reader: &'static str,
    pub business_engine: &'static str,
    pub conflict_writer: Option<&'static str>,
    pub target_read_time_ms: u128,
    pub feedback_read_time_ms: u128,
    pub conflict_write_time_ms: u128,
    pub plan_write_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
pub struct S3ExplanationFeedbackExecutionReport {
    pub output_workbook: PathBuf,
    pub backup_workbook: Option<PathBuf>,
    pub rollback_cleanup_warning: Option<String>,
    pub imported: usize,
    pub unmatched: usize,
    pub conflicts: usize,
    pub planned_cell_updates: usize,
    pub conflict_report_path: Option<PathBuf>,
    pub edit: Option<EditExecutionReport>,
    pub copied_without_edits: bool,
    pub target_read_time_ms: u128,
    pub feedback_read_time_ms: u128,
    pub conflict_write_time_ms: u128,
    pub writeback_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct InputSnapshot {
    path: PathBuf,
    sha256: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct PlanCounts {
    imported: usize,
    unmatched: usize,
    conflicts: usize,
    planned_cell_updates: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct PlanDocument {
    schema_version: u32,
    feature: String,
    status: String,
    created_at: String,
    applied: bool,
    target: InputSnapshot,
    feedback_sources: Vec<InputSnapshot>,
    candidate_output_workbook: PathBuf,
    counts: PlanCounts,
    conflict_report_path: Option<PathBuf>,
    edit_plan: Option<ExcelEditPlan>,
}

#[derive(Clone, Debug, Serialize)]
struct ConflictRow([String; 10]);

#[derive(Clone, Debug)]
struct TargetRow {
    row_number: usize,
    organization_name: CellValue,
    indicator_name: CellValue,
}

#[derive(Clone, Debug)]
struct HeaderColumns {
    row_number: usize,
    indexes: HashMap<String, usize>,
}

#[derive(Clone, Debug)]
struct ValueGrid {
    name: String,
    max_row: usize,
    rows: BTreeMap<usize, BTreeMap<usize, CellValue>>,
}

impl From<SheetValues> for ValueGrid {
    fn from(sheet: SheetValues) -> Self {
        let mut rows = BTreeMap::<usize, BTreeMap<usize, CellValue>>::new();
        for (address, value) in sheet.values {
            if let Some((row, column)) = parse_cell_address(&address) {
                rows.entry(row).or_default().insert(column, value);
            }
        }
        Self {
            name: sheet.name,
            max_row: sheet.max_row,
            rows,
        }
    }
}

impl ValueGrid {
    fn cell(&self, row: usize, column: usize) -> CellValue {
        self.rows
            .get(&row)
            .and_then(|cells| cells.get(&column))
            .cloned()
            .unwrap_or(CellValue::Empty)
    }

    fn has_header_text(&self, expected: &str) -> bool {
        (1..=self.max_row.min(10)).any(|row| {
            self.rows.get(&row).is_some_and(|cells| {
                cells.values().any(|value| {
                    !matches!(value, CellValue::Empty) && python_string(value).trim() == expected
                })
            })
        })
    }

    fn headers(&self) -> Result<HeaderColumns> {
        for row_number in 1..=self.max_row.min(10) {
            let mut indexes = HashMap::new();
            let mut ordered_names = Vec::new();
            if let Some(cells) = self.rows.get(&row_number) {
                for (column, value) in cells {
                    if matches!(value, CellValue::Empty) {
                        continue;
                    }
                    let name = python_string(value).trim().to_owned();
                    if name.is_empty() {
                        continue;
                    }
                    if !indexes.contains_key(&name) {
                        ordered_names.push(name.clone());
                    }
                    // Match Python's dict comprehension: a duplicate header
                    // retains its first insertion position but points to the
                    // last column carrying that text.
                    indexes.insert(name, *column);
                }
            }
            if !REQUIRED_ID_HEADERS
                .iter()
                .all(|name| indexes.contains_key(*name))
            {
                continue;
            }
            let value_header = ordered_names
                .iter()
                .find(|name| name.as_str() == DATA_VALUE_HEADER || name.starts_with("本期金额"));
            let change_header = ordered_names
                .iter()
                .find(|name| name.starts_with(CHANGE_AMOUNT_PREFIX));
            if let (Some(value_header), Some(change_header)) = (value_header, change_header) {
                if indexes.contains_key(REASON_HEADER) {
                    let value_column = indexes[value_header];
                    let change_column = indexes[change_header];
                    indexes.insert(DATA_VALUE_HEADER.to_owned(), value_column);
                    indexes.insert(CHANGE_AMOUNT_PREFIX.to_owned(), change_column);
                    return Ok(HeaderColumns {
                        row_number,
                        indexes,
                    });
                }
            }
        }
        bail!("说明文件缺少机构代码、机构名称、指标代码、指标名称、数据值、增减额或变动原因列")
    }
}

/// Compute S3-F03 matching and conflict outcomes without changing the target.
/// The JSON plan contains the proposed cell updates and source hashes so a
/// later, separately validated ExcelModifier step can reject stale inputs.
pub fn plan_explanation_feedback_import(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    candidate_output_path: impl AsRef<Path>,
    plan_output_path: impl AsRef<Path>,
) -> Result<S3ExplanationFeedbackPlanSummary> {
    plan_explanation_feedback_import_internal(
        target_path,
        feedback_paths,
        candidate_output_path,
        plan_output_path,
        true,
        None,
    )
    .map(|(summary, _)| summary)
}

fn plan_explanation_feedback_import_internal(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    candidate_output_path: impl AsRef<Path>,
    plan_output_path: impl AsRef<Path>,
    persist_plan: bool,
    mut on_progress: Option<&mut dyn FnMut(&str)>,
) -> Result<(S3ExplanationFeedbackPlanSummary, Option<ExcelEditPlan>)> {
    let started = Instant::now();
    let target_path = target_path.as_ref();
    let candidate_output_path = candidate_output_path.as_ref();
    let plan_output_path = plan_output_path.as_ref();
    ensure!(!feedback_paths.is_empty(), "请至少选择一个机构反馈说明文件");
    ensure_workbook_path(target_path, true)?;
    ensure!(
        target_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 目标说明文件目前仅接受 .xlsx；为保护宏和其他工作簿对象，未尝试写入 .xlsm"
    );
    ensure_workbook_path(candidate_output_path, false)?;
    ensure!(
        candidate_output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 候选输出必须使用 .xlsx"
    );
    ensure!(
        !candidate_output_path.exists(),
        "候选输出路径已存在，未覆盖：{}",
        candidate_output_path.display()
    );
    ensure!(
        plan_output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
        "S3-F03 计划输出必须使用 .json"
    );
    ensure!(
        !plan_output_path.exists(),
        "导入计划路径已存在，未覆盖：{}",
        plan_output_path.display()
    );
    ensure_distinct_paths(&[target_path, candidate_output_path, plan_output_path])?;
    for feedback_path in feedback_paths {
        ensure_workbook_path(feedback_path, true)?;
        // 反馈文件只读，扩展名放开到 .xls/.xlsb/.ods 等 calamine 能读的格式
        // （真实机构反馈常为 .xls）；目标说明工作簿仍只允许 .xlsx，写回走
        // 现有 ExcelEditPlan 编辑后端，保持格式不动。
        ensure!(
            feedback_path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    matches!(
                        extension.to_ascii_lowercase().as_str(),
                        "xlsx" | "xlsm" | "xls" | "xlsb" | "ods"
                    )
                }),
            "S3-F03 反馈工作簿必须是 calamine 可读的表格格式（.xlsx/.xlsm/.xls/.xlsb/.ods）：{}",
            feedback_path.display()
        );
    }

    let snapshot_paths = std::iter::once(target_path.to_path_buf())
        .chain(feedback_paths.iter().cloned())
        .collect::<Vec<_>>();
    let before_hashes = snapshot_paths
        .iter()
        .map(|path| sha256_file(path))
        .collect::<Result<Vec<_>>>()?;

    if let Some(callback) = on_progress.as_deref_mut() {
        callback(&format!("读取目标说明工作簿：{}", target_path.display()));
    }
    let target_read_started = Instant::now();
    let target_sheets = read_value_columns_with_iso_dates(target_path, usize::MAX)
        .with_context(|| format!("读取目标说明工作簿：{}", target_path.display()))?;
    let target_read_time_ms = target_read_started.elapsed().as_millis();
    let target_sheet = find_explanation_sheet(target_sheets)?;
    if let Some(callback) = on_progress.as_deref_mut() {
        callback(&format!("目标说明工作表={} 已读取", target_sheet.name));
    }
    let target_headers = target_sheet.headers()?;

    let mut targets = HashMap::<Identity, Vec<TargetRow>>::new();
    for row_number in target_headers.row_number + 1..=target_sheet.max_row {
        let key = identity(&target_sheet, row_number, &target_headers);
        if identity_is_nonempty(&key) {
            targets.entry(key).or_default().push(TargetRow {
                row_number,
                organization_name: target_sheet
                    .cell(row_number, target_headers.indexes["机构名称"]),
                indicator_name: target_sheet.cell(row_number, target_headers.indexes["指标名称"]),
            });
        }
    }

    let feedback_read_started = Instant::now();
    let mut feedback_order = Vec::<Identity>::new();
    let mut feedback_by_key = HashMap::<Identity, BTreeMap<String, BTreeSet<String>>>::new();
    for (index, feedback_path) in feedback_paths.iter().enumerate() {
        if let Some(callback) = on_progress.as_deref_mut() {
            callback(&format!(
                "读取机构反馈 [{}/{}]：{}",
                index + 1,
                feedback_paths.len(),
                feedback_path.display()
            ));
        }
        let sheets = read_value_columns_with_iso_dates(feedback_path, usize::MAX)
            .with_context(|| format!("读取机构反馈工作簿：{}", feedback_path.display()))?;
        let sheet = find_explanation_sheet(sheets)?;
        if let Some(callback) = on_progress.as_deref_mut() {
            callback(&format!(
                "机构反馈文件={} 工作表={} 已读取",
                feedback_path.display(),
                sheet.name
            ));
        }
        let headers = sheet.headers()?;
        let source_name = feedback_path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        let reason_column = headers.indexes[REASON_HEADER];
        for row_number in headers.row_number + 1..=sheet.max_row {
            let key = identity(&sheet, row_number, &headers);
            let reason_cell = sheet.cell(row_number, reason_column);
            let reason = if python_truthy(&reason_cell) {
                python_string(&reason_cell).trim().to_owned()
            } else {
                String::new()
            };
            if !identity_is_nonempty(&key) || reason.is_empty() {
                continue;
            }
            let reasons = feedback_by_key.entry(key.clone()).or_insert_with(|| {
                feedback_order.push(key.clone());
                BTreeMap::new()
            });
            reasons
                .entry(reason)
                .or_default()
                .insert(source_name.clone());
        }
    }
    let feedback_read_time_ms = feedback_read_started.elapsed().as_millis();
    if let Some(callback) = on_progress.as_deref_mut() {
        callback("按机构、指标和数据值匹配反馈原因并检测冲突");
    }

    let mut imported = 0usize;
    let mut unmatched = 0usize;
    let mut conflicts = 0usize;
    let mut operations = Vec::new();
    let mut conflict_rows = Vec::new();
    let target_reason_column = target_headers.indexes[REASON_HEADER];
    for key in feedback_order {
        let reasons = feedback_by_key
            .get(&key)
            .expect("feedback order and index are populated together");
        let Some(target_rows) = targets.get(&key) else {
            unmatched += 1;
            continue;
        };
        if reasons.len() != 1 || target_rows.len() != 1 {
            conflicts += 1;
            let reasons_text = reasons
                .iter()
                .map(|(reason, sources)| {
                    format!(
                        "{reason}（来源：{}）",
                        sources.iter().cloned().collect::<Vec<_>>().join("、")
                    )
                })
                .collect::<Vec<_>>()
                .join("\n");
            let mut causes = Vec::new();
            if target_rows.len() != 1 {
                causes.push(format!("目标说明文件存在 {} 行相同标识", target_rows.len()));
            }
            if reasons.len() != 1 {
                causes.push("多个反馈文件给出不同变动原因".to_owned());
            }
            let sample = &target_rows[0];
            conflict_rows.push(ConflictRow([
                "冲突".to_owned(),
                key.0.clone(),
                python_str_for_conflict(&sample.organization_name),
                key.1.clone(),
                python_str_for_conflict(&sample.indicator_name),
                key.2.trim_start_matches('#').to_owned(),
                key.3.trim_start_matches('#').to_owned(),
                target_rows
                    .iter()
                    .map(|row| row.row_number.to_string())
                    .collect::<Vec<_>>()
                    .join("、"),
                reasons_text,
                causes.join("；"),
            ]));
            continue;
        }
        let reason = reasons.keys().next().expect("one reason exists").clone();
        operations.push(EditOperation::SetCellValue {
            sheet: target_sheet.name.clone(),
            cell: cell_address(target_rows[0].row_number, target_reason_column),
            value: EditCellValue::Text(reason),
        });
        imported += 1;
    }

    let after_hashes = snapshot_paths
        .iter()
        .map(|path| sha256_file(path))
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        before_hashes == after_hashes,
        "规划期间输入工作簿发生变化，未写出导入计划"
    );

    let candidate_path = normalize_new_path(candidate_output_path)?;
    let plan_path = normalize_new_path(plan_output_path)?;
    ensure_distinct_paths(&[target_path, &candidate_path, &plan_path])?;
    let conflict_report_path = if conflict_rows.is_empty() {
        None
    } else {
        Some(conflict_report_path(target_path)?)
    };

    let edit_plan = if operations.is_empty() {
        None
    } else {
        let profile = inspect_xlsx_profile(target_path)
            .with_context(|| format!("检查目标工作簿结构：{}", target_path.display()))?;
        ensure!(
            profile.format == WorkbookFormat::Xlsx,
            "S3-F03 ExcelEditPlan 目前仅支持 XLSX 目标"
        );
        let request_id = format!(
            "s3-f03-{}-{}",
            Local::now().format("%Y%m%d%H%M%S"),
            PLAN_COUNTER.fetch_add(1, Ordering::Relaxed)
        );
        let plan = ExcelEditPlan {
            request_id,
            input: target_path.to_path_buf(),
            output: candidate_path.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: profile,
            operations,
            preserve: required_preservation(),
            drop_allowed: BTreeMap::new(),
        };
        plan.validate().context("校验 S3-F03 ExcelEditPlan")?;
        Some(plan)
    };

    let snapshots = snapshot_paths
        .iter()
        .zip(after_hashes)
        .map(|(path, sha256)| InputSnapshot {
            path: path.clone(),
            sha256,
        })
        .collect::<Vec<_>>();
    let mut snapshots = snapshots.into_iter();
    let target_snapshot = snapshots.next().expect("target snapshot exists");
    let feedback_snapshots = snapshots.collect::<Vec<_>>();
    if persist_plan {
        if let Some(callback) = on_progress.as_deref_mut() {
            callback("写入反馈冲突报告与回写计划");
        }
    }
    let (document, execution_edit_plan) = if persist_plan {
        (
            Some(PlanDocument {
                schema_version: 1,
                feature: "S3-F03".to_owned(),
                status: "planned_not_applied".to_owned(),
                created_at: Local::now().to_rfc3339(),
                applied: false,
                target: target_snapshot,
                feedback_sources: feedback_snapshots,
                candidate_output_workbook: candidate_path.clone(),
                counts: PlanCounts {
                    imported,
                    unmatched,
                    conflicts,
                    planned_cell_updates: imported,
                },
                conflict_report_path: conflict_report_path.clone(),
                edit_plan,
            }),
            None,
        )
    } else {
        (None, edit_plan)
    };

    let conflict_write_started = Instant::now();
    let conflict_temp = if let Some(path) = &conflict_report_path {
        let temporary = reserve_temporary_path(path, "xlsx")?;
        if let Err(error) = write_conflict_workbook(&temporary, &conflict_rows) {
            let _ = fs::remove_file(&temporary);
            return Err(error);
        }
        Some(temporary)
    } else {
        None
    };
    let conflict_write_time_ms = if conflict_temp.is_some() {
        conflict_write_started.elapsed().as_millis()
    } else {
        0
    };

    let plan_write_started = Instant::now();
    let plan_write_time_ms = if let Some(document) = &document {
        let plan_temp = reserve_temporary_path(&plan_path, "json")?;
        let plan_write_result = (|| -> Result<()> {
            let mut file = OpenOptions::new()
                .write(true)
                .truncate(true)
                .open(&plan_temp)
                .with_context(|| format!("打开临时导入计划：{}", plan_temp.display()))?;
            serde_json::to_writer_pretty(&mut file, document).context("序列化 S3-F03 导入计划")?;
            file.write_all(b"\n").context("写入导入计划换行")?;
            file.sync_all().context("刷新导入计划到磁盘")?;
            Ok(())
        })();
        if let Err(error) = plan_write_result {
            let _ = fs::remove_file(&plan_temp);
            if let Some(final_path) = &conflict_report_path {
                let _ = fs::remove_file(final_path);
            }
            return Err(error);
        }
        if let (Some(temporary), Some(final_path)) = (&conflict_temp, &conflict_report_path) {
            if let Err(error) = publish_without_overwrite(temporary, final_path) {
                let _ = fs::remove_file(&plan_temp);
                return Err(error);
            }
        }
        if let Err(error) = publish_without_overwrite(&plan_temp, &plan_path) {
            if let Some(final_path) = &conflict_report_path {
                let _ = fs::remove_file(final_path);
            }
            return Err(error);
        }
        plan_write_started.elapsed().as_millis()
    } else {
        if let (Some(temporary), Some(final_path)) = (&conflict_temp, &conflict_report_path) {
            publish_without_overwrite(temporary, final_path)?;
        }
        0
    };

    let summary = S3ExplanationFeedbackPlanSummary {
        status: "planned_not_applied",
        applied: false,
        target_workbook: target_path.to_path_buf(),
        plan_file: plan_path,
        candidate_output_workbook: candidate_path,
        imported,
        unmatched,
        conflicts,
        planned_cell_updates: imported,
        conflict_report_path,
        reader: READER,
        business_engine: BUSINESS_ENGINE,
        conflict_writer: (!conflict_rows.is_empty()).then_some(CONFLICT_WRITER),
        target_read_time_ms,
        feedback_read_time_ms,
        conflict_write_time_ms,
        plan_write_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    };
    Ok((summary, execution_edit_plan))
}

/// Retained independent-copy workflow for callers that explicitly request it.
/// The existing plan and conflict report remain the source of matching and
/// conflict behavior. The candidate is published only after Excel-Core
/// succeeds and all original input hashes still match their pre-run values.
/// The S3-F03 GUI uses execute_explanation_feedback_import, which follows the
/// confirmed backup-then-update-in-place contract.
pub fn execute_explanation_feedback_import_to_copy(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    output_directory: impl AsRef<Path>,
    modifier: &ExcelModifier,
) -> Result<S3ExplanationFeedbackExecutionReport> {
    execute_explanation_feedback_import_to_copy_with_progress(
        target_path,
        feedback_paths,
        output_directory,
        modifier,
        |_| {},
    )
}

pub fn execute_explanation_feedback_import_to_copy_with_progress<F>(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    output_directory: impl AsRef<Path>,
    modifier: &ExcelModifier,
    mut on_progress: F,
) -> Result<S3ExplanationFeedbackExecutionReport>
where
    F: FnMut(&str),
{
    let started = Instant::now();
    let target_path = target_path.as_ref();
    let output_directory = output_directory.as_ref();
    ensure_workbook_path(target_path, true)?;
    ensure!(
        target_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 目标说明文件目前仅接受 .xlsx"
    );
    ensure!(!feedback_paths.is_empty(), "请至少选择一个机构反馈说明文件");
    fs::create_dir_all(output_directory)
        .with_context(|| format!("创建 S3-F03 输出目录：{}", output_directory.display()))?;
    let output_directory = fs::canonicalize(output_directory)
        .with_context(|| format!("解析 S3-F03 输出目录：{}", output_directory.display()))?;
    // normalize_new_path strips the Windows verbatim prefix when it
    // canonicalizes the candidate parent, and the edit report's output path is
    // compared against the local candidate below; keep both representations
    // identical or the same location compares unequal.
    let output_directory = strip_windows_verbatim_prefix(output_directory);
    ensure!(output_directory.is_dir(), "S3-F03 输出路径不是目录");

    let snapshot_paths = std::iter::once(target_path.to_path_buf())
        .chain(feedback_paths.iter().cloned())
        .collect::<Vec<_>>();
    let source_hashes = snapshot_paths
        .iter()
        .map(|path| sha256_file(path))
        .collect::<Result<Vec<_>>>()?;
    let output_path = next_feedback_output_path(&output_directory, target_path)?;
    let staging = S3FeedbackStagingDirectory::create(&output_directory)?;
    let candidate_path = staging.path().join("反馈写回结果.xlsx");
    let plan_path = staging.path().join("反馈写回计划.json");
    on_progress("构建回写方案：读取目标与反馈工作簿、匹配说明行");
    let (summary, edit_plan) = plan_explanation_feedback_import_internal(
        target_path,
        feedback_paths,
        &candidate_path,
        &plan_path,
        false,
        Some(&mut on_progress),
    )
    .context("规划 S3-F03 机构说明回写")?;

    ensure!(
        source_hashes
            == snapshot_paths
                .iter()
                .map(|path| sha256_file(path))
                .collect::<Result<Vec<_>>>()?,
        "S3-F03 规划期间输入工作簿发生变化，未应用回写"
    );
    on_progress("写入候选说明工作簿");
    let writeback_started = Instant::now();
    let edit = if let Some(edit_plan) = edit_plan {
        let edit = modifier.execute(&edit_plan).with_context(|| {
            summary
                .conflict_report_path
                .as_ref()
                .map(|path| {
                    format!(
                        "通过 Excel-Core 写入 S3-F03 机构反馈说明；冲突明细见 {}",
                        path.display()
                    )
                })
                .unwrap_or_else(|| "通过 Excel-Core 写入 S3-F03 机构反馈说明".to_owned())
        })?;
        ensure!(
            edit.output == candidate_path && candidate_path.is_file(),
            "S3-F03 Excel-Core 未生成请求的候选说明文件"
        );
        Some(edit)
    } else {
        fs::copy(target_path, &candidate_path)
            .with_context(|| format!("复制无待写入内容的说明文件：{}", target_path.display()))?;
        None
    };
    ensure!(
        source_hashes
            == snapshot_paths
                .iter()
                .map(|path| sha256_file(path))
                .collect::<Result<Vec<_>>>()?,
        "S3-F03 写入期间源说明或反馈文件发生变化，候选结果已丢弃"
    );
    let writeback_time_ms = writeback_started.elapsed().as_millis();
    on_progress("验证目标与反馈输入未变化");
    publish_staged_file(&candidate_path, &output_path)?;
    on_progress(&format!("发布回写结果：{}", output_path.display()));
    Ok(S3ExplanationFeedbackExecutionReport {
        output_workbook: output_path,
        backup_workbook: None,
        rollback_cleanup_warning: None,
        imported: summary.imported,
        unmatched: summary.unmatched,
        conflicts: summary.conflicts,
        planned_cell_updates: summary.planned_cell_updates,
        conflict_report_path: summary.conflict_report_path,
        edit,
        copied_without_edits: summary.planned_cell_updates == 0,
        target_read_time_ms: summary.target_read_time_ms,
        feedback_read_time_ms: summary.feedback_read_time_ms,
        conflict_write_time_ms: summary.conflict_write_time_ms,
        writeback_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    })
}

/// Apply institution feedback to the selected explanation workbook in place.
/// A byte-for-byte backup is created beside the target before the candidate
/// workbook replaces it. If publication fails, the original target is restored
/// from a same-directory rollback copy; both backup and recovery paths are
/// included in errors when restoration itself fails.
pub fn execute_explanation_feedback_import(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    modifier: &ExcelModifier,
) -> Result<S3ExplanationFeedbackExecutionReport> {
    execute_explanation_feedback_import_with_progress(target_path, feedback_paths, modifier, |_| {})
}

pub fn execute_explanation_feedback_import_with_progress<F>(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    modifier: &ExcelModifier,
    mut on_progress: F,
) -> Result<S3ExplanationFeedbackExecutionReport>
where
    F: FnMut(&str),
{
    let started = Instant::now();
    let target_path = target_path.as_ref();
    on_progress("校验原地更新目标和机构反馈文件");
    ensure_workbook_path(target_path, true)?;
    ensure!(
        target_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 原地写回目前仅接受 .xlsx 目标说明文件"
    );
    ensure!(!feedback_paths.is_empty(), "请至少选择一个机构反馈说明文件");

    let target_path = strip_windows_verbatim_prefix(
        fs::canonicalize(target_path)
            .with_context(|| format!("解析 S3-F03 目标说明文件：{}", target_path.display()))?,
    );
    let feedback_paths = feedback_paths
        .iter()
        .map(|path| {
            fs::canonicalize(path)
                .with_context(|| format!("解析 S3-F03 机构反馈文件：{}", path.display()))
        })
        .collect::<Result<Vec<_>>>()?;
    let target_identity = normalized_identity(&target_path)?;
    for feedback_path in &feedback_paths {
        ensure!(
            normalized_identity(feedback_path)? != target_identity,
            "S3-F03 目标说明文件不能同时作为机构反馈输入"
        );
    }

    let parent = target_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .context("S3-F03 目标说明文件缺少父目录")?;
    let mut staging = S3FeedbackStagingDirectory::create(parent)?;
    let before_hashes = std::iter::once(target_path.clone())
        .chain(feedback_paths.iter().cloned())
        .map(|path| sha256_file(&path))
        .collect::<Result<Vec<_>>>()?;
    let target_permissions = fs::metadata(&target_path)
        .with_context(|| format!("读取目标说明文件属性：{}", target_path.display()))?
        .permissions();

    let staged_report = execute_explanation_feedback_import_to_copy_with_progress(
        &target_path,
        &feedback_paths,
        staging.path(),
        modifier,
        &mut on_progress,
    )
    .context("生成 S3-F03 原地写回候选文件")?;
    let candidate_path = staged_report.output_workbook;
    ensure!(
        candidate_path.starts_with(staging.path()) && candidate_path.is_file(),
        "S3-F03 候选说明文件未留在目标文件同目录暂存区"
    );
    fs::set_permissions(&candidate_path, target_permissions.clone())
        .with_context(|| format!("保留目标说明文件权限：{}", candidate_path.display()))?;

    on_progress("校验目标与反馈输入未被并发修改");
    let current_hashes = std::iter::once(target_path.clone())
        .chain(feedback_paths.iter().cloned())
        .map(|path| sha256_file(&path))
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        current_hashes == before_hashes,
        "S3-F03 写回前目标说明或反馈输入发生变化，未更新目标文件"
    );

    on_progress("备份原目标说明文件并校验备份");
    let backup_path = create_feedback_backup(&target_path, &target_permissions)?;
    ensure!(
        sha256_file(&backup_path)? == before_hashes[0],
        "S3-F03 备份校验失败；原目标未更新，备份文件保留在 {}",
        backup_path.display()
    );
    let final_hashes = std::iter::once(target_path.clone())
        .chain(feedback_paths.iter().cloned())
        .map(|path| sha256_file(&path))
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        final_hashes == before_hashes,
        "S3-F03 备份期间输入文件发生变化；目标未更新，备份文件保留在 {}",
        backup_path.display()
    );

    let rollback_path = staging.path().join(
        target_path
            .file_name()
            .context("S3-F03 目标说明文件缺少文件名")?,
    );
    on_progress("原地替换目标说明文件（保留备份并启用回滚保护）");
    let rollback_cleanup_warning = publish_candidate_in_place(
        &candidate_path,
        &target_path,
        &rollback_path,
        &backup_path,
        &mut staging,
    )?;
    on_progress(&format!(
        "机构反馈已写回目标说明文件；原文件备份：{}",
        backup_path.display()
    ));

    Ok(S3ExplanationFeedbackExecutionReport {
        output_workbook: target_path,
        backup_workbook: Some(backup_path),
        rollback_cleanup_warning,
        imported: staged_report.imported,
        unmatched: staged_report.unmatched,
        conflicts: staged_report.conflicts,
        planned_cell_updates: staged_report.planned_cell_updates,
        conflict_report_path: staged_report.conflict_report_path,
        edit: staged_report.edit,
        copied_without_edits: staged_report.copied_without_edits,
        target_read_time_ms: staged_report.target_read_time_ms,
        feedback_read_time_ms: staged_report.feedback_read_time_ms,
        conflict_write_time_ms: staged_report.conflict_write_time_ms,
        writeback_time_ms: staged_report.writeback_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    })
}

fn create_feedback_backup(target: &Path, permissions: &std::fs::Permissions) -> Result<PathBuf> {
    let parent = target.parent().context("目标说明文件缺少父目录")?;
    let stem = target
        .file_stem()
        .and_then(|value| value.to_str())
        .filter(|value| !value.trim().is_empty())
        .unwrap_or("说明文件");
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%3f");
    for _ in 0..128 {
        let sequence = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let path = parent.join(format!("{stem}_导入前备份_{timestamp}_{sequence}.xlsx"));
        let mut backup = match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("创建 S3-F03 原说明文件备份：{}", path.display()))
            }
        };
        let copy_result = (|| -> Result<()> {
            let mut source = File::open(target)
                .with_context(|| format!("打开待备份说明文件：{}", target.display()))?;
            std::io::copy(&mut source, &mut backup)
                .with_context(|| format!("复制 S3-F03 原说明文件到备份：{}", path.display()))?;
            backup
                .sync_all()
                .with_context(|| format!("刷新 S3-F03 原说明文件备份：{}", path.display()))?;
            Ok(())
        })();
        drop(backup);
        if let Err(error) = copy_result {
            let _ = fs::remove_file(&path);
            return Err(error);
        }
        fs::set_permissions(&path, permissions.clone())
            .with_context(|| format!("保留 S3-F03 备份文件权限：{}", path.display()))?;
        return Ok(path);
    }
    bail!("无法为 S3-F03 创建不覆盖现有文件的原说明备份")
}

fn publish_candidate_in_place(
    candidate: &Path,
    target: &Path,
    rollback: &Path,
    backup: &Path,
    staging: &mut S3FeedbackStagingDirectory,
) -> Result<Option<String>> {
    ensure!(!rollback.exists(), "S3-F03 回滚暂存路径已存在");
    fs::rename(target, rollback).with_context(|| {
        format!(
            "S3-F03 无法暂存原目标说明文件；备份保留在 {}",
            backup.display()
        )
    })?;
    if let Err(install_error) = fs::rename(candidate, target) {
        if let Err(restore_error) = fs::rename(rollback, target) {
            staging.preserve();
            bail!(
                "S3-F03 更新失败且自动回滚失败；原文件备份：{}；可恢复原件：{}；候选文件：{}；发布错误：{install_error}; 回滚错误：{restore_error}",
                backup.display(),
                rollback.display(),
                candidate.display()
            );
        }
        bail!(
            "S3-F03 更新失败，原目标文件已恢复；安全备份保留在 {}；发布错误：{install_error}",
            backup.display()
        );
    }
    Ok(fs::remove_file(rollback).err().map(|error| {
        format!(
            "S3-F03 已更新目标；回滚暂存文件清理失败：{} ({error})",
            rollback.display()
        )
    }))
}

fn next_feedback_output_path(output_directory: &Path, target_path: &Path) -> Result<PathBuf> {
    let stem = target_path
        .file_stem()
        .and_then(|value| value.to_str())
        .filter(|value| !value.trim().is_empty())
        .unwrap_or("说明文件");
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%3f");
    for suffix in 0..128 {
        let name = if suffix == 0 {
            format!("{stem}_回写_{timestamp}.xlsx")
        } else {
            format!("{stem}_回写_{timestamp}_{suffix}.xlsx")
        };
        let path = output_directory.join(name);
        if !path.exists() {
            return Ok(path);
        }
    }
    bail!("无法为 S3-F03 生成唯一的回写输出文件名")
}

fn publish_staged_file(staged: &Path, output: &Path) -> Result<()> {
    fs::hard_link(staged, output).with_context(|| {
        format!(
            "安全发布 S3-F03 回写结果失败（目标已存在或文件系统不支持硬链接）：{}",
            output.display()
        )
    })?;
    fs::remove_file(staged).with_context(|| format!("清理 S3-F03 暂存文件：{}", staged.display()))
}

struct S3FeedbackStagingDirectory {
    path: PathBuf,
    preserve_on_drop: bool,
}

impl S3FeedbackStagingDirectory {
    fn create(parent: &Path) -> Result<Self> {
        for _ in 0..128 {
            let sequence = STAGING_COUNTER.fetch_add(1, Ordering::Relaxed);
            let path = parent.join(format!(".audit-s3-f3-{}-{sequence}", std::process::id()));
            match fs::create_dir(&path) {
                Ok(()) => {
                    return Ok(Self {
                        path,
                        preserve_on_drop: false,
                    });
                }
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建 S3-F03 暂存目录：{}", path.display()))
                }
            }
        }
        bail!("无法创建唯一的 S3-F03 暂存目录")
    }

    fn path(&self) -> &Path {
        &self.path
    }

    fn preserve(&mut self) {
        self.preserve_on_drop = true;
    }
}

impl Drop for S3FeedbackStagingDirectory {
    fn drop(&mut self) {
        if !self.preserve_on_drop {
            let _ = fs::remove_dir_all(&self.path);
        }
    }
}

fn ensure_workbook_path(path: &Path, must_exist: bool) -> Result<()> {
    if must_exist {
        ensure!(path.is_file(), "工作簿不存在：{}", path.display());
    } else if path.exists() {
        bail!("目标路径已存在：{}", path.display());
    }
    Ok(())
}

fn ensure_distinct_paths(paths: &[&Path]) -> Result<()> {
    let identities = paths
        .iter()
        .map(|path| normalized_identity(path))
        .collect::<Result<BTreeSet<_>>>()?;
    ensure!(
        identities.len() == paths.len(),
        "输入、候选输出和计划输出必须使用不同路径"
    );
    Ok(())
}

fn normalized_identity(path: &Path) -> Result<String> {
    let absolute = if path.exists() {
        strip_windows_verbatim_prefix(
            fs::canonicalize(path).with_context(|| format!("规范化路径 {}", path.display()))?,
        )
    } else {
        normalize_new_path(path)?
    };
    let mut value = absolute.to_string_lossy().replace('/', "\\");
    if cfg!(windows) {
        value.make_ascii_lowercase();
    }
    Ok(value)
}

fn normalize_new_path(path: &Path) -> Result<PathBuf> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let name = path.file_name().context("输出路径必须包含文件名")?;
    let absolute_parent = if parent.exists() {
        strip_windows_verbatim_prefix(
            fs::canonicalize(parent)
                .with_context(|| format!("规范化输出目录 {}", parent.display()))?,
        )
    } else {
        std::env::current_dir()
            .context("获取当前目录")?
            .join(parent)
    };
    Ok(absolute_parent.join(name))
}

fn strip_windows_verbatim_prefix(path: PathBuf) -> PathBuf {
    #[cfg(windows)]
    {
        let value = path.to_string_lossy();
        if let Some(rest) = value.strip_prefix("\\\\?\\") {
            if let Some(unc) = rest.strip_prefix("UNC\\") {
                return PathBuf::from(format!("\\\\{unc}"));
            }
            return PathBuf::from(rest);
        }
    }
    path
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut file =
        File::open(path).with_context(|| format!("打开输入工作簿 {}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = [0u8; 64 * 1024];
    loop {
        let read = file
            .read(&mut buffer)
            .with_context(|| format!("读取输入工作簿 {}", path.display()))?;
        if read == 0 {
            break;
        }
        digest.update(&buffer[..read]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

fn find_explanation_sheet(sheets: Vec<SheetValues>) -> Result<ValueGrid> {
    let grids = sheets.into_iter().map(ValueGrid::from).collect::<Vec<_>>();
    if let Some(sheet) = grids.iter().find(|sheet| sheet.name == "说明文件") {
        return Ok(sheet.clone());
    }
    grids
        .into_iter()
        .find(ValueGrid::has_institution_code_header)
        .context("未找到说明文件工作表或“机构代码”表头")
}

impl ValueGrid {
    fn has_institution_code_header(&self) -> bool {
        self.has_header_text("机构代码")
    }
}

fn identity(sheet: &ValueGrid, row: usize, headers: &HeaderColumns) -> Identity {
    (
        code_value(&sheet.cell(row, headers.indexes["机构代码"])),
        code_value(&sheet.cell(row, headers.indexes["指标代码"])),
        amount_value(&sheet.cell(row, headers.indexes["数据值"])),
        amount_value(&sheet.cell(row, headers.indexes["增减额"])),
    )
}

fn identity_is_nonempty(key: &Identity) -> bool {
    key.0.len() > 0 || key.1.len() > 0 || key.2.len() > 0 || key.3.len() > 0
}

fn code_value(value: &CellValue) -> String {
    let mut text = python_string(value).trim().replace('\'', "");
    if let Some(digits) = text.strip_suffix(".0") {
        if !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()) {
            text = digits.to_owned();
        }
    }
    text.to_uppercase()
}

fn amount_value(value: &CellValue) -> String {
    if matches!(value, CellValue::Empty) {
        return String::new();
    }
    let text = python_string(value).trim().to_owned();
    match quantize_half_up_four_places(&text) {
        DecimalQuantization::Rounded(value) => format!("#{value}"),
        DecimalQuantization::Invalid(value) => value,
        DecimalQuantization::NotDecimal => text,
    }
}

enum DecimalQuantization {
    Rounded(String),
    Invalid(String),
    NotDecimal,
}

fn quantize_half_up_four_places(text: &str) -> DecimalQuantization {
    if text.eq_ignore_ascii_case("nan") || text.eq_ignore_ascii_case("+nan") {
        return DecimalQuantization::Rounded("NaN".to_owned());
    }
    if text.eq_ignore_ascii_case("-nan") {
        return DecimalQuantization::Rounded("-NaN".to_owned());
    }
    if text.eq_ignore_ascii_case("infinity") || text.eq_ignore_ascii_case("+infinity") {
        return DecimalQuantization::Invalid("Infinity".to_owned());
    }
    if text.eq_ignore_ascii_case("-infinity") {
        return DecimalQuantization::Invalid("-Infinity".to_owned());
    }

    let (negative, unsigned) = if let Some(rest) = text.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = text.strip_prefix('+') {
        (false, rest)
    } else {
        (false, text)
    };
    if unsigned.is_empty() {
        return DecimalQuantization::NotDecimal;
    }
    let mut exponent_parts = unsigned.split(['e', 'E']);
    let mantissa = exponent_parts.next().unwrap_or_default();
    let exponent_text = exponent_parts.next();
    if exponent_parts.next().is_some() {
        return DecimalQuantization::NotDecimal;
    }
    let exponent = match exponent_text {
        Some(value) => match value.parse::<i32>() {
            Ok(value) => value,
            Err(_) => return DecimalQuantization::NotDecimal,
        },
        None => 0,
    };
    let mut decimal_parts = mantissa.split('.');
    let integer_part = decimal_parts.next().unwrap_or_default();
    let fractional_part = decimal_parts.next().unwrap_or_default();
    if decimal_parts.next().is_some()
        || (!integer_part.bytes().all(|byte| byte.is_ascii_digit()))
        || (!fractional_part.bytes().all(|byte| byte.is_ascii_digit()))
        || (integer_part.is_empty() && fractional_part.is_empty())
    {
        return DecimalQuantization::NotDecimal;
    }
    let coefficient = format!("{integer_part}{fractional_part}");
    let digits = coefficient.trim_start_matches('0');
    if digits.is_empty() {
        return DecimalQuantization::Rounded(if negative {
            "-0.0000".to_owned()
        } else {
            "0.0000".to_owned()
        });
    }

    let shift = exponent - fractional_part.len() as i32 + 4;
    if shift > 10_000 || shift < -10_000 {
        return DecimalQuantization::Invalid(decimal_to_string(text));
    }
    let mut scaled = if shift >= 0 {
        format!("{digits}{}", "0".repeat(shift as usize))
    } else {
        let discarded = (-shift) as usize;
        let keep = digits.len().saturating_sub(discarded);
        let mut retained = if keep == 0 {
            "0".to_owned()
        } else {
            digits[..keep].to_owned()
        };
        let first_discarded = if keep < digits.len() && discarded <= digits.len() {
            digits.as_bytes()[keep]
        } else {
            b'0'
        };
        if first_discarded >= b'5' {
            increment_decimal_digits(&mut retained);
        }
        retained
    };
    let precision = scaled.trim_start_matches('0').len().max(1);
    if precision > 28 {
        return DecimalQuantization::Invalid(decimal_to_string(text));
    }
    if scaled.len() <= 4 {
        scaled = format!("{}{}", "0".repeat(5 - scaled.len()), scaled);
    }
    let split = scaled.len() - 4;
    let integer = scaled[..split].trim_start_matches('0');
    let integer = if integer.is_empty() { "0" } else { integer };
    let sign = if negative { "-" } else { "" };
    DecimalQuantization::Rounded(format!("{sign}{integer}.{}", &scaled[split..]))
}

fn increment_decimal_digits(value: &mut String) {
    let mut bytes = value.as_bytes().to_vec();
    let mut index = bytes.len();
    while index > 0 {
        index -= 1;
        if bytes[index] < b'9' {
            bytes[index] += 1;
            *value = String::from_utf8(bytes).expect("decimal digits remain UTF-8");
            return;
        }
        bytes[index] = b'0';
    }
    bytes.insert(0, b'1');
    *value = String::from_utf8(bytes).expect("decimal digits remain UTF-8");
}

fn decimal_to_string(text: &str) -> String {
    if let Some(index) = text.find(['e', 'E']) {
        let (mantissa, exponent) = text.split_at(index);
        let exponent = exponent[1..].parse::<i32>().unwrap_or(0);
        let sign = if exponent >= 0 { "+" } else { "" };
        format!("{}E{sign}{exponent}", mantissa.trim_start_matches('+'))
    } else {
        text.to_owned()
    }
}

fn python_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => python_float_string(*value),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeIso(value) => value.replace('T', " "),
        CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => format_excel_error(value).to_owned(),
    }
}

fn python_str_for_conflict(value: &CellValue) -> String {
    if matches!(value, CellValue::Empty) {
        "None".to_owned()
    } else {
        python_string(value)
    }
}

fn python_float_string(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value == f64::INFINITY {
        return "inf".to_owned();
    }
    if value == f64::NEG_INFINITY {
        return "-inf".to_owned();
    }
    let value = format!("{value:?}");
    let Some((mantissa, exponent)) = value.split_once('e') else {
        return value;
    };
    let mantissa = mantissa.strip_suffix(".0").unwrap_or(mantissa);
    let exponent = exponent.parse::<i32>().unwrap_or(0);
    let sign = if exponent >= 0 { "+" } else { "" };
    let exponent_digits = exponent.abs().to_string();
    let exponent_digits = if exponent_digits.len() < 2 {
        format!("0{exponent_digits}")
    } else {
        exponent_digits
    };
    format!(
        "{mantissa}e{sign}{}{exponent_digits}",
        if exponent < 0 { "-" } else { "" }
    )
}

fn python_truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => false,
        CellValue::Text(value) => !value.is_empty(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => *value != 0.0,
        CellValue::Integer(value) => *value != 0,
        CellValue::Boolean(value) => *value,
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => !value.is_empty(),
        CellValue::Error(value) => !value.is_empty(),
    }
}

fn format_excel_error(value: &str) -> &str {
    let normalized = value
        .trim()
        .trim_start_matches('#')
        .trim_end_matches(['!', '?'])
        .replace('_', "")
        .replace('-', "")
        .replace(' ', "")
        .to_ascii_lowercase();
    match normalized.as_str() {
        "ref" => "#REF!",
        "value" => "#VALUE!",
        "na" | "notavailable" => "#N/A",
        "div0" | "divzero" => "#DIV/0!",
        "name" => "#NAME?",
        "num" => "#NUM!",
        "null" => "#NULL!",
        "spill" => "#SPILL!",
        "calc" => "#CALC!",
        _ => value,
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let bytes = address.as_bytes();
    let split = bytes.iter().position(|byte| byte.is_ascii_digit())?;
    if split == 0
        || !bytes[..split].iter().all(u8::is_ascii_alphabetic)
        || !bytes[split..].iter().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = bytes[..split].iter().try_fold(0usize, |column, byte| {
        column
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    let row = address[split..].parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut label = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        label.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{label}{row}")
}

fn required_preservation() -> BTreeSet<PreservationRequirement> {
    [
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::UnmodifiedPackageParts,
        PreservationRequirement::SheetIdentityAndOrder,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ]
    .into_iter()
    .collect()
}

fn conflict_report_path(target_path: &Path) -> Result<PathBuf> {
    let parent = target_path
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = target_path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("S3-F03");
    Ok(parent.join(format!(
        "{stem}_导入冲突清单_{}.xlsx",
        Local::now().format("%Y%m%d_%H%M%S")
    )))
}

fn write_conflict_workbook(path: &Path, rows: &[ConflictRow]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet
        .set_name("冲突明细")
        .context("设置 S3-F03 冲突表名称")?;
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0xC0_00_00));
    for (column, header) in CONFLICT_HEADERS.iter().enumerate() {
        sheet.write_string_with_format(0, column as u16, *header, &header_format)?;
    }
    for (index, row) in rows.iter().enumerate() {
        for (column, value) in row.0.iter().enumerate() {
            sheet.write_string((index + 1) as u32, column as u16, value)?;
        }
    }
    sheet
        .set_freeze_panes(1, 0)
        .context("冻结 S3-F03 冲突表表头")?;
    sheet
        .autofilter(0, 0, rows.len() as u32, 9)
        .context("设置 S3-F03 冲突表筛选")?;
    for (column, width) in CONFLICT_WIDTHS.iter().enumerate() {
        sheet
            .set_column_width_pixels(column as u16, (width * 7.0).round() as u32)
            .with_context(|| format!("设置 S3-F03 冲突清单第 {} 列宽", column + 1))?;
    }
    workbook.save(path).context("保存 S3-F03 冲突清单")
}

fn reserve_temporary_path(output: &Path, extension: &str) -> Result<PathBuf> {
    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent).with_context(|| format!("创建输出目录：{}", parent.display()))?;
    let stem = output
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("s3-f03");
    for _ in 0..100 {
        let sequence = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".{stem}.{}.{}.tmp.{extension}",
            std::process::id(),
            sequence
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(_) => return Ok(candidate),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| format!("创建临时输出 {}", candidate.display()))
            }
        }
    }
    bail!("无法为输出保留唯一临时路径：{}", output.display())
}

fn publish_without_overwrite(temporary: &Path, output: &Path) -> Result<()> {
    fs::hard_link(temporary, output).with_context(|| {
        format!(
            "安全发布输出失败（目标已存在或文件系统不支持硬链接）：{}",
            output.display()
        )
    })?;
    fs::remove_file(temporary).with_context(|| format!("清理临时文件：{}", temporary.display()))
}

#[cfg(test)]
mod tests {
    use super::{
        amount_value, code_value, create_feedback_backup, execute_explanation_feedback_import,
        execute_explanation_feedback_import_to_copy, publish_candidate_in_place,
        quantize_half_up_four_places, required_preservation, sha256_file, DecimalQuantization,
        S3FeedbackStagingDirectory,
    };
    use anyhow::Result;
    use excel_editor::{
        BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperationKind, EngineId,
        ExcelEditPlan, ExcelModifier, OutputValidator, UmyaXlsxBackend, UmyaXlsxValidator,
        WorkbookFormat,
    };
    use excel_reader::CellValue;
    use excel_writer::Workbook;
    use std::{
        fs,
        path::Path,
        sync::Arc,
        time::{SystemTime, UNIX_EPOCH},
    };

    struct AcceptAllValidator;

    impl OutputValidator for AcceptAllValidator {
        fn validate(&self, _plan: &ExcelEditPlan, _candidate_output: &Path) -> Result<()> {
            Ok(())
        }
    }

    fn write_explanation(path: &Path, reason: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("说明文件").unwrap();
        for (column, header) in [
            "机构代码",
            "机构名称",
            "指标代码",
            "指标名称",
            "数据值",
            "增减额(亿元)",
            "变动原因",
        ]
        .iter()
        .enumerate()
        {
            sheet.write_string(0, column as u16, *header).unwrap();
        }
        for (column, value) in ["0001", "测试机构", "I001", "测试指标"].iter().enumerate() {
            sheet.write_string(1, column as u16, *value).unwrap();
        }
        sheet.write_number(1, 4, 100.0).unwrap();
        sheet.write_number(1, 5, 20.0).unwrap();
        sheet.write_string(1, 6, reason).unwrap();
        workbook.save(path).unwrap();
    }

    #[test]
    fn s3_feedback_identity_normalizes_codes_and_half_up_amounts() {
        assert_eq!(
            code_value(&CellValue::Text(" '6k0u.0 ".to_owned())),
            "6K0U.0"
        );
        assert_eq!(code_value(&CellValue::Number(7020.0)), "7020");
        assert_eq!(amount_value(&CellValue::Number(0.418462)), "#0.4185");
        assert_eq!(
            amount_value(&CellValue::Text("-0.00005".to_owned())),
            "#-0.0001"
        );
        assert_eq!(
            amount_value(&CellValue::Text("0.000049".to_owned())),
            "#0.0000"
        );
        assert_eq!(
            amount_value(&CellValue::Text("-0.000049".to_owned())),
            "#-0.0000"
        );
        assert_eq!(
            amount_value(&CellValue::Text("description".to_owned())),
            "description"
        );
    }

    #[test]
    fn decimal_quantization_handles_leading_zeroes_and_exponents() {
        assert!(matches!(
            quantize_half_up_four_places("5e-5"),
            DecimalQuantization::Rounded(value) if value == "0.0001"
        ));
        assert!(matches!(
            quantize_half_up_four_places("5e-6"),
            DecimalQuantization::Rounded(value) if value == "0.0000"
        ));
        assert!(matches!(
            quantize_half_up_four_places("1.23455e2"),
            DecimalQuantization::Rounded(value) if value == "123.4550"
        ));
    }

    #[test]
    fn execution_without_a_passed_route_keeps_inputs_and_publishes_no_workbook() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f3_execution_{suffix}"));
        let output = root.join("output");
        fs::create_dir_all(&output).unwrap();
        let target = root.join("说明文件.xlsx");
        let feedback = root.join("机构反馈.xlsx");
        write_explanation(&target, "");
        write_explanation(&feedback, "核实后调整");
        let before = [
            sha256_file(&target).unwrap(),
            sha256_file(&feedback).unwrap(),
        ];
        let modifier =
            ExcelModifier::new("test-platform", Vec::new(), Arc::new(AcceptAllValidator)).unwrap();

        let error = execute_explanation_feedback_import(&target, &[feedback.clone()], &modifier)
            .expect_err("S3-F03 must not apply without a Passed Excel-Core route");
        assert!(format!("{error:#}").contains("NoEligibleBackend"));
        assert_eq!(
            before,
            [
                sha256_file(&target).unwrap(),
                sha256_file(&feedback).unwrap()
            ]
        );
        assert!(fs::read_dir(&output).unwrap().next().is_none());

        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn successful_execution_writes_feedback_to_a_new_workbook_and_keeps_sources_unchanged(
    ) -> Result<()> {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f3_success_{suffix}"));
        let output = root.join("output");
        fs::create_dir_all(&output)?;
        let target = root.join("说明文件.xlsx");
        let feedback = root.join("机构反馈.xlsx");
        write_explanation(&target, "");
        write_explanation(&feedback, "核实后调整");
        let before = [sha256_file(&target)?, sha256_file(&feedback)?];

        // The synthetic Passed capability exists only to exercise ExcelModifier's
        // successful route in this focused test. It is not production evidence.
        let platform = format!("{}-{}", std::env::consts::OS, std::env::consts::ARCH);
        let profile = excel_editor::inspect_xlsx_profile(&target)?;
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Passed,
            format: WorkbookFormat::Xlsx,
            operations: [EditOperationKind::SetCellValue].into_iter().collect(),
            supported_features: profile.features_present.clone(),
            preserved: required_preservation(),
            platform: platform.clone(),
            evidence: Some(CapabilityEvidence {
                sample_id: "synthetic-s3-f03-feedback-test".to_owned(),
                report_path: "test-only://s3-f03-feedback-execution".to_owned(),
                adapter_version: "test fixture".to_owned(),
                engine_version: "test fixture".to_owned(),
                platform,
            }),
        };
        let backend = Arc::new(UmyaXlsxBackend::new(capability)?);
        let modifier = ExcelModifier::new(
            format!("{}-{}", std::env::consts::OS, std::env::consts::ARCH),
            vec![backend],
            Arc::new(UmyaXlsxValidator),
        )?;

        let report = execute_explanation_feedback_import_to_copy(
            &target,
            &[feedback.clone()],
            &output,
            &modifier,
        )?;

        assert_eq!(report.imported, 1);
        assert_eq!(report.planned_cell_updates, 1);
        assert_eq!(report.conflicts, 0);
        assert!(report.output_workbook.is_file());
        assert_ne!(report.output_workbook, target);
        assert_eq!(
            before,
            [sha256_file(&target)?, sha256_file(&feedback)?],
            "the target explanation and institution feedback workbooks must remain unchanged"
        );
        let workbook = excel_editor::open_xlsx(&report.output_workbook)?;
        let sheet = workbook.sheet_by_name("说明文件")?;
        assert_eq!(
            sheet.cell("G2").map(|cell| cell.value().to_string()),
            Some("核实后调整".to_owned())
        );

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn successful_execution_backs_up_then_updates_the_target_in_place() -> Result<()> {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f3_in_place_{suffix}"));
        fs::create_dir_all(&root)?;
        let target = root.join("说明文件.xlsx");
        let feedback = root.join("机构反馈.xlsx");
        write_explanation(&target, "");
        write_explanation(&feedback, "核实后调整");
        let target_hash_before = sha256_file(&target)?;
        let feedback_hash_before = sha256_file(&feedback)?;

        let platform = format!("{}-{}", std::env::consts::OS, std::env::consts::ARCH);
        let profile = excel_editor::inspect_xlsx_profile(&target)?;
        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Passed,
            format: WorkbookFormat::Xlsx,
            operations: [EditOperationKind::SetCellValue].into_iter().collect(),
            supported_features: profile.features_present.clone(),
            preserved: required_preservation(),
            platform: platform.clone(),
            evidence: Some(CapabilityEvidence {
                sample_id: "synthetic-s3-f03-in-place-test".to_owned(),
                report_path: "test-only://s3-f03-in-place".to_owned(),
                adapter_version: "test fixture".to_owned(),
                engine_version: "test fixture".to_owned(),
                platform: platform.clone(),
            }),
        };
        let backend = Arc::new(UmyaXlsxBackend::new(capability)?);
        let modifier = ExcelModifier::new(platform, vec![backend], Arc::new(UmyaXlsxValidator))?;

        let report = execute_explanation_feedback_import(&target, &[feedback.clone()], &modifier)?;

        assert_eq!(report.output_workbook, target);
        assert_eq!(report.imported, 1);
        assert_eq!(report.planned_cell_updates, 1);
        assert_eq!(report.conflicts, 0);
        assert_ne!(sha256_file(&target)?, target_hash_before);
        assert_eq!(sha256_file(&feedback)?, feedback_hash_before);
        let backup = report.backup_workbook.expect("pre-write target backup");
        assert!(backup.is_file());
        assert_eq!(sha256_file(&backup)?, target_hash_before);
        let workbook = excel_editor::open_xlsx(&report.output_workbook)?;
        let sheet = workbook.sheet_by_name("说明文件")?;
        assert_eq!(
            sheet.cell("G2").map(|cell| cell.value().to_string()),
            Some("核实后调整".to_owned())
        );

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn conflicting_execution_publishes_conflict_report_and_preserves_target() -> Result<()> {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f3_conflicts_{suffix}"));
        fs::create_dir_all(&root)?;
        let target = root.join("说明文件.xlsx");
        let feedback_a = root.join("机构反馈甲.xlsx");
        let feedback_b = root.join("机构反馈乙.xlsx");
        write_explanation(&target, "待机构填写");
        write_explanation(&feedback_a, "原因甲");
        write_explanation(&feedback_b, "原因乙");
        let target_hash_before = sha256_file(&target)?;
        let feedback_hashes_before = [sha256_file(&feedback_a)?, sha256_file(&feedback_b)?];
        let modifier =
            ExcelModifier::new("test-platform", Vec::new(), Arc::new(AcceptAllValidator))?;

        let report = execute_explanation_feedback_import(
            &target,
            &[feedback_a.clone(), feedback_b.clone()],
            &modifier,
        )?;

        assert_eq!(report.imported, 0);
        assert_eq!(report.conflicts, 1);
        assert_eq!(report.planned_cell_updates, 0);
        assert!(report.conflict_write_time_ms <= report.total_time_ms);
        assert_eq!(sha256_file(&target)?, target_hash_before);
        assert_eq!(
            [sha256_file(&feedback_a)?, sha256_file(&feedback_b)?],
            feedback_hashes_before
        );
        let backup = report.backup_workbook.expect("pre-write target backup");
        assert_eq!(sha256_file(&backup)?, target_hash_before);
        let conflict_path = report
            .conflict_report_path
            .expect("conflict workbook should be published");
        assert!(conflict_path.is_file());
        let conflict_book = excel_editor::open_xlsx(&conflict_path)?;
        let sheet = conflict_book.sheet_by_name("冲突明细")?;
        assert_eq!(
            sheet.cell("A2").map(|cell| cell.value().to_string()),
            Some("冲突".to_owned())
        );
        let reasons = sheet
            .cell("I2")
            .map(|cell| cell.value().to_string())
            .expect("conflict report should include feedback reasons");
        assert!(reasons.contains("原因甲"));
        assert!(reasons.contains("原因乙"));
        assert_eq!(
            sheet.cell("J2").map(|cell| cell.value().to_string()),
            Some("多个反馈文件给出不同变动原因".to_owned())
        );

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn failed_candidate_install_restores_the_target_and_retains_its_backup() -> Result<()> {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f3_rollback_{suffix}"));
        fs::create_dir_all(&root)?;
        let target = root.join("说明文件.xlsx");
        fs::write(&target, b"original target bytes")?;
        let original = fs::read(&target)?;
        let permissions = fs::metadata(&target)?.permissions();
        let backup = create_feedback_backup(&target, &permissions)?;
        let mut staging = S3FeedbackStagingDirectory::create(&root)?;
        let rollback = staging.path().join("说明文件.xlsx");
        let missing_candidate = staging.path().join("missing-candidate.xlsx");

        let error = publish_candidate_in_place(
            &missing_candidate,
            &target,
            &rollback,
            &backup,
            &mut staging,
        )
        .expect_err("missing staged output must abort publication and roll back");

        assert!(format!("{error:#}").contains("原目标文件已恢复"));
        assert_eq!(fs::read(&target)?, original);
        assert_eq!(fs::read(&backup)?, original);
        assert!(!rollback.exists());
        fs::remove_dir_all(root)?;
        Ok(())
    }
}
