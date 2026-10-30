//! S1 conditional-format extraction from evaluator results into Audit issues.
//!
//! This is the business-output stage for the existing static Rust evaluator.
//! Template configuration loading, the remaining S1 rule engine, history, and
//! workbook writeback remain separate migration work.

use crate::{
    conditional_format::{
        evaluate_xlsx_workbook_with_dump, ConditionalDiagnostic, ConditionalEvaluationReport,
        ConditionalTrigger, EvaluationStatus,
    },
    models::{CopyRange, Issue},
    template::normalize_template_name,
};
use anyhow::{bail, Result};
use excel_reader::{CellValue, SheetDump};
use rayon::prelude::*;
use serde::Serialize;
use serde_json::{json, Value};
use std::{
    collections::BTreeSet,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Clone, Debug, Default)]
pub struct S1ConditionalMetadata {
    pub period: String,
    pub batch_id: String,
    pub audit_time: String,
    pub org_code: String,
    pub org_name: String,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct S1ConditionalContextSummary {
    pub unresolved_indicator_count: usize,
    pub comments_unverified_sheet_count: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ConditionalAuditResult {
    pub status: &'static str,
    pub issues: Vec<Issue>,
    pub evaluation: ConditionalEvaluationReport,
    pub context: S1ConditionalContextSummary,
    pub issue_build_time_ms: u128,
}

#[derive(Clone, Debug)]
pub struct S1ConditionalBatchInput {
    pub workbook_path: PathBuf,
    pub source_file: PathBuf,
    pub metadata: S1ConditionalMetadata,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ConditionalBatchFileSummary {
    pub source_file: String,
    pub status: &'static str,
    pub error: Option<String>,
    pub issue_count: usize,
    pub reader: &'static str,
    pub rules_total: usize,
    pub rules_supported: usize,
    pub scanned_cells: usize,
    pub diagnostics: Vec<ConditionalDiagnostic>,
    pub unresolved_indicator_count: usize,
    pub comments_unverified_sheet_count: usize,
    pub read_time_ms: u128,
    pub evaluation_time_ms: u128,
    pub issue_build_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ConditionalBatchAuditResult {
    pub status: &'static str,
    pub issues: Vec<Issue>,
    pub files: Vec<S1ConditionalBatchFileSummary>,
    pub total_elapsed_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ConditionalBatchProgress {
    pub status: &'static str,
    pub file_index: usize,
    pub files_total: usize,
    pub source_file: String,
    pub started_at: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub finished_at: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub elapsed_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reader: Option<&'static str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rules_total: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rules_supported: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scanned_cells: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub issue_count: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagnostics: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub read_time_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub evaluation_time_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub issue_build_time_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
}

/// Run the static conditional-format evaluator and construct the S1-F01 issue
/// rows that its Python counterpart feeds into AuditResult/history processing.
/// The workbook is read once; no recalculation or workbook mutation occurs.
pub fn extract_conditional_issues(
    workbook_path: impl AsRef<Path>,
    source_file: impl AsRef<Path>,
    scan_areas: &[String],
    structure_ranges: &[CopyRange],
    metadata: &S1ConditionalMetadata,
) -> Result<S1ConditionalAuditResult> {
    let workbook_path = workbook_path.as_ref();
    let source_file = source_file.as_ref();
    let (mut evaluation, workbook) = evaluate_xlsx_workbook_with_dump(workbook_path, scan_areas)?;
    let build_started = Instant::now();
    let source_stem = source_file
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let normalized_source = normalize_template_name(&source_stem);
    let workbook_key = if normalized_source.is_empty() {
        source_stem
    } else {
        normalized_source
    };

    let mut issues = Vec::with_capacity(evaluation.triggered.len());
    let mut issue_build_diagnostics = Vec::new();
    let mut unresolved_indicator_count = 0usize;
    let mut comments_unverified_sheets = BTreeSet::new();
    for trigger in &evaluation.triggered {
        let Some(sheet) = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == trigger.sheet)
        else {
            issue_build_diagnostics.push(issue_build_diagnostic(
                trigger,
                format!("条件格式结果引用了未知工作表：{}", trigger.sheet),
            ));
            continue;
        };
        let (column, row) = match parse_cell_address(&trigger.cell) {
            Some(position) => position,
            None => {
                issue_build_diagnostics.push(issue_build_diagnostic(
                    trigger,
                    format!("无法解析条件格式触发单元格：{}", trigger.cell),
                ));
                continue;
            }
        };
        let check_field = resolve_indicator(sheet, structure_ranges, row, column);
        if check_field.is_empty() {
            unresolved_indicator_count += 1;
        }
        let comment_supported = sheet
            .comments_status
            .starts_with("read_by_namespace_aware_ooxml_supplement");
        if !comment_supported {
            comments_unverified_sheets.insert(sheet.name.clone());
        }
        let check_type = if comment_supported {
            sheet
                .comments
                .get(&trigger.cell)
                .map(|comment| comment.trim().to_owned())
                .unwrap_or_default()
        } else {
            String::new()
        };
        let target_value = match issue_value(sheet.values.get(&trigger.cell)) {
            Ok(value) => value,
            Err(error) => {
                issue_build_diagnostics.push(issue_build_diagnostic(
                    trigger,
                    format!("无法读取条件格式触发单元格的审核值：{error:#}"),
                ));
                continue;
            }
        };

        let identity = [
            workbook_key.as_str(),
            trigger.sheet.as_str(),
            trigger.cell.as_str(),
            if check_field.is_empty() {
                "校验指标未识别"
            } else {
                check_field.as_str()
            },
        ]
        .join("｜");
        issues.push(Issue {
            issue_id: identity,
            period: metadata.period.clone(),
            batch_id: metadata.batch_id.clone(),
            audit_time: metadata.audit_time.clone(),
            triggered: true,
            status: String::new(),
            first_seen_period: String::new(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: metadata.org_code.clone(),
            org_name: metadata.org_name.clone(),
            report_code: String::new(),
            sheet_name: trigger.sheet.clone(),
            rule_id: "条件格式填充".to_owned(),
            severity: "条件格式触发".to_owned(),
            formula_cell: trigger.cell.clone(),
            target_cell: trigger.cell.clone(),
            target_value,
            formula_result: json!(""),
            message: trigger.message.clone(),
            source_file: source_file.display().to_string(),
            audit_file: workbook_path.display().to_string(),
            check_field,
            check_type: check_type.clone(),
            comparison_value: json!(""),
            reference_value: json!(""),
            difference_value: json!(""),
            detail: trigger.message.clone(),
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        });
    }
    evaluation.diagnostics.extend(issue_build_diagnostics);

    let issue_build_time_ms = build_started.elapsed().as_millis();
    let context = S1ConditionalContextSummary {
        unresolved_indicator_count,
        comments_unverified_sheet_count: comments_unverified_sheets.len(),
    };
    let has_partial_context = context.unresolved_indicator_count > 0
        || context.comments_unverified_sheet_count > 0
        || !evaluation
            .read_diagnostics
            .normalized_error_cells
            .is_empty();
    let status = if evaluation.diagnostics.is_empty() && !has_partial_context {
        "completed"
    } else {
        "completed_with_diagnostics"
    };
    Ok(S1ConditionalAuditResult {
        status,
        issues,
        evaluation,
        context,
        issue_build_time_ms,
    })
}

fn issue_build_diagnostic(trigger: &ConditionalTrigger, reason: String) -> ConditionalDiagnostic {
    ConditionalDiagnostic {
        sheet: trigger.sheet.clone(),
        cell: trigger.cell.clone(),
        applies_to: trigger.applies_to.clone(),
        rule_type: trigger.rule_type.clone(),
        formula: trigger.formula.clone(),
        priority: trigger.priority,
        status: EvaluationStatus::Error,
        reason: format!("条件格式触发结果未能转换为审核异常；已跳过该单元格：{reason}"),
    }
}

/// Run the same conditional-format issue stage for a selected set of reports,
/// preserving input order and per-file institution metadata for the S1 batch.
pub fn extract_conditional_issues_batch<F>(
    inputs: &[S1ConditionalBatchInput],
    scan_areas: &[String],
    structure_ranges: &[CopyRange],
    on_progress: F,
) -> Result<S1ConditionalBatchAuditResult>
where
    F: Fn(&S1ConditionalBatchProgress) + Send + Sync,
{
    if inputs.is_empty() {
        bail!("请至少指定一个待审核工作簿");
    }

    let batch_started = Instant::now();
    let reports = inputs
        .par_iter()
        .enumerate()
        .map(|(offset, input)| {
            let file_index = offset + 1;
            let source_name = path_basename(&input.source_file);
            let started_at = chrono::Local::now().to_rfc3339();
            let file_started = Instant::now();
            on_progress(&S1ConditionalBatchProgress {
                status: "started",
                file_index,
                files_total: inputs.len(),
                source_file: source_name.clone(),
                started_at: started_at.clone(),
                finished_at: None,
                elapsed_ms: None,
                reader: None,
                rules_total: None,
                rules_supported: None,
                scanned_cells: None,
                issue_count: None,
                diagnostics: None,
                read_time_ms: None,
                evaluation_time_ms: None,
                issue_build_time_ms: None,
                error: None,
            });

            let report = match extract_conditional_issues(
                &input.workbook_path,
                &input.source_file,
                scan_areas,
                structure_ranges,
                &input.metadata,
            ) {
                Ok(report) => report,
                Err(error) => {
                    on_progress(&S1ConditionalBatchProgress {
                        status: "failed",
                        file_index,
                        files_total: inputs.len(),
                        source_file: source_name,
                        started_at,
                        finished_at: Some(chrono::Local::now().to_rfc3339()),
                        elapsed_ms: Some(file_started.elapsed().as_millis()),
                        reader: None,
                        rules_total: None,
                        rules_supported: None,
                        scanned_cells: None,
                        issue_count: None,
                        diagnostics: None,
                        read_time_ms: None,
                        evaluation_time_ms: None,
                        issue_build_time_ms: None,
                        error: Some(format!("{error:#}")),
                    });
                    return Err(format!(
                        "S1 条件格式审核失败：{}：{error:#}",
                        input.workbook_path.display()
                    ));
                }
            };

            let issue_count = report.issues.len();
            let diagnostics_count = report.evaluation.diagnostics.len();
            on_progress(&S1ConditionalBatchProgress {
                status: "completed",
                file_index,
                files_total: inputs.len(),
                source_file: source_name.clone(),
                started_at,
                finished_at: Some(chrono::Local::now().to_rfc3339()),
                elapsed_ms: Some(file_started.elapsed().as_millis()),
                reader: Some(report.evaluation.reader),
                rules_total: Some(report.evaluation.rules_total),
                rules_supported: Some(report.evaluation.rules_supported),
                scanned_cells: Some(report.evaluation.scanned_cells),
                issue_count: Some(issue_count),
                diagnostics: Some(diagnostics_count),
                read_time_ms: Some(report.evaluation.read_time_ms),
                evaluation_time_ms: Some(report.evaluation.evaluation_time_ms),
                issue_build_time_ms: Some(report.issue_build_time_ms),
                error: None,
            });
            Ok((source_name, report))
        })
        .collect::<Vec<_>>();

    let mut issues = Vec::new();
    let mut files = Vec::with_capacity(reports.len());
    let mut has_diagnostics = false;

    for result in reports {
        match result {
            Ok((source_name, report)) => {
                let total_time_ms = report.evaluation.total_time_ms + report.issue_build_time_ms;
                has_diagnostics |= report.status == "completed_with_diagnostics";
                let issue_count = report.issues.len();
                files.push(S1ConditionalBatchFileSummary {
                    source_file: source_name,
                    status: report.status,
                    error: None,
                    issue_count,
                    reader: report.evaluation.reader,
                    rules_total: report.evaluation.rules_total,
                    rules_supported: report.evaluation.rules_supported,
                    scanned_cells: report.evaluation.scanned_cells,
                    diagnostics: report.evaluation.diagnostics,
                    unresolved_indicator_count: report.context.unresolved_indicator_count,
                    comments_unverified_sheet_count: report.context.comments_unverified_sheet_count,
                    read_time_ms: report.evaluation.read_time_ms,
                    evaluation_time_ms: report.evaluation.evaluation_time_ms,
                    issue_build_time_ms: report.issue_build_time_ms,
                    total_time_ms,
                });
                issues.extend(report.issues);
            }
            Err(error) => {
                has_diagnostics = true;
                let source_file = error
                    .split_once("：")
                    .and_then(|(_, remainder)| remainder.split_once("："))
                    .map(|(path, _)| path.rsplit(['\\', '/']).next().unwrap_or(path))
                    .unwrap_or("unknown");
                files.push(S1ConditionalBatchFileSummary {
                    source_file: source_file.to_owned(),
                    status: "failed",
                    error: Some(error),
                    issue_count: 0,
                    reader: "unavailable",
                    rules_total: 0,
                    rules_supported: 0,
                    scanned_cells: 0,
                    diagnostics: Vec::new(),
                    unresolved_indicator_count: 0,
                    comments_unverified_sheet_count: 0,
                    read_time_ms: 0,
                    evaluation_time_ms: 0,
                    issue_build_time_ms: 0,
                    total_time_ms: 0,
                });
            }
        }
    }

    Ok(S1ConditionalBatchAuditResult {
        status: if has_diagnostics {
            "completed_with_diagnostics"
        } else {
            "completed"
        },
        issues,
        files,
        total_elapsed_ms: batch_started.elapsed().as_millis(),
    })
}

fn path_basename(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}

fn issue_value(raw: Option<&Value>) -> Result<Value> {
    let empty = json!({"type":"empty"});
    let value = CellValue::from_probe_json(raw.unwrap_or(&empty))?;
    Ok(match value {
        CellValue::Empty => Value::Null,
        CellValue::Text(value) | CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => {
            json!(value)
        }
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => json!(value),
        CellValue::Integer(value) => json!(value),
        CellValue::Boolean(value) => json!(value),
        CellValue::Error(value) => json!(value),
    })
}

fn resolve_indicator(
    sheet: &SheetDump,
    structure_ranges: &[CopyRange],
    row: i32,
    column: i32,
) -> String {
    let bounds = structure_ranges
        .iter()
        .filter(|area| area.sheet_name == sheet.name)
        .filter_map(|area| parse_cell_range(&area.address))
        .collect::<Vec<_>>();
    let Some(min_row) = bounds.iter().map(|item| item.1).min() else {
        return String::new();
    };
    let Some(min_column) = bounds.iter().map(|item| item.0).min() else {
        return String::new();
    };

    let mut left = Vec::new();
    for current_column in min_column..column {
        if let Some(text) = display_value(sheet, row, current_column) {
            if !left.contains(&text) {
                left.push(text);
            }
        }
    }
    let mut top = Vec::new();
    for current_row in min_row..row {
        if let Some(text) = display_value(sheet, current_row, column) {
            if !top.contains(&text) {
                top.push(text);
            }
        }
    }
    let mut parts = Vec::new();
    if !left.is_empty() {
        parts.push(left.join("_"));
    }
    if !top.is_empty() {
        parts.push(top.join("_"));
    }
    parts.join("｜")
}

fn display_value(sheet: &SheetDump, row: i32, column: i32) -> Option<String> {
    let mut anchor = (row, column);
    for merge in &sheet.merges {
        let Some(bounds) = parse_cell_range(merge) else {
            continue;
        };
        if bounds.0 <= column && column <= bounds.2 && bounds.1 <= row && row <= bounds.3 {
            anchor = (bounds.1, bounds.0);
            break;
        }
    }
    let address = cell_address(anchor.1, anchor.0);
    let raw = sheet.values.get(&address)?;
    let value = CellValue::from_probe_json(raw).ok()?;
    let text = match value {
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Boolean(true) => "True".to_owned(),
        CellValue::Boolean(false) => "False".to_owned(),
        CellValue::Empty
        | CellValue::Number(_)
        | CellValue::Integer(_)
        | CellValue::DateTimeSerial(_)
        | CellValue::DateTimeIso(_)
        | CellValue::DurationIso(_)
        | CellValue::Error(_) => return None,
    };
    if text.is_empty() || is_numeric_text(&text) {
        return None;
    }
    Some(text)
}

fn is_numeric_text(value: &str) -> bool {
    let compact = value.replace(',', "");
    let digits = if compact.starts_with('+') || compact.starts_with('-') {
        &compact[1..]
    } else {
        &compact
    };
    if digits.is_empty() {
        return false;
    }
    let mut parts = digits.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some() || whole.is_empty() || !whole.chars().all(|ch| ch.is_ascii_digit()) {
        return false;
    }
    fraction.map_or(true, |fraction| {
        !fraction.is_empty() && fraction.chars().all(|ch| ch.is_ascii_digit())
    })
}

fn parse_cell_address(value: &str) -> Option<(i32, i32)> {
    let address = value.replace('$', "");
    let split = address.find(|ch: char| ch.is_ascii_digit())?;
    if split == 0 {
        return None;
    }
    let (letters, digits) = address.split_at(split);
    if !letters.chars().all(|ch| ch.is_ascii_alphabetic())
        || !digits.chars().all(|ch| ch.is_ascii_digit())
    {
        return None;
    }
    let mut column = 0i32;
    for ch in letters.chars() {
        column = column * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
    }
    let row = digits.parse::<i32>().ok()?;
    (column > 0 && row > 0).then_some((column, row))
}

fn cell_address(column: i32, row: i32) -> String {
    let mut number = column;
    let mut letters = String::new();
    while number > 0 {
        number -= 1;
        letters.insert(0, char::from(b'A' + (number % 26) as u8));
        number /= 26;
    }
    format!("{letters}{row}")
}

/// Parse one A1 cell or range and return (min column, min row, max column,
/// max row), matching the 1-based coordinates used by the Python resolver.
fn parse_cell_range(value: &str) -> Option<(i32, i32, i32, i32)> {
    let mut endpoints = value.split(':');
    let first = parse_cell_address(endpoints.next()?.trim())?;
    let last = match endpoints.next() {
        Some(value) => parse_cell_address(value.trim())?,
        None => first,
    };
    if endpoints.next().is_some() {
        return None;
    }
    Some((
        first.0.min(last.0),
        first.1.min(last.1),
        first.0.max(last.0),
        first.1.max(last.1),
    ))
}

#[cfg(test)]
mod tests {
    use super::{
        extract_conditional_issues_batch, resolve_indicator, CopyRange, S1ConditionalBatchInput,
        S1ConditionalMetadata,
    };
    use excel_reader::SheetDump;
    use serde_json::json;
    use std::{path::PathBuf, sync::Mutex};

    fn conditional_fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_conditional_eval.xlsx")
    }

    #[test]
    fn resolves_indicators_using_structure_ranges_and_merged_anchors() {
        let sheet = SheetDump {
            name: "Data".to_owned(),
            values: [
                ("A1".to_owned(), json!({"type":"string","value":"贷款余额"})),
                ("B1".to_owned(), json!({"type":"string","value":"本期"})),
                ("B2".to_owned(), json!({"type":"number","value":1000.0})),
            ]
            .into_iter()
            .collect(),
            formulas: Default::default(),
            formula_status: "read_by_calamine".to_owned(),
            merges: vec!["A1:A3".to_owned()],
            merge_status: "read_by_calamine".to_owned(),
            comments: Default::default(),
            comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
        };
        let structure_ranges = [CopyRange {
            sheet_name: "Data".to_owned(),
            address: "A1:B4".to_owned(),
        }];

        assert_eq!(
            resolve_indicator(&sheet, &structure_ranges, 3, 2),
            "贷款余额｜本期"
        );
    }

    #[test]
    fn numeric_display_text_is_not_used_as_an_indicator() {
        let sheet = SheetDump {
            name: "Data".to_owned(),
            values: [
                ("A1".to_owned(), json!({"type":"string","value":"1,234.5"})),
                ("A2".to_owned(), json!({"type":"string","value":"余额"})),
            ]
            .into_iter()
            .collect(),
            formulas: Default::default(),
            formula_status: "read_by_calamine".to_owned(),
            merges: Vec::new(),
            merge_status: "read_by_calamine".to_owned(),
            comments: Default::default(),
            comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
        };
        let structure_ranges = [CopyRange {
            sheet_name: "Data".to_owned(),
            address: "A1:B3".to_owned(),
        }];

        assert_eq!(resolve_indicator(&sheet, &structure_ranges, 2, 2), "余额");
    }

    #[test]
    fn batch_keeps_per_file_issue_identity_and_reports_engine_timings() {
        let fixture = conditional_fixture();
        let inputs = [
            S1ConditionalBatchInput {
                workbook_path: fixture.clone(),
                source_file: PathBuf::from("甲银行_20260731.xlsx"),
                metadata: S1ConditionalMetadata {
                    period: "2026-07-31".to_owned(),
                    batch_id: "batch-1".to_owned(),
                    audit_time: "2026-10-24 12:00:00".to_owned(),
                    org_code: "9144".to_owned(),
                    org_name: "甲银行".to_owned(),
                },
            },
            S1ConditionalBatchInput {
                workbook_path: fixture,
                source_file: PathBuf::from("乙银行_20260731.xlsx"),
                metadata: S1ConditionalMetadata {
                    period: "2026-07-31".to_owned(),
                    batch_id: "batch-1".to_owned(),
                    audit_time: "2026-10-24 12:00:00".to_owned(),
                    org_code: "9145".to_owned(),
                    org_name: "乙银行".to_owned(),
                },
            },
        ];
        let events = Mutex::new(Vec::new());
        let report = extract_conditional_issues_batch(&inputs, &[], &[], |event| {
            events.lock().unwrap().push(event.clone());
        })
        .expect("extract the selected workbook batch");

        assert_eq!(report.files.len(), 2);
        assert_eq!(report.issues.len(), 8);
        assert_eq!(report.files[0].source_file, "甲银行_20260731.xlsx");
        assert_eq!(report.files[1].source_file, "乙银行_20260731.xlsx");
        assert!(report
            .files
            .iter()
            .all(|file| file.reader == "calamine+namespace_aware_ooxml"));
        assert!(report
            .files
            .iter()
            .all(|file| file.total_time_ms >= file.read_time_ms));
        assert_eq!(report.issues[0].org_code, "9144");
        assert_eq!(report.issues[4].org_code, "9145");
        assert!(report.issues[0].issue_id.starts_with("甲银行｜"));
        assert!(report.issues[4].issue_id.starts_with("乙银行｜"));

        let events = events.into_inner().unwrap();
        assert_eq!(events.len(), 4);
        assert_eq!(
            events
                .iter()
                .filter(|event| event.status == "started")
                .count(),
            2
        );
        assert_eq!(
            events
                .iter()
                .filter(|event| event.status == "completed")
                .count(),
            2
        );
        let first_completed = events
            .iter()
            .find(|event| event.file_index == 1 && event.status == "completed")
            .expect("first file completion event");
        assert_eq!(first_completed.issue_count, Some(4));
        assert_eq!(
            first_completed.reader,
            Some("calamine+namespace_aware_ooxml")
        );
        assert!(events
            .iter()
            .any(|event| event.file_index == 2 && event.source_file == "乙银行_20260731.xlsx"));
    }

    #[test]
    fn batch_retains_successful_files_when_an_independent_workbook_fails() {
        let valid = conditional_fixture();
        let invalid = std::env::temp_dir().join(format!(
            "missing-s1-conditional-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        let inputs = [
            S1ConditionalBatchInput {
                workbook_path: valid,
                source_file: PathBuf::from("有效机构.xlsx"),
                metadata: S1ConditionalMetadata::default(),
            },
            S1ConditionalBatchInput {
                workbook_path: invalid,
                source_file: PathBuf::from("损坏机构.xlsx"),
                metadata: S1ConditionalMetadata::default(),
            },
        ];

        let report = extract_conditional_issues_batch(&inputs, &[], &[], |_| {})
            .expect("per-file errors must not stop the remaining batch");
        assert_eq!(report.status, "completed_with_diagnostics");
        assert_eq!(report.files.len(), 2);
        assert_eq!(report.issues.len(), 4);
        assert_eq!(report.files[0].status, "completed_with_diagnostics");
        assert!(
            !report.files[0].diagnostics.is_empty(),
            "the retained successful workbook must keep its conditional-format diagnostics"
        );
        assert_eq!(report.files[1].status, "failed");
        assert!(report.files[1]
            .error
            .as_deref()
            .unwrap_or_default()
            .contains("missing-s1-conditional"));
    }
}
