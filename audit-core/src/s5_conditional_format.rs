//! S5 conditional-format evaluation and result projection.
//!
//! This reuses the bounded Rust OOXML evaluator used by S1. A rule is only
//! reported as triggered when evaluation returns true; unsupported and failed
//! evaluations remain explicit result rows.

use crate::conditional_format::{
    evaluate_xlsx_workbook_with_dump, ConditionalDiagnostic, ConditionalTrigger,
    EvaluationStatus,
};
use anyhow::{Context, Result};
use excel_reader::CellValue;
use serde::Serialize;
use std::path::Path;

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S5ConditionalFormatFinding {
    pub institution_name: String,
    pub report_item: String,
    pub institution_code: String,
    pub sheet_name: String,
    pub cell: String,
    pub cell_value: String,
    pub status: String,
    pub rule_type: String,
    pub formula: String,
    pub applies_to: String,
    pub priority: u32,
    pub source_file: String,
    pub detail: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S5ConditionalFormatRun {
    pub status: String,
    pub reader: String,
    pub rules_total: usize,
    pub rules_supported: usize,
    pub scanned_cells: usize,
    pub triggered_count: usize,
    pub unsupported_count: usize,
    pub error_count: usize,
    pub read_time_ms: u128,
    pub evaluation_time_ms: u128,
    pub elapsed_ms: u128,
    pub findings: Vec<S5ConditionalFormatFinding>,
}

/// Evaluate the selected current-period workbook's OOXML conditional rules.
/// No Office/LibreOffice recalculation is performed. The evaluator reports
/// unsupported rules and formula errors as findings instead of treating them
/// as non-triggered.
pub fn run_s5_conditional_format(
    source_path: impl AsRef<Path>,
    institution_name: &str,
    report_item: &str,
    institution_code: &str,
) -> Result<S5ConditionalFormatRun> {
    let source_path = source_path.as_ref();
    let (evaluation, workbook) = evaluate_xlsx_workbook_with_dump(source_path, &[])
        .with_context(|| format!("读取并计算 S5 条件格式：{}", source_path.display()))?;

    let source_file = source_path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| source_path.to_string_lossy().into_owned());
    let mut findings = Vec::with_capacity(evaluation.triggered.len() + evaluation.diagnostics.len());

    for trigger in &evaluation.triggered {
        findings.push(trigger_finding(
            trigger,
            &workbook.sheets,
            institution_name,
            report_item,
            institution_code,
            &source_file,
        )?);
    }
    for diagnostic in &evaluation.diagnostics {
        findings.push(diagnostic_finding(
            diagnostic,
            &workbook.sheets,
            institution_name,
            report_item,
            institution_code,
            &source_file,
        )?);
    }

    let unsupported_count = evaluation
        .diagnostics
        .iter()
        .filter(|item| item.status == EvaluationStatus::Unsupported)
        .count();
    let error_count = evaluation
        .diagnostics
        .iter()
        .filter(|item| item.status == EvaluationStatus::Error)
        .count();
    let status = if unsupported_count + error_count > 0 {
        "PARTIAL"
    } else if evaluation.rules_total == 0 {
        "COMPLETED_NO_RULES"
    } else {
        "COMPLETED"
    };

    Ok(S5ConditionalFormatRun {
        status: status.to_owned(),
        reader: evaluation.reader.to_owned(),
        rules_total: evaluation.rules_total,
        rules_supported: evaluation.rules_supported,
        scanned_cells: evaluation.scanned_cells,
        triggered_count: evaluation.triggered.len(),
        unsupported_count,
        error_count,
        read_time_ms: evaluation.read_time_ms,
        evaluation_time_ms: evaluation.evaluation_time_ms,
        elapsed_ms: evaluation.total_time_ms,
        findings,
    })
}

fn trigger_finding(
    trigger: &ConditionalTrigger,
    sheets: &[excel_reader::SheetDump],
    institution_name: &str,
    report_item: &str,
    institution_code: &str,
    source_file: &str,
) -> Result<S5ConditionalFormatFinding> {
    let mut detail = trigger.message.clone();
    if !trigger.unsupported_reason.is_empty() {
        detail.push_str("；同一单元格另有未判断规则：");
        detail.push_str(&trigger.unsupported_reason);
    }
    Ok(S5ConditionalFormatFinding {
        institution_name: institution_name.to_owned(),
        report_item: report_item.to_owned(),
        institution_code: institution_code.to_owned(),
        sheet_name: trigger.sheet.clone(),
        cell: trigger.cell.clone(),
        cell_value: cell_value_text(sheets, &trigger.sheet, &trigger.cell)?,
        status: if trigger.confidence == "PARTIAL" {
            "TRIGGERED_PARTIAL".to_owned()
        } else {
            "TRIGGERED".to_owned()
        },
        rule_type: trigger.rule_type.clone(),
        formula: trigger.formula.clone(),
        applies_to: trigger.applies_to.clone(),
        priority: trigger.priority,
        source_file: source_file.to_owned(),
        detail,
    })
}

fn diagnostic_finding(
    diagnostic: &ConditionalDiagnostic,
    sheets: &[excel_reader::SheetDump],
    institution_name: &str,
    report_item: &str,
    institution_code: &str,
    source_file: &str,
) -> Result<S5ConditionalFormatFinding> {
    let status = match diagnostic.status {
        EvaluationStatus::Unsupported => "UNSUPPORTED",
        EvaluationStatus::Error => "ERROR",
        EvaluationStatus::False => "NOT_TRIGGERED",
        EvaluationStatus::True => "TRIGGERED",
    };
    Ok(S5ConditionalFormatFinding {
        institution_name: institution_name.to_owned(),
        report_item: report_item.to_owned(),
        institution_code: institution_code.to_owned(),
        sheet_name: diagnostic.sheet.clone(),
        cell: diagnostic.cell.clone(),
        cell_value: cell_value_text(sheets, &diagnostic.sheet, &diagnostic.cell)?,
        status: status.to_owned(),
        rule_type: diagnostic.rule_type.clone(),
        formula: diagnostic.formula.clone(),
        applies_to: diagnostic.applies_to.clone(),
        priority: diagnostic.priority,
        source_file: source_file.to_owned(),
        detail: diagnostic.reason.clone(),
    })
}

fn cell_value_text(
    sheets: &[excel_reader::SheetDump],
    sheet_name: &str,
    cell: &str,
) -> Result<String> {
    let raw = sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .with_context(|| format!("条件格式结果引用了未知工作表：{sheet_name}"))?
        .values
        .get(cell);
    let Some(raw) = raw else {
        return Ok(String::new());
    };
    let value = CellValue::from_probe_json(raw)
        .with_context(|| format!("读取 S5 条件格式触发值 {sheet_name}!{cell}"))?;
    Ok(match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value)
        | CellValue::DateTimeIso(value)
        | CellValue::DurationIso(value)
        | CellValue::Error(value) => value,
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => value.to_string(),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => value.to_string(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_conditional_eval.xlsx")
    }

    #[test]
    fn maps_only_true_rules_as_triggers_and_keeps_unsupported_results_visible() {
        let result = run_s5_conditional_format(
            fixture(),
            "测试机构",
            "单位贷款",
            "I001",
        )
        .expect("evaluate S5 conditional formats");

        assert_eq!(result.status, "PARTIAL");
        assert_eq!(result.rules_total, 7);
        assert_eq!(result.rules_supported, 6);
        assert_eq!(result.scanned_cells, 12);
        assert_eq!(result.triggered_count, 4);
        assert_eq!(result.unsupported_count, 3);
        assert_eq!(result.error_count, 1);
        assert_eq!(
            result
                .findings
                .iter()
                .filter(|finding| finding.status == "TRIGGERED")
                .count(),
            4
        );
        assert_eq!(
            result
                .findings
                .iter()
                .filter(|finding| finding.status == "UNSUPPORTED")
                .count(),
            3
        );
        assert_eq!(
            result
                .findings
                .iter()
                .filter(|finding| finding.status == "ERROR")
                .count(),
            1
        );
        assert_eq!(result.findings[0].institution_code, "I001");
        assert_eq!(result.findings[0].cell, "B2");
        assert!(!result.findings[0].cell_value.is_empty());
    }
}
