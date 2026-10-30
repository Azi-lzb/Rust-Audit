//! S5 projection of the existing read-only template/cache error scanner.
//! No formulas are injected, calculated, or written to the source workbook.

use crate::{s1_validation_formula_scan::S1ValidationFormulaScanPlan, s5_service::S5RuleOutcome};
use serde_json::{json, Value};
use std::path::{Path, PathBuf};

pub struct S5ValidationFormulaScanner {
    plan: Result<S1ValidationFormulaScanPlan, String>,
}

pub struct S5ValidationFormulaRun {
    pub outcomes: Vec<S5RuleOutcome>,
    pub summary: Value,
}

impl S5ValidationFormulaScanner {
    pub fn new(template_paths: &[PathBuf]) -> Self {
        Self {
            plan: S1ValidationFormulaScanPlan::load(template_paths)
                .map_err(|error| format!("{error:#}")),
        }
    }

    pub fn scan(
        &self,
        source: &Path,
        business: &str,
        period: &str,
        institution_code: &str,
        institution_name: &str,
    ) -> S5ValidationFormulaRun {
        let base = |rule_id: String, code: &str, sheet: &str, cell: &str, message: String| {
            S5RuleOutcome {
                origin_source_path: source.to_string_lossy().into_owned(),
                rule_id,
                revision: 1,
                rule_name: "模板校验公式错误扫描".to_owned(),
                business: business.to_owned(),
                period: period.to_owned(),
                institution_code: institution_code.to_owned(),
                institution_name: institution_name.to_owned(),
                severity: "错误".to_owned(),
                check_type: "校验公式扫描失败".to_owned(),
                metric: "模板校验公式".to_owned(),
                sheet_name: sheet.to_owned(),
                status: "CALCULATION_FAILED".to_owned(),
                error_code: Some(code.to_owned()),
                error_message: message,
                error_evidence: json!({"source_path": source.to_string_lossy(), "sheet": sheet, "cell": cell}),
                ..S5RuleOutcome::default()
            }
        };
        let report = self
            .plan
            .as_ref()
            .map_err(Clone::clone)
            .and_then(|plan| plan.scan(source).map_err(|error| format!("{error:#}")));
        match report {
            Ok(report) => {
                let mut outcomes = Vec::new();
                for error in &report.errors {
                    let message = match error.origin.as_str() {
                        "template_formula" => format!(
                            "模板校验公式本身含 Excel 错误值或引用 {}；原公式：{}",
                            error.value, error.formula
                        ),
                        "template_formula_and_source_cached_value" => format!(
                            "模板校验公式及源文件结果均包含 Excel 错误值 {}；原公式：{}",
                            error.value, error.formula
                        ),
                        _ => format!(
                            "模板校验公式在源文件中返回 Excel 错误值 {}；原公式：{}",
                            error.value, error.formula
                        ),
                    };
                    let mut outcome = base(
                        format!("TEMPLATE_FORMULA_ERROR:{}!{}", error.sheet, error.cell),
                        "TEMPLATE_FORMULA_ERROR",
                        &error.sheet,
                        &error.cell,
                        message,
                    );
                    outcome.check_type = "校验公式错误".to_owned();
                    outcome.metric = "模板校验公式结果".to_owned();
                    outcome.current_value = Some(json!(error.value));
                    outcome.error_evidence["template_formula"] = json!(error.formula);
                    outcome.error_evidence["origin"] = json!(error.origin);
                    outcome.error_evidence["template_paths"] = json!(error.template_paths);
                    outcome.error_evidence["formula_error_tokens"] =
                        json!(error.formula_error_tokens);
                    outcomes.push(outcome);
                }
                for sheet in &report.missing_sheets {
                    outcomes.push(base(
                        format!("TEMPLATE_FORMULA_SCAN:{sheet}"),
                        "TEMPLATE_VALIDATION_SHEET_MISSING",
                        sheet,
                        "",
                        format!("本次源文件缺少模板校验公式所在工作表：{sheet}"),
                    ));
                }
                S5ValidationFormulaRun {
                    outcomes,
                    summary: json!(report),
                }
            }
            Err(message) => S5ValidationFormulaRun {
                outcomes: vec![base(
                    "TEMPLATE_FORMULA_SCAN:FAILED".to_owned(),
                    "TEMPLATE_FORMULA_SCAN_FAILED",
                    "",
                    "",
                    format!("模板校验公式错误扫描未完成：{message}"),
                )],
                summary: json!({"status": "FAILED", "error": message}),
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cached_validation_errors_are_projected_as_owned_s5_failures_without_recalculation() {
        let stamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("s5-validation-cache-{stamp}"));
        std::fs::create_dir_all(&root).unwrap();
        let template = root.join("template.xlsx");
        let source = root.join("source.xlsx");
        let mut book = excel_writer::Workbook::new();
        book.add_worksheet()
            .set_name("数据表")
            .unwrap()
            .write_formula(0, 0, "=1/0")
            .unwrap();
        book.define_name("校验区域_1", "'数据表'!$A$1").unwrap();
        book.save(&template).unwrap();
        let mut book = excel_writer::Workbook::new();
        book.add_worksheet()
            .set_name("数据表")
            .unwrap()
            .write_string(0, 0, "#DIV/0!")
            .unwrap();
        book.save(&source).unwrap();
        let before = std::fs::read(&source).unwrap();
        let result = S5ValidationFormulaScanner::new(&[template]).scan(
            &source,
            "业务",
            "2026-07-31",
            "I001",
            "机构",
        );
        assert_eq!(result.outcomes.len(), 1);
        let outcome = &result.outcomes[0];
        assert_eq!(outcome.status, "CALCULATION_FAILED");
        assert_eq!(
            outcome.error_code.as_deref(),
            Some("TEMPLATE_FORMULA_ERROR")
        );
        assert_eq!(outcome.current_value, Some(json!("#DIV/0!")));
        assert_eq!(outcome.error_evidence["cell"], "A1");
        assert_eq!(Path::new(&outcome.origin_source_path), source.as_path());
        assert_eq!(std::fs::read(&source).unwrap(), before);
        std::fs::remove_dir_all(root).unwrap();
    }
}
