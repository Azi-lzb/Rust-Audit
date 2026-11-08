//! Open-Data-Audit project package.
//!
//! Shared spreadsheet operations are provided by the workspace crates. This
//! package owns only Audit-specific probes, fixtures, rules, and future business
//! workflows.

// The GUI state literal in web_bridge.rs is one large `json!` expression; the
// serde_json macro expansion needs more recursion depth than the default 128.
#![recursion_limit = "512"]

pub mod acceptance_runner;
pub mod app_settings;
pub mod central_csv;
pub mod conditional_format;
pub mod config_update;
pub mod discovery;
pub mod excel_open_repair;
pub mod excel_route;
pub mod models;
pub mod name_config;
pub mod native_file_dialog;
#[cfg(windows)]
pub mod native_office_com;
#[cfg(windows)]
pub mod native_recalculation;
pub mod period_pairs;
pub mod report_source;
pub mod s1_conditional_audit;
pub mod s1_formula_audit;
pub mod s1_formula_copy;
pub mod s1_history;
pub mod s1_issue_chain;
pub mod s1_recalculation;
pub mod s1_recovery;
pub mod s1_region_summary;
pub mod s1_result_workbook;
pub mod s1_structure_compare;
pub mod s1_summary_merge;
pub mod s1_template;
pub mod s1_template_inspection;
pub mod s1_template_rules;
pub mod s1_validation_formula_scan;
pub mod s2_analysis;
pub mod s2_config;
pub mod s2_config_check;
pub mod s2_exporter;
mod s2_expression;
pub mod s2_external;
pub mod s2_external_import;
pub mod s2_import;
pub mod s2_period;
pub mod s2_rule_engine;
pub mod s2_service;
pub mod s2_summary;
pub mod s2_validation;
pub mod s3_current_value_check;
pub mod s3_execution_compare;
pub mod s3_explanation_export;
pub mod s3_explanation_feedback;
pub mod s3_indicator_split;
pub mod template;
pub mod unit_conversion;
pub mod web_bridge;
