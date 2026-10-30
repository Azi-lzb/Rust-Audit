//! 固定 S1 校验区域扫描的两种输入口径（Goal C）。
//!
//! 业务事实（矩阵 S1F1.11）：
//! - GUI 完整流程对**重算后的审核副本**扫描校验区域（execute_s1_f1_workflow
//!   使用 `plan.scan(&execution.recalculation.output)`）；
//! - CLI `s1-audit-directory` 的保存缓存诊断对**源文件**的缓存值扫描
//!   （open-data-audit.rs 使用 `plan.scan(&input.source_file)`）。
//!
//! 本测试用“源缓存与审核副本内容不同”的合成样本锁定两种口径各自的读数，
//! 供用户在 ALIGNMENT_DECISION_LIST_20261029.md 的 AL-03 上做统一决定；
//! 在用户答复前，两条调用路径的行为都不得改动。

use anyhow::Result;
use excel_writer::Workbook;
use serde_json::json;
use std::{fs, path::Path, path::PathBuf, time::{SystemTime, UNIX_EPOCH}};

fn temporary_directory(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s1_scan_scope_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary directory");
    path
}

/// 模板：校验区域 = 检查表!A1:A2，两条 IF 规则。
fn write_template(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("检查表").expect("name template sheet");
    sheet
        .write_formula(0, 0, r#"=IF(TRUE,"错误|字段|指标|说明","通过")"#)
        .expect("write rule 1");
    sheet
        .write_formula(1, 0, r#"=IF(TRUE,"","通过")"#)
        .expect("write rule 2");
    workbook
        .define_name("校验区域", "='检查表'!$A$1:$A$2")
        .expect("define validation area");
    workbook.save(path).expect("save template");
}

/// 源文件：A1 缓存“通过”，A2 缓存 #REF!（源侧错误）。
fn write_source(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("检查表").expect("name source sheet");
    sheet.write_string(0, 0, "通过").expect("write A1 cache");
    sheet.write_string(1, 0, "#REF!").expect("write A2 cache");
    workbook.save(path).expect("save source");
}

/// 审核副本（重算结果口径）：A1 缓存 #DIV/0!（副本侧重算出的 Excel 错误），
/// A2 为空（源侧 #REF! 在副本中不存在）。
fn write_audit_copy(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("检查表").expect("name audit copy sheet");
    sheet.write_string(0, 0, "#DIV/0!").expect("write A1 recalc result");
    sheet.write_string(1, 0, "").expect("write A2 recalc result");
    workbook.save(path).expect("save audit copy");
}

#[test]
fn scan_scope_follows_the_workbook_argument_it_is_given() -> Result<()> {
    let directory = temporary_directory("scope");
    let template = directory.join("template.xlsx");
    let source = directory.join("source.xlsx");
    let audit_copy = directory.join("audit-copy.xlsx");
    write_template(&template);
    write_source(&source);
    write_audit_copy(&audit_copy);

    let plan =
        audit_core::s1_validation_formula_scan::S1ValidationFormulaScanPlan::load(&[template])?;

    // CLI 口径：扫描源文件缓存值 → 报告源侧 A2 的 #REF!，不报副本侧 A1。
    let from_source = plan.scan(&source)?;
    assert_eq!(from_source.status, "COMPLETED");
    let source_cells: Vec<String> = from_source
        .errors
        .iter()
        .map(|error| error.cell.clone())
        .collect();
    assert_eq!(
        source_cells,
        vec!["A2".to_owned()],
        "源文件口径必须报告源缓存 A2 的 #REF!"
    );

    // GUI 口径：扫描重算副本 → 报告副本侧 A1 的规则异常，不报源侧 A2。
    let from_copy = plan.scan(&audit_copy)?;
    assert_eq!(from_copy.status, "COMPLETED");
    let copy_cells: Vec<String> = from_copy
        .errors
        .iter()
        .map(|error| error.cell.clone())
        .collect();
    assert_eq!(
        copy_cells,
        vec!["A1".to_owned()],
        "重算副本口径必须报告副本 A1 的 Excel 错误"
    );

    // 同一样本、同一模板，仅输入不同即产生不同异常集合——这就是
    // AL-03 需要用户统一的口径差异。
    assert_ne!(source_cells, copy_cells);

    fs::remove_dir_all(directory).expect("remove scan-scope fixtures");
    Ok(())
}

/// 帮助文档用途的断言：两条调用路径在源码中的取参表达式保持各自语义。
#[test]
fn scan_call_sites_keep_their_documented_arguments() {
    let gui_bridge = include_str!("../src/web_bridge.rs");
    assert!(
        gui_bridge.contains(".scan(audit_copy)"),
        "GUI 完整流程必须继续扫描重算审核副本（用户统一口径前不得改动）"
    );
    let cli = include_str!("../src/bin/open-data-audit.rs");
    assert!(
        cli.contains("plan.scan(&input.source_file)"),
        "CLI 保存缓存诊断必须继续扫描源文件（用户统一口径前不得改动）"
    );
    // JSON 合同供决策清单引用
    let payload = json!({
        "gui": "recalculation.output（重算副本）",
        "cli": "source_file（源文件缓存值）"
    });
    assert_eq!(payload["gui"], "recalculation.output（重算副本）");
}
