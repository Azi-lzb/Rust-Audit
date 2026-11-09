//! S3-F05 CLI: comparison workbook → financial forms, through the production
//! Excel-Core route. Expected values come from the synthetic fixtures copied
//! from the S3-F05 unit tests, not from CLI output.

use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_directory(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s3_f05_cli_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary directory");
    path
}

fn write_template(path: &Path) {
    let mut workbook = Workbook::new();
    let forms = workbook.add_worksheet();
    forms.set_name("报表清单").unwrap();
    for (column, value) in ["报表代码", "报表名称", "频度", "批次", "是否启用写入"]
        .iter()
        .enumerate()
    {
        forms.write_string(0, column as u16, *value).unwrap();
    }
    for (row, code, name) in [
        (1, "测试A1", "测试表"),
        (2, "空表A1", "无数据表"),
        (3, "关闭A2", "关闭表"),
    ] {
        forms.write_string(row, 0, code).unwrap();
        forms.write_string(row, 1, name).unwrap();
        forms.write_string(row, 2, "月").unwrap();
        forms.write_string(row, 3, "1").unwrap();
    }
    forms.write_string(1, 4, "是").unwrap();
    forms.write_string(2, 4, "是").unwrap();
    forms.write_number(3, 4, 0.0).unwrap();

    let instructions = workbook.add_worksheet();
    instructions.set_name("使用说明").unwrap();
    instructions.write_string(0, 0, "说明").unwrap();

    let settings = workbook.add_worksheet();
    settings.set_name("转表设置").unwrap();
    settings.write_string(0, 0, "参数").unwrap();
    settings.write_string(0, 1, "值").unwrap();
    settings.write_string(1, 0, "隐藏空行").unwrap();
    settings.write_string(1, 1, "是").unwrap();
    settings.write_string(2, 0, "空表删除").unwrap();
    settings.write_string(2, 1, "是").unwrap();

    let populated = workbook.add_worksheet();
    populated.set_name("测试A1").unwrap();
    write_form_sheet_headers(populated);
    populated.write_string(1, 0, "I001").unwrap();
    populated.write_string(2, 0, "I002").unwrap();

    let empty = workbook.add_worksheet();
    empty.set_name("空表A1").unwrap();
    write_form_sheet_headers(empty);
    empty.write_string(1, 0, "NO-MATCH").unwrap();

    let disabled = workbook.add_worksheet();
    disabled.set_name("关闭A2").unwrap();
    write_form_sheet_headers(disabled);
    disabled.write_string(1, 0, "I001").unwrap();

    workbook.save(path).unwrap();
}

fn write_form_sheet_headers(sheet: &mut excel_writer::Worksheet) {
    for (column, value) in [
        "指标代码",
        "指标名称",
        "行序号",
        "余额人民币",
        "上期-余额人民币",
        "增减额-余额人民币",
        "环比-余额人民币",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *value).unwrap();
    }
}

fn write_comparison(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("2026年8月对比结果").unwrap();
    let headers = [
        "业务类",
        "数据日期",
        "机构类代码",
        "机构类名称",
        "地区代码",
        "地区名称",
        "指标顺序码",
        "指标代码",
        "指标名称",
        "数据属性",
        "币种",
        "频度",
        "批次",
        "数据值",
        "上期值",
        "增减额(亿元)",
        "环比",
        "备注",
        "是否说明",
        "说明内容",
        "计算过程",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header).unwrap();
    }
    for (row, indicator, frequency, batch, current, previous) in [
        (1, "I001", "月", "1", 120.0, 100.0),
        (2, "I003", "月", "1", 5.0, 4.0),
        (3, "I001", "季", "2", 999.0, 888.0),
    ] {
        for (column, value) in [
            (0, "人民币"),
            (1, "2026-08-31"),
            (2, "0001"),
            (3, "机构"),
            (4, "1100000"),
            (5, "地区"),
            (6, indicator),
            (7, indicator),
            (8, "指标"),
            (9, "余额"),
            (10, "人民币"),
        ] {
            sheet.write_string(row, column, value).unwrap();
        }
        sheet.write_string(row, 11, frequency).unwrap();
        sheet.write_string(row, 12, batch).unwrap();
        sheet.write_number(row, 13, current).unwrap();
        sheet.write_number(row, 14, previous).unwrap();
    }
    workbook.save(path).unwrap();
}

fn write_common_config(path: &Path) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("环比警戒").unwrap();
    for (column, header) in [
        "下限",
        "上限",
        "备注文字",
        "填充颜色",
        "是否整行填充",
        "金额变动阈值",
        "是否说明",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(0, column as u16, *header).unwrap();
    }
    sheet.write_string(1, 0, "10").unwrap();
    sheet.write_string(1, 1, "30").unwrap();
    sheet.write_string(1, 3, "6").unwrap();
    workbook.save(path).unwrap();
}

#[test]
fn cli_generates_financial_forms_and_keeps_inputs_unchanged() {
    let directory = temporary_directory("success");
    let comparison = directory.join("比较结果.xlsx");
    let template = directory.join("3.3配置.xlsx");
    let common = directory.join("3.0配置.xlsx");
    let output = directory.join("forms");
    write_comparison(&comparison);
    write_template(&template);
    write_common_config(&common);
    let comparison_before = fs::read(&comparison).unwrap();
    let template_before = fs::read(&template).unwrap();
    let common_before = fs::read(&common).unwrap();

    let result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
        .arg("s3-generate-financial-forms")
        .arg("--comparison-workbook")
        .arg(&comparison)
        .arg("--common-config")
        .arg(&common)
        .arg("--forms-config")
        .arg(&template)
        .arg("--output-directory")
        .arg(&output)
        .arg("--developer-view")
        .output()
        .expect("run S3-F05 CLI");
    let stdout = String::from_utf8_lossy(&result.stdout);
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(result.status.success(), "stdout={stdout} stderr={stderr}");
    let report: Value = serde_json::from_slice(&result.stdout).expect("valid S3-F05 CLI JSON");
    assert_eq!(report["status"], "completed");
    assert_eq!(report["feature"], "S3-F05");
    assert_eq!(report["engine"], "rust_s3_indicator_split");
    assert_eq!(report["comparison_rows"], 3);
    assert_eq!(report["institution_count"], 1);
    assert_eq!(report["matched_indicators"], 1);
    assert_eq!(report["edits"].as_array().unwrap().len(), 1);
    // 用户确认（2026-11-08）：S3-F05 auto 路由 OOXML 直改优先，Umya 已退出
    // 该功能（定义名守卫会拒绝 DeleteColumn 等操作）。
    assert_eq!(
        report["edits"][0]["engine"], "ooxml_patch",
        "the production route must use the OOXML patch backend (Umya excluded from S3-F05)"
    );

    // One published output directory holding one generated form workbook.
    let output_directory = PathBuf::from(
        report["output_directory"]
            .as_str()
            .expect("output directory path"),
    );
    assert!(output_directory.is_dir());
    let output_files = report["output_files"].as_array().unwrap();
    assert_eq!(output_files.len(), 1);
    let form_path = PathBuf::from(output_files[0].as_str().unwrap());
    assert!(form_path.is_file());

    // Read back with the independent reader. The template's original
    // 上期/增减额/环比 columns are dropped and the three indicator blocks are
    // rewritten as 上期-/增减额-/环比- + the base header, so the generated form
    // holds 指标代码, 指标名称, 行序号, 余额人民币 and those three blocks.
    let dump = excel_reader::read_workbook(&form_path).expect("reopen generated form");
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "测试A1")
        .expect("generated form sheet");
    let value = |cell: &str| {
        sheet
            .values
            .get(cell)
            .map(|value| value["value"].to_string())
            .unwrap_or_default()
    };
    assert_eq!(value("A2"), "\"I001\"");
    assert_eq!(value("D2"), "120.0");
    assert_eq!(value("E2"), "100.0");
    assert_eq!(value("F2"), "20.0");
    assert_eq!(value("G2"), "0.2");
    let headers = [
        "指标代码",
        "指标名称",
        "行序号",
        "余额人民币",
        "上期-余额人民币",
        "增减额-余额人民币",
        "环比-余额人民币",
    ];
    for (index, header) in headers.iter().enumerate() {
        let column = (b'A' + index as u8) as char;
        assert_eq!(
            value(&format!("{column}1")),
            format!("\"{header}\""),
            "column {column} header after the form rewrite"
        );
    }
    assert!(dump
        .sheets
        .iter()
        .all(|sheet| sheet.name != "空表A1" && sheet.name != "关闭A2"));
    assert!(dump.sheets.iter().all(|sheet| sheet.name != "使用说明"));

    // Sources are never modified.
    assert_eq!(fs::read(&comparison).unwrap(), comparison_before);
    assert_eq!(fs::read(&template).unwrap(), template_before);
    assert_eq!(fs::read(&common).unwrap(), common_before);

    // Developer view reports one started and one completed stage.
    let events = stderr
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(events.iter().any(|event| {
        event["feature"] == "S3-F05"
            && event["step"] == "indicator_split"
            && event["status"] == "started"
    }));
    assert!(events.iter().any(|event| {
        event["feature"] == "S3-F05"
            && event["step"] == "indicator_split"
            && event["status"] == "completed"
            && event["matched_indicators"] == 1
    }));

    fs::remove_dir_all(directory).expect("remove S3-F05 CLI files");
}

#[test]
fn cli_refuses_missing_configuration_before_writing_outputs() {
    let directory = temporary_directory("missing_config");
    let comparison = directory.join("比较结果.xlsx");
    let template = directory.join("3.3配置.xlsx");
    let missing_common = directory.join("不存在的3.0配置.xlsx");
    let output = directory.join("forms");
    write_comparison(&comparison);
    write_template(&template);

    let result = Command::new(env!("CARGO_BIN_EXE_AuditTools"))
        .arg("s3-generate-financial-forms")
        .arg("--comparison-workbook")
        .arg(&comparison)
        .arg("--common-config")
        .arg(&missing_common)
        .arg("--forms-config")
        .arg(&template)
        .arg("--output-directory")
        .arg(&output)
        .output()
        .expect("run S3-F05 CLI with a missing config");
    assert!(
        !result.status.success(),
        "a missing 3.0 config must fail the run"
    );
    assert!(
        !output.exists(),
        "no output directory may be created when configuration is missing"
    );
    fs::remove_dir_all(directory).expect("remove S3-F05 failure files");
}
