use excel_reader::{read_value_columns, CellValue, SheetValues};
use excel_writer::Workbook;
use std::{env, error::Error, fs, path::Path};

fn write_sheet(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[&[&str]],
) -> Result<(), Box<dyn Error>> {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name)?;
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            sheet.write_string(row as u32 + 1, column as u16, *value)?;
        }
    }
    Ok(())
}

fn write_s3_configs(root: &Path) -> Result<(), Box<dyn Error>> {
    let mut common = Workbook::new();
    write_sheet(
        &mut common,
        "运行参数",
        &["参数", "值"],
        &[&["默认源数据单位", "元"], &["默认目标单位", "元"]],
    )?;
    write_sheet(&mut common, "单位换算例外", &["指标代码", "启用"], &[])?;
    write_sheet(
        &mut common,
        "环比警戒",
        &[
            "下限",
            "上限",
            "备注文字",
            "填充颜色",
            "是否整行填充",
            "金额变动阈值",
            "是否说明",
        ],
        &[],
    )?;
    common.save(root.join("s3_common_config.xlsx"))?;

    let mut feature = Workbook::new();
    write_sheet(
        &mut feature,
        "指标参照",
        &["指标代码", "表单代码", "行序号", "指标别名"],
        &[&["I001", "F001", "1", "模拟贷款余额"]],
    )?;
    write_sheet(
        &mut feature,
        "规则动作",
        &[
            "规则编号",
            "规则说明",
            "规则动作",
            "指标代码",
            "数据属性",
            "币种",
            "比较指标代码",
            "人民币阈值(亿元)",
            "美元阈值(亿美元)",
            "变幅阈值(%)",
            "来源分组",
            "适用表单",
            "适用频度",
            "适用批次",
            "适用场景",
            "详细说明",
            "启用",
            "停用原因",
            "备注",
        ],
        &[&[
            "DEMO-S3-CHANGE",
            "指标发生变动时提示",
            "指标有变动需说明。",
            "I001",
            "余额",
            "",
            "",
            "0",
            "0",
            "0",
            "",
            "",
            "月报",
            "第1批",
            "",
            "本样例按配置对任意非零变动触发提示。",
            "是",
            "",
            "",
        ]],
    )?;
    write_sheet(
        &mut feature,
        "表达式校验",
        &[
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ],
        &[],
    )?;
    feature.save(root.join("s3_feature_config.xlsx"))?;
    Ok(())
}

fn write_s3_csvs(root: &Path) -> Result<(), Box<dyn Error>> {
    let headers = "业务类,数据日期,机构类代码,机构类名称,地区代码,地区名称,指标顺序码,指标代码,指标名称,数据属性,币种,频度,批次,数据值\n";
    fs::write(
        root.join("s3_previous.csv"),
        format!("{headers}C001,2026-07-31,ORG1,模拟机构,R001,模拟地区,1,I001,模拟贷款余额,余额,CNY,月报,第1批,100\n"),
    )?;
    fs::write(
        root.join("s3_current_normal.csv"),
        format!("{headers}C001,2026-08-31,ORG1,模拟机构,R001,模拟地区,1,I001,模拟贷款余额,余额,CNY,月报,第1批,100\n"),
    )?;
    fs::write(
        root.join("s3_current_exception.csv"),
        format!("{headers}C001,2026-08-31,ORG1,模拟机构,R001,模拟地区,1,I001,模拟贷款余额,余额,CNY,月报,第1批,120\n"),
    )?;
    Ok(())
}

type S3F02InputRow<'a> = (
    &'a str,
    &'a str,
    &'a str,
    &'a str,
    &'a str,
    &'a str,
    f64,
    f64,
    f64,
    &'a str,
    &'a str,
);

fn write_s3_f02_comparison(path: &Path, rows: &[S3F02InputRow<'_>]) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("比较结果")?;
    let headers = [
        "机构类代码",
        "机构类名称",
        "地区代码",
        "地区名称",
        "指标代码",
        "指标名称",
        "数据值",
        "上期值",
        "增减额",
        "输出说明",
        "说明内容",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (index, row) in rows.iter().enumerate() {
        let output_row = index as u32 + 1;
        let (org_code, org_name, region_code, region_name, indicator_code, indicator_name) =
            (row.0, row.1, row.2, row.3, row.4, row.5);
        for (column, value) in [
            org_code,
            org_name,
            region_code,
            region_name,
            indicator_code,
            indicator_name,
        ]
        .iter()
        .enumerate()
        {
            sheet.write_string(output_row, column as u16, *value)?;
        }
        sheet.write_number(output_row, 6, row.6)?;
        sheet.write_number(output_row, 7, row.7)?;
        sheet.write_number(output_row, 8, row.8)?;
        sheet.write_string(output_row, 9, row.9)?;
        sheet.write_string(output_row, 10, row.10)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn write_s3_f02_common_config(path: &Path) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    write_sheet(
        &mut workbook,
        "运行参数",
        &["参数", "值"],
        &[&["默认源数据单位", "元"], &["默认目标单位", "元"]],
    )?;
    write_sheet(
        &mut workbook,
        "机构地区参照",
        &["机构类代码", "地区代码", "机构名称"],
        &[
            &["ORG1001", "R001", "配置示例银行"],
            &["ORG2002", "R001", "多地区银行东区"],
            &["ORG2002", "R002", "多地区银行西区"],
        ],
    )?;
    workbook.save(path)?;
    Ok(())
}

fn generate_s3_f02_samples(root: &Path) -> Result<(), Box<dyn Error>> {
    if root.exists() {
        return Err(format!(
            "refusing to overwrite existing S3-F02 sample directory: {}",
            root.display()
        )
        .into());
    }
    fs::create_dir_all(root.join("input/normal"))?;
    fs::create_dir_all(root.join("input/multi_region"))?;
    fs::create_dir_all(root.join("outputs/normal"))?;
    fs::create_dir_all(root.join("outputs/multi_region"))?;
    write_s3_f02_common_config(&root.join("s3_f02_common_config.xlsx"))?;
    write_s3_f02_comparison(
        &root.join("input/normal/comparison_normal.xlsx"),
        &[
            (
                "ORG1001",
                "原始机构名",
                "R001",
                "东城区",
                "20202001",
                "贷款余额",
                1500.0,
                1400.0,
                100.0,
                "是",
                "示例说明：贷款余额增加100万元。",
            ),
            (
                "ORG1001",
                "原始机构名",
                "R001",
                "东城区",
                "20202002",
                "存款余额",
                800.0,
                820.0,
                -20.0,
                "否",
                "此行未勾选，应排除。",
            ),
        ],
    )?;
    write_s3_f02_comparison(
        &root.join("input/multi_region/comparison_multi_region.xlsx"),
        &[
            (
                "ORG2002",
                "多地区银行",
                "R001",
                "东城区",
                "20203001",
                "东区余额",
                260.0,
                250.0,
                10.0,
                "是",
                "东区余额说明。",
            ),
            (
                "ORG2002",
                "多地区银行",
                "R002",
                "西城区",
                "20203002",
                "西区余额",
                340.0,
                330.0,
                10.0,
                "✓",
                "西区余额说明。",
            ),
        ],
    )?;
    println!("generated synthetic S3-F02 cases at {}", root.display());
    Ok(())
}

fn value_text(sheet: &SheetValues, address: &str) -> Option<String> {
    match sheet.values.get(address)? {
        CellValue::Text(value) => Some(value.clone()),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => Some(value.to_string()),
        CellValue::Integer(value) => Some(value.to_string()),
        CellValue::Boolean(value) => Some(value.to_string()),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => Some(value.clone()),
        CellValue::Error(value) => Some(value.clone()),
        CellValue::Empty => None,
    }
}

fn verify_s3_f02_outputs(root: &Path) -> Result<(), Box<dyn Error>> {
    let normal = read_value_columns(root.join("outputs/normal/s3_f02_normal.xlsx"), usize::MAX)?;
    let normal = normal
        .iter()
        .find(|sheet| sheet.name == "说明文件")
        .ok_or("S3-F02 normal output sheet missing")?;
    if normal.max_row != 3
        || value_text(normal, "A3").as_deref() != Some("ORG1001")
        || value_text(normal, "B3").as_deref() != Some("配置示例银行（东城区）")
        || value_text(normal, "C3").as_deref() != Some("20202001")
        || value_text(normal, "E3").as_deref() != Some("1500")
        || value_text(normal, "G3").as_deref() != Some("100")
        || value_text(normal, "H3").as_deref() != Some("示例说明：贷款余额增加100万元。")
    {
        return Err("S3-F02 normal output did not match its independently specified row".into());
    }

    let multi_region = read_value_columns(
        root.join("outputs/multi_region/s3_f02_multi_region.xlsx"),
        usize::MAX,
    )?;
    let multi_region = multi_region
        .iter()
        .find(|sheet| sheet.name == "说明文件")
        .ok_or("S3-F02 multi-region output sheet missing")?;
    if multi_region.max_row != 4
        || value_text(multi_region, "A3").as_deref() != Some("ORG2002R001")
        || value_text(multi_region, "B3").as_deref() != Some("多地区银行东区（东城区）")
        || value_text(multi_region, "A4").as_deref() != Some("ORG2002R002")
        || value_text(multi_region, "B4").as_deref() != Some("多地区银行西区（西城区）")
        || value_text(multi_region, "H4").as_deref() != Some("西区余额说明。")
    {
        return Err(
            "S3-F02 multi-region output did not match its independently specified rows".into(),
        );
    }
    println!("verified: S3-F02 normal 1 selected row; multi-region 2 distinct rows");
    Ok(())
}

fn write_s3_f04_config(root: &Path) -> Result<(), Box<dyn Error>> {
    let mut common = Workbook::new();
    write_sheet(
        &mut common,
        "运行参数",
        &["参数", "值"],
        &[&["默认源数据单位", "元"], &["默认目标单位", "元"]],
    )?;
    write_sheet(&mut common, "单位换算例外", &["指标代码", "启用"], &[])?;
    common.save(root.join("s3_f04_common_config.xlsx"))?;

    let mut feature = Workbook::new();
    // The legacy Python engine treats every nonempty flag (including “否”) as enabled.
    // Keep false inversion/logical/disabled flags blank in this synthetic configuration.
    write_sheet(
        &mut feature,
        "跨期数值核对",
        &[
            "校验编码",
            "校验名称",
            "机构类代码",
            "地区代码",
            "前提条件",
            "left",
            "right",
            "校验公式",
            "校验提示",
            "取反标识",
            "无误是否提示",
            "是否逻辑校验",
            "禁用",
            "备注",
        ],
        &[&[
            "DEMO-S3-F04-DIMENSION",
            "同一期不同批次余额核对",
            "^ORG1$",
            "^REG0001$",
            "",
            "[I001,余额,CNY,月报,第1批]",
            "[I001,余额,CNY,月报,第2批]",
            "left > right",
            "第一批余额高于第二批，请核实",
            "",
            "是",
            "",
            "",
            "合成验收：频度与批次分别保留，不得合并。",
        ]],
    )?;
    feature.save(root.join("s3_f04_feature_config.xlsx"))?;
    Ok(())
}

fn write_s3_f04_current(path: &Path, first_batch_value: f64) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("本期导数")?;
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
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    for (row, batch, value) in [(1_u32, "第1批", first_batch_value), (2_u32, "第2批", 100.0)] {
        let text = [
            "人民币",
            "2026-08-31",
            "ORG1",
            "模拟机构",
            "REG0001",
            "模拟地区",
            "1",
            "I001",
            "模拟贷款余额",
            "余额",
            "CNY",
            "月报",
            batch,
        ];
        for (column, value) in text.iter().enumerate() {
            sheet.write_string(row, column as u16, *value)?;
        }
        sheet.write_number(row, 13, value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn generate_s3_f04_samples(root: &Path) -> Result<(), Box<dyn Error>> {
    if root.exists() {
        return Err(format!(
            "refusing to overwrite existing S3-F04 sample directory: {}",
            root.display()
        )
        .into());
    }
    fs::create_dir_all(root)?;
    fs::create_dir(root.join("outputs"))?;
    write_s3_f04_config(root)?;
    write_s3_f04_current(&root.join("s3_f04_current_normal.xlsx"), 100.0)?;
    write_s3_f04_current(&root.join("s3_f04_current_exception.xlsx"), 120.0)?;
    println!(
        "generated synthetic S3-F04 XLSX cases at {}",
        root.display()
    );
    Ok(())
}

fn write_s1_f02_template(path: &Path) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("本地校验结果")?;
    for (column, header) in [
        "批次",
        "表单名称",
        "规则编号",
        "规则类型",
        "规则描述",
        "校验字段",
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(1, column as u16 + 1, *header)?;
    }
    workbook.define_name(
        "任意行汇总区域",
        "='本地校验结果'!$B$3:$G$10",
    )?;
    workbook.define_name("表头区域", "='本地校验结果'!$B$2:$G$2")?;
    workbook.save(path)?;
    Ok(())
}

fn write_s1_f02_source(
    path: &Path,
    rows: &[&[&str]],
    mismatched_header: bool,
) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("本地校验结果")?;
    let headers = [
        "批次",
        "表单名称",
        "规则编号",
        if mismatched_header {
            "错误规则类型"
        } else {
            "规则类型"
        },
        "规则描述",
        "校验字段",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(1, column as u16 + 1, *header)?;
    }
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, value) in row.iter().enumerate() {
            sheet.write_string(row_index as u32 + 2, column_index as u16 + 1, *value)?;
        }
    }
    workbook.save(path)?;
    Ok(())
}

fn write_s1_f02_history(path: &Path) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("本地校验结果")?;
    let headers = [
        "来源文件",
        "来源工作表",
        "批次",
        "表单名称",
        "规则编号",
        "规则类型",
        "规则描述",
        "校验字段",
        "历史触发条数",
        "历史分类说明",
        "审核结果",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    let rows = [
        ("甲机构", 1.0, 10.0, "类别甲", "通过"),
        ("甲机构", 1.0, 12.0, "类别乙", "未通过"),
        ("甲机构", 2.0, 4.0, "边界复核", "通过"),
        ("乙机构", 2.0, 8.0, "异常复核", "复查"),
    ];
    for (row_index, (source, rule_id, count, category, result)) in rows.iter().enumerate() {
        let row = row_index as u32 + 1;
        for (column, value) in [
            *source,
            "本地校验结果",
            "2026-07",
            "存量单位贷款",
            "",
            "阈值",
            "不应低于",
            "贷款余额",
            "",
            *category,
            *result,
        ]
        .iter()
        .enumerate()
        {
            if column != 4 && column != 8 {
                sheet.write_string(row, column as u16, *value)?;
            }
        }
        sheet.write_number(row, 4, *rule_id)?;
        sheet.write_number(row, 8, *count)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn generate_s1_f02_samples(root: &Path) -> Result<(), Box<dyn Error>> {
    if root.exists() {
        return Err(format!(
            "refusing to overwrite existing sample directory: {}",
            root.display()
        )
        .into());
    }
    let normal_input = root.join("input/normal");
    let exception_input = root.join("input/structure_exception");
    fs::create_dir_all(&normal_input)?;
    fs::create_dir_all(&exception_input)?;
    fs::create_dir_all(root.join("outputs/normal"))?;
    fs::create_dir_all(root.join("outputs/structure_exception"))?;
    write_s1_f02_template(&root.join("template.xlsx"))?;
    write_s1_f02_history(&root.join("history.xlsx"))?;

    let alpha_rows = [
        &[
            "2026-07",
            "存量单位贷款",
            "1.0",
            "阈值",
            "不应低于",
            "贷款余额",
        ][..],
        &[
            "2026-07",
            "存量单位贷款",
            "2",
            "阈值",
            "不应低于",
            "贷款余额",
        ][..],
    ];
    let beta_rows = [
        &[
            "2026-07",
            "存量单位贷款",
            "1",
            "阈值",
            "不应低于",
            "贷款余额",
        ][..],
        &[
            "2026-07",
            "存量单位贷款",
            "2",
            "阈值",
            "不应低于",
            "贷款余额",
        ][..],
        &[
            "2026-07",
            "存量单位贷款",
            "3",
            "比例",
            "不应超过",
            "不良率",
        ][..],
    ];
    for input_dir in [&normal_input, &exception_input] {
        write_s1_f02_source(&input_dir.join("甲机构.xlsx"), &alpha_rows, false)?;
        write_s1_f02_source(&input_dir.join("乙机构.xlsx"), &beta_rows, false)?;
    }
    write_s1_f02_source(
        &exception_input.join("结构异常.xlsx"),
        &alpha_rows,
        true,
    )?;
    println!("generated synthetic S1-F02 samples at {}", root.display());
    Ok(())
}

fn sample_text<'a>(sheet: &'a excel_reader::SheetDump, address: &str) -> Option<&'a str> {
    sheet.values.get(address).and_then(|value| value["value"].as_str())
}

fn verify_s1_f02_outputs(root: &Path) -> Result<(), Box<dyn Error>> {
    for relative in [
        "outputs/normal/s1_f02_summary.xlsx",
        "outputs/structure_exception/s1_f02_summary.xlsx",
    ] {
        let workbook = excel_reader::read_workbook(root.join(relative))?;
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "本地校验结果")
            .ok_or_else(|| format!("S1-F02 result sheet missing: {relative}"))?;
        for address in ["A1", "K1", "A2", "A6"] {
            if sample_text(sheet, address).is_none() {
                return Err(format!("S1-F02 output missing {address}: {relative}").into());
            }
        }
        println!("readable: {relative}");
    }
    let normal = excel_reader::read_workbook(root.join("outputs/normal/s1_f02_summary.xlsx"))?;
    let sheet = normal
        .sheets
        .iter()
        .find(|sheet| sheet.name == "本地校验结果")
        .expect("checked S1-F02 normal result sheet");
    let alpha_rule_one = (2..=6).find(|row| {
        sample_text(sheet, &format!("A{row}")) == Some("甲机构")
            && sample_text(sheet, &format!("E{row}")) == Some("1.0")
    });
    let Some(row) = alpha_rule_one else {
        return Err("S1-F02 output is missing 甲机构 rule 1.0".into());
    };
    if sample_text(sheet, &format!("I{row}")) != Some("10|12")
        || sample_text(sheet, &format!("J{row}")) != Some("类别甲|类别乙")
        || sample_text(sheet, &format!("K{row}")) != Some("通过|未通过")
    {
        return Err("S1-F02 normal output history enrichment differs from independent expectation".into());
    }
    println!("verified: S1-F02 normal output has 5 rows and expected history enrichment");
    Ok(())
}

fn verify_outputs(root: &Path) -> Result<(), Box<dyn Error>> {
    let files = [
        "outputs/s3_normal.xlsx",
        "outputs/s3_exception.xlsx",
    ];
    for relative in files {
        let path = root.join(relative);
        let workbook = excel_reader::read_workbook(&path)?;
        if workbook.sheets.is_empty() {
            return Err(format!("workbook contains no readable sheets: {}", path.display()).into());
        }
        println!("readable: {} ({} sheets)", relative, workbook.sheets.len());
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = env::args().skip(1);
    let command_or_output = args.next().ok_or(
        "usage: cargo run -p audit-core --example manual_acceptance_samples -- <new-output-directory> | --s1-f02 <new-sample-directory> | --s3-f02 <new-sample-directory> | --s3-f04 <new-output-directory> | --verify <sample-directory> | --verify-s1-f02 <sample-directory> | --verify-s3-f02 <sample-directory>",
    )?;
    if command_or_output == "--s1-f02" {
        let path = args
            .next()
            .ok_or("--s1-f02 requires a new sample directory")?;
        return generate_s1_f02_samples(Path::new(&path));
    }
    if command_or_output == "--verify-s1-f02" {
        let path = args
            .next()
            .ok_or("--verify-s1-f02 requires a sample directory")?;
        return verify_s1_f02_outputs(Path::new(&path));
    }
    if command_or_output == "--s3-f02" {
        let path = args
            .next()
            .ok_or("--s3-f02 requires a new sample directory")?;
        return generate_s3_f02_samples(Path::new(&path));
    }
    if command_or_output == "--verify-s3-f02" {
        let path = args
            .next()
            .ok_or("--verify-s3-f02 requires a sample directory")?;
        return verify_s3_f02_outputs(Path::new(&path));
    }
    if command_or_output == "--s3-f04" {
        let path = args
            .next()
            .ok_or("--s3-f04 requires a new sample directory")?;
        return generate_s3_f04_samples(Path::new(&path));
    }
    if command_or_output == "--verify" {
        let path = args.next().ok_or("--verify requires a sample directory")?;
        return verify_outputs(Path::new(&path));
    }
    let output = command_or_output;
    let root = Path::new(&output);
    if root.exists() {
        return Err(format!(
            "refusing to overwrite existing sample directory: {}",
            root.display()
        )
        .into());
    }
    fs::create_dir_all(root)?;
    write_s3_configs(root)?;
    write_s3_csvs(root)?;
    println!(
        "generated synthetic Audit acceptance inputs at {}",
        root.display()
    );
    Ok(())
}
