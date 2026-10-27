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

fn write_s5_config(root: &Path) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    write_sheet(
        &mut workbook,
        "审核规则",
        &[
            "规则ID",
            "规则修订",
            "是否启用",
            "业务类型",
            "默认数据源",
            "默认Sheet",
            "规则名称",
            "审核级别",
            "校验类型",
            "校验指标",
            "审核表达式",
            "表达式布尔值取反",
            "规则描述",
            "当前值表达式",
            "对比值表达式",
            "差值表达式",
            "取值字典",
            "生效起始日",
            "生效截止日",
            "适用范围",
            "适用表达式",
        ],
        &[&[
            "DEMO-S5-001",
            "1",
            "是",
            "单位贷款",
            "CUR",
            "贷款表",
            "期末余额偏离演示值",
            "错误",
            "数值校验",
            "经营贷款余额",
            "ABS({本期} - 120) <= 1",
            "否",
            "本期余额在120万元上下1万元内通过，超出范围时触发。",
            "{本期}",
            "120",
            "{本期} - 120",
            r#"{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期"))}"#,
            "",
            "",
            "",
            "",
        ]],
    )?;
    write_sheet(
        &mut workbook,
        "模板指标目录",
        &[
            "业务类型",
            "工作表名称",
            "行路径",
            "列路径",
            "行表头标题",
            "数据类型",
            "指标类型",
            "来源单位",
            "标准单位",
            "存储尺度",
            "允许空值",
            "精度",
            "口径依据",
            "确认状态",
            "布局版本",
        ],
        &[&[
            "单位贷款",
            "贷款表",
            "贷款产品类别\n经营贷款",
            "余额\n本期",
            "一级\n二级",
            "number",
            "amount",
            "万元",
            "万元",
            "native",
            "否",
            "2",
            "手工模拟值及明确阈值",
            "是",
            "manual-v1",
        ]],
    )?;
    write_sheet(
        &mut workbook,
        "配置版本",
        &[
            "配置版本号",
            "状态",
            "发布日期",
            "生效起始日",
            "生效截止日",
            "发布说明",
        ],
        &[&[
            "manual-v1",
            "启用",
            "2026-10-27",
            "",
            "",
            "模拟人工验收配置",
        ]],
    )?;
    workbook.save(root.join("s5_config.xlsx"))?;
    Ok(())
}

fn write_s5_source(path: &Path, current: f64) -> Result<(), Box<dyn Error>> {
    write_s5_source_without_conditional_format(path, current)
}

fn write_s5_source_with_conditional_format(
    path: &Path,
    current: f64,
    add_conditional_format: bool,
) -> Result<(), Box<dyn Error>> {
    if !add_conditional_format {
        return write_s5_source_without_conditional_format(path, current);
    }

    let mut workbook = umya_spreadsheet::new_file();
    let basic = workbook.get_sheet_by_name_mut("Sheet1")?;
    basic.set_name("基本信息");
    basic.get_cell_mut("A1").set_value("机构代码");
    basic.get_cell_mut("B1").set_value("I001");
    basic.get_cell_mut("A2").set_value("统计期间");
    basic.get_cell_mut("B2").set_value("2026-08-31");

    let data = workbook.new_sheet("贷款表")?;
    data.get_cell_mut("A1").set_value("模拟单位贷款数据");
    data.get_cell_mut("C2").set_value("余额");
    data.get_cell_mut("A3").set_value("一级");
    data.get_cell_mut("B3").set_value("二级");
    data.get_cell_mut("C3").set_value("本期");
    data.get_cell_mut("D3").set_value("上期");
    data.get_cell_mut("A4").set_value("贷款产品类别");
    data.get_cell_mut("B4").set_value("经营贷款");
    data.get_cell_mut("C4").set_value_number(current);
    data.get_cell_mut("D4").set_value_number(100.0);

    let mut formula = umya_spreadsheet::Formula::default();
    formula.set_string_value("C4>120");
    let mut rule = umya_spreadsheet::ConditionalFormattingRule::default();
    rule.set_type(umya_spreadsheet::ConditionalFormatValues::Expression)
        .set_priority(1)
        .set_formula(formula);
    let mut conditional_format = umya_spreadsheet::ConditionalFormatting::default();
    conditional_format
        .sequence_of_references_mut()
        .set_sqref("C4");
    conditional_format.add_conditional_collection(rule);
    workbook
        .get_sheet_by_name_mut("贷款表")?
        .add_conditional_formatting_collection(conditional_format);
    umya_spreadsheet::writer::xlsx::write(&workbook, path)?;
    Ok(())
}

fn write_s5_source_without_conditional_format(
    path: &Path,
    current: f64,
) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let info = workbook.add_worksheet();
    info.set_name("基本信息")?;
    info.write_string(0, 0, "机构代码")?;
    info.write_string(0, 1, "I001")?;
    info.write_string(1, 0, "统计期间")?;
    info.write_string(1, 1, "2026-08-31")?;

    let data = workbook.add_worksheet();
    data.set_name("贷款表")?;
    data.write_string(0, 0, "模拟单位贷款数据")?;
    data.write_string(1, 2, "余额")?;
    data.write_string(2, 0, "一级")?;
    data.write_string(2, 1, "二级")?;
    data.write_string(2, 2, "本期")?;
    data.write_string(2, 3, "上期")?;
    data.write_string(3, 0, "贷款产品类别")?;
    data.write_string(3, 1, "经营贷款")?;
    data.write_number(3, 2, current)?;
    data.write_number(3, 3, 100.0)?;
    workbook.save(path)?;
    Ok(())
}

fn write_s5_gui_layout_template(path: &Path) -> Result<(), Box<dyn Error>> {
    if path.exists() {
        return Err(format!(
            "refusing to overwrite existing S5 GUI template: {}",
            path.display()
        )
        .into());
    }
    let mut workbook = Workbook::new();
    let info = workbook.add_worksheet();
    info.set_name("基本信息")?;
    info.write_string(0, 0, "机构代码")?;
    info.write_string(0, 1, "I001")?;
    info.write_string(1, 0, "统计期间")?;
    info.write_string(1, 1, "2026-08-31")?;

    let data = workbook.add_worksheet();
    data.set_name("贷款表")?;
    data.write_string(0, 0, "模拟单位贷款数据")?;
    data.write_string(1, 2, "余额")?;
    data.write_string(2, 0, "一级")?;
    data.write_string(2, 1, "二级")?;
    data.write_string(2, 2, "本期")?;
    data.write_string(2, 3, "上期")?;
    data.write_string(3, 0, "贷款产品类别")?;
    data.write_string(3, 1, "经营贷款")?;
    data.write_number(3, 2, 120.0)?;
    data.write_number(3, 3, 100.0)?;
    workbook.define_name("行索引区域_1", "='贷款表'!$A$4:$B$4")?;
    workbook.define_name("列索引区域_1", "='贷款表'!$C$3:$D$3")?;
    workbook.define_name("数字区域_1", "='贷款表'!$C$4:$D$4")?;
    workbook.save(path)?;
    Ok(())
}

fn generate_s5_gui_template(template_dir: &Path) -> Result<(), Box<dyn Error>> {
    if !template_dir.is_dir() {
        return Err(format!(
            "S5 GUI template directory does not exist: {}",
            template_dir.display()
        )
        .into());
    }
    let path = template_dir.join("单位贷款_合成布局模板.xlsx");
    write_s5_gui_layout_template(&path)?;
    println!("generated S5 GUI layout template at {}", path.display());
    Ok(())
}

fn generate_s5_conditional_samples(root: &Path) -> Result<(), Box<dyn Error>> {
    if root.exists() {
        return Err(format!(
            "refusing to overwrite existing sample directory: {}",
            root.display()
        )
        .into());
    }
    fs::create_dir_all(root.join("s5_templates"))?;
    write_s5_config(root)?;
    write_s5_source_with_conditional_format(&root.join("s5_source_normal.xlsx"), 120.0, true)?;
    write_s5_source_with_conditional_format(&root.join("s5_source_exception.xlsx"), 125.0, true)?;
    println!(
        "generated S5 conditional-format samples at {}",
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
        "outputs/s5_normal_corrected/demo-s5-normal-corrected/审核结果.xlsx",
        "outputs/s5_exception/demo-s5-exception/审核结果.xlsx",
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

fn verify_s5_conditional_outputs(root: &Path) -> Result<(), Box<dyn Error>> {
    let cases = [
        (
            "outputs/s5_conditional_normal_ready/s5-cf-normal-ready-20261027/审核结果.xlsx",
            false,
        ),
        (
            "outputs/s5_conditional_exception_ready/s5-cf-exception-ready-20261027/审核结果.xlsx",
            true,
        ),
    ];
    for (relative, expect_trigger) in cases {
        let path = root.join(relative);
        let workbook = excel_reader::read_workbook(&path)?;
        let conditional = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "条件格式触发")
            .ok_or_else(|| format!("conditional-format sheet missing: {}", path.display()))?;
        let statuses = conditional
            .values
            .iter()
            .filter(|(address, _)| {
                address
                    .strip_prefix('G')
                    .and_then(|row| row.parse::<usize>().ok())
                    .is_some_and(|row| row >= 2)
            })
            .filter_map(|(_, value)| value["value"].as_str())
            .collect::<Vec<_>>();
        if expect_trigger {
            if statuses.as_slice() != ["TRIGGERED"] {
                return Err(format!(
                    "expected one conditional-format trigger in {} but read {:?}",
                    path.display(),
                    statuses
                )
                .into());
            }
            for (address, expected) in [
                ("D2", "贷款表"),
                ("E2", "C4"),
                ("F2", "125"),
                ("G2", "TRIGGERED"),
                ("I2", "C4>120"),
            ] {
                let actual = conditional
                    .values
                    .get(address)
                    .and_then(|value| value["value"].as_str())
                    .unwrap_or_default();
                if actual != expected {
                    return Err(format!(
                        "unexpected {address} in {}: expected {expected:?}, got {actual:?}",
                        path.display()
                    )
                    .into());
                }
            }
        } else if !statuses.is_empty() {
            return Err(format!(
                "normal sample unexpectedly triggered conditional formatting in {}: {:?}",
                path.display(),
                statuses
            )
            .into());
        }
        println!(
            "verified: {} (conditional trigger: {})",
            relative, expect_trigger
        );
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = env::args().skip(1);
    let command_or_output = args.next().ok_or(
        "usage: cargo run -p audit-core --example manual_acceptance_samples -- <new-output-directory> | --s1-f02 <new-sample-directory> | --s3-f02 <new-sample-directory> | --s3-f04 <new-output-directory> | --s5-conditional <new-output-directory> | --s5-gui-template <existing-template-directory> | --verify <sample-directory> | --verify-s1-f02 <sample-directory> | --verify-s3-f02 <sample-directory> | --verify-s5-conditional <sample-directory>",
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
    if command_or_output == "--s5-conditional" {
        let path = args
            .next()
            .ok_or("--s5-conditional requires a new sample directory")?;
        return generate_s5_conditional_samples(Path::new(&path));
    }
    if command_or_output == "--s5-gui-template" {
        let path = args
            .next()
            .ok_or("--s5-gui-template requires an existing template directory")?;
        return generate_s5_gui_template(Path::new(&path));
    }
    if command_or_output == "--verify" {
        let path = args.next().ok_or("--verify requires a sample directory")?;
        return verify_outputs(Path::new(&path));
    }
    if command_or_output == "--verify-s5-conditional" {
        let path = args
            .next()
            .ok_or("--verify-s5-conditional requires a sample directory")?;
        return verify_s5_conditional_outputs(Path::new(&path));
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
    write_s5_config(root)?;
    write_s5_source(&root.join("s5_source_normal.xlsx"), 120.0)?;
    write_s5_source(&root.join("s5_source_exception.xlsx"), 125.0)?;
    println!(
        "generated synthetic Audit acceptance inputs at {}",
        root.display()
    );
    Ok(())
}
