//! Generate a compact, user-operable S2-F01 external-comparison sample.
//!
//! The report inputs are existing synthetic workbooks from this workspace;
//! the S2 configuration and external workbooks are generated with the Rust
//! writer. No source-project or real business data is used.

use anyhow::{Context, Result, ensure};
use excel_writer::{Workbook, Worksheet};
use std::{env, fs, path::PathBuf};

const FEATURE_DIR: &str = "manual_acceptance/20261027/s2_f01";
const OUTPUT_DIR: &str = "manual_acceptance/20261027/s2_external_hit";

fn main() -> Result<()> {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let root = env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| manifest.join(OUTPUT_DIR));
    ensure!(
        !root.exists(),
        "S2 外部核对样例目录已存在，拒绝覆盖：{}",
        root.display()
    );

    let fixture_root = manifest.join(FEATURE_DIR);
    let current_source = fixture_root.join("input/current");
    let previous_source = fixture_root.join("input/previous");
    let current_target = root.join("input/current");
    let previous_target = root.join("input/previous");
    fs::create_dir_all(&current_target)?;
    fs::create_dir_all(&previous_target)?;
    copy_single_xlsx(&current_source, &current_target)?;
    copy_single_xlsx(&previous_source, &previous_target)?;

    write_config(&root.join("s2_external_config.xlsx"))?;
    write_external(&root.join("external_normal.xlsx"), 1_234_567.89)?;
    write_external(&root.join("external_mismatch.xlsx"), 1_234_500.0)?;
    fs::create_dir_all(root.join("outputs/normal"))?;
    fs::create_dir_all(root.join("outputs/external_mismatch"))?;
    fs::write(
        root.join("expected-results.txt"),
        r#"S2-F01 外部核对样例独立预期（按输入和配置计算）

配置：指标 20202001（存款余额）使用规则 MANUAL-EXT-001；比较方式“等于”，容差 0，级别“硬性”。本期源数据值为 1,234,567.89 元。源/外部/输出单位均为元。

正常：外部值 1,234,567.89；绝对差 0，规则不触发。
异常：外部值 1,234,500.00；绝对差 67.89，大于容差 0；规则 MANUAL-EXT-001 应触发 1 条硬性外部核对异常，报表项目为“存款项目”。

输入使用 workspace 自有合成 S2 报表，另由本例以 Rust excel_writer 生成配置与外部工作簿。预期由上述数值和配置独立推导，不从 Rust 输出复制。
        "#,
    )?;

    println!("生成 S2-F01 外部核对样例：{}", root.display());
    Ok(())
}

fn copy_single_xlsx(source_dir: &PathBuf, target_dir: &PathBuf) -> Result<()> {
    let mut files = fs::read_dir(source_dir)
        .with_context(|| format!("读取合成报表目录：{}", source_dir.display()))?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.is_file()
                && path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
        })
        .collect::<Vec<_>>();
    files.sort();
    ensure!(
        files.len() == 1,
        "预期合成报表目录恰有一个 XLSX，实际 {}：{}",
        files.len(),
        source_dir.display()
    );
    let source = &files[0];
    let target = target_dir.join(
        source
            .file_name()
            .context("合成报表文件名为空")?,
    );
    fs::copy(source, &target).with_context(|| {
        format!(
            "复制合成报表到独立样例目录：{} -> {}",
            source.display(),
            target.display()
        )
    })?;
    Ok(())
}

fn write_text_sheet(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[&[&str]],
) -> Result<()> {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name)?;
    write_headers(sheet, headers)?;
    for (row_index, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            sheet.write_string(row_index as u32 + 1, column as u16, *value)?;
        }
    }
    Ok(())
}

fn write_headers(sheet: &mut Worksheet, headers: &[&str]) -> Result<()> {
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header)?;
    }
    Ok(())
}

fn write_config(path: &std::path::Path) -> Result<()> {
    let mut workbook = Workbook::new();
    write_text_sheet(
        &mut workbook,
        "通用设置",
        &["设置项", "设置值", "说明"],
        &[
            &["源数据单位", "元", "合成样例"],
            &["外部文件单位", "元", "合成样例"],
            &["输出文件单位", "元", "合成样例"],
            &["是否隐藏无变动指标", "否", "保留正常结果"],
            &["是否隐藏与大集中一致指标", "否", "保留正常结果"],
            &["文字变动默认级别", "硬性", ""],
            &["本期缺失默认级别", "硬性", ""],
            &["上期缺失默认级别", "硬性", ""],
        ],
    )?;
    write_text_sheet(
        &mut workbook,
        "指标参照",
        &[
            "指标代码",
            "指标名称",
            "表单代码",
            "数据属性",
            "环比启用",
            "环比策略组",
            "最小变动值",
            "禁用",
        ],
        &[&["20202001", "存款余额", "20202", "余额", "否", "", "0", "否"]],
    )?;
    write_text_sheet(
        &mut workbook,
        "机构参照",
        &[
            "机构名称",
            "机构代码",
            "机构类别",
            "承接行",
            "归属行",
            "报表项目",
            "禁用",
        ],
        &[&["甲银行", "9144", "银行", "", "", "存款项目", "否"]],
    )?;
    write_text_sheet(
        &mut workbook,
        "环比策略",
        &[
            "策略组",
            "规则编号",
            "数据属性",
            "下限",
            "上限",
            "审核级别",
            "分类说明",
            "禁用",
        ],
        &[],
    )?;
    write_text_sheet(
        &mut workbook,
        "外部核对规则",
        &[
            "规则编号",
            "指标代码",
            "比较方式",
            "外部指标名称",
            "容差",
            "级别",
            "启用",
        ],
        &[&[
            "MANUAL-EXT-001",
            "20202001",
            "等于",
            "存款余额",
            "0",
            "硬性",
            "是",
        ]],
    )?;
    write_text_sheet(
        &mut workbook,
        "校验规则",
        &[
            "规则编号",
            "类型",
            "描述",
            "规则内容",
            "Thd值(万元)",
            "取反",
            "级别",
            "禁用",
            "备注",
        ],
        &[],
    )?;
    workbook.save(path)?;
    Ok(())
}

fn write_external(path: &std::path::Path, value: f64) -> Result<()> {
    let mut workbook = Workbook::new();
    let data = workbook.add_worksheet();
    data.set_name("集中系统数据")?;
    write_headers(data, &["报表项目", "存款余额"])?;
    data.write_string(1, 0, "存款项目")?;
    data.write_number(1, 1, value)?;

    let mapping = workbook.add_worksheet();
    mapping.set_name("机构参照表")?;
    write_headers(
        mapping,
        &["统一社会信用代码", "报表项目", "禁用"],
    )?;
    mapping.write_string(1, 0, "9144")?;
    mapping.write_string(1, 1, "存款项目")?;
    mapping.write_string(1, 2, "否")?;
    workbook.save(path)?;
    Ok(())
}
