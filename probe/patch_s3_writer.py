import io

p = 'audit-core/src/s3_execution_compare.rs'
s = io.open(p, encoding='utf-8').read()

old = "use excel_writer::{Color, Format, FormatAlign, Workbook, Worksheet};"
new = "use excel_writer::{Color, ExcelDateTime, Format, FormatAlign, Workbook, Worksheet};"
assert s.count(old) == 1
s = s.replace(old, new)

old_loop = """            for (index, row) in rows.iter().enumerate() {
                let output_row = index as u32 + 1;
                let record = &row.record;
                let fill_color = row
                    .fill_color
                    .filter(|_| !row.fill_scope.is_empty())
                    .and_then(vba_color_index_rgb);
                // 指标名称显示列：设置开启时去首尾空格（不影响匹配与说明文字）。
                let display_indicator_name = if row.indicator_name_out.is_empty() {
                    record.indicator_name.as_str()
                } else {
                    row.indicator_name_out.as_str()
                };
                let display_indicator_name = if trim_indicator_names {
                    display_indicator_name.trim()
                } else {
                    display_indicator_name
                };
                let text_values = [
                    (0, record.biz_class.as_str()),
                    (1, record.record_date.as_str()),
                    (2, record.org_code.as_str()),
                    (3, record.org_name.as_str()),
                    (4, record.region_code.as_str()),
                    (5, record.region_name.as_str()),
                    (
                        6,
                        if row.order_code_out.is_empty() {
                            record.indicator.as_str()
                        } else {
                            row.order_code_out.as_str()
                        },
                    ),
                    (7, record.indicator.as_str()),
                    (8, display_indicator_name),
                    (9, record.data_attr.as_str()),
                    (10, record.currency.as_str()),
                    (11, record.frequency.as_str()),
                    (12, record.batch.as_str()),
                    (17, row.remark.as_str()),
                    (18, row.need_explanation.as_str()),
                    (19, row.explanation_content.as_str()),
                    (20, row.calculation_process.as_str()),
                ];
                for (column, value) in text_values {
                    let applies_fill = fill_color.is_some()
                        && ((row.fill_scope == "row" && column <= 17)
                            || (row.fill_scope != "row" && column == 17));
                    if column == 1 || applies_fill {
                        let format = if column == 1 {
                            if applies_fill {
                                formats
                                    .date_by_fill
                                    .get(&fill_color.expect("fill applies to date"))
                                    .expect("date fill format was precomputed")
                            } else {
                                &formats.date
                            }
                        } else {
                            formats
                                .text_by_fill
                                .get(&fill_color.expect("fill applies to text"))
                                .expect("text fill format was precomputed")
                        };
                        worksheet.write_string_with_format(output_row, column, value, format)?;
                    } else {
                        worksheet.write_string(output_row, column, value)?;
                    }
                }
                write_central_value(
                    worksheet,
                    output_row,
                    13,
                    &record.value,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    14,
                    &row.previous_value,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    15,
                    &row.change,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    16,
                    &row.ratio,
                    Some("0.00%"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                worksheet.write_blank(output_row, 21, &formats.blank)?;
            }"""

new_loop = """            for (index, row) in rows.iter().enumerate() {
                let output_row = index as u32 + 1;
                let record = &row.record;
                let fill_color = row
                    .fill_color
                    .filter(|_| !row.fill_scope.is_empty())
                    .and_then(vba_color_index_rgb);
                // 指标名称显示列：设置开启时去首尾空格（不影响匹配与说明文字）。
                let display_indicator_name = if row.indicator_name_out.is_empty() {
                    record.indicator_name.as_str()
                } else {
                    row.indicator_name_out.as_str()
                };
                let display_indicator_name = if trim_indicator_names {
                    display_indicator_name.trim()
                } else {
                    display_indicator_name
                };
                // 源生产 exporter 口径：数值显示单元格统一按 15 位有效数字
                // 落盘，环比按写入的增减额/上期值重算后再取 15 位；规则匹配
                // 仍使用行内原始值。
                let display_value = display_central_15g(&record.value);
                let display_previous = display_central_15g(&row.previous_value);
                let display_change = display_central_15g(&row.change);
                let display_ratio =
                    display_ratio_cell(&row.ratio, &display_change, &display_previous);
                let text_values = [
                    (0, record.biz_class.as_str()),
                    (2, record.org_code.as_str()),
                    (3, record.org_name.as_str()),
                    (4, record.region_code.as_str()),
                    (5, record.region_name.as_str()),
                    (
                        6,
                        if row.order_code_out.is_empty() {
                            record.indicator.as_str()
                        } else {
                            row.order_code_out.as_str()
                        },
                    ),
                    (7, record.indicator.as_str()),
                    (8, display_indicator_name),
                    (9, record.data_attr.as_str()),
                    (10, record.currency.as_str()),
                    (11, record.frequency.as_str()),
                    (17, row.remark.as_str()),
                    (18, row.need_explanation.as_str()),
                    (19, row.explanation_content.as_str()),
                    (20, row.calculation_process.as_str()),
                ];
                for (column, value) in text_values {
                    let applies_fill = fill_color.is_some()
                        && ((row.fill_scope == "row" && column <= 17)
                            || (row.fill_scope != "row" && column == 17));
                    if applies_fill {
                        let format = formats
                            .text_by_fill
                            .get(&fill_color.expect("fill applies to text"))
                            .expect("text fill format was precomputed");
                        worksheet.write_string_with_format(output_row, column, value, format)?;
                    } else {
                        worksheet.write_string(output_row, column, value)?;
                    }
                }
                // 数据日期：写原生 Excel 日期值（源生产写 date 对象）。
                {
                    let applies_fill = fill_color.is_some() && row.fill_scope == "row";
                    let format = if applies_fill {
                        formats
                            .date_by_fill
                            .get(&fill_color.expect("fill applies to date"))
                            .expect("date fill format was precomputed")
                    } else {
                        &formats.date
                    };
                    match excel_datetime(&record.record_date) {
                        Some(datetime) => {
                            worksheet
                                .write_datetime_with_format(output_row, 1, &datetime, format)?;
                        }
                        None => {
                            worksheet.write_string_with_format(
                                output_row,
                                1,
                                record.record_date.as_str(),
                                format,
                            )?;
                        }
                    }
                }
                // 批次：规范整数写数值，其余保持文本（源生产 _batch_cell）。
                {
                    let applies_fill = fill_color.is_some() && row.fill_scope == "row";
                    if let Some(batch) = canonical_batch_number(&record.batch) {
                        if applies_fill {
                            let format = formats
                                .central
                                .get(&(None, fill_color))
                                .context("S3-F01 批次数值格式缓存缺少所需格式")?;
                            worksheet
                                .write_number_with_format(output_row, 12, batch as f64, format)?;
                        } else {
                            worksheet.write_number(output_row, 12, batch as f64)?;
                        }
                    } else if applies_fill {
                        let format = formats
                            .text_by_fill
                            .get(&fill_color.expect("fill applies to text"))
                            .expect("text fill format was precomputed");
                        worksheet.write_string_with_format(
                            output_row,
                            12,
                            record.batch.as_str(),
                            format,
                        )?;
                    } else {
                        worksheet.write_string(output_row, 12, record.batch.as_str())?;
                    }
                }
                write_central_value(
                    worksheet,
                    output_row,
                    13,
                    &display_value,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    14,
                    &display_previous,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    15,
                    &display_change,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    16,
                    &display_ratio,
                    Some("0.00%"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                    &formats,
                )?;
                worksheet.write_blank(output_row, 21, &formats.blank)?;
            }"""

assert s.count(old_loop) == 1, f"loop anchor {s.count(old_loop)}"
s = s.replace(old_loop, new_loop)

anchor = "struct ExecutionCompareFormats {"
helpers = """/// 源生产 `_excel_number_cell`/`_ratio_cell`：浮点按 15 位有效数字落盘。
fn excel_15g(value: f64) -> f64 {
    format!("{value:.14e}").parse::<f64>().unwrap_or(value)
}

fn display_central_15g(value: &CentralValue) -> CentralValue {
    match value {
        CentralValue::Number(number) => CentralValue::Number(excel_15g(*number)),
        other => other.clone(),
    }
}

/// 环比显示值：可整除时按写入后的增减额/上期值重算（源生产
/// `_comparison_output_cells`），否则保留原值；结果统一取 15 位。
fn display_ratio_cell(
    ratio: &CentralValue,
    change: &CentralValue,
    previous: &CentralValue,
) -> CentralValue {
    if let (CentralValue::Number(change), CentralValue::Number(previous)) = (change, previous) {
        if previous != &0.0 {
            return CentralValue::Number(excel_15g(change / previous));
        }
    }
    display_central_15g(ratio)
}

/// ISO 日期文本 -> 原生 Excel 日期值；解析失败保持文本写入。
fn excel_datetime(text: &str) -> Option<ExcelDateTime> {
    ExcelDateTime::parse_from_str(text.trim()).ok()
}

/// 源生产 `_batch_cell`：规范整数（0 或无前导零十进制）写数值。
fn canonical_batch_number(batch: &str) -> Option<i64> {
    let text = batch.trim();
    if text.is_empty() {
        return None;
    }
    if text == "0" {
        return Some(0);
    }
    let first = text.chars().next()?;
    if !first.is_ascii_digit() || first == '0' {
        return None;
    }
    if !text.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    text.parse::<i64>().ok()
}

struct ExecutionCompareFormats {"""
assert s.count(anchor) == 1
s = s.replace(anchor, helpers)

io.open(p, 'w', encoding='utf-8', newline='').write(s)
print('patched write loop + helpers')
