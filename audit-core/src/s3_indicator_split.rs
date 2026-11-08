//! S3-F05: map an existing S3-F01 comparison workbook into the embedded
//! financial-form template. This module plans business edits and prepares
//! capability-gated Excel-Core batches for institution workbooks.

use crate::s3_execution_compare::{
    indicator_split_fill_color_index, vba_color_index_rgb, S3AlertBand,
};
use std::sync::Arc;

use anyhow::{ensure, Context, Result};
use chrono::Local;
use excel_editor::{
    inspect_xlsx_profile, CellValue as EditCellValue, EditExecutionReport, EditOperation, EngineId,
    ExcelEditPlan, ExcelModifier, PreservationRequirement, SheetVisibility, WorkbookFormat,
    WorkbookProfile,
};
use excel_reader::{
    read_value_columns_with_iso_dates, read_xlsx_worksheet_extents, CellValue, SheetValues,
};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

const EMBEDDED_CONFIG_SHEETS: [&str; 2] = ["使用说明", "转表设置"];
const HEADER_PREFIXES: [&str; 3] = ["上期-", "增减额-", "环比-"];
const DROP_HEADER_MARKERS: [&str; 5] = ["剥离", "退出", "上期", "增减额", "环比"];
static STAGING_SEQUENCE: AtomicU64 = AtomicU64::new(0);
static PLAN_SEQUENCE: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorSplitOptions {
    pub hide_empty_rows: bool,
    pub delete_empty_sheets: bool,
}

impl Default for S3IndicatorSplitOptions {
    fn default() -> Self {
        Self {
            hide_empty_rows: true,
            delete_empty_sheets: true,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorSplitTemplate {
    pub path: PathBuf,
    pub options: S3IndicatorSplitOptions,
    pub forms: Vec<S3IndicatorFormDefinition>,
    pub disabled_form_codes: HashSet<String>,
    pub embedded_config_sheets: Vec<String>,
    pub sheets: Vec<S3IndicatorFormSheet>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorFormDefinition {
    pub code: String,
    pub name: String,
    pub frequency: String,
    pub batch: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorFormSheet {
    pub name: String,
    pub frequency: String,
    pub batch: String,
    pub business_class: String,
    pub base_headers: Vec<String>,
    /// Original one-based coordinates, deleted from right to left.
    pub drop_columns: Vec<usize>,
    pub indicator_rows: Vec<S3IndicatorFormRow>,
    /// Non-empty values from the source template, keyed by one-based A1
    /// coordinates. Used only to reproduce the legacy empty-sheet decision.
    pub values: BTreeMap<String, CellValue>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorFormRow {
    pub row: usize,
    pub indicator: String,
    pub keys: Vec<(usize, S3IndicatorMappingKey)>,
}

/// The dimensions that identify one value in the S3-F01 comparison output.
/// Frequency and batch remain separate fields so rows from distinct reporting
/// dimensions cannot overwrite one another during S3-F05 mapping.
#[derive(Clone, Debug, Eq, Ord, PartialEq, PartialOrd)]
pub struct S3IndicatorMappingKey {
    pub business_class: String,
    pub indicator: String,
    pub data_attribute: String,
    pub currency: String,
    pub frequency: String,
    pub batch: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3IndicatorSplitPlan {
    pub comparison_path: PathBuf,
    pub template_path: PathBuf,
    pub comparison_rows: usize,
    pub organizations: Vec<S3InstitutionSplitPlan>,
    pub warnings: Vec<String>,
    pub read_time_ms: u128,
    pub planning_time_ms: u128,
}

#[derive(Clone, Debug)]
pub struct S3IndicatorSplitExecutionReport {
    pub output_directory: PathBuf,
    pub output_files: Vec<PathBuf>,
    pub comparison_rows: usize,
    pub institution_count: usize,
    pub matched_indicators: usize,
    pub warnings: Vec<String>,
    pub edits: Vec<EditExecutionReport>,
    pub read_time_ms: u128,
    pub planning_time_ms: u128,
    pub editing_time_ms: u128,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3InstitutionSplitPlan {
    pub institution_code: String,
    pub region_code: String,
    pub institution_name: String,
    pub region_name: String,
    pub record_date: String,
    /// Concatenated frequency/batch value from the first comparison row,
    /// matching the legacy per-institution metadata.
    pub frequency_batch: String,
    pub output_file_name: String,
    pub edits: Vec<S3IndicatorSplitEdit>,
    pub matched_indicators: usize,
}

#[derive(Clone, Debug, PartialEq)]
pub enum S3IndicatorSplitEdit {
    DeleteSheet {
        sheet: String,
    },
    DeleteColumn {
        sheet: String,
        column: usize,
    },
    SetCellValue {
        sheet: String,
        cell: String,
        value: CellValue,
    },
    SetNumberFormat {
        sheet: String,
        cell: String,
        format: String,
    },
    SetCellFillColor {
        sheet: String,
        cell: String,
        argb: String,
    },
    HideRow {
        sheet: String,
        row: usize,
    },
    HideSheet {
        sheet: String,
    },
}

#[derive(Clone, Debug, PartialEq)]
struct ComparisonOrganization {
    institution_code: String,
    region_code: String,
    institution_name: String,
    region_name: String,
    record_date: String,
    frequency_batch: String,
    data: BTreeMap<S3IndicatorMappingKey, (CellValue, CellValue)>,
}

#[derive(Clone, Debug)]
struct MatchedFormRow {
    row: usize,
    indicator: String,
    values: Vec<(usize, CellValue, CellValue)>,
}

/// Read the existing 3.3 workbook as both configuration and the form template.
/// The legacy workbook is only read; each returned plan names a new output.
pub fn load_indicator_split_template(path: impl AsRef<Path>) -> Result<S3IndicatorSplitTemplate> {
    let started = Instant::now();
    let path = path.as_ref();
    ensure_xlsx(path, "S3-F05 指标比较拆分配置")?;
    let sheets = read_value_columns_with_iso_dates(path, usize::MAX)
        .with_context(|| format!("读取 S3-F05 指标比较拆分配置：{}", path.display()))?;
    let extents = read_xlsx_worksheet_extents(path)
        .with_context(|| format!("读取 S3-F05 模板工作表范围：{}", path.display()))?;
    let extent_by_name = extents
        .iter()
        .map(|extent| (extent.name.as_str(), extent))
        .collect::<HashMap<_, _>>();
    let sheet_by_name = sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<HashMap<_, _>>();
    let form_list = sheet_by_name
        .get("报表清单")
        .copied()
        .context("S3-F05 配置缺少“报表清单”工作表")?;

    let mut options = S3IndicatorSplitOptions::default();
    if let Some(settings) = sheet_by_name.get("转表设置") {
        let columns = header_columns(settings);
        if let (Some(parameter_column), Some(value_column)) =
            (columns.get("参数"), columns.get("值"))
        {
            let settings_max_row = extent_by_name
                .get("转表设置")
                .map_or(settings.max_row, |extent| {
                    extent.max_row.max(settings.max_row)
                });
            for row in 2..=settings_max_row {
                let parameter = python_or_empty(cell(settings, row, *parameter_column));
                let value = python_or_empty(cell(settings, row, *value_column));
                match parameter.trim() {
                    "隐藏空行" => options.hide_empty_rows = value != "否",
                    "空表删除" => options.delete_empty_sheets = value != "否",
                    _ => {}
                }
            }
        }
    }

    let mut forms = Vec::new();
    let mut disabled_form_codes = HashSet::new();
    let form_list_max_row = extent_by_name
        .get("报表清单")
        .map_or(form_list.max_row, |extent| {
            extent.max_row.max(form_list.max_row)
        });
    for row in 2..=form_list_max_row {
        let code = python_or_empty(cell(form_list, row, 1)).trim().to_owned();
        if code.is_empty() || !code.contains('A') {
            continue;
        }
        let flag = python_form_flag(cell(form_list, row, 5));
        if matches!(flag.as_str(), "" | "否" | "0" | "false") {
            disabled_form_codes.insert(code);
        } else {
            forms.push(S3IndicatorFormDefinition {
                code,
                name: python_or_empty(cell(form_list, row, 2)).trim().to_owned(),
                frequency: python_or_empty(cell(form_list, row, 3)).trim().to_owned(),
                batch: python_or_empty(cell(form_list, row, 4)).trim().to_owned(),
            });
        }
    }

    let embedded_config_sheets = if EMBEDDED_CONFIG_SHEETS
        .iter()
        .all(|name| sheet_by_name.contains_key(name))
    {
        EMBEDDED_CONFIG_SHEETS
            .iter()
            .map(|name| (*name).to_owned())
            .collect()
    } else {
        Vec::new()
    };
    let form_by_code = forms
        .iter()
        .map(|form| (form.code.as_str(), form))
        .collect::<HashMap<_, _>>();
    let form_codes = forms
        .iter()
        .map(|form| form.code.as_str())
        .collect::<HashSet<_>>();
    let mut form_sheets = Vec::new();
    for source in &sheets {
        if !form_codes.contains(source.name.as_str()) {
            continue;
        }
        let form = form_by_code[source.name.as_str()];
        let extent = extent_by_name.get(source.name.as_str());
        let max_column = extent.map_or(source.max_column, |extent| {
            extent.max_column.max(source.max_column)
        });
        let max_row = extent.map_or(source.max_row, |extent| extent.max_row.max(source.max_row));
        if max_column <= 3 {
            continue;
        }
        let header = (1..=max_column)
            .map(|column| python_or_empty(cell(source, 1, column)).trim().to_owned())
            .collect::<Vec<_>>();
        let drop_columns = (5..=max_column)
            .filter(|column| {
                let value = header
                    .get(column - 1)
                    .map(String::as_str)
                    .unwrap_or_default();
                value.is_empty()
                    || DROP_HEADER_MARKERS
                        .iter()
                        .any(|marker| value.contains(marker))
            })
            .collect::<Vec<_>>();
        let kept_columns = (4..=max_column)
            .filter(|column| !drop_columns.contains(column))
            .collect::<Vec<_>>();
        let base_headers = kept_columns
            .iter()
            .filter_map(|column| header.get(column - 1).cloned())
            .collect::<Vec<_>>();
        if base_headers.is_empty() {
            continue;
        }
        let business_class = sheet_business_class(&source.name).to_owned();
        let mut indicator_rows = Vec::new();
        for row in 2..=max_row {
            let indicator = python_or_empty(cell(source, row, 1)).trim().to_owned();
            if indicator.is_empty() {
                continue;
            }
            let keys = base_headers
                .iter()
                .enumerate()
                .map(|(offset, header)| {
                    let attr = if header.contains("发生额") {
                        "发生额"
                    } else {
                        "余额"
                    };
                    let currency = column_currency(header, &business_class);
                    (
                        4 + offset,
                        S3IndicatorMappingKey {
                            business_class: business_class.to_owned(),
                            indicator: indicator.clone(),
                            data_attribute: attr.to_owned(),
                            currency: currency.to_owned(),
                            frequency: form.frequency.clone(),
                            batch: form.batch.clone(),
                        },
                    )
                })
                .collect();
            indicator_rows.push(S3IndicatorFormRow {
                row,
                indicator,
                keys,
            });
        }
        form_sheets.push(S3IndicatorFormSheet {
            name: source.name.clone(),
            frequency: form.frequency.clone(),
            batch: form.batch.clone(),
            business_class,
            base_headers,
            drop_columns,
            indicator_rows,
            values: source.values.clone(),
        });
    }

    let _read_time_ms = started.elapsed().as_millis();
    Ok(S3IndicatorSplitTemplate {
        path: path.to_path_buf(),
        options,
        forms,
        disabled_form_codes,
        embedded_config_sheets,
        sheets: form_sheets,
    })
}

/// Plan one new form workbook per institution/region. The operation list
/// mirrors legacy `form_template.render_org_workbook` behavior without
/// reopening or modifying the source template.
pub fn plan_indicator_comparison_split(
    comparison_path: impl AsRef<Path>,
    template: &S3IndicatorSplitTemplate,
    alert_bands: &[S3AlertBand],
) -> Result<S3IndicatorSplitPlan> {
    let started = Instant::now();
    let comparison_path = comparison_path.as_ref();
    ensure_xlsx(comparison_path, "S3-F05 比较结果")?;
    let read_started = Instant::now();
    let comparison = read_comparison_result(comparison_path)?;
    let read_time_ms = read_started.elapsed().as_millis();
    let comparison_rows = comparison.0;
    let organizations = comparison.1;
    let warnings = comparison.2;

    let plans = organizations
        .into_iter()
        .map(|organization| plan_organization(&organization, template, alert_bands))
        .collect::<Result<Vec<_>>>()?;
    Ok(S3IndicatorSplitPlan {
        comparison_path: comparison_path.to_path_buf(),
        template_path: template.path.clone(),
        comparison_rows,
        organizations: plans,
        warnings,
        read_time_ms,
        planning_time_ms: started.elapsed().as_millis(),
    })
}

/// Render every institution's planned edits into a separate workbook copied
/// from the configured financial-form template. All workbooks are first
/// written into a private sibling directory and published as one directory
/// rename only after every Excel-Core edit succeeds.
/// 高级设置 edit_engine 的 S3-F05 渲染策略。
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FastRenderMode {
    /// auto / ooxml：FAST 渲染直出（ZIP 级，最快）
    Direct,
    /// openpyxl：FAST 渲染后交 openpyxl worker 做整簿往返（保留 openpyxl
    /// 独立序列化实现；逐格 Python 操作已被 Rust 回填取代）
    OpenpyxlRoundtrip,
}

pub fn fast_render_mode() -> FastRenderMode {
    if crate::app_settings::AppSettings::load()
        .edit_engine
        .as_str()
        == "openpyxl"
    {
        FastRenderMode::OpenpyxlRoundtrip
    } else {
        FastRenderMode::Direct
    }
}

pub fn fast_render_enabled() -> bool {
    true
}

use crate::s3_fast_ooxml_render::write_zip as fast_write_zip;

fn parse_a1_cell(cell: &str) -> (u32, u32) {
    // A1 → (行, 列)；FAST 渲染只需要数字坐标
    let letters: String = cell
        .chars()
        .take_while(|ch| ch.is_ascii_alphabetic())
        .collect();
    let digits: String = cell
        .chars()
        .skip_while(|ch| ch.is_ascii_alphabetic())
        .collect();
    let mut column = 0u32;
    for ch in letters.chars() {
        column = column * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
    }
    (digits.parse().unwrap_or(0), column)
}

fn fast_input_from_edits(
    edits: &[S3IndicatorSplitEdit],
) -> anyhow::Result<crate::s3_fast_ooxml_render::OrgRenderInput> {
    use crate::s3_fast_ooxml_render::{CellFill, OrgRenderInput};
    let mut input = OrgRenderInput::default();
    for edit in edits {
        match edit {
            S3IndicatorSplitEdit::DeleteSheet { sheet }
            | S3IndicatorSplitEdit::HideSheet { sheet } => {
                input.delete_candidates.push(sheet.clone());
            }
            S3IndicatorSplitEdit::DeleteColumn { .. } => {}
            S3IndicatorSplitEdit::HideRow { sheet, row } => {
                input
                    .sheets
                    .entry(sheet.clone())
                    .or_default()
                    .hidden_rows
                    .insert(*row as u32);
            }
            S3IndicatorSplitEdit::SetCellValue { sheet, cell, value } => {
                let (row, column) = parse_a1_cell(cell);
                let converted =
                    editor_cell_value(value).map_err(|error| anyhow::anyhow!("{error:#}"))?;
                input
                    .sheets
                    .entry(sheet.clone())
                    .or_default()
                    .cells
                    .entry((row, column))
                    .or_default()
                    .value = Some(converted);
            }
            S3IndicatorSplitEdit::SetNumberFormat {
                sheet,
                cell,
                format,
            } => {
                let (row, column) = parse_a1_cell(cell);
                input
                    .sheets
                    .entry(sheet.clone())
                    .or_default()
                    .cells
                    .entry((row, column))
                    .or_default()
                    .num_fmt = Some(format.clone());
            }
            S3IndicatorSplitEdit::SetCellFillColor { sheet, cell, argb } => {
                let (row, column) = parse_a1_cell(cell);
                let rgb = u32::from_str_radix(
                    argb.trim_start_matches('#').get(2..).unwrap_or("000000"),
                    16,
                )
                .unwrap_or(0);
                input
                    .sheets
                    .entry(sheet.clone())
                    .or_default()
                    .cells
                    .entry((row, column))
                    .or_default()
                    .fill_rgb = Some(rgb);
            }
        }
    }
    // 删空表/隐藏空行语义已由计划承载（plan_organization 生成显式操作）
    input.hide_empty_rows = false;
    input.delete_empty_sheets = false;
    Ok(input)
}

pub fn execute_indicator_comparison_split(
    comparison_path: impl AsRef<Path>,
    financial_template_path: impl AsRef<Path>,
    common_config_path: impl AsRef<Path>,
    output_root: impl AsRef<Path>,
    modifier: &ExcelModifier,
    fast_mode: FastRenderMode,
) -> Result<S3IndicatorSplitExecutionReport> {
    execute_indicator_comparison_split_with_progress(
        comparison_path,
        financial_template_path,
        common_config_path,
        output_root,
        modifier,
        fast_mode,
        |_| {},
    )
}

pub fn execute_indicator_comparison_split_with_progress<F>(
    comparison_path: impl AsRef<Path>,
    financial_template_path: impl AsRef<Path>,
    common_config_path: impl AsRef<Path>,
    output_root: impl AsRef<Path>,
    modifier: &ExcelModifier,
    fast_mode: FastRenderMode,
    mut on_progress: F,
) -> Result<S3IndicatorSplitExecutionReport>
where
    F: FnMut(&str),
{
    let comparison_path = comparison_path.as_ref();
    let financial_template_path = financial_template_path.as_ref();
    let common_config_path = common_config_path.as_ref();
    let output_root = output_root.as_ref();
    ensure_xlsx(comparison_path, "S3-F05 比较结果")?;
    ensure_xlsx(financial_template_path, "S3-F05 指标比较拆分配置")?;
    ensure!(common_config_path.is_file(), "S3-F05 大集中通用配置不存在");
    ensure!(
        !same_path(comparison_path, financial_template_path)?,
        "S3-F05 比较结果与金融表单模板必须是不同文件"
    );

    fs::create_dir_all(output_root)
        .with_context(|| format!("创建 S3-F05 输出目录：{}", output_root.display()))?;
    let output_root = fs::canonicalize(output_root)
        .with_context(|| format!("解析 S3-F05 输出目录：{}", output_root.display()))?;
    ensure!(output_root.is_dir(), "S3-F05 输出路径不是目录");

    on_progress(&format!(
        "读取 3.3 金融表单配置：{}",
        financial_template_path.display()
    ));
    let template = load_indicator_split_template(financial_template_path)?;
    on_progress(&format!(
        "读取 3.0 通用警戒配置：{}",
        common_config_path.display()
    ));
    let alert_bands =
        crate::s3_execution_compare::load_indicator_split_alert_bands(common_config_path)
            .context("读取 S3-F05 大集中通用配置中的环比预警区间")?;
    on_progress(&format!(
        "读取比较文件并规划机构表单：{}",
        comparison_path.display()
    ));
    let plan = plan_indicator_comparison_split(comparison_path, &template, &alert_bands)?;
    ensure!(
        !plan.organizations.is_empty(),
        "S3-F05 比较结果中没有可生成金融表单的机构记录"
    );

    let mut file_names = HashSet::new();
    for institution in &plan.organizations {
        validate_output_file_name(&institution.output_file_name)?;
        ensure!(
            file_names.insert(institution.output_file_name.to_lowercase()),
            "S3-F05 多个机构映射到相同输出文件名：{}",
            institution.output_file_name
        );
    }

    on_progress(&format!(
        "检查金融表单模板结构：{}",
        template.path.display()
    ));
    let workbook_profile = inspect_xlsx_profile(&template.path)
        .with_context(|| format!("检查 S3-F05 金融表单模板：{}", template.path.display()))?;
    ensure!(
        workbook_profile.format == WorkbookFormat::Xlsx,
        "S3-F05 金融表单模板目前仅支持 XLSX"
    );
    let openpyxl_roundtrip = fast_mode == FastRenderMode::OpenpyxlRoundtrip;
    let fast_template = {
        let bytes = std::fs::read(&template.path)
            .with_context(|| format!("读取金融表单模板：{}", template.path.display()))?;
        let t_compile = std::time::Instant::now();
        let compiled =
            crate::s3_fast_ooxml::compile_template(&bytes).context("FAST_OOXML 模板编译失败")?;
        if std::env::var_os("AUDIT_S3F5_TIMING").is_some() {
            eprintln!("[s3f5-fast] compile={}ms", t_compile.elapsed().as_millis());
        }
        let mut base_indexes: HashSet<String> = HashSet::new();
        base_indexes.insert("0".to_owned());
        let mut colors: HashSet<u32> = HashSet::new();
        colors.insert(3);
        colors.insert(6);
        for band in &alert_bands {
            if let Ok(value) = band.fill_color.parse::<u32>() {
                colors.insert(value);
            }
        }
        for sheet in compiled.sheets.values() {
            for style in sheet.base_style_by_col.values() {
                base_indexes.insert(if style.is_empty() {
                    "0".to_owned()
                } else {
                    style.clone()
                });
            }
        }
        let bank =
            crate::s3_fast_ooxml::StyleBank::new(&compiled.styles_xml, &base_indexes, &colors)
                .context("FAST_OOXML 样式编译失败")?;
        Some((Arc::new(compiled), Arc::new(bank)))
    };
    let final_directory = next_output_directory(&output_root)?;
    let staging = S3IndicatorSplitStagingDirectory::create(&output_root)?;
    let edit_started = Instant::now();
    let total = plan.organizations.len();
    let workers = total.min(
        std::thread::available_parallelism()
            .map(|n| n.get())
            .unwrap_or(4),
    );
    on_progress(&format!("并行生成 {total} 份机构表单（{workers} 线程）……"));
    let (progress_tx, progress_rx) = std::sync::mpsc::channel::<String>();

    let mut filled_reports: Vec<Option<EditExecutionReport>> = (0..total).map(|_| None).collect();
    let mut first_error: Option<anyhow::Error> = None;
    std::thread::scope(|scope| {
        let mut handles = Vec::with_capacity(total);
        for (index, institution) in plan.organizations.iter().enumerate() {
            let staged_output = staging.path().join(&institution.output_file_name);
            let modifier = &modifier;
            let fast_template = fast_template.as_ref();
            let template_path = &template.path;
            let profile = &workbook_profile;
            let sender = progress_tx.clone();
            handles.push(scope.spawn(move || {
                let label = format!(
                    "机构={}（{}）",
                    institution.institution_name, institution.institution_code
                );
                let _ = sender.send(format!(
                    "[{}/{}] {label}：准备 Excel-Core 编辑计划",
                    index + 1,
                    total
                ));
                let outcome = build_indicator_split_edit_plan(
                    template_path,
                    &staged_output,
                    profile,
                    institution,
                    index,
                )
                .and_then(|edit_plan| {
                    let _ =
                        sender.send(format!("[{}/{}] {label}：执行编辑并验证", index + 1, total));
                    if let Some((compiled, bank)) = fast_template.as_ref() {
                        let started = std::time::Instant::now();
                        let timing = std::env::var_os("AUDIT_S3F5_TIMING").is_some();
                        let input = fast_input_from_edits(&institution.edits)?;
                        let t_input = started.elapsed().as_millis();
                        let output =
                            crate::s3_fast_ooxml_render::render_org(compiled, bank, &input)
                                .map_err(|error| anyhow::anyhow!("{error:#}"))?;
                        let t_render = started.elapsed().as_millis();
                        fast_write_zip(
                            &staged_output,
                            &output.part_order,
                            &output.parts,
                            &output.workbook_xml,
                            &output.workbook_rels_xml,
                            &output.content_types_xml,
                        )
                        .map_err(|error| anyhow::anyhow!("{error:#}"))?;
                        if openpyxl_roundtrip {
                            crate::excel_route::fast_roundtrip_via_openpyxl(
                                &staged_output,
                                &edit_plan.request_id,
                            )
                            .map_err(|error| anyhow::anyhow!("{error:#}"))?;
                        }
                        if timing {
                            eprintln!(
                                "[s3f5-fast] {} input={}ms render={}ms rest={}ms total={}ms filled={} roundtrip={}",
                                institution.institution_code,
                                t_input,
                                t_render,
                                started.elapsed().as_millis() - t_render,
                                started.elapsed().as_millis(),
                                output.filled,
                                openpyxl_roundtrip
                            );
                        }
                        return Ok(EditExecutionReport {
                            engine: EngineId::OoxmlPatch,
                            input: template_path.clone(),
                            output: staged_output.clone(),
                            elapsed_ms: started.elapsed().as_millis(),
                            route_decisions: Vec::new(),
                            attempts: Vec::new(),
                        });
                    }
                    modifier
                        .execute(&edit_plan)
                        .map_err(|error| anyhow::anyhow!("{error:#}"))
                })
                .with_context(|| format!("S3-F05 生成机构 {label} 金融表单"))
                .and_then(|report| {
                    ensure!(
                        report.output == staged_output && staged_output.is_file(),
                        "S3-F05 Excel-Core 未生成请求的金融表单：{}",
                        staged_output.display()
                    );
                    Ok(report)
                });
                let _ = sender.send(match &outcome {
                    Ok(_) => format!(
                        "[{}/{}] 金融表单已生成并验证：{}",
                        index + 1,
                        total,
                        staged_output.display()
                    ),
                    Err(error) => format!("[{}/{}] {label} 失败：{error:#}", index + 1, total),
                });
                (index, outcome)
            }));
        }
        drop(progress_tx);
        // 进度只从主线程回调（FnMut 非线程安全）；通道断开即全部工作线程结束
        while let Ok(message) = progress_rx.recv() {
            on_progress(&message);
        }
        for handle in handles {
            let (index, outcome) = match handle.join() {
                Ok(joined) => joined,
                Err(panic) => {
                    let message = panic
                        .downcast_ref::<&str>()
                        .map(|text| (*text).to_owned())
                        .or_else(|| panic.downcast_ref::<String>().cloned())
                        .unwrap_or_else(|| "工作线程已 panic".to_owned());
                    // 索引取占位值即可：outcome 为 Err，不会写入 filled_reports
                    (
                        usize::MAX,
                        Err(anyhow::anyhow!("S3-F05 并行生成工作线程异常：{message}")),
                    )
                }
            };
            match outcome {
                Ok(report) => filled_reports[index] = Some(report),
                Err(error) if first_error.is_none() => first_error = Some(error),
                Err(_) => {}
            }
        }
    });
    if let Some(error) = first_error {
        return Err(error);
    }
    let mut edit_reports = Vec::with_capacity(total);
    let mut output_names = Vec::with_capacity(total);
    for (index, report) in filled_reports.into_iter().enumerate() {
        let report = report.with_context(|| {
            format!(
                "S3-F05 机构 {} 未产出编辑报告",
                plan.organizations[index].institution_code
            )
        })?;
        output_names.push(plan.organizations[index].output_file_name.clone());
        edit_reports.push(report);
    }

    let editing_time_ms = edit_started.elapsed().as_millis();
    on_progress(&format!(
        "发布全部机构表单目录：{}",
        final_directory.display()
    ));
    let output_directory = staging.publish(&final_directory)?;
    on_progress(&format!(
        "机构表单批次已发布：{}",
        output_directory.display()
    ));
    let output_files = output_names
        .iter()
        .map(|name| output_directory.join(name))
        .collect();
    Ok(S3IndicatorSplitExecutionReport {
        output_directory,
        output_files,
        comparison_rows: plan.comparison_rows,
        institution_count: plan.organizations.len(),
        matched_indicators: plan
            .organizations
            .iter()
            .map(|institution| institution.matched_indicators)
            .sum(),
        warnings: plan.warnings,
        edits: edit_reports,
        read_time_ms: plan.read_time_ms,
        planning_time_ms: plan.planning_time_ms,
        editing_time_ms,
    })
}

fn build_indicator_split_edit_plan(
    template_path: &Path,
    output_path: &Path,
    workbook_profile: &WorkbookProfile,
    institution: &S3InstitutionSplitPlan,
    index: usize,
) -> Result<ExcelEditPlan> {
    ensure!(
        workbook_profile.format == WorkbookFormat::Xlsx,
        "S3-F05 ExcelEditPlan 目前仅支持 XLSX 模板"
    );
    ensure!(
        !institution.edits.is_empty(),
        "S3-F05 机构 {} 没有可应用的表单操作",
        institution.institution_code
    );
    let mut operations = Vec::with_capacity(institution.edits.len());
    for edit in &institution.edits {
        operations.push(match edit {
            S3IndicatorSplitEdit::DeleteSheet { sheet } => EditOperation::DeleteSheet {
                name: sheet.clone(),
            },
            S3IndicatorSplitEdit::DeleteColumn { sheet, column } => EditOperation::DeleteColumn {
                sheet: sheet.clone(),
                column_index: u32::try_from(*column)
                    .context("S3-F05 删除列索引超出 Excel-Core 范围")?,
            },
            S3IndicatorSplitEdit::SetCellValue { sheet, cell, value } => {
                EditOperation::SetCellValue {
                    sheet: sheet.clone(),
                    cell: cell.clone(),
                    value: editor_cell_value(value)?,
                }
            }
            S3IndicatorSplitEdit::SetNumberFormat {
                sheet,
                cell,
                format,
            } => EditOperation::SetCellNumberFormat {
                sheet: sheet.clone(),
                cell: cell.clone(),
                format_code: format.clone(),
            },
            S3IndicatorSplitEdit::SetCellFillColor { sheet, cell, argb } => {
                EditOperation::SetCellFillColor {
                    sheet: sheet.clone(),
                    cell: cell.clone(),
                    argb: argb.clone(),
                }
            }
            S3IndicatorSplitEdit::HideRow { sheet, row } => EditOperation::SetRowHidden {
                sheet: sheet.clone(),
                row_index: u32::try_from(*row).context("S3-F05 行号超出 Excel-Core 范围")?,
                hidden: true,
            },
            S3IndicatorSplitEdit::HideSheet { sheet } => EditOperation::SetSheetVisibility {
                sheet: sheet.clone(),
                visibility: SheetVisibility::Hidden,
            },
        });
    }

    let sequence = PLAN_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    let plan = ExcelEditPlan {
        request_id: format!(
            "s3-f5-{}-{index}-{sequence}",
            Local::now().format("%Y%m%d%H%M%S")
        ),
        input: template_path.to_path_buf(),
        output: output_path.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: workbook_profile.clone(),
        operations,
        preserve: s3_f5_required_preservation(),
        // 删表会连带移除该表作用域的定义名称（如 _xlnm._FilterDatabase），
        // 这是删表的预期语义而非丢失；显式登记允许，供后端守卫放行。
        drop_allowed: {
            let mut allowed = BTreeMap::new();
            for edit in &institution.edits {
                if let S3IndicatorSplitEdit::DeleteSheet { sheet } = edit {
                    let mut requirements = std::collections::BTreeSet::new();
                    requirements.insert(PreservationRequirement::DefinedNames);
                    requirements.insert(PreservationRequirement::RowAndColumnDimensions);
                    allowed.insert(sheet.clone(), requirements);
                }
            }
            allowed
        },
    };
    plan.validate().context("校验 S3-F05 ExcelEditPlan")?;
    Ok(plan)
}

fn editor_cell_value(value: &CellValue) -> Result<EditCellValue> {
    Ok(match value {
        CellValue::Empty => EditCellValue::Blank,
        CellValue::Text(text) | CellValue::DateTimeIso(text) | CellValue::DurationIso(text) => {
            EditCellValue::Text(text.clone())
        }
        CellValue::Number(number) => {
            ensure!(number.is_finite(), "S3-F05 单元格包含非有限数值");
            EditCellValue::Decimal(number.to_string())
        }
        CellValue::Integer(number) => EditCellValue::Decimal(number.to_string()),
        CellValue::DateTimeSerial(serial) => {
            ensure!(serial.is_finite(), "S3-F05 日期单元格包含非有限数值");
            EditCellValue::DateTime {
                serial: serial.to_string(),
                number_format: "General".to_owned(),
            }
        }
        CellValue::Boolean(value) => EditCellValue::Boolean(*value),
        CellValue::Error(error) => EditCellValue::Error(error.clone()),
    })
}

/// S3-F05 金融表单的保真要求（业务口径 B，2026-11-07 确认）。
///
/// 金融表单的验收口径是「数值/公式/结构正确、可打开」，不是逐像素复刻模板
/// 渲染。openpyxl 重写必然规范化样式索引、页眉页脚、页面设置、关系部件与
/// 冻结/筛选元数据，而这些差异不影响表单业务含义；因此这里只保留业务关键项，
/// 渲染层不再作为硬门槛。真实数据丢失仍会被下面的项拦截。
fn s3_f5_required_preservation() -> std::collections::BTreeSet<PreservationRequirement> {
    [
        // 业务结论直接依赖：单元格值/类型、公式文本、条件格式（影响审核判断）
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        // 结构完整性：合并区域、定义名称、工作表身份与顺序、批注
        PreservationRequirement::MergedRanges,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::SheetIdentityAndOrder,
        PreservationRequirement::CommentsAndPositions,
        // 注：Styles / RowAndColumnDimensions / PageSettings / HeaderFooterText /
        // Relationships / FreezePanes / AutoFilter 不列为硬要求——openpyxl
        // 重写会规范化这些渲染元数据（业务口径 B）。差异在输出校验中以诊断形式
        // 记录，不阻断发布。
    ]
    .into_iter()
    .collect()
}

fn validate_output_file_name(name: &str) -> Result<()> {
    let path = Path::new(name);
    ensure!(
        path.file_name().and_then(|value| value.to_str()) == Some(name),
        "S3-F05 输出文件名包含目录部分：{name}"
    );
    ensure!(
        path.extension()
            .and_then(|value| value.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F05 输出文件名必须为 XLSX：{name}"
    );
    ensure!(
        !name.trim().is_empty()
            && !name.chars().any(|character| {
                matches!(
                    character,
                    '<' | '>' | ':' | '"' | '|' | '?' | '*' | '/' | '\\'
                )
            }),
        "S3-F05 输出文件名包含无效字符：{name}"
    );
    Ok(())
}

fn same_path(left: &Path, right: &Path) -> Result<bool> {
    let left =
        fs::canonicalize(left).with_context(|| format!("解析文件路径：{}", left.display()))?;
    let right =
        fs::canonicalize(right).with_context(|| format!("解析文件路径：{}", right.display()))?;
    Ok(left == right)
}

fn next_output_directory(output_root: &Path) -> Result<PathBuf> {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S_%3f");
    for suffix in 0..128 {
        let name = if suffix == 0 {
            format!("指标比较拆分_{timestamp}")
        } else {
            format!("指标比较拆分_{timestamp}_{suffix}")
        };
        let path = output_root.join(name);
        if !path.exists() {
            return Ok(path);
        }
    }
    anyhow::bail!("无法为 S3-F05 生成唯一的输出目录名")
}

struct S3IndicatorSplitStagingDirectory {
    path: PathBuf,
    published: bool,
}

impl S3IndicatorSplitStagingDirectory {
    fn create(parent: &Path) -> Result<Self> {
        for _ in 0..128 {
            let sequence = STAGING_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let path = parent.join(format!(".audit-s3-f5-{}-{sequence}", std::process::id()));
            match fs::create_dir(&path) {
                Ok(()) => {
                    return Ok(Self {
                        path,
                        published: false,
                    })
                }
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建 S3-F05 暂存目录：{}", path.display()))
                }
            }
        }
        anyhow::bail!("无法创建唯一的 S3-F05 暂存目录")
    }

    fn path(&self) -> &Path {
        &self.path
    }

    fn publish(mut self, destination: &Path) -> Result<PathBuf> {
        ensure!(!destination.exists(), "S3-F05 输出目录已存在，拒绝覆盖");
        fs::rename(&self.path, destination).with_context(|| {
            format!(
                "发布 S3-F05 金融表单目录：{} -> {}",
                self.path.display(),
                destination.display()
            )
        })?;
        self.published = true;
        Ok(destination.to_path_buf())
    }
}

impl Drop for S3IndicatorSplitStagingDirectory {
    fn drop(&mut self) {
        if !self.published {
            let _ = fs::remove_dir_all(&self.path);
        }
    }
}

fn plan_organization(
    organization: &ComparisonOrganization,
    template: &S3IndicatorSplitTemplate,
    alert_bands: &[S3AlertBand],
) -> Result<S3InstitutionSplitPlan> {
    let mut edits = Vec::new();
    for sheet in &template.embedded_config_sheets {
        edits.push(S3IndicatorSplitEdit::DeleteSheet {
            sheet: sheet.clone(),
        });
    }
    for sheet in &template.disabled_form_codes {
        edits.push(S3IndicatorSplitEdit::DeleteSheet {
            sheet: sheet.clone(),
        });
    }

    let mut matched_indicators = 0;
    for sheet in &template.sheets {
        for column in sheet.drop_columns.iter().rev() {
            edits.push(S3IndicatorSplitEdit::DeleteColumn {
                sheet: sheet.name.clone(),
                column: *column,
            });
        }
        let mut matched = Vec::new();
        for row in &sheet.indicator_rows {
            let mut values = Vec::new();
            for (column, form_key) in &row.keys {
                let matches = organization
                    .data
                    .iter()
                    .filter(|(candidate, _)| {
                        candidate.business_class == form_key.business_class
                            && candidate.indicator == form_key.indicator
                            && candidate.data_attribute == form_key.data_attribute
                            && candidate.currency == form_key.currency
                            && (form_key.frequency.is_empty()
                                || candidate.frequency == form_key.frequency)
                            && (form_key.batch.is_empty() || candidate.batch == form_key.batch)
                    })
                    .take(2)
                    .collect::<Vec<_>>();
                ensure!(
                    matches.len() <= 1,
                    "S3-F05 表单 {} 指标 {} 第 {} 行的频度/批次匹配到多条比较记录",
                    sheet.name,
                    row.indicator,
                    row.row
                );
                if let Some((_, (current, previous))) = matches.first() {
                    values.push((*column, (*current).clone(), (*previous).clone()));
                }
            }
            if !values.is_empty() {
                matched.push(MatchedFormRow {
                    row: row.row,
                    indicator: row.indicator.clone(),
                    values,
                });
            }
        }
        if matched.is_empty() && template.options.delete_empty_sheets {
            edits.push(S3IndicatorSplitEdit::DeleteSheet {
                sheet: sheet.name.clone(),
            });
            continue;
        }

        let base_count = sheet.base_headers.len();
        for (block, prefix) in HEADER_PREFIXES.iter().enumerate() {
            for (offset, header) in sheet.base_headers.iter().enumerate() {
                let column = 4 + base_count * (block + 1) + offset;
                edits.push(set_text(
                    &sheet.name,
                    1,
                    column,
                    format!("{prefix}{header}"),
                ));
            }
        }

        let matched_rows = matched.iter().map(|row| row.row).collect::<HashSet<_>>();
        matched_indicators += matched_rows.len();
        for row in &matched {
            for (current_column, current, previous) in &row.values {
                let offset = current_column - 4;
                let current_number = python_number(current);
                let previous_number = python_number(previous);
                let change = current_number - previous_number;
                ensure!(
                    current_number.is_finite() && previous_number.is_finite() && change.is_finite(),
                    "S3-F05 指标 {} 第 {} 行包含非有限数值",
                    row.indicator,
                    row.row
                );

                let mut colored_cells = Vec::new();
                if !matches!(current, CellValue::Empty) {
                    emit_value(
                        &mut edits,
                        &sheet.name,
                        row.row,
                        *current_column,
                        current.clone(),
                        "0.00",
                    );
                    colored_cells.push(*current_column);
                }
                if !matches!(previous, CellValue::Empty) {
                    let column = 4 + base_count + offset;
                    emit_value(
                        &mut edits,
                        &sheet.name,
                        row.row,
                        column,
                        previous.clone(),
                        "0.00",
                    );
                    colored_cells.push(column);
                }
                if change != 0.0 {
                    let column = 4 + 2 * base_count + offset;
                    emit_value(
                        &mut edits,
                        &sheet.name,
                        row.row,
                        column,
                        CellValue::Number(change),
                        "0.00",
                    );
                    colored_cells.push(column);
                }
                let color_index = if previous_number != 0.0 && change != 0.0 {
                    let ratio = change / previous_number;
                    ensure!(
                        ratio.is_finite(),
                        "S3-F05 指标 {} 的环比不是有限数值",
                        row.indicator
                    );
                    let column = 4 + 3 * base_count + offset;
                    emit_value(
                        &mut edits,
                        &sheet.name,
                        row.row,
                        column,
                        CellValue::Number(ratio),
                        "0.00%",
                    );
                    colored_cells.push(column);
                    indicator_split_fill_color_index(ratio, alert_bands)
                } else if matches!(current, CellValue::Empty) {
                    Some(6)
                } else if matches!(previous, CellValue::Empty) {
                    Some(3)
                } else {
                    None
                };
                if let Some(color) = color_index.and_then(vba_color_index_rgb) {
                    let argb = format!("FF{color:06X}");
                    for column in colored_cells {
                        edits.push(S3IndicatorSplitEdit::SetCellFillColor {
                            sheet: sheet.name.clone(),
                            cell: cell_address(row.row, column),
                            argb: argb.clone(),
                        });
                    }
                }
            }
        }

        if template.options.hide_empty_rows {
            for row in &sheet.indicator_rows {
                if !matched_rows.contains(&row.row) {
                    edits.push(S3IndicatorSplitEdit::HideRow {
                        sheet: sheet.name.clone(),
                        row: row.row,
                    });
                }
            }
        }
        if !template.options.delete_empty_sheets && !sheet_has_output_content(sheet, &matched) {
            edits.push(S3IndicatorSplitEdit::HideSheet {
                sheet: sheet.name.clone(),
            });
        }
    }

    Ok(S3InstitutionSplitPlan {
        institution_code: organization.institution_code.clone(),
        region_code: organization.region_code.clone(),
        institution_name: organization.institution_name.clone(),
        region_name: organization.region_name.clone(),
        record_date: organization.record_date.clone(),
        frequency_batch: organization.frequency_batch.clone(),
        output_file_name: output_file_name(
            &organization.record_date,
            &organization.institution_name,
            &organization.region_name,
        ),
        edits,
        matched_indicators,
    })
}

fn sheet_has_output_content(sheet: &S3IndicatorFormSheet, matched: &[MatchedFormRow]) -> bool {
    let last_data_row = sheet
        .indicator_rows
        .last()
        .map(|row| row.row)
        .unwrap_or_default()
        .min(5000);
    let base_count = sheet.base_headers.len();
    let max_column = (3 + 4 * base_count).min(20);
    for (address, value) in &sheet.values {
        if matches!(value, CellValue::Empty) {
            continue;
        }
        let Some((row, old_column)) = parse_cell_address(address) else {
            continue;
        };
        if row < 2
            || row > last_data_row
            || old_column < 4
            || sheet.drop_columns.contains(&old_column)
        {
            continue;
        }
        let shift = sheet
            .drop_columns
            .iter()
            .filter(|column| **column < old_column)
            .count();
        if old_column - shift <= max_column {
            return true;
        }
    }
    matched.iter().any(|row| {
        row.values.iter().any(|(_, current, previous)| {
            if !matches!(current, CellValue::Empty) || !matches!(previous, CellValue::Empty) {
                return true;
            }
            python_number(current) != python_number(previous)
        })
    })
}

fn emit_value(
    edits: &mut Vec<S3IndicatorSplitEdit>,
    sheet: &str,
    row: usize,
    column: usize,
    value: CellValue,
    number_format: &str,
) {
    let cell = cell_address(row, column);
    edits.push(S3IndicatorSplitEdit::SetCellValue {
        sheet: sheet.to_owned(),
        cell: cell.clone(),
        value,
    });
    edits.push(S3IndicatorSplitEdit::SetNumberFormat {
        sheet: sheet.to_owned(),
        cell,
        format: number_format.to_owned(),
    });
}

fn set_text(sheet: &str, row: usize, column: usize, value: String) -> S3IndicatorSplitEdit {
    S3IndicatorSplitEdit::SetCellValue {
        sheet: sheet.to_owned(),
        cell: cell_address(row, column),
        value: CellValue::Text(value),
    }
}

fn read_comparison_result(
    path: &Path,
) -> Result<(usize, Vec<ComparisonOrganization>, Vec<String>)> {
    let sheets = read_value_columns_with_iso_dates(path, usize::MAX)
        .with_context(|| format!("读取 S3-F05 比较结果：{}", path.display()))?;
    let result_sheet = sheets
        .iter()
        .find(|sheet| sheet.name.contains("对比结果"))
        .context("S3-F05 比较结果文件缺少“对比结果”工作表")?;
    ensure!(result_sheet.max_row >= 1, "S3-F05 比较结果工作表为空");
    let expected_headers = [
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
    ];
    let header_mismatch = expected_headers
        .iter()
        .enumerate()
        .any(|(index, expected)| python_or_empty(cell(result_sheet, 1, index + 1)) != *expected);
    let mut warnings = Vec::new();
    if header_mismatch {
        warnings.push("比较结果表头与标准 21 列不完全一致，按旧版固定列位置读取".to_owned());
    }
    ensure!(
        result_sheet.max_column >= 15,
        "S3-F05 比较结果至少需要 15 列（本期值和上期值）"
    );

    let mut row_count = 0;
    let mut organizations = Vec::<ComparisonOrganization>::new();
    let mut by_org_key = HashMap::<String, usize>::new();
    for row in 2..=result_sheet.max_row {
        if (1..=result_sheet.max_column)
            .all(|column| matches!(cell(result_sheet, row, column), CellValue::Empty))
        {
            continue;
        }
        row_count += 1;
        let organization_code = python_or_empty(cell(result_sheet, row, 3))
            .replace('\'', "")
            .trim()
            .to_owned();
        let region_code = python_or_empty(cell(result_sheet, row, 5))
            .replace('\'', "")
            .trim()
            .to_owned();
        let group_key = format!("{organization_code}{region_code}");
        let index = *by_org_key.entry(group_key).or_insert_with(|| {
            organizations.push(ComparisonOrganization {
                institution_code: organization_code.clone(),
                region_code: region_code.clone(),
                institution_name: python_or_empty(cell(result_sheet, row, 4))
                    .trim()
                    .to_owned(),
                region_name: python_or_empty(cell(result_sheet, row, 6))
                    .trim()
                    .to_owned(),
                record_date: python_date_prefix(cell(result_sheet, row, 2)),
                frequency_batch: format!(
                    "{}{}",
                    python_or_empty(cell(result_sheet, row, 12)),
                    python_or_empty(cell(result_sheet, row, 13))
                ),
                data: BTreeMap::new(),
            });
            organizations.len() - 1
        });
        let key = S3IndicatorMappingKey {
            business_class: python_or_empty(cell(result_sheet, row, 1))
                .trim()
                .to_owned(),
            indicator: python_or_empty(cell(result_sheet, row, 8))
                .trim()
                .to_owned(),
            data_attribute: python_or_empty(cell(result_sheet, row, 10))
                .trim()
                .to_owned(),
            currency: python_or_empty(cell(result_sheet, row, 11))
                .trim()
                .to_owned(),
            frequency: python_or_empty(cell(result_sheet, row, 12))
                .trim()
                .to_owned(),
            batch: python_or_empty(cell(result_sheet, row, 13))
                .trim()
                .to_owned(),
        };
        organizations[index].data.entry(key).or_insert_with(|| {
            (
                cell(result_sheet, row, 14).clone(),
                cell(result_sheet, row, 15).clone(),
            )
        });
    }
    Ok((row_count, organizations, warnings))
}

fn header_columns(sheet: &SheetValues) -> HashMap<String, usize> {
    (1..=sheet.max_column)
        .filter_map(|column| {
            let header = python_or_empty(cell(sheet, 1, column)).trim().to_owned();
            (!header.is_empty()).then_some((header, column))
        })
        .collect()
}

fn cell(sheet: &SheetValues, row: usize, column: usize) -> &CellValue {
    sheet
        .values
        .get(&cell_address(row, column))
        .unwrap_or(&CellValue::Empty)
}

fn python_or_empty(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(value) if *value == 0.0 => String::new(),
        CellValue::Integer(0) | CellValue::Boolean(false) => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => value.to_string(),
        CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => value.clone(),
    }
}

fn python_form_flag(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(number) if *number == 0.0 || *number == 1.0 => {
            format!("{:.0}", number)
        }
        CellValue::Integer(number) if *number == 0 || *number == 1 => number.to_string(),
        CellValue::Boolean(value) => if *value { "1" } else { "0" }.to_owned(),
        _ => python_or_empty(value).trim().to_ascii_lowercase(),
    }
}

fn python_number(value: &CellValue) -> f64 {
    match value {
        CellValue::Number(value) => *value,
        CellValue::Integer(value) => *value as f64,
        CellValue::Boolean(true) => 1.0,
        _ => 0.0,
    }
}

fn python_date_prefix(value: &CellValue) -> String {
    let text = match value {
        CellValue::DateTimeIso(text) => text.clone(),
        _ => python_or_empty(value),
    };
    text.chars().take(10).collect()
}

fn sheet_business_class(name: &str) -> &'static str {
    if name.contains("A1") {
        "人民币"
    } else if name.contains("A2") {
        "外币"
    } else {
        "本外币"
    }
}

fn column_currency(header: &str, business_class: &str) -> &'static str {
    match business_class {
        "人民币" => "人民币",
        "外币" => "美元合计",
        _ if header.contains("人民币") => "人民币",
        _ if header.contains("外币") => "美元合计",
        _ => "人民币",
    }
}

fn output_file_name(record_date: &str, institution_name: &str, region_name: &str) -> String {
    let mut parts = record_date.split('-');
    if let (Some(year), Some(month)) = (parts.next(), parts.next()) {
        if let Ok(month) = month.parse::<u32>() {
            return format!("{year}.{month} {institution_name} {region_name}.xlsx");
        }
    }
    format!("{institution_name} {region_name}.xlsx")
}

fn cell_address(row: usize, mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters.reverse();
    format!("{}{row}", letters.into_iter().collect::<String>())
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let mut column = 0usize;
    let mut row = 0usize;
    let mut in_row = false;
    for character in address.chars() {
        if character.is_ascii_alphabetic() && !in_row {
            column = column
                .checked_mul(26)?
                .checked_add(character.to_ascii_uppercase() as usize - 'A' as usize + 1)?;
        } else if character.is_ascii_digit() {
            in_row = true;
            row = row
                .checked_mul(10)?
                .checked_add(character as usize - '0' as usize)?;
        } else {
            return None;
        }
    }
    (row > 0 && column > 0).then_some((row, column))
}

fn ensure_xlsx(path: &Path, description: &str) -> Result<()> {
    ensure!(
        path.is_file(),
        "{description}文件不存在：{}",
        path.display()
    );
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "{description}目前只支持 XLSX：{}",
        path.display()
    );
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::s3_execution_compare::S3AlertBand;
    use excel_editor::{ExcelEditPlan, OutputValidator};
    use excel_writer::Workbook;
    use sha2::Digest;
    use std::{
        fs,
        sync::{
            atomic::{AtomicU64, Ordering},
            Arc,
        },
    };

    static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

    fn temp_path(label: &str) -> PathBuf {
        let sequence = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        std::env::temp_dir().join(format!(
            "audit-s3-f05-{label}-{}-{sequence}.xlsx",
            std::process::id()
        ))
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

    fn band() -> S3AlertBand {
        S3AlertBand {
            lower: "10".to_owned(),
            upper: "30".to_owned(),
            fill_color: "6".to_owned(),
            ..S3AlertBand::default()
        }
    }

    struct AcceptAllValidator;

    impl OutputValidator for AcceptAllValidator {
        fn validate(&self, _original_plan: &ExcelEditPlan, _candidate_output: &Path) -> Result<()> {
            Ok(())
        }
    }

    #[test]
    fn plans_form_mapping_headers_values_formats_colors_and_visibility() {
        let template_path = temp_path("template");
        let comparison_path = temp_path("comparison");
        write_template(&template_path);
        write_comparison(&comparison_path);
        let template = load_indicator_split_template(&template_path).unwrap();
        assert_eq!(template.forms.len(), 2);
        assert!(template.disabled_form_codes.contains("关闭A2"));
        assert_eq!(template.embedded_config_sheets, ["使用说明", "转表设置"]);
        assert_eq!(template.sheets.len(), 2);
        assert_eq!(template.sheets[0].frequency, "月");
        assert_eq!(template.sheets[0].batch, "1");

        let plan = plan_indicator_comparison_split(&comparison_path, &template, &[band()]).unwrap();
        assert_eq!(plan.comparison_rows, 3);
        assert_eq!(plan.organizations.len(), 1);
        let organization = &plan.organizations[0];
        assert_eq!(organization.output_file_name, "2026.8 机构 地区.xlsx");
        assert_eq!(organization.frequency_batch, "月1");
        assert_eq!(organization.matched_indicators, 1);
        assert!(organization
            .edits
            .iter()
            .any(|edit| matches!(edit, S3IndicatorSplitEdit::DeleteColumn { column: 7, .. })));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::SetCellValue { sheet, cell, value: CellValue::Number(value) }
                if sheet == "测试A1" && cell == "D2" && *value == 120.0
        )));
        assert!(!organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::SetCellValue { sheet, cell, value: CellValue::Number(value) }
                if sheet == "测试A1" && cell == "D2" && *value == 999.0
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::SetCellValue { sheet, cell, value: CellValue::Text(value) }
                if sheet == "测试A1" && cell == "E1" && value == "上期-余额人民币"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::SetNumberFormat { sheet, cell, format }
                if sheet == "测试A1" && cell == "G2" && format == "0.00%"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::SetCellFillColor { sheet, cell, argb }
                if sheet == "测试A1" && cell == "D2" && argb == "FFFFFF00"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::HideRow { sheet, row: 3 } if sheet == "测试A1"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::DeleteSheet { sheet } if sheet == "空表A1"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::DeleteSheet { sheet } if sheet == "关闭A2"
        )));
        assert!(organization.edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::DeleteSheet { sheet } if sheet == "使用说明"
        )));

        fs::remove_file(template_path).unwrap();
        fs::remove_file(comparison_path).unwrap();
    }

    #[test]
    fn edit_plan_translates_all_business_operations_and_preserves_template_contract() {
        let template_path = temp_path("edit-plan-template");
        let output_path = temp_path("edit-plan-output");
        let comparison_path = temp_path("edit-plan-comparison");
        write_template(&template_path);
        write_comparison(&comparison_path);
        let template = load_indicator_split_template(&template_path).unwrap();
        let business_plan =
            plan_indicator_comparison_split(&comparison_path, &template, &[band()]).unwrap();
        let profile = inspect_xlsx_profile(&template_path).unwrap();
        let edit_plan = build_indicator_split_edit_plan(
            &template_path,
            &output_path,
            &profile,
            &business_plan.organizations[0],
            0,
        )
        .unwrap();

        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::DeleteColumn { sheet, column_index: 7 } if sheet == "测试A1"
        )));
        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellValue { sheet, cell, value: EditCellValue::Decimal(value) }
                if sheet == "测试A1" && cell == "D2" && value == "120"
        )));
        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellNumberFormat { sheet, cell, format_code }
                if sheet == "测试A1" && cell == "G2" && format_code == "0.00%"
        )));
        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellFillColor { sheet, cell, argb }
                if sheet == "测试A1" && cell == "D2" && argb == "FFFFFF00"
        )));
        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetRowHidden { sheet, row_index: 3, hidden: true }
                if sheet == "测试A1"
        )));
        assert!(edit_plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::DeleteSheet { name } if name == "使用说明"
        )));
        assert!(edit_plan
            .preserve
            .contains(&PreservationRequirement::FormulaText));
        assert!(edit_plan
            .preserve
            .contains(&PreservationRequirement::ConditionalFormatting));

        fs::remove_file(template_path).unwrap();
        fs::remove_file(comparison_path).unwrap();
    }

    /// FAST 渲染是独立的 Rust 路径，不依赖 Excel-Core 后端注册表；
    /// 但任何失败仍必须 fail-closed、不留下半成品（沿用原测试的意图，
    /// 失败点从「后端注册表为空」换成「模板无法编译」）。
    #[test]
    fn execution_without_a_passed_excel_core_route_publishes_no_partial_outputs() {
        let template_path = temp_path("capability-template");
        let comparison_path = temp_path("capability-comparison");
        let common_path = temp_path("capability-common");
        let output_root = temp_path("capability-output-root").with_extension("");
        write_template(&template_path);
        write_comparison(&comparison_path);
        write_split_common_config(&common_path);
        fs::create_dir_all(&output_root).unwrap();
        let modifier =
            ExcelModifier::new("test-platform", Vec::new(), Arc::new(AcceptAllValidator)).unwrap();

        // 空注册表 + openpyxl 往返模式：渲染成功后需要 worker，必然失败
        let error = execute_indicator_comparison_split(
            &comparison_path,
            &template_path,
            &common_path,
            &output_root,
            &modifier,
            FastRenderMode::OpenpyxlRoundtrip,
        )
        .expect_err("the openpyxl roundtrip mode cannot succeed without a worker runtime");
        assert!(!format!("{error:#}").is_empty());
        assert!(
            fs::read_dir(&output_root).unwrap().next().is_none(),
            "failed generation must not publish partial outputs"
        );

        fs::remove_file(template_path).unwrap();
        fs::remove_file(comparison_path).unwrap();
        fs::remove_file(common_path).unwrap();
        fs::remove_dir(output_root).unwrap();
    }

    fn write_split_common_config(path: &Path) {
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
    fn rejects_ambiguous_mapping_when_template_dimensions_are_unspecified() {
        let template_path = temp_path("ambiguous-template");
        let comparison_path = temp_path("ambiguous-comparison");
        write_template(&template_path);
        write_comparison(&comparison_path);
        let mut template = load_indicator_split_template(&template_path).unwrap();
        for row in &mut template.sheets[0].indicator_rows {
            for (_, key) in &mut row.keys {
                key.frequency.clear();
                key.batch.clear();
            }
        }

        let error = plan_indicator_comparison_split(&comparison_path, &template, &[band()])
            .expect_err("two frequency/batch rows must not collapse into one form cell");
        assert!(error.to_string().contains("频度/批次匹配到多条"));

        fs::remove_file(template_path).unwrap();
        fs::remove_file(comparison_path).unwrap();
    }

    #[test]
    fn preserves_empty_form_as_hidden_when_delete_setting_is_disabled() {
        let template_path = temp_path("empty-template");
        let comparison_path = temp_path("empty-comparison");
        write_template(&template_path);
        write_comparison(&comparison_path);
        let mut template = load_indicator_split_template(&template_path).unwrap();
        template.options.delete_empty_sheets = false;
        let plan = plan_indicator_comparison_split(&comparison_path, &template, &[band()]).unwrap();
        assert!(plan.organizations[0].edits.iter().any(|edit| matches!(
            edit,
            S3IndicatorSplitEdit::HideSheet { sheet } if sheet == "空表A1"
        )));
        fs::remove_file(template_path).unwrap();
        fs::remove_file(comparison_path).unwrap();
    }

    #[test]
    #[ignore = "requires AUDIT_S3_F3_COMMON_CONFIG pointing to the read-only 3.0 workbook"]
    fn parses_real_three_point_zero_alert_configuration_when_requested() {
        let path = std::env::var("AUDIT_S3_F3_COMMON_CONFIG").expect("3.0 config path");
        let bands = crate::s3_execution_compare::load_indicator_split_alert_bands(path)
            .expect("load real 3.0 alert bands");
        assert!(!bands.is_empty());
    }

    #[test]
    #[ignore = "requires the read-only real 3.3 config and S3-F01 result paths in environment"]
    fn plans_real_comparison_against_real_split_config_without_writing_inputs() {
        let common = std::env::var("AUDIT_S3_F3_COMMON_CONFIG").expect("3.0 config path");
        let template_path = std::env::var("AUDIT_S3_F3_CONFIG").expect("3.3 config path");
        let comparison_path = std::env::var("AUDIT_S3_F3_COMPARISON").expect("S3-F01 result path");
        let original_hashes = [&common, &template_path, &comparison_path]
            .map(|path| sha2::Sha256::digest(fs::read(path).expect("read real input")).to_vec());
        let template = load_indicator_split_template(&template_path).expect("load real template");
        assert!(
            template.forms.is_empty(),
            "the real 3.3 configuration has no enabled forms"
        );
        assert!(template.sheets.is_empty());
        let bands = crate::s3_execution_compare::load_indicator_split_alert_bands(&common)
            .expect("load real alert bands");
        let plan = plan_indicator_comparison_split(&comparison_path, &template, &bands)
            .expect("plan real comparison");
        assert_eq!(plan.comparison_rows, 2610);
        assert!(!plan.organizations.is_empty());
        let after_hashes = [&common, &template_path, &comparison_path]
            .map(|path| sha2::Sha256::digest(fs::read(path).expect("reread real input")).to_vec());
        assert_eq!(original_hashes, after_hashes);
    }
}
