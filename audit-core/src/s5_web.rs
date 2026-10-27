//! S5 F1 adapter for the existing Audit GUI.
//!
//! This module keeps the GUI entry thin and runs the existing Rust S5
//! configuration, retrieval, rule, and output services. It does not calculate
//! rules in JavaScript or modify selected input workbooks.

use crate::{
    s5_business::match_s5_business,
    s5_config::{load_configuration, Configuration, SourceContract},
    s5_config_dsl::{calls_in, parse_expression},
    s5_conditional_format::run_s5_conditional_format,
    s5_external_layout::ExternalLayout,
    s5_outputs::{
        write_run_outputs_with_conditional_format, S5OutputRequest, S5OutputSource,
    },
    s5_period_recognition::detect_s5_period,
    s5_retrieval::{
        S5CurrentSourceBinding, S5ExternalSourceBinding, S5Retrieval, S5SourceDescriptor,
    },
    s5_scan::ScanBook,
    s5_service::run_rules,
    s5_template_matching::match_s5_template,
};
use serde_json::{json, Value};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

const CODE_LABELS: [&str; 10] = [
    "统一社会信用代码",
    "法人机构统一社会信用代码",
    "统一信用代码",
    "社会信用代码",
    "社会统一信用代码",
    "机构代码",
    "机构编码",
    "机构编号",
    "金融机构代码",
    "金融机构编码",
];
const NAME_LABELS: [&str; 4] = ["法人机构名称", "机构名称", "机构全称", "法人名称"];

#[derive(Clone, Debug)]
pub struct S5WebRunRequest {
    pub config_path: PathBuf,
    pub source_path: PathBuf,
    pub template_dir: Option<PathBuf>,
    pub template_path: Option<PathBuf>,
    pub external_path: Option<PathBuf>,
    pub input_dir: PathBuf,
    pub output_dir: PathBuf,
}

#[derive(Clone, Debug)]
struct InstitutionIdentity {
    code: String,
    name: String,
}

/// Resolve the same configuration and folder defaults used by the legacy S5
/// panel. Settings that require user input remain blank on startup.
pub fn initial_state(config_path: &Path) -> Value {
    let config_dir = config_path.parent().unwrap_or_else(|| Path::new("."));
    let config_names = [
        "5.逐笔统计系统_审核配置.xlsx",
        "S5_逐笔统计系统_审核配置.xlsx",
    ];
    let selected_config = config_names
        .iter()
        .map(|name| config_dir.join(name))
        .find(|path| path.is_file())
        .unwrap_or_else(|| config_dir.join(config_names[0]));
    let named_templates = config_dir.join("模板文件_命名区域");
    let legacy_templates = config_dir.join("S5_模板文件");
    let template_dir = if named_templates.is_dir() {
        named_templates
    } else if legacy_templates.is_dir() {
        legacy_templates
    } else {
        PathBuf::new()
    };
    let external_path = external_default(config_dir).unwrap_or_default();

    json!({
        "input_dir": "",
        "template_dir": path_text(&template_dir),
        "output_dir": "",
        "config_path": path_text(&selected_config),
        "template_path": "",
        "external_path": path_text(&external_path),
        "template_manual": false,
        "period_manual": false,
        "business_manual": false,
        "output_pinned": false,
        "business": "",
        "period": "",
        "external_period": "",
        "previous_files": [],
        "previous_period": "",
        "selected_files": [],
        "files": [],
        "configuration": {},
        "task": {},
        "progress": {},
        "configurationCheckTask": {},
        "recognition": {"errors": {}},
        "maintenanceTools": {},
        "maintenanceTask": {}
    })
}

fn external_default(config_dir: &Path) -> Option<PathBuf> {
    let fixed = config_dir.join("外部文件-大集中相关数据.xlsx");
    if fixed.is_file() {
        return Some(fixed);
    }
    let mut candidates = fs::read_dir(config_dir)
        .ok()?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.is_file()
                && path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with("外部文件-大集中相关数据"))
                && path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
        })
        .collect::<Vec<_>>();
    candidates.sort();
    candidates.into_iter().next()
}

fn path_text(path: &Path) -> String {
    if path.as_os_str().is_empty() {
        String::new()
    } else {
        path.to_string_lossy().into_owned()
    }
}

pub fn input_files(input_dir: &Path) -> Result<Vec<PathBuf>, String> {
    if !input_dir.is_dir() {
        return Err(format!("S5 源数据目录不存在：{}", input_dir.display()));
    }
    let mut paths = fs::read_dir(input_dir)
        .map_err(|error| format!("扫描 S5 源数据目录失败：{error}"))?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.is_file()
                && !path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with("~$"))
                && path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| {
                        extension.eq_ignore_ascii_case("xlsx")
                            || extension.eq_ignore_ascii_case("xlsm")
                    })
        })
        .collect::<Vec<_>>();
    paths.sort();
    Ok(paths)
}

pub fn classify_file_rows(paths: &[PathBuf], selected: &[String]) -> Vec<Value> {
    let selected: BTreeSet<&str> = selected.iter().map(String::as_str).collect();
    paths
        .iter()
        .map(|path| {
            let path_string = path_text(path);
            let period = crate::s5_period_recognition::s5_file_period_hint(path);
            json!({
                "path": path_string,
                "name": path.file_name().unwrap_or_default().to_string_lossy(),
                "template": "未识别",
                "match_status": "pending",
                "business": "",
                "period": period,
                "match_detail": ""
            })
        })
        .map(|mut row| {
            if selected.contains(row["path"].as_str().unwrap_or_default()) {
                row["selected"] = json!(true);
            }
            row
        })
        .collect()
}

/// Match selected files to the existing S5 template, business, and reporting
/// period recognizers. Any failed recognition remains explicit in the result.
pub fn recognize(
    config_path: &Path,
    selected_files: &[PathBuf],
    input_dir: &Path,
    template_dir: Option<&Path>,
    manual_template: Option<&Path>,
) -> Result<Value, String> {
    if selected_files.is_empty() {
        return Err("请至少勾选一个 S5 本期工作簿".to_owned());
    }
    let configuration = load_configuration(config_path, false, "")
        .map_err(|error| format!("读取 S5 审核配置失败：{error}"))?;
    let mut result = json!({"errors": {}});
    let template = if let Some(template) = manual_template {
        if !template.is_file() {
            return Err(format!("手动选择的 S5 模板不存在：{}", template.display()));
        }
        Some(template.to_path_buf())
    } else if let Some(directory) = template_dir {
        match match_s5_template(selected_files, directory) {
            Ok(path) => Some(path),
            Err(error) => {
                result["errors"]["template"] = json!(error.to_string());
                None
            }
        }
    } else {
        result["errors"]["template"] = json!("请选择 S5 模板目录或手动指定匹配模板");
        None
    };
    if let Some(template) = &template {
        result["template_path"] = json!(path_text(template));
        match match_s5_business(&configuration, selected_files, template) {
            Ok(business) => result["business"] = json!(business),
            Err(error) => result["errors"]["business"] = json!(error),
        }
    }
    match detect_s5_period(selected_files, input_dir) {
        Ok(detection) => {
            result["period"] = json!(detection.period);
            result["period_source"] = json!(detection.source);
            result["period_reader"] = json!(detection.reader);
        }
        Err(error) => result["errors"]["period"] = json!(error.to_string()),
    }
    Ok(result)
}

/// Execute S5 F1 for a selected current workbook using the existing Rust
/// runtime and output contract.
pub fn run(request: &S5WebRunRequest) -> Result<Value, String> {
    let started = Instant::now();
    ensure_file(&request.config_path, "S5 审核配置")?;
    ensure_file(&request.source_path, "S5 本期工作簿")?;
    ensure_xlsx_family(&request.config_path, "S5 审核配置")?;
    ensure_xlsx_family(&request.source_path, "S5 本期工作簿")?;
    if let Some(path) = &request.template_path {
        ensure_file(path, "S5 匹配模板")?;
        ensure_xlsx_family(path, "S5 匹配模板")?;
    }
    if let Some(path) = &request.external_path {
        ensure_file(path, "S5 外部工作簿")?;
        ensure_xlsx_family(path, "S5 外部工作簿")?;
    }

    let configuration = load_configuration(&request.config_path, false, "")
        .map_err(|error| format!("读取 S5 审核配置失败：{error}"))?;
    let template_path = match request.template_path.as_deref() {
        Some(path) => path.to_path_buf(),
        None => {
            let directory = request.template_dir.as_deref().ok_or_else(|| {
                "S5 未指定匹配模板，且没有可用模板目录；请先识别模板或手动选择模板".to_owned()
            })?;
            match_s5_template(std::slice::from_ref(&request.source_path), directory)
                .map_err(|error| format!("S5 模板识别失败：{error}"))?
        }
    };
    let business = match_s5_business(
        &configuration,
        std::slice::from_ref(&request.source_path),
        &template_path,
    )
    .map_err(|error| format!("S5 业务识别失败：{error}"))?;
    if !configuration
        .rules
        .iter()
        .any(|rule| rule.business == business && rule.enabled)
    {
        return Err(format!("S5 业务“{business}”没有启用规则"));
    }
    let detection = detect_s5_period(
        std::slice::from_ref(&request.source_path),
        &request.input_dir,
    )
    .map_err(|error| format!("S5 报告期识别失败：{error}"))?;
    let identity = resolve_institution_identity(
        &request.source_path,
        &configuration,
        &business,
        request.external_path.as_deref(),
    )?;
    let current_source_id = source_id_for_business(&configuration, &business)?;
    let external_source_ids = external_source_ids_for_business(&configuration, &business);
    if external_source_ids.len() > 1 {
        return Err(format!(
            "当前 S5 GUI 一次选择一个外部工作簿，但业务“{business}”启用规则引用多个逻辑外部来源：{}；请先在配置中统一数据源或扩展选择入口",
            external_source_ids.join("、")
        ));
    }
    if let Some(source_id) = external_source_ids.first() {
        if request.external_path.is_none() {
            return Err(format!(
                "S5 业务“{business}”的启用规则引用外部来源 {source_id}，请先选择大集中系统数据"
            ));
        }
    }
    let external_source_id = external_source_ids.first().cloned();
    let current_binding = S5CurrentSourceBinding {
        source: S5SourceDescriptor {
            path: &request.source_path,
            source_id: &current_source_id,
            period: &detection.period,
            institution_code: &identity.code,
            institution_name: &identity.name,
        },
        business: &business,
    };
    let external_binding = request
        .external_path
        .as_deref()
        .zip(external_source_id.as_deref())
        .map(|(path, source_id)| S5ExternalSourceBinding {
            source: S5SourceDescriptor {
                path,
                source_id,
                period: &detection.period,
                institution_code: &identity.code,
                institution_name: &identity.name,
            },
            business: &business,
        });
    let external_bindings: Vec<S5ExternalSourceBinding<'_>> =
        external_binding.into_iter().collect();
    let default_external_source = external_source_id.as_deref();
    let template_dir = request.template_dir.as_deref();
    let mut retrieval = S5Retrieval::bind_with_current_sources_and_external_sources_and_templates(
        configuration.clone(),
        &[current_binding],
        None,
        &external_bindings,
        default_external_source,
        &detection.period,
        Some(&template_path),
        template_dir,
    )
    .map_err(|failure| format!("{}：{}", failure.code, failure.message))?;
    let source_sha256_before = retrieval.source_sha256().to_owned();
    let (outcomes, status_counts) = run_rules(
        &configuration,
        &mut retrieval,
        &detection.period,
        std::slice::from_ref(&business),
        &identity.code,
        &identity.name,
    )
    .map_err(|error| format!("S5 规则执行失败：{error}"))?;
    let conditional_format = run_s5_conditional_format(
        &request.source_path,
        &identity.name,
        &business,
        &identity.code,
    )
    .map_err(|error| format!("S5 条件格式读取/求值失败：{error:#}"))?;

    let external_snapshots = retrieval.external_source_snapshots();
    let additional_sources = external_snapshots
        .iter()
        .map(|source| S5OutputSource {
            role: "external",
            path: &source.path,
            source_id: &source.source_id,
            sha256_before: &source.sha256,
        })
        .collect::<Vec<_>>();
    let templates = retrieval.numeric_region_template_sources();
    let selected_templates = templates
        .iter()
        .map(|template| S5OutputSource {
            role: "layout",
            path: &template.path,
            source_id: &template.source_id,
            sha256_before: &template.sha256,
        })
        .collect::<Vec<_>>();
    let structure_observations = retrieval.structure_observations();
    let schema_snapshot_dir = schema_snapshot_dir();
    let run_id = format!("s5-{}", chrono::Local::now().format("%Y%m%d_%H%M%S_%3f"));
    let output_dir = if request.output_dir.as_os_str().is_empty() {
        request.input_dir.join("执行结果")
    } else {
        request.output_dir.clone()
    };
    if output_dir.exists() && !output_dir.is_dir() {
        return Err(format!("S5 输出位置不是目录：{}", output_dir.display()));
    }
    let output = write_run_outputs_with_conditional_format(
        &S5OutputRequest {
            run_id: &run_id,
            period: &detection.period,
            institution_code: &identity.code,
            institution_name: &identity.name,
            outcomes: &outcomes,
            status_counts: &status_counts,
            data_quality_warnings: retrieval.data_quality_warnings(),
            source_path: &request.source_path,
            source_id: &current_source_id,
            source_sha256_before: &source_sha256_before,
            additional_sources: &additional_sources,
            selected_templates: &selected_templates,
            config_path: &request.config_path,
            config_sha256_before: &configuration.sha256,
            config_version: &configuration.version,
            structure_observations: &structure_observations,
            schema_snapshot_dir: Some(&schema_snapshot_dir),
        },
        &output_dir,
        Some(&conditional_format),
    )
    .map_err(|error| format!("写入 S5 审核结果失败：{error}"))?;
    let elapsed_ms = started.elapsed().as_millis();
    Ok(json!({
        "status": output.run_status,
        "business": business,
        "period": detection.period,
        "institution_code": identity.code,
        "institution_name": identity.name,
        "rule_count": outcomes.len(),
        "status_counts": output.status_counts,
        "triggered_rows": output.triggered_rows,
        "failure_rows": output.failure_rows,
        "conditional_formatting": {
            "status": conditional_format.status,
            "reader": conditional_format.reader,
            "rules_total": conditional_format.rules_total,
            "rules_supported": conditional_format.rules_supported,
            "scanned_cells": conditional_format.scanned_cells,
            "triggered_count": conditional_format.triggered_count,
            "unsupported_count": conditional_format.unsupported_count,
            "error_count": conditional_format.error_count,
            "elapsed_ms": conditional_format.elapsed_ms,
        },
        "conditional_format_trigger_count": output.conditional_format_trigger_count,
        "conditional_format_diagnostic_count": output.conditional_format_diagnostic_count,
        "run_dir": path_text(&output.run_dir),
        "result_workbook": path_text(&output.output_files.result_workbook),
        "run_ledger": path_text(&output.output_files.run_ledger),
        "input_integrity": output.input_integrity,
        "configuration_integrity": output.configuration_integrity,
        "elapsed_ms": elapsed_ms,
        "engine": "rust_s5_configured_audit",
        "reader": "calamine via excel-reader",
        "writer": "rust_xlsxwriter",
    }))
}

fn source_id_for_business(configuration: &Configuration, business: &str) -> Result<String, String> {
    let mut source_ids = BTreeSet::new();
    for rule in configuration
        .rules
        .iter()
        .filter(|rule| rule.business == business && rule.enabled)
    {
        let default_source = rule.default_source.trim();
        if !default_source.is_empty()
            && !matches!(default_source, "current" | "CURRENT" | "本期")
        {
            source_ids.insert(default_source.to_owned());
        }
        for expression in [
            rule.audit_expression.as_str(),
            rule.current_expression.as_str(),
            rule.comparison_expression.as_str(),
            rule.difference_expression.as_str(),
            rule.value_dictionary.as_str(),
        ] {
            let Ok(node) = parse_expression(expression) else {
                continue;
            };
            for call in calls_in(&node)
                .into_iter()
                .filter(|call| call.name == "V_CUR")
            {
                let source_id = call
                    .keywords
                    .iter()
                    .find(|(name, _)| name == "source")
                    .map(|(_, value)| value.trim())
                    .filter(|value| {
                        !value.is_empty() && !matches!(*value, "current" | "CURRENT" | "本期")
                    });
                if let Some(source_id) = source_id {
                    source_ids.insert(source_id.to_owned());
                }
            }
        }
    }

    if source_ids.len() > 1 {
        return Err(format!(
            "当前 S5 GUI 每次选择一个本期工作簿，但业务“{business}”启用了多个逻辑本期数据源：{}",
            source_ids.into_iter().collect::<Vec<_>>().join("、")
        ));
    }

    let rule_source_id = source_ids.into_iter().next();
    let contract_source_id = current_contract(configuration, business)
        .map(|contract| contract.source_id.trim())
        .filter(|source_id| !source_id.is_empty());
    if let (Some(rule_source_id), Some(contract_source_id)) =
        (rule_source_id.as_deref(), contract_source_id)
    {
        if rule_source_id != contract_source_id {
            return Err(format!(
                "S5 业务“{business}”的规则数据源“{rule_source_id}”与本期来源配置“{contract_source_id}”不一致"
            ));
        }
    }

    Ok(contract_source_id
        .map(str::to_owned)
        .or(rule_source_id)
        .unwrap_or_else(|| "CURRENT".to_owned()))
}

fn external_source_ids_for_business(configuration: &Configuration, business: &str) -> Vec<String> {
    let mut source_ids = BTreeSet::new();
    for rule in configuration
        .rules
        .iter()
        .filter(|rule| rule.business == business && rule.enabled)
    {
        for expression in [
            rule.audit_expression.as_str(),
            rule.current_expression.as_str(),
            rule.comparison_expression.as_str(),
            rule.difference_expression.as_str(),
            rule.value_dictionary.as_str(),
        ] {
            let Ok(node) = parse_expression(expression) else {
                continue;
            };
            for call in calls_in(&node)
                .into_iter()
                .filter(|call| call.name == "V_EXT")
            {
                let source_id = call
                    .keywords
                    .iter()
                    .find(|(name, _)| name == "source")
                    .map(|(_, value)| value.trim())
                    .filter(|value| !value.is_empty())
                    .unwrap_or("EXT");
                source_ids.insert(source_id.to_owned());
            }
        }
    }
    source_ids.into_iter().collect()
}

fn resolve_institution_identity(
    source_path: &Path,
    configuration: &Configuration,
    business: &str,
    external_path: Option<&Path>,
) -> Result<InstitutionIdentity, String> {
    let source = ScanBook::open(source_path)
        .map_err(|error| format!("读取本期工作簿机构身份失败：{error}"))?;
    let contract = current_contract(configuration, business);
    let allowed_sheet = contract
        .map(|contract| contract.identity_sheet.trim())
        .filter(|sheet| !sheet.is_empty());
    if let Some(sheet) = allowed_sheet {
        if !source.sheet_names().iter().any(|name| name == sheet) {
            return Err(format!("本期工作簿缺少配置指定的机构身份工作表：{sheet}"));
        }
    }
    let code_labels: BTreeSet<String> = CODE_LABELS
        .iter()
        .map(|label| normalize_label(label))
        .chain(
            contract
                .map(|contract| contract.institution_code_label.trim())
                .filter(|label| !label.is_empty())
                .map(normalize_label),
        )
        .collect();
    let name_labels: BTreeSet<String> = NAME_LABELS
        .iter()
        .map(|label| normalize_label(label))
        .chain(
            contract
                .map(|contract| contract.institution_name_label.trim())
                .filter(|label| !label.is_empty())
                .map(normalize_label),
        )
        .collect();

    let sheets = allowed_sheet
        .map(|sheet| vec![sheet.to_owned()])
        .unwrap_or_else(|| source.sheet_names().to_vec());
    let mut codes = Vec::new();
    let mut names = Vec::new();
    for sheet in sheets {
        for row in 1..=source.sheet_max_row(&sheet).min(12) {
            for column in 1..=source.sheet_max_col(&sheet).min(20) {
                if source.sheet_has_formula(&sheet, row, column) {
                    continue;
                }
                let label = normalize_label(&source.cell_raw_text(&sheet, row, column));
                if label.is_empty() {
                    continue;
                }
                let value = source.adjacent_value_for_label(&sheet, row, column);
                if value.trim().is_empty() || value.trim_start().starts_with('=') {
                    continue;
                }
                let priority = usize::from(label.starts_with("法人")) + 1;
                if code_labels.contains(&label) {
                    codes.push((priority, value.trim().to_owned()));
                }
                if name_labels.contains(&label) {
                    names.push((priority, value.trim().to_owned()));
                }
            }
        }
    }
    let source_code = unique_highest_priority(codes, "本期工作簿包含互相冲突的机构代码")?;
    let source_name = unique_highest_priority(names, "本期工作簿包含互相冲突的机构名称")?;
    let file_name = source_path
        .file_stem()
        .and_then(|name| name.to_str())
        .unwrap_or_default();
    let file_name_identity = file_name
        .split_once('_')
        .map(|(prefix, _)| prefix.trim())
        .filter(|prefix| !prefix.is_empty())
        .unwrap_or_default();
    let mut identity = InstitutionIdentity {
        code: source_code.unwrap_or_default(),
        name: source_name.unwrap_or_else(|| file_name_identity.to_owned()),
    };

    if let Some(external_path) = external_path {
        if let Ok(external_book) = ScanBook::open(external_path) {
            if let Ok(layout) = ExternalLayout::discover(&external_book) {
                match layout.resolve_institution_identity(&identity.code, &identity.name) {
                    Ok((code, name)) => {
                        identity.code = code;
                        if !name.is_empty() {
                            identity.name = name;
                        }
                    }
                    Err(error) if !valid_institution_code(&identity.code) => {
                        return Err(format!("无法通过外部机构参照解析身份：{}", error.message));
                    }
                    Err(error)
                        if matches!(
                            error.code.as_str(),
                            "INSTITUTION_DISABLED"
                                | "INSTITUTION_REFERENCE_CONFLICT"
                                | "IDENTITY_MAPPING_AMBIGUOUS"
                        ) =>
                    {
                        return Err(format!("外部机构参照校验失败：{}", error.message));
                    }
                    Err(_) => {}
                }
            }
        }
    }
    if !valid_institution_code(&identity.code) {
        return Err(
            "本期工作簿未识别到有效机构代码；请检查身份字段，或选择含机构参照表的大集中数据"
                .to_owned(),
        );
    }
    Ok(identity)
}

fn current_contract<'a>(
    configuration: &'a Configuration,
    business: &str,
) -> Option<&'a SourceContract> {
    let rule_sources: BTreeSet<&str> = configuration
        .rules
        .iter()
        .filter(|rule| rule.business == business && rule.enabled)
        .map(|rule| rule.default_source.trim())
        .filter(|source| !source.is_empty() && !matches!(*source, "current" | "本期"))
        .collect();
    if rule_sources.len() == 1 {
        if let Some(contract) = configuration.source_contracts.iter().find(|contract| {
            is_current_source_type(&contract.source_type)
                && contract.source_id == *rule_sources.iter().next().unwrap()
        }) {
            return Some(contract);
        }
    }
    let current_contracts: Vec<&SourceContract> = configuration
        .source_contracts
        .iter()
        .filter(|contract| is_current_source_type(&contract.source_type))
        .collect();
    if current_contracts.len() == 1 {
        current_contracts.into_iter().next()
    } else {
        None
    }
}

fn is_current_source_type(source_type: &str) -> bool {
    matches!(source_type.trim(), "current" | "本期")
}

fn unique_highest_priority(
    candidates: Vec<(usize, String)>,
    conflict: &str,
) -> Result<Option<String>, String> {
    let Some(priority) = candidates.iter().map(|(priority, _)| *priority).max() else {
        return Ok(None);
    };
    let distinct: BTreeSet<String> = candidates
        .into_iter()
        .filter(|(candidate_priority, _)| *candidate_priority == priority)
        .map(|(_, value)| value)
        .collect();
    if distinct.len() > 1 {
        return Err(conflict.to_owned());
    }
    Ok(distinct.into_iter().next())
}

fn normalize_label(value: &str) -> String {
    value
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect::<String>()
        .trim_end_matches([':', '：'])
        .to_lowercase()
}

fn valid_institution_code(code: &str) -> bool {
    let code = code.trim();
    (code.len() == 18 && code.bytes().all(|byte| byte.is_ascii_alphanumeric()))
        || ((4..=32).contains(&code.len())
            && code.bytes().all(|byte| byte.is_ascii_alphanumeric())
            && code.bytes().any(|byte| byte.is_ascii_alphabetic()))
}

fn ensure_file(path: &Path, label: &str) -> Result<(), String> {
    if path.is_file() {
        Ok(())
    } else {
        Err(format!("{label}不存在：{}", path.display()))
    }
}

fn ensure_xlsx_family(path: &Path, label: &str) -> Result<(), String> {
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default();
    if extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm") {
        Ok(())
    } else {
        Err(format!("{label}仅支持 XLSX/XLSM：{}", path.display()))
    }
}

fn schema_snapshot_dir() -> PathBuf {
    std::env::var_os("BASE_AUDIT_DATA_DIR")
        .filter(|value| !value.is_empty())
        .map(PathBuf::from)
        .or_else(|| {
            std::env::var_os("USERPROFILE")
                .or_else(|| std::env::var_os("HOME"))
                .map(|home| PathBuf::from(home).join(".audit-tools"))
        })
        .unwrap_or_else(|| PathBuf::from(".audit-tools"))
        .join("s5")
        .join(".schema_snapshots")
}

#[cfg(test)]
mod tests {
    use super::{resolve_institution_identity, run, source_id_for_business, S5WebRunRequest};
    use crate::s5_config::{Configuration, RuleDefinition};
    use excel_writer::Workbook;
    use std::{fs, path::Path, time::SystemTime};

    fn minimal_configuration() -> Configuration {
        Configuration {
            rules: Vec::new(),
            metrics: Vec::new(),
            external_metrics: Vec::new(),
            source_contracts: Vec::new(),
            applicability_scopes: Vec::new(),
            version: String::new(),
            version_status: String::new(),
            effective_from: String::new(),
            effective_to: String::new(),
            release_date: String::new(),
            release_note: String::new(),
            sha256: String::new(),
        }
    }

    fn source_rule(default_source: &str) -> RuleDefinition {
        RuleDefinition {
            rule_id: "TEST-S5-SOURCE".to_owned(),
            revision: 1,
            enabled: true,
            business: "单位贷款".to_owned(),
            default_source: default_source.to_owned(),
            default_sheet: "贷款表".to_owned(),
            name: "测试规则".to_owned(),
            severity: "硬性".to_owned(),
            check_type: "数值校验".to_owned(),
            metric: "余额".to_owned(),
            audit_expression: "True".to_owned(),
            description: String::new(),
            audit_expression_inverted: false,
            current_expression: String::new(),
            comparison_expression: String::new(),
            difference_expression: String::new(),
            value_dictionary: String::new(),
            effective_from: String::new(),
            effective_to: String::new(),
            applicable_expression: String::new(),
            applicability_scope: String::new(),
            numeric_mode: String::new(),
            formula_cell: String::new(),
        }
    }

    fn unique_temp_dir() -> std::path::PathBuf {
        let stamp = SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .expect("system clock after unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!("audit-s5-web-identity-{stamp}"))
    }

    fn write_identity_book(path: &Path, second_code: Option<&str>) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("封面").expect("name identity sheet");
        sheet
            .write_string(0, 0, "统一社会信用代码")
            .expect("write code label");
        sheet
            .write_string(0, 1, "913205007040999999")
            .expect("write code");
        sheet
            .write_string(1, 0, "机构名称")
            .expect("write name label");
        sheet.write_string(1, 1, "示例机构").expect("write name");
        if let Some(code) = second_code {
            sheet
                .write_string(2, 0, "社会信用代码")
                .expect("write second code label");
            sheet.write_string(2, 1, code).expect("write second code");
        }
        workbook.save(path).expect("save identity workbook");
    }

    #[test]
    fn reads_institution_identity_from_the_source_banner() {
        let directory = unique_temp_dir();
        fs::create_dir_all(&directory).expect("create temp directory");
        let source = directory.join("示例机构_20260731.xlsx");
        write_identity_book(&source, None);
        let result =
            resolve_institution_identity(&source, &minimal_configuration(), "测试业务", None)
                .expect("resolve source identity");
        assert_eq!(result.code, "913205007040999999");
        assert_eq!(result.name, "示例机构");
        fs::remove_dir_all(directory).expect("remove temp directory");
    }

    #[test]
    fn rejects_conflicting_institution_codes_from_source_banner() {
        let directory = unique_temp_dir();
        fs::create_dir_all(&directory).expect("create temp directory");
        let source = directory.join("示例机构_20260731.xlsx");
        write_identity_book(&source, Some("913205007040888888"));
        let error =
            resolve_institution_identity(&source, &minimal_configuration(), "测试业务", None)
                .expect_err("conflicting source identity is rejected");
        assert!(error.contains("互相冲突的机构代码"));
        fs::remove_dir_all(directory).expect("remove temp directory");
    }

    #[test]
    fn binds_the_unique_configured_current_source_without_a_source_contract() {
        let mut configuration = minimal_configuration();
        configuration.rules.push(source_rule("CUR"));

        assert_eq!(
            source_id_for_business(&configuration, "单位贷款").unwrap(),
            "CUR"
        );
    }

    #[test]
    fn rejects_multiple_current_sources_for_the_single_workbook_gui_entry() {
        let mut configuration = minimal_configuration();
        configuration.rules.push(source_rule("CUR"));
        configuration.rules.push(source_rule("CUR_2"));

        let error = source_id_for_business(&configuration, "单位贷款")
            .expect_err("one GUI workbook cannot satisfy multiple current source ids");
        assert!(error.contains("多个逻辑本期数据源"));
        assert!(error.contains("CUR"));
        assert!(error.contains("CUR_2"));
    }

    #[test]
    #[ignore = "requires explicit paths to read-only config-real S5 inputs"]
    fn runs_s5_gui_adapter_with_a_real_read_only_source() {
        let source = std::path::PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_SOURCE").expect("set AUDIT_S5_REAL_SOURCE"),
        );
        let config = std::path::PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_CONFIG").expect("set AUDIT_S5_REAL_CONFIG"),
        );
        let template_dir = std::path::PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_TEMPLATE_DIR").expect("set AUDIT_S5_REAL_TEMPLATE_DIR"),
        );
        let external = std::path::PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_EXTERNAL").expect("set AUDIT_S5_REAL_EXTERNAL"),
        );
        let output = unique_temp_dir();
        let source_before = fs::read(&source).expect("snapshot real S5 source");
        let input_dir = source
            .parent()
            .expect("real S5 source has parent")
            .to_path_buf();
        let summary = run(&S5WebRunRequest {
            config_path: config,
            source_path: source.clone(),
            template_dir: Some(template_dir),
            template_path: None,
            external_path: Some(external),
            input_dir,
            output_dir: output.clone(),
        })
        .expect("run real S5 GUI workflow");
        assert_eq!(summary["status"], "COMPLETED");
        let result_path = summary["result_workbook"]
            .as_str()
            .expect("S5 result workbook path");
        assert!(Path::new(result_path).is_file());
        let result_workbook =
            excel_reader::read_workbook(result_path).expect("read generated S5 result workbook");
        let result_sheet = result_workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "核查表校验结果")
            .expect("V2 result sheet");
        let expected_headers = [
            "工作簿名",
            "工作表名",
            "定位单元格",
            "错误类型",
            "校验类型",
            "校验指标",
            "描述",
            "当前值",
            "对比值",
            "差值",
            "规则编号",
            "历史校验说明",
            "审核意见",
        ];
        for (column, expected) in expected_headers.iter().enumerate() {
            let address = format!("{}1", (b'A' + column as u8) as char);
            assert_eq!(
                result_sheet
                    .values
                    .get(&address)
                    .and_then(|value| value.get("value"))
                    .and_then(serde_json::Value::as_str),
                Some(*expected),
                "real result header {address}"
            );
        }
        assert_eq!(
            fs::read(&source).expect("read real S5 source after run"),
            source_before
        );
        println!(
            "real S5 GUI run: rules={}, triggered={}, failures={}, elapsed_ms={}",
            summary["rule_count"].as_u64().unwrap_or_default(),
            summary["triggered_rows"].as_u64().unwrap_or_default(),
            summary["failure_rows"].as_u64().unwrap_or_default(),
            summary["elapsed_ms"].as_u64().unwrap_or_default()
        );
        fs::remove_dir_all(output).expect("remove isolated real S5 output");
    }
}
