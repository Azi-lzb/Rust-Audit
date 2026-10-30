//! S5 F1 adapter for the existing Audit GUI.
//!
//! This module keeps the GUI entry thin and runs the existing Rust S5
//! configuration, retrieval, rule, and output services. It does not calculate
//! rules in JavaScript or modify selected input workbooks.

use crate::{
    s5_business::match_s5_business,
    s5_conditional_format::run_s5_conditional_format_with_template,
    s5_conditional_format::S5ConditionalFormatRun,
    s5_config::{load_configuration, Configuration, SourceContract},
    s5_config_dsl::{calls_in, parse_expression},
    s5_external_layout::ExternalLayout,
    s5_outputs::{
        file_sha256, write_batch_outputs_with_conditional_format, S5OutputRequest, S5OutputSource,
    },
    s5_period_recognition::detect_s5_period,
    s5_retrieval::{
        S5CurrentSourceBinding, S5ExternalSourceBinding, S5Retrieval, S5SourceDescriptor,
    },
    s5_scan::ScanBook,
    s5_service::{run_rules, select_rules_for_period, S5RuleOutcome},
    s5_template_matching::match_s5_template,
    s5_validation_formula_scan::S5ValidationFormulaScanner,
};
use serde_json::{json, Value};
use std::{
    collections::{BTreeMap, BTreeSet},
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
    pub additional_source_paths: Vec<PathBuf>,
    pub previous_source_paths: Vec<PathBuf>,
    pub previous_period: String,
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
    report_item: String,
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
        "previous_dir": "",
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

/// Execute S5 F1 for the selected current-workbook batch using the existing Rust
/// runtime and output contract.
pub fn run(request: &S5WebRunRequest) -> Result<Value, String> {
    let started = Instant::now();
    ensure_file(&request.config_path, "S5 审核配置")?;
    ensure_xlsx_family(&request.config_path, "S5 审核配置")?;
    let source_paths = normalize_s5_paths(
        std::iter::once(&request.source_path).chain(request.additional_source_paths.iter()),
        "S5 本期工作簿",
    )?;
    let previous_paths = normalize_s5_paths(request.previous_source_paths.iter(), "S5 上期工作簿")?;
    for path in request.template_path.iter() {
        ensure_file(path, "S5 匹配模板")?;
        ensure_xlsx_family(path, "S5 匹配模板")?;
    }
    if let Some(path) = &request.external_path {
        ensure_file(path, "S5 外部工作簿")?;
        ensure_xlsx_family(path, "S5 外部工作簿")?;
    }
    if !previous_paths.is_empty() && request.previous_period.trim().is_empty() {
        return Err("选择 S5 上期文件时必须填写上期报告期".to_owned());
    }

    let configuration = load_configuration(&request.config_path, false, "")
        .map_err(|error| format!("读取 S5 审核配置失败：{error}"))?;
    let template_path = match request.template_path.as_deref() {
        Some(path) => path.to_path_buf(),
        None => {
            let directory = request.template_dir.as_deref().ok_or_else(|| {
                "S5 未指定匹配模板，且没有可用模板目录；请先识别模板或手动选择模板".to_owned()
            })?;
            match_s5_template(&source_paths, directory)
                .map_err(|error| format!("S5 模板识别失败：{error}"))?
        }
    };
    let business = match_s5_business(&configuration, &source_paths, &template_path)
        .map_err(|error| format!("S5 业务识别失败：{error}"))?;
    if !configuration
        .rules
        .iter()
        .any(|rule| rule.business == business && rule.enabled)
    {
        return Err(format!("S5 业务“{business}”没有启用规则"));
    }
    let detection = detect_s5_period(&source_paths, &request.input_dir)
        .map_err(|error| format!("S5 报告期识别失败：{error}"))?;
    validate_previous_period(
        &request.previous_period,
        &detection.period,
        !previous_paths.is_empty(),
    )?;

    let mut preflight_configuration = configuration.clone();
    preflight_configuration.rules = select_rules_for_period(
        &configuration,
        &detection.period,
        std::slice::from_ref(&business),
    )
    .map_err(|error| format!("选择本期 S5 规则失败：{error}"))?
    .into_iter()
    .filter(|selected| selected.status_override.is_none() && selected.rule.enabled)
    .map(|selected| selected.rule.clone())
    .collect();
    let current_source_id = source_id_for_business(&preflight_configuration, &business)?;
    let previous_source_id = if previous_paths.is_empty() {
        current_source_id.clone()
    } else {
        previous_source_id_for_business(&preflight_configuration, &business, &current_source_id)?
    };
    let external_source_ids = external_source_ids_for_business(&preflight_configuration, &business);
    if external_source_ids.len() > 1 {
        return Err(format!(
            "当前 S5 GUI 使用一个外部工作簿，但业务“{business}”启用规则引用多个逻辑外部来源：{}；请按来源拆分审核配置或使用 CLI 多来源绑定",
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
    let external_source_id = request.external_path.as_ref().map(|_| {
        external_source_ids
            .first()
            .cloned()
            .unwrap_or_else(|| "EXT".to_owned())
    });

    let mut current_hashes = BTreeMap::new();
    for path in &source_paths {
        current_hashes.insert(path.clone(), file_sha256(path)?);
    }
    let mut previous_hashes = BTreeMap::new();
    for path in &previous_paths {
        previous_hashes.insert(path.clone(), file_sha256(path)?);
    }
    let external_hash = request
        .external_path
        .as_deref()
        .map(file_sha256)
        .transpose()?;
    let template_hash = file_sha256(&template_path)?;
    let formula_scanner = S5ValidationFormulaScanner::new(std::slice::from_ref(&template_path));

    let mut previous_file_issues = Vec::new();
    let mut previous_identities = Vec::with_capacity(previous_paths.len());
    for path in &previous_paths {
        let identity = resolve_previous_institution_identity(
            path,
            &preflight_configuration,
            &business,
            request.external_path.as_deref(),
        );
        if let Err(error) = &identity {
            previous_file_issues.push(json!({
                "source": path_text(path),
                "status": "failed",
                "error": error,
            }));
        }
        previous_identities.push((path.clone(), identity));
    }

    let mut outcomes = Vec::new();
    let mut data_quality_warnings: Vec<Value> = previous_file_issues
        .iter()
        .map(|issue| {
            json!({
                "kind": "previous_source_identity_failed",
                "source_file": Path::new(issue["source"].as_str().unwrap_or_default())
                    .file_name()
                    .unwrap_or_default()
                    .to_string_lossy(),
                "message": issue["error"].as_str().unwrap_or("无法识别上期工作簿机构身份"),
            })
        })
        .collect();
    let mut structure_observations = Vec::new();
    let mut conditional_runs = Vec::with_capacity(source_paths.len());
    let mut file_results = Vec::with_capacity(source_paths.len());
    let mut selected_template_paths = BTreeSet::new();
    selected_template_paths.insert(template_path.clone());
    let template_dir = request.template_dir.as_deref();

    for source_path in &source_paths {
        let identity = match resolve_institution_identity(
            source_path,
            &preflight_configuration,
            &business,
            request.external_path.as_deref(),
        ) {
            Ok(identity) => identity,
            Err(error) => {
                outcomes.push(file_failure_outcome(
                    source_path,
                    &business,
                    &detection.period,
                    "SOURCE_IDENTITY_FAILED",
                    error.clone(),
                ));
                file_results.push(json!({
                    "source": path_text(source_path),
                    "status": "failed",
                    "error": error,
                }));
                continue;
            }
        };

        let matched_previous: Vec<&PathBuf> = previous_identities
            .iter()
            .filter_map(|(path, result)| {
                result
                    .as_ref()
                    .ok()
                    .filter(|previous| previous.code.eq_ignore_ascii_case(&identity.code))
                    .map(|_| path)
            })
            .collect();
        if matched_previous.len() > 1 {
            let error = format!(
                "机构 {} 匹配到多份 S5 上期文件，请保留该机构的一份上期文件",
                identity.code
            );
            outcomes.push(file_failure_outcome(
                source_path,
                &business,
                &detection.period,
                "PREVIOUS_SOURCE_AMBIGUOUS",
                error.clone(),
            ));
            file_results.push(json!({
                "source": path_text(source_path),
                "status": "failed",
                "institution_code": identity.code,
                "institution_name": identity.name,
                "error": error,
            }));
            continue;
        }
        let previous_path = matched_previous.first().copied();
        if !previous_paths.is_empty() && previous_path.is_none() {
            data_quality_warnings.push(json!({
                "kind": "previous_source_unmatched",
                "source_file": source_path.file_name().unwrap_or_default().to_string_lossy(),
                "institution_code": identity.code,
                "message": "未匹配到同机构上期文件；涉及 V_PRE 的规则会报告取数失败",
            }));
        }

        let current_binding = S5CurrentSourceBinding {
            source: S5SourceDescriptor {
                path: source_path,
                source_id: &current_source_id,
                period: &detection.period,
                institution_code: &identity.code,
                institution_name: &identity.name,
            },
            business: &business,
        };
        let previous_descriptor = previous_path.map(|path| S5SourceDescriptor {
            path,
            source_id: &previous_source_id,
            period: request.previous_period.trim(),
            institution_code: &identity.code,
            institution_name: &identity.name,
        });
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
        let mut retrieval =
            match S5Retrieval::bind_with_current_sources_and_external_sources_and_templates(
                configuration.clone(),
                &[current_binding],
                previous_descriptor,
                &external_bindings,
                external_source_id.as_deref(),
                &detection.period,
                Some(&template_path),
                template_dir,
            ) {
                Ok(retrieval) => retrieval,
                Err(failure) => {
                    let error = format!("{}：{}", failure.code, failure.message);
                    outcomes.push(file_failure_outcome(
                        source_path,
                        &business,
                        &detection.period,
                        &failure.code,
                        error.clone(),
                    ));
                    file_results.push(json!({
                        "source": path_text(source_path),
                        "status": "failed",
                        "institution_code": identity.code,
                        "institution_name": identity.name,
                        "error": error,
                    }));
                    continue;
                }
            };

        let (mut current_outcomes, _) = match run_rules(
            &configuration,
            &mut retrieval,
            &detection.period,
            std::slice::from_ref(&business),
            &identity.code,
            &identity.name,
        ) {
            Ok(result) => result,
            Err(error) => {
                let message = format!("S5 规则执行失败：{error}");
                outcomes.push(file_failure_outcome(
                    source_path,
                    &business,
                    &detection.period,
                    "RULE_EXECUTION_FAILED",
                    message.clone(),
                ));
                file_results.push(json!({
                    "source": path_text(source_path),
                    "status": "failed",
                    "institution_code": identity.code,
                    "institution_name": identity.name,
                    "error": message,
                }));
                continue;
            }
        };
        let current_outcome_count = current_outcomes.len();
        for outcome in &mut current_outcomes {
            outcome.origin_source_path = path_text(source_path);
        }
        outcomes.append(&mut current_outcomes);
        let mut formula_scan = formula_scanner.scan(
            source_path,
            &business,
            &detection.period,
            &identity.code,
            &identity.name,
        );
        outcomes.append(&mut formula_scan.outcomes);
        data_quality_warnings.extend(retrieval.data_quality_warnings().iter().cloned());
        let mut observations = retrieval.structure_observations();
        for observation in &mut observations {
            observation.context = Some(json!({
                "current_source_path": path_text(source_path),
                "institution_code": identity.code,
                "institution_name": identity.name,
            }));
        }
        structure_observations.extend(observations);
        selected_template_paths.extend(
            retrieval
                .numeric_region_template_sources()
                .into_iter()
                .map(|template| template.path),
        );

        let conditional = match run_s5_conditional_format_with_template(
            source_path,
            Some(&template_path),
            &identity.name,
            &identity.report_item,
            &identity.code,
        ) {
            Ok(report) => report,
            Err(error) => conditional_format_failure(
                source_path,
                &identity.name,
                &identity.report_item,
                &identity.code,
                &format!("{error:#}"),
            ),
        };
        file_results.push(json!({
            "source": path_text(source_path),
            "status": "completed",
            "institution_code": identity.code,
            "institution_name": identity.name,
            "previous_source": previous_path.map(|path| path_text(path)),
            "rule_count": current_outcome_count,
            "template_formula_scan": formula_scan.summary,
            "conditional_format_triggers": conditional.triggered_count,
            "conditional_format_diagnostics": conditional.unsupported_count + conditional.error_count,
        }));
        conditional_runs.push(conditional);
    }

    let status_counts = count_statuses(&outcomes);
    let conditional_format = merge_conditional_format_runs(conditional_runs);
    let first_source = &source_paths[0];
    let first_source_hash = current_hashes
        .get(first_source)
        .expect("hash captured for primary S5 source");
    let mut additional_sources = Vec::new();
    for path in source_paths.iter().skip(1) {
        additional_sources.push((
            "current",
            current_source_id.as_str().to_owned(),
            path.clone(),
            current_hashes[path].clone(),
        ));
    }
    for path in &previous_paths {
        additional_sources.push((
            "previous",
            previous_source_id.clone(),
            path.clone(),
            previous_hashes[path].clone(),
        ));
    }
    if let (Some(path), Some(hash), Some(source_id)) = (
        request.external_path.as_ref(),
        external_hash.as_ref(),
        external_source_id.as_ref(),
    ) {
        additional_sources.push(("external", source_id.clone(), path.clone(), hash.clone()));
    }
    let additional_sources = additional_sources
        .iter()
        .map(|(role, source_id, path, hash)| S5OutputSource {
            role: *role,
            path,
            source_id: source_id.as_str(),
            sha256_before: hash.as_str(),
        })
        .collect::<Vec<_>>();
    let selected_template_owned = selected_template_paths
        .iter()
        .map(|path| {
            let hash = if path == &template_path {
                template_hash.clone()
            } else {
                file_sha256(path).unwrap_or_else(|_| "MISSING".to_owned())
            };
            (path.clone(), hash)
        })
        .collect::<Vec<_>>();
    let selected_templates = selected_template_owned
        .iter()
        .map(|(path, hash)| S5OutputSource {
            role: "layout",
            path,
            source_id: "TEMPLATE",
            sha256_before: hash,
        })
        .collect::<Vec<_>>();
    let schema_snapshot_dir = schema_snapshot_dir();
    let run_id = format!("s5-{}", chrono::Local::now().format("%Y%m%d_%H%M%S_%3f"));
    let summary_institution_code = if source_paths.len() == 1 {
        file_results[0]["institution_code"]
            .as_str()
            .unwrap_or_default()
    } else {
        ""
    };
    let summary_institution_name = if source_paths.len() == 1 {
        file_results[0]["institution_name"]
            .as_str()
            .unwrap_or_default()
    } else {
        "多个机构"
    };
    let output_dir = if request.output_dir.as_os_str().is_empty() {
        request.input_dir.join("执行结果")
    } else {
        request.output_dir.clone()
    };
    if output_dir.exists() && !output_dir.is_dir() {
        return Err(format!("S5 输出位置不是目录：{}", output_dir.display()));
    }
    let files_processed = file_results
        .iter()
        .filter(|record| record["status"] == "completed")
        .count();
    let files_failed = file_results.len() - files_processed;
    let batch = json!({
        "file_count": file_results.len(),
        "files_processed": files_processed,
        "files_failed": files_failed,
        "files": file_results,
        "previous_source_issues": previous_file_issues,
    });
    let output = write_batch_outputs_with_conditional_format(
        &S5OutputRequest {
            run_id: &run_id,
            period: &detection.period,
            institution_code: summary_institution_code,
            institution_name: summary_institution_name,
            outcomes: &outcomes,
            status_counts: &status_counts,
            data_quality_warnings: &data_quality_warnings,
            source_path: first_source,
            source_id: &current_source_id,
            source_sha256_before: first_source_hash,
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
        &batch,
    )
    .map_err(|error| format!("写入 S5 审核结果失败：{error}"))?;
    let elapsed_ms = started.elapsed().as_millis();
    Ok(json!({
        "status": output.run_status,
        "business": business,
        "period": detection.period,
        "institution_code": if source_paths.len() == 1 { json!(summary_institution_code) } else { Value::Null },
        "institution_name": summary_institution_name,
        "file_count": file_results.len(),
        "files_processed": files_processed,
        "files_failed": files_failed,
        "previous_source_issues": previous_file_issues,
        "files": file_results,
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

fn normalize_s5_paths<'a>(
    paths: impl Iterator<Item = &'a PathBuf>,
    label: &str,
) -> Result<Vec<PathBuf>, String> {
    let mut normalized = Vec::new();
    let mut seen = BTreeSet::new();
    for path in paths {
        ensure_file(path, label)?;
        ensure_xlsx_family(path, label)?;
        let canonical = fs::canonicalize(path)
            .map_err(|error| format!("解析{label}路径失败 {}：{error}", path.display()))?;
        if seen.insert(canonical.clone()) {
            normalized.push(canonical);
        }
    }
    if normalized.is_empty() && label == "S5 本期工作簿" {
        return Err("请至少勾选一份 S5 本期工作簿".to_owned());
    }
    Ok(normalized)
}

fn validate_previous_period(previous: &str, current: &str, selected: bool) -> Result<(), String> {
    if !selected {
        return Ok(());
    }
    let previous = chrono::NaiveDate::parse_from_str(previous.trim(), "%Y-%m-%d")
        .map_err(|_| "选择上期文件时，请填写有效上期报告期（YYYY-MM-DD）".to_owned())?;
    let current = chrono::NaiveDate::parse_from_str(current.trim(), "%Y-%m-%d")
        .map_err(|_| format!("S5 本期报告期不是有效日期：{current}"))?;
    if previous >= current {
        return Err("S5 上期报告期必须早于本期".to_owned());
    }
    Ok(())
}

fn file_failure_outcome(
    source_path: &Path,
    business: &str,
    period: &str,
    error_code: &str,
    error_message: String,
) -> S5RuleOutcome {
    S5RuleOutcome {
        origin_source_path: path_text(source_path),
        rule_id: format!(
            "FILE:{}",
            source_path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
        ),
        rule_name: "源文件处理失败".to_owned(),
        business: business.to_owned(),
        period: period.to_owned(),
        severity: "错误".to_owned(),
        check_type: "输入文件检查".to_owned(),
        metric: "源工作簿".to_owned(),
        status: "RETRIEVAL_FAILED".to_owned(),
        error_code: Some(error_code.to_owned()),
        error_evidence: json!({
            "source_path": source_path.to_string_lossy(),
            "sheet": "",
            "cell": "",
        }),
        error_message,
        ..S5RuleOutcome::default()
    }
}

fn count_statuses(outcomes: &[S5RuleOutcome]) -> BTreeMap<String, usize> {
    let mut counts = BTreeMap::new();
    for outcome in outcomes {
        *counts.entry(outcome.status.clone()).or_insert(0) += 1;
    }
    counts
}

fn conditional_format_failure(
    source_path: &Path,
    institution_name: &str,
    business: &str,
    institution_code: &str,
    error: &str,
) -> S5ConditionalFormatRun {
    S5ConditionalFormatRun {
        status: "PARTIAL".to_owned(),
        reader: "OOXML conditional-format evaluator".to_owned(),
        rules_total: 0,
        rules_supported: 0,
        scanned_cells: 0,
        triggered_count: 0,
        unsupported_count: 0,
        error_count: 1,
        read_time_ms: 0,
        evaluation_time_ms: 0,
        elapsed_ms: 0,
        findings: vec![crate::s5_conditional_format::S5ConditionalFormatFinding {
            institution_name: institution_name.to_owned(),
            report_item: business.to_owned(),
            institution_code: institution_code.to_owned(),
            sheet_name: String::new(),
            cell: String::new(),
            cell_value: String::new(),
            status: "ERROR".to_owned(),
            rule_type: "读取失败".to_owned(),
            formula: String::new(),
            applies_to: String::new(),
            priority: 0,
            source_file: source_path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .into_owned(),
            detail: error.to_owned(),
        }],
    }
}

fn merge_conditional_format_runs(runs: Vec<S5ConditionalFormatRun>) -> S5ConditionalFormatRun {
    let mut merged = S5ConditionalFormatRun {
        status: "COMPLETED_NO_RULES".to_owned(),
        reader: "Calamine + OOXML".to_owned(),
        rules_total: 0,
        rules_supported: 0,
        scanned_cells: 0,
        triggered_count: 0,
        unsupported_count: 0,
        error_count: 0,
        read_time_ms: 0,
        evaluation_time_ms: 0,
        elapsed_ms: 0,
        findings: Vec::new(),
    };
    for run in runs {
        merged.rules_total += run.rules_total;
        merged.rules_supported += run.rules_supported;
        merged.scanned_cells += run.scanned_cells;
        merged.triggered_count += run.triggered_count;
        merged.unsupported_count += run.unsupported_count;
        merged.error_count += run.error_count;
        merged.read_time_ms += run.read_time_ms;
        merged.evaluation_time_ms += run.evaluation_time_ms;
        merged.elapsed_ms += run.elapsed_ms;
        merged.findings.extend(run.findings);
    }
    if merged.unsupported_count + merged.error_count > 0 {
        merged.status = "PARTIAL".to_owned();
    } else if merged.rules_total > 0 {
        merged.status = "COMPLETED".to_owned();
    }
    merged
}

fn source_id_for_business(configuration: &Configuration, business: &str) -> Result<String, String> {
    let source_ids = current_source_ids_for_business(configuration, business);
    if source_ids.len() > 1 {
        return Err(format!(
            "S5 GUI 每批使用一个逻辑本期来源，但业务“{business}”启用了多个逻辑本期数据源：{}",
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
            return Err(format!("S5 业务“{business}”的规则数据源“{rule_source_id}”与本期来源配置“{contract_source_id}”不一致"));
        }
    }
    Ok(contract_source_id
        .map(str::to_owned)
        .or(rule_source_id)
        .unwrap_or_else(|| "CURRENT".to_owned()))
}

fn current_source_ids_for_business(
    configuration: &Configuration,
    business: &str,
) -> BTreeSet<String> {
    let mut source_ids = BTreeSet::new();
    let mut has_current_call = false;
    let mut source_free_defaults = BTreeSet::new();
    for rule in configuration
        .rules
        .iter()
        .filter(|rule| rule.business == business && rule.enabled)
    {
        let default_source = rule.default_source.trim();
        let mut uses_value_source = false;
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
            for call in calls_in(&node) {
                uses_value_source |= matches!(call.name.as_str(), "V_CUR" | "V_PRE" | "V_EXT");
                if call.name != "V_CUR" {
                    continue;
                }
                has_current_call = true;
                let source_id = call
                    .keywords
                    .iter()
                    .find(|(name, _)| name == "source")
                    .map(|(_, value)| value.trim())
                    .filter(|value| !value.is_empty());
                if let Some(source_id) = source_id {
                    source_ids.insert(source_id.to_owned());
                } else if !default_source.is_empty() {
                    source_ids.insert(default_source.to_owned());
                }
            }
        }
        if !uses_value_source && !default_source.is_empty() {
            source_free_defaults.insert(default_source.to_owned());
        }
    }

    if !has_current_call {
        source_ids.extend(source_free_defaults);
    }
    source_ids
}

fn previous_source_id_for_business(
    configuration: &Configuration,
    business: &str,
    current_source_id: &str,
) -> Result<String, String> {
    let mut source_ids = BTreeSet::new();
    for rule in configuration
        .rules
        .iter()
        .filter(|rule| rule.business == business && rule.enabled)
    {
        let mut previous_call_uses_default_source = false;
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
                .filter(|call| call.name == "V_PRE")
            {
                if let Some(source_id) = call
                    .keywords
                    .iter()
                    .find(|(name, _)| name == "source")
                    .map(|(_, value)| value.trim())
                    .filter(|source_id| !source_id.is_empty())
                {
                    source_ids.insert(source_id.to_owned());
                } else {
                    previous_call_uses_default_source = true;
                }
            }
        }
        if previous_call_uses_default_source {
            let default_source = rule.default_source.trim();
            if !default_source.is_empty() {
                source_ids.insert(default_source.to_owned());
            }
        }
    }
    if source_ids.len() > 1 {
        return Err(format!(
            "当前 S5 GUI 每家机构只绑定一份上期文件，但业务“{business}”的 V_PRE 规则引用多个逻辑来源：{}",
            source_ids.into_iter().collect::<Vec<_>>().join("、")
        ));
    }
    if let Some(source_id) = source_ids.into_iter().next() {
        return Ok(source_id);
    }
    let previous_contracts: Vec<&SourceContract> = configuration
        .source_contracts
        .iter()
        .filter(|contract| is_previous_source_type(&contract.source_type))
        .collect();
    if previous_contracts.len() > 1 {
        return Err(format!(
            "S5 配置为上期数据定义了多个来源契约：{}",
            previous_contracts
                .iter()
                .map(|contract| contract.source_id.as_str())
                .collect::<Vec<_>>()
                .join("、")
        ));
    }
    Ok(previous_contracts
        .first()
        .map(|contract| contract.source_id.clone())
        .unwrap_or_else(|| current_source_id.to_owned()))
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
    resolve_institution_identity_for_role(
        source_path,
        configuration,
        business,
        external_path,
        "current",
    )
}

fn resolve_previous_institution_identity(
    source_path: &Path,
    configuration: &Configuration,
    business: &str,
    external_path: Option<&Path>,
) -> Result<InstitutionIdentity, String> {
    resolve_institution_identity_for_role(
        source_path,
        configuration,
        business,
        external_path,
        "previous",
    )
}

fn resolve_institution_identity_for_role(
    source_path: &Path,
    configuration: &Configuration,
    business: &str,
    external_path: Option<&Path>,
    role: &str,
) -> Result<InstitutionIdentity, String> {
    let role_label = if role == "previous" {
        "上期"
    } else {
        "本期"
    };
    let source = ScanBook::open(source_path)
        .map_err(|error| format!("读取 S5 {role_label}工作簿机构身份失败：{error}"))?;
    let catalog = external_path
        .map(|path| {
            let book =
                ScanBook::open(path).map_err(|error| format!("读取外部机构参照失败：{error}"))?;
            ExternalLayout::institution_catalog(&book)
                .map_err(|error| format!("{}：{}", error.code, error.message))
        })
        .transpose()?
        .flatten();
    let contract = if role == "previous" {
        previous_contract(configuration, business)
    } else {
        current_contract(configuration, business)
    };
    let allowed_sheet = contract
        .map(|contract| contract.identity_sheet.trim())
        .filter(|sheet| !sheet.is_empty());
    if let Some(sheet) = allowed_sheet {
        if !source.sheet_names().iter().any(|name| name == sheet) {
            return Err(format!(
                "S5 {role_label}工作簿缺少配置指定的机构身份工作表：{sheet}"
            ));
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
                if code_labels.contains(&label)
                    && (valid_institution_code(value.trim())
                        || catalog
                            .as_ref()
                            .is_some_and(|catalog| catalog.contains_institution_code(value.trim())))
                {
                    codes.push((priority, value.trim().to_owned()));
                }
                if name_labels.contains(&label) {
                    names.push((priority, value.trim().to_owned()));
                }
            }
        }
    }
    let source_code = unique_highest_priority(
        codes,
        &format!("S5 {role_label}工作簿包含互相冲突的机构代码"),
    )?;
    let name_priority = names
        .iter()
        .map(|(priority, _)| *priority)
        .max()
        .unwrap_or_default();
    let mut seen_names = BTreeSet::new();
    let source_names = names
        .into_iter()
        .filter(|(priority, _)| *priority == name_priority)
        .map(|(_, name)| name)
        .filter(|name| seen_names.insert(normalize_label(name)))
        .collect::<Vec<_>>();
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
        name: source_names
            .first()
            .cloned()
            .unwrap_or_else(|| file_name_identity.to_owned()),
        report_item: String::new(),
    };

    let identity_mapped = catalog.is_some();
    if let Some(catalog) = catalog {
        let names = std::iter::once(file_name_identity.to_owned())
            .chain(source_names)
            .collect::<Vec<_>>();
        let (code, name, report_item) = catalog
            .resolve_institution_identity_from_names(&identity.code, &names)
            .map_err(|error| format!("{}：{}", error.code, error.message))?;
        if !code.is_empty() {
            identity.code = code;
        }
        identity.report_item = report_item;
        if !name.is_empty() {
            identity.name = name;
        }
    }
    if identity.code.is_empty() || (!identity_mapped && !valid_institution_code(&identity.code)) {
        return Err(
            "S5 工作簿未识别到有效机构代码；请检查身份字段，或选择含机构参照表的大集中数据"
                .to_owned(),
        );
    }
    Ok(identity)
}

fn current_contract<'a>(
    configuration: &'a Configuration,
    business: &str,
) -> Option<&'a SourceContract> {
    let rule_sources = current_source_ids_for_business(configuration, business);
    if rule_sources.len() == 1 {
        if let Some(contract) = configuration.source_contracts.iter().find(|contract| {
            is_current_source_type(&contract.source_type)
                && &contract.source_id == rule_sources.iter().next().unwrap()
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

fn previous_contract<'a>(
    configuration: &'a Configuration,
    business: &str,
) -> Option<&'a SourceContract> {
    let source_id = previous_source_id_for_business(configuration, business, "CURRENT").ok()?;
    configuration
        .source_contracts
        .iter()
        .find(|contract| {
            is_previous_source_type(&contract.source_type) && contract.source_id == source_id
        })
        .or_else(|| {
            let mut matches = configuration
                .source_contracts
                .iter()
                .filter(|contract| is_previous_source_type(&contract.source_type));
            let first = matches.next()?;
            matches.next().is_none().then_some(first)
        })
}

fn is_current_source_type(source_type: &str) -> bool {
    matches!(source_type.trim(), "current" | "本期")
}

fn is_previous_source_type(source_type: &str) -> bool {
    matches!(source_type.trim(), "previous" | "上期")
}

fn unique_highest_priority(
    candidates: Vec<(usize, String)>,
    conflict: &str,
) -> Result<Option<String>, String> {
    let Some(priority) = candidates.iter().map(|(priority, _)| *priority).max() else {
        return Ok(None);
    };
    let distinct: BTreeMap<String, String> = candidates
        .into_iter()
        .filter(|(candidate_priority, _)| *candidate_priority == priority)
        .map(|(_, value)| (normalize_label(&value), value))
        .collect();
    if distinct.len() > 1 {
        return Err(conflict.to_owned());
    }
    Ok(distinct.into_values().next())
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
    use super::{
        normalize_s5_paths, previous_source_id_for_business, resolve_institution_identity, run,
        source_id_for_business, validate_previous_period, S5WebRunRequest,
    };
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
    fn resolves_the_configured_vpre_source_and_matches_python_batch_defaults() {
        let mut configuration = minimal_configuration();
        let mut rule = source_rule("PREVIOUS_BOOK");
        rule.audit_expression = "V_PRE(\"贷款表\", \"合计\", \"本期\") > 0".to_owned();
        configuration.rules.push(rule);

        assert_eq!(
            source_id_for_business(&configuration, "单位贷款").unwrap(),
            "CURRENT"
        );

        assert_eq!(
            previous_source_id_for_business(&configuration, "单位贷款", "CURRENT_BOOK").unwrap(),
            "PREVIOUS_BOOK"
        );
    }

    #[test]
    fn explicit_vpre_source_overrides_the_rule_default_source() {
        let mut configuration = minimal_configuration();
        let mut rule = source_rule("CURRENT_BOOK");
        rule.audit_expression =
            "V_PRE(\"贷款表\", \"合计\", \"本期\", source=\"PREVIOUS_BOOK\") > 0".to_owned();
        configuration.rules.push(rule);

        assert_eq!(
            previous_source_id_for_business(&configuration, "单位贷款", "CURRENT_BOOK").unwrap(),
            "PREVIOUS_BOOK"
        );
    }

    #[test]
    fn current_source_preflight_uses_only_the_selected_revision_and_explicit_call_source() {
        let mut configuration = minimal_configuration();
        let mut old = source_rule("OLD");
        old.effective_to = "2026-06-30".to_owned();
        old.audit_expression = "V_CUR(\"贷款表\", \"合计\", \"本期\") > 0".to_owned();
        let mut current = old.clone();
        current.revision = 2;
        current.default_source = "UNUSED_DEFAULT".to_owned();
        current.effective_from = "2026-07-01".to_owned();
        current.effective_to.clear();
        current.audit_expression =
            "V_CUR(\"贷款表\", \"合计\", \"本期\", source=\"CURRENT_BOOK\") > 0".to_owned();
        configuration.rules = vec![old, current];
        let chosen = crate::s5_service::select_rules_for_period(
            &configuration,
            "2026-07-31",
            &["单位贷款".to_owned()],
        )
        .unwrap();
        let active = chosen
            .into_iter()
            .filter(|selection| selection.status_override.is_none() && selection.rule.enabled)
            .map(|selection| selection.rule.clone())
            .collect();
        configuration.rules = active;
        assert_eq!(
            source_id_for_business(&configuration, "单位贷款").unwrap(),
            "CURRENT_BOOK"
        );
    }

    #[test]
    fn selected_external_reference_rejects_unmapped_identity_and_invalid_flags() {
        let directory = unique_temp_dir();
        fs::create_dir_all(&directory).unwrap();
        let source = directory.join("示例机构_20260731.xlsx");
        let external = directory.join("external.xlsx");
        write_identity_book(&source, None);
        let write_reference = |code: &str, flag: &str| {
            let mut book = Workbook::new();
            let sheet = book.add_worksheet();
            sheet.set_name("机构参照表").unwrap();
            for (column, header) in ["统一社会信用代码", "机构名称", "报表项目", "禁用"]
                .iter()
                .enumerate()
            {
                sheet.write_string(0, column as u16, *header).unwrap();
            }
            for (column, value) in [code, "其他机构", "外部报表行", flag].iter().enumerate()
            {
                sheet.write_string(1, column as u16, *value).unwrap();
            }
            book.save(&external).unwrap();
        };
        write_reference("913205007040888888", "否");
        let error = resolve_institution_identity(
            &source,
            &minimal_configuration(),
            "单位贷款",
            Some(&external),
        )
        .unwrap_err();
        assert!(error.contains("IDENTITY_NOT_MAPPED"));
        write_reference("913205007040999999", "无效禁用值");
        let error = resolve_institution_identity(
            &source,
            &minimal_configuration(),
            "单位贷款",
            Some(&external),
        )
        .unwrap_err();
        assert!(error.contains("REFERENCE_FLAG_INVALID"));
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn validates_optional_s5_previous_report_period() {
        validate_previous_period("2026-06-30", "2026-07-31", true)
            .expect("previous report period precedes current period");
        assert!(validate_previous_period("2026-07-31", "2026-07-31", true).is_err());
        assert!(validate_previous_period("", "2026-07-31", true).is_err());
        validate_previous_period("", "2026-07-31", false)
            .expect("previous period is optional without previous files");
    }

    #[test]
    fn deduplicates_selected_s5_workbooks_without_opening_them() {
        let directory = unique_temp_dir();
        fs::create_dir_all(&directory).expect("create temp directory");
        let source = directory.join("示例机构_20260731.xlsx");
        fs::write(&source, b"workbook placeholder").expect("write placeholder input");
        let selected = vec![source.clone(), source.clone()];

        let normalized = normalize_s5_paths(selected.iter(), "S5 本期工作簿")
            .expect("normalize selected source files");
        assert_eq!(normalized, vec![fs::canonicalize(&source).unwrap()]);
        fs::remove_dir_all(directory).expect("remove temp directory");
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
            additional_source_paths: Vec::new(),
            previous_source_paths: Vec::new(),
            previous_period: String::new(),
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
