//! Config-driven manual acceptance runner for the existing Audit entry points.
//!
//! The runner owns orchestration and evidence only. S1-F01 is routed through
//! its existing GUI workflow worker; S2 and S3 use the shipped CLI
//! subcommands. It never edits a supplied input workbook. S3-F03 receives a
//! runner-owned copy because its established contract updates its target in
//! place after creating a backup.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    io::Read,
    path::{Path, PathBuf},
    process::Command,
    time::Instant,
};

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AcceptanceManifest {
    pub s1: Option<S1AcceptanceInputs>,
    pub s2: Option<S2AcceptanceInputs>,
    pub s3: Option<S3AcceptanceInputs>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S1AcceptanceInputs {
    pub source_directory: PathBuf,
    pub template_directory: PathBuf,
    pub mapping_config: Option<PathBuf>,
    pub external_data: Option<PathBuf>,
    pub sidecar: Option<PathBuf>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S2AcceptanceInputs {
    pub current_directory: PathBuf,
    pub previous_directory: PathBuf,
    pub config: PathBuf,
    pub external: Option<PathBuf>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S3AcceptanceInputs {
    #[serde(default)]
    pub current_data: Vec<PathBuf>,
    #[serde(default)]
    pub previous_data: Vec<PathBuf>,
    pub common_config: Option<PathBuf>,
    pub execution_config: Option<PathBuf>,
    pub current_value_config: Option<PathBuf>,
    pub forms_config: Option<PathBuf>,
    pub comparison_workbook: Option<PathBuf>,
    pub feedback_target_workbook: Option<PathBuf>,
    #[serde(default)]
    pub feedback_workbooks: Vec<PathBuf>,
    #[serde(default)]
    pub expression_schema: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct AcceptanceRunReport {
    pub schema_version: u32,
    pub status: String,
    pub run_id: String,
    pub started_at: String,
    pub finished_at: String,
    pub elapsed_ms: u128,
    pub output_directory: PathBuf,
    pub steps: Vec<AcceptanceStepReport>,
    pub input_files: Vec<InputHashReport>,
    pub generated_files: Vec<PathBuf>,
}

#[derive(Clone, Debug, Serialize)]
pub struct AcceptanceStepReport {
    pub name: String,
    pub status: String,
    pub started_at: String,
    pub elapsed_ms: u128,
    pub exit_status: Option<i32>,
    pub input_files_unchanged: Option<bool>,
    pub changed_inputs: Vec<PathBuf>,
    pub generated_files: Vec<PathBuf>,
    pub key_counts: BTreeMap<String, Value>,
    pub stdout: String,
    pub stderr: String,
    pub error_original: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct InputHashReport {
    pub path: PathBuf,
    pub sha256_before: String,
    pub sha256_after: Option<String>,
    pub unchanged: Option<bool>,
}

/// Run selected workflows from a manifest and save JSON and Markdown evidence.
/// The output directory must be new and disjoint from every input path.
pub fn run_acceptance_manifest(
    manifest_path: &Path,
    output_directory: &Path,
) -> Result<AcceptanceRunReport> {
    let manifest_path = fs::canonicalize(manifest_path)
        .with_context(|| format!("读取验收清单 {}", manifest_path.display()))?;
    let manifest_text = fs::read_to_string(&manifest_path)
        .with_context(|| format!("读取验收清单内容 {}", manifest_path.display()))?;
    let mut manifest: AcceptanceManifest = serde_json::from_str(&manifest_text)
        .with_context(|| format!("解析验收清单 JSON {}", manifest_path.display()))?;
    let manifest_parent = manifest_path.parent().context("验收清单没有父目录")?;
    resolve_manifest_paths(&mut manifest, manifest_parent);

    let output_directory = absolute_path(output_directory)?;
    ensure!(
        !output_directory.exists(),
        "验收输出目录已存在，拒绝覆盖：{}",
        output_directory.display()
    );
    let output_parent = output_directory
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .context("验收输出目录必须有父目录")?;
    ensure!(output_parent.is_dir(), "验收输出目录的父目录不存在");
    let output_parent = fs::canonicalize(output_parent).context("解析验收输出父目录")?;
    let output_directory = output_parent.join(
        output_directory
            .file_name()
            .context("验收输出目录不能是文件系统根目录")?,
    );

    let input_roots = manifest_input_roots(&manifest, &manifest_path);
    for root in &input_roots {
        let canonical = fs::canonicalize(root)
            .with_context(|| format!("验收输入不存在：{}", root.display()))?;
        let disjoint = if canonical.is_dir() {
            !output_directory.starts_with(&canonical) && !canonical.starts_with(&output_directory)
        } else {
            canonical != output_directory && !canonical.starts_with(&output_directory)
        };
        ensure!(
            disjoint,
            "验收输出目录与输入路径重叠：{} 与 {}",
            output_directory.display(),
            canonical.display()
        );
    }

    fs::create_dir(&output_directory)
        .with_context(|| format!("创建验收输出目录 {}", output_directory.display()))?;
    let initial_hashes = hash_input_roots(&input_roots)?;
    ensure!(
        manifest.s1.is_some() || manifest.s2.is_some() || manifest.s3.is_some(),
        "验收清单至少需要提供 S1、S2 或 S3 中的一项"
    );

    let run_started = Instant::now();
    let started_at = Local::now().to_rfc3339();
    let run_id = Local::now().format("%Y%m%d_%H%M%S_%3f").to_string();
    let executable = std::env::current_exe().context("定位当前 Audit CLI")?;
    let mut steps = Vec::new();

    if let Some(inputs) = manifest.s1.as_ref() {
        let step_dir = output_directory.join("steps").join("S1-F01");
        let before = hash_input_roots(&input_roots)?;
        let step_started_at = Local::now().to_rfc3339();
        let step_started = Instant::now();
        let output = crate::web_bridge::run_s1_f1_for_acceptance(
            &inputs.source_directory,
            &inputs.template_directory,
            inputs.mapping_config.as_deref(),
            inputs.external_data.as_deref(),
            &step_dir,
            inputs.sidecar.as_deref(),
        );
        steps.push(finish_api_step(
            "S1-F01 模板审核",
            step_started_at,
            step_started.elapsed(),
            before,
            &input_roots,
            &step_dir,
            output,
        )?);
    }

    if let Some(inputs) = manifest.s2.as_ref() {
        let step_dir = output_directory.join("steps").join("S2-F01");
        let mut args = vec!["s2".to_owned()];
        push_arg(&mut args, "--current-dir", &inputs.current_directory);
        push_arg(&mut args, "--previous-dir", &inputs.previous_directory);
        push_arg(&mut args, "--config", &inputs.config);
        push_arg(&mut args, "--output-dir", &step_dir);
        push_optional_path(&mut args, "--external", inputs.external.as_deref());
        steps.push(run_cli_step(
            "S2-F01 两期报表审核",
            &executable,
            &args,
            &step_dir,
            &input_roots,
        )?);
    }

    if let Some(inputs) = manifest.s3.as_ref() {
        run_s3_steps(
            inputs,
            &executable,
            &output_directory.join("steps"),
            &input_roots,
            &mut steps,
        )?;
    }

    let final_hashes = hash_input_roots(&input_roots)?;
    let input_files = compare_hashes(&initial_hashes, &final_hashes);
    let input_mutated = input_files.iter().any(|item| item.unchanged != Some(true));
    let any_execution = steps
        .iter()
        .any(|step| matches!(step.status.as_str(), "completed" | "failed"));
    if !any_execution {
        steps.push(failed_precondition_step(
            "验收运行",
            "清单没有提供可执行的业务步骤".to_owned(),
        ));
    }
    let failed =
        !any_execution || steps.iter().any(|step| step.status == "failed") || input_mutated;
    let generated_files = list_files(&output_directory)?
        .into_iter()
        .filter(|path| {
            path.file_name().is_some_and(|name| {
                name != "acceptance-report.json" && name != "acceptance-report.md"
            })
        })
        .collect();
    let report = AcceptanceRunReport {
        schema_version: 1,
        status: if failed { "failed" } else { "completed" }.to_owned(),
        run_id,
        started_at,
        finished_at: Local::now().to_rfc3339(),
        elapsed_ms: run_started.elapsed().as_millis(),
        output_directory: output_directory.clone(),
        steps,
        input_files,
        generated_files,
    };
    let json_path = output_directory.join("acceptance-report.json");
    fs::write(&json_path, serde_json::to_vec_pretty(&report)?)
        .with_context(|| format!("写入验收 JSON 报告 {}", json_path.display()))?;
    let markdown_path = output_directory.join("acceptance-report.md");
    fs::write(&markdown_path, render_markdown(&report))
        .with_context(|| format!("写入验收 Markdown 报告 {}", markdown_path.display()))?;
    Ok(report)
}

fn run_s3_steps(
    inputs: &S3AcceptanceInputs,
    executable: &Path,
    root: &Path,
    input_roots: &[PathBuf],
    steps: &mut Vec<AcceptanceStepReport>,
) -> Result<()> {
    let f1_requested = !inputs.current_data.is_empty()
        || !inputs.previous_data.is_empty()
        || inputs.execution_config.is_some();
    let f1_output = root.join("S3-F01").join("comparison.xlsx");
    let mut comparison = inputs.comparison_workbook.clone();
    if f1_requested {
        if inputs.current_data.is_empty()
            || inputs.previous_data.is_empty()
            || inputs.common_config.is_none()
            || inputs.execution_config.is_none()
        {
            steps.push(failed_precondition_step(
                "S3-F01 两期数据→比较文件",
                "S3-F01 需要 currentData、previousData、commonConfig(3.0) 和 executionConfig(3.1)"
                    .to_owned(),
            ));
        } else {
            let mut args = vec!["s3-execution-compare".to_owned()];
            push_repeated_paths(&mut args, "--current-data", &inputs.current_data);
            push_repeated_paths(&mut args, "--previous-data", &inputs.previous_data);
            push_arg(
                &mut args,
                "--common-config",
                inputs.common_config.as_ref().unwrap(),
            );
            push_arg(
                &mut args,
                "--feature-config",
                inputs.execution_config.as_ref().unwrap(),
            );
            push_arg(&mut args, "--output-workbook", &f1_output);
            if let Some(schema) = inputs.expression_schema.as_deref() {
                push_arg_text(&mut args, "--expression-schema", schema);
            }
            args.push("--developer-view".to_owned());
            let step = run_cli_step(
                "S3-F01 两期数据→比较文件",
                executable,
                &args,
                &root.join("S3-F01"),
                input_roots,
            )?;
            if step.status == "completed" {
                comparison = Some(f1_output.clone());
            }
            steps.push(step);
        }
    }

    let f2_output = root.join("S3-F02").join("institution-explanations.xlsx");
    if let Some(comparison_path) = comparison.as_ref() {
        let mut args = vec!["s3-export-explanations".to_owned()];
        push_arg(&mut args, "--comparison-file", comparison_path);
        push_arg(&mut args, "--output-workbook", &f2_output);
        push_optional_path(
            &mut args,
            "--config-workbook",
            inputs.common_config.as_deref(),
        );
        args.push("--developer-view".to_owned());
        steps.push(run_cli_step(
            "S3-F02 比较文件→说明文件",
            executable,
            &args,
            &root.join("S3-F02"),
            input_roots,
        )?);
    } else {
        steps.push(not_requested_step(
            "S3-F02 比较文件→说明文件",
            "未提供比较文件，且 S3-F01 未产生比较文件".to_owned(),
        ));
    }

    if inputs.feedback_workbooks.is_empty() {
        steps.push(not_requested_step(
            "S3-F03 机构说明→说明文件",
            "未提供机构反馈工作簿".to_owned(),
        ));
    } else {
        let target = inputs
            .feedback_target_workbook
            .as_ref()
            .or_else(|| f2_output.is_file().then_some(&f2_output));
        if let Some(target) = target {
            let step_dir = root.join("S3-F03");
            fs::create_dir_all(&step_dir)?;
            let target_copy = step_dir.join("explanation-target-copy.xlsx");
            fs::copy(target, &target_copy).with_context(|| {
                format!("为 S3-F03 创建 runner 自有目标副本：{}", target.display())
            })?;
            let mut args = vec!["s3-import-explanation-feedback".to_owned()];
            push_arg(&mut args, "--target-workbook", &target_copy);
            push_repeated_paths(&mut args, "--feedback-workbook", &inputs.feedback_workbooks);
            args.push("--developer-view".to_owned());
            steps.push(run_cli_step(
                "S3-F03 机构反馈原地更新（runner 副本+备份）",
                executable,
                &args,
                &step_dir,
                input_roots,
            )?);
        } else {
            steps.push(failed_precondition_step(
                "S3-F03 机构说明→说明文件",
                "提供反馈工作簿时，必须同时提供 feedbackTargetWorkbook 或生成/提供比较文件以导出说明文件".to_owned(),
            ));
        }
    }

    let f4_requested = inputs.current_value_config.is_some();
    if f4_requested {
        if inputs.current_data.is_empty() || inputs.common_config.is_none() {
            steps.push(failed_precondition_step(
                "S3-F04 本期数值核对",
                "S3-F04 需要 currentData、commonConfig(3.0) 和 currentValueConfig(3.2)".to_owned(),
            ));
        } else {
            let step_dir = root.join("S3-F04");
            let mut args = vec!["s3-check-current-values".to_owned()];
            push_repeated_paths(&mut args, "--current-csv", &inputs.current_data);
            push_repeated_paths(&mut args, "--previous-csv", &inputs.previous_data);
            push_arg(
                &mut args,
                "--config-workbook",
                inputs.common_config.as_ref().unwrap(),
            );
            push_arg(
                &mut args,
                "--config-workbook",
                inputs.current_value_config.as_ref().unwrap(),
            );
            push_arg(
                &mut args,
                "--output-workbook",
                &step_dir.join("current-value-check.xlsx"),
            );
            push_arg(&mut args, "--log-workbook", &step_dir.join("run-log.xlsx"));
            args.push("--developer-view".to_owned());
            steps.push(run_cli_step(
                "S3-F04 本期数值核对",
                executable,
                &args,
                &step_dir,
                input_roots,
            )?);
        }
    } else {
        steps.push(not_requested_step(
            "S3-F04 本期数值核对",
            "未提供 3.2 本期数值核对配置".to_owned(),
        ));
    }

    if inputs.forms_config.is_some() {
        if let (Some(comparison_path), Some(common_config)) =
            (comparison.as_ref(), inputs.common_config.as_ref())
        {
            let step_dir = root.join("S3-F05");
            let mut args = vec!["s3-generate-financial-forms".to_owned()];
            push_arg(&mut args, "--comparison-workbook", comparison_path);
            push_arg(&mut args, "--common-config", common_config);
            push_arg(
                &mut args,
                "--forms-config",
                inputs.forms_config.as_ref().unwrap(),
            );
            push_arg(&mut args, "--output-directory", &step_dir);
            args.push("--developer-view".to_owned());
            steps.push(run_cli_step(
                "S3-F05 比较文件→金融表单",
                executable,
                &args,
                &step_dir,
                input_roots,
            )?);
        } else {
            steps.push(failed_precondition_step(
                "S3-F05 比较文件→金融表单",
                "S3-F05 需要 3.0 通用配置以及 S3-F01 生成或 comparisonWorkbook 指定的比较文件"
                    .to_owned(),
            ));
        }
    } else {
        steps.push(not_requested_step(
            "S3-F05 比较文件→金融表单",
            "未提供 3.3 指标比较拆分配置".to_owned(),
        ));
    }
    Ok(())
}

fn run_cli_step(
    name: &str,
    executable: &Path,
    args: &[String],
    step_directory: &Path,
    input_roots: &[PathBuf],
) -> Result<AcceptanceStepReport> {
    let started_at = Local::now().to_rfc3339();
    let started = Instant::now();
    fs::create_dir_all(step_directory).with_context(|| format!("创建 {} 输出目录", name))?;
    let before_hashes = hash_input_roots(input_roots)?;
    let output = Command::new(executable)
        .args(args)
        .output()
        .with_context(|| format!("启动 {} CLI 入口", name))?;
    let stdout = String::from_utf8_lossy(&output.stdout).into_owned();
    let stderr = String::from_utf8_lossy(&output.stderr).into_owned();
    let after_hashes = hash_input_roots(input_roots)?;
    let changed_inputs = changed_paths(&before_hashes, &after_hashes);
    let generated_files = list_files(step_directory)?;
    let output_missing = output.status.success() && generated_files.is_empty();
    let json_output = serde_json::from_str::<Value>(stdout.trim()).ok();
    let mut error_original = (!output.status.success()).then(|| {
        let raw = stderr.trim();
        if raw.is_empty() {
            stdout.trim().to_owned()
        } else {
            raw.to_owned()
        }
    });
    if !changed_inputs.is_empty() {
        let changed = changed_inputs
            .iter()
            .map(|path| path.display().to_string())
            .collect::<Vec<_>>()
            .join("; ");
        let notice = format!("验收步骤检测到输入哈希变化：{changed}");
        error_original = Some(match error_original {
            Some(previous) => format!("{previous}\n{notice}"),
            None => notice,
        });
    }
    if output_missing {
        error_original = Some(match error_original {
            Some(previous) => format!("{previous}\nCLI 返回成功但没有生成文件"),
            None => "CLI 返回成功但没有生成文件".to_owned(),
        });
    }
    Ok(AcceptanceStepReport {
        name: name.to_owned(),
        status: if output.status.success() && changed_inputs.is_empty() && !output_missing {
            "completed"
        } else {
            "failed"
        }
        .to_owned(),
        started_at,
        elapsed_ms: started.elapsed().as_millis(),
        exit_status: output.status.code(),
        input_files_unchanged: Some(changed_inputs.is_empty()),
        changed_inputs,
        generated_files,
        key_counts: json_output
            .as_ref()
            .map(extract_key_counts)
            .unwrap_or_default(),
        stdout,
        stderr,
        error_original,
    })
}

fn finish_api_step(
    name: &str,
    started_at: String,
    elapsed: std::time::Duration,
    before_hashes: BTreeMap<PathBuf, String>,
    input_roots: &[PathBuf],
    step_directory: &Path,
    output: Result<Value>,
) -> Result<AcceptanceStepReport> {
    let after_hashes = hash_input_roots(input_roots)?;
    let changed_inputs = changed_paths(&before_hashes, &after_hashes);
    let (stdout, stderr, exit_status, json_output, mut error_original) = match output {
        Ok(value) => {
            let completed = value["status"] == "completed";
            let error = (!completed).then(|| {
                let preflight_error = value["preflight"]["error"].as_str().unwrap_or_default();
                if !preflight_error.is_empty() {
                    preflight_error.to_owned()
                } else {
                    value["ui_status"]
                        .as_str()
                        .unwrap_or("S1-F01未完成")
                        .to_owned()
                }
            });
            (
                serde_json::to_string_pretty(&value)?,
                String::new(),
                Some(if completed { 0 } else { 2 }),
                Some(value),
                error,
            )
        }
        Err(error) => (
            String::new(),
            format!("{error:#}"),
            Some(2),
            None,
            Some(format!("{error:#}")),
        ),
    };
    if !changed_inputs.is_empty() {
        let notice = format!(
            "验收步骤检测到输入哈希变化：{}",
            changed_inputs
                .iter()
                .map(|path| path.display().to_string())
                .collect::<Vec<_>>()
                .join("; ")
        );
        error_original = Some(match error_original {
            Some(previous) => format!("{previous}\n{notice}"),
            None => notice,
        });
    }
    let success = exit_status == Some(0) && changed_inputs.is_empty();
    let generated_files = if step_directory.exists() {
        list_files(step_directory)?
    } else {
        Vec::new()
    };
    Ok(AcceptanceStepReport {
        name: name.to_owned(),
        status: if success { "completed" } else { "failed" }.to_owned(),
        started_at,
        elapsed_ms: elapsed.as_millis(),
        exit_status,
        input_files_unchanged: Some(changed_inputs.is_empty()),
        changed_inputs,
        generated_files,
        key_counts: json_output
            .as_ref()
            .map(extract_key_counts)
            .unwrap_or_default(),
        stdout,
        stderr,
        error_original,
    })
}

fn not_requested_step(name: &str, reason: String) -> AcceptanceStepReport {
    AcceptanceStepReport {
        name: name.to_owned(),
        status: "not_requested".to_owned(),
        started_at: Local::now().to_rfc3339(),
        elapsed_ms: 0,
        exit_status: None,
        input_files_unchanged: None,
        changed_inputs: Vec::new(),
        generated_files: Vec::new(),
        key_counts: BTreeMap::new(),
        stdout: String::new(),
        stderr: String::new(),
        error_original: Some(reason),
    }
}

fn failed_precondition_step(name: &str, reason: String) -> AcceptanceStepReport {
    let mut report = not_requested_step(name, reason.clone());
    report.status = "failed".to_owned();
    report.exit_status = Some(2);
    report
}

fn extract_key_counts(value: &Value) -> BTreeMap<String, Value> {
    const COUNT_KEYS: &[&str] = &[
        "issues",
        "selected_rows",
        "selectedRows",
        "imported",
        "unmatched",
        "conflicts",
        "data_rows",
        "dataRows",
        "triggered",
        "triggered_count",
        "triggeredCount",
        "formulaIssues",
        "conditionalIssues",
        "totalIssues",
        "rows",
        "matched",
        "filesTotal",
        "structurePassed",
        "skipped",
    ];
    fn visit(value: &Value, counts: &mut BTreeMap<String, Value>) {
        match value {
            Value::Object(object) => {
                for (key, child) in object {
                    if COUNT_KEYS.contains(&key.as_str()) {
                        let count = if let Some(array) = child.as_array() {
                            Some(json!(array.len()))
                        } else if child.is_number() || child.is_boolean() {
                            Some(child.clone())
                        } else {
                            None
                        };
                        if let Some(count) = count {
                            counts.insert(key.clone(), count);
                        }
                    }
                    visit(child, counts);
                }
            }
            Value::Array(array) => array.iter().for_each(|child| visit(child, counts)),
            _ => {}
        }
    }
    let mut counts = BTreeMap::new();
    visit(value, &mut counts);
    counts
}

fn manifest_input_roots(manifest: &AcceptanceManifest, manifest_path: &Path) -> Vec<PathBuf> {
    let mut roots = BTreeSet::from([manifest_path.to_path_buf()]);
    if let Some(s1) = &manifest.s1 {
        roots.insert(s1.source_directory.clone());
        roots.insert(s1.template_directory.clone());
        add_option(&mut roots, s1.mapping_config.as_ref());
        add_option(&mut roots, s1.external_data.as_ref());
        add_option(&mut roots, s1.sidecar.as_ref());
    }
    if let Some(s2) = &manifest.s2 {
        roots.insert(s2.current_directory.clone());
        roots.insert(s2.previous_directory.clone());
        roots.insert(s2.config.clone());
        add_option(&mut roots, s2.external.as_ref());
    }
    if let Some(s3) = &manifest.s3 {
        roots.extend(s3.current_data.iter().cloned());
        roots.extend(s3.previous_data.iter().cloned());
        for path in [
            s3.common_config.as_ref(),
            s3.execution_config.as_ref(),
            s3.current_value_config.as_ref(),
            s3.forms_config.as_ref(),
            s3.comparison_workbook.as_ref(),
            s3.feedback_target_workbook.as_ref(),
        ] {
            add_option(&mut roots, path);
        }
        roots.extend(s3.feedback_workbooks.iter().cloned());
    }
    roots.into_iter().collect()
}

fn add_option(roots: &mut BTreeSet<PathBuf>, value: Option<&PathBuf>) {
    if let Some(value) = value {
        roots.insert(value.clone());
    }
}

fn resolve_manifest_paths(manifest: &mut AcceptanceManifest, base: &Path) {
    fn resolve(path: &mut PathBuf, base: &Path) {
        if path.is_relative() {
            *path = base.join(&*path);
        }
    }
    if let Some(s1) = &mut manifest.s1 {
        resolve(&mut s1.source_directory, base);
        resolve(&mut s1.template_directory, base);
        for path in [
            &mut s1.mapping_config,
            &mut s1.external_data,
            &mut s1.sidecar,
        ]
        .into_iter()
        .flatten()
        {
            resolve(path, base);
        }
    }
    if let Some(s2) = &mut manifest.s2 {
        resolve(&mut s2.current_directory, base);
        resolve(&mut s2.previous_directory, base);
        resolve(&mut s2.config, base);
        if let Some(path) = &mut s2.external {
            resolve(path, base);
        }
    }
    if let Some(s3) = &mut manifest.s3 {
        for path in s3
            .current_data
            .iter_mut()
            .chain(s3.previous_data.iter_mut())
        {
            resolve(path, base);
        }
        for path in [
            &mut s3.common_config,
            &mut s3.execution_config,
            &mut s3.current_value_config,
            &mut s3.forms_config,
            &mut s3.comparison_workbook,
            &mut s3.feedback_target_workbook,
        ]
        .into_iter()
        .flatten()
        {
            resolve(path, base);
        }
        for path in &mut s3.feedback_workbooks {
            resolve(path, base);
        }
    }
}

fn hash_input_roots(roots: &[PathBuf]) -> Result<BTreeMap<PathBuf, String>> {
    let mut files = BTreeSet::new();
    for root in roots {
        let canonical = fs::canonicalize(root)
            .with_context(|| format!("验收输入不可读：{}", root.display()))?;
        if canonical.is_dir() {
            files.extend(list_files(&canonical)?);
        } else if canonical.is_file() {
            files.insert(canonical);
        } else {
            bail!("验收输入不是文件或目录：{}", root.display());
        }
    }
    let mut hashes = BTreeMap::new();
    for path in files {
        hashes.insert(path.clone(), sha256_file(&path)?);
    }
    Ok(hashes)
}

fn compare_hashes(
    before: &BTreeMap<PathBuf, String>,
    after: &BTreeMap<PathBuf, String>,
) -> Vec<InputHashReport> {
    let paths = before
        .keys()
        .chain(after.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    paths
        .into_iter()
        .map(|path| {
            let initial = before.get(&path).cloned();
            let final_hash = after.get(&path).cloned();
            InputHashReport {
                path,
                sha256_before: initial.clone().unwrap_or_default(),
                unchanged: Some(initial.is_some() && initial == final_hash),
                sha256_after: final_hash,
            }
        })
        .collect()
}

fn changed_paths(
    before: &BTreeMap<PathBuf, String>,
    after: &BTreeMap<PathBuf, String>,
) -> Vec<PathBuf> {
    compare_hashes(before, after)
        .into_iter()
        .filter(|record| record.unchanged != Some(true))
        .map(|record| record.path)
        .collect()
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut file =
        fs::File::open(path).with_context(|| format!("打开输入文件 {}", path.display()))?;
    let mut hasher = Sha256::new();
    let mut buffer = [0u8; 64 * 1024];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        hasher.update(&buffer[..read]);
    }
    Ok(format!("{:x}", hasher.finalize()))
}

fn list_files(root: &Path) -> Result<Vec<PathBuf>> {
    if root.is_file() {
        return Ok(vec![fs::canonicalize(root)?]);
    }
    if !root.exists() {
        return Ok(Vec::new());
    }
    let mut pending = vec![root.to_path_buf()];
    let mut files = Vec::new();
    while let Some(directory) = pending.pop() {
        for entry in
            fs::read_dir(&directory).with_context(|| format!("枚举目录 {}", directory.display()))?
        {
            let entry = entry?;
            let kind = entry.file_type()?;
            if kind.is_dir() {
                pending.push(entry.path());
            } else if kind.is_file() {
                files.push(fs::canonicalize(entry.path())?);
            }
        }
    }
    files.sort();
    Ok(files)
}

fn absolute_path(path: &Path) -> Result<PathBuf> {
    if path.is_absolute() {
        Ok(path.to_path_buf())
    } else {
        Ok(std::env::current_dir()?.join(path))
    }
}

fn push_arg(args: &mut Vec<String>, name: &str, value: &Path) {
    args.push(name.to_owned());
    args.push(value.to_string_lossy().into_owned());
}

fn push_arg_text(args: &mut Vec<String>, name: &str, value: &str) {
    args.push(name.to_owned());
    args.push(value.to_owned());
}

fn push_arg_opt(args: &mut Vec<String>, name: &str, value: Option<&str>) {
    if let Some(value) = value.filter(|value| !value.is_empty()) {
        args.push(name.to_owned());
        args.push(value.to_owned());
    }
}

fn push_optional_path(args: &mut Vec<String>, name: &str, value: Option<&Path>) {
    if let Some(value) = value {
        push_arg(args, name, value);
    }
}

fn push_repeated(args: &mut Vec<String>, name: &str, values: &[String]) {
    for value in values {
        args.push(name.to_owned());
        args.push(value.clone());
    }
}

fn push_repeated_paths(args: &mut Vec<String>, name: &str, values: &[PathBuf]) {
    for value in values {
        push_arg(args, name, value);
    }
}

fn render_markdown(report: &AcceptanceRunReport) -> String {
    let mut text = format!(
        "# Rust-Audit 业务验收报告\n\n- 状态：{}\n- 运行编号：{}\n- 开始：{}\n- 结束：{}\n- 耗时：{} ms\n- 输出目录：`{}`\n\n",
        report.status,
        report.run_id,
        report.started_at,
        report.finished_at,
        report.elapsed_ms,
        report.output_directory.display()
    );
    text.push_str("## 业务步骤\n\n| 步骤 | 状态 | 退出码 | 耗时 | 输入未变 | 输出文件数 |\n|---|---|---:|---:|---|---:|\n");
    for step in &report.steps {
        text.push_str(&format!(
            "| {} | {} | {} | {} ms | {} | {} |\n",
            step.name.replace('|', "\\|"),
            step.status,
            step.exit_status
                .map_or_else(|| "—".to_owned(), |value| value.to_string()),
            step.elapsed_ms,
            step.input_files_unchanged
                .map_or_else(|| "—".to_owned(), |value| value.to_string()),
            step.generated_files.len()
        ));
        if let Some(error) = step.error_original.as_deref() {
            text.push_str(&format!("\n错误原文：\n\n```text\n{}\n```\n", error));
        }
        if !step.key_counts.is_empty() {
            text.push_str(&format!(
                "\n关键计数：`{}`\n",
                serde_json::to_string(&step.key_counts).unwrap_or_default()
            ));
        }
    }
    text.push_str("\n## 输入文件哈希\n\n| 输入文件 | SHA-256 未变 |\n|---|---|\n");
    for input in &report.input_files {
        text.push_str(&format!(
            "| `{}` | {} |\n",
            input.path.display(),
            input
                .unchanged
                .map_or_else(|| "未验证".to_owned(), |value| value.to_string())
        ));
    }
    text.push_str("\n## 生成文件\n\n");
    for file in &report.generated_files {
        text.push_str(&format!("- `{}`\n", file.display()));
    }
    text
}

#[cfg(test)]
mod tests {
    use super::{changed_paths, extract_key_counts};
    use serde_json::json;
    use std::{collections::BTreeMap, path::PathBuf};

    #[test]
    fn runner_counts_nested_business_metrics_and_issue_arrays() {
        let value = json!({
            "summary": {"formulaIssues": 3, "conditionalIssues": 1},
            "selected_rows": 4,
            "issues": [{"cell": "A1"}, {"cell": "B2"}]
        });
        let counts = extract_key_counts(&value);
        assert_eq!(counts["formulaIssues"], json!(3));
        assert_eq!(counts["conditionalIssues"], json!(1));
        assert_eq!(counts["selected_rows"], json!(4));
        assert_eq!(counts["issues"], json!(2));
    }

    #[test]
    fn runner_detects_modified_added_and_removed_input_files() {
        let before = BTreeMap::from([
            (PathBuf::from("same.xlsx"), "a".to_owned()),
            (PathBuf::from("changed.xlsx"), "b".to_owned()),
            (PathBuf::from("removed.xlsx"), "c".to_owned()),
        ]);
        let after = BTreeMap::from([
            (PathBuf::from("same.xlsx"), "a".to_owned()),
            (PathBuf::from("changed.xlsx"), "changed".to_owned()),
            (PathBuf::from("added.xlsx"), "d".to_owned()),
        ]);
        assert_eq!(
            changed_paths(&before, &after),
            vec![
                PathBuf::from("added.xlsx"),
                PathBuf::from("changed.xlsx"),
                PathBuf::from("removed.xlsx")
            ]
        );
    }
}
