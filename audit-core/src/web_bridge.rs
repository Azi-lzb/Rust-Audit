//! Local HTTP bridge for the bundled, unchanged Audit web frontend.
//!
//! This bridge exposes migrated Audit workflows through the existing frontend.

use crate::{
    discovery::{detect_period_from_workbooks, organisation_from_name},
    models::{CopyRange, Issue},
    native_file_dialog::{PickKind, PickRequest},
    name_config::{
        load_s1_conditional_mappings, named_areas_from_defined_names,
        resolve_s1_conditional_area_plan,
    },
    period_pairs::{list_period_pairs, report_source_warning},
    s1_conditional_audit::{
        extract_conditional_issues, S1ConditionalAuditResult, S1ConditionalMetadata,
    },
    s1_formula_audit::{extract_s1_formula_issues, S1FormulaMetadata},
    s1_formula_copy::{
        execute_s1_formula_copy_with_manual_recovery,
        execute_s1_formula_copy_with_manual_recovery_session, S1FormulaCopyOutcome,
    },
    s1_history::{process_s1_history, read_s1_history_workbook, S1HistoryRequest},
    s1_issue_chain::{merge_s1_issue_sources, S1IssueFilePair},
    s1_recalculation::{s1_recalculated_filename, S1RecalculationSession},
    s1_recovery::{
        file_sha256, stable_file_id, validate_recovery_fingerprints,
        workbook_recovery_fingerprints, S1RecoveryFile, S1RecoveryManifest, S1RecoveryStatus,
        S1RecoveryStore, S1RecoveryTaskSummary,
    },
    s1_region_summary::{s1_region_summary_source_workbooks, summarize_s1_named_regions_selected},
    s1_result_workbook::{write_s1_result_workbook, S1_RESULT_SHEET},
    s1_structure_compare::{
        compare_s1_source_structure, read_s1_structure_template_with_required_sheets,
        S1StructureTemplate,
    },
    s1_template::{load_template_profiles, recommend_template_with_profiles},
    s1_template_inspection::{inspect_s1_template, INSPECTION_ENGINE},
    s1_template_rules::{
        assess_s1_template_health, load_s1_template_rule_plan, S1TemplateRulePlan,
    },
    s1_validation_formula_scan::{
        issues_from_validation_formula_scan, S1ValidationFormulaScanPlan,
    },
    s2_config_check::check_s2_config,
    s2_service::{run_s2_audit, s2_engine_for_stage, S2ProgressEvent, S2RunRequest},
    s3_current_value_check::{
        input_reader_for_paths, python_regex_scope_matcher, run_current_value_check_files,
    },
    s3_execution_compare::{run_execution_compare_files, S3ExpressionSchema},
    s3_explanation_export::export_selected_explanations,
    s3_explanation_feedback::execute_explanation_feedback_import,
    s3_indicator_split::execute_indicator_comparison_split,
    s5_config_check::validate_s5_config,
    s5_web::{self, S5WebRunRequest},
};
use anyhow::{bail, ensure, Context, Result};
use axum::{
    body::Bytes,
    extract::{Path as RoutePath, State},
    http::{header, HeaderMap, HeaderValue, StatusCode},
    response::{Html, IntoResponse, Response},
    routing::{get, post},
    Json, Router,
};
use excel_editor::ExcelModifier;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
#[cfg(any(windows, target_os = "macos", target_os = "linux"))]
use std::process::Command;
use std::{
    fs,
    net::{IpAddr, Ipv4Addr, SocketAddr},
    path::{Path, PathBuf},
    sync::{Arc, Mutex, MutexGuard},
    thread,
    time::Instant,
};
use tokio::{net::TcpListener, runtime::Builder};

const ORIGINAL_HTML: &str = include_str!("../frontend/web/index.html");
const APP_ICON_PNG: &[u8] = include_bytes!("../frontend/assets/app-icon.png");
const APP_ICON_ICO: &[u8] = include_bytes!("../frontend/assets/app-icon.ico");
const APP_TITLE: &str = "Open-Data-Audit (Rust)";

const BRIDGE_JS: &str = r#"
<script>
(function () {
  function callApi(method, args) {
    return fetch('/api/' + method, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(args || [])
    }).then(function (response) {
      return response.json().then(function (data) {
        if (!response.ok || (data && data.error)) {
          throw new Error((data && data.error) || ('HTTP ' + response.status));
        }
        return data.result;
      });
    });
  }
  var api = new Proxy({}, {
    get: function (target, name) {
      if (typeof name !== 'string') { return undefined; }
      return function () {
        return callApi(name, Array.prototype.slice.call(arguments));
      };
    }
  });
  window.bridge = { api: api };
  window.addEventListener('load', function () {
    window.dispatchEvent(new Event('bridgeready'));
    if (typeof openCollectWorkbench === 'function') { openCollectWorkbench(); }
  });
})();
</script>
"#;

#[derive(Clone)]
struct AppState {
    ui: Arc<Mutex<UiState>>,
}

struct UiState {
    value: Value,
    started_at: Option<Instant>,
}

impl AppState {
    fn new(config_path: PathBuf) -> Self {
        let config_path = config_path.to_string_lossy().into_owned();
        let s5_state = s5_web::initial_state(Path::new(&config_path));
        let central_common_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.0大集中通用配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let central_cross_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.2大集中本期数值核对_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let central_comparison_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.1大集中执行比较_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let central_forms_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.3大集中指标比较拆分_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let history_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("逐笔统计系统_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let config_report = check_s2_config(&config_path);
        let openpyxl_status = crate::app_settings::openpyxl_evidence_status();
    let mut state = json!({
            "appInfo": {"title": APP_TITLE},
            "busy": false,
            "cancelRequested": false,
            "status": "Rust 迁移版已启动；当前接通报表采集 S2 功能",
            "log": [],
            "filePickerMode": "自动",
            "filePickerModeOptions": ["自动", "系统原生", "浏览器内置"],
            "uiTheme": "日间",
            "uiAccent": "默认蓝",
            "confirmBeforeRun": true,
            "showRunDetailLogs": false,
            "showPerformanceDiagnostics": false,
            "exportRunLogs": false,
            "recursiveDepth": -1,
            "developerAccess": {},
            "developerAccessMessage": "",
            "distributionUnlocked": false,
            "distributionProfile": "s1s2",
            "configIssues": {},
            "sourceFiles": [],
            "selectedFiles": [],
            "s1F1Preflight": {},
            "s1F1OutputPath": "",
            "s1F3RecoveryTasks": [],
            "s1F3PendingCount": 0,
            "s1F3LostCopyTasks": [],
            "s1F3LostCopyCount": 0,
            "s1F3RecoveryError": "",
            "historyConfig": history_config,
            "extraFiles": [],
            "mixedTemplates": [],
            "explanationFiles": [],
            "s5": s5_state,
            "periodSummary": {},
            "periodProgress": {
                "status": "idle",
                "stage": "",
                "label": "",
                "percent": 0,
                "current": 0,
                "total": 0,
                "message": "",
                "elapsedSeconds": 0
            },
            "centralCrossOutputAuto": true,
            "centralOutputAuto": true,
            "centralFormOutputAuto": true,
            "centralOutput": "",
            "centralFormOutput": "",
            "centralCompareFile": "",
            "centralExplainFile": "",
            "centralCommonConfig": central_common_config,
            "centralFormsConfig": central_forms_config,
            "centralFormOutputLastResult": "",
            "advancedSettings": {
                "calcEngine": crate::app_settings::AppSettings::load().calc_engine.as_str(),
                "editEngine": crate::app_settings::AppSettings::load().edit_engine.as_str(),
                "calcEngineOptions": crate::app_settings::CALC_ENGINE_LABELS
                    .iter()
                    .map(|(value, label)| json!({"value": value, "label": label}))
                    .collect::<Vec<_>>(),
                "editEngineOptions": crate::app_settings::EDIT_ENGINE_LABELS
                    .iter()
                    .map(|(value, label)| json!({"value": value, "label": label}))
                    .collect::<Vec<_>>(),
                "settingsPath": crate::app_settings::AppSettings::settings_path_for_display(),
                "openpyxlEvidence": {
                    "available": openpyxl_status.0,
                    "detail": openpyxl_status.1
                }
            },
            "engines": [{"value": "自动", "label": "Rust 迁移功能"}]
        });
        let s2_state = json!({
            "input": "",
            "templateDir": "",
            "template": "",
            "templateManual": false,
            "external": "",
            "output": "",
            "outputPinned": false,
            "pcCurDir": "",
            "pcPreDir": "",
            "pcCentral": "",
            "pcConfig": config_path,
            "pcOutput": "",
            "pcOutputAuto": true,
            "periodPairs": [],
            "periodSourceWarning": "",
            "periodConfigCheck": config_report
        });
        let s3_state = json!({
            "centralCrossCur": "",
            "centralCrossConfig": central_cross_config,
            "centralCrossLastResult": ""
        });
        let s3_execution_state = json!({
            "centralCur": "",
            "centralPre": "",
            "centralComparisonConfig": central_comparison_config,
            "centralExpressionSchema": "FIVE_SEGMENT_V1"
        });
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s2_state
                    .as_object()
                    .expect("S2 state is a JSON object")
                    .clone(),
            );
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s3_state
                    .as_object()
                    .expect("S3 state is a JSON object")
                    .clone(),
            );
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s3_execution_state
                    .as_object()
                    .expect("S3 execution compare state is a JSON object")
                    .clone(),
            );
        // 恢复上一次会话保存的用户偏好与路径（无效路径仅清空自身字段）。
        crate::app_settings::apply_preferences_to_state(&mut state);
        Self {
            ui: Arc::new(Mutex::new(UiState {
                value: state,
                started_at: None,
            })),
        }
    }

    fn lock(&self) -> MutexGuard<'_, UiState> {
        self.ui
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn get_state(&self) -> Value {
        let recovery = S1RecoveryStore::user_default()
            .and_then(|store| Ok((store.pending_tasks()?, store.lost_copy_tasks()?)));
        let mut ui = self.lock();
        apply_s1_f3_recovery_state(&mut ui.value, recovery);
        let mut value = ui.value.clone();
        if let Some(started_at) = ui.started_at {
            if value["periodProgress"]["status"] == "running" {
                value["periodProgress"]["elapsedSeconds"] =
                    json!(started_at.elapsed().as_secs_f64());
            }
        }
        value
    }
}

fn apply_s1_f3_recovery_state(
    state: &mut Value,
    recovery: Result<(Vec<S1RecoveryTaskSummary>, Vec<S1RecoveryTaskSummary>)>,
) {
    match recovery {
        Ok((pending, lost_copies)) => {
            state["s1F3PendingCount"] = json!(pending.len());
            state["s1F3RecoveryTasks"] = json!(pending);
            state["s1F3LostCopyCount"] = json!(lost_copies.len());
            state["s1F3LostCopyTasks"] = json!(lost_copies);
            state["s1F3RecoveryError"] = json!("");
        }
        Err(error) => {
            state["s1F3PendingCount"] = json!(0);
            state["s1F3RecoveryTasks"] = json!([]);
            state["s1F3LostCopyCount"] = json!(0);
            state["s1F3LostCopyTasks"] = json!([]);
            state["s1F3RecoveryError"] = json!(format!("{error:#}"));
        }
    }
}

/// Run the local-only frontend/API service. It never binds to a LAN address.
pub fn serve_web_ui(port: u16, config_path: PathBuf) -> Result<()> {
    let runtime = Builder::new_multi_thread()
        .enable_all()
        .build()
        .context("create Rust web runtime")?;
    runtime.block_on(async move {
        let address = SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port);
        let listener = TcpListener::bind(address)
            .await
            .with_context(|| format!("listen on http://{address}"))?;
        let address = listener
            .local_addr()
            .context("read the bound Audit UI address")?;
        let app = router(AppState::new(config_path));
        let url = format!("http://{address}");
        println!("Open-Data-Audit Rust web UI: {url}");
        if let Err(error) = open_default_browser(&url) {
            eprintln!("无法自动打开浏览器，请手动访问 {url}：{error}");
        }
        axum::serve(listener, app)
            .await
            .context("serve Audit web UI")
    })
}

fn open_default_browser(url: &str) -> std::io::Result<()> {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;

        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        Command::new("cmd")
            .args(["/C", "start", "", url])
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()?;
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open").arg(url).spawn()?;
    }

    #[cfg(target_os = "linux")]
    {
        Command::new("xdg-open").arg(url).spawn()?;
    }

    #[cfg(not(any(windows, target_os = "macos", target_os = "linux")))]
    {
        let _ = url;
        return Err(std::io::Error::new(
            std::io::ErrorKind::Unsupported,
            "automatic browser launch is unsupported on this platform",
        ));
    }

    Ok(())
}

fn router(state: AppState) -> Router {
    Router::new()
        .route("/", get(index))
        .route("/favicon.ico", get(favicon))
        .route("/assets/app-icon.png", get(app_icon_png))
        .route("/api/{method}", post(call_api))
        .with_state(state)
}

fn render_index() -> String {
    ORIGINAL_HTML
        .replace("__AUDIT_APP_TITLE__", APP_TITLE)
        .replace("__AUDIT_ICON_URL__", "/assets/app-icon.png")
        .replacen("</head>", &format!("{BRIDGE_JS}</head>"), 1)
}

async fn index() -> Response {
    if !ORIGINAL_HTML.contains("</head>") {
        return (
            StatusCode::INTERNAL_SERVER_ERROR,
            "Bundled Audit frontend is missing its head element.",
        )
            .into_response();
    }
    Html(render_index()).into_response()
}

async fn app_icon_png() -> Response {
    (
        [(header::CONTENT_TYPE, HeaderValue::from_static("image/png"))],
        APP_ICON_PNG,
    )
        .into_response()
}

async fn favicon() -> Response {
    (
        [(
            header::CONTENT_TYPE,
            HeaderValue::from_static("image/x-icon"),
        )],
        APP_ICON_ICO,
    )
        .into_response()
}

async fn call_api(
    State(state): State<AppState>,
    RoutePath(method): RoutePath<String>,
    headers: HeaderMap,
    body: Bytes,
) -> Response {
    if let Err(error) = validate_origin(&headers) {
        return error.into_response();
    }
    let args = match serde_json::from_slice::<Value>(&body) {
        Ok(Value::Array(args)) => args,
        Ok(_) => {
            return ApiError::new(StatusCode::BAD_REQUEST, "请求体必须是 JSON 数组")
                .into_response();
        }
        Err(error) => {
            return ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("请求体不是有效的 JSON：{error}"),
            )
            .into_response();
        }
    };

    if method == "pick_paths" {
        // 原生选择器需要 await（Windows 阻塞线程/Linux 门户），在进入同步
        // 分发前单独处理；不改状态、不做 busy 限制，路径校验仍由后续
        // choose_* 接口执行。
        return match pick_paths(&state, &args).await {
            Ok(result) => Json(json!({"result": result})).into_response(),
            Err(error) => error.into_response(),
        };
    }

    match dispatch_api(&state, &method, &args) {
        Ok(result) => Json(json!({"result": result})).into_response(),
        Err(error) => error.into_response(),
    }
}

fn validate_origin(headers: &HeaderMap) -> std::result::Result<(), ApiError> {
    let Some(origin) = headers.get(header::ORIGIN) else {
        return Ok(());
    };
    let Some(host) = headers
        .get(header::HOST)
        .and_then(|value| value.to_str().ok())
    else {
        return Err(ApiError::new(StatusCode::FORBIDDEN, "本机页面来源校验失败"));
    };
    let host_name = host
        .split(':')
        .next()
        .unwrap_or_default()
        .trim_matches(['[', ']']);
    if !matches!(host_name, "127.0.0.1" | "localhost") {
        return Err(ApiError::new(
            StatusCode::FORBIDDEN,
            "仅允许本机页面调用接口",
        ));
    }
    let origin = origin.to_str().unwrap_or_default();
    if origin != format!("http://{host}") {
        return Err(ApiError::new(
            StatusCode::FORBIDDEN,
            "仅允许同源本机页面调用接口",
        ));
    }
    Ok(())
}

fn dispatch_api(
    state: &AppState,
    method: &str,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    match method {
        "get_state" => Ok(state.get_state()),
        "choose_s5" => choose_s5(state, args),
        "update_s5" => update_s5(state, args),
        "recognize_s5_inputs" => recognize_s5_inputs(state, args),
        "start_s5" => start_s5(state, args),
        "check_s5_config" => check_s5_config(state, args),
        "get_s5_config_check" => get_s5_config_check(state),
        "check_period_audit_config" => check_period_audit_config(state),
        "choose_period_compare" => choose_period_compare(state, args),
        "choose_central" => choose_central(state, args),
        "choose_folder" => choose_folder(state, args),
        "choose_file" => choose_file(state, args),
        "choose_history_config" => choose_history_config(state, args),
        "clear_path" => clear_path(state, args),
        "recognize" => recognize_s1_sources(state),
        "set_selected_files" => set_selected_files(state, args),
        "inspect_template_formulas" => inspect_template_formulas(state),
        "start_dag_flow" => start_dag_flow(state, args),
        "list_s1_f3_tasks" => list_s1_f3_tasks(state),
        "start_s1_f3_workflow" => start_s1_f3_workflow(state, args),
        "start_central_comparison" => start_central_comparison(state),
        "start_central_explanation_export" => start_central_explanation_export(state),
        "import_central_explanation_feedback" => {
            start_central_explanation_feedback_import(state, args)
        }
        "start_central_indicator_split" => start_central_indicator_split(state),
        "check_central_forms_config" => check_central_forms_config(state),
        "check_central_common_config" => check_central_common_config(state),
        "check_central_expression_config" => check_central_expression_config(state),
        "check_central_cross_formulas" => check_central_cross_formulas(state),
        "start_central_cross_period" => start_central_cross_period(state),
        "toggle_central_output_pin" => toggle_central_output_pin(state, args),
        "toggle_period_output_pin" => toggle_period_output_pin(state),
        "browse_dir" => browse_dir(state, args),
        "start_period_compare_v2" => start_period_compare_v2(state),
        "update" => update_state(state, args),
        _ => Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版尚未实现接口：{method}"),
        )),
    }
}

fn choose_s5(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_s5 缺少路径类型"))?;
    if !matches!(
        field,
        "input_dir"
            | "previous_dir"
            | "template_dir"
            | "output_dir"
            | "config_path"
            | "template_path"
            | "external_path"
    ) {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("未知的 S5 路径字段：{field}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(state.lock().value["s5"][field].clone());
    }
    let path = PathBuf::from(unquote_path(selected));
    let is_directory = matches!(
        field,
        "input_dir" | "previous_dir" | "template_dir" | "output_dir"
    );
    if is_directory && !path.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选 S5 目录不存在：{}", path.display()),
        ));
    }
    if !is_directory {
        if !path.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("所选 S5 文件不存在：{}", path.display()),
            ));
        }
        let extension = path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default();
        let accepted = if field == "config_path" {
            extension.eq_ignore_ascii_case("xlsx")
        } else {
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
        };
        if !accepted {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                if field == "config_path" {
                    "S5 审核配置必须是 XLSX 工作簿".to_owned()
                } else {
                    "S5 工作簿仅支持 XLSX/XLSM".to_owned()
                },
            ));
        }
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S5 路径失败：{error}"),
        )
    })?;
    update_s5(state, &[json!({(field): path.to_string_lossy()})])
}

fn update_s5(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let values = args
        .first()
        .and_then(Value::as_object)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "update_s5 需要 JSON 对象"))?;
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改 S5 设置",
        ));
    }
    let form = ui.value["s5"]
        .as_object_mut()
        .expect("S5 UI state is initialized as an object");
    let previous_input = form
        .get("input_dir")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let previous_directory = form
        .get("previous_dir")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let previous_config = form
        .get("config_path")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let previous_template_dir = form
        .get("template_dir")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let previous_template = form
        .get("template_path")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let mut input_changed = false;
    let mut previous_directory_changed = false;
    let mut config_changed = false;
    let mut template_changed = false;
    for (key, value) in values {
        match key.as_str() {
            "input_dir" | "previous_dir" | "template_dir" | "output_dir" | "config_path"
            | "template_path" | "external_path" | "business" | "period" | "external_period"
            | "previous_period" => {
                let text = value.as_str().unwrap_or_default().trim().to_owned();
                match key.as_str() {
                    "input_dir" => input_changed = text != previous_input,
                    "previous_dir" => previous_directory_changed = text != previous_directory,
                    "config_path" => config_changed = text != previous_config,
                    "template_dir" => template_changed = text != previous_template_dir,
                    "template_path" => template_changed = text != previous_template,
                    _ => {}
                }
                form.insert(key.clone(), json!(text));
            }
            "output_pinned" | "template_manual" | "period_manual" | "business_manual" => {
                form.insert(key.clone(), json!(value.as_bool().unwrap_or(false)));
            }
            "selected_files" | "previous_files" => {
                let Some(paths) = value.as_array() else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "S5 选择文件必须是路径列表",
                    ));
                };
                let allowed: std::collections::BTreeSet<String> = if key == "selected_files" {
                    form.get("files")
                        .and_then(Value::as_array)
                        .into_iter()
                        .flatten()
                        .filter_map(|row| row["path"].as_str().map(str::to_owned))
                        .collect()
                } else {
                    let previous_dir = form
                        .get("previous_dir")
                        .and_then(Value::as_str)
                        .filter(|path| !path.is_empty())
                        .ok_or_else(|| {
                            ApiError::new(StatusCode::BAD_REQUEST, "请先选择 S5 上期文件目录")
                        })?;
                    s5_web::input_files(Path::new(previous_dir))
                        .map_err(|error| ApiError::new(StatusCode::BAD_REQUEST, error))?
                        .into_iter()
                        .map(|path| path.to_string_lossy().into_owned())
                        .collect()
                };
                let mut selected = Vec::new();
                for path in paths.iter().filter_map(Value::as_str) {
                    if !allowed.contains(path) {
                        return Err(ApiError::new(
                            StatusCode::BAD_REQUEST,
                            if key == "selected_files" {
                                format!("S5 选择的本期文件不在当前源目录清单中：{path}")
                            } else {
                                format!("S5 选择的上期文件不在上期目录清单中：{path}")
                            },
                        ));
                    }
                    if !selected.iter().any(|known| known == path) {
                        selected.push(path.to_owned());
                    }
                }
                form.insert(key.clone(), json!(selected));
            }
            _ => {}
        }
    }
    let input_dir = form
        .get("input_dir")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let files = if input_dir.is_empty() {
        Vec::new()
    } else {
        s5_web::input_files(Path::new(&input_dir))
            .map_err(|error| ApiError::new(StatusCode::BAD_REQUEST, error))?
    };
    let file_rows = s5_web::classify_file_rows(&files, &[]);
    let old_selected: Vec<String> = form
        .get("selected_files")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(Value::as_str)
        .map(str::to_owned)
        .collect();
    let allowed: std::collections::BTreeSet<String> = files
        .iter()
        .map(|path| path.to_string_lossy().into_owned())
        .collect();
    let selected = if input_changed {
        allowed.iter().cloned().collect::<Vec<_>>()
    } else {
        old_selected
            .into_iter()
            .filter(|path| allowed.contains(path))
            .collect()
    };
    form.insert("files".to_owned(), json!(file_rows));
    form.insert("selected_files".to_owned(), json!(selected));
    let previous_dir = form
        .get("previous_dir")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let previous_files = if previous_dir.is_empty() {
        Vec::new()
    } else {
        s5_web::input_files(Path::new(&previous_dir))
            .map_err(|error| ApiError::new(StatusCode::BAD_REQUEST, error))?
            .into_iter()
            .map(|path| path.to_string_lossy().into_owned())
            .collect::<Vec<_>>()
    };
    let old_previous_files: Vec<String> = form
        .get("previous_files")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(Value::as_str)
        .map(str::to_owned)
        .collect();
    let previous_allowed: std::collections::BTreeSet<&str> =
        previous_files.iter().map(String::as_str).collect();
    let selected_previous = if previous_directory_changed {
        previous_files.clone()
    } else {
        old_previous_files
            .into_iter()
            .filter(|path| previous_allowed.contains(path.as_str()))
            .collect()
    };
    form.insert("previous_files".to_owned(), json!(selected_previous));
    if input_changed {
        if form.get("output_pinned") != Some(&json!(true)) {
            form.insert(
                "output_dir".to_owned(),
                json!(if input_dir.is_empty() {
                    String::new()
                } else {
                    Path::new(&input_dir)
                        .join("执行结果")
                        .to_string_lossy()
                        .into_owned()
                }),
            );
        }
        form.insert("template_path".to_owned(), json!(""));
        form.insert("template_manual".to_owned(), json!(false));
        form.insert("business".to_owned(), json!(""));
        form.insert("business_manual".to_owned(), json!(false));
        form.insert("period".to_owned(), json!(""));
        form.insert("period_manual".to_owned(), json!(false));
        form.insert("recognition".to_owned(), json!({"errors": {}}));
    }
    if template_changed || config_changed {
        form.insert("recognition".to_owned(), json!({"errors": {}}));
        if config_changed {
            form.insert("business".to_owned(), json!(""));
            form.insert("business_manual".to_owned(), json!(false));
        }
    }
    if !values.contains_key("output_pinned") && values.contains_key("output_dir") {
        form.insert("output_pinned".to_owned(), json!(true));
    }
    if values.get("output_pinned").and_then(Value::as_bool) == Some(false) {
        form.insert(
            "output_dir".to_owned(),
            json!(if input_dir.is_empty() {
                String::new()
            } else {
                Path::new(&input_dir)
                    .join("执行结果")
                    .to_string_lossy()
                    .into_owned()
            }),
        );
    }
    if values.contains_key("template_path") {
        form.insert(
            "template_manual".to_owned(),
            json!(!form
                .get("template_path")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .is_empty()),
        );
    }
    if values.contains_key("business") {
        form.insert(
            "business_manual".to_owned(),
            json!(!form
                .get("business")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .is_empty()),
        );
    }
    if values.contains_key("period") {
        form.insert(
            "period_manual".to_owned(),
            json!(!form
                .get("period")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .is_empty()),
        );
    }
    let snapshot = Value::Object(form.clone());
    push_log(&mut ui.value, "已更新 S5 文件与配置设置".to_owned());
    Ok(snapshot)
}

fn recognize_s5_inputs(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let preserve_template = args.first().and_then(Value::as_bool).unwrap_or(false);
    let (config_path, input_dir, template_dir, template_path, selected_files, busy) = {
        let ui = state.lock();
        let form = &ui.value["s5"];
        (
            PathBuf::from(form["config_path"].as_str().unwrap_or_default()),
            PathBuf::from(form["input_dir"].as_str().unwrap_or_default()),
            form["template_dir"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            form["template_path"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            form["selected_files"]
                .as_array()
                .into_iter()
                .flatten()
                .filter_map(Value::as_str)
                .map(PathBuf::from)
                .collect::<Vec<_>>(),
            ui.value["busy"] == true,
        )
    };
    if busy {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能识别 S5 模板",
        ));
    }
    if preserve_template && template_path.is_none() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "请先手动选择匹配模板",
        ));
    }
    let recognition = s5_web::recognize(
        &config_path,
        &selected_files,
        &input_dir,
        template_dir.as_deref(),
        preserve_template
            .then_some(template_path.as_deref())
            .flatten(),
    )
    .map_err(|error| ApiError::new(StatusCode::BAD_REQUEST, error))?;
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务已开始，不能更新 S5 识别结果",
        ));
    }
    let form = ui.value["s5"]
        .as_object_mut()
        .expect("S5 UI state is initialized as an object");
    if !preserve_template {
        form.insert("template_manual".to_owned(), json!(false));
        form.insert(
            "template_path".to_owned(),
            recognition
                .get("template_path")
                .cloned()
                .unwrap_or(json!("")),
        );
    }
    if form.get("business_manual") != Some(&json!(true)) {
        form.insert("business".to_owned(), recognition["business"].clone());
    }
    if form.get("period_manual") != Some(&json!(true)) {
        form.insert("period".to_owned(), recognition["period"].clone());
    }
    form.insert("recognition".to_owned(), recognition.clone());
    if let Some(rows) = form.get_mut("files").and_then(Value::as_array_mut) {
        let selected: std::collections::BTreeSet<String> = selected_files
            .iter()
            .map(|path| path.to_string_lossy().into_owned())
            .collect();
        for row in rows {
            let is_selected = row["path"]
                .as_str()
                .is_some_and(|path| selected.contains(path));
            row["selected"] = json!(is_selected);
            if is_selected {
                row["template"] = recognition["template_path"]
                    .as_str()
                    .and_then(|path| Path::new(path).file_name())
                    .map(|name| json!(name.to_string_lossy()))
                    .unwrap_or(json!("未识别"));
                row["match_status"] = json!(if recognition["errors"]["template"].is_null() {
                    "matched"
                } else {
                    "unmatched"
                });
                row["business"] = recognition["business"].clone();
                row["period"] = recognition["period"].clone();
                row["match_detail"] = json!(recognition["errors"]);
            }
        }
    }
    let snapshot = Value::Object(form.clone());
    push_log(&mut ui.value, "S5 模板、业务和报告期识别完成".to_owned());
    Ok(snapshot)
}

fn start_s5(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let action = arg_string(args, 0).unwrap_or_default();
    if action != "S5-F01" {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            "S5 目前只支持 F1 汇总核查表校验",
        ));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Ok(json!({"accepted": false, "message": "已有任务正在运行，请等待完成"}));
    }
    let form = ui.value["s5"].clone();
    let selected: Vec<PathBuf> = form["selected_files"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(Value::as_str)
        .map(PathBuf::from)
        .collect();
    if selected.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "请至少勾选一份 S5 本期工作簿",
        ));
    }
    let mut selected = selected
        .iter()
        .map(|path| {
            fs::canonicalize(path).map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("解析 S5 本期工作簿失败 {}：{error}", path.display()),
                )
            })
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let mut seen_current = std::collections::BTreeSet::new();
    selected.retain(|path| seen_current.insert(path.clone()));
    let source_path = selected.remove(0);
    let additional_source_paths = selected;
    let previous_source_paths: Vec<PathBuf> = form["previous_files"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(Value::as_str)
        .map(PathBuf::from)
        .map(|path| {
            fs::canonicalize(&path).map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("解析 S5 上期工作簿失败 {}：{error}", path.display()),
                )
            })
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let mut seen_previous = std::collections::BTreeSet::new();
    let previous_source_paths = previous_source_paths
        .into_iter()
        .filter(|path| seen_previous.insert(path.clone()))
        .collect::<Vec<_>>();
    let previous_period = form["previous_period"]
        .as_str()
        .unwrap_or_default()
        .to_owned();
    if !previous_source_paths.is_empty() && previous_period.trim().is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "选择 S5 上期文件目录时必须填写上期报告期",
        ));
    }
    let input_dir = form["input_dir"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
        .or_else(|| source_path.parent().map(Path::to_path_buf))
        .unwrap_or_default();
    let output_dir = form["output_dir"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| input_dir.join("执行结果"));
    let config_path = form["config_path"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请选择 S5 审核配置工作簿"))?;
    let template_path = form["template_path"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from);
    let template_dir = form["template_dir"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from);
    let external_path = form["external_path"]
        .as_str()
        .filter(|path| !path.is_empty())
        .map(PathBuf::from);
    let files_total = additional_source_paths.len() + 1;
    let request = S5WebRunRequest {
        config_path,
        source_path,
        additional_source_paths,
        previous_source_paths,
        previous_period,
        template_dir,
        template_path,
        external_path,
        input_dir,
        output_dir,
    };
    let task_id = next_s5_task_id();
    ui.value["busy"] = json!(true);
    ui.value["cancelRequested"] = json!(false);
    ui.value["status"] = json!("S5-F01 正在运行 Rust 配置化审核……");
    ui.value["periodProgress"] = json!({
        "status": "running", "stage": "audit", "label": "S5-F01 配置化审核",
        "percent": 5, "current": 0, "total": files_total,
        "message": "读取配置并准备所选工作簿批次", "elapsedSeconds": 0
    });
    ui.value["s5"]["task"] = json!({
        "task_id": task_id, "system": "S5", "function": action,
        "task_state": "queued", "files_total": files_total,
        "started_at_epoch": chrono::Local::now().timestamp_millis() as f64 / 1000.0
    });
    ui.value["s5"]["progress"] = json!({
        "phase": "queued", "message": "准备审核", "percent": 0,
        "file_index": 0, "files_total": files_total
    });
    push_log(
        &mut ui.value,
        "S5-F01 已启动：Rust 配置、动态取数与规则审核".to_owned(),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S5-F01 engine=rust_s5_configured_audit；reader=calamine via excel-reader；writer=rust_xlsxwriter；本期文件={} 份；主输入={}；上期文件={} 份（报告期 {}）；配置={}；输出目录={}",
            files_total,
            request.source_path.display(),
            request.previous_source_paths.len(),
            request.previous_period,
            request.config_path.display(),
            request.output_dir.display()
        ),
    );
    drop(ui);
    let worker_state = state.clone();
    let worker_task_id = task_id.clone();
    thread::spawn(move || run_s5_worker(worker_state, worker_task_id, request));
    Ok(json!({"accepted": true, "task_id": task_id}))
}

fn run_s5_worker(state: AppState, task_id: String, request: S5WebRunRequest) {
    let started = Instant::now();
    let requested_files_total = request.additional_source_paths.len() + 1;
    {
        let mut ui = state.lock();
        ui.value["s5"]["task"]["task_state"] = json!("running");
        ui.value["s5"]["progress"]["phase"] = json!("audit");
        ui.value["s5"]["progress"]["message"] = json!("正在读取配置、定位表格并执行审核规则");
        ui.value["s5"]["progress"]["percent"] = json!(10);
        push_log(&mut ui.value, "S5-F01 正在读取配置与本期工作簿".to_owned());
    }
    let result = s5_web::run(&request);
    let elapsed = started.elapsed().as_secs_f64();
    let mut ui = state.lock();
    match result {
        Ok(result) => {
            let run_status = result["status"].as_str().unwrap_or("FAILED");
            let failure_rows = result["failure_rows"].as_u64().unwrap_or_default();
            let files_total = result["file_count"].as_u64().unwrap_or_default();
            let files_processed = result["files_processed"].as_u64().unwrap_or_default();
            let files_failed = result["files_failed"].as_u64().unwrap_or_default();
            let previous_source_issues = result["previous_source_issues"]
                .as_array()
                .map(Vec::len)
                .unwrap_or_default();
            let conditional_diagnostics = result["conditional_formatting"]["unsupported_count"]
                .as_u64()
                .unwrap_or_default()
                + result["conditional_formatting"]["error_count"]
                    .as_u64()
                    .unwrap_or_default();
            let completed = matches!(
                run_status,
                "COMPLETED"
                    | "COMPLETED_WITH_CONDITIONAL_FORMAT_LIMITATIONS"
                    | "COMPLETED_WITH_ISSUES"
            );
            let task_state = if !completed {
                "failed"
            } else if failure_rows > 0 || conditional_diagnostics > 0 || previous_source_issues > 0
            {
                "partial_failed"
            } else {
                "succeeded"
            };
            let message = format!(
                "S5-F01 {}：文件处理 {}/{}（失败 {}），上期身份失败 {} 份，业务 {}，报告期 {}，规则 {} 条，业务触发 {} 条，执行失败 {} 条，条件格式触发 {} 条、待处理诊断 {} 条",
                if task_state == "succeeded" {
                    "完成"
                } else if task_state == "failed" {
                    "失败"
                } else {
                    "完成但需处理"
                },
                files_processed,
                files_total,
                files_failed,
                previous_source_issues,
                result["business"].as_str().unwrap_or(""),
                result["period"].as_str().unwrap_or(""),
                result["rule_count"].as_u64().unwrap_or_default(),
                result["triggered_rows"].as_u64().unwrap_or_default(),
                failure_rows,
                result["conditional_format_trigger_count"]
                    .as_u64()
                    .unwrap_or_default(),
                conditional_diagnostics
            );
            let output_path = result["result_workbook"].as_str().unwrap_or_default();
            ui.value["s5"]["task"] = json!({
                "task_id": task_id,
                "system": "S5",
                "function": "S5-F01",
                "task_state": task_state,
                "files_total": files_total,
                "message": message,
                "result": result,
                "result_path": output_path,
                "elapsed_seconds": elapsed
            });
            ui.value["s5"]["progress"] = json!({
                "phase": if task_state == "failed" { "failed" } else { "completed" }, "message": message, "percent": 100,
                "file_index": files_total, "files_total": files_total, "elapsedSeconds": elapsed
            });
            ui.value["status"] = json!(message);
            ui.value["periodProgress"] = json!({
                "status": if task_state == "failed" { "failed" } else { "completed" }, "stage": "audit", "label": "S5-F01 配置化审核",
                "percent": 100, "current": files_total, "total": files_total,
                "message": message, "elapsedSeconds": elapsed
            });
            push_log(&mut ui.value, message.clone());
            if let Some(files) = result["files"].as_array() {
                for file in files {
                    if file["status"] == "failed" {
                        push_log(
                            &mut ui.value,
                            format!(
                                "S5-F01 文件处理失败：{}；{}",
                                file["source"].as_str().unwrap_or_default(),
                                file["error"].as_str().unwrap_or("未知原因")
                            ),
                        );
                    }
                }
            }
            if let Some(issues) = result["previous_source_issues"].as_array() {
                for issue in issues {
                    push_log(
                        &mut ui.value,
                        format!(
                            "S5-F01 上期文件身份识别失败：{}；{}",
                            issue["source"].as_str().unwrap_or_default(),
                            issue["error"].as_str().unwrap_or("未知原因")
                        ),
                    );
                }
            }
            push_log(&mut ui.value, format!("S5-F01 结果工作簿：{output_path}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S5-F01 结束：engine={}；reader={}；writer={}；状态={}；耗时 {:.2} 秒；结果={}",
                    result["engine"]
                        .as_str()
                        .unwrap_or("rust_s5_configured_audit"),
                    result["reader"]
                        .as_str()
                        .unwrap_or("calamine via excel-reader"),
                    result["writer"].as_str().unwrap_or("rust_xlsxwriter"),
                    run_status,
                    elapsed,
                    output_path
                ),
            );
        }
        Err(error) => {
            let message = format!("S5-F01 执行失败：{error}");
            ui.value["s5"]["task"] = json!({
                "task_id": task_id,
                "system": "S5",
                "function": "S5-F01",
                "task_state": "failed",
                "files_total": requested_files_total,
                "message": message,
                "elapsed_seconds": elapsed
            });
            ui.value["s5"]["progress"]["message"] = json!(message);
            ui.value["s5"]["progress"]["elapsedSeconds"] = json!(elapsed);
            ui.value["status"] = json!(message);
            ui.value["periodProgress"]["status"] = json!("failed");
            ui.value["periodProgress"]["message"] = json!(message);
            ui.value["periodProgress"]["elapsedSeconds"] = json!(elapsed);
            push_log(&mut ui.value, message.clone());
            push_detail_log(
                &mut ui.value,
                format!(
                    "S5-F01 失败；engine=rust_s5_configured_audit；耗时 {:.2} 秒；错误：{error}",
                    elapsed
                ),
            );
        }
    }
    ui.value["busy"] = json!(false);
    ui.value["cancelRequested"] = json!(false);
}

fn check_s5_config(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let scope = arg_string(args, 0).unwrap_or("current_business").to_owned();
    if !matches!(scope.as_str(), "current_business" | "all_enabled") {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("未知的 S5 配置检查范围：{scope}"),
        ));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Ok(json!({"accepted": false, "message": "其他任务正在运行"}));
    }
    let form = ui.value["s5"].clone();
    let config_path = PathBuf::from(form["config_path"].as_str().unwrap_or_default());
    if !config_path.is_file() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "请选择 S5 审核配置文件",
        ));
    }
    if scope == "current_business"
        && form["template_path"]
            .as_str()
            .unwrap_or_default()
            .is_empty()
    {
        return Ok(json!({
            "accepted": false,
            "message": "当前业务检查需要先点击“识别模板”或手动选择匹配模板"
        }));
    }
    if scope == "all_enabled" && form["template_dir"].as_str().unwrap_or_default().is_empty() {
        return Ok(json!({"accepted": false, "message": "全量检查需要有效的模板文件目录"}));
    }
    let task_id = next_s5_task_id();
    ui.value["s5"]["configurationCheckTask"] = json!({
        "task_id": task_id,
        "task_state": "queued",
        "scope": scope,
        "scope_label": if scope == "all_enabled" { "全量启用配置" } else { "当前业务启用配置" },
        "message": "正在准备 S5 配置检查",
        "progress": {"percent": 0, "message": "准备读取审核配置"}
    });
    drop(ui);
    let worker_state = state.clone();
    let worker_task_id = task_id.clone();
    thread::spawn(move || {
        {
            let mut ui = worker_state.lock();
            ui.value["s5"]["configurationCheckTask"]["task_state"] = json!("running");
            ui.value["s5"]["configurationCheckTask"]["progress"] = json!({
                "percent": 10, "message": "正在检查配置、模板与外部引用"
            });
        }
        let result = validate_s5_config(
            &config_path,
            form["template_path"].as_str().unwrap_or_default(),
            form["template_dir"].as_str().unwrap_or_default(),
            "",
            form["external_path"].as_str().unwrap_or_default(),
            form["business"].as_str().unwrap_or_default(),
            &scope,
        );
        let mut ui = worker_state.lock();
        match result {
            Ok(report) => {
                let valid = report["valid"].as_bool().unwrap_or(false);
                let complete = report["complete"].as_bool().unwrap_or(false);
                let task_state = if !valid {
                    "failed"
                } else if !complete {
                    "incomplete"
                } else {
                    "succeeded"
                };
                ui.value["s5"]["configurationCheckTask"] = json!({
                    "task_id": worker_task_id,
                    "task_state": task_state,
                    "scope": scope,
                    "message": if valid && complete { "S5 配置检查完成" } else { "S5 配置检查发现问题或未验证项" },
                    "progress": {"percent": 100, "message": "配置检查完成"},
                    "report": report
                });
            }
            Err(error) => {
                ui.value["s5"]["configurationCheckTask"] = json!({
                    "task_id": worker_task_id,
                    "task_state": "failed",
                    "scope": scope,
                    "message": error.to_string(),
                    "progress": {"percent": 100, "message": "配置检查失败"}
                });
            }
        }
    });
    Ok(json!({"accepted": true, "task_id": task_id}))
}

fn get_s5_config_check(state: &AppState) -> std::result::Result<Value, ApiError> {
    Ok(state.lock().value["s5"]["configurationCheckTask"].clone())
}

fn next_s5_task_id() -> String {
    use std::sync::atomic::{AtomicU64, Ordering};
    static SEQUENCE: AtomicU64 = AtomicU64::new(1);
    format!(
        "s5-{}-{}",
        chrono::Local::now().timestamp_millis(),
        SEQUENCE.fetch_add(1, Ordering::Relaxed)
    )
}

/// 系统原生文件/目录选择的公共入口（`pick_paths` API）。
/// 不修改状态、不受 busy 限制；返回 {canceled, paths, backend, message}。
/// 原生能力不可用/超时时 `message` 带可读诊断，`canceled=false, paths=[]`，
/// 由前端回退页面内选择器；用户取消时 `canceled=true`，前端不再回退。
/// 选中的路径不在此校验——仍由各业务的 choose_* 接口按既有规则校验。
async fn pick_paths(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let options = args
        .first()
        .cloned()
        .unwrap_or_else(|| Value::Object(serde_json::Map::new()));
    let title = options
        .get("title")
        .and_then(Value::as_str)
        .unwrap_or("选择位置")
        .trim()
        .to_owned();
    let kind = match options.get("select").and_then(Value::as_str) {
        Some("dir") => PickKind::Directory,
        Some("multi") => PickKind::MultiFile,
        _ => PickKind::File,
    };
    let mode = options.get("mode").and_then(Value::as_str).unwrap_or("");
    let (start, start_notice) =
        crate::native_file_dialog::normalize_start(options.get("start").and_then(Value::as_str));
    let request = PickRequest {
        title: if title.is_empty() { "选择位置".to_owned() } else { title },
        kind,
        extensions: browse_extensions(mode),
        start,
    };
    let outcome = crate::native_file_dialog::pick(request).await;
    // 克隆后合并：outcome 仍需整体借用序列化；起始路径提示与诊断可并存。
    let message = outcome.message.clone().or(start_notice);
    if let Some(text) = &message {
        // 诊断写入运行日志但不改变业务状态，便于事后定位原生能力问题。
        let mut ui = state.lock();
        push_log(&mut ui.value, format!("文件选择：{text}"));
    }
    let mut value = crate::native_file_dialog::outcome_to_value(&outcome);
    if let Some(text) = message {
        value["message"] = json!(text);
    }
    Ok(value)
}

fn choose_folder(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_folder 缺少路径类型"))?;
    if !matches!(field, "input" | "templateDir" | "output") {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持选择目录：{field}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(state.lock().value[field].clone());
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选目录不存在：{}", path.display()),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析所选目录失败：{error}"),
        )
    })?;
    Ok(json!(path.to_string_lossy()))
}

fn choose_history_config(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let selected = arg_string(args, 0).unwrap_or_default().trim();
    if selected.is_empty() {
        let current = state.lock().value["historyConfig"]
            .as_str()
            .unwrap_or_default()
            .to_owned();
        if current.is_empty() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "请先选择逐笔统计系统配置工作簿",
            ));
        }
        return Ok(json!(current));
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_file()
        || !path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "逐笔统计系统配置必须是存在的 .xlsx 文件：{}",
                path.display()
            ),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析逐笔统计系统配置失败：{error}"),
        )
    })?;
    let value = path.to_string_lossy().into_owned();
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改逐笔统计系统配置",
        ));
    }
    ui.value["historyConfig"] = json!(value);
    push_log(&mut ui.value, "已选择逐笔统计系统历史审核配置".to_owned());
    drop(ui);
    persist_state_preferences(state)?;
    Ok(json!(value))
}

fn recognize_s1_sources(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (input_dir, template_dir, manual_template, busy) = {
        let ui = state.lock();
        (
            ui.value["input"].as_str().unwrap_or_default().to_owned(),
            ui.value["templateDir"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
            ui.value["templateManual"] == true,
            ui.value["busy"] == true,
        )
    };
    if busy {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能识别模板",
        ));
    }
    let input_dir = PathBuf::from(input_dir);
    if !input_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "请先选择有效的源数据目录",
        ));
    }
    let files = s1_region_summary_source_workbooks(&input_dir, true)
        .map_err(|error| ApiError::internal(format!("扫描 S1 源数据目录失败：{error:#}")))?;
    let template_dir = PathBuf::from(template_dir);
    let can_match_templates = template_dir.is_dir();
    let template_profiles = if can_match_templates {
        Some(
            load_template_profiles(&template_dir)
                .map_err(|error| ApiError::internal(format!("识别 S1 模板失败：{error:#}")))?,
        )
    } else {
        None
    };
    let mut source_rows = Vec::with_capacity(files.len());
    let mut matched_templates = std::collections::BTreeSet::new();
    for path in &files {
        let suggestion = if can_match_templates {
            recommend_template_with_profiles(
                template_profiles
                    .as_deref()
                    .expect("matching requires loaded template profiles"),
                std::slice::from_ref(path),
            )
        } else {
            crate::s1_template::TemplateSuggestion {
                template_path: None,
                confidence: 0.0,
                matched: false,
                details: "请先选择有效的模板文件目录".to_owned(),
                alternatives: Vec::new(),
            }
        };
        let matched_template = suggestion
            .matched
            .then(|| suggestion.template_path.clone())
            .flatten();
        let template_label = matched_template
            .as_ref()
            .and_then(|path| path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| "未识别".to_owned());
        if let Some(path) = &matched_template {
            matched_templates.insert(path.clone());
        }
        source_rows.push(json!({
            "name": path.file_name().unwrap_or_default().to_string_lossy(),
            "path": path,
            "template": template_label,
            "templateMatched": suggestion.matched,
            "templateConfidence": suggestion.confidence,
            "templateReason": suggestion.details
        }));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核已开始，不能更新待处理文件列表",
        ));
    }
    ui.value["sourceFiles"] = json!(source_rows);
    ui.value["selectedFiles"] = json!(files);
    ui.value["mixedTemplates"] = json!(matched_templates
        .iter()
        .filter_map(|path| path.file_name())
        .map(|name| name.to_string_lossy().into_owned())
        .collect::<Vec<_>>());
    if !manual_template && matched_templates.len() == 1 {
        if let Some(template) = matched_templates.iter().next() {
            ui.value["template"] = json!(template.to_string_lossy());
        }
    }
    ui.value["status"] = json!(format!("已扫描 {} 个 S1 源工作簿", files.len()));
    push_log(
        &mut ui.value,
        format!(
            "S1 源文件扫描完成：{} 个工作簿；列表包含模板匹配结果和未匹配原因",
            files.len()
        ),
    );
    Ok(ui.value.clone())
}

fn set_selected_files(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(paths) = args.first().and_then(Value::as_array) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "set_selected_files 需要路径数组",
        ));
    };
    let selected = paths
        .iter()
        .map(|value| {
            value
                .as_str()
                .map(str::to_owned)
                .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "已选源文件路径必须是字符串"))
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改待处理文件",
        ));
    }
    let available = ui.value["sourceFiles"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|file| file["path"].as_str())
        .collect::<std::collections::HashSet<_>>();
    if let Some(path) = selected
        .iter()
        .find(|path| !available.contains(path.as_str()))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("待处理文件不在当前扫描清单中：{path}"),
        ));
    }
    ui.value["selectedFiles"] = json!(selected);
    Ok(json!(true))
}

fn choose_file(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_file 缺少路径类型"))?;
    if !matches!(field, "template" | "external") {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持选择文件：{field}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(json!(""));
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_file() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选工作簿不存在：{}", path.display()),
        ));
    }
    if !path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
        })
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            if field == "template" {
                "S1 模板目前只支持 .xlsx 和 .xlsm"
            } else {
                "S1 外部数据目前只支持 .xlsx 和 .xlsm"
            },
        ));
    }

    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能更改路径",
        ));
    }
    let value = path.to_string_lossy().into_owned();
    ui.value[field] = json!(value);
    if field == "template" {
        ui.value["templateManual"] = json!(true);
        push_log(&mut ui.value, "S1 模板：已选择待检查模板".to_owned());
    } else {
        push_log(
            &mut ui.value,
            "S1 外部数据：已选择工作簿；F1 当前预检不会修改或导入该文件".to_owned(),
        );
    }
    Ok(ui.value[field].clone())
}

#[derive(Clone)]
struct S1F1PreflightRequest {
    source_files: Vec<PathBuf>,
    source_directory: PathBuf,
    template_dir: PathBuf,
    manual_template: Option<PathBuf>,
    mapping_config: Option<PathBuf>,
    external_data: Option<PathBuf>,
    output_directory: PathBuf,
    batch_id: String,
    run_directory: PathBuf,
    planned_result_path: PathBuf,
    sidecar_path: Option<PathBuf>,
    calc_engine: String,
    edit_engine: String,
    request_preparation_us: u128,
    source_directory_scan_us: u128,
}

struct S1F1TemplatePreflight {
    structure: S1StructureTemplate,
    rule_plan: S1TemplateRulePlan,
    scan_areas: Vec<String>,
    structure_ranges: Vec<CopyRange>,
    scan_area_count: usize,
    structure_area_count: usize,
    health_passed: bool,
    health_errors: Vec<String>,
}

struct S1F1AcceptedSource {
    source: PathBuf,
    template: PathBuf,
    preflight: Arc<S1F1TemplatePreflight>,
}

struct S1F1OutputDirectoryGuard {
    path: PathBuf,
    keep: bool,
}

impl S1F1OutputDirectoryGuard {
    fn create(path: PathBuf) -> Result<Self> {
        fs::create_dir(&path)
            .with_context(|| format!("创建本次 S1-F01 审核目录 {}", path.display()))?;
        Ok(Self { path, keep: false })
    }

    fn keep(&mut self) {
        self.keep = true;
    }
}

impl Drop for S1F1OutputDirectoryGuard {
    fn drop(&mut self) {
        if !self.keep {
            let _ = fs::remove_dir_all(&self.path);
        }
    }
}

/// Execute S1-F01 through the same request and worker path used by the GUI.
/// This adapter is consumed by the acceptance runner so it can exercise the
/// complete template-match, structure-preflight, formula-copy, recalculation,
/// issue-merge and result-publication flow without duplicating business logic.
pub fn run_s1_f1_for_acceptance(
    source_directory: &Path,
    template_directory: &Path,
    mapping_config: Option<&Path>,
    external_data: Option<&Path>,
    output_directory: &Path,
    sidecar: Option<&Path>,
) -> Result<Value> {
    let state_config = mapping_config
        .map(Path::to_path_buf)
        .unwrap_or_else(|| output_directory.join(".acceptance-config-placeholder.xlsx"));
    let state = AppState::new(state_config);
    let arguments = json!({
        "input": source_directory,
        "templateDir": template_directory,
        "mappingConfig": mapping_config,
        "external": external_data,
        "output": output_directory,
        "sidecar": sidecar,
    });
    start_dag_flow(&state, &[json!("dag:汇总核查表校验"), arguments, json!([])])
        .map_err(|error| anyhow::anyhow!(error.message))?;

    while state.get_state()["busy"] == true {
        std::thread::sleep(std::time::Duration::from_millis(100));
    }

    let snapshot = state.get_state();
    Ok(json!({
        "status": snapshot["s1F1Preflight"]["status"],
        "result_path": snapshot["s1F1OutputPath"],
        "preflight": snapshot["s1F1Preflight"],
        "logs": snapshot["log"],
        "ui_status": snapshot["status"],
    }))
}

fn start_s1_f1_workflow(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let request_preparation_started = Instant::now();
    let values = args.get(1).and_then(Value::as_object).ok_or_else(|| {
        ApiError::new(StatusCode::BAD_REQUEST, "启动 S1-F01 需要当前 GUI 路径设置")
    })?;
    let input_dir = values
        .get("input")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择 S1-F01 源数据目录"))?;
    if !input_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F01 源数据目录不存在：{}", input_dir.display()),
        ));
    }
    let input_dir = fs::canonicalize(&input_dir).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F01 源目录失败：{error}"),
        )
    })?;
    let source_directory_scan_started = Instant::now();
    let source_files = s1_region_summary_source_workbooks(&input_dir, true)
        .map_err(|error| ApiError::internal(format!("扫描 S1-F01 源文件失败：{error:#}")))?;
    let source_directory_scan_us = source_directory_scan_started.elapsed().as_micros();
    if source_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "源数据目录中没有可审核的 .xlsx 工作簿",
        ));
    }

    let requested_selection = args
        .get(2)
        .and_then(Value::as_array)
        .map(|paths| {
            paths
                .iter()
                .map(|value| {
                    value.as_str().map(PathBuf::from).ok_or_else(|| {
                        ApiError::new(StatusCode::BAD_REQUEST, "已选源文件路径必须是字符串")
                    })
                })
                .collect::<std::result::Result<Vec<_>, _>>()
        })
        .transpose()?
        .unwrap_or_default();
    let selected_files = if requested_selection.is_empty() {
        source_files
    } else {
        let eligible = source_files
            .iter()
            .map(|path| web_path_identity(path))
            .collect::<std::collections::HashSet<_>>();
        let mut selected = Vec::new();
        for path in requested_selection {
            let canonical = fs::canonicalize(&path).map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("已选 S1-F01 源文件不可用：{}：{error}", path.display()),
                )
            })?;
            if !eligible.contains(&web_path_identity(&canonical)) {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("已选文件不在当前 S1-F01 源目录中：{}", path.display()),
                ));
            }
            selected.push(canonical);
        }
        selected.sort();
        selected.dedup_by(|left, right| web_path_identity(left) == web_path_identity(right));
        selected
    };
    if selected_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "没有选中 S1-F01 源文件",
        ));
    }

    let template_dir = values
        .get("templateDir")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择 S1-F01 模板目录"))?;
    if !template_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F01 模板目录不存在：{}", template_dir.display()),
        ));
    }
    let template_dir = fs::canonicalize(&template_dir).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F01 模板目录失败：{error}"),
        )
    })?;
    let (manual_template, state_mapping_config) = {
        let ui = state.lock();
        let manual = (ui.value["templateManual"] == true)
            .then(|| values.get("template").and_then(Value::as_str))
            .flatten()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from);
        let config = ui.value["historyConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from);
        (manual, config)
    };
    let mapping_config = values
        .get("mappingConfig")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .or(state_mapping_config);
    if let Some(template) = manual_template.as_ref() {
        if !template.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("手动选择的 S1-F01 模板不存在：{}", template.display()),
            ));
        }
    }
    if let Some(config) = mapping_config.as_ref() {
        if !config.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("S1-F01 配置工作簿不存在：{}", config.display()),
            ));
        }
    }
    let external_data = values
        .get("external")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from);
    if let Some(path) = external_data.as_ref() {
        if !path.is_file()
            || !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                })
        {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!(
                    "S1-F01 外部数据必须是存在的 .xlsx 或 .xlsm 文件：{}",
                    path.display()
                ),
            ));
        }
    }

    let output_directory = values
        .get("output")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| input_dir.join("执行结果"));
    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F01 输出位置不是目录：{}", output_directory.display()),
        ));
    }
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S_%3f").to_string();
    let batch_id = format!("S1F1-{stamp}");
    let run_directory = output_directory.join(format!("S1-F1_审核结果_{stamp}"));
    let planned_result_path = run_directory.join("汇总核查表校验.xlsx");
    let sidecar_path = if cfg!(windows) {
        None
    } else {
        values
            .get("sidecar")
            .and_then(Value::as_str)
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(find_s1_recalculation_sidecar)
    };
    let settings = crate::app_settings::AppSettings::load();
    {
        let mut ui = state.lock();
        push_log(
            &mut ui.value,
            format!(
                "S1-F01 高级设置：计算引擎={}；修改引擎={}（设置文件 {}）",
                settings.calc_engine.label(),
                settings.edit_engine.label(),
                crate::app_settings::AppSettings::settings_path_for_display()
            ),
        );
    }
    let request_preparation_us = request_preparation_started.elapsed().as_micros();
    let request = S1F1PreflightRequest {
        source_files: selected_files.clone(),
        source_directory: input_dir,
        template_dir,
        manual_template,
        mapping_config,
        external_data,
        output_directory,
        batch_id,
        run_directory,
        planned_result_path,
        sidecar_path,
        calc_engine: settings.calc_engine.sidecar_engine().to_owned(),
        edit_engine: settings.edit_engine.as_str().to_owned(),
        request_preparation_us,
        source_directory_scan_us,
    };

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行 S1-F01",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在识别 S1-F01 模板并检查表结构……");
        ui.value["s1F1OutputPath"] = json!("");
        ui.value["s1F1Preflight"] = json!({});
        push_log(
            &mut ui.value,
            format!("S1-F01 开始预检：已选 {} 个源工作簿", selected_files.len()),
        );
        push_detail_log(
            &mut ui.value,
            format!(
                "S1-F01 输入：{}；模板目录：{}；配置：{}；外部数据：{}；预期结果路径：{}",
                selected_files.len(),
                request.template_dir.display(),
                request
                    .mapping_config
                    .as_deref()
                    .map(|path| path.display().to_string())
                    .unwrap_or_else(|| "内置映射".to_owned()),
                request
                    .external_data
                    .as_deref()
                    .map(|path| path.display().to_string())
                    .unwrap_or_else(|| "未选择".to_owned()),
                request.planned_result_path.display()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s1-f1-workflow".to_owned())
        .spawn(move || run_s1_f1_worker(worker_state, request))
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["s1F1OutputPath"] = json!("");
        ui.value["status"] = json!("S1-F01 启动失败");
        push_log(&mut ui.value, format!("启动 S1-F01 失败：{error}"));
        return Err(ApiError::internal(format!("启动 S1-F01 失败：{error}")));
    }
    Ok(json!({"accepted": true}))
}

fn list_s1_f3_tasks(state: &AppState) -> std::result::Result<Value, ApiError> {
    let snapshot = state.get_state();
    if !snapshot["s1F3RecoveryError"]
        .as_str()
        .unwrap_or_default()
        .is_empty()
    {
        return Err(ApiError::internal(format!(
            "读取 S1-F3 持久化任务失败：{}",
            snapshot["s1F3RecoveryError"].as_str().unwrap_or_default()
        )));
    }
    Ok(snapshot["s1F3RecoveryTasks"].clone())
}

fn start_s1_f3_workflow(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let run_id = arg_string(args, 0)
        .filter(|value| !value.trim().is_empty())
        .map(str::to_owned)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "S1-F3 缺少恢复 run_id"))?;
    let file_id = arg_string(args, 1)
        .filter(|value| !value.trim().is_empty())
        .map(str::to_owned)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "S1-F3 缺少恢复 file_id"))?;
    let store = S1RecoveryStore::user_default()
        .map_err(|error| ApiError::internal(format!("初始化 S1-F3 恢复任务失败：{error:#}")))?;
    let pending = store
        .pending_tasks()
        .map_err(|error| ApiError::internal(format!("扫描 S1-F3 恢复任务失败：{error:#}")))?;
    if !pending
        .iter()
        .any(|task| task.run_id == run_id && task.file_id == file_id)
    {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "所选 S1-F3 任务已完成、副本不可用或不再等待恢复；请刷新任务列表",
        ));
    }
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再使用 S1-F3",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在验证 S1-F3 人工重算副本并恢复审核……");
        push_log(
            &mut ui.value,
            format!("S1-F3 开始恢复：run_id={run_id}，file_id={file_id}"),
        );
    }
    let worker_state = state.clone();
    let worker_run_id = run_id.clone();
    let worker_file_id = file_id.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s1-f3-recovery".to_owned())
        .spawn(move || run_s1_f3_worker(worker_state, worker_run_id, worker_file_id))
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S1-F3 启动失败");
        push_log(&mut ui.value, format!("启动 S1-F3 失败：{error}"));
        return Err(ApiError::internal(format!("启动 S1-F3 失败：{error}")));
    }
    Ok(json!({"accepted": true}))
}

fn s1_recovery_result_path(manifest: &S1RecoveryManifest, input_sha256: &str) -> PathBuf {
    let run_hash = format!("{:x}", Sha256::digest(manifest.run_id.as_bytes()));
    let run_key = &run_hash[..32];
    let input_key = input_sha256.get(..32).unwrap_or(input_sha256);
    manifest
        .run_directory
        .join(format!("汇总核查表校验_恢复_{run_key}_{input_key}.xlsx"))
}

fn run_s1_f3_worker(state: AppState, run_id: String, file_id: String) {
    run_s1_f3_worker_with_store(state, run_id, file_id, S1RecoveryStore::user_default());
}

fn run_s1_f3_worker_with_store(
    state: AppState,
    run_id: String,
    file_id: String,
    store_result: Result<S1RecoveryStore>,
) {
    let result = (|| -> Result<(S1RecoveryManifest, PathBuf, S1ResultWriteSummary)> {
        let store = store_result
            .as_ref()
            .map_err(|error| anyhow::anyhow!("初始化 S1-F3 恢复状态目录：{error:#}"))?;
        let mut manifest = store.load(&run_id)?;
        let file_index = manifest
            .files
            .iter()
            .position(|file| file.file_id == file_id)
            .context("S1-F3 恢复清单中不存在所选工作簿")?;
        ensure!(
            manifest.files[file_index].status.is_pending(),
            "所选 S1-F3 工作簿已不处于待恢复状态"
        );
        manifest.files[file_index].status = S1RecoveryStatus::ReadyToResume;
        manifest.files[file_index].reason = Some("正在验证人工重算后的审核副本".to_owned());
        store.save(&mut manifest)?;

        let audit_copy = manifest.files[file_index]
            .audit_copy
            .clone()
            .context("S1-F3 任务缺少审核副本路径")?;
        ensure!(
            audit_copy.is_file(),
            "S1-F3 审核副本不存在：{}",
            audit_copy.display()
        );
        let template = &manifest.files[file_index].template_file;
        ensure!(
            template.is_file(),
            "S1-F3 模板文件不存在：{}",
            template.display()
        );
        ensure!(
            file_sha256(template)? == manifest.files[file_index].template_sha256,
            "S1-F3 模板内容已变化；为避免按新规则审核旧副本，任务保持待恢复"
        );
        if let Some(config) = manifest.mapping_config.as_deref() {
            ensure!(
                config.is_file(),
                "S1-F3 映射配置不存在：{}",
                config.display()
            );
            ensure!(
                manifest.mapping_config_sha256.as_deref() == Some(file_sha256(config)?.as_str()),
                "S1-F3 映射配置内容已变化；为避免按新配置审核旧副本，任务保持待恢复"
            );
        }

        let (
            current_fingerprint,
            current_normalized_fingerprint,
            current_semantic_fingerprint,
            formula_count,
        ) = workbook_recovery_fingerprints(&audit_copy)?;
        ensure!(formula_count > 0, "S1-F3 副本中没有可验证的公式");
        let expected_fingerprint = manifest.files[file_index]
            .formula_fingerprint
            .as_deref()
            .context("S1-F3 恢复任务缺少原始公式指纹，无法验证公式完整性")?;
        let expected_normalized_fingerprint = manifest.files[file_index]
            .formula_fingerprint_normalized
            .as_deref();
        validate_recovery_fingerprints(
            expected_fingerprint,
            expected_normalized_fingerprint,
            manifest.files[file_index]
                .audit_copy_semantic_fingerprint
                .as_deref(),
            &current_fingerprint,
            &current_normalized_fingerprint,
            &current_semantic_fingerprint,
        )?;
        let cache_inventory = excel_reader::inspect_formula_cache_nodes(&audit_copy)
            .context("检查 S1-F3 副本公式缓存节点")?;
        ensure!(
            cache_inventory.formula_cell_count == formula_count,
            "S1-F3 副本公式数量与 OOXML 公式缓存清单不一致"
        );
        if !cache_inventory.missing_cache_cells.is_empty() {
            manifest.files[file_index].status = S1RecoveryStatus::WaitingManualRecalc;
            manifest.files[file_index].reason = Some(format!(
                "公式缓存仍缺失 {} 个，示例：{}!{}；请重新计算、保存并关闭副本后重试。",
                cache_inventory.missing_cache_cells.len(),
                cache_inventory.missing_cache_cells[0].sheet_name,
                cache_inventory.missing_cache_cells[0].cell
            ));
            store.save(&mut manifest)?;
            bail!(
                "{}",
                manifest.files[file_index]
                    .reason
                    .as_deref()
                    .unwrap_or_default()
            );
        }
        let recalculated_workbook =
            excel_reader::read_workbook_with_unknown_error_tokens_as_text(&audit_copy)
                .context("读取 S1-F3 人工重算副本的公式结果")?
                .workbook;
        let result_coverage = excel_reader::formula_result_value_coverage(&recalculated_workbook);
        if result_coverage.formula_cell_count != formula_count
            || result_coverage.typed_result_cell_count != formula_count
            || !result_coverage.missing_result_cells.is_empty()
        {
            let example = result_coverage
                .missing_result_cells
                .first()
                .map(|location| format!("{}!{}", location.sheet_name, location.cell))
                .unwrap_or_else(|| {
                    format!(
                        "公式 {} 个，可读结果 {} 个",
                        result_coverage.formula_cell_count, result_coverage.typed_result_cell_count
                    )
                });
            let reason = format!(
                "公式缓存节点存在，但仍有公式结果无法读取（{}）；请在 Excel 中强制完整重算、保存并关闭副本后重试。",
                example
            );
            manifest.files[file_index].status = S1RecoveryStatus::WaitingManualRecalc;
            manifest.files[file_index].reason = Some(reason.clone());
            store.save(&mut manifest)?;
            bail!("{reason}");
        }

        let config = manifest.mapping_config.as_deref();
        let preflight =
            s1_f1_template_preflight(template, config).context("重新载入原 S1-F1 模板审核规则")?;
        ensure!(
            preflight.health_passed,
            "S1-F3 原模板体检未通过，不能恢复审核"
        );
        let item = S1F1AcceptedSource {
            source: manifest.files[file_index].source_file.clone(),
            template: template.clone(),
            preflight: Arc::new(preflight),
        };
        let org = organisation_from_name(&item.source);
        let metadata = S1FormulaMetadata {
            period: manifest.period.clone(),
            batch_id: manifest.batch_id.clone(),
            audit_time: manifest.audit_time.clone(),
            org_code: org.0,
            org_name: org.1,
        };
        let extracted = extract_s1_audit_copy(&item, &item.source, &audit_copy, &metadata)
            .context("复用 S1-F1 公式和条件格式审核提取")?;
        manifest.files[file_index].formula_issues = extracted.formula_issues;
        manifest.files[file_index].conditional_issues = extracted.conditional_issues;
        manifest.files[file_index].conditional_diagnostics = extracted.diagnostics;
        manifest.files[file_index].issues_extracted = true;
        manifest.files[file_index].audit_partial = extracted.partial;
        manifest.files[file_index].recalculation_engine = Some("manual_recalculation".to_owned());
        manifest.files[file_index].reason = extracted.reason;
        manifest.files[file_index].status = S1RecoveryStatus::ReadyToResume;
        store.save(&mut manifest)?;

        let summary = match write_s1_recovery_manifest_result(&store, &mut manifest) {
            Ok(summary) => summary,
            Err(error) => {
                manifest.files[file_index].status = S1RecoveryStatus::ResumeFailed;
                manifest.files[file_index].reason = Some(format!(
                    "审核提取已保存，但结果工作簿生成失败；可以重试 S1-F3：{error:#}"
                ));
                store.save(&mut manifest)?;
                return Err(error).context("S1-F3 写入恢复后的最终结果工作簿");
            }
        };
        manifest.result_path = Some(summary.path.clone());
        for file in &mut manifest.files {
            if file.issues_extracted
                && matches!(
                    file.status,
                    S1RecoveryStatus::ReadyToResume | S1RecoveryStatus::ResumeFailed
                )
            {
                file.status = if file.audit_partial {
                    S1RecoveryStatus::Partial
                } else {
                    S1RecoveryStatus::Completed
                };
            }
        }
        store.save(&mut manifest)?;
        Ok((manifest, summary.path.clone(), summary))
    })();

    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok((manifest, result_path, summary)) => {
            ui.value["s1F1OutputPath"] = json!(result_path.to_string_lossy());
            ui.value["status"] = json!(format!(
                "S1-F3 恢复完成：结果已保存到 {}；剩余待处理 {} 个",
                result_path.display(),
                manifest
                    .files
                    .iter()
                    .filter(|file| file.status.is_pending())
                    .count()
            ));
            ui.value["s1F1Preflight"] = json!({
                "status": "recovered",
                "runId": manifest.run_id,
                "resultPath": result_path,
                "formulaIssues": summary.formula_issue_count,
                "conditionalIssues": summary.conditional_issue_count,
                "historyRows": summary.history_rows,
                "totalIssues": summary.total_issue_count,
                "manualRecalculation": true,
                "cacheValidation": "formula text unchanged except for equivalent LibreOffice TRUE/FALSE literal serialization; all formula cache nodes are present and each result is readable by Calamine; result freshness depends on the selected calculation engine recalculating and saving the workbook",
                "auditInputValidation": if manifest.files.iter().any(|file| file.file_id == file_id && file.audit_copy_semantic_fingerprint.is_some()) { "non-formula cell values, conditional formatting, merged ranges and defined names match the pre-recalculation snapshot; LibreOffice style IDs/default serialization are normalized" } else { "legacy recovery manifest: non-formula input snapshot unavailable; formula integrity was checked" }
            });
            push_log(
                &mut ui.value,
                format!("S1-F3 结果文件：{}", result_path.display()),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F3 使用已保存的审核副本恢复：公式异常={}；条件格式异常={}；输出={}。未重新读取原始报表、匹配模板、复制公式或调用自动重算引擎。",
                    summary.formula_issue_count,
                    summary.conditional_issue_count,
                    result_path.display()
                ),
            );
        }
        Err(error) => {
            let reason = format!("{error:#}");
            if let Ok(store) = store_result.as_ref() {
                if let Ok(mut manifest) = store.load(&run_id) {
                    if let Some(file) = manifest
                        .files
                        .iter_mut()
                        .find(|file| file.file_id == file_id)
                    {
                        if file.status == S1RecoveryStatus::ReadyToResume {
                            file.status = S1RecoveryStatus::ResumeFailed;
                        }
                        if file.reason.as_deref() == Some("正在验证人工重算后的审核副本")
                        {
                            file.reason = Some(reason.clone());
                        }
                        let _ = store.save(&mut manifest);
                    }
                }
            }
            ui.value["status"] = json!(format!("S1-F3 恢复未完成：{reason}"));
            push_log(
                &mut ui.value,
                format!("S1-F3 恢复失败，可在修复后重试：{reason}"),
            );
        }
    }
}

fn run_s1_f1_worker(state: AppState, request: S1F1PreflightRequest) {
    let preflight_started = Instant::now();
    let mut files = Vec::with_capacity(request.source_files.len());
    let mut accepted = Vec::new();
    let template_profile_started = Instant::now();
    let template_profiles = if request.manual_template.is_none() && !request.source_files.is_empty()
    {
        Some(load_template_profiles(&request.template_dir))
    } else {
        None
    };
    let template_profile_load_us = template_profile_started.elapsed().as_micros();
    let mut template_matching_us = 0u128;
    let mut template_preflight_us = 0u128;
    let mut source_structure_compare_us = 0u128;
    let mut template_preflight_cache = std::collections::HashMap::<
        PathBuf,
        std::result::Result<Arc<S1F1TemplatePreflight>, String>,
    >::new();
    let mut structure_passed = 0usize;
    let mut skipped = 0usize;
    for source in &request.source_files {
        let template_match_started = Instant::now();
        let source_name = source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| source.display().to_string());
        let template_result = if let Some(template) = request.manual_template.as_ref() {
            Ok((template.clone(), "使用手动选择模板".to_owned()))
        } else {
            match template_profiles
                .as_ref()
                .expect("automatic matching loads profiles for non-empty source lists")
            {
                Ok(profiles) => {
                    let suggestion =
                        recommend_template_with_profiles(profiles, std::slice::from_ref(source));
                    let details = suggestion.details;
                    match suggestion.template_path.filter(|_| suggestion.matched) {
                        Some(path) => Ok((path, details)),
                        None => Err(anyhow::anyhow!(details)),
                    }
                }
                Err(error) => Err(anyhow::anyhow!("{error:#}")),
            }
        };
        template_matching_us += template_match_started.elapsed().as_micros();
        let file_report = match template_result {
            Err(error) => {
                skipped += 1;
                json!({
                    "sourceFile": source_name,
                    "status": "skipped_template_mismatch",
                    "reason": format!("模板未匹配：{error:#}"),
                    "template": Value::Null
                })
            }
            Ok((template_path, match_details)) => {
                let template_path = match fs::canonicalize(&template_path) {
                    Ok(path) => path,
                    Err(error) => {
                        skipped += 1;
                        files.push(json!({
                            "sourceFile": source_name,
                            "status": "skipped_template_unavailable",
                            "template": template_path,
                            "reason": format!("无法读取匹配模板：{error}")
                        }));
                        continue;
                    }
                };
                let template_preflight_started = Instant::now();
                let preflight_result = cached_s1_f1_template_preflight(
                    &mut template_preflight_cache,
                    &template_path,
                    request.mapping_config.as_deref(),
                );
                template_preflight_us += template_preflight_started.elapsed().as_micros();
                match preflight_result {
                    Err(error) => {
                        skipped += 1;
                        json!({
                            "sourceFile": source_name,
                            "status": "skipped_template_invalid",
                            "template": template_path,
                            "matchDetails": match_details,
                            "reason": format!("模板规则或表结构配置不可用：{error:#}")
                        })
                    }
                    Ok(preflight) if !preflight.health_passed => {
                        skipped += 1;
                        json!({
                            "sourceFile": source_name,
                            "status": "skipped_template_health_failed",
                            "template": template_path,
                            "matchDetails": match_details,
                            "scanAreas": preflight.scan_area_count,
                            "structureAreas": preflight.structure_area_count,
                            "reason": format!("模板体检未通过：{}", preflight.health_errors.join("；"))
                        })
                    }
                    Ok(preflight) => {
                        let structure_compare_started = Instant::now();
                        let comparison = compare_s1_source_structure(source, &preflight.structure);
                        source_structure_compare_us +=
                            structure_compare_started.elapsed().as_micros();
                        match comparison {
                            Ok(comparison) if comparison.matched => {
                                structure_passed += 1;
                                let scan_area_count = preflight.scan_area_count;
                                let structure_area_count = preflight.structure_area_count;
                                accepted.push(S1F1AcceptedSource {
                                    source: source.clone(),
                                    template: template_path.clone(),
                                    preflight,
                                });
                                json!({
                                    "sourceFile": source_name,
                                    "sourcePath": source,
                                    "status": "structure_passed",
                                    "template": template_path,
                                    "matchDetails": match_details,
                                    "scanAreas": scan_area_count,
                                    "structureAreas": structure_area_count,
                                    "structureMatch": comparison
                                })
                            }
                            Ok(comparison) => {
                                skipped += 1;
                                let reason = if !comparison.missing_sheets.is_empty() {
                                    format!(
                                        "缺少模板要求的工作表：{}",
                                        comparison.missing_sheets.join("、")
                                    )
                                } else {
                                    comparison.diagnostic.clone().unwrap_or_else(|| {
                                        format!(
                                            "固定标签匹配率 {:.0}% ({}/{})",
                                            comparison.score * 100.0,
                                            comparison.matched_labels,
                                            comparison.checked_labels
                                        )
                                    })
                                };
                                json!({
                                    "sourceFile": source_name,
                                    "status": "skipped_structure_mismatch",
                                    "template": template_path,
                                    "matchDetails": match_details,
                                    "structureMatch": comparison,
                                    "reason": format!("表结构不匹配，已跳过：{reason}")
                                })
                            }
                            Err(error) => {
                                skipped += 1;
                                json!({
                                    "sourceFile": source_name,
                                    "status": "skipped_structure_read_error",
                                    "template": template_path,
                                    "matchDetails": match_details,
                                    "reason": format!("无法读取源文件结构，已跳过：{error:#}")
                                })
                            }
                        }
                    }
                }
            }
        };
        files.push(file_report);
    }
    let preflight_timings = json!({
        "unit": "microseconds",
        "templateProfileLoadUs": template_profile_load_us,
        "templateMatchingUs": template_matching_us,
        "templatePreflightUs": template_preflight_us,
        "sourceStructureCompareUs": source_structure_compare_us,
        "unclassifiedPreflightUs": preflight_started.elapsed().as_micros()
            .saturating_sub(template_profile_load_us)
            .saturating_sub(template_matching_us)
            .saturating_sub(template_preflight_us)
            .saturating_sub(source_structure_compare_us),
        "totalUs": preflight_started.elapsed().as_micros(),
        "sourceFiles": request.source_files.len(),
        "structurePassed": structure_passed,
        "skipped": skipped
    });
    let request_timings = json!({
        "unit": "microseconds",
        "sourceDirectoryScanUs": request.source_directory_scan_us,
        "requestPreparationOtherUs": request
            .request_preparation_us
            .saturating_sub(request.source_directory_scan_us),
        "requestPreparationTotalUs": request.request_preparation_us
    });

    if structure_passed == 0 {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["s1F1OutputPath"] = json!("");
        ui.value["s1F1Preflight"] = json!({
            "status": "no_structure_passed",
            "filesTotal": request.source_files.len(),
            "structurePassed": 0,
            "skipped": skipped,
            "files": files.clone(),
            "outputDirectory": request.output_directory,
            "plannedResultPath": request.planned_result_path,
            "resultPath": Value::Null,
            "externalData": request.external_data,
            "recalculationSidecarFound": request.sidecar_path.is_some(),
            "recalculationSidecar": request.sidecar_path,
            "performanceTimings": {"request": request_timings, "preflight": preflight_timings}
        });
        ui.value["status"] = json!(format!(
            "S1-F01 预检结束：结构通过 0 个、跳过 {} 个；未生成审核结果",
            skipped
        ));
        for file in files {
            let name = file["sourceFile"].as_str().unwrap_or("未知文件");
            let reason = file["reason"].as_str().unwrap_or("未提供原因");
            push_log(&mut ui.value, format!("{name}：{reason}"));
        }
        push_log(
            &mut ui.value,
            format!(
                "S1-F01 未生成审核结果；预期路径：{}",
                request.planned_result_path.display()
            ),
        );
        return;
    }

    {
        let mut ui = state.lock();
        ui.value["status"] = json!("模板与表结构预检通过，正在复制公式并执行正式审核……");
        push_log(
            &mut ui.value,
            format!(
                "S1-F01 模板与表结构预检通过：{} 个文件；开始执行正式审核",
                structure_passed
            ),
        );
    }

    match execute_s1_f1_workflow(&request, &accepted) {
        Ok(output) => {
            let performance_timings = json!({
                "request": request_timings,
                "preflight": preflight_timings,
                "workflow": output.performance_timings
            });
            for file in &mut files {
                if file["status"] == "structure_passed" {
                    let source_path = file["sourcePath"].as_str().unwrap_or_default();
                    let source_key = web_path_identity(Path::new(source_path));
                    if let Some(result) = output.file_results.iter().find(|result| {
                        result["sourceFile"]
                            .as_str()
                            .is_some_and(|path| web_path_identity(Path::new(path)) == source_key)
                    }) {
                        file["status"] = result["status"].clone();
                        file["reason"] = result["reason"].clone();
                        file["auditWorkbook"] = result["auditWorkbook"].clone();
                    }
                }
            }
            let mut ui = state.lock();
            ui.value["busy"] = json!(false);
            ui.value["s1F1OutputPath"] = json!(output
                .result_path
                .as_ref()
                .map(|path| path.to_string_lossy().into_owned())
                .unwrap_or_default());
            ui.value["s1F1Preflight"] = json!({
                "status": if output.waiting_manual_recalc > 0 || output.partial_files > 0 || output.failed_files > 0 { "partial" } else { "completed" },
                "filesTotal": request.source_files.len(),
                "structurePassed": structure_passed,
                "reviewedFiles": output.file_results.iter().filter(|file| matches!(file["status"].as_str(), Some("COMPLETED" | "PARTIAL"))).count(),
                "skipped": skipped,
                "files": files.clone(),
                "fileResults": output.file_results,
                "outputDirectory": request.output_directory,
                "runDirectory": request.run_directory,
                "plannedResultPath": request.planned_result_path,
                "resultPath": output.result_path,
                "resultError": output.result_error,
                "formulaIssues": output.formula_issue_count,
                "conditionalIssues": output.conditional_issue_count,
                "historyRows": output.history_rows,
                "totalIssues": output.total_issue_count,
                "recalculationEngines": output.recalculation_engines,
                "auditCopies": output.audit_copies,
                "waitingManualRecalc": output.waiting_manual_recalc,
                "partialFiles": output.partial_files,
                "failedFiles": output.failed_files,
                "manualRecalcInstructions": "自动公式重算不可用或失败，已保留待计算的审核副本。请使用 Microsoft Excel 打开审核副本：1. 按 Ctrl+Alt+F9 强制重新计算；2. 等待 Excel 完成计算；3. 按 Ctrl+S 保存；4. 关闭工作簿；5. 返回 Rust-Audit，使用 S1-F3『从审核副本汇总校验核查表』继续审核。如果使用 WPS 或 LibreOffice，请采用对应软件支持的重新计算方式并保存。请勿修改原始报表文件。",
                "externalData": request.external_data,
                "recalculationSidecarFound": request.sidecar_path.is_some(),
                "recalculationSidecar": request.sidecar_path,
                "performanceTimings": performance_timings
            });
            for file in &files {
                let name = file["sourceFile"].as_str().unwrap_or("未知文件");
                let status = file["status"].as_str().unwrap_or("未知状态");
                let reason = file["reason"].as_str().unwrap_or("");
                push_log(
                    &mut ui.value,
                    format!(
                        "{name}：{status}{}",
                        if reason.is_empty() {
                            String::new()
                        } else {
                            format!("；{reason}")
                        }
                    ),
                );
            }
            ui.value["status"] = json!(if output.waiting_manual_recalc > 0 {
                format!(
                    "S1-F01 部分完成：{} 个待人工重算、{} 个部分审核、{} 个失败；已生成结果 {}",
                    output.waiting_manual_recalc,
                    output.partial_files,
                    output.failed_files,
                    output
                        .result_path
                        .as_ref()
                        .map(|path| path.display().to_string())
                        .unwrap_or_else(|| "无可发布结果".to_owned())
                )
            } else if output.result_path.is_some() {
                format!(
                    "S1-F01 审核完成：公式异常 {} 条、条件格式异常 {} 条；结果已保存",
                    output.formula_issue_count, output.conditional_issue_count
                )
            } else {
                format!(
                    "S1-F01 未生成核查表：待恢复 {} 个、失败 {} 个；审核副本与状态已保留",
                    output.waiting_manual_recalc, output.failed_files
                )
            });
            if let Some(result_path) = output.result_path.as_ref() {
                push_log(
                    &mut ui.value,
                    format!("S1-F01 结果文件：{}", result_path.display()),
                );
            } else if output.waiting_manual_recalc > 0 {
                push_log(&mut ui.value, "自动公式重算不可用或失败。请按 Ctrl+Alt+F9、等待计算、Ctrl+S 保存并关闭审核副本，然后使用 S1-F3 恢复审核。请勿修改原始报表文件。".to_owned());
            }
            push_log(
                &mut ui.value,
                format!("S1-F01 审核副本目录：{}", request.run_directory.display()),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F01 正式审核结束：数据期={}；审核副本={}；公式异常={}；条件格式异常={}；读取历史行={}；本期合计={}；真实重算引擎={}；结果工作表={}；输出={}；待人工重算={}；部分完成={}；失败={}。原始源文件保持只读。",
                    output.period,
                    output.audit_copies.len(),
                    output.formula_issue_count,
                    output.conditional_issue_count,
                    output.history_rows,
                    output.total_issue_count,
                    output.recalculation_engines.join("、"),
                    output.result_sheet.as_deref().unwrap_or("未生成"),
                    output.result_path.as_ref().map(|path| path.display().to_string()).unwrap_or_else(|| "无".to_owned()),
                    output.waiting_manual_recalc,
                    output.partial_files,
                    output.failed_files
                ),
            );
        }
        Err(error) => {
            let raw_reason = format!("{error:#}");
            let reason = if raw_reason.contains("no qualified backend") {
                format!(
                    "Excel-Core 当前没有通过工作簿保留验证的 S1-F1 公式复制路由，已拒绝修改工作簿：{raw_reason}"
                )
            } else {
                raw_reason
            };
            for file in &mut files {
                if file["status"] == "structure_passed" {
                    file["status"] = json!("review_failed");
                    file["reason"] = json!(reason);
                }
            }
            let mut ui = state.lock();
            ui.value["busy"] = json!(false);
            ui.value["s1F1OutputPath"] = json!("");
            ui.value["s1F1Preflight"] = json!({
                "status": "failed",
                "filesTotal": request.source_files.len(),
                "structurePassed": structure_passed,
                "skipped": skipped,
                "files": files.clone(),
                "outputDirectory": request.output_directory,
                "runDirectory": request.run_directory,
                "plannedResultPath": request.planned_result_path,
                "resultPath": Value::Null,
                "externalData": request.external_data,
                "recalculationSidecarFound": request.sidecar_path.is_some(),
                "recalculationSidecar": request.sidecar_path,
                "performanceTimings": {
                    "request": request_timings,
                    "preflight": preflight_timings
                },
                "error": reason
            });
            for file in &files {
                let name = file["sourceFile"].as_str().unwrap_or("未知文件");
                let status = file["status"].as_str().unwrap_or("未知状态");
                if status == "review_failed" {
                    push_log(&mut ui.value, format!("{name}：{reason}"));
                } else {
                    push_log(
                        &mut ui.value,
                        format!(
                            "{name}：{}",
                            file["reason"].as_str().unwrap_or("未提供原因")
                        ),
                    );
                }
            }
            ui.value["status"] = json!("S1-F01 正式审核失败；未生成审核结果");
            push_log(
                &mut ui.value,
                format!(
                    "S1-F01 未生成审核结果；预期路径：{}",
                    request.planned_result_path.display()
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F01 正式审核失败；本次新建审核目录已清理；原始源文件保持只读；未发布部分结果。阻断原因：{reason}"
                ),
            );
        }
    }
}

struct S1F1WorkflowOutput {
    period: String,
    result_path: Option<PathBuf>,
    result_sheet: Option<String>,
    formula_issue_count: usize,
    conditional_issue_count: usize,
    history_rows: usize,
    total_issue_count: usize,
    audit_copies: Vec<Value>,
    recalculation_engines: Vec<String>,
    file_results: Vec<Value>,
    waiting_manual_recalc: usize,
    partial_files: usize,
    failed_files: usize,
    result_error: Option<String>,
    performance_timings: Value,
}

struct S1ExtractedAuditCopy {
    formula_issues: Vec<Issue>,
    conditional_issues: Vec<Issue>,
    diagnostics: Vec<Value>,
    partial: bool,
    reason: Option<String>,
    formula_extraction_elapsed_us: u128,
    conditional_format_extraction_elapsed_us: u128,
}

fn extract_s1_audit_copy(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    metadata: &S1FormulaMetadata,
) -> Result<S1ExtractedAuditCopy> {
    extract_s1_audit_copy_with_conditional_extractor(
        item,
        source_file,
        audit_copy,
        metadata,
        |workbook_path, source_file, scan_areas, structure_ranges, conditional_metadata| {
            extract_conditional_issues(
                workbook_path,
                source_file,
                scan_areas,
                structure_ranges,
                conditional_metadata,
            )
        },
    )
}

fn extract_s1_audit_copy_with_conditional_extractor<F>(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    metadata: &S1FormulaMetadata,
    conditional_extractor: F,
) -> Result<S1ExtractedAuditCopy>
where
    F: FnOnce(
        &Path,
        &Path,
        &[String],
        &[CopyRange],
        &S1ConditionalMetadata,
    ) -> Result<S1ConditionalAuditResult>,
{
    let formula_extraction_started = Instant::now();
    let formula_result = extract_s1_formula_issues(
        &item.preflight.rule_plan.definition,
        audit_copy,
        Some(source_file),
        metadata,
    )
    .with_context(|| format!("提取公式审核异常失败：{}", source_file.display()))?;

    let validation_plan =
        S1ValidationFormulaScanPlan::load(std::slice::from_ref(&item.template))
            .with_context(|| format!("读取校验区域公式失败：{}", item.template.display()))?;
    let validation_report = validation_plan
        .scan(audit_copy)
        .with_context(|| format!("扫描校验区域结果失败：{}", source_file.display()))?;
    let formula_extraction_elapsed_us = formula_extraction_started.elapsed().as_micros();
    let mut formula_issues = formula_result.issues;
    let validation_issues =
        issues_from_validation_formula_scan(&validation_report, source_file, audit_copy, metadata);
    for candidate in validation_issues {
        if !formula_issues.iter().any(|existing| {
            existing.issue_id == candidate.issue_id
                && existing.formula_result == candidate.formula_result
                && existing.source_file == candidate.source_file
        }) {
            formula_issues.push(candidate);
        }
    }

    let conditional_extraction_started = Instant::now();
    let conditional_result = conditional_extractor(
        audit_copy,
        source_file,
        &item.preflight.scan_areas,
        &item.preflight.structure_ranges,
        &S1ConditionalMetadata {
            period: metadata.period.clone(),
            batch_id: metadata.batch_id.clone(),
            audit_time: metadata.audit_time.clone(),
            org_code: metadata.org_code.clone(),
            org_name: metadata.org_name.clone(),
        },
    );
    let conditional_format_extraction_elapsed_us =
        conditional_extraction_started.elapsed().as_micros();

    let (
        conditional_issues,
        mut diagnostics,
        conditional_partial,
        conditional_diagnostic_count,
        unresolved_indicator_count,
        conditional_error,
    ) = match conditional_result {
        Ok(conditional_result) => (
            conditional_result.issues,
            conditional_result
                .evaluation
                .diagnostics
                .iter()
                .map(|diagnostic| json!({"kind":"conditional_format", "diagnostic":diagnostic}))
                .collect::<Vec<_>>(),
            conditional_result.status != "completed",
            conditional_result.evaluation.diagnostics.len(),
            conditional_result.context.unresolved_indicator_count,
            None,
        ),
        Err(error) => {
            let message = format!("{error:#}");
            (
                Vec::new(),
                vec![json!({
                    "kind":"conditional_format",
                    "stage":"audit_extraction",
                    "status":"error",
                    "message":message
                })],
                true,
                1,
                0,
                Some(message),
            )
        }
    };
    if !formula_result
        .summary
        .read_diagnostics
        .normalized_error_cells
        .is_empty()
    {
        diagnostics.push(json!({
            "kind":"formula_value_reader",
            "diagnostics":formula_result.summary.read_diagnostics
        }));
    }
    if !validation_report
        .read_diagnostics
        .normalized_error_cells
        .is_empty()
    {
        diagnostics.push(json!({
            "kind":"validation_value_reader",
            "diagnostics":validation_report.read_diagnostics
        }));
    }
    let partial = conditional_partial
        || !formula_result
            .summary
            .read_diagnostics
            .normalized_error_cells
            .is_empty()
        || !validation_report
            .read_diagnostics
            .normalized_error_cells
            .is_empty();
    let reason = partial.then(|| {
        let conditional_failure = conditional_error
            .as_ref()
            .map(|error| format!("；条件格式阶段失败，已保留公式审核结果：{error}"))
            .unwrap_or_default();
        format!(
            "审核结果含部分诊断：条件格式诊断 {} 条、未识别指标 {} 个、未知公式错误值 {} 个{}",
            conditional_diagnostic_count,
            unresolved_indicator_count,
            formula_result
                .summary
                .read_diagnostics
                .normalized_error_cells
                .len()
                + validation_report
                    .read_diagnostics
                    .normalized_error_cells
                    .len(),
            conditional_failure
        )
    });

    Ok(S1ExtractedAuditCopy {
        formula_issues,
        conditional_issues,
        diagnostics,
        partial,
        reason,
        formula_extraction_elapsed_us,
        conditional_format_extraction_elapsed_us,
    })
}

struct S1ResultWriteSummary {
    path: PathBuf,
    sheet: String,
    formula_issue_count: usize,
    conditional_issue_count: usize,
    history_rows: usize,
    total_issue_count: usize,
}

struct PreparedS1ManifestResult {
    issues: Vec<Issue>,
    formula_issue_count: usize,
    conditional_issue_count: usize,
    history_rows: usize,
    total_issue_count: usize,
}

fn prepare_s1_manifest_result(manifest: &S1RecoveryManifest) -> Result<PreparedS1ManifestResult> {
    let files = manifest
        .files
        .iter()
        .filter(|file| {
            file.issues_extracted
                && matches!(
                    file.status,
                    S1RecoveryStatus::Completed
                        | S1RecoveryStatus::Partial
                        | S1RecoveryStatus::ReadyToResume
                        | S1RecoveryStatus::ResumeFailed
                )
        })
        .collect::<Vec<_>>();
    anyhow::ensure!(
        !files.is_empty(),
        "没有已经提取完成的 S1 文件结果可写入核查表"
    );
    let file_pairs = files
        .iter()
        .filter_map(|file| {
            file.audit_copy
                .as_ref()
                .map(|audit_workbook| S1IssueFilePair {
                    source_file: file.source_file.clone(),
                    audit_workbook: audit_workbook.clone(),
                })
        })
        .collect::<Vec<_>>();
    let formula_issues = files
        .iter()
        .flat_map(|file| file.formula_issues.iter().cloned())
        .collect::<Vec<_>>();
    let conditional_issues = files
        .iter()
        .flat_map(|file| file.conditional_issues.iter().cloned())
        .collect::<Vec<_>>();
    let formula_issue_count = formula_issues.len();
    let conditional_issue_count = conditional_issues.len();
    let merged = merge_s1_issue_sources(&file_pairs, &formula_issues, &conditional_issues)
        .context("按源文件顺序合流公式与条件格式异常")?;
    let history_rows = manifest.history_issues.len();
    let history = process_s1_history(S1HistoryRequest {
        period: manifest.period.clone(),
        batch_id: manifest.batch_id.clone(),
        audit_time: manifest.audit_time.clone(),
        audited_org_codes: None,
        current_issues: merged,
        history: manifest.history_issues.clone(),
    });
    Ok(PreparedS1ManifestResult {
        total_issue_count: history.current_issues.len(),
        issues: history.current_issues,
        formula_issue_count,
        conditional_issue_count,
        history_rows,
    })
}

fn summary_for_s1_result(
    path: &Path,
    prepared: &PreparedS1ManifestResult,
) -> Result<S1ResultWriteSummary> {
    let path = fs::canonicalize(path)
        .with_context(|| format!("解析 S1 结果工作簿路径 {}", path.display()))?;
    Ok(S1ResultWriteSummary {
        path,
        sheet: S1_RESULT_SHEET.to_owned(),
        formula_issue_count: prepared.formula_issue_count,
        conditional_issue_count: prepared.conditional_issue_count,
        history_rows: prepared.history_rows,
        total_issue_count: prepared.total_issue_count,
    })
}

fn write_s1_manifest_result(
    manifest: &S1RecoveryManifest,
    result_path: &Path,
) -> Result<S1ResultWriteSummary> {
    let prepared = prepare_s1_manifest_result(manifest)?;
    write_s1_result_workbook(result_path, None, &prepared.issues)
        .context("写入 S1-F01 最终审核结果")?;
    summary_for_s1_result(result_path, &prepared)
}

fn write_s1_recovery_manifest_result(
    store: &S1RecoveryStore,
    manifest: &mut S1RecoveryManifest,
) -> Result<S1ResultWriteSummary> {
    let prepared = prepare_s1_manifest_result(manifest)?;
    let input_bytes = serde_json::to_vec(&prepared.issues).context("序列化 S1-F3 结果输入")?;
    let input_sha256 = format!("{:x}", Sha256::digest(input_bytes));
    let result_path = s1_recovery_result_path(manifest, &input_sha256);
    let result_name = result_path
        .file_name()
        .and_then(|name| name.to_str())
        .context("生成 S1-F3 结果文件名")?;
    let candidate_path = result_path.with_file_name(format!(".{result_name}.candidate.xlsx"));

    if manifest.recovery_result_path.as_deref() != Some(result_path.as_path()) {
        manifest.recovery_result_path = Some(result_path.clone());
        manifest.recovery_result_candidate_path = Some(candidate_path.clone());
        manifest.recovery_result_input_sha256 = Some(input_sha256.clone());
        manifest.recovery_result_sha256 = None;
        store.save(manifest).context("保存 S1-F3 结果发布意图")?;
    }
    ensure!(
        manifest.recovery_result_input_sha256.as_deref() == Some(input_sha256.as_str()),
        "S1-F3 结果发布意图与当前审核结果不一致"
    );

    if result_path.exists() {
        let expected_sha256 = manifest
            .recovery_result_sha256
            .as_deref()
            .context("S1-F3 结果文件已存在，但持久化状态缺少发布摘要；拒绝覆盖或复用")?;
        ensure!(
            file_sha256(&result_path)? == expected_sha256,
            "S1-F3 已发布结果文件摘要变化；拒绝复用或覆盖"
        );
        return summary_for_s1_result(&result_path, &prepared);
    }

    let persisted_candidate = manifest
        .recovery_result_candidate_path
        .as_deref()
        .context("S1-F3 结果发布状态缺少候选文件路径")?;
    ensure!(
        persisted_candidate == candidate_path,
        "S1-F3 候选结果路径与当前恢复结果不一致"
    );
    if candidate_path.exists() {
        match manifest.recovery_result_sha256.as_deref() {
            Some(expected_sha256) => ensure!(
                file_sha256(&candidate_path)? == expected_sha256,
                "S1-F3 候选结果文件摘要变化；拒绝发布"
            ),
            None => {
                // A crash may leave an incomplete private candidate before its
                // digest was persisted. It is safe to discard only this
                // app-owned, deterministic candidate path and regenerate it.
                let expected_candidate_name = format!(".{result_name}.candidate.xlsx");
                ensure!(
                    candidate_path.parent() == Some(manifest.run_directory.as_path())
                        && candidate_path.file_name().and_then(|name| name.to_str())
                            == Some(expected_candidate_name.as_str()),
                    "S1-F3 中断候选路径不在本次运行目录，拒绝删除"
                );
                fs::remove_file(&candidate_path).context("清理摘要尚未持久化的 S1-F3 中断候选")?;
            }
        }
    }

    if !candidate_path.exists() {
        manifest.recovery_result_sha256 = None;
        store.save(manifest).context("保存 S1-F3 候选创建状态")?;
        write_s1_result_workbook(&candidate_path, None, &prepared.issues)
            .context("生成 S1-F3 暂存结果工作簿")?;
        let candidate_sha256 = file_sha256(&candidate_path)?;
        manifest.recovery_result_sha256 = Some(candidate_sha256);
        store.save(manifest).context("持久化 S1-F3 候选结果摘要")?;
    }

    ensure!(
        !result_path.exists(),
        "S1-F3 结果路径在发布前被其他文件占用；拒绝覆盖"
    );
    fs::rename(&candidate_path, &result_path).with_context(|| {
        format!(
            "发布 S1-F3 结果工作簿 {} → {}",
            candidate_path.display(),
            result_path.display()
        )
    })?;
    let published_sha256 = file_sha256(&result_path)?;
    ensure!(
        manifest.recovery_result_sha256.as_deref() == Some(published_sha256.as_str()),
        "发布后的 S1-F3 结果摘要与候选摘要不一致"
    );
    summary_for_s1_result(&result_path, &prepared)
}

fn execute_s1_f1_workflow(
    request: &S1F1PreflightRequest,
    accepted: &[S1F1AcceptedSource],
) -> Result<S1F1WorkflowOutput> {
    let adapter_started = Instant::now();
    let session_initialization_started = Instant::now();
    let mut recalculation_session = if cfg!(windows) {
        Some(S1RecalculationSession::new_native(&request.calc_engine)?)
    } else if matches!(request.calc_engine.as_str(), "auto" | "lo") {
        // Linux：一个审核批次创建一个由 sidecar 拥有的隔离 LibreOffice 会话
        // （rules/13 §4.5），逐本打开/重算/保存/关闭；创建失败时保持旧的
        // 逐本 sidecar 路径作为有界回退。
        match request
            .sidecar_path
            .as_deref()
            .ok_or_else(|| anyhow::anyhow!("缺少公式重算 sidecar"))
            .and_then(|sidecar| S1RecalculationSession::new(sidecar, &request.calc_engine))
        {
            Ok(session) => Some(session),
            Err(error) => {
                eprintln!(
                    "S1-F01 Linux 公式重算会话创建失败，回退逐本 sidecar 路径：{error:#}"
                );
                None
            }
        }
    } else {
        None
    };
    let session_initialization_us = session_initialization_started.elapsed().as_micros();
    let mut output = execute_s1_f1_workflow_with_dependencies(
        request,
        accepted,
        None,
        |item, request, audit_copy, modifier| {
            if let Some(session) = recalculation_session.as_mut() {
                execute_s1_formula_copy_with_manual_recovery_session(
                    &item.source,
                    &item.template,
                    audit_copy,
                    request.mapping_config.as_deref(),
                    request.external_data.as_deref(),
                    true,
                    modifier,
                    session,
                )
            } else {
                execute_s1_formula_copy_with_manual_recovery(
                    &item.source,
                    &item.template,
                    audit_copy,
                    request.mapping_config.as_deref(),
                    request.external_data.as_deref(),
                    request.sidecar_path.as_deref(),
                    modifier,
                    &request.calc_engine,
                )
            }
        },
    )?;
    let session_shutdown_started = Instant::now();
    drop(recalculation_session);
    let session_shutdown_us = session_shutdown_started.elapsed().as_micros();
    output.performance_timings["recalculationSessionInitializationUs"] =
        json!(session_initialization_us);
    output.performance_timings["recalculationSessionShutdownUs"] =
        json!(session_shutdown_us);
    output.performance_timings["workflowAdapterTotalUs"] =
        json!(adapter_started.elapsed().as_micros());
    Ok(output)
}

fn execute_s1_f1_workflow_with_dependencies<F>(
    request: &S1F1PreflightRequest,
    accepted: &[S1F1AcceptedSource],
    recovery_store: Option<S1RecoveryStore>,
    mut copy_executor: F,
) -> Result<S1F1WorkflowOutput>
where
    F: FnMut(
        &S1F1AcceptedSource,
        &S1F1PreflightRequest,
        &Path,
        &ExcelModifier,
    ) -> Result<S1FormulaCopyOutcome>,
{
    let workflow_started = Instant::now();
    anyhow::ensure!(
        !accepted.is_empty(),
        "S1-F01 没有通过模板及表结构预检的源文件"
    );
    let sources = accepted
        .iter()
        .map(|item| item.source.clone())
        .collect::<Vec<_>>();
    let period_detection_started = Instant::now();
    let detection = detect_period_from_workbooks(&request.source_directory, &sources);
    anyhow::ensure!(
        !detection.conflict && !detection.period.trim().is_empty(),
        "无法唯一识别 S1-F01 数据期：{}",
        detection.details
    );
    let period = detection.period;
    let period_detection_us = period_detection_started.elapsed().as_micros();
    let workflow_setup_started = Instant::now();
    fs::create_dir_all(&request.output_directory).with_context(|| {
        format!(
            "创建 S1-F01 输出目录 {}",
            request.output_directory.display()
        )
    })?;
    fs::create_dir(&request.run_directory).with_context(|| {
        format!(
            "创建本次 S1-F01 审核目录 {}",
            request.run_directory.display()
        )
    })?;

    // The production route registers the qualified backend for the GUI's
    // 高级设置 edit-engine selection (Umya by default); on unverified
    // platforms it stays empty and the router fails closed before any edit.
    let (modifier, capability) =
        crate::excel_route::production_modifier_with_edit_engine("S1-F01", &request.edit_engine)
            .context("初始化 S1-F01 Excel-Core 路由")?;
    anyhow::ensure!(
        capability.is_some(),
        "S1-F01 当前平台没有已登记的 Excel-Core Passed 后端，编辑批次保持失败关闭"
    );

    let batch_id = request.batch_id.clone();
    let audit_time = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
    let history_rows = if let Some(path) = request.mapping_config.as_deref() {
        read_s1_history_workbook(path, None)
            .with_context(|| {
                format!(
                    "从 S1-F01 配置/历史工作簿读取历史审核记录：{}",
                    path.display()
                )
            })?
            .issues
    } else {
        Vec::new()
    };
    let mapping_config_sha256 = request
        .mapping_config
        .as_deref()
        .map(file_sha256)
        .transpose()
        .context("计算 S1-F01 映射配置摘要")?;
    let store = match recovery_store {
        Some(store) => store,
        None => S1RecoveryStore::user_default().context("初始化 S1-F3 持久化恢复任务目录")?,
    };
    let mut recovery_files = Vec::with_capacity(accepted.len());
    for (index, item) in accepted.iter().enumerate() {
        recovery_files.push(S1RecoveryFile::new(
            item.source.clone(),
            item.template.clone(),
            file_sha256(&item.template)
                .with_context(|| format!("计算模板摘要：{}", item.template.display()))?,
            Some(
                request
                    .run_directory
                    .join(s1_recalculated_filename(index + 1, &item.source)?),
            ),
            S1RecoveryStatus::Copying,
        ));
    }
    let mut manifest = S1RecoveryManifest::new(
        request.batch_id.clone(),
        request.output_directory.clone(),
        request.run_directory.clone(),
        period.clone(),
        batch_id.clone(),
        audit_time.clone(),
        request.mapping_config.clone(),
        mapping_config_sha256,
        history_rows,
        Vec::new(),
        recovery_files,
    );
    manifest.external_data = request.external_data.clone();
    store
        .save(&mut manifest)
        .context("保存 S1-F01 正式审核前的恢复任务清单")?;

    let mut audit_copies = Vec::with_capacity(accepted.len());
    let mut recalculation_engines = Vec::new();
    let mut file_results = Vec::with_capacity(accepted.len());
    let workflow_setup_us = workflow_setup_started.elapsed().as_micros();
    let mut copy_executor_us = 0u128;
    let mut excel_core_edit_us = 0u128;
    let mut copy_preparation_other_us = 0u128;
    let mut external_recalculation_us = 0u128;
    let mut issue_extraction_us = 0u128;
    let mut formula_issue_extraction_us = 0u128;
    let mut conditional_format_extraction_us = 0u128;
    let mut issue_extraction_other_us = 0u128;
    let mut file_timings = Vec::with_capacity(accepted.len());

    for item in accepted {
        let file_id = stable_file_id(&item.source);
        let file_index = manifest
            .files
            .iter()
            .position(|file| file.file_id == file_id)
            .context("恢复清单缺少当前 S1-F01 源文件")?;
        let audit_copy = manifest.files[file_index]
            .audit_copy
            .clone()
            .context("恢复清单缺少审核副本路径")?;
        let copy_started = Instant::now();
        let outcome = copy_executor(item, request, &audit_copy, &modifier);
        let file_copy_executor_us = copy_started.elapsed().as_micros();
        copy_executor_us += file_copy_executor_us;
        let org = organisation_from_name(&item.source);
        let metadata = S1FormulaMetadata {
            period: period.clone(),
            batch_id: batch_id.clone(),
            audit_time: audit_time.clone(),
            org_code: org.0,
            org_name: org.1,
        };
        match outcome {
            Ok(S1FormulaCopyOutcome::WaitingManualRecalc(copy)) => {
                let editor_elapsed_us = copy.edit.elapsed_ms.saturating_mul(1_000);
                excel_core_edit_us += editor_elapsed_us;
                let copy_other_us = file_copy_executor_us.saturating_sub(editor_elapsed_us);
                copy_preparation_other_us += copy_other_us;
                file_timings.push(json!({
                    "sourceFile": item.source,
                    "copyAndRecalculationPreparationUs": file_copy_executor_us,
                    "excelCoreEditAndValidationUs": editor_elapsed_us,
                    "copyPreparationAndValidationOtherUs": copy_other_us,
                    "externalRecalculationUs": Value::Null,
                    "formulaIssueExtractionUs": Value::Null,
                    "conditionalFormatExtractionUs": Value::Null
                }));
                let (fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
                    match workbook_recovery_fingerprints(&copy.audit_copy) {
                        Ok(value) if value.3 > 0 => value,
                        Ok(_) => {
                            manifest.files[file_index].status = S1RecoveryStatus::Failed;
                            manifest.files[file_index].reason =
                                Some("已保留副本，但未发现可恢复的公式。".to_owned());
                            file_results.push(json!({
                                "sourceFile": item.source,
                                "status": manifest.files[file_index].status.as_str(),
                                "reason": manifest.files[file_index].reason,
                                "auditWorkbook": copy.audit_copy
                            }));
                            store.save(&mut manifest)?;
                            continue;
                        }
                        Err(error) => {
                            manifest.files[file_index].status = S1RecoveryStatus::Failed;
                            manifest.files[file_index].reason = Some(format!(
                                "审核副本无法建立公式和非公式数据完整性基线：{error:#}"
                            ));
                            file_results.push(json!({
                                "sourceFile": item.source,
                                "status": manifest.files[file_index].status.as_str(),
                                "reason": manifest.files[file_index].reason,
                                "auditWorkbook": copy.audit_copy
                            }));
                            store.save(&mut manifest)?;
                            continue;
                        }
                    };
                manifest.files[file_index].audit_copy = Some(copy.audit_copy.clone());
                manifest.files[file_index].formula_fingerprint = Some(fingerprint);
                manifest.files[file_index].formula_fingerprint_normalized =
                    Some(normalized_fingerprint);
                manifest.files[file_index].audit_copy_semantic_fingerprint =
                    Some(semantic_fingerprint);
                manifest.files[file_index].status = S1RecoveryStatus::WaitingManualRecalc;
                manifest.files[file_index].reason = Some(copy.reason.clone());
                manifest.files[file_index].recalculation_engine = Some("manual_pending".to_owned());
                audit_copies.push(json!({
                    "sourceFile": item.source,
                    "template": item.template,
                    "auditWorkbook": copy.audit_copy,
                    "editorEngine": copy.edit.engine.to_string(),
                    "recalculationEngine": "manual_pending",
                    "formulaCellsPrepared": copy.formula_cells_prepared,
                    "externalSheetsImported": copy.external_sheets_imported,
                    "reason": copy.reason
                }));
                file_results.push(json!({
                    "sourceFile": item.source,
                    "status": manifest.files[file_index].status.as_str(),
                    "reason": manifest.files[file_index].reason,
                    "auditWorkbook": copy.audit_copy,
                    "formulaCells": formula_count
                }));
                store.save(&mut manifest)?;
            }
            Ok(S1FormulaCopyOutcome::Recalculated(execution)) => {
                let editor_elapsed_us = execution.edit.elapsed_ms.saturating_mul(1_000);
                let recalculation_elapsed_us =
                    execution.recalculation.elapsed_ms.saturating_mul(1_000);
                excel_core_edit_us += editor_elapsed_us;
                external_recalculation_us += recalculation_elapsed_us;
                let copy_other_us = file_copy_executor_us
                    .saturating_sub(editor_elapsed_us)
                    .saturating_sub(recalculation_elapsed_us);
                copy_preparation_other_us += copy_other_us;
                let audit_workbook = execution.recalculation.output.clone();
                let (fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
                    match workbook_recovery_fingerprints(&audit_workbook) {
                        Ok(value) if value.3 > 0 => value,
                        Ok(_) => {
                            manifest.files[file_index].status = S1RecoveryStatus::ResumeFailed;
                            manifest.files[file_index].reason = Some(
                                "重算输出中未发现公式，保留副本并等待 S1-F3 复核。".to_owned(),
                            );
                            manifest.files[file_index].audit_copy = Some(audit_workbook.clone());
                            file_results.push(json!({
                                "sourceFile": item.source,
                                "status": manifest.files[file_index].status.as_str(),
                                "reason": manifest.files[file_index].reason,
                                "auditWorkbook": audit_workbook
                            }));
                            store.save(&mut manifest)?;
                            continue;
                        }
                        Err(error) => {
                            manifest.files[file_index].status = S1RecoveryStatus::ResumeFailed;
                            manifest.files[file_index].reason = Some(format!(
                                "重算输出无法建立公式和非公式数据完整性基线：{error:#}"
                            ));
                            manifest.files[file_index].audit_copy = Some(audit_workbook.clone());
                            file_results.push(json!({
                                "sourceFile": item.source,
                                "status": manifest.files[file_index].status.as_str(),
                                "reason": manifest.files[file_index].reason,
                                "auditWorkbook": audit_workbook
                            }));
                            store.save(&mut manifest)?;
                            continue;
                        }
                    };
                manifest.files[file_index].audit_copy = Some(audit_workbook.clone());
                manifest.files[file_index].formula_fingerprint = Some(fingerprint);
                manifest.files[file_index].formula_fingerprint_normalized =
                    Some(normalized_fingerprint);
                manifest.files[file_index].audit_copy_semantic_fingerprint =
                    Some(semantic_fingerprint);
                manifest.files[file_index].recalculation_engine =
                    Some(execution.recalculation.engine.clone());
                let extraction_started = Instant::now();
                let extracted =
                    extract_s1_audit_copy(item, &item.source, &audit_workbook, &metadata);
                let file_issue_extraction_us = extraction_started.elapsed().as_micros();
                issue_extraction_us += file_issue_extraction_us;
                let (file_formula_extraction_us, file_conditional_extraction_us) =
                    extracted.as_ref().map_or((0, 0), |extracted| {
                        (
                            extracted.formula_extraction_elapsed_us,
                            extracted.conditional_format_extraction_elapsed_us,
                        )
                    });
                formula_issue_extraction_us += file_formula_extraction_us;
                conditional_format_extraction_us += file_conditional_extraction_us;
                issue_extraction_other_us += file_issue_extraction_us
                    .saturating_sub(file_formula_extraction_us)
                    .saturating_sub(file_conditional_extraction_us);
                file_timings.push(json!({
                    "sourceFile": item.source,
                    "copyAndRecalculationPreparationUs": file_copy_executor_us,
                    "excelCoreEditAndValidationUs": editor_elapsed_us,
                    "copyPreparationAndValidationOtherUs": copy_other_us,
                    "externalRecalculationUs": recalculation_elapsed_us,
                    "issueExtractionTotalUs": file_issue_extraction_us,
                    "formulaIssueExtractionUs": file_formula_extraction_us,
                    "conditionalFormatExtractionUs": file_conditional_extraction_us,
                    "issueExtractionOtherUs": file_issue_extraction_us
                        .saturating_sub(file_formula_extraction_us)
                        .saturating_sub(file_conditional_extraction_us)
                }));
                match extracted {
                    Ok(extracted) => {
                        manifest.files[file_index].formula_issues = extracted.formula_issues;
                        manifest.files[file_index].conditional_issues =
                            extracted.conditional_issues;
                        manifest.files[file_index].conditional_diagnostics = extracted.diagnostics;
                        manifest.files[file_index].status = S1RecoveryStatus::ReadyToResume;
                        manifest.files[file_index].issues_extracted = true;
                        manifest.files[file_index].audit_partial = extracted.partial;
                        manifest.files[file_index].reason = extracted.reason;
                    }
                    Err(error) => {
                        manifest.files[file_index].status = S1RecoveryStatus::ResumeFailed;
                        manifest.files[file_index].reason = Some(format!(
                            "自动重算完成，但审核异常提取失败；可用 S1-F3 重试：{error:#}"
                        ));
                    }
                }
                if !recalculation_engines
                    .iter()
                    .any(|engine| engine == &execution.recalculation.engine)
                {
                    recalculation_engines.push(execution.recalculation.engine.clone());
                }
                audit_copies.push(json!({
                    "sourceFile": item.source,
                    "template": item.template,
                    "auditWorkbook": audit_workbook,
                    "editorEngine": execution.edit.engine.to_string(),
                    "recalculationEngine": execution.recalculation.engine,
                    "formulaCellsPrepared": execution.formula_cells_prepared,
                    "externalSheetsImported": execution.external_sheets_imported,
                    "formulaCells": formula_count
                }));
                file_results.push(json!({
                    "sourceFile": item.source,
                    "status": manifest.files[file_index].status.as_str(),
                    "reason": manifest.files[file_index].reason,
                    "auditWorkbook": audit_workbook
                }));
                store.save(&mut manifest)?;
            }
            Err(error) => {
                copy_preparation_other_us += file_copy_executor_us;
                file_timings.push(json!({
                    "sourceFile": item.source,
                    "copyAndRecalculationPreparationUs": file_copy_executor_us,
                    "excelCoreEditAndValidationUs": Value::Null,
                    "copyPreparationAndValidationOtherUs": file_copy_executor_us,
                    "externalRecalculationUs": Value::Null,
                    "issueExtractionTotalUs": 0,
                    "formulaIssueExtractionUs": 0,
                    "conditionalFormatExtractionUs": 0
                }));
                manifest.files[file_index].status = S1RecoveryStatus::Failed;
                manifest.files[file_index].reason =
                    Some(format!("公式复制或重算准备失败：{error:#}"));
                file_results.push(json!({
                    "sourceFile": item.source,
                    "status": manifest.files[file_index].status.as_str(),
                    "reason": manifest.files[file_index].reason
                }));
                store.save(&mut manifest)?;
            }
        }
    }

    let mut formula_issue_count = manifest
        .files
        .iter()
        .map(|file| file.formula_issues.len())
        .sum::<usize>();
    let mut conditional_issue_count = manifest
        .files
        .iter()
        .map(|file| file.conditional_issues.len())
        .sum::<usize>();
    let mut history_row_count = manifest.history_issues.len();
    let mut total_issue_count = formula_issue_count + conditional_issue_count;
    let mut result_path = None;
    let mut result_sheet = None;
    let mut result_error = None;
    let has_extractable_files = manifest.files.iter().any(|file| {
        file.issues_extracted
            && matches!(
                file.status,
                S1RecoveryStatus::ReadyToResume
                    | S1RecoveryStatus::Completed
                    | S1RecoveryStatus::Partial
            )
    });
    let result_merge_write_started = Instant::now();
    if has_extractable_files {
        match write_s1_manifest_result(&manifest, &request.planned_result_path) {
            Ok(summary) => {
                result_path = Some(summary.path.clone());
                result_sheet = Some(summary.sheet);
                formula_issue_count = summary.formula_issue_count;
                conditional_issue_count = summary.conditional_issue_count;
                history_row_count = summary.history_rows;
                total_issue_count = summary.total_issue_count;
                manifest.result_path = Some(summary.path);
                for file in &mut manifest.files {
                    if file.status == S1RecoveryStatus::ReadyToResume && file.issues_extracted {
                        file.status = if file.audit_partial {
                            S1RecoveryStatus::Partial
                        } else {
                            S1RecoveryStatus::Completed
                        };
                    }
                }
                store.save(&mut manifest)?;
            }
            Err(error) => {
                result_error = Some(format!("生成审核核查表失败：{error:#}"));
                for file in &mut manifest.files {
                    if file.status == S1RecoveryStatus::ReadyToResume && file.issues_extracted {
                        file.status = S1RecoveryStatus::ResumeFailed;
                        file.reason = result_error.clone();
                    }
                }
                store.save(&mut manifest)?;
            }
        }
    }
    let result_merge_history_write_us = result_merge_write_started.elapsed().as_micros();
    let waiting_manual_recalc = manifest
        .files
        .iter()
        .filter(|file| file.status == S1RecoveryStatus::WaitingManualRecalc)
        .count();
    let partial_files = manifest
        .files
        .iter()
        .filter(|file| file.status == S1RecoveryStatus::Partial)
        .count();
    let failed_files = manifest
        .files
        .iter()
        .filter(|file| {
            matches!(
                file.status,
                S1RecoveryStatus::Failed | S1RecoveryStatus::ResumeFailed
            )
        })
        .count();
    file_results.clear();
    for file in &manifest.files {
        file_results.push(json!({
            "sourceFile": file.source_file,
            "status": file.status.as_str(),
            "reason": file.reason,
            "auditWorkbook": file.audit_copy,
            "issuesExtracted": file.issues_extracted,
            "partial": file.audit_partial
        }));
    }
    let workflow_total_us = workflow_started.elapsed().as_micros();
    let workflow_other_and_recovery_persistence_us = workflow_total_us
        .saturating_sub(period_detection_us)
        .saturating_sub(workflow_setup_us)
        .saturating_sub(copy_executor_us)
        .saturating_sub(issue_extraction_us)
        .saturating_sub(result_merge_history_write_us);
    let performance_timings = json!({
        "unit": "microseconds",
        "periodDetectionUs": period_detection_us,
        "workflowSetupAndInitialManifestUs": workflow_setup_us,
        "copyAndRecalculationPreparationTotalUs": copy_executor_us,
        "excelCoreEditAndOutputValidationUs": excel_core_edit_us,
        "copyPreparationAndValidationOtherUs": copy_preparation_other_us,
        "externalExcelRecalculationUs": external_recalculation_us,
        "issueExtractionTotalUs": issue_extraction_us,
        "formulaAndValidationIssueExtractionUs": formula_issue_extraction_us,
        "conditionalFormatIssueExtractionUs": conditional_format_extraction_us,
        "issueExtractionOtherUs": issue_extraction_other_us,
        "resultMergeHistoryAndWorkbookWriteUs": result_merge_history_write_us,
        "otherWorkflowAndRecoveryPersistenceUs": workflow_other_and_recovery_persistence_us,
        "workflowTotalUs": workflow_total_us,
        "perFile": file_timings
    });
    Ok(S1F1WorkflowOutput {
        period,
        result_path,
        result_sheet,
        formula_issue_count,
        conditional_issue_count,
        history_rows: history_row_count,
        total_issue_count,
        audit_copies,
        recalculation_engines,
        file_results,
        waiting_manual_recalc,
        partial_files,
        failed_files,
        result_error,
        performance_timings,
    })
}

fn s1_f1_template_preflight(
    template_path: &Path,
    mapping_config: Option<&Path>,
) -> Result<S1F1TemplatePreflight> {
    let rule_plan = load_s1_template_rule_plan(template_path, mapping_config)
        .context("读取 S1-F01 模板公式规则")?;
    let mappings = load_s1_conditional_mappings(mapping_config, template_path)
        .context("读取 S1-F01 条件格式与表结构映射")?;
    let defined_names =
        excel_reader::read_defined_names(template_path).context("读取 S1-F01 模板命名区域")?;
    let sheet_names =
        excel_reader::read_sheet_names(template_path).context("读取 S1-F01 模板工作表目录")?;
    let areas = named_areas_from_defined_names(&defined_names);
    let area_plan = resolve_s1_conditional_area_plan(&mappings, None, &areas, &sheet_names)
        .context("解析 S1-F01 模板扫描区域")?;
    if area_plan.structure_ranges.is_empty() {
        anyhow::bail!("模板没有可用的表结构区域，不能把源文件当作结构通过");
    }
    let required_sheets = area_plan
        .scan_areas
        .iter()
        .filter_map(|area| s1_f1_scan_area_sheet(area))
        .chain(
            area_plan
                .structure_ranges
                .iter()
                .map(|range| range.sheet_name.clone()),
        )
        .collect::<Vec<_>>();
    let structure = read_s1_structure_template_with_required_sheets(
        template_path,
        &area_plan.structure_ranges,
        &required_sheets,
    )
    .context("读取 S1-F01 模板表结构标签")?;
    let health = assess_s1_template_health(&rule_plan, &area_plan.structure_ranges);
    let health_errors = health
        .items
        .iter()
        .filter(|item| item.level == "错误")
        .map(|item| format!("{} {}：{}", item.sheet_name, item.location, item.message))
        .collect::<Vec<_>>();
    let scan_area_count = area_plan.scan_areas.len();
    let structure_area_count = area_plan.structure_ranges.len();
    Ok(S1F1TemplatePreflight {
        structure,
        rule_plan,
        scan_areas: area_plan.scan_areas,
        structure_ranges: area_plan.structure_ranges,
        scan_area_count,
        structure_area_count,
        health_passed: health.passed,
        health_errors,
    })
}

fn cached_s1_f1_template_preflight(
    cache: &mut std::collections::HashMap<
        PathBuf,
        std::result::Result<Arc<S1F1TemplatePreflight>, String>,
    >,
    template_path: &Path,
    mapping_config: Option<&Path>,
) -> Result<Arc<S1F1TemplatePreflight>> {
    let cached = cache.entry(template_path.to_path_buf()).or_insert_with(|| {
        s1_f1_template_preflight(template_path, mapping_config)
            .map(Arc::new)
            .map_err(|error| format!("{error:#}"))
    });
    match cached {
        Ok(preflight) => Ok(Arc::clone(preflight)),
        Err(message) => Err(anyhow::anyhow!(message.clone())),
    }
}

fn s1_f1_scan_area_sheet(specification: &str) -> Option<String> {
    let (sheet_name, _) = specification.split_once('!')?;
    let value = sheet_name.trim();
    let value = value
        .strip_prefix('\'')
        .and_then(|name| name.strip_suffix('\''))
        .map(|name| name.replace("''", "'"))
        .unwrap_or_else(|| value.to_owned());
    (!value.is_empty()).then_some(value)
}

fn web_path_identity(path: &Path) -> String {
    let resolved = fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
    let identity = resolved.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    {
        identity.to_lowercase()
    }
    #[cfg(not(windows))]
    {
        identity
    }
}

fn find_s1_recalculation_sidecar() -> Option<PathBuf> {
    if let Some(configured) = std::env::var_os("AUDIT_S1_RECALC_SIDECAR") {
        let path = PathBuf::from(configured);
        if path.is_file() {
            return fs::canonicalize(path).ok();
        }
    }
    let executable = std::env::current_exe().ok()?;
    let parent = executable.parent()?;
    let name = if cfg!(windows) {
        "lo-convert.exe"
    } else {
        "lo-convert"
    };
    let candidate = parent.join(name);
    candidate
        .is_file()
        .then(|| fs::canonicalize(candidate).ok())
        .flatten()
}

fn start_dag_flow(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let workflow = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "start_dag_flow 缺少流程名称"))?;
    if workflow == "dag:汇总核查表校验" {
        return start_s1_f1_workflow(state, args);
    }
    if workflow != "dag:汇总校验结果说明" {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版尚未接通工作流：{workflow}"),
        ));
    }
    let values = args.get(1).and_then(Value::as_object).ok_or_else(|| {
        ApiError::new(StatusCode::BAD_REQUEST, "启动 S1-F02 需要当前 GUI 路径设置")
    })?;
    let input_dir = values
        .get("input")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择 S1-F02 源数据目录"))?;
    if !input_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F02 源数据目录不存在：{}", input_dir.display()),
        ));
    }
    let input_dir = fs::canonicalize(&input_dir).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F02 源目录失败：{error}"),
        )
    })?;
    let source_files = s1_region_summary_source_workbooks(&input_dir, true)
        .map_err(|error| ApiError::internal(format!("扫描 S1-F02 源文件失败：{error:#}")))?;
    if source_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "源数据目录及其子目录中没有可汇总的 .xlsx 文件",
        ));
    }
    let requested_selection = args
        .get(2)
        .and_then(Value::as_array)
        .map(|values| {
            values
                .iter()
                .map(|value| {
                    value.as_str().map(PathBuf::from).ok_or_else(|| {
                        ApiError::new(StatusCode::BAD_REQUEST, "已选源文件路径必须是字符串")
                    })
                })
                .collect::<std::result::Result<Vec<_>, _>>()
        })
        .transpose()?
        .unwrap_or_default();
    let selected_files = if requested_selection.is_empty() {
        source_files
    } else {
        let eligible = source_files
            .iter()
            .map(|path| fs::canonicalize(path))
            .collect::<std::io::Result<std::collections::HashSet<_>>>()
            .map_err(|error| ApiError::internal(format!("解析 S1-F02 源文件路径失败：{error}")))?;
        let mut selected = Vec::new();
        for path in requested_selection {
            let canonical = fs::canonicalize(&path).map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("已选 S1-F02 源文件不可用：{}：{error}", path.display()),
                )
            })?;
            if !eligible.contains(&canonical) {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!(
                        "已选文件不在当前可汇总源目录或扫描范围内：{}",
                        path.display()
                    ),
                ));
            }
            selected.push(canonical);
        }
        selected.sort();
        selected.dedup();
        selected
    };
    if selected_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "没有选中可汇总的 S1-F02 源文件",
        ));
    }

    let template_path = values
        .get("template")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from);
    let template_path = match template_path {
        Some(path) => path,
        None => {
            let template_dir = values
                .get("templateDir")
                .and_then(Value::as_str)
                .filter(|path| !path.trim().is_empty())
                .map(PathBuf::from)
                .filter(|path| path.is_dir())
                .ok_or_else(|| {
                    ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "请先选择匹配模板，或选择模板目录后自动识别",
                    )
                })?;
            let template_profiles = load_template_profiles(&template_dir).map_err(|error| {
                ApiError::internal(format!("自动识别 S1-F02 模板失败：{error:#}"))
            })?;
            let mut matched = std::collections::BTreeSet::new();
            let mut unmatched = Vec::new();
            for source in &selected_files {
                let suggestion = recommend_template_with_profiles(
                    &template_profiles,
                    std::slice::from_ref(source),
                );
                if suggestion.matched {
                    if let Some(path) = suggestion.template_path {
                        matched.insert(path);
                    } else {
                        unmatched.push(source.display().to_string());
                    }
                } else {
                    unmatched.push(source.display().to_string());
                }
            }
            if !unmatched.is_empty() || matched.len() != 1 {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!(
                        "自动识别未得到唯一模板；未匹配文件 {} 个、匹配模板 {} 个。请按模板筛选文件或手动选择模板。",
                        unmatched.len(),
                        matched.len()
                    ),
                ));
            }
            matched.into_iter().next().expect("one matched template")
        }
    };
    if !template_path.is_file()
        || !template_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "S1-F02 模板必须是存在的 .xlsx 文件：{}",
                template_path.display()
            ),
        ));
    }
    let template_path = fs::canonicalize(&template_path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F02 模板失败：{error}"),
        )
    })?;

    let history_workbook = state.lock().value["historyConfig"]
        .as_str()
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                "请先选择逐笔统计系统配置，以关联历史审核结果",
            )
        })?;
    if !history_workbook.is_file()
        || !history_workbook
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "逐笔统计系统历史配置必须是存在的 .xlsx 文件：{}",
                history_workbook.display()
            ),
        ));
    }
    let history_workbook = fs::canonicalize(&history_workbook).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析历史审核配置失败：{error}"),
        )
    })?;

    let output_directory = values
        .get("output")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| input_dir.join("执行结果"));
    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F02 输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory)
        .map_err(|error| ApiError::internal(format!("创建 S1-F02 输出目录失败：{error}")))?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析 S1-F02 输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("校验结果与报送说明汇总_{stamp}.xlsx"));
    if output_workbook.exists() {
        output_workbook = output_directory.join(format!(
            "校验结果与报送说明汇总_{stamp}_{:03}.xlsx",
            chrono::Local::now().timestamp_subsec_millis()
        ));
    }
    if output_workbook.exists() {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "S1-F02 时间戳输出文件已存在，未覆盖；请稍后重试",
        ));
    }

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行 S1-F02",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在汇总校验结果说明，请勿关闭窗口……");
        ui.value["template"] = json!(template_path.to_string_lossy());
        ui.value["auditSummaryOutput"] = json!(output_workbook.to_string_lossy());
        push_log(
            &mut ui.value,
            format!("S1-F02 开始汇总：已选 {} 个源工作簿", selected_files.len()),
        );
        push_detail_log(
            &mut ui.value,
            format!(
                "S1-F02 输入：{}；模板：{}；历史配置：{}；reader=calamine+namespace_aware_ooxml；writer=rust_xlsxwriter；输出：{}",
                selected_files.len(),
                template_path.display(),
                history_workbook.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s1-region-summary".to_owned())
        .spawn(move || {
            run_s1_region_summary_worker(
                worker_state,
                template_path,
                input_dir,
                output_workbook,
                history_workbook,
                selected_files,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S1-F02 汇总启动失败");
        push_log(&mut ui.value, format!("启动 S1-F02 后台任务失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S1-F02 后台任务失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s1_region_summary_worker(
    state: AppState,
    template_path: PathBuf,
    input_dir: PathBuf,
    output_workbook: PathBuf,
    history_workbook: PathBuf,
    selected_files: Vec<PathBuf>,
) {
    let started = Instant::now();
    let feature_names = ["任意行汇总".to_owned()];
    let result = summarize_s1_named_regions_selected(
        &template_path,
        &input_dir,
        &output_workbook,
        Some(&history_workbook),
        Some(&history_workbook),
        &feature_names,
        true,
        &selected_files,
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(summary) => {
            let path = summary.output_path.to_string_lossy().into_owned();
            ui.value["auditSummaryOutput"] = json!(path);
            ui.value["status"] = json!(format!("S1-F02 汇总完成：{} 行", summary.data_rows));
            push_log(
                &mut ui.value,
                format!(
                    "S1-F02 汇总完成：{} 个源工作簿，结构通过 {} 个、部分继续 {} 个、跳过 {} 个；{} 张输出表、{} 行；文件 {}",
                    summary.source_workbooks,
                    summary.structure_preflight_passed_workbooks,
                    summary.structure_preflight_partial_workbooks,
                    summary.structure_preflight_skipped_workbooks,
                    summary.output_sheets,
                    summary.data_rows,
                    summary.output_path.display()
                ),
            );
            for detail in &summary.structure_preflight_skipped_details {
                push_log(&mut ui.value, format!("S1-F02 结构预检跳过文件：{detail}"));
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F02 完成：reader={}；writer={}；结构预检={} 个，通过={}，部分继续={}，跳过={}；历史匹配={} 行；扫描={} ms；模板读取={} ms；结构预检={} ms；源读取={} ms；汇总={} ms；历史富化={} ms；写入={} ms；读回校验={} ms；任务总计={} ms",
                    summary.reader,
                    summary.writer,
                    summary.structure_preflight_checked_workbooks,
                    summary.structure_preflight_passed_workbooks,
                    summary.structure_preflight_partial_workbooks,
                    summary.structure_preflight_skipped_workbooks,
                    summary.history_rows_matched,
                    summary.discovery_time_ms,
                    summary.template_read_time_ms,
                    summary.structure_preflight_time_ms,
                    summary.source_read_time_ms,
                    summary.calculation_time_ms,
                    summary.history_elapsed_ms,
                    summary.write_time_ms,
                    summary.validation_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("S1-F02 汇总失败");
            push_log(
                &mut ui.value,
                format!("S1-F02 汇总失败，未报告成功：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F02 失败：reader=calamine+namespace_aware_ooxml；writer=rust_xlsxwriter；elapsed_ms={}；error={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn choose_central(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_central 缺少路径类型"))?;
    let is_file = match field {
        "centralCur"
        | "centralPre"
        | "centralCompareFile"
        | "centralExplainFile"
        | "centralCommonConfig"
        | "centralComparisonConfig"
        | "centralFormsConfig"
        | "centralCrossCur"
        | "centralCrossConfig" => true,
        "centralOutput" | "centralFormOutput" => false,
        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("Rust 迁移版暂不支持选择大集中路径：{field}"),
            ));
        }
    };
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(state.lock().value[field].clone());
    }
    let path = PathBuf::from(unquote_path(selected));
    if is_file {
        if !path.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("所选大集中工作簿不存在：{}", path.display()),
            ));
        }
        let extension = path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default();
        let supported = match field {
            "centralCur" | "centralPre" => {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            }
            "centralCompareFile" => {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }
            "centralExplainFile"
            | "centralCommonConfig"
            | "centralComparisonConfig"
            | "centralFormsConfig" => extension.eq_ignore_ascii_case("xlsx"),
            "centralCrossCur" => {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            }
            "centralCrossConfig" => extension.eq_ignore_ascii_case("xlsx"),
            _ => false,
        };
        if !supported {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                match field {
                    "centralCur" | "centralPre" => "大集中本期/上期数据目前只支持 CSV、XLSX、XLS",
                    "centralCompareFile" => "S3-F02 比较结果目前只支持 .xlsx 和 .xlsm",
                    "centralExplainFile" => "说明工作簿目前只支持 .xlsx",
                    "centralCommonConfig" => "大集中通用配置目前只支持 .xlsx",
                    "centralFormsConfig" => "大集中指标比较拆分配置目前只支持 .xlsx",
                    "centralCrossCur" => "S3-F04 只支持 CSV、XLSX、XLS 输入",
                    "centralCrossConfig" => "本期数值核对配置目前只支持 .xlsx",
                    _ => "大集中执行比较配置目前只支持 .xlsx",
                },
            ));
        }
    } else if !path.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选大集中输出目录不存在：{}", path.display()),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析大集中路径失败：{error}"),
        )
    })?;

    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改大集中路径",
        ));
    }
    let value = path.to_string_lossy().into_owned();
    ui.value[field] = json!(value);
    match field {
        "centralCur" => {
            if ui.value["centralOutputAuto"] == true {
                ui.value["centralOutput"] = json!(path
                    .parent()
                    .unwrap_or_else(|| Path::new("."))
                    .join("执行结果")
                    .to_string_lossy());
            }
        }
        "centralCompareFile" => {
            let parent = path
                .parent()
                .map(|parent| parent.to_string_lossy().into_owned())
                .unwrap_or_default();
            if ui.value["centralOutputAuto"] == true {
                ui.value["centralOutput"] = json!(parent);
            }
            if ui.value["centralFormOutputAuto"] == true {
                ui.value["centralFormOutput"] = json!(parent);
            }
        }
        "centralOutput" => ui.value["centralOutputAuto"] = json!(false),
        "centralFormOutput" => ui.value["centralFormOutputAuto"] = json!(false),
        _ => {}
    }
    drop(ui);
    push_log_for_state(state, &format!("S3 路径已选择：{field}"));
    persist_state_preferences(state)?;
    Ok(state.lock().value[field].clone())
}

fn start_central_comparison(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (current_file, previous_file, common_config, feature_config, output_directory, schema) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行大集中比较",
            ));
        }
        let current_file = ui.value["centralCur"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期大集中数据"))?;
        let previous_file = ui.value["centralPre"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择上期大集中数据"))?;
        for (label, path) in [("本期数据", &current_file), ("上期数据", &previous_file)] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}文件不存在：{}", path.display()),
                ));
            }
            let supported = path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("csv")
                        || extension.eq_ignore_ascii_case("xlsx")
                        || extension.eq_ignore_ascii_case("xls")
                });
            if !supported {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}只支持 CSV、XLSX、XLS：{}", path.display()),
                ));
            }
        }

        let common_config = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置（3.0）")
            })?;
        let feature_config = ui.value["centralComparisonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中执行比较配置（3.1）")
            })?;
        for (label, path) in [
            ("大集中通用配置（3.0）", &common_config),
            ("大集中执行比较配置（3.1）", &feature_config),
        ] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}目前只支持 .xlsx"),
                ));
            }
        }
        let schema = match ui.value["centralExpressionSchema"].as_str() {
            Some("FIVE_SEGMENT_V1") | None | Some("") => S3ExpressionSchema::FiveSegment,
            Some("LEGACY_8") => S3ExpressionSchema::LegacyEightSegment,
            Some(value) => {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("不支持的大集中表达式语法：{value}"),
                ));
            }
        };
        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| current_file.parent().map(|parent| parent.join("执行结果")))
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (
            current_file,
            previous_file,
            common_config,
            feature_config,
            output_directory,
            schema,
        )
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("大集中比较输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建大集中比较输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析大集中比较输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("比较结果_{stamp}.xlsx"));
    if output_workbook.exists() {
        let suffix = chrono::Local::now().timestamp_subsec_millis();
        output_workbook = output_directory.join(format!("比较结果_{stamp}_{suffix:03}.xlsx"));
        if output_workbook.exists() {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "大集中比较时间戳输出仍有同名文件，请稍后重试",
            ));
        }
    }

    let input_paths = [current_file.clone(), previous_file.clone()];
    let input_reader = input_reader_for_paths(&input_paths);
    let schema_name = match schema {
        S3ExpressionSchema::FiveSegment => "five-segment",
        S3ExpressionSchema::LegacyEightSegment => "legacy-eight-segment",
    };
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行大集中比较",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在执行大集中比较，请勿关闭窗口……");
        push_log(&mut ui.value, "S3-F01 开始大集中执行比较".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F01 开始：input_reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=Rust local ({schema_name})；本期 {}；上期 {}；3.0配置 {}；3.1配置 {}；输出 {}",
                current_file.display(),
                previous_file.display(),
                common_config.display(),
                feature_config.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-execution-compare".to_owned())
        .spawn(move || {
            run_s3_execution_compare_worker(
                worker_state,
                current_file,
                previous_file,
                common_config,
                feature_config,
                output_workbook,
                schema,
                input_reader,
                schema_name,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("大集中执行比较失败");
        push_log(&mut ui.value, format!("启动 S3-F01 比较线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F01 比较线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

#[allow(clippy::too_many_arguments)]
fn run_s3_execution_compare_worker(
    state: AppState,
    current_file: PathBuf,
    previous_file: PathBuf,
    common_config: PathBuf,
    feature_config: PathBuf,
    output_workbook: PathBuf,
    schema: S3ExpressionSchema,
    input_reader: String,
    schema_name: &'static str,
) {
    let started = Instant::now();
    let current_paths = [current_file];
    let previous_paths = [previous_file];
    let progress_state = state.clone();
    let result = run_execution_compare_files(
        &current_paths,
        &previous_paths,
        &common_config,
        &feature_config,
        &output_workbook,
        schema,
        |step| {
            let mut ui = progress_state.lock();
            ui.value["status"] = json!(format!("S3-F01：{step}"));
            push_log(&mut ui.value, format!("S3-F01：{step}"));
            push_detail_log(
                &mut ui.value,
                format!("S3-F01 阶段：{step}；expression={schema_name}"),
            );
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_path.to_string_lossy().into_owned();
            ui.value["centralCompareFile"] = json!(output);
            if ui.value["centralFormOutputAuto"] == true {
                if let Some(parent) = report.output_path.parent() {
                    ui.value["centralFormOutput"] = json!(parent.to_string_lossy());
                }
            }
            ui.value["status"] = json!(format!("执行比较完成：{} 行", report.rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F01 执行比较完成：{} 行，结果文件 {}",
                    report.rows,
                    report.output_path.display()
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F01 完成：reader={}/{}, writer={}，本期/上期记录={}/{}，结果行={}，动作规则触发={}/{}/{}，表达式规则触发={}/{}，表达式错误={}，配置读取={} ms，输入读取={} ms，规则计算={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.config_reader,
                    report.input_reader,
                    report.writer,
                    report.current_records,
                    report.previous_records,
                    report.rows,
                    report.triggered_action_rules,
                    report.applicable_action_rules,
                    report.enabled_action_rules,
                    report.triggered_expression_rules,
                    report.enabled_expression_rules,
                    report.expression_errors,
                    report.config_read_time_ms,
                    report.input_read_time_ms,
                    report.calculation_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("大集中执行比较失败");
            push_log(
                &mut ui.value,
                format!("S3-F01 执行比较失败，未报告成功结果：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F01 失败：reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=Rust local ({schema_name})，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn start_central_explanation_export(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (comparison_file, config_workbook, output_directory) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再导出指标说明",
            ));
        }
        let comparison_file = ui.value["centralCompareFile"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择比较结果.xlsx"))?;
        if !comparison_file.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("比较结果文件不存在：{}", comparison_file.display()),
            ));
        }
        let config_workbook = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from);
        let config_workbook = config_workbook
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置.xlsx"))?;
        if !config_workbook.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("大集中通用配置不存在：{}", config_workbook.display()),
            ));
        }
        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .or_else(|| {
                ui.value["centralFormOutput"]
                    .as_str()
                    .filter(|path| !path.trim().is_empty())
            })
            .map(PathBuf::from)
            .or_else(|| comparison_file.parent().map(Path::to_path_buf))
            .map(|path| {
                if ui.value["centralOutput"]
                    .as_str()
                    .is_some_and(|value| !value.is_empty())
                    || ui.value["centralFormOutput"]
                        .as_str()
                        .is_some_and(|value| !value.is_empty())
                {
                    path
                } else {
                    path.join("执行结果")
                }
            })
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (comparison_file, config_workbook, output_directory)
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("指标说明输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建指标说明输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_workbook = output_directory.join(format!(
        "指标说明_{}.xlsx",
        chrono::Local::now().format("%Y%m%d_%H%M%S")
    ));
    if output_workbook.exists() {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            format!(
                "说明文件输出路径已存在，未覆盖：{}",
                output_workbook.display()
            ),
        ));
    }

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再导出指标说明",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在导出选定指标说明……");
        push_log(&mut ui.value, "S3-F02 开始导出选定指标说明".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F02 输入比较结果 {}；读取引擎 Calamine；输出引擎 rust_xlsxwriter；大集中通用配置 {}；输出 {}",
                comparison_file.display(),
                config_workbook.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    let thread_output = output_workbook.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-explanation-export".to_owned())
        .spawn(move || {
            run_s3_explanation_export_worker(
                worker_state,
                comparison_file,
                Some(config_workbook),
                thread_output,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("指标说明导出失败");
        push_log(&mut ui.value, format!("启动 S3-F02 导出线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F02 导出线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s3_explanation_export_worker(
    state: AppState,
    comparison_file: PathBuf,
    config_workbook: Option<PathBuf>,
    output_workbook: PathBuf,
) {
    let started = Instant::now();
    let result = export_selected_explanations(
        &comparison_file,
        &output_workbook,
        config_workbook.as_deref(),
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_workbook.to_string_lossy().into_owned();
            ui.value["centralExplainFile"] = json!(output);
            ui.value["status"] = json!(format!("指标说明导出完成：{} 条", report.selected_rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F02 指标说明导出完成：{} 条，文件 {}",
                    report.selected_rows,
                    report.output_workbook.display()
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F02 完成：reader={}，writer={}，选中行={}，比较读取={} ms，配置读取={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.reader,
                    report.writer,
                    report.selected_rows,
                    report.comparison_read_time_ms,
                    report.config_read_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("指标说明导出失败");
            push_log(&mut ui.value, format!("S3-F02 指标说明导出失败：{error:#}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F02 失败：reader=Calamine，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn start_central_explanation_feedback_import(
    state: &AppState,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    let (target_workbook, feedback_workbooks) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请稍后再导入 S3-F03 机构反馈",
            ));
        }
        let target_workbook = ui.value["centralExplainFile"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择说明文件.xlsx"))?;
        if !target_workbook.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("目标说明文件不存在：{}", target_workbook.display()),
            ));
        }
        let selected = args.first().and_then(Value::as_array).ok_or_else(|| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                "请从文件选择器批量选择机构反馈工作簿",
            )
        })?;
        if selected.is_empty() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "请至少选择一个机构反馈工作簿",
            ));
        }
        let mut feedback_workbooks = Vec::with_capacity(selected.len());
        for value in selected {
            let path = value
                .as_str()
                .filter(|path| !path.trim().is_empty())
                .map(PathBuf::from)
                .ok_or_else(|| {
                    ApiError::new(StatusCode::BAD_REQUEST, "反馈工作簿路径必须是非空文本")
                })?;
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("机构反馈工作簿不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                })
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("机构反馈工作簿只支持 XLSX/XLSM：{}", path.display()),
                ));
            }
            feedback_workbooks.push(path);
        }
        (target_workbook, feedback_workbooks)
    };
    let modifier = s3_production_modifier("S3-F03")?;

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请稍后再导入 S3-F03 机构反馈",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在匹配并写回机构说明……");
        push_log(
            &mut ui.value,
            format!(
                "S3-F03 开始导入机构反馈：{} 个文件；目标说明文件 {}",
                feedback_workbooks.len(),
                target_workbook.display()
            ),
        );
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F03 将先在目标文件同目录备份，再原地更新目标说明文件；反馈输入保持只读；Excel-Core Passed 后端注册数={}",
                crate::excel_route::registered_backend_count(
                    &crate::excel_route::runtime_platform()
                )
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-explanation-feedback-import".to_owned())
        .spawn(move || {
            run_s3_explanation_feedback_worker(
                worker_state,
                target_workbook,
                feedback_workbooks,
                modifier,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S3-F03 机构反馈导入启动失败");
        push_log(
            &mut ui.value,
            format!("启动 S3-F03 反馈导入线程失败：{error}"),
        );
        return Err(ApiError::internal(format!(
            "启动 S3-F03 反馈导入线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s3_explanation_feedback_worker(
    state: AppState,
    target_workbook: PathBuf,
    feedback_workbooks: Vec<PathBuf>,
    modifier: ExcelModifier,
) {
    let started = Instant::now();
    let result =
        execute_explanation_feedback_import(&target_workbook, &feedback_workbooks, &modifier);
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_workbook.to_string_lossy().into_owned();
            ui.value["centralExplainFile"] = json!(output);
            ui.value["status"] = json!(format!(
                "S3-F03 完成：写入 {} 条，未匹配 {} 条，冲突 {} 条",
                report.imported, report.unmatched, report.conflicts
            ));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F03 机构说明写回完成：写入 {} 条，未匹配 {} 条，冲突 {} 条；目标文件已原地更新：{}",
                    report.imported,
                    report.unmatched,
                    report.conflicts,
                    report.output_workbook.display()
                ),
            );
            if let Some(path) = &report.backup_workbook {
                push_log(
                    &mut ui.value,
                    format!("S3-F03 更新前备份：{}", path.display()),
                );
            }
            if let Some(warning) = &report.rollback_cleanup_warning {
                push_log(&mut ui.value, warning.clone());
            }
            if let Some(path) = &report.conflict_report_path {
                push_log(
                    &mut ui.value,
                    format!("S3-F03 冲突明细文件：{}", path.display()),
                );
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F03 完成：reader=calamine；写回 Excel-Core={}；无待更新单元格={}；目标读取={} ms；反馈读取={} ms；写回={} ms；任务总计={} ms；目标={}",
                    report.edit.as_ref().map(|edit| edit.engine.to_string()).unwrap_or_else(|| "无待更新单元格".to_owned()),
                    report.copied_without_edits,
                    report.target_read_time_ms,
                    report.feedback_read_time_ms,
                    report.writeback_time_ms,
                    started.elapsed().as_millis(),
                    report.output_workbook.display()
                ),
            );
            if let Some(path) = &report.backup_workbook {
                push_detail_log(
                    &mut ui.value,
                    format!("S3-F03 备份路径：{}", path.display()),
                );
            }
        }
        Err(error) => {
            ui.value["status"] = json!("S3-F03 机构反馈导入失败");
            push_log(
                &mut ui.value,
                format!("S3-F03 未发布反馈写回结果：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F03 失败，未报告成功；反馈输入保持不变；若更新阶段需要恢复，错误信息包含备份/恢复路径；任务总计={} ms；错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

/// 在不持有外部锁保护的前提下向状态日志追加一条消息（内部短暂加锁）。
fn push_log_for_state(state: &AppState, message: &str) {
    let mut ui = state.lock();
    push_log(&mut ui.value, message.to_owned());
}

/// 将当前状态中的用户偏好与路径快照写入用户级设置文件。
/// 供 choose_*/update 等改状态入口在成功后调用；失败时返回明确错误，
/// 由前端提示，不静默丢弃用户选择。
/// 注意：调用方必须已释放 `state.lock()` 守卫（本函数内部会重新加锁）。
fn persist_state_preferences(state: &AppState) -> std::result::Result<(), ApiError> {
    let (calc_engine, edit_engine, snapshot) = {
        let ui = state.lock();
        (
            ui.value["advancedSettings"]["calcEngine"]
                .as_str()
                .unwrap_or("auto")
                .to_owned(),
            ui.value["advancedSettings"]["editEngine"]
                .as_str()
                .unwrap_or("umya")
                .to_owned(),
            ui.value.clone(),
        )
    };
    let preferences: std::collections::BTreeMap<String, Value> = snapshot
        .as_object()
        .map(|object| {
            object
                .iter()
                .filter(|(key, _)| crate::app_settings::preference_spec(key).is_some())
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect()
        })
        .unwrap_or_default();
    crate::app_settings::AppSettings::store_preferences(
        &calc_engine,
        &edit_engine,
        preferences,
    )
    .map(|_| ())
    .map_err(|error| ApiError::internal(format!("保存用户设置失败：{error:#}")))
}

fn s3_production_modifier(feature: &str) -> std::result::Result<ExcelModifier, ApiError> {
    // Read the persisted 高级设置 selection; background workers see the last
    // saved value even while the GUI state lock is held by a running task.
    let edit_engine = crate::app_settings::AppSettings::load()
        .edit_engine
        .as_str()
        .to_owned();
    let (modifier, capability) =
        crate::excel_route::production_modifier_with_edit_engine(feature, &edit_engine).map_err(
            |error| ApiError::internal(format!("初始化 {feature} Excel-Core 路由失败：{error}")),
        )?;
    if capability.is_none() {
        return Err(ApiError::new(
            StatusCode::INTERNAL_SERVER_ERROR,
            format!("{feature} 当前平台没有已登记的 Excel-Core Passed 后端，编辑批次保持失败关闭"),
        ));
    }
    Ok(modifier)
}



fn start_central_indicator_split(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (comparison_file, common_config, forms_config, output_directory) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再生成 S3-F05 金融表单",
            ));
        }
        let comparison_file = ui.value["centralCompareFile"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择比较结果.xlsx"))?;
        let common_config = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择 3.0 大集中通用配置.xlsx")
            })?;
        let forms_config = ui.value["centralFormsConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "请先选择 3.3 大集中指标比较拆分配置.xlsx",
                )
            })?;
        let output_directory = ui.value["centralFormOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| comparison_file.parent().map(Path::to_path_buf))
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择表单输出目录"))?;
        for (label, path) in [
            ("比较结果", &comparison_file),
            ("3.0 大集中通用配置", &common_config),
            ("3.3 指标比较拆分配置", &forms_config),
        ] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}文件不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}目前只支持 XLSX：{}", path.display()),
                ));
            }
        }
        if output_directory.exists() && !output_directory.is_dir() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("金融表单输出位置不是目录：{}", output_directory.display()),
            ));
        }
        (
            comparison_file,
            common_config,
            forms_config,
            output_directory,
        )
    };

    let modifier = s3_production_modifier("S3-F05")?;

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请稍后再执行 S3-F05",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["centralFormOutputLastResult"] = json!("");
        ui.value["status"] = json!("正在生成 S3-F05 金融表单……");
        push_log(&mut ui.value, "S3-F05 开始生成机构金融表单".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F05 输入比较结果 {}；3.0 通用配置 {}；3.3 表单配置 {}；输出根目录 {}；Excel-Core Passed 后端注册数={}",
                comparison_file.display(),
                common_config.display(),
                forms_config.display(),
                output_directory.display(),
                crate::excel_route::registered_backend_count(
                    &crate::excel_route::runtime_platform()
                )
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-indicator-split".to_owned())
        .spawn(move || {
            run_s3_indicator_split_worker(
                worker_state,
                comparison_file,
                common_config,
                forms_config,
                output_directory,
                modifier,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S3-F05 任务启动失败");
        push_log(&mut ui.value, format!("启动 S3-F05 任务线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F05 任务线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn check_central_forms_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let path = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能检查 S3-F05 配置",
            ));
        }
        ui.value["centralFormsConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "请先选择 3.3 大集中指标比较拆分配置.xlsx",
                )
            })?
    };
    let started = Instant::now();
    let template =
        crate::s3_indicator_split::load_indicator_split_template(&path).map_err(|error| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("S3-F05 3.3 配置检查失败：{error:#}"),
            )
        })?;
    let enabled_forms = template.forms.len();
    let disabled_forms = template.disabled_form_codes.len();
    let sheet_count = template.sheets.len();
    let report_text = format!(
        "3.3 大集中指标比较拆分配置检查通过\n\n配置文件：{}\n启用报表：{}\n停用报表：{}\n表单工作表：{}\n隐藏空行：{}\n空表删除：{}\n\n此检查只验证配置和内嵌表单可读取，不代表 Excel-Core 已具备生成能力。",
        path.display(),
        enabled_forms,
        disabled_forms,
        sheet_count,
        if template.options.hide_empty_rows { "是" } else { "否" },
        if template.options.delete_empty_sheets { "是" } else { "否" },
    );
    let elapsed_ms = started.elapsed().as_millis();
    let mut ui = state.lock();
    push_log(
        &mut ui.value,
        format!(
            "S3-F05 3.3 配置检查通过：启用报表 {}，停用报表 {}，表单工作表 {}",
            enabled_forms, disabled_forms, sheet_count
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S3-F05 配置读取与结构检查耗时 {elapsed_ms} ms；文件 {}",
            path.display()
        ),
    );
    Ok(json!({"passed": true, "report_text": report_text}))
}

fn central_config_path(
    state: &AppState,
    busy_label: &str,
    field: &str,
    missing_label: &str,
) -> std::result::Result<PathBuf, ApiError> {
    let ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            format!("任务正在运行，暂不能检查{busy_label}配置"),
        ));
    }
    ui.value[field]
        .as_str()
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, format!("请先选择{missing_label}")))
}

fn check_central_common_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let path = central_config_path(
        state,
        " S3-F01 3.0 通用",
        "centralCommonConfig",
        "3.0 大集中通用配置.xlsx",
    )?;
    let started = Instant::now();
    let alert_bands = crate::s3_execution_compare::load_indicator_split_alert_bands(&path)
        .map_err(|error| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("S3-F01 3.0 通用配置检查失败：{error:#}"),
            )
        })?;
    let report_text = format!(
        "3.0 大集中通用配置检查通过\n\n配置文件：{}\n环比警戒区间：{} 条\n机构地区参照等辅助工作表随配置一起读取。\n\n此检查只验证 3.0 配置可读取、警戒区间结构完整，不代表已选择 3.1 配置或已执行比较。",
        path.display(),
        alert_bands.len(),
    );
    let elapsed_ms = started.elapsed().as_millis();
    let mut ui = state.lock();
    push_log(
        &mut ui.value,
        format!(
            "S3-F01 3.0 配置检查通过：环比警戒区间 {} 条",
            alert_bands.len()
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S3-F01 3.0 配置检查耗时 {elapsed_ms} ms；文件 {}",
            path.display()
        ),
    );
    Ok(json!({"passed": true, "report_text": report_text}))
}

fn check_central_expression_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let common_path = central_config_path(
        state,
        " S3-F01 3.0",
        "centralCommonConfig",
        "3.0 大集中通用配置.xlsx",
    )?;
    let feature_path = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能检查 S3-F01 3.1 配置",
            ));
        }
        ui.value["centralComparisonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "请先选择 3.1 大集中执行比较配置.xlsx",
                )
            })?
    };
    let started = Instant::now();
    let config =
        crate::s3_execution_compare::load_execution_compare_config(&common_path, &feature_path)
            .map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("S3-F01 3.1 执行比较配置检查失败：{error:#}"),
                )
            })?;
    let action_rule_count = config.action_rules.len();
    let expression_rule_count = config.expression_rules.len();
    let report_text = format!(
        "3.1 大集中执行比较配置检查通过\n\n配置文件：{}\n联动 3.0 通用配置：{}\n动作规则：{} 条\n表达式规则：{} 条（含兼容八段与五段两种结构）\n豁免指标：{} 项\n\n此检查只验证配置与 3.0 可联合读取、表达式结构可解析，不代表已执行比较。",
        feature_path.display(),
        common_path.display(),
        action_rule_count,
        expression_rule_count,
        config.exempt_indicators.len(),
    );
    let elapsed_ms = started.elapsed().as_millis();
    let mut ui = state.lock();
    push_log(
        &mut ui.value,
        format!(
            "S3-F01 3.1 配置检查通过：动作规则 {}，表达式规则 {}",
            action_rule_count, expression_rule_count
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S3-F01 3.1 配置检查耗时 {elapsed_ms} ms；文件 {}",
            feature_path.display()
        ),
    );
    Ok(json!({"passed": true, "report_text": report_text}))
}

fn check_central_cross_formulas(state: &AppState) -> std::result::Result<Value, ApiError> {
    let path = central_config_path(
        state,
        " S3-F04 3.2",
        "centralCrossConfig",
        "3.2 大集中本期数值核对配置.xlsx",
    )?;
    let started = Instant::now();
    let config = crate::s3_current_value_check::load_current_value_check_config(&[path.clone()])
        .map_err(|error| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("S3-F04 3.2 配置检查失败：{error:#}"),
            )
        })?;
    let enabled = config
        .rules
        .iter()
        .filter(|rule| rule.disabled.is_empty())
        .count();
    let disabled = config.rules.len() - enabled;
    let report_text = format!(
        "3.2 大集中本期数值核对配置检查通过\n\n配置文件：{}\n核对规则：{} 条（启用 {}，停用 {}）\n单位换算：{} → {}\n豁免指标：{} 项\n\n此检查只验证配置可读取、规则结构完整，不代表已选择本期数据或已执行核对。",
        path.display(),
        config.rules.len(),
        enabled,
        disabled,
        config.source_unit,
        config.target_unit,
        config.exempt_indicators.len(),
    );
    let elapsed_ms = started.elapsed().as_millis();
    let mut ui = state.lock();
    push_log(
        &mut ui.value,
        format!(
            "S3-F04 3.2 配置检查通过：核对规则 {}（启用 {}）",
            config.rules.len(),
            enabled
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S3-F04 3.2 配置检查耗时 {elapsed_ms} ms；文件 {}",
            path.display()
        ),
    );
    Ok(json!({"passed": true, "report_text": report_text}))
}

fn run_s3_indicator_split_worker(
    state: AppState,
    comparison_file: PathBuf,
    common_config: PathBuf,
    forms_config: PathBuf,
    output_directory: PathBuf,
    modifier: ExcelModifier,
) {
    let started = Instant::now();
    let result = execute_indicator_comparison_split(
        &comparison_file,
        &forms_config,
        &common_config,
        &output_directory,
        &modifier,
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_directory.to_string_lossy().into_owned();
            ui.value["centralFormOutputLastResult"] = json!(output);
            ui.value["status"] = json!(format!(
                "S3-F05 完成：{} 家机构，{} 个指标匹配",
                report.institution_count, report.matched_indicators
            ));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F05 金融表单生成完成：{} 家机构，输出目录 {}",
                    report.institution_count,
                    report.output_directory.display()
                ),
            );
            for output_file in &report.output_files {
                push_log(
                    &mut ui.value,
                    format!("S3-F05 输出文件：{}", output_file.display()),
                );
            }
            for warning in &report.warnings {
                push_log(&mut ui.value, format!("S3-F05 提示：{warning}"));
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F05 完成：比较行={}；机构={}；匹配指标={}；Excel-Core 编辑器数={}；读取={} ms；规划={} ms；编辑={} ms；任务总计={} ms；输出目录={}",
                    report.comparison_rows,
                    report.institution_count,
                    report.matched_indicators,
                    report.edits.len(),
                    report.read_time_ms,
                    report.planning_time_ms,
                    report.editing_time_ms,
                    started.elapsed().as_millis(),
                    report.output_directory.display()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("S3-F05 失败");
            push_log(&mut ui.value, format!("S3-F05 未生成金融表单：{error:#}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F05 失败；未发布输出；任务总计={} ms；错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn start_central_cross_period(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (current_file, config_workbooks, output_directory, write_flow_logs) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行本期数值核对",
            ));
        }
        let current_file = ui.value["centralCrossCur"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期数值核对数据文件")
            })?;
        if !current_file.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("本期数值核对文件不存在：{}", current_file.display()),
            ));
        }
        if !current_file
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            })
        {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "S3-F04 只支持 CSV、XLSX、XLS 输入",
            ));
        }

        let common_config = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置（3.0）")
            })?;
        let cross_config = ui.value["centralCrossConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期数值核对配置（3.2）")
            })?;
        for (label, path) in [
            ("大集中通用配置", &common_config),
            ("本期数值核对配置", &cross_config),
        ] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}目前只支持 .xlsx"),
                ));
            }
        }

        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| current_file.parent().map(|parent| parent.join("执行结果")))
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (
            current_file,
            vec![common_config, cross_config],
            output_directory,
            ui.value["exportRunLogs"] == true,
        )
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "本期数值核对输出位置不是目录：{}",
                output_directory.display()
            ),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建本期数值核对输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析本期数值核对输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("数值核对结果_{stamp}.xlsx"));
    let mut log_workbook = write_flow_logs
        .then(|| output_directory.join(format!("本期数值核对_运行日志_{stamp}.xlsx")));
    if output_workbook.exists() || log_workbook.as_ref().is_some_and(|path| path.exists()) {
        let suffix = chrono::Local::now().timestamp_subsec_millis();
        output_workbook = output_directory.join(format!("数值核对结果_{stamp}_{suffix:03}.xlsx"));
        log_workbook = write_flow_logs.then(|| {
            output_directory.join(format!("本期数值核对_运行日志_{stamp}_{suffix:03}.xlsx"))
        });
        if output_workbook.exists() || log_workbook.as_ref().is_some_and(|path| path.exists()) {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "本期数值核对时间戳输出仍有同名文件，请稍后重试",
            ));
        }
    }

    let input_reader = input_reader_for_paths(std::slice::from_ref(&current_file));
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行本期数值核对",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在执行本期数值核对，请勿关闭窗口……");
        push_log(&mut ui.value, "S3-F04 开始本期数值核对".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F04 开始：reader={}，writer=rust_xlsxwriter；当前数据 {}；通用配置 {}；本期核对配置 {}；结果输出 {}；运行日志={}",
                input_reader,
                current_file.display(),
                config_workbooks[0].display(),
                config_workbooks[1].display(),
                output_workbook.display(),
                log_workbook.is_some()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-current-value-check".to_owned())
        .spawn(move || {
            run_s3_current_value_check_worker(
                worker_state,
                current_file,
                config_workbooks,
                output_workbook,
                log_workbook,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("本期数值核对失败");
        push_log(&mut ui.value, format!("启动 S3-F04 任务线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F04 任务线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s3_current_value_check_worker(
    state: AppState,
    current_file: PathBuf,
    config_workbooks: Vec<PathBuf>,
    output_workbook: PathBuf,
    log_workbook: Option<PathBuf>,
) {
    let started = Instant::now();
    let input_reader = input_reader_for_paths(std::slice::from_ref(&current_file));
    let result = run_current_value_check_files(
        std::slice::from_ref(&current_file),
        &[],
        &config_workbooks,
        &output_workbook,
        log_workbook.as_deref(),
        python_regex_scope_matcher(),
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_workbook.to_string_lossy().into_owned();
            ui.value["centralCrossLastResult"] = json!(output);
            ui.value["status"] = json!(format!("本期数值核对完成：{} 行", report.rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F04 本期数值核对完成：{} 行，结果文件 {}",
                    report.rows,
                    report.output_workbook.display()
                ),
            );
            if let Some(log_path) = report.log_workbook.as_ref() {
                push_log(
                    &mut ui.value,
                    format!("S3-F04 运行日志文件：{}", log_path.display()),
                );
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F04 完成：reader={}/{}, writer={}，规则={}/{}/{}，输入记录={}/{}，结果行={}，配置读取={} ms，输入读取={} ms，规则计算={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.config_reader,
                    report.input_reader,
                    report.writer,
                    report.enabled_rules,
                    report.total_rules,
                    report.disabled_rules,
                    report.current_records,
                    report.previous_records,
                    report.rows,
                    report.config_read_time_ms,
                    report.input_read_time_ms,
                    report.calculation_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("本期数值核对失败");
            push_log(&mut ui.value, format!("S3-F04 本期数值核对失败：{error:#}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F04 失败：reader={input_reader}，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn toggle_central_output_pin(
    state: &AppState,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "缺少大集中输出目录类型"))?;
    let (auto_key, source_key) = match kind {
        "centralOutput" => ("centralOutputAuto", "centralCur"),
        "centralFormOutput" => ("centralFormOutputAuto", "centralCompareFile"),
        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("Rust 迁移版暂不支持固定该输出目录：{kind}"),
            ));
        }
    };
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改输出目录",
        ));
    }
    let auto = ui.value[auto_key] != true;
    ui.value[auto_key] = json!(auto);
    if auto {
        if let Some(source) = ui.value[source_key]
            .as_str()
            .filter(|path| !path.is_empty())
        {
            if let Some(parent) = Path::new(source).parent() {
                ui.value[kind] = json!(parent.join("执行结果").to_string_lossy());
            }
        }
    }
    Ok(json!(!auto))
}

fn inspect_template_formulas(state: &AppState) -> std::result::Result<Value, ApiError> {
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，请完成后再检查模板公式",
        ));
    }
    let template = ui.value["template"]
        .as_str()
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择或自动识别匹配模板"))?;
    let started = Instant::now();
    let (report, read_elapsed_ms) = inspect_s1_template(&template).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("模板公式检查失败：{error:#}"),
        )
    })?;
    let formula_count = report.formula_count;
    let error_count = report.errors.len();
    let ranges = report
        .named_range_groups
        .iter()
        .filter(|group| group.count > 0)
        .map(|group| format!("{} {}个", group.family, group.count))
        .collect::<Vec<_>>()
        .join("、");
    let other_count = report.other_names.len();
    let internal_count = report.internal_names.len();
    let value = serde_json::to_value(report)
        .map_err(|error| ApiError::internal(format!("序列化模板检查结果失败：{error}")))?;
    push_log(
        &mut ui.value,
        format!(
            "检查模板：共 {formula_count} 个公式；明确错误 {error_count} 处；命名区域 {}；其他 {other_count} 个；内部名称 {internal_count} 个",
            if ranges.is_empty() { "无" } else { &ranges }
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S1-F01 模板检查步骤：读取公式文本与命名区域；引擎 {INSPECTION_ENGINE}；读取耗时 {read_elapsed_ms} ms；总耗时 {} ms",
            started.elapsed().as_millis()
        ),
    );
    Ok(value)
}

fn clear_path(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    if let Some(
        field @ ("centralCompareFile"
        | "centralExplainFile"
        | "centralCur"
        | "centralPre"
        | "centralCommonConfig"
        | "centralComparisonConfig"
        | "centralFormsConfig"
        | "centralCrossCur"
        | "centralCrossConfig"
        | "centralOutput"
        | "centralFormOutput"),
    ) = arg_string(args, 0)
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能清除大集中路径",
            ));
        }
        ui.value[field] = json!("");
        push_log(&mut ui.value, format!("S3 路径已清除：{field}"));
        return Ok(json!(true));
    }
    if arg_string(args, 0) == Some("template") {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能清除路径",
            ));
        }
        ui.value["template"] = json!("");
        ui.value["templateManual"] = json!(false);
        push_log(&mut ui.value, "S1 模板：已清除模板选择".to_owned());
        return Ok(json!(true));
    }
    if let Some(field @ ("input" | "templateDir" | "external" | "output" | "historyConfig")) =
        arg_string(args, 0)
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能清除路径",
            ));
        }
        ui.value[field] = json!("");
        if field == "input" && ui.value["outputPinned"] != true {
            ui.value["output"] = json!("");
        } else if field == "output" {
            ui.value["outputPinned"] = json!(false);
            let input = ui.value["input"].as_str().unwrap_or_default().to_owned();
            ui.value["output"] = json!(s1_output_directory_following_input(&input));
        }
        if field == "historyConfig" {
            push_log(&mut ui.value, "逐笔统计系统历史配置路径已清除".to_owned());
        }
        return Ok(json!(true));
    }
    clear_period_path(state, args)
}
fn check_period_audit_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let config_path = state.lock().value["pcConfig"]
        .as_str()
        .unwrap_or_default()
        .to_owned();
    let started = Instant::now();
    let report = check_s2_config(config_path);
    let passed = report.passed;
    let blocking_errors = report.blocking_errors;
    let warnings = report.warnings;
    let sheet_count = report.sheet_stats.len();
    let indicator_count = report.indicator_count;
    let strategy_count = report.strategy_count;
    let elapsed_ms = started.elapsed().as_millis();
    let value = serde_json::to_value(report)
        .map_err(|error| ApiError::internal(format!("序列化配置检查结果失败：{error}")))?;
    let mut ui = state.lock();
    ui.value["periodConfigCheck"] = value.clone();
    push_log(
        &mut ui.value,
        format!(
            "S2 配置检查{}：{} 项阻断错误、{} 项警告；耗时 {} ms",
            if passed { "通过" } else { "未通过" },
            blocking_errors,
            warnings,
            elapsed_ms
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S2-F01 配置检查步骤：reader=calamine；工作表={sheet_count}；指标={indicator_count}；环比策略={strategy_count}；blocking_errors={blocking_errors}；warnings={warnings}；elapsed_ms={elapsed_ms}"
        ),
    );
    Ok(value)
}

fn choose_period_compare(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0).ok_or_else(|| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            "choose_period_compare 缺少路径类型",
        )
    })?;
    if !matches!(
        kind,
        "pcCurDir" | "pcPreDir" | "pcCentral" | "pcConfig" | "pcOutput"
    ) {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("不支持的报表采集路径类型：{kind}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim().to_owned();
    if selected.is_empty() {
        return Ok(state.lock().value[kind].clone());
    }
    {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能更改路径",
            ));
        }
    }
    let selected_path = PathBuf::from(unquote_path(&selected));
    let is_file = matches!(kind, "pcCentral" | "pcConfig");
    if (is_file && !selected_path.is_file()) || (!is_file && !selected_path.is_dir()) {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "所选{}不存在或类型不正确：{}",
                if is_file { "文件" } else { "目录" },
                selected_path.display()
            ),
        ));
    }
    let absolute = selected_path
        .canonicalize()
        .unwrap_or(selected_path)
        .to_string_lossy()
        .into_owned();
    let mut ui = state.lock();
    ui.value[kind] = json!(absolute);
    if kind == "pcCurDir" && ui.value["pcOutputAuto"] == true {
        ui.value["pcOutput"] = json!(Path::new(&absolute).join("执行结果").to_string_lossy());
    } else if kind == "pcOutput" {
        ui.value["pcOutputAuto"] = json!(false);
    }
    push_log(
        &mut ui.value,
        format!("报表采集：{}已选择", period_path_label(kind)),
    );
    let chosen = ui.value[kind].clone();
    drop(ui);
    refresh_period_pairs(state);
    if kind == "pcConfig" {
        let _ = check_period_audit_config(state)?;
    }
    persist_state_preferences(state)?;
    Ok(chosen)
}

fn toggle_period_output_pin(state: &AppState) -> std::result::Result<Value, ApiError> {
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Ok(json!(ui.value["pcOutputAuto"] != true));
    }
    let auto = ui.value["pcOutputAuto"] != true;
    ui.value["pcOutputAuto"] = json!(auto);
    if auto {
        if let Some(current) = ui.value["pcCurDir"].as_str() {
            if !current.is_empty() {
                ui.value["pcOutput"] = json!(Path::new(current).join("执行结果").to_string_lossy());
            }
        }
    }
    push_log(
        &mut ui.value,
        if auto {
            "报表采集：输出目录跟随本期目录".to_owned()
        } else {
            "报表采集：输出目录已固定".to_owned()
        },
    );
    Ok(json!(!auto))
}

fn clear_period_path(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "clear_path 缺少路径类型"))?;
    if !matches!(
        kind,
        "pcCurDir" | "pcPreDir" | "pcCentral" | "pcConfig" | "pcOutput"
    ) {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持清除路径：{kind}"),
        ));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能清除路径",
        ));
    }
    ui.value[kind] = json!("");
    if kind == "pcCurDir" && ui.value["pcOutputAuto"] == true {
        ui.value["pcOutput"] = json!("");
    }
    if kind == "pcOutput" {
        ui.value["pcOutputAuto"] = json!(false);
    }
    push_log(
        &mut ui.value,
        format!("报表采集：已清除{}", period_path_label(kind)),
    );
    drop(ui);
    refresh_period_pairs(state);
    if kind == "pcConfig" {
        let _ = check_period_audit_config(state)?;
    }
    Ok(json!(true))
}

fn browse_dir(_state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let requested = arg_string(args, 0).unwrap_or_default();
    let mode = arg_string(args, 1).unwrap_or_default();
    let root = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let cleaned = unquote_path(requested.trim());
    let mut target = if cleaned.is_empty() {
        root.clone()
    } else {
        let candidate = PathBuf::from(cleaned);
        if candidate.is_absolute() {
            candidate
        } else {
            root.join(candidate)
        }
    };
    target = match target.canonicalize() {
        Ok(path) => path,
        Err(_) => {
            return Ok(browse_error(
                &root,
                &root,
                "路径不存在或不是可访问的目录",
                "",
            ));
        }
    };
    let extensions = browse_extensions(mode);
    let mut selected_file = String::new();
    if target.is_file() {
        if is_hidden_name(&target) || !extension_allowed(&target, extensions) {
            return Ok(browse_error(
                &target.parent().unwrap_or(&root),
                &root,
                "该文件不符合当前选择类型",
                "",
            ));
        }
        selected_file = target.to_string_lossy().into_owned();
        target = target.parent().unwrap_or(&root).to_path_buf();
    } else if !target.is_dir() {
        return Ok(browse_error(
            &root,
            &root,
            "路径不存在或不是可访问的目录",
            "",
        ));
    }

    let mut directories = Vec::new();
    let mut files = Vec::new();
    let entries = match fs::read_dir(&target) {
        Ok(entries) => entries,
        Err(error) => {
            return Ok(browse_error(
                &target,
                &root,
                &error.to_string(),
                &selected_file,
            ));
        }
    };
    let mut entries = entries.filter_map(|entry| entry.ok()).collect::<Vec<_>>();
    entries.sort_by_key(|entry| entry.file_name().to_string_lossy().to_lowercase());
    for entry in entries {
        let name = entry.file_name().to_string_lossy().into_owned();
        if is_hidden_name(Path::new(&name)) || name == "__pycache__" {
            continue;
        }
        let Ok(file_type) = entry.file_type() else {
            continue;
        };
        let path = entry.path();
        if file_type.is_dir() {
            directories.push(json!({"name": name, "path": path.to_string_lossy()}));
        } else if file_type.is_file() && extension_allowed(&path, extensions) {
            let size = entry.metadata().map(|metadata| metadata.len()).unwrap_or(0);
            let size_text = if size >= 1024 {
                format!("{} KB", (size + 512) / 1024)
            } else {
                format!("{size} B")
            };
            files.push(json!({"name": name, "path": path.to_string_lossy(), "size": size_text}));
        }
    }
    let parent = target
        .parent()
        .filter(|path| *path != target)
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    Ok(json!({
        "path": target.to_string_lossy(),
        "parent": parent,
        "dirs": directories,
        "files": files,
        "root": root.to_string_lossy(),
        "selectedFile": selected_file,
        "error": ""
    }))
}

fn start_period_compare_v2(state: &AppState) -> std::result::Result<Value, ApiError> {
    let request = {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Ok(json!(false));
        }
        let current = ui.value["pcCurDir"].as_str().unwrap_or_default().to_owned();
        let previous = ui.value["pcPreDir"].as_str().unwrap_or_default().to_owned();
        if current.is_empty() || previous.is_empty() {
            push_log(
                &mut ui.value,
                "跨期比较失败：请先选择本期与上期目录".to_owned(),
            );
            ui.value["status"] = json!("跨期比较失败：请先选择本期与上期目录");
            return Ok(json!(false));
        }
        let current_path = PathBuf::from(&current);
        let previous_path = PathBuf::from(&previous);
        if !current_path.is_dir() || !previous_path.is_dir() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "本期或上期目录已不存在",
            ));
        }
        let config = ui.value["pcConfig"].as_str().unwrap_or_default().to_owned();
        if config.is_empty() || !Path::new(&config).is_file() {
            ui.value["periodConfigCheck"] = serde_json::to_value(check_s2_config(&config))
                .unwrap_or_else(|_| json!({"passed": false, "message": "配置文件不可读取"}));
            ui.value["status"] = json!("跨期比较失败：报表采集系统配置不存在");
            push_log(
                &mut ui.value,
                "跨期比较失败：请先选择有效的报表采集系统配置".to_owned(),
            );
            return Ok(json!(false));
        }
        let config_report = check_s2_config(&config);
        ui.value["periodConfigCheck"] = serde_json::to_value(&config_report)
            .map_err(|error| ApiError::internal(format!("序列化配置检查结果失败：{error}")))?;
        if !config_report.passed {
            ui.value["status"] = json!("跨期比较失败：报表采集系统配置存在阻断错误");
            push_log(&mut ui.value, config_report.message.clone());
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "config", "label": "检查审核配置",
                "percent": 0, "current": 0, "total": 0,
                "message": config_report.message, "elapsedSeconds": 0
            });
            return Ok(json!(false));
        }
        let external = ui.value["pcCentral"]
            .as_str()
            .filter(|path| !path.is_empty())
            .map(str::to_owned);
        if let Some(external) = external.as_deref() {
            if !Path::new(external).is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "大集中数据文件已不存在",
                ));
            }
        }
        let output = ui.value["pcOutput"]
            .as_str()
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
            .unwrap_or_else(|| current_path.join("执行结果"));
        ui.value["busy"] = json!(true);
        ui.value["cancelRequested"] = json!(false);
        ui.value["periodSummary"] = json!({});
        ui.value["status"] = json!("正在检查审核配置……");
        ui.value["periodProgress"] = json!({
            "status": "running", "stage": "config", "label": "检查审核配置",
            "percent": 0, "current": 0, "total": 0,
            "message": "准备执行 Rust S2 审核", "elapsedSeconds": 0
        });
        ui.started_at = Some(Instant::now());
        let request = S2RunRequest {
            current_directory: current_path,
            previous_directory: previous_path,
            config_path: PathBuf::from(config),
            output_directory: output,
            external_workbook: external.map(PathBuf::from),
        };
        push_log(
            &mut ui.value,
            format!(
                "Rust S2 开始：本期 {}；上期 {}；输出 {}",
                request.current_directory.display(),
                request.previous_directory.display(),
                request.output_directory.display()
            ),
        );
        request
    };

    let worker_state = state.clone();
    thread::Builder::new()
        .name("audit-s2-service".to_owned())
        .spawn(move || run_s2_worker(worker_state, request))
        .map_err(|error| {
            let mut ui = state.lock();
            ui.value["busy"] = json!(false);
            ui.value["status"] = json!("无法启动 Rust S2 审核线程");
            ApiError::internal(format!("启动审核线程失败：{error}"))
        })?;
    Ok(json!(true))
}

fn run_s2_worker(state: AppState, request: S2RunRequest) {
    let worker_started = Instant::now();
    let mut active_stage: Option<(String, String, Instant)> = None;
    let mut on_progress = |event: S2ProgressEvent| {
        let stage_changed = active_stage
            .as_ref()
            .map(|(stage, _, _)| stage != &event.stage)
            .unwrap_or(true);
        let mut ui = state.lock();
        if stage_changed {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段完成：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let engine = s2_engine_for_stage(&event.stage).unwrap_or("未映射");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 阶段开始：stage={} label={} engine={} progress={}/{}",
                    event.stage, event.label, engine, event.current, event.total
                ),
            );
            active_stage = Some((event.stage.clone(), engine.to_owned(), Instant::now()));
        }
        let stage_elapsed = active_stage
            .as_ref()
            .map(|(_, _, started)| started.elapsed().as_secs_f64())
            .unwrap_or_default();
        ui.value["status"] = json!(format!("{}……", event.label));
        ui.value["periodProgress"] = progress_value(
            &event,
            "running",
            worker_started.elapsed().as_secs_f64(),
            stage_elapsed,
        );
    };
    let result = run_s2_audit(&request, Some(&mut on_progress));
    drop(on_progress);
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.started_at = None;
    match result {
        Ok(result) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段完成：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let summary = result.analysis.summary;
            let report_path = result.report_path.to_string_lossy().into_owned();
            ui.value["periodSummary"] = json!({
                "result": report_path,
                "institutions": summary.institutions,
                "forms": summary.forms,
                "indicators": summary.indicators,
                "abnormal": summary.abnormal_findings,
                "totalFindings": summary.total_findings,
                "byType": summary.by_type,
                "bySeverity": summary.by_severity
            });
            ui.value["periodProgress"] = json!({
                "status": "completed", "stage": "completed", "label": "审核完成",
                "percent": 100, "current": 1, "total": 1,
                "message": format!("报告已生成：{report_path}"),
                "elapsedSeconds": worker_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较完成：已生成审核结果");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务完成：reader=Calamine writer=rust_xlsxwriter findings={} total_ms={}",
                    summary.total_findings,
                    worker_started.elapsed().as_millis()
                ),
            );
            push_log(&mut ui.value, format!("审核完成：{}", report_path));
        }
        Err(error) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段失败：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let message = format!("跨期比较失败：{error:#}");
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "failed", "label": "审核失败",
                "percent": 0, "current": 0, "total": 0,
                "message": message,
                "elapsedSeconds": worker_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较失败");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务失败：total_ms={} error={error:#}",
                    worker_started.elapsed().as_millis()
                ),
            );
            push_log(&mut ui.value, message);
        }
    }
}

fn progress_value(
    event: &S2ProgressEvent,
    status: &str,
    elapsed: f64,
    stage_elapsed: f64,
) -> Value {
    let engine = s2_engine_for_stage(&event.stage);
    let label = engine.map_or_else(
        || event.label.clone(),
        |engine| {
            format!(
                "{} · {engine} · 本阶段已运行 {:.2} 秒",
                event.label, stage_elapsed
            )
        },
    );
    json!({
        "status": status,
        "stage": event.stage,
        "label": label,
        "engine": engine,
        "percent": event.percent.clamp(0.0, 99.0),
        "current": event.current,
        "total": event.total,
        "message": event.message,
        "elapsedSeconds": elapsed,
        "stageElapsedSeconds": stage_elapsed
    })
}

fn update_state(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(updates) = args.first().and_then(Value::as_object) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "update 需要一个 JSON 对象",
        ));
    };
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能更改设置",
        ));
    }
    let previous_input = ui.value["input"].as_str().unwrap_or_default().to_owned();
    let mut input_changed = false;
    let mut explicit_output = false;
    for (key, value) in updates {
        match key.as_str() {
            "input" | "templateDir" | "template" | "external" | "output" | "historyConfig" => {
                let Some(path) = value.as_str() else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是字符串"),
                    ));
                };
                ui.value[key] = json!(path.trim());
                if key == "input" {
                    input_changed = path.trim() != previous_input;
                } else if key == "output" && !path.trim().is_empty() {
                    explicit_output = true;
                }
                if key == "template" && !path.trim().is_empty() {
                    ui.value["templateManual"] = json!(true);
                }
            }
            "outputPinned" => {
                let Some(pinned) = value.as_bool() else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "outputPinned 必须是布尔值",
                    ));
                };
                ui.value["outputPinned"] = json!(pinned);
            }
            "uiTheme" | "uiAccent" => {
                if !value.is_string() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是字符串"),
                    ));
                }
                ui.value[key] = value.clone();
            }
            "centralExpressionSchema"
                if matches!(value.as_str(), Some("FIVE_SEGMENT_V1" | "LEGACY_8")) =>
            {
                ui.value[key] = value.clone();
            }
            "centralExpressionSchema" => {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "centralExpressionSchema 必须是 FIVE_SEGMENT_V1 或 LEGACY_8",
                ));
            }
            "confirmBeforeRun"
            | "showRunDetailLogs"
            | "showPerformanceDiagnostics"
            | "exportRunLogs" => {
                if !value.is_boolean() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是布尔值"),
                    ));
                }
                ui.value[key] = value.clone();
            }
            "filePickerMode"
                if matches!(
                    value.as_str(),
                    Some("自动") | Some("系统原生") | Some("浏览器内置")
                ) =>
            {
                ui.value[key] = value.clone();
            }
            "filePickerMode" => {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "文件/目录选择方式必须是：自动、系统原生 或 浏览器内置",
                ));
            }
            "advancedSettings" => {
                let Some(selection) = value.as_object() else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "advancedSettings 必须是对象",
                    ));
                };
                let calc = selection
                    .get("calcEngine")
                    .and_then(Value::as_str)
                    .unwrap_or(
                        ui.value["advancedSettings"]["calcEngine"]
                            .as_str()
                            .unwrap_or("auto"),
                    );
                let edit = selection
                    .get("editEngine")
                    .and_then(Value::as_str)
                    .unwrap_or(
                        ui.value["advancedSettings"]["editEngine"]
                            .as_str()
                            .unwrap_or("umya"),
                    );
                if edit == "openpyxl" {
                    let (available, detail) = crate::app_settings::openpyxl_evidence_status();
                    if !available {
                        return Err(ApiError::new(
                            StatusCode::BAD_REQUEST,
                            format!("openpyxl 修改引擎不可用：{detail}"),
                        ));
                    }
                }
                let settings = crate::app_settings::AppSettings::store(calc, edit)
                    .map_err(|error| {
                        ApiError::new(StatusCode::BAD_REQUEST, format!("{error:#}"))
                    })?;
                ui.value["advancedSettings"]["calcEngine"] =
                    json!(settings.calc_engine.as_str());
                ui.value["advancedSettings"]["editEngine"] =
                    json!(settings.edit_engine.as_str());
                push_log(
                    &mut ui.value,
                    format!(
                        "高级设置已更新：计算引擎={}；修改引擎={}；设置文件 {}",
                        settings.calc_engine.label(),
                        settings.edit_engine.label(),
                        crate::app_settings::AppSettings::settings_path_for_display()
                    ),
                );
            }
            _ => {
                return Err(ApiError::new(
                    StatusCode::NOT_IMPLEMENTED,
                    format!("Rust 迁移版暂不支持修改设置：{key}"),
                ));
            }
        }
    }
    // 本轮更新涉及用户偏好（白名单键）时，把最终状态整体快照持久化；
    // 失败时明确报错，不静默丢弃用户修改。此处持有 ui 守卫，因此直接
    // 以 ui.value 构建快照，不得改调内部会重新加锁的 persist_state_preferences。
    let touches_preferences = updates
        .keys()
        .any(|key| crate::app_settings::preference_spec(key).is_some());
    if touches_preferences {
        let calc_engine = ui.value["advancedSettings"]["calcEngine"]
            .as_str()
            .unwrap_or("auto")
            .to_owned();
        let edit_engine = ui.value["advancedSettings"]["editEngine"]
            .as_str()
            .unwrap_or("umya")
            .to_owned();
        let preferences: std::collections::BTreeMap<String, Value> = ui
            .value
            .as_object()
            .map(|object| {
                object
                    .iter()
                    .filter(|(key, _)| crate::app_settings::preference_spec(key).is_some())
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect()
            })
            .unwrap_or_default();
        crate::app_settings::AppSettings::store_preferences(
            &calc_engine,
            &edit_engine,
            preferences,
        )
        .map_err(|error| {
            ApiError::internal(format!("保存用户设置失败：{error:#}"))
        })?;
    }
    let input = ui.value["input"].as_str().unwrap_or_default().to_owned();
    if let Some(pinned) = updates.get("outputPinned").and_then(Value::as_bool) {
        if !pinned {
            ui.value["output"] = json!(s1_output_directory_following_input(&input));
        } else if ui.value["output"].as_str().unwrap_or_default().is_empty() {
            ui.value["output"] = json!(s1_output_directory_following_input(&input));
        }
    } else if explicit_output {
        // Choosing an output folder is an explicit request to keep it fixed.
        ui.value["outputPinned"] = json!(true);
    } else if input_changed && ui.value["outputPinned"] != true {
        ui.value["output"] = json!(s1_output_directory_following_input(&input));
    }
    Ok(ui.value.clone())
}

fn s1_output_directory_following_input(input: &str) -> String {
    if input.trim().is_empty() {
        String::new()
    } else {
        Path::new(input.trim())
            .join("执行结果")
            .to_string_lossy()
            .into_owned()
    }
}

fn browse_extensions(mode: &str) -> &'static [&'static str] {
    match mode {
        "excel" => &[".xlsx", ".xlsm", ".xls"],
        "data" => &[".csv", ".xlsx"],
        "config" => &[".xlsx"],
        "excel_ooxml" => &[".xlsx", ".xlsm"],
        "central_data" => &[".csv", ".xlsx", ".xls"],
        "csv" => &[".csv"],
        "report_data" => &[".xlsx", ".xls"],
        _ => &[],
    }
}

fn refresh_period_pairs(state: &AppState) {
    let (current, previous, old_warning) = {
        let ui = state.lock();
        (
            ui.value["pcCurDir"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            ui.value["pcPreDir"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            ui.value["periodSourceWarning"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
        )
    };
    let warning = report_source_warning(current.as_deref(), previous.as_deref());
    let pairs = list_period_pairs(current.as_deref(), previous.as_deref());
    let mut ui = state.lock();
    ui.value["periodSourceWarning"] = json!(warning);
    match pairs {
        Ok(pairs) => {
            ui.value["periodPairs"] = serde_json::to_value(pairs).unwrap_or_else(|_| json!([]));
        }
        Err(error) => {
            ui.value["periodPairs"] = json!([]);
            push_log(&mut ui.value, format!("两期文件配对失败：{error:#}"));
        }
    }
    if !warning.is_empty() && warning != old_warning {
        push_log(&mut ui.value, warning);
    }
}

fn extension_allowed(path: &Path, extensions: &[&str]) -> bool {
    extensions.is_empty()
        || path
            .file_name()
            .and_then(|name| name.to_str())
            .map(|name| {
                let name = name.to_lowercase();
                extensions.iter().any(|extension| name.ends_with(extension))
            })
            .unwrap_or(false)
}

fn is_hidden_name(path: &Path) -> bool {
    path.file_name()
        .and_then(|name| name.to_str())
        .map(|name| name.starts_with('.') || name.starts_with("~$"))
        .unwrap_or(false)
}

fn browse_error(target: &Path, root: &Path, error: &str, selected_file: &str) -> Value {
    let parent = target
        .parent()
        .filter(|path| *path != target)
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    json!({
        "path": target.to_string_lossy(),
        "parent": parent,
        "dirs": [],
        "files": [],
        "root": root.to_string_lossy(),
        "selectedFile": selected_file,
        "error": error
    })
}

fn unquote_path(value: &str) -> String {
    let mut text = value.trim().to_owned();
    for _ in 0..3 {
        let pairs = [("\"", "\""), ("'", "'"), ("“", "”"), ("‘", "’")];
        let mut changed = false;
        for (open, close) in pairs {
            if text.starts_with(open)
                && text.ends_with(close)
                && text.len() >= open.len() + close.len()
            {
                text = text[open.len()..text.len() - close.len()].trim().to_owned();
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    text
}

fn arg_string(args: &[Value], index: usize) -> Option<&str> {
    args.get(index).and_then(Value::as_str)
}

fn period_path_label(kind: &str) -> &'static str {
    match kind {
        "pcCurDir" => "本期目录",
        "pcPreDir" => "上期目录",
        "pcCentral" => "大集中数据",
        "pcConfig" => "报表采集系统配置",
        "pcOutput" => "输出目录",
        _ => "路径",
    }
}

fn push_detail_log(state: &mut Value, message: String) {
    let Some(log) = state.get_mut("log").and_then(Value::as_array_mut) else {
        return;
    };
    log.push(json!({"text": message, "detail": true}));
    if log.len() > 200 {
        let excess = log.len() - 200;
        log.drain(0..excess);
    }
}
fn push_log(state: &mut Value, message: String) {
    let Some(log) = state.get_mut("log").and_then(Value::as_array_mut) else {
        return;
    };
    log.push(json!({"text": message, "detail": false}));
    if log.len() > 200 {
        let excess = log.len() - 200;
        log.drain(0..excess);
    }
}

#[derive(Debug)]
struct ApiError {
    status: StatusCode,
    message: String,
}

impl ApiError {
    fn new(status: StatusCode, message: impl Into<String>) -> Self {
        Self {
            status,
            message: message.into(),
        }
    }

    fn internal(message: impl Into<String>) -> Self {
        Self::new(StatusCode::INTERNAL_SERVER_ERROR, message)
    }
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        (self.status, Json(json!({"error": self.message}))).into_response()
    }
}

#[cfg(test)]
mod tests {
    use super::{
        apply_s1_f3_recovery_state, browse_dir, cached_s1_f1_template_preflight, dispatch_api,
        execute_s1_f1_workflow_with_dependencies, extract_s1_audit_copy_with_conditional_extractor,
        progress_value, render_index, run_s1_f3_worker_with_store, s1_f1_template_preflight,
        run_s1_f1_for_acceptance,
        write_s1_recovery_manifest_result, AppState, S1F1AcceptedSource, S1F1OutputDirectoryGuard,
        S1F1PreflightRequest,
    };
    use crate::central_csv::EXPECTED_HEADERS;
    use crate::models::Issue;
    use crate::s1_formula_audit::S1FormulaMetadata;
    use crate::s1_formula_copy::{S1FormulaCopyExecutionReport, S1FormulaCopyOutcome};
    use crate::s1_recalculation::S1RecalculationReport;
    use crate::s1_recovery::{
        file_sha256, workbook_recovery_fingerprints, S1RecoveryFile, S1RecoveryManifest,
        S1RecoveryStatus, S1RecoveryStore, S1RecoveryTaskSummary,
    };
    use crate::s2_service::S2ProgressEvent;
    use anyhow::Context;
    use axum::http::StatusCode;
    use excel_writer::Workbook;
    use serde_json::{json, Value};
    use std::{
        fs,
        path::{Path, PathBuf},
        time::{Duration, SystemTime, UNIX_EPOCH},
    };

    fn test_state() -> AppState {
        AppState::new(PathBuf::from("missing-test-config.xlsx"))
    }

    #[test]
    fn s1_output_directory_tracks_input_until_explicitly_pinned() {
        let state = test_state();
        let source_a = std::env::temp_dir().join("rust-audit-s1-source-a");
        let source_b = std::env::temp_dir().join("rust-audit-s1-source-b");
        let explicit_output = std::env::temp_dir().join("rust-audit-s1-fixed-output");
        let source_a_text = source_a.to_string_lossy().into_owned();
        let source_b_text = source_b.to_string_lossy().into_owned();
        let explicit_output_text = explicit_output.to_string_lossy().into_owned();

        dispatch_api(&state, "update", &[json!({"input": source_a_text})])
            .expect("select a source directory");
        assert_eq!(
            state.get_state()["output"],
            json!(source_a.join("执行结果").to_string_lossy())
        );
        assert_eq!(state.get_state()["outputPinned"], false);

        dispatch_api(&state, "update", &[json!({"output": explicit_output_text})])
            .expect("select an explicit output directory");
        assert_eq!(state.get_state()["output"], json!(explicit_output_text));
        assert_eq!(state.get_state()["outputPinned"], true);

        dispatch_api(&state, "update", &[json!({"input": source_b_text})])
            .expect("change the source directory while output is pinned");
        assert_eq!(state.get_state()["output"], json!(explicit_output_text));

        dispatch_api(&state, "update", &[json!({"outputPinned": false})])
            .expect("switch output back to source-following mode");
        assert_eq!(
            state.get_state()["output"],
            json!(source_b.join("执行结果").to_string_lossy())
        );
        assert_eq!(state.get_state()["outputPinned"], false);

        dispatch_api(&state, "clear_path", &[json!("input")]).expect("clear the source directory");
        assert_eq!(state.get_state()["output"], "");
        assert_eq!(state.get_state()["outputPinned"], false);

        let invalid = dispatch_api(&state, "update", &[json!({"outputPinned": "yes"})])
            .expect_err("reject a non-boolean output pin state");
        assert_eq!(invalid.status, StatusCode::BAD_REQUEST);
    }

    #[test]
    fn s1_f3_state_payload_tracks_persisted_tasks_empty_state_and_read_errors() {
        let task = S1RecoveryTaskSummary {
            run_id: "run-01".to_owned(),
            file_id: "file-01".to_owned(),
            source_file: PathBuf::from("机构甲.xlsx"),
            audit_copy: Some(PathBuf::from("run-01/机构甲-审核副本.xlsx")),
            status: "WAITING_MANUAL_RECALC",
            reason: Some("等待人工重新计算".to_owned()),
            updated_at: "2026-10-29T12:00:00Z".to_owned(),
        };
        let mut state = json!({});
        let lost_copy = S1RecoveryTaskSummary {
            run_id: "run-02".to_owned(),
            file_id: "file-lost".to_owned(),
            source_file: PathBuf::from("机构乙.xlsx"),
            audit_copy: Some(PathBuf::from("run-02/机构乙-审核副本.xlsx")),
            status: "FAILED",
            reason: Some("待恢复任务的审核副本已丢失或路径无效".to_owned()),
            updated_at: "2026-10-29T12:01:00Z".to_owned(),
        };

        apply_s1_f3_recovery_state(&mut state, Ok((vec![task], vec![lost_copy])));
        assert_eq!(state["s1F3PendingCount"], 1);
        assert_eq!(state["s1F3RecoveryTasks"][0]["run_id"], "run-01");
        assert_eq!(state["s1F3RecoveryTasks"][0]["file_id"], "file-01");
        assert_eq!(state["s1F3LostCopyCount"], 1);
        assert_eq!(state["s1F3LostCopyTasks"][0]["status"], "FAILED");
        assert_eq!(state["s1F3RecoveryError"], "");

        apply_s1_f3_recovery_state(&mut state, Ok((Vec::new(), Vec::new())));
        assert_eq!(state["s1F3PendingCount"], 0);
        assert!(state["s1F3RecoveryTasks"].as_array().unwrap().is_empty());
        assert_eq!(state["s1F3LostCopyCount"], 0);
        assert!(state["s1F3LostCopyTasks"].as_array().unwrap().is_empty());
        assert_eq!(state["s1F3RecoveryError"], "");

        apply_s1_f3_recovery_state(
            &mut state,
            Err(anyhow::anyhow!("corrupt recovery manifest")),
        );
        assert_eq!(state["s1F3PendingCount"], 0);
        assert!(state["s1F3RecoveryTasks"].as_array().unwrap().is_empty());
        assert_eq!(state["s1F3LostCopyCount"], 0);
        assert!(state["s1F3LostCopyTasks"].as_array().unwrap().is_empty());
        assert!(state["s1F3RecoveryError"]
            .as_str()
            .unwrap()
            .contains("corrupt recovery manifest"));
    }

    fn recovery_result_issue(source_file: &Path, audit_file: &Path) -> Issue {
        Issue {
            issue_id: "S1-RECOVERY-001".to_owned(),
            period: "2026-07".to_owned(),
            batch_id: "recovery-test".to_owned(),
            audit_time: "2026-10-29 12:00:00".to_owned(),
            triggered: true,
            status: "新增".to_owned(),
            first_seen_period: "2026-07".to_owned(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: "1001".to_owned(),
            org_name: "测试机构".to_owned(),
            report_code: "R01".to_owned(),
            sheet_name: "检查表".to_owned(),
            rule_id: "R-001".to_owned(),
            severity: "错误".to_owned(),
            formula_cell: "D2".to_owned(),
            target_cell: "B2".to_owned(),
            target_value: json!(10),
            formula_result: json!("#DIV/0!"),
            message: "合成恢复结果".to_owned(),
            source_file: source_file.to_string_lossy().into_owned(),
            audit_file: audit_file.to_string_lossy().into_owned(),
            check_field: "金额".to_owned(),
            check_type: "公式校验".to_owned(),
            comparison_value: json!(""),
            reference_value: json!(""),
            difference_value: json!(""),
            detail: "初始恢复结果".to_owned(),
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        }
    }

    fn write_s1_f3_recovery_fixture(
        path: &Path,
        cached_result: Option<&str>,
    ) -> anyhow::Result<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("检查表")?;
        worksheet.write_formula(
            0,
            0,
            "=IF(TRUE,\"错误|公式核查|测试指标|测试异常\",\"通过\")",
        )?;
        if let Some(result) = cached_result {
            worksheet.set_formula_result(0, 0, result);
        }
        worksheet.write_string(0, 1, "目标字段")?;
        workbook.define_name("校验区域", "='检查表'!$A$1")?;
        workbook.define_name("条件格式区域", "='检查表'!$A$1")?;
        workbook.define_name("表结构区域", "='检查表'!$B$1")?;
        workbook.save(path)?;
        Ok(())
    }

    #[test]
    fn s1_f1_template_preflight_cache_reuses_one_immutable_plan_per_path() -> anyhow::Result<()> {
        let directory = S1F3RealProbeDirectory::create()?;
        let template = directory.0.join("template.xlsx");
        write_s1_f3_recovery_fixture(&template, None)?;
        let mut cache = std::collections::HashMap::new();

        let first = cached_s1_f1_template_preflight(&mut cache, &template, None)?;
        let second = cached_s1_f1_template_preflight(&mut cache, &template, None)?;

        assert!(std::sync::Arc::ptr_eq(&first, &second));
        assert_eq!(cache.len(), 1);
        assert_eq!(first.scan_area_count, second.scan_area_count);
        assert_eq!(first.structure_area_count, second.structure_area_count);
        Ok(())
    }

    #[test]
    #[ignore = "requires Windows Excel COM for a complete synthetic S1-F1 GUI batch"]
    #[cfg(windows)]
    fn s1_f1_gui_batch_reuses_template_preflight_and_preserves_legal_recalc_state(
    ) -> anyhow::Result<()> {
        let directory = S1F3RealProbeDirectory::create()?;
        let input = directory.0.join("input");
        let templates = directory.0.join("templates");
        let output = directory.0.join("output");
        fs::create_dir_all(&input)?;
        fs::create_dir_all(&templates)?;
        let template = templates.join("template.xlsx");
        write_s1_f3_recovery_fixture(&template, None)?;
        let mut source_hashes = Vec::new();
        for index in 1..=3 {
            let source = input.join(format!("机构甲_2026-07_{index}.xlsx"));
            write_s1_f3_recovery_fixture(&source, None)?;
            source_hashes.push((source.clone(), file_sha256(&source)?));
        }
        let template_hash = file_sha256(&template)?;

        let started = std::time::Instant::now();
        let report = run_s1_f1_for_acceptance(&input, &templates, None, None, &output, None)?;
        eprintln!(
            "synthetic S1-F1 shared-template GUI batch: files=3 elapsed_ms={}",
            started.elapsed().as_millis()
        );

        assert_eq!(report["preflight"]["filesTotal"], 3);
        assert_eq!(report["preflight"]["structurePassed"], 3);
        match report["status"].as_str() {
            Some("completed") => {
                assert_eq!(report["preflight"]["reviewedFiles"], 3);
                assert!(PathBuf::from(report["result_path"].as_str().unwrap()).is_file());
            }
            Some("partial") => {
                assert_eq!(report["preflight"]["waitingManualRecalc"], 3);
                assert_eq!(report["result_path"], "");
                assert!(report["logs"].as_array().unwrap().iter().any(|entry| {
                    entry["text"]
                        .as_str()
                        .is_some_and(|text| text.contains("自动公式重算不可用或失败"))
                }));
            }
            status => panic!("unexpected S1-F1 acceptance status {status:?}: {report}"),
        }
        assert_eq!(file_sha256(&template)?, template_hash);
        for (source, before) in source_hashes {
            assert_eq!(file_sha256(&source)?, before);
        }
        Ok(())
    }

    struct S1F3RealProbeDirectory(PathBuf);

    impl S1F3RealProbeDirectory {
        fn create() -> anyhow::Result<Self> {
            let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let path = std::env::temp_dir().join(format!("audit_s1_f3_real_{suffix}"));
            fs::create_dir_all(&path)?;
            Ok(Self(path))
        }
    }

    impl Drop for S1F3RealProbeDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    fn required_s1_real_path(name: &str) -> anyhow::Result<PathBuf> {
        let path = std::env::var_os(name)
            .map(PathBuf::from)
            .with_context(|| format!("real S1 acceptance requires {name}"))?;
        anyhow::ensure!(path.is_file(), "{name} is not a file");
        Ok(path)
    }

    #[test]
    #[ignore = "requires local read-only S1-F1 workbooks/configuration and Windows Excel COM"]
    #[cfg(windows)]
    fn real_s1_f3_recalculates_and_resumes_a_real_s1_f1_audit_copy() -> anyhow::Result<()> {
        let source = required_s1_real_path("AUDIT_S1_REAL_SOURCE")?;
        let template = required_s1_real_path("AUDIT_S1_REAL_TEMPLATE")?;
        let external = required_s1_real_path("AUDIT_S1_REAL_EXTERNAL")?;
        let mapping = required_s1_real_path("AUDIT_S1_REAL_MAPPING")?;
        let sidecar = required_s1_real_path("AUDIT_S1_REAL_SIDECAR")?;
        let original_hashes = [
            file_sha256(&source)?,
            file_sha256(&template)?,
            file_sha256(&external)?,
            file_sha256(&mapping)?,
        ];

        let temp = S1F3RealProbeDirectory::create()?;
        let output_directory = temp.0.join("output");
        let run_id = "s1-f3-real-resume";
        let run_directory = output_directory.join(run_id);
        let state_directory = temp.0.join("state");
        fs::create_dir_all(&run_directory)?;

        let audit_copy = run_directory.join("real-s1-f1-audit-copy.xlsx");
        let modifier = crate::excel_route::production_modifier("S1-F3 real workbook recovery")?.0;
        let manual_copy = crate::s1_formula_copy::execute_s1_formula_copy_with_manual_recovery(
            &source,
            &template,
            &audit_copy,
            Some(&mapping),
            Some(&external),
            None,
            &modifier,
            "auto",
        )?;
        let (formula_fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
            match manual_copy {
                S1FormulaCopyOutcome::WaitingManualRecalc(report) => {
                    anyhow::ensure!(
                        report.audit_copy == audit_copy,
                        "S1-F1 did not preserve the expected independent audit copy"
                    );
                    workbook_recovery_fingerprints(&report.audit_copy)?
                }
                S1FormulaCopyOutcome::Recalculated(_) => {
                    anyhow::bail!("S1-F1 unexpectedly recalculated without a sidecar")
                }
            };
        anyhow::ensure!(formula_count > 0, "real S1-F1 copy has no formulas");

        // Model the user opening the saved S1-F1 copy in Excel, forcing a full
        // recalculation, saving it, and closing Excel. The sidecar writes a
        // separate temporary output; only that output is copied back over this
        // disposable audit copy. No source, template, external workbook or
        // mapping configuration is opened for writing.
        let recalculated_path = run_directory.join("excel-recalculated-copy.xlsx");
        let recalculation = crate::s1_recalculation::recalculate_s1_formula_copy(
            &sidecar,
            &audit_copy,
            &recalculated_path,
            "auto",
        )?;
        anyhow::ensure!(
            recalculation.engine == "Excel.Application",
            "Windows real S1-F3 acceptance expected Excel COM; got {}",
            recalculation.engine
        );
        fs::copy(&recalculated_path, &audit_copy)?;
        let audit_copy_hash = file_sha256(&audit_copy)?;

        let mut recovery_file = S1RecoveryFile::new(
            source.clone(),
            template.clone(),
            file_sha256(&template)?,
            Some(audit_copy.clone()),
            S1RecoveryStatus::WaitingManualRecalc,
        );
        recovery_file.formula_fingerprint = Some(formula_fingerprint);
        recovery_file.formula_fingerprint_normalized = Some(normalized_fingerprint);
        recovery_file.audit_copy_semantic_fingerprint = Some(semantic_fingerprint);
        recovery_file.reason = Some("真实样本 S1-F1 副本已由 Excel COM 重算并保存".to_owned());
        let file_id = recovery_file.file_id.clone();
        let mut manifest = S1RecoveryManifest::new(
            run_id,
            output_directory,
            run_directory,
            "2026-07",
            "s1-f3-real-resume",
            "2026-10-29 00:00:00",
            Some(mapping.clone()),
            Some(file_sha256(&mapping)?),
            Vec::new(),
            Vec::new(),
            vec![recovery_file],
        );
        manifest.external_data = Some(external.clone());
        let first_store = S1RecoveryStore::new(&state_directory);
        first_store.save(&mut manifest)?;

        // Construct a fresh store to cover the persisted-task path used after
        // an application restart, then invoke the same worker as the GUI.
        let restarted_store = S1RecoveryStore::new(&state_directory);
        let state = test_state();
        {
            let mut ui = state.lock();
            ui.value["busy"] = json!(true);
        }
        run_s1_f3_worker_with_store(
            state.clone(),
            run_id.to_owned(),
            file_id,
            Ok(restarted_store.clone()),
        );

        let ui = state.lock();
        anyhow::ensure!(ui.value["busy"] == false, "S1-F3 worker remained busy");
        anyhow::ensure!(
            ui.value["s1F1Preflight"]["status"] == "recovered",
            "S1-F3 real-workbook recovery failed: {}",
            ui.value["status"]
        );
        let result_path = ui.value["s1F1OutputPath"]
            .as_str()
            .context("S1-F3 did not publish the real-workbook result path")?
            .to_owned();
        drop(ui);

        let completed = restarted_store.load(run_id)?;
        anyhow::ensure!(
            matches!(
                completed.files[0].status,
                S1RecoveryStatus::Completed | S1RecoveryStatus::Partial
            ),
            "unexpected S1-F3 completion state: {}",
            completed.files[0].status.as_str()
        );
        anyhow::ensure!(completed.files[0].issues_extracted);
        anyhow::ensure!(restarted_store.pending_tasks()?.is_empty());
        anyhow::ensure!(completed.result_path.as_deref() == Some(Path::new(&result_path)));
        anyhow::ensure!(Path::new(&result_path).is_file());
        anyhow::ensure!(file_sha256(&audit_copy)? == audit_copy_hash);
        anyhow::ensure!(
            [
                file_sha256(&source)?,
                file_sha256(&template)?,
                file_sha256(&external)?,
                file_sha256(&mapping)?,
            ] == original_hashes,
            "S1-F3 changed a source workbook, template, external workbook or mapping config"
        );

        let output = excel_reader::read_workbook(&result_path)?;
        let result_sheet = output
            .sheets
            .iter()
            .find(|sheet| sheet.name == crate::s1_result_workbook::S1_RESULT_SHEET)
            .context("S1-F3 real result workbook is missing the audit result sheet")?;
        anyhow::ensure!(
            result_sheet
                .values
                .get("A1")
                .and_then(|cell| cell["value"].as_str())
                == Some("工作簿名")
                && result_sheet
                    .values
                    .get("M1")
                    .and_then(|cell| cell["value"].as_str())
                    == Some("审核意见"),
            "S1-F3 real result workbook does not have the expected 13-column output headers"
        );

        eprintln!(
            "real S1-F3 recovery passed: engine={}, formula_cells={}, formula_issues={}, conditional_issues={}, diagnostics={}, status={}",
            recalculation.engine,
            formula_count,
            completed.files[0].formula_issues.len(),
            completed.files[0].conditional_issues.len(),
            completed.files[0].conditional_diagnostics.len(),
            completed.files[0].status.as_str()
        );
        Ok(())
    }

    #[test]
    fn conditional_extraction_failure_keeps_formula_findings_as_partial_result(
    ) -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_conditional_partial_{suffix}"));
        fs::create_dir_all(&root)?;
        let source = root.join("机构甲_部分审核.xlsx");
        let template = root.join("template.xlsx");
        let audit_copy = root.join("audit-copy.xlsx");
        write_s1_f3_recovery_fixture(&source, None)?;
        write_s1_f3_recovery_fixture(&template, None)?;
        write_s1_f3_recovery_fixture(&audit_copy, Some("错误|公式核查|测试指标|测试异常"))?;
        let preflight = s1_f1_template_preflight(&template, None)?;
        let item = S1F1AcceptedSource {
            source: source.clone(),
            template,
            preflight: std::sync::Arc::new(preflight),
        };
        let extracted = extract_s1_audit_copy_with_conditional_extractor(
            &item,
            &source,
            &audit_copy,
            &S1FormulaMetadata {
                period: "2026-07".to_owned(),
                batch_id: "conditional-partial-test".to_owned(),
                audit_time: "2026-10-29 12:00:00".to_owned(),
                org_code: "1001".to_owned(),
                org_name: "测试机构".to_owned(),
            },
            |_, _, _, _, _| anyhow::bail!("合成条件格式局部读取失败"),
        )?;

        assert_eq!(extracted.formula_issues.len(), 1);
        assert!(extracted.conditional_issues.is_empty());
        assert!(extracted.partial);
        assert!(extracted
            .reason
            .as_deref()
            .unwrap_or_default()
            .contains("已保留公式审核结果"));
        assert_eq!(extracted.diagnostics.len(), 1);
        assert_eq!(extracted.diagnostics[0]["kind"], "conditional_format");
        assert_eq!(extracted.diagnostics[0]["status"], "error");

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f1_batch_publishes_completed_file_and_persists_other_file_for_manual_recalc(
    ) -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_f1_mixed_recalc_{suffix}"));
        let input_dir = root.join("input");
        let template_dir = root.join("templates");
        let output_dir = root.join("output");
        let run_dir = output_dir.join("S1-F1_mixed");
        let state_dir = root.join("state");
        fs::create_dir_all(&input_dir)?;
        fs::create_dir_all(&template_dir)?;

        let calculated_source = input_dir.join("机构甲_2026-07_自动重算.xlsx");
        let pending_source = input_dir.join("机构乙_2026-07_待人工重算.xlsx");
        let template = template_dir.join("template.xlsx");
        write_s1_f3_recovery_fixture(&calculated_source, None)?;
        write_s1_f3_recovery_fixture(&pending_source, None)?;
        write_s1_f3_recovery_fixture(&template, None)?;

        let accepted = vec![
            S1F1AcceptedSource {
                source: calculated_source.clone(),
                template: template.clone(),
                preflight: std::sync::Arc::new(s1_f1_template_preflight(&template, None)?),
            },
            S1F1AcceptedSource {
                source: pending_source.clone(),
                template: template.clone(),
                preflight: std::sync::Arc::new(s1_f1_template_preflight(&template, None)?),
            },
        ];
        let request = S1F1PreflightRequest {
            source_files: vec![calculated_source.clone(), pending_source.clone()],
            source_directory: input_dir,
            template_dir,
            manual_template: Some(template.clone()),
            mapping_config: None,
            external_data: None,
            output_directory: output_dir,
            batch_id: "s1-f1-mixed-recalc".to_owned(),
            run_directory: run_dir,
            planned_result_path: root.join("output/S1-F1_mixed/汇总核查表校验.xlsx"),
            sidecar_path: None,
            calc_engine: "auto".to_owned(),
            edit_engine: "umya".to_owned(),
            request_preparation_us: 0,
            source_directory_scan_us: 0,
        };
        let store = S1RecoveryStore::new(state_dir);
        let output = execute_s1_f1_workflow_with_dependencies(
            &request,
            &accepted,
            Some(store.clone()),
            |item, request, audit_copy, modifier| {
                let outcome = crate::s1_formula_copy::execute_s1_formula_copy_with_manual_recovery(
                    &item.source,
                    &item.template,
                    audit_copy,
                    request.mapping_config.as_deref(),
                    request.external_data.as_deref(),
                    None,
                    modifier,
                    &request.calc_engine,
                )?;
                match outcome {
                    S1FormulaCopyOutcome::WaitingManualRecalc(manual)
                        if item.source == calculated_source =>
                    {
                        // Model an external engine that completed and saved this
                        // one synthetic workbook while leaving the other waiting.
                        write_s1_f3_recovery_fixture(
                            &manual.audit_copy,
                            Some("错误|公式核查|测试指标|自动重算异常"),
                        )?;
                        let recalculated_path = manual.audit_copy.clone();
                        Ok(S1FormulaCopyOutcome::Recalculated(
                            S1FormulaCopyExecutionReport {
                                edit: manual.edit,
                                recalculation: S1RecalculationReport {
                                    engine: "synthetic_external_engine".to_owned(),
                                    input: recalculated_path.clone(),
                                    output: recalculated_path,
                                    elapsed_ms: 1,
                                },
                                formula_cells_prepared: manual.formula_cells_prepared,
                                cached_values_cleared: manual.cached_values_cleared,
                                external_sheets_imported: manual.external_sheets_imported,
                            },
                        ))
                    }
                    waiting @ S1FormulaCopyOutcome::WaitingManualRecalc(_) => Ok(waiting),
                    S1FormulaCopyOutcome::Recalculated(_) => {
                        anyhow::bail!("没有 sidecar 的合成路径意外报告了重算成功")
                    }
                }
            },
        )?;

        assert_eq!(output.waiting_manual_recalc, 1);
        assert_eq!(output.partial_files, 0);
        assert_eq!(output.failed_files, 0);
        assert_eq!(output.formula_issue_count, 1);
        assert!(output.performance_timings["workflowTotalUs"].is_number());
        assert!(output.performance_timings["copyAndRecalculationPreparationTotalUs"].is_number());
        assert!(output.performance_timings["externalExcelRecalculationUs"].is_number());
        assert_eq!(
            output.performance_timings["perFile"]
                .as_array()
                .unwrap()
                .len(),
            2
        );
        assert!(
            output.performance_timings["perFile"][0]["excelCoreEditAndValidationUs"].is_number()
        );
        let result_path = output.result_path.context("混合批次未生成成功文件结果")?;
        assert!(result_path.is_file());

        let manifest = store.load("s1-f1-mixed-recalc")?;
        let completed = manifest
            .files
            .iter()
            .find(|file| file.source_file == calculated_source)
            .context("恢复清单缺少自动重算成功文件")?;
        assert_eq!(completed.status, S1RecoveryStatus::Completed);
        assert!(completed.issues_extracted);
        assert_eq!(completed.formula_issues.len(), 1);
        let pending = manifest
            .files
            .iter()
            .find(|file| file.source_file == pending_source)
            .context("恢复清单缺少待人工重算文件")?;
        assert_eq!(pending.status, S1RecoveryStatus::WaitingManualRecalc);
        assert!(!pending.issues_extracted);
        assert!(pending.audit_copy.as_deref().is_some_and(Path::is_file));

        let pending_tasks = S1RecoveryStore::new(store.directory()).pending_tasks()?;
        assert_eq!(pending_tasks.len(), 1);
        assert_eq!(pending_tasks[0].file_id, pending.file_id);
        let result = excel_reader::read_workbook(&result_path)?;
        let result_sheet = result
            .sheets
            .iter()
            .find(|sheet| sheet.name == crate::s1_result_workbook::S1_RESULT_SHEET)
            .context("混合批次输出缺少本期审核结果工作表")?;
        assert!(result_sheet.values.values().any(|value| {
            value["value"]
                .as_str()
                .is_some_and(|text| text.contains("自动重算异常"))
        }));
        assert!(!result_sheet.values.values().any(|value| {
            value["value"]
                .as_str()
                .is_some_and(|text| text.contains("待人工重算"))
        }));

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f3_worker_resumes_cached_copy_after_store_restart_and_publishes_real_result(
    ) -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_f3_worker_{suffix}"));
        let output_dir = root.join("output");
        let run_dir = output_dir.join("run-01");
        let state_dir = root.join("state");
        fs::create_dir_all(&run_dir)?;

        let source = root.join("机构甲_测试报表.xlsx");
        let template = root.join("template.xlsx");
        let audit_copy = run_dir.join("机构甲_测试报表_审核副本.xlsx");
        write_s1_f3_recovery_fixture(&source, None)?;
        write_s1_f3_recovery_fixture(&template, None)?;
        // This synthetic cache represents a workbook that the user has already
        // recalculated and saved. It verifies the S1-F3 resume stage only; it
        // does not stand in for Excel/LibreOffice recalculation acceptance.
        write_s1_f3_recovery_fixture(&audit_copy, Some("错误|公式核查|测试指标|测试异常"))?;
        let source_sha256 = file_sha256(&source)?;
        let template_sha256 = file_sha256(&template)?;
        let audit_copy_sha256 = file_sha256(&audit_copy)?;

        let (formula_fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
            workbook_recovery_fingerprints(&audit_copy)?;
        assert_eq!(formula_count, 1);
        let mut recovery_file = S1RecoveryFile::new(
            source.clone(),
            template.clone(),
            template_sha256.clone(),
            Some(audit_copy.clone()),
            S1RecoveryStatus::WaitingManualRecalc,
        );
        recovery_file.formula_fingerprint = Some(formula_fingerprint);
        recovery_file.formula_fingerprint_normalized = Some(normalized_fingerprint);
        recovery_file.audit_copy_semantic_fingerprint = Some(semantic_fingerprint);
        let file_id = recovery_file.file_id.clone();
        let mut manifest = S1RecoveryManifest::new(
            "s1-f3-worker-restart-test",
            output_dir,
            run_dir,
            "2026-07",
            "s1-f3-worker-test",
            "2026-10-29 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![recovery_file],
        );
        let first_store = S1RecoveryStore::new(&state_dir);
        first_store.save(&mut manifest)?;
        assert_eq!(first_store.pending_tasks()?.len(), 1);

        // A newly constructed store represents a fresh application process.
        let restarted_store = S1RecoveryStore::new(&state_dir);
        let state = test_state();
        {
            let mut ui = state.lock();
            ui.value["busy"] = json!(true);
        }
        run_s1_f3_worker_with_store(
            state.clone(),
            "s1-f3-worker-restart-test".to_owned(),
            file_id,
            Ok(restarted_store.clone()),
        );

        let ui = state.lock();
        assert_eq!(ui.value["busy"], false);
        assert_eq!(ui.value["s1F1Preflight"]["status"], "recovered");
        assert_eq!(ui.value["s1F1Preflight"]["formulaIssues"], 1);
        assert_eq!(ui.value["s1F1Preflight"]["conditionalIssues"], 0);
        drop(ui);

        let completed = restarted_store.load("s1-f3-worker-restart-test")?;
        assert_eq!(completed.files[0].status, S1RecoveryStatus::Completed);
        assert!(completed.files[0].issues_extracted);
        assert_eq!(completed.files[0].formula_issues.len(), 1);
        assert!(completed.files[0].conditional_issues.is_empty());
        assert!(restarted_store.pending_tasks()?.is_empty());

        let result_path = completed.result_path.context("恢复后缺少结果路径")?;
        assert!(result_path.is_file());
        let output = excel_reader::read_workbook(&result_path)?;
        let result_sheet = output
            .sheets
            .iter()
            .find(|sheet| sheet.name == crate::s1_result_workbook::S1_RESULT_SHEET)
            .context("恢复结果缺少本期审核结果工作表")?;
        assert!(result_sheet.values.len() > 13);
        assert_eq!(
            result_sheet
                .values
                .get("A1")
                .and_then(|cell| cell["value"].as_str()),
            Some("工作簿名")
        );
        assert_eq!(
            result_sheet
                .values
                .get("M1")
                .and_then(|cell| cell["value"].as_str()),
            Some("审核意见")
        );
        assert!(result_sheet.values.values().any(|value| matches!(
            value["value"].as_str(),
            Some(text) if text.contains("测试异常")
        )));

        // F3 reads the original report and saved audit copy but never edits either.
        assert_eq!(file_sha256(&source)?, source_sha256);
        assert_eq!(file_sha256(&template)?, template_sha256);
        assert_eq!(file_sha256(&audit_copy)?, audit_copy_sha256);

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f3_worker_keeps_missing_cache_task_pending_without_publishing_result(
    ) -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_f3_missing_cache_{suffix}"));
        let output_dir = root.join("output");
        let run_dir = output_dir.join("run-01");
        let state_dir = root.join("state");
        fs::create_dir_all(&run_dir)?;

        let source = root.join("机构乙_测试报表.xlsx");
        let template = root.join("template.xlsx");
        let calculated_copy = run_dir.join("calculated-before-cache-clear.xlsx");
        let audit_copy = run_dir.join("机构乙_测试报表_审核副本.xlsx");
        write_s1_f3_recovery_fixture(&source, None)?;
        write_s1_f3_recovery_fixture(&template, None)?;
        write_s1_f3_recovery_fixture(&calculated_copy, Some("错误|公式核查|测试指标|测试异常"))?;
        ooxml_patch::clear_formula_caches_for_recalculation(&calculated_copy, &audit_copy)?;
        let (formula_fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
            workbook_recovery_fingerprints(&audit_copy)?;
        assert_eq!(formula_count, 1);
        assert_eq!(
            excel_reader::inspect_formula_cache_nodes(&audit_copy)?
                .missing_cache_cells
                .len(),
            1
        );
        let audit_copy_sha256 = file_sha256(&audit_copy)?;

        let mut recovery_file = S1RecoveryFile::new(
            source,
            template.clone(),
            file_sha256(&template)?,
            Some(audit_copy.clone()),
            S1RecoveryStatus::WaitingManualRecalc,
        );
        recovery_file.formula_fingerprint = Some(formula_fingerprint);
        recovery_file.formula_fingerprint_normalized = Some(normalized_fingerprint);
        recovery_file.audit_copy_semantic_fingerprint = Some(semantic_fingerprint);
        let file_id = recovery_file.file_id.clone();
        let mut manifest = S1RecoveryManifest::new(
            "s1-f3-missing-cache-test",
            output_dir,
            run_dir.clone(),
            "2026-07",
            "s1-f3-missing-cache-test",
            "2026-10-29 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![recovery_file],
        );
        let store = S1RecoveryStore::new(&state_dir);
        store.save(&mut manifest)?;

        let state = test_state();
        {
            let mut ui = state.lock();
            ui.value["busy"] = json!(true);
        }
        run_s1_f3_worker_with_store(
            state.clone(),
            "s1-f3-missing-cache-test".to_owned(),
            file_id,
            Ok(store.clone()),
        );

        let ui = state.lock();
        assert_eq!(ui.value["busy"], false);
        assert!(ui.value["status"]
            .as_str()
            .unwrap()
            .contains("公式缓存仍缺失"));
        drop(ui);

        let pending = store.pending_tasks()?;
        assert_eq!(pending.len(), 1);
        assert_eq!(pending[0].status, "WAITING_MANUAL_RECALC");
        assert!(pending[0]
            .reason
            .as_deref()
            .is_some_and(|reason| reason.contains("公式缓存仍缺失")));
        let waiting = store.load("s1-f3-missing-cache-test")?;
        assert!(waiting.result_path.is_none());
        assert!(!waiting.files[0].issues_extracted);
        assert_eq!(
            waiting.files[0].status,
            S1RecoveryStatus::WaitingManualRecalc
        );
        assert_eq!(file_sha256(&audit_copy)?, audit_copy_sha256);
        assert_eq!(
            fs::read_dir(&run_dir)?
                .filter_map(Result::ok)
                .filter(|entry| entry
                    .file_name()
                    .to_string_lossy()
                    .contains("汇总核查表校验_恢复_"))
                .count(),
            0
        );

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f3_worker_reports_corrupt_copy_and_preserves_it_for_repair() -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_f3_corrupt_copy_{suffix}"));
        let output_dir = root.join("output");
        let run_dir = output_dir.join("run-01");
        let state_dir = root.join("state");
        fs::create_dir_all(&run_dir)?;

        let source = root.join("机构丙_测试报表.xlsx");
        let template = root.join("template.xlsx");
        let audit_copy = run_dir.join("机构丙_测试报表_审核副本.xlsx");
        write_s1_f3_recovery_fixture(&source, None)?;
        write_s1_f3_recovery_fixture(&template, None)?;
        write_s1_f3_recovery_fixture(&audit_copy, Some("错误|公式核查|测试指标|测试异常"))?;
        let (formula_fingerprint, normalized_fingerprint, semantic_fingerprint, formula_count) =
            workbook_recovery_fingerprints(&audit_copy)?;
        assert_eq!(formula_count, 1);
        let corrupt_bytes = b"not an xlsx workbook; retain for manual repair";
        fs::write(&audit_copy, corrupt_bytes)?;

        let mut recovery_file = S1RecoveryFile::new(
            source,
            template.clone(),
            file_sha256(&template)?,
            Some(audit_copy.clone()),
            S1RecoveryStatus::WaitingManualRecalc,
        );
        recovery_file.formula_fingerprint = Some(formula_fingerprint);
        recovery_file.formula_fingerprint_normalized = Some(normalized_fingerprint);
        recovery_file.audit_copy_semantic_fingerprint = Some(semantic_fingerprint);
        let file_id = recovery_file.file_id.clone();
        let mut manifest = S1RecoveryManifest::new(
            "s1-f3-corrupt-copy-test",
            output_dir,
            run_dir,
            "2026-07",
            "s1-f3-corrupt-copy-test",
            "2026-10-29 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![recovery_file],
        );
        let store = S1RecoveryStore::new(&state_dir);
        store.save(&mut manifest)?;

        let state = test_state();
        {
            let mut ui = state.lock();
            ui.value["busy"] = json!(true);
        }
        run_s1_f3_worker_with_store(
            state.clone(),
            "s1-f3-corrupt-copy-test".to_owned(),
            file_id,
            Ok(store.clone()),
        );

        let ui = state.lock();
        assert_eq!(ui.value["busy"], false);
        assert!(ui.value["status"]
            .as_str()
            .unwrap()
            .contains("S1-F3 恢复未完成"));
        drop(ui);

        let manifest = store.load("s1-f3-corrupt-copy-test")?;
        assert_eq!(manifest.files[0].status, S1RecoveryStatus::ResumeFailed);
        assert!(manifest.files[0]
            .reason
            .as_deref()
            .is_some_and(|reason| reason.contains("读取审核副本完整性指纹")));
        assert!(manifest.result_path.is_none());
        assert!(!manifest.files[0].issues_extracted);
        assert_eq!(fs::read(&audit_copy)?, corrupt_bytes);
        let pending = store.pending_tasks()?;
        assert_eq!(pending.len(), 1);
        assert_eq!(pending[0].status, "RESUME_FAILED");
        assert!(excel_reader::read_workbook(&audit_copy).is_err());

        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f3_result_publication_is_idempotent_across_store_restart() -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_f3_publish_{suffix}"));
        let run_dir = root.join("output/run");
        fs::create_dir_all(&run_dir)?;
        let source = root.join("source.xlsx");
        let audit_copy = run_dir.join("audit-copy.xlsx");
        let mut file = S1RecoveryFile::new(
            source.clone(),
            root.join("template.xlsx"),
            "template-hash".to_owned(),
            Some(audit_copy.clone()),
            S1RecoveryStatus::ReadyToResume,
        );
        file.issues_extracted = true;
        file.formula_issues
            .push(recovery_result_issue(&source, &audit_copy));
        let mut manifest = S1RecoveryManifest::new(
            "recovery-publish-test",
            root.join("output"),
            run_dir.clone(),
            "2026-07",
            "recovery-test",
            "2026-10-29 12:00:00",
            None,
            None,
            Vec::new(),
            Vec::new(),
            vec![file],
        );
        let store = S1RecoveryStore::new(root.join("state"));
        store.save(&mut manifest)?;

        let first = write_s1_recovery_manifest_result(&store, &mut manifest)?;
        let first_bytes = fs::read(&first.path)?;
        assert_eq!(first.formula_issue_count, 1);
        assert_eq!(first.total_issue_count, 1);
        assert!(excel_reader::read_workbook(&first.path).is_ok());

        let restarted_store = S1RecoveryStore::new(root.join("state"));
        let mut restored = restarted_store.load("recovery-publish-test")?;
        let second = write_s1_recovery_manifest_result(&restarted_store, &mut restored)?;
        assert_eq!(first.path, second.path);
        assert_eq!(fs::read(&second.path)?, first_bytes);
        let published_results = fs::read_dir(&run_dir)?
            .filter_map(Result::ok)
            .map(|entry| entry.path())
            .filter(|path| {
                path.extension().and_then(|extension| extension.to_str()) == Some("xlsx")
                    && path
                        .file_name()
                        .is_some_and(|name| name.to_string_lossy().contains("_恢复_"))
            })
            .count();
        assert_eq!(published_results, 1);

        restored.files[0].formula_issues[0].detail = "修订后的审核结论".to_owned();
        restarted_store.save(&mut restored)?;
        let revised = write_s1_recovery_manifest_result(&restarted_store, &mut restored)?;
        assert_ne!(revised.path, first.path);
        assert_eq!(fs::read(&first.path)?, first_bytes);
        assert_eq!(
            fs::read_dir(&run_dir)?
                .filter_map(Result::ok)
                .map(|entry| entry.path())
                .filter(|path| {
                    path.extension().and_then(|extension| extension.to_str()) == Some("xlsx")
                        && path
                            .file_name()
                            .is_some_and(|name| name.to_string_lossy().contains("_恢复_"))
                })
                .count(),
            2,
            "a changed issue set gets a separate report without overwriting the previous result"
        );
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn s1_f1_failed_run_directory_is_removed_without_touching_neighbors() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let parent = std::env::temp_dir().join(format!("audit_s1_f1_output_guard_{suffix}"));
        fs::create_dir_all(&parent).unwrap();
        let neighbor = parent.join("user-file.txt");
        fs::write(&neighbor, b"keep").unwrap();
        let run = parent.join("S1-F1_审核结果_测试");
        {
            let _guard = S1F1OutputDirectoryGuard::create(run.clone()).unwrap();
            fs::write(run.join("partial.xlsx"), b"partial").unwrap();
        }
        assert!(!run.exists());
        assert_eq!(fs::read(&neighbor).unwrap(), b"keep");
        fs::remove_dir_all(parent).unwrap();
    }

    #[test]
    fn s1_f1_successful_run_directory_is_retained() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let parent = std::env::temp_dir().join(format!("audit_s1_f1_output_keep_{suffix}"));
        fs::create_dir_all(&parent).unwrap();
        let run = parent.join("S1-F1_审核结果_测试");
        {
            let mut guard = S1F1OutputDirectoryGuard::create(run.clone()).unwrap();
            fs::write(run.join("result.xlsx"), b"result").unwrap();
            guard.keep();
        }
        assert!(run.join("result.xlsx").is_file());
        fs::remove_dir_all(parent).unwrap();
    }

    fn write_s1_region_source(path: &Path) {
        let mut workbook = Workbook::new();
        let data = workbook.add_worksheet();
        data.set_name("Data").expect("name S1 summary source sheet");
        data.write_string(1, 1, "财务")
            .expect("write S1 summary structure group header");
        data.write_string(2, 1, "单位")
            .expect("write S1 summary structure row header");
        data.write_string(2, 2, "金额")
            .expect("write S1 summary structure value header");
        data.write_string(4, 1, "甲")
            .expect("write S1 summary row label");
        data.write_number(4, 2, 10.0)
            .expect("write S1 summary row value");
        data.write_string(6, 1, "乙")
            .expect("write second S1 summary row label");
        data.write_number(6, 2, 20.0)
            .expect("write second S1 summary row value");
        data.write_string(0, 5, "测试分组")
            .expect("write S1 local metadata");
        let meta = workbook.add_worksheet();
        meta.set_name("Meta")
            .expect("name S1 summary metadata sheet");
        meta.write_string(0, 0, "测试批次")
            .expect("write S1 global metadata");
        workbook.save(path).expect("save S1 summary source");
    }

    fn write_s3_comparison(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet
            .set_name("比较结果")
            .expect("set comparison sheet name");
        let headers = [
            "机构类代码",
            "机构类名称",
            "地区代码",
            "地区名称",
            "指标代码",
            "指标名称",
            "数据值",
            "上期值",
            "输出说明",
            "说明内容",
            "增减额",
        ];
        for (column, header) in headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write comparison header");
        }
        let row = [
            "12345",
            "示例机构",
            "4401",
            "示例地区",
            "I-01",
            "示例指标",
            "100",
            "90",
            "是",
            "需要说明",
            "10",
        ];
        for (column, value) in row.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write selected comparison row");
        }
        workbook.save(path).expect("save comparison workbook");
    }

    fn write_s3_common_config(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet
            .set_name("机构地区参照")
            .expect("set organization reference sheet name");
        for (column, header) in ["机构类代码", "地区代码", "机构名称"].iter().enumerate()
        {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write organization reference header");
        }
        for (column, value) in ["12345", "4401", "配置机构名称"].iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write organization reference row");
        }
        workbook
            .save(path)
            .expect("save common configuration workbook");
    }

    fn write_s3_current_value_common_config(path: &Path) {
        let mut workbook = Workbook::new();
        let parameters = workbook.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        for (column, header) in ["参数", "值"].iter().enumerate() {
            parameters.write_string(0, column as u16, *header).unwrap();
        }
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = workbook.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        exceptions.write_string(0, 0, "指标代码").unwrap();
        workbook.save(path).unwrap();
    }

    /// 3.0 common config with all sheets the S3-F01 loaders require,
    /// including one 环比警戒 band so the check reports a non-zero count.
    fn write_s3_full_common_config(path: &Path) {
        let mut workbook = Workbook::new();
        let parameters = workbook.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        for (column, header) in ["参数", "值"].iter().enumerate() {
            parameters.write_string(0, column as u16, *header).unwrap();
        }
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = workbook.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        for (column, header) in ["指标代码", "启用"].iter().enumerate() {
            exceptions.write_string(0, column as u16, *header).unwrap();
        }

        let alerts = workbook.add_worksheet();
        alerts.set_name("环比警戒").unwrap();
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
            alerts.write_string(0, column as u16, *header).unwrap();
        }
        for (column, value) in ["0.1", "0.2", "环比增幅超限", "FFFF00", "否", "0", "是"]
            .iter()
            .enumerate()
        {
            alerts.write_string(1, column as u16, *value).unwrap();
        }
        workbook.save(path).unwrap();
    }

    /// 3.1 feature config with the sheets the S3-F01 loader requires.
    fn write_s3_execution_feature_config(path: &Path) {
        let mut workbook = Workbook::new();
        let indicators = workbook.add_worksheet();
        indicators.set_name("指标参照").unwrap();
        for (column, header) in ["指标代码", "表单代码", "行序号", "指标别名"]
            .iter()
            .enumerate()
        {
            indicators.write_string(0, column as u16, *header).unwrap();
        }
        for (column, value) in ["I001", "F001", "1", "合成指标"].iter().enumerate() {
            indicators.write_string(1, column as u16, *value).unwrap();
        }

        let actions = workbook.add_worksheet();
        actions.set_name("规则动作").unwrap();
        for (column, header) in [
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
        ]
        .iter()
        .enumerate()
        {
            actions.write_string(0, column as u16, *header).unwrap();
        }

        let expressions = workbook.add_worksheet();
        expressions.set_name("表达式校验").unwrap();
        for (column, header) in [
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ]
        .iter()
        .enumerate()
        {
            expressions.write_string(0, column as u16, *header).unwrap();
        }
        for (column, value) in ["E1", "合成表达式", "left > right", "否", "0", "是", ""]
            .iter()
            .enumerate()
        {
            expressions.write_string(1, column as u16, *value).unwrap();
        }
        workbook.save(path).unwrap();
    }

    fn write_s3_current_value_cross_config(path: &Path) {
        let mut workbook = Workbook::new();
        let rules = workbook.add_worksheet();
        rules.set_name("跨期数值核对").unwrap();
        let headers = [
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
        ];
        for (column, header) in headers.iter().enumerate() {
            rules.write_string(0, column as u16, *header).unwrap();
        }
        let values = [
            "R1",
            "Web 集成测试规则",
            "^0001$",
            "1234567",
            "",
            "[I1,余额,人民币,月,2]",
            "0",
            "left > right",
            "请核实",
            "",
            "是",
            "",
            "",
            "synthetic web fixture",
        ];
        for (column, value) in values.iter().enumerate() {
            rules.write_string(1, column as u16, *value).unwrap();
        }
        workbook.save(path).unwrap();
    }

    fn write_s3_current_value_input_xlsx(path: &Path) {
        let mut workbook = Workbook::new();
        let cover = workbook.add_worksheet();
        cover.set_name("说明").unwrap();
        cover.write_string(0, 0, "非导数工作表").unwrap();

        let data = workbook.add_worksheet();
        data.set_name("本期导数").unwrap();
        for (column, header) in EXPECTED_HEADERS.iter().enumerate() {
            data.write_string(0, column as u16, *header).unwrap();
        }
        let values = [
            "人民币",
            "2026-08-31",
            "0001",
            "机构",
            "1234567",
            "地区",
            "1",
            "I1",
            "指标",
            "余额",
            "人民币",
            "月",
            "2",
        ];
        for (column, value) in values.iter().enumerate() {
            data.write_string(2, column as u16, *value).unwrap();
        }
        data.write_number(2, 13, 200_000_000.0).unwrap();
        workbook.save(path).unwrap();
    }

    fn write_s3_execution_configs(common_path: &Path, feature_path: &Path) {
        let mut common = Workbook::new();
        let parameters = common.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        for (column, header) in ["参数", "值"].iter().enumerate() {
            parameters.write_string(0, column as u16, *header).unwrap();
        }
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = common.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        exceptions.write_string(0, 0, "指标代码").unwrap();
        exceptions.write_string(0, 1, "启用").unwrap();

        let alerts = common.add_worksheet();
        alerts.set_name("环比警戒").unwrap();
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
            alerts.write_string(0, column as u16, *header).unwrap();
        }
        common.save(common_path).unwrap();

        let mut feature = Workbook::new();
        let indicators = feature.add_worksheet();
        indicators.set_name("指标参照").unwrap();
        for (column, header) in ["指标代码", "表单代码", "行序号", "指标别名"]
            .iter()
            .enumerate()
        {
            indicators.write_string(0, column as u16, *header).unwrap();
        }
        let actions = feature.add_worksheet();
        actions.set_name("规则动作").unwrap();
        for (column, header) in [
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
        ]
        .iter()
        .enumerate()
        {
            actions.write_string(0, column as u16, *header).unwrap();
        }
        let expressions = feature.add_worksheet();
        expressions.set_name("表达式校验").unwrap();
        for (column, header) in [
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ]
        .iter()
        .enumerate()
        {
            expressions.write_string(0, column as u16, *header).unwrap();
        }
        feature.save(feature_path).unwrap();
    }

    fn write_s3_execution_csv(path: &Path, date: &str, amount: &str) {
        let row = format!(
            "业务一,{date},ORG1,示例机构,R001,示例地区,1,I001,指标一,余额,CNY,月,1,{amount}"
        );
        fs::write(path, format!("{}\n{row}\n", EXPECTED_HEADERS.join(","))).unwrap();
    }

    #[test]
    fn served_page_is_original_markup_plus_runtime_bridge() {
        let served = render_index();
        let served_markup = served.replace(super::BRIDGE_JS, "");
        let expected_markup = super::ORIGINAL_HTML
            .replace("__AUDIT_APP_TITLE__", super::APP_TITLE)
            .replace("__AUDIT_ICON_URL__", "/assets/app-icon.png");
        assert_eq!(served_markup, expected_markup);
        assert!(!served.contains("__AUDIT_APP_TITLE__"));
        assert!(!served.contains("__AUDIT_ICON_URL__"));
        assert!(served_markup.contains("centralCrossCur:'central_data'"));
        assert!(served_markup.contains("本期数值核对文件（CSV/XLSX/XLS）"));
        assert!(!served_markup.contains("XLS/XLSX 尚未接通"));
        assert!(served_markup.contains("id=\"s1F3RecoveryButton\" hidden"));
        assert!(served_markup.contains("id=\"s1F3RecoveryPanel\" hidden"));
        assert!(served_markup.contains("#s1F3RecoveryButton[hidden]{display:none!important}"));
        assert!(served.contains("button.hidden=tasks.length===0"));
        assert!(served.contains("s1F3LostCopyTasks"));
        assert!(served.contains("warning.hidden=warningMessages.length===0"));
        assert!(served.contains("请重新运行 S1-F1"));
        assert!(served.contains("window.bridge = { api: api }"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_comparison()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_explanation_export()"));
        assert!(
            super::ORIGINAL_HTML
                .contains("bridge.api.import_central_explanation_feedback(picked.paths)")
        );
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_indicator_split()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_cross_period()"));
        assert!(super::ORIGINAL_HTML.contains("onclick=\"runCentralForms()\""));
        assert!(super::ORIGINAL_HTML.contains("centralFormsConfig:'config'"));
        assert!(super::ORIGINAL_HTML.contains("最近输出目录："));
    }

    #[test]
    fn s3_f5_accepts_3_3_xlsx_config_and_can_clear_it() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s3_f5_picker_{suffix}"));
        fs::create_dir_all(&root).unwrap();
        let config = root.join("3.3配置.xlsx");
        let unsupported = root.join("3.3配置.xlsm");
        fs::write(&config, b"picker fixture").unwrap();
        fs::write(&unsupported, b"unsupported picker fixture").unwrap();
        let state = test_state();

        dispatch_api(
            &state,
            "choose_central",
            &[json!("centralFormsConfig"), json!(config.to_string_lossy())],
        )
        .expect("select S3-F05 3.3 XLSX config");
        assert_eq!(
            state.get_state()["centralFormsConfig"],
            json!(fs::canonicalize(&config).unwrap().to_string_lossy())
        );
        let error = dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralFormsConfig"),
                json!(unsupported.to_string_lossy()),
            ],
        )
        .expect_err("3.3 form configuration is XLSX-only");
        assert!(error.message.contains("只支持 .xlsx"));
        dispatch_api(&state, "clear_path", &[json!("centralFormsConfig")])
            .expect("clear S3-F05 config selection");
        assert_eq!(state.get_state()["centralFormsConfig"], "");

        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn s3_f04_picker_accepts_only_its_supported_input_formats() {
        let extensions = super::browse_extensions("central_data");
        assert!(super::ORIGINAL_HTML.contains("centralCrossCur:'central_data'"));
        for filename in ["current.csv", "current.xlsx", "current.xls"] {
            assert!(
                super::extension_allowed(Path::new(filename), extensions),
                "{filename} should be selectable"
            );
        }
        for filename in ["current.xlsm", "current.txt"] {
            assert!(
                !super::extension_allowed(Path::new(filename), extensions),
                "{filename} should be rejected"
            );
        }
    }

    #[test]
    fn s5_gui_path_selection_scans_workbooks_and_rejects_other_formats() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s5-selection-{suffix}"));
        fs::create_dir_all(&root).expect("create S5 selection fixture directory");
        let source = root.join("current.xlsx");
        let ignored = root.join("~$locked.xlsx");
        let other_format = root.join("not-a-workbook.csv");
        write_s1_region_source(&source);
        write_s1_region_source(&ignored);
        fs::write(&other_format, "ignored").expect("write non-Excel file");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_s5",
            &[json!("input_dir"), json!(root.to_string_lossy())],
        )
        .expect("select S5 input directory");
        let form = state.get_state()["s5"].clone();
        assert_eq!(form["files"].as_array().unwrap().len(), 1);
        assert_eq!(form["selected_files"].as_array().unwrap().len(), 1);
        assert_eq!(form["files"][0]["name"], "current.xlsx");

        let error = dispatch_api(
            &state,
            "choose_s5",
            &[
                json!("external_path"),
                json!(other_format.to_string_lossy()),
            ],
        )
        .expect_err("reject CSV as S5 workbook");
        assert!(error.message.contains("仅支持 XLSX/XLSM"));
        fs::remove_dir_all(&root).expect("remove S5 selection fixture directory");
    }

    #[test]
    fn s5_gui_previous_directory_selects_supported_workbooks_for_identity_matching() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s5-previous-{suffix}"));
        let previous_dir = root.join("previous");
        fs::create_dir_all(&previous_dir).expect("create S5 previous directory");
        let previous = previous_dir.join("previous.xlsx");
        let ignored = previous_dir.join("notes.txt");
        write_s1_region_source(&previous);
        fs::write(&ignored, "ignored").expect("write unrelated file");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_s5",
            &[json!("previous_dir"), json!(previous_dir.to_string_lossy())],
        )
        .expect("select S5 previous directory");
        let form = state.get_state()["s5"].clone();
        assert_eq!(form["previous_files"].as_array().unwrap().len(), 1);
        let expected_previous = fs::canonicalize(previous)
            .unwrap()
            .to_string_lossy()
            .into_owned();
        assert_eq!(
            form["previous_files"][0].as_str(),
            Some(expected_previous.as_str())
        );

        dispatch_api(
            &state,
            "update_s5",
            &[json!({"previous_period": "2026-06-30"})],
        )
        .expect("save previous report period");
        assert_eq!(state.get_state()["s5"]["previous_period"], "2026-06-30");
        fs::remove_dir_all(&root).expect("remove S5 previous fixture directory");
    }

    #[test]
    fn s3_current_value_check_runs_through_the_web_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_current_value_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3-F04 temporary directory");
        let current = root.join("current.csv");
        let xlsx_current = root.join("current.xlsx");
        let common_config = root.join("3.0大集中通用配置.xlsx");
        let cross_config = root.join("3.2大集中本期数值核对_配置.xlsx");
        let output = root.join("output");
        fs::create_dir_all(&output).expect("create S3-F04 output directory");
        let csv = format!(
            "{}\n人民币,2026-08-31,0001,机构,1234567,地区,1,I1,指标,余额,人民币,月,2,200000000\n",
            EXPECTED_HEADERS.join(",")
        );
        fs::write(&current, &csv).expect("write synthetic current CSV");
        write_s3_current_value_input_xlsx(&xlsx_current);
        write_s3_current_value_common_config(&common_config);
        write_s3_current_value_cross_config(&cross_config);

        let state = test_state();
        let selected_xlsx = dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCrossCur"),
                json!(xlsx_current.to_string_lossy()),
            ],
        )
        .expect("XLSX input should be selectable for S3-F04");
        assert_eq!(
            selected_xlsx,
            json!(fs::canonicalize(&xlsx_current).unwrap().to_string_lossy())
        );

        for (field, path) in [
            ("centralCrossCur", &current),
            ("centralCommonConfig", &common_config),
            ("centralCrossConfig", &cross_config),
            ("centralOutput", &output),
        ] {
            dispatch_api(
                &state,
                "choose_central",
                &[json!(field), json!(path.to_string_lossy())],
            )
            .unwrap_or_else(|error| panic!("select {field}: {}", error.message));
        }
        dispatch_api(&state, "update", &[json!({"exportRunLogs": true})])
            .expect("enable run-log workbook for the integration check");
        assert_eq!(
            dispatch_api(&state, "start_central_cross_period", &[])
                .expect("start S3-F04 through its Web API"),
            json!(true)
        );

        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(final_state["status"], "本期数值核对完成：1 行");
        let result_paths = fs::read_dir(&output)
            .unwrap()
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with("数值核对结果_"))
            })
            .collect::<Vec<_>>();
        assert_eq!(result_paths.len(), 1);
        let canonical_result = fs::canonicalize(&result_paths[0])
            .unwrap()
            .to_string_lossy()
            .into_owned();
        assert_eq!(
            final_state["centralCrossLastResult"].as_str(),
            Some(canonical_result.as_str())
        );
        let result_sheet = excel_reader::read_value_columns(&result_paths[0], 17)
            .unwrap()
            .remove(0);
        assert_eq!(result_sheet.max_row, 2);
        assert_eq!(
            result_sheet.values.get("O2"),
            Some(&excel_reader::CellValue::Text(
                "一边有，一边无;请核实".to_owned()
            ))
        );
        assert_eq!(
            fs::read_dir(&output)
                .unwrap()
                .filter_map(|entry| entry.ok().map(|entry| entry.path()))
                .filter(|path| {
                    path.file_name()
                        .and_then(|name| name.to_str())
                        .is_some_and(|name| name.starts_with("本期数值核对_运行日志_"))
                })
                .count(),
            1
        );
        let logs = final_state["log"].as_array().unwrap();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine/csv")
                        && text.contains("writer=rust_xlsxwriter")
                        && text.contains("工作簿写入=")
                })
        }));
        assert_eq!(fs::read(&current).unwrap(), csv.as_bytes());
        let xlsx_before = fs::read(&xlsx_current).unwrap();
        assert!(fs::read(&common_config).is_ok());
        assert!(fs::read(&cross_config).is_ok());

        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCrossCur"),
                json!(xlsx_current.to_string_lossy()),
            ],
        )
        .expect("select the XLSX input for an end-to-end Web run");
        dispatch_api(&state, "start_central_cross_period", &[])
            .expect("start S3-F04 XLSX run through its Web API");
        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let xlsx_state = state.get_state();
        assert_eq!(xlsx_state["busy"], false);
        assert_eq!(xlsx_state["status"], "本期数值核对完成：1 行");
        assert!(xlsx_state["log"].as_array().unwrap().iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine") && text.contains("结果行=1")
                })
        }));
        assert_eq!(fs::read(&xlsx_current).unwrap(), xlsx_before);
        fs::remove_dir_all(root).expect("remove S3-F04 temporary files");
    }

    #[test]
    fn s3_config_check_apis_report_structure_without_running_a_batch() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_config_check_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3 config-check directory");
        let common = root.join("3.0大集中通用配置.xlsx");
        let feature = root.join("3.1大集中执行比较_配置.xlsx");
        let cross = root.join("3.2大集中本期数值核对_配置.xlsx");
        write_s3_full_common_config(&common);
        write_s3_execution_feature_config(&feature);
        write_s3_current_value_cross_config(&cross);

        let state = test_state();
        for (field, path) in [
            ("centralCommonConfig", &common),
            ("centralComparisonConfig", &feature),
            ("centralCrossConfig", &cross),
        ] {
            dispatch_api(
                &state,
                "choose_central",
                &[json!(field), json!(path.to_string_lossy())],
            )
            .unwrap_or_else(|error| panic!("select {field}: {}", error.message));
        }

        let common_check = dispatch_api(&state, "check_central_common_config", &[])
            .expect("3.0 config check must be implemented");
        assert_eq!(common_check["passed"], true);
        let common_text = common_check["report_text"].as_str().unwrap();
        assert!(common_text.contains("3.0 大集中通用配置检查通过"));
        assert!(common_text.contains("环比警戒区间：1 条"));

        let expression_check = dispatch_api(&state, "check_central_expression_config", &[])
            .expect("3.1 config check must be implemented");
        assert_eq!(expression_check["passed"], true);
        let expression_text = expression_check["report_text"].as_str().unwrap();
        assert!(expression_text.contains("3.1 大集中执行比较配置检查通过"));
        assert!(expression_text.contains("动作规则：0 条"));
        assert!(expression_text.contains("表达式规则：1 条"));

        let cross_check = dispatch_api(&state, "check_central_cross_formulas", &[])
            .expect("3.2 config check must be implemented");
        assert_eq!(cross_check["passed"], true);
        let cross_text = cross_check["report_text"].as_str().unwrap();
        assert!(cross_text.contains("3.2 大集中本期数值核对配置检查通过"));
        assert!(cross_text.contains("核对规则：1 条（启用 1，停用 0）"));
        assert!(cross_text.contains("单位换算：元 → 亿元"));

        // Checks are read-only: no busy task, no result files, inputs untouched.
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(fs::read_dir(&root).unwrap().count(), 3);
        assert_eq!(
            final_state["log"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|entry| entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("配置检查通过")))
                .count(),
            3
        );
        fs::remove_dir_all(root).expect("remove S3 config-check files");
    }

    /// AUDIT_STATE_DIR is process-global; tests that touch it must serialize.
    static ADVANCED_SETTINGS_ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn advanced_settings_update_persists_and_logs_selection() {
        let _guard = ADVANCED_SETTINGS_ENV_LOCK.lock().unwrap();
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let state_dir = std::env::temp_dir().join(format!(
            "audit_web_adv_settings_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&state_dir).expect("create isolated settings directory");
        std::env::set_var("AUDIT_STATE_DIR", &state_dir);

        let state = test_state();
        let updated = dispatch_api(
            &state,
            "update",
            &[json!({"advancedSettings": {"calcEngine": "wps", "editEngine": "umya"}})],
        )
        .expect("valid advanced-settings update must succeed");
        assert_eq!(updated["advancedSettings"]["calcEngine"], "wps");
        assert_eq!(updated["advancedSettings"]["editEngine"], "umya");
        let logs = updated["log"].as_array().unwrap();
        assert!(logs.iter().any(|entry| {
            entry["text"]
                .as_str()
                .is_some_and(|text| text.contains("高级设置已更新") && text.contains("WPS 表格"))
        }));
        let settings_file = state_dir.join("settings.json");
        assert!(settings_file.is_file(), "selection must be persisted");
        let persisted: Value =
            serde_json::from_slice(&fs::read(&settings_file).expect("read settings file"))
                .expect("valid settings JSON");
        assert_eq!(persisted["calc_engine"], "wps");

        // The next state build reflects the persisted selection (worker threads
        // read the file, not the GUI state).
        assert_eq!(
            crate::app_settings::AppSettings::load().calc_engine.as_str(),
            "wps"
        );

        std::env::remove_var("AUDIT_STATE_DIR");
        fs::remove_dir_all(&state_dir).expect("remove isolated settings directory");
    }

    #[test]
    fn advanced_settings_rejects_unknown_engines_and_unavailable_openpyxl() {
        let _guard = ADVANCED_SETTINGS_ENV_LOCK.lock().unwrap();
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let state_dir = std::env::temp_dir().join(format!(
            "audit_web_adv_settings_reject_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&state_dir).expect("create isolated settings directory");
        std::env::set_var("AUDIT_STATE_DIR", &state_dir);

        let state = test_state();
        let invalid = dispatch_api(
            &state,
            "update",
            &[json!({"advancedSettings": {"calcEngine": "pywin32"}})],
        )
        .expect_err("unknown calc engine must be rejected");
        assert_eq!(invalid.status, StatusCode::BAD_REQUEST);
        assert!(invalid.message.contains("auto、excel、wps 或 libreoffice"));

        let openpyxl_error = dispatch_api(
            &state,
            "update",
            &[json!({"advancedSettings": {"editEngine": "openpyxl"}})],
        )
        .expect_err("openpyxl without positive evidence must be rejected");
        assert_eq!(openpyxl_error.status, StatusCode::BAD_REQUEST);
        assert!(
            openpyxl_error.message.contains("openpyxl 修改引擎不可用"),
            "rejection must carry the evidence reason: {}",
            openpyxl_error.message
        );
        // A rejected store must not change the persisted selection.
        assert_eq!(
            crate::app_settings::AppSettings::load().calc_engine.as_str(),
            "auto"
        );

        std::env::remove_var("AUDIT_STATE_DIR");
        fs::remove_dir_all(&state_dir).expect("remove isolated settings directory");
    }

    #[test]
    fn s3_config_check_apis_require_a_selected_config_file() {
        let state = test_state();
        for (method, message) in [
            ("check_central_common_config", "3.0 大集中通用配置"),
            ("check_central_expression_config", "3.0 大集中通用配置"),
            ("check_central_cross_formulas", "3.2 大集中本期数值核对配置"),
        ] {
            let error = dispatch_api(&state, method, &[]).expect_err("missing config must fail");
            assert_eq!(error.status, StatusCode::BAD_REQUEST);
            assert!(
                error.message.contains(message),
                "{method}: unexpected message {}",
                error.message
            );
        }
    }

    #[test]
    fn s2_web_run_persists_stage_engine_and_elapsed_detail_logs() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s2_progress_{}_{}",
            std::process::id(),
            suffix
        ));
        let current = root.join("current");
        let previous = root.join("previous");
        let output = root.join("output");
        fs::create_dir_all(&current).expect("create S2 current directory");
        fs::create_dir_all(&previous).expect("create S2 previous directory");
        fs::create_dir_all(&output).expect("create S2 output directory");

        let fixture_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let source_report = fixture_root.join("s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
        fs::copy(
            &source_report,
            current.join("9144#2026-06-30#x#20202#甲银行.xlsx"),
        )
        .expect("copy current S2 fixture");
        fs::copy(
            &source_report,
            previous.join("9144#2026-03-31#x#20202#甲银行.xlsx"),
        )
        .expect("copy previous S2 fixture");
        let config = fixture_root.join("s2_config/check_valid.xlsx");
        let external = fixture_root.join("s2_external_import/plain_layout.xlsx");
        let state = AppState::new(config.clone());

        for (kind, path) in [
            ("pcCurDir", current.as_path()),
            ("pcPreDir", previous.as_path()),
            ("pcConfig", config.as_path()),
            ("pcCentral", external.as_path()),
            ("pcOutput", output.as_path()),
        ] {
            dispatch_api(
                &state,
                "choose_period_compare",
                &[json!(kind), json!(path.to_string_lossy())],
            )
            .expect("select S2 Web API input path");
        }
        assert_eq!(
            dispatch_api(&state, "start_period_compare_v2", &[]).expect("start S2 Web API run"),
            json!(true)
        );

        for _ in 0..400 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(final_state["periodProgress"]["status"], "completed");
        let report = PathBuf::from(final_state["periodSummary"]["result"].as_str().unwrap());
        assert!(report.is_file());

        let logs = final_state["log"].as_array().expect("S2 Web logs");
        for stage in [
            "config",
            "current",
            "previous",
            "validate",
            "period",
            "rules",
            "external_read",
            "external",
            "export",
            "save",
            "complete",
        ] {
            let engine = crate::s2_service::s2_engine_for_stage(stage).unwrap();
            assert!(
                logs.iter().any(|entry| {
                    entry["detail"] == true
                        && entry["text"].as_str().is_some_and(|text| {
                            text.contains(&format!("S2 阶段开始：stage={stage}"))
                                && text.contains(&format!("engine={engine}"))
                        })
                }),
                "missing start detail for S2 stage {stage}"
            );
            assert!(
                logs.iter().any(|entry| {
                    entry["detail"] == true
                        && entry["text"].as_str().is_some_and(|text| {
                            text.contains(&format!("S2 阶段完成：stage={stage}"))
                                && text.contains(&format!("engine={engine}"))
                                && text.contains("elapsed_ms=")
                        })
                }),
                "missing completion detail for S2 stage {stage}"
            );
        }
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S2 任务完成：reader=Calamine"))
        }));
        fs::remove_dir_all(root).expect("remove temporary S2 Web API files");
    }

    #[test]
    fn s3_explanation_export_runs_through_the_existing_frontend_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_explanation_{}_{}",
            std::process::id(),
            suffix
        ));
        let comparison = root.join("comparison.xlsx");
        let common_config = root.join("3.0大集中通用配置.xlsx");
        let output = root.join("output");
        fs::create_dir_all(&output).expect("create S3 output directory");
        write_s3_comparison(&comparison);
        write_s3_common_config(&common_config);

        let state = test_state();
        let selected = dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCompareFile"),
                json!(comparison.to_string_lossy()),
            ],
        )
        .expect("select comparison file through the Web API");
        assert_eq!(
            selected,
            json!(fs::canonicalize(&comparison).unwrap().to_string_lossy())
        );
        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCommonConfig"),
                json!(common_config.to_string_lossy()),
            ],
        )
        .expect("select common configuration through the Web API");
        dispatch_api(
            &state,
            "choose_central",
            &[json!("centralOutput"), json!(output.to_string_lossy())],
        )
        .expect("select output directory through the Web API");
        assert_eq!(
            dispatch_api(&state, "start_central_explanation_export", &[])
                .expect("start S3-F02 through the Web API"),
            json!(true)
        );

        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(
            final_state["status"],
            "指标说明导出完成：1 条",
            "S3 Web API logs: {}",
            serde_json::to_string(&final_state["log"]).unwrap()
        );
        let output_path = PathBuf::from(final_state["centralExplainFile"].as_str().unwrap());
        assert!(output_path.is_file());
        let canonical_output = fs::canonicalize(&output).expect("canonical output directory");
        assert_eq!(output_path.parent(), Some(canonical_output.as_path()));
        assert!(output_path
            .file_name()
            .unwrap()
            .to_string_lossy()
            .starts_with("指标说明_"));
        let output_sheets = excel_reader::read_value_columns(&output_path, 8)
            .expect("read generated explanation workbook");
        let explanation = output_sheets
            .iter()
            .find(|sheet| sheet.name == "说明文件")
            .expect("explanation worksheet");
        assert_eq!(
            explanation.values.get("B3"),
            Some(&excel_reader::CellValue::Text(
                "配置机构名称（示例地区）".to_owned()
            )),
            "the Web entry must apply the selected 3.0 common organization mapping"
        );
        let logs = final_state["log"].as_array().unwrap();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == false
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("导出完成：1 条"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine")
                        && text.contains("writer=rust_xlsxwriter")
                        && text.contains("配置读取=")
                })
        }));
        fs::remove_dir_all(root).expect("remove temporary S3 Web API files");
    }

    #[test]
    fn s3_explanation_export_requires_the_python_compatible_common_config() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_missing_config_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3 input directory");
        let comparison = root.join("comparison.xlsx");
        write_s3_comparison(&comparison);

        let state = test_state();
        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCompareFile"),
                json!(comparison.to_string_lossy()),
            ],
        )
        .expect("select comparison file through the Web API");
        let error = dispatch_api(&state, "start_central_explanation_export", &[])
            .expect_err("export must require the common config used by Python");

        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("大集中通用配置"));
        assert_eq!(state.get_state()["busy"], false);
        fs::remove_dir_all(root).expect("remove temporary S3 missing-config files");
    }

    #[test]
    fn state_reports_s2_paths_and_unmigrated_status() {
        let state = test_state();
        let value = state.get_state();
        assert_eq!(value["distributionProfile"], "s1s2");
        assert!(value["pcConfig"].is_string());
        assert_eq!(value["filePickerMode"], "自动");
        assert_eq!(value["busy"], false);
        assert_eq!(value["template"], "");
    }

    #[test]
    fn s2_progress_view_includes_stage_engine_and_elapsed_time() {
        let event = S2ProgressEvent {
            stage: "current".to_owned(),
            label: "读取本期报表".to_owned(),
            percent: 12.0,
            current: 1,
            total: 2,
            message: "读取工作簿".to_owned(),
        };
        let value = progress_value(&event, "running", 1.5, 0.25);

        assert_eq!(value["engine"], "Calamine");
        assert_eq!(value["elapsedSeconds"], 1.5);
        assert_eq!(value["stageElapsedSeconds"], 0.25);
        assert!(value["label"].as_str().unwrap().contains("Calamine"));
        assert!(value["label"].as_str().unwrap().contains("本阶段已运行"));
    }

    #[test]
    fn s1_template_inspection_is_reachable_through_the_existing_ui_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "audit_web_template_inspection_{}_{}.xlsx",
            std::process::id(),
            suffix
        ));
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("set sheet");
        sheet
            .write_formula(0, 0, "=1+#REF!")
            .expect("write formula");
        workbook
            .define_name("校验区域_1", "='检查表'!$A$1")
            .expect("define named range");
        workbook.save(&path).expect("save workbook");

        let state = test_state();
        let chosen = dispatch_api(
            &state,
            "choose_file",
            &[json!("template"), json!(path.to_string_lossy())],
        )
        .expect("choose template");
        assert_eq!(chosen, json!(path.to_string_lossy()));
        let report =
            dispatch_api(&state, "inspect_template_formulas", &[]).expect("inspect template");
        assert_eq!(report["template"], path.to_string_lossy().as_ref());
        assert_eq!(report["formulaCount"], 1);
        assert_eq!(report["errors"][0]["kind"], "公式");
        assert_eq!(report["errors"][0]["token"], "#REF!");
        assert_eq!(report["namedRangeGroups"][0]["count"], 1);
        let logs = state.get_state()["log"].as_array().unwrap().clone();
        assert!(logs.iter().any(|entry| entry["detail"] == false));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("calamine+namespace_aware_ooxml"))
        }));
        fs::remove_file(path).expect("remove workbook");
    }

    #[test]
    fn s1_template_inspection_respects_busy_state() {
        let state = test_state();
        state.lock().value["busy"] = json!(true);
        let error = dispatch_api(&state, "inspect_template_formulas", &[]).unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::CONFLICT);
    }

    #[test]
    fn s2_config_check_logs_user_result_and_developer_reader_timing() {
        let config = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s2_config")
            .join("check_valid.xlsx");
        let state = AppState::new(config);

        let report = dispatch_api(&state, "check_period_audit_config", &[])
            .expect("check S2 config through the Web API");
        assert_eq!(report["passed"], true);

        let logs = state.get_state()["log"].as_array().unwrap().clone();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == false
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S2 配置检查通过"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("S2-F01 配置检查步骤")
                        && text.contains("reader=calamine")
                        && text.contains("指标=")
                        && text.contains("elapsed_ms=")
                })
        }));
    }

    #[test]
    fn path_browser_returns_the_frontend_contract() {
        let state = test_state();
        let value = browse_dir(&state, &[json!(""), json!("report_data")]).unwrap();
        assert!(value["path"].is_string());
        assert!(value["parent"].is_string());
        assert!(value["dirs"].is_array());
        assert!(value["files"].is_array());
        assert!(value["selectedFile"].is_string());
        assert!(value["error"].is_string());
    }

    #[test]
    fn s3_execution_compare_reports_missing_gui_inputs() {
        let error = dispatch_api(&test_state(), "start_central_comparison", &[]).unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("本期大集中数据"));
    }

    #[test]
    fn s3_execution_compare_runs_through_the_existing_gui_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_execution_compare_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3-F01 web test directory");
        let current = root.join("current.csv");
        let previous = root.join("previous.csv");
        let common = root.join("3.0大集中通用配置.xlsx");
        let feature = root.join("3.1大集中执行比较_配置.xlsx");
        write_s3_execution_csv(&current, "2026-08-31", "1000");
        write_s3_execution_csv(&previous, "2026-07-31", "900");
        write_s3_execution_configs(&common, &feature);
        let current_before = fs::read(&current).unwrap();
        let previous_before = fs::read(&previous).unwrap();
        let common_before = fs::read(&common).unwrap();
        let feature_before = fs::read(&feature).unwrap();

        let state = test_state();
        for (field, path) in [
            ("centralCur", &current),
            ("centralPre", &previous),
            ("centralCommonConfig", &common),
            ("centralComparisonConfig", &feature),
        ] {
            dispatch_api(
                &state,
                "choose_central",
                &[json!(field), json!(path.to_string_lossy())],
            )
            .unwrap_or_else(|error| panic!("select {field}: {}", error.message));
        }
        dispatch_api(
            &state,
            "update",
            &[json!({"centralExpressionSchema": "FIVE_SEGMENT_V1"})],
        )
        .expect("set the legacy GUI expression schema");
        assert_eq!(
            dispatch_api(&state, "start_central_comparison", &[]).unwrap(),
            json!(true)
        );

        let started = std::time::Instant::now();
        while state.get_state()["busy"] == true && started.elapsed() < Duration::from_secs(10) {
            std::thread::sleep(Duration::from_millis(10));
        }
        let value = state.get_state();
        assert_eq!(value["busy"], false, "background GUI task should finish");
        assert_eq!(value["status"], "执行比较完成：1 行");
        let output = PathBuf::from(
            value["centralCompareFile"]
                .as_str()
                .expect("GUI should publish comparison output path"),
        );
        assert!(
            output.is_file(),
            "GUI task should create the result workbook"
        );
        assert_eq!(
            value["centralFormOutput"],
            output.parent().unwrap().to_string_lossy().as_ref()
        );
        let logs = value["log"].as_array().expect("GUI log array");
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("writer=rust_xlsxwriter"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S3-F01 完成：reader=calamine/csv"))
        }));
        assert_eq!(fs::read(&current).unwrap(), current_before);
        assert_eq!(fs::read(&previous).unwrap(), previous_before);
        assert_eq!(fs::read(&common).unwrap(), common_before);
        assert_eq!(fs::read(&feature).unwrap(), feature_before);
        fs::remove_dir_all(&root).expect("remove synthetic S3-F01 web test directory");
    }

    #[test]
    fn s1_f2_runs_through_the_existing_gui_workflow_without_modifying_inputs() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s1-f2-{suffix}"));
        let input_dir = root.join("input");
        let template_dir = root.join("templates");
        let config_dir = root.join("config");
        fs::create_dir_all(&input_dir).expect("create S1-F02 input directory");
        fs::create_dir_all(&template_dir).expect("create S1-F02 template directory");
        fs::create_dir_all(&config_dir).expect("create S1-F02 config directory");

        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s1_region_summary_template.xlsx");
        let template = template_dir.join("summary_template.xlsx");
        let history = config_dir.join("逐笔统计系统_配置.xlsx");
        let source = input_dir.join("机构甲_20260731.xlsx");
        fs::copy(&fixture, &template).expect("copy template fixture");
        fs::copy(&fixture, &history).expect("copy history config fixture");
        write_s1_region_source(&source);
        let source_before = fs::read(&source).expect("snapshot S1 source");
        let history_before = fs::read(&history).expect("snapshot history config");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_folder",
            &[json!("input"), json!(input_dir.to_string_lossy())],
        )
        .expect("choose S1-F02 input directory");
        dispatch_api(
            &state,
            "choose_folder",
            &[json!("templateDir"), json!(template_dir.to_string_lossy())],
        )
        .expect("choose S1-F02 template directory");
        dispatch_api(
            &state,
            "update",
            &[json!({
                "input": input_dir.to_string_lossy(),
                "templateDir": template_dir.to_string_lossy()
            })],
        )
        .expect("update selected S1-F02 paths");
        dispatch_api(&state, "recognize", &[]).expect("scan S1-F02 source workbooks");
        assert_eq!(
            state.get_state()["sourceFiles"].as_array().unwrap().len(),
            1
        );
        assert_eq!(
            state.get_state()["selectedFiles"].as_array().unwrap().len(),
            1
        );
        dispatch_api(
            &state,
            "choose_file",
            &[json!("template"), json!(template.to_string_lossy())],
        )
        .expect("select S1-F02 template");
        dispatch_api(
            &state,
            "choose_history_config",
            &[json!(history.to_string_lossy())],
        )
        .expect("select S1 history config");

        let selected = state.get_state()["selectedFiles"].clone();
        dispatch_api(
            &state,
            "start_dag_flow",
            &[
                json!("dag:汇总校验结果说明"),
                json!({
                    "input": input_dir.to_string_lossy(),
                    "templateDir": template_dir.to_string_lossy(),
                    "template": template.to_string_lossy(),
                    "output": ""
                }),
                selected,
            ],
        )
        .expect("start S1-F02 through the original GUI API");

        let started = std::time::Instant::now();
        while state.get_state()["busy"] == true && started.elapsed() < Duration::from_secs(10) {
            std::thread::sleep(Duration::from_millis(10));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert!(
            final_state["status"]
                .as_str()
                .unwrap_or_default()
                .starts_with("S1-F02 汇总完成"),
            "{}",
            final_state["status"]
        );
        let output = PathBuf::from(
            final_state["auditSummaryOutput"]
                .as_str()
                .expect("S1-F02 output path"),
        );
        assert!(output.is_file());
        assert_eq!(
            output.parent().map(fs::canonicalize).transpose().unwrap(),
            Some(fs::canonicalize(input_dir.join("执行结果")).unwrap())
        );
        assert_eq!(
            super::s1_region_summary_source_workbooks(&input_dir, true)
                .expect("rediscover S1 sources after output")
                .len(),
            1,
            "generated execution output must not become a later source file"
        );
        assert!(!excel_reader::read_workbook(&output)
            .expect("read generated S1-F02 result")
            .sheets
            .is_empty());
        assert_eq!(
            fs::read(&source).expect("read source after S1-F02"),
            source_before
        );
        assert_eq!(
            fs::read(&history).expect("read history after S1-F02"),
            history_before
        );
        assert!(final_state["log"].as_array().unwrap().iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("reader=calamine+namespace_aware_ooxml"))
        }));
        fs::remove_dir_all(&root).expect("remove synthetic S1-F02 web test directory");
    }

    #[test]
    fn s1_f2_gui_workflow_rejects_missing_history_configuration() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s1-f2-no-history-{suffix}"));
        let input_dir = root.join("input");
        fs::create_dir_all(&input_dir).expect("create S1-F02 input directory");
        let template = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s1_region_summary_template.xlsx");
        write_s1_region_source(&input_dir.join("source.xlsx"));
        let state = test_state();
        let error = dispatch_api(
            &state,
            "start_dag_flow",
            &[
                json!("dag:汇总校验结果说明"),
                json!({
                    "input": input_dir.to_string_lossy(),
                    "template": template.to_string_lossy()
                }),
                json!([]),
            ],
        )
        .unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("逐笔统计系统配置"));
        assert_eq!(state.get_state()["busy"], false);
        fs::remove_dir_all(&root).expect("remove S1-F02 missing config test directory");
    }

    #[test]
    fn invalid_path_kind_is_rejected() {
        let error = dispatch_api(
            &test_state(),
            "choose_period_compare",
            &[
                Value::String("unknown".to_owned()),
                Value::String("/tmp".to_owned()),
            ],
        )
        .unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
    }

    #[test]
    fn choosing_and_clearing_period_directories_refreshes_pair_preview() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-pairs-{suffix}"));
        let current = root.join("current");
        let previous = root.join("previous");
        fs::create_dir_all(&current).expect("create current directory");
        fs::create_dir_all(&previous).expect("create previous directory");
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
        fs::copy(
            &fixture,
            current.join("9144#2026-06-30#x#20202#甲银行.xlsx"),
        )
        .expect("copy current workbook");
        fs::copy(
            &fixture,
            previous.join("9144#2026-03-31#x#20202#甲银行.xlsx"),
        )
        .expect("copy previous workbook");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_period_compare",
            &[json!("pcCurDir"), json!(current.to_string_lossy())],
        )
        .expect("select current directory");
        let current_only = state.get_state();
        assert_eq!(current_only["periodPairs"].as_array().unwrap().len(), 2);
        assert!(current_only["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["side"] == "cur"));

        dispatch_api(
            &state,
            "choose_period_compare",
            &[json!("pcPreDir"), json!(previous.to_string_lossy())],
        )
        .expect("select previous directory");
        let matched = state.get_state();
        assert!(matched["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["matched"] == true));
        assert_eq!(matched["periodPairs"][0]["curDate"], "2026-06-30");
        assert_eq!(matched["periodPairs"][0]["preDate"], "2026-03-31");

        dispatch_api(&state, "clear_path", &[json!("pcPreDir")]).expect("clear previous");
        let cleared = state.get_state();
        assert!(cleared["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["side"] == "cur"));
        let _ = fs::remove_dir_all(root);
    }
}
