//! Local HTTP bridge for the bundled, unchanged Audit web frontend.
//!
//! This bridge exposes migrated Audit workflows through the existing frontend.

#[cfg(test)]
use crate::s1_region_summary::s1_region_summary_source_workbooks;
use crate::{
    discovery::{detect_period_from_workbooks, organisation_from_name, ScanDepth},
    models::{CopyRange, Issue},
    name_config::{
        load_s1_conditional_mappings, named_areas_from_defined_names,
        resolve_s1_conditional_area_plan, FeatureMapping,
    },
    native_file_dialog::{PickKind, PickRequest},
    period_pairs::{list_period_pairs_with_scan_depth, report_source_warning_with_scan_depth},
    s1_conditional_audit::{
        extract_conditional_issues, extract_conditional_issues_with_display_failure,
        extract_conditional_issues_with_display_scan, S1ConditionalAuditResult,
        S1ConditionalMetadata,
    },
    s1_formula_audit::{extract_s1_formula_issues, S1FormulaMetadata},
    s1_formula_copy::{
        execute_s1_formula_copy_batch_with_manual_recovery,
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
    s1_region_summary::{
        s1_region_summary_source_workbooks_with_depth,
        summarize_s1_named_regions_selected_with_progress,
    },
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
    s2_service::{
        run_s2_audit_with_scan_depth, s2_engine_for_stage, S2ProgressEvent, S2RunRequest,
    },
    s3_current_value_check::{
        input_reader_for_paths, python_regex_scope_matcher,
        run_current_value_check_files_with_progress,
    },
    s3_execution_compare::{run_execution_compare_files, S3ExpressionSchema},
    s3_explanation_export::export_selected_explanations_with_progress,
    s3_explanation_feedback::execute_explanation_feedback_import_with_progress,
    s3_indicator_split::execute_indicator_comparison_split_with_progress,
    s3_indicator_split::fast_render_mode,
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
    collections::VecDeque,
    fs,
    net::{IpAddr, Ipv4Addr, SocketAddr},
    path::{Path, PathBuf},
    sync::{Arc, Mutex, MutexGuard, OnceLock},
    thread,
    time::{Duration, Instant},
};
use tokio::{net::TcpListener, runtime::Builder};

const ORIGINAL_HTML: &str = include_str!("../frontend/web/index.html");
const APP_ICON_PNG: &[u8] = include_bytes!("../frontend/assets/app-icon.png");
const APP_ICON_ICO: &[u8] = include_bytes!("../frontend/assets/app-icon.ico");
/// 界面标题「审核工具 V<版本>」，版本来自 VERSION_MANIFEST.json。
fn app_title() -> &'static str {
    static TITLE: std::sync::OnceLock<String> = std::sync::OnceLock::new();
    TITLE.get_or_init(crate::app_identity::app_title)
}

fn system_environment_info(openpyxl_status: &(bool, String)) -> Value {
    static INFO: OnceLock<Value> = OnceLock::new();
    INFO.get_or_init(|| collect_system_environment_info(openpyxl_status))
        .clone()
}

fn collect_system_environment_info(openpyxl_status: &(bool, String)) -> Value {
    let logical_processors = std::thread::available_parallelism()
        .map(|count| count.get())
        .unwrap_or(0);
    let architecture = match std::env::consts::ARCH {
        "x86_64" => "x86_64（AMD64）".to_owned(),
        "aarch64" => "aarch64（ARM64）".to_owned(),
        "x86" => "x86（32 位）".to_owned(),
        other => other.to_owned(),
    };
    let mut hardware = if logical_processors == 0 {
        "逻辑处理器数量不可用".to_owned()
    } else {
        format!("逻辑处理器 {logical_processors} 个")
    };
    #[cfg(target_os = "linux")]
    if let Some(memory_gib) = linux_memory_gib() {
        hardware.push_str(&format!("；物理内存约 {memory_gib:.1} GiB"));
    }
    hardware.push_str("；显卡信息未采集");

    let (excel, wps, com) = if cfg!(windows) {
        (
            "Windows 可选 COM 引擎；诊断页不会启动 Excel".to_owned(),
            "Windows 可选 COM 引擎；诊断页不会启动 WPS".to_owned(),
            "本机 COM 接口；任务执行时按配置尝试".to_owned(),
        )
    } else {
        (
            "不适用（Windows 专用引擎）".to_owned(),
            "不适用（Windows 专用引擎）".to_owned(),
            "不适用（当前系统非 Windows）".to_owned(),
        )
    };
    let openpyxl = openpyxl_status.1.clone();
    let python = executable_in_path(&["python.exe", "python3", "python"])
        .map(|path| format!("检测到可选 Python：{}；主程序不依赖 Python", path.display()))
        .unwrap_or_else(|| "PATH 未检测到 Python；主程序不依赖 Python".to_owned());
    let libreoffice = executable_in_path(&["soffice.exe", "soffice", "libreoffice"])
        .map(|path| format!("检测到：{}", path.display()))
        .unwrap_or_else(|| "PATH 未检测到 soffice/libreoffice".to_owned());

    json!({
        "app": format!(
            "{}；{} 构建；本机服务 + 默认浏览器",
            app_title(),
            if cfg!(debug_assertions) { "Debug" } else { "Release" }
        ),
        "os": host_os_description(),
        "architecture": architecture,
        "locale": host_locale_description(),
        "glibc": host_glibc_description(),
        "hardware": hardware,
        "excel": excel,
        "wps": wps,
        "com": com,
        "libreoffice": libreoffice,
        "python": python,
        "pywin32": "不使用（Office COM 由本机接口调用）",
        "webview2": "不适用（当前界面由系统默认浏览器承载）",
        "openpyxl": openpyxl
    })
}

fn executable_in_path(names: &[&str]) -> Option<PathBuf> {
    let path = std::env::var_os("PATH")?;
    for directory in std::env::split_paths(&path) {
        for name in names {
            let candidate = directory.join(name);
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

fn host_os_description() -> String {
    #[cfg(target_os = "linux")]
    {
        if let Ok(contents) = fs::read_to_string("/etc/os-release") {
            if let Some(value) = contents.lines().find_map(|line| {
                line.strip_prefix("PRETTY_NAME=")
                    .map(|value| value.trim_matches('"').to_owned())
            }) {
                return value;
            }
        }
        return "Linux（无法读取 /etc/os-release）".to_owned();
    }
    #[cfg(windows)]
    {
        if let Ok(output) = std::process::Command::new("cmd")
            .args(["/C", "ver"])
            .output()
        {
            let value = String::from_utf8_lossy(&output.stdout).trim().to_owned();
            if let Some(version) = value
                .split(|character: char| !character.is_ascii_digit() && character != '.')
                .find(|part| {
                    part.matches('.').count() >= 2
                        && part
                            .chars()
                            .all(|character| character.is_ascii_digit() || character == '.')
                })
            {
                return format!("Microsoft Windows {version}");
            }
        }
        return "Microsoft Windows".to_owned();
    }
    #[cfg(target_os = "macos")]
    {
        if let Ok(output) = std::process::Command::new("sw_vers")
            .arg("-productVersion")
            .output()
        {
            let value = String::from_utf8_lossy(&output.stdout).trim().to_owned();
            if !value.is_empty() {
                return format!("macOS {value}");
            }
        }
        "macOS（系统版本读取失败）".to_owned()
    }
}

fn host_locale_description() -> String {
    for key in ["LC_ALL", "LC_CTYPE", "LANG"] {
        if let Ok(value) = std::env::var(key) {
            if !value.trim().is_empty() {
                return format!("{key}={}", value.trim());
            }
        }
    }
    if cfg!(windows) {
        "Windows 系统区域；环境变量未设置".to_owned()
    } else {
        "环境变量未设置".to_owned()
    }
}

fn host_glibc_description() -> String {
    #[cfg(target_os = "linux")]
    {
        if let Ok(output) = std::process::Command::new("getconf")
            .arg("GNU_LIBC_VERSION")
            .output()
        {
            let value = String::from_utf8_lossy(&output.stdout).trim().to_owned();
            if output.status.success() && !value.is_empty() {
                return value;
            }
        }
        return "未检测到 glibc 版本（可能使用 musl 或缺少 getconf）".to_owned();
    }
    #[cfg(not(target_os = "linux"))]
    {
        "不适用（非 Linux）".to_owned()
    }
}

#[cfg(target_os = "linux")]
fn linux_memory_gib() -> Option<f64> {
    let contents = fs::read_to_string("/proc/meminfo").ok()?;
    let kib = contents.lines().find_map(|line| {
        let value = line.strip_prefix("MemTotal:")?.split_whitespace().next()?;
        value.parse::<u64>().ok()
    })?;
    Some(kib as f64 / 1024.0 / 1024.0)
}

fn scan_depth_from_value(value: &Value) -> ScanDepth {
    match value.as_i64().unwrap_or(-1) {
        0 => ScanDepth::RootOnly,
        depth if depth > 0 => ScanDepth::MaxDirectoryDepth(depth as usize),
        _ => ScanDepth::Full,
    }
}

fn current_scan_depth(state: &AppState) -> ScanDepth {
    scan_depth_from_value(&state.lock().value["recursiveDepth"])
}

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
        #[cfg(test)]
        tests::isolate_settings_state_dir();
        let config_path = config_path.to_string_lossy().into_owned();
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
            .and_then(|parent| {
                ["1.逐笔统计系统_配置.xlsx", "逐笔统计系统_配置.xlsx"]
                    .into_iter()
                    .map(|name| parent.join(name))
                    .find(|path| path.is_file())
            })
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let config_report = check_s2_config(&config_path);
        let openpyxl_status = crate::app_settings::openpyxl_selection_status();
        let system_environment_info = system_environment_info(&openpyxl_status);
        let mut state = json!({
            "appInfo": {"title": app_title()},
            "systemEnvironmentInfo": system_environment_info,
            "busy": false,
            "currentWork": Value::Null,
            "cancelRequested": false,
            "status": if distribution_profile() == "s1s2" {
                "审核工具已启动；S1、S2 入口已接入"
            } else {
                "审核工具已启动；S1、S2、S3 入口已接入"
            },
            "log": [],
            "logVersion": 0,
            "lastRunSummary": Value::Null,
            "filePickerMode": "自动",
            "filePickerModeOptions": ["自动", "系统原生", "浏览器内置"],
            "uiTheme": "日间",
            "uiAccent": "默认蓝",
            "confirmBeforeRun": true,
            "showRunDetailLogs": false,
            "showPerformanceDiagnostics": false,
            "exportRunLogs": false,
            "s1ConditionalFormatSource": "audit_copy",
            "conditionalFormatIncludeBlankCells": true,
            "centralTolerance": 100.0,
            "centralTrimIndicatorNames": false,
            "centralAutoOpenResult": false,
            "recursiveDepth": -1,
            "developerAccess": {},
            "developerAccessMessage": "",
            "distributionUnlocked": false,
            "distributionProfile": distribution_profile(),
            "configUpdateSupported": crate::config_update::supported(),
            "configIssues": {},
            "sourceFiles": [],
            "selectedFiles": [],
            "s1F1Preflight": {},
            "s1F1OutputPath": "",
            "historyConfig": history_config,
            "extraFiles": [],
            "mixedTemplates": [],
            "explanationFiles": [],
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
            "centralOutput": "",
            "centralCompareFile": "",
            "centralExplainFile": "",
            "centralCommonConfig": central_common_config,
            "centralFormsConfig": central_forms_config,
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
            "engines": [{"value": "自动", "label": "审核工具功能"}]
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
        // 测试构建下跳过：并行测试会把各自选择的路径持久化到同一个测试
        // 设置文件，恢复会把别的测试的临时路径回灌进本测试状态。
        #[cfg(not(test))]
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
        let mut ui = self.lock();
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
        println!("审核工具 web UI: {url}");
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

/// 发行版档位：默认 S1+S2+S3 全量。
///
/// 用「可执行文件同级的 `edition-s1s2` 标记文件」切换为 S1+S2 版，
/// 避免为两个版本分别构建（同一 exe，打包时放/不放标记文件即可）。
/// 若标记文件内容为 `s1s2s3` 则仍按全量处理，便于日后扩展。
fn distribution_profile() -> &'static str {
    let Ok(executable) = std::env::current_exe() else {
        return "s1s2s3";
    };
    let Some(directory) = executable.parent() else {
        return "s1s2s3";
    };
    let marker = directory.join("edition-s1s2");
    if !marker.is_file() {
        return "s1s2s3";
    }
    match std::fs::read_to_string(&marker) {
        Ok(text) if text.trim().eq_ignore_ascii_case("s1s2s3") => "s1s2s3",
        Ok(_) => "s1s2",
        Err(_) => "s1s2",
    }
}

fn render_index() -> String {
    ORIGINAL_HTML
        .replace("__AUDIT_APP_TITLE__", app_title())
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
        "check_period_audit_config" => check_period_audit_config(state),
        "choose_period_compare" => choose_period_compare(state, args),
        "choose_central" => choose_central(state, args),
        "choose_folder" => choose_folder(state, args),
        "choose_file" => choose_file(state, args),
        "choose_history_config" => choose_history_config(state, args),
        "clear_path" => clear_path(state, args),
        "open_path" => open_path(state, args),
        "recognize" => recognize_s1_sources(state),
        "set_selected_files" => set_selected_files(state, args),
        "get_faq" => Ok(crate::faq::items(distribution_profile())),
        "get_config_update" => Ok(crate::config_update::inspect()),
        "apply_config_update" => {
            let names = args
                .first()
                .and_then(Value::as_array)
                .map(|items| {
                    items
                        .iter()
                        .filter_map(Value::as_str)
                        .map(str::to_owned)
                        .collect::<Vec<_>>()
                })
                .unwrap_or_default();
            Ok(crate::config_update::apply_updates(&names))
        }
        "append_extra_files" => append_extra_files(state, args),
        "remove_extra_file" => remove_extra_file(state, args),
        "inspect_template_formulas" => inspect_template_formulas(state),
        "start_dag_flow" => start_dag_flow(state, args),
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
        "cancel_execution" => cancel_execution(state),
        "update" => update_state(state, args),
        _ => Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("当前审核工具尚未实现接口：{method}"),
        )),
    }
}

fn open_path(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let requested = arg_string(args, 0)
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "打开路径不能为空"))?;
    let path = PathBuf::from(unquote_path(requested));
    if !path.exists() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("要打开的文件或目录不存在：{}", path.display()),
        ));
    }
    open_path_with_default_app(&path).map_err(|error| {
        ApiError::internal(format!(
            "无法使用系统默认程序打开 {}：{error}",
            path.display()
        ))
    })?;
    push_log(
        &mut state.lock().value,
        format!("已在系统默认程序中打开：{}", path.display()),
    );
    Ok(json!({"opened": true, "path": path.to_string_lossy()}))
}

#[cfg(windows)]
fn open_path_with_default_app(path: &Path) -> std::io::Result<()> {
    use std::{ffi::c_void, os::windows::ffi::OsStrExt, ptr};

    #[link(name = "shell32")]
    extern "system" {
        fn ShellExecuteW(
            window: *mut c_void,
            operation: *const u16,
            file: *const u16,
            parameters: *const u16,
            directory: *const u16,
            show_command: i32,
        ) -> *mut c_void;
    }

    let operation = "open".encode_utf16().chain(Some(0)).collect::<Vec<_>>();
    let file = path
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect::<Vec<_>>();
    let result = unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            operation.as_ptr(),
            file.as_ptr(),
            ptr::null(),
            ptr::null(),
            1,
        )
    };
    if result as isize <= 32 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::Other,
            format!("ShellExecuteW 返回错误码 {}", result as usize),
        ));
    }
    Ok(())
}

#[cfg(target_os = "macos")]
fn open_path_with_default_app(path: &Path) -> std::io::Result<()> {
    Command::new("open").arg(path).spawn().map(|_| ())
}

#[cfg(target_os = "linux")]
fn open_path_with_default_app(path: &Path) -> std::io::Result<()> {
    Command::new("xdg-open").arg(path).spawn().map(|_| ())
}

#[cfg(not(any(windows, target_os = "macos", target_os = "linux")))]
fn open_path_with_default_app(_path: &Path) -> std::io::Result<()> {
    Err(std::io::Error::new(
        std::io::ErrorKind::Unsupported,
        "当前平台不支持通过系统默认程序打开路径",
    ))
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
        title: if title.is_empty() {
            "选择位置".to_owned()
        } else {
            title
        },
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
            format!("当前审核工具暂不支持选择该目录：{field}"),
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
    let (input_dir, template_dir, manual_template, busy, scan_depth) = {
        let ui = state.lock();
        (
            ui.value["input"].as_str().unwrap_or_default().to_owned(),
            ui.value["templateDir"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
            ui.value["templateManual"] == true,
            ui.value["busy"] == true,
            scan_depth_from_value(&ui.value["recursiveDepth"]),
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
    let files = s1_region_summary_source_workbooks_with_depth(&input_dir, scan_depth)
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
    // 手动追加的源目录外文件并入清单并同样参与模板匹配（旧 Audit 口径：
    // 直接位于输入目录下的追加文件由目录扫描覆盖，不再重复并入）。
    let extra_files = crate::app_settings::AppSettings::load().extra_files();
    let mut all_paths = files.clone();
    for extra in &extra_files {
        let extra_path = PathBuf::from(extra);
        if !extra_path.is_file() || extra_path.parent() == Some(&input_dir) {
            continue;
        }
        let known = files
            .iter()
            .chain(all_paths.iter())
            .any(|path| web_path_identity(path) == web_path_identity(&extra_path));
        if !known {
            all_paths.push(extra_path);
        }
    }
    for path in &all_paths {
        let is_extra = !files
            .iter()
            .any(|scanned| web_path_identity(scanned) == web_path_identity(path));
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
            "templateReason": suggestion.details,
            "extra": is_extra
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
    ui.value["selectedFiles"] = json!(all_paths);
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

fn refresh_s1_sources_for_scan_depth(state: &AppState) {
    let (input, previous_sources, previous_selected) = {
        let ui = state.lock();
        let previous_sources = ui.value["sourceFiles"]
            .as_array()
            .into_iter()
            .flatten()
            .filter_map(|file| file["path"].as_str().map(str::to_owned))
            .collect::<std::collections::HashSet<_>>();
        let previous_selected = ui.value["selectedFiles"]
            .as_array()
            .into_iter()
            .flatten()
            .filter_map(Value::as_str)
            .map(str::to_owned)
            .collect::<std::collections::HashSet<_>>();
        (
            ui.value["input"].as_str().unwrap_or_default().to_owned(),
            previous_sources,
            previous_selected,
        )
    };

    if input.trim().is_empty() {
        let mut ui = state.lock();
        ui.value["sourceFiles"] = json!([]);
        ui.value["selectedFiles"] = json!([]);
        ui.value["mixedTemplates"] = json!([]);
        return;
    }

    if let Err(error) = recognize_s1_sources(state) {
        let mut ui = state.lock();
        ui.value["sourceFiles"] = json!([]);
        ui.value["selectedFiles"] = json!([]);
        ui.value["mixedTemplates"] = json!([]);
        ui.value["status"] = json!("递归深度已保存，但 S1 源文件列表刷新失败");
        push_error_log(
            &mut ui.value,
            format!("递归深度调整后刷新 S1 源文件列表失败：{}", error.message),
        );
        return;
    }

    let mut ui = state.lock();
    let selected = ui.value["sourceFiles"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|file| file["path"].as_str())
        .filter(|path| !previous_sources.contains(*path) || previous_selected.contains(*path))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    ui.value["selectedFiles"] = json!(selected);
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

/// 手动追加/移除待处理文件后统一持久化 `extraFiles`（其余偏好原样保留）。
fn persist_extra_files(state: &AppState, files: &[String]) -> std::result::Result<(), ApiError> {
    let settings = crate::app_settings::AppSettings::load();
    let mut preferences = settings.preferences.clone();
    preferences.insert("extraFiles".to_owned(), json!(files));
    crate::app_settings::AppSettings::store_preferences(
        settings.calc_engine.as_str(),
        settings.edit_engine.as_str(),
        preferences,
    )
    .map_err(|error| ApiError::internal(format!("保存用户设置失败：{error:#}")))?;
    let mut ui = state.lock();
    ui.value["extraFiles"] = json!(files);
    Ok(())
}

/// 待处理清单：批量追加源数据目录外的工作簿（跨会话保留），随后重建
/// 源文件清单并入清单。与旧 Audit 口径一致：只接受存在的文件并去重。
fn append_extra_files(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(paths) = args.first().and_then(Value::as_array) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "append_extra_files 需要路径数组",
        ));
    };
    let requested = paths
        .iter()
        .map(|value| {
            value
                .as_str()
                .map(PathBuf::from)
                .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "追加路径必须是字符串"))
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能追加待处理文件",
            ));
        }
    }
    let existing: std::collections::HashSet<String> = {
        let ui = state.lock();
        ui.value["extraFiles"]
            .as_array()
            .into_iter()
            .flatten()
            .filter_map(|value| value.as_str().map(str::to_owned))
            .collect()
    };
    let mut merged = existing;
    let mut added = 0usize;
    for path in requested {
        let Ok(resolved) = fs::canonicalize(&path) else {
            continue;
        };
        if !resolved.is_file() {
            continue;
        }
        let text = resolved.to_string_lossy().into_owned();
        if merged.insert(text) {
            added += 1;
        }
    }
    let mut files: Vec<String> = merged.into_iter().collect();
    files.sort();
    persist_extra_files(state, &files)?;
    {
        let mut ui = state.lock();
        push_log(&mut ui.value, format!("已追加 {added} 个待处理文件"));
    }
    refresh_s1_source_list(state);
    Ok(state.get_state())
}

/// 从追加清单移除一个手动追加的文件（不影响源数据目录扫描结果）。
fn remove_extra_file(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(path) = args.first().and_then(Value::as_str) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "remove_extra_file 需要一个路径",
        ));
    };
    {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能修改待处理文件",
            ));
        }
    }
    let previous: Vec<String> = {
        let ui = state.lock();
        ui.value["extraFiles"]
            .as_array()
            .into_iter()
            .flatten()
            .filter_map(|value| value.as_str().map(str::to_owned))
            .collect()
    };
    let remaining: Vec<String> = previous
        .iter()
        .filter(|item| item.as_str() != path)
        .cloned()
        .collect();
    if remaining.len() != previous.len() {
        persist_extra_files(state, &remaining)?;
        {
            let mut ui = state.lock();
            let file_name = Path::new(path)
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| path.to_owned());
            push_log(&mut ui.value, format!("已移除追加文件：{file_name}"));
        }
        refresh_s1_source_list(state);
    }
    Ok(state.get_state())
}

/// 在输入目录与模板目录有效时重建源文件清单；目录未就绪时静默跳过
/// （追加的文件仍保留在设置里，等目录可用后随下次识别一起并入）。
fn refresh_s1_source_list(state: &AppState) {
    let input_dir = {
        let ui = state.lock();
        ui.value["input"].as_str().unwrap_or_default().to_owned()
    };
    if !PathBuf::from(input_dir.trim()).is_dir() {
        return;
    }
    let _ = recognize_s1_sources(state);
}

fn choose_file(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_file 缺少路径类型"))?;
    if !matches!(field, "template" | "external") {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("当前审核工具暂不支持选择该文件：{field}"),
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
    conditional_format_source: String,
    request_preparation_us: u128,
    source_directory_scan_us: u128,
}

struct S1F1TemplatePreflight {
    structure: S1StructureTemplate,
    rule_plan: S1TemplateRulePlan,
    conditional_mappings: Vec<FeatureMapping>,
    scan_areas: Vec<String>,
    scan_area_source: String,
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
    let source_files =
        s1_region_summary_source_workbooks_with_depth(&input_dir, current_scan_depth(state))
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
    let conditional_format_source = settings.s1_conditional_format_source().to_owned();
    {
        let mut ui = state.lock();
        push_log(
            &mut ui.value,
            format!(
                "S1-F01 高级设置：计算引擎={}；修改引擎={}；条件格式提取来源={}（设置文件 {}）",
                settings.calc_engine.label(),
                settings.edit_engine.label(),
                if conditional_format_source == "audit_copy" {
                    "复制校验区域并重算后的审核副本"
                } else {
                    "原始源文件"
                },
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
        conditional_format_source,
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
        clear_run_summary(&mut ui.value);
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
        .spawn(move || run_s1_f1_worker(worker_state, request, request_preparation_started))
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["s1F1OutputPath"] = json!("");
        ui.value["status"] = json!("S1-F01 启动失败");
        push_error_log(&mut ui.value, format!("启动 S1-F01 失败：{error}"));
        return Err(ApiError::internal(format!("启动 S1-F01 失败：{error}")));
    }
    Ok(json!({"accepted": true}))
}

fn s1_f1_end_to_end_total_us(task_started: Instant) -> u128 {
    task_started.elapsed().as_micros()
}

fn run_s1_f1_worker(state: AppState, request: S1F1PreflightRequest, task_started: Instant) {
    let work_slot: CurrentWorkSlot = std::sync::Arc::new(std::sync::Mutex::new(json!({
        "feature": "S1-F01",
        "total": request.source_files.len(),
    })));
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
    for (source_index, source) in request.source_files.iter().enumerate() {
        let template_match_started = Instant::now();
        let source_name = source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| source.display().to_string());
        let source_label = source
            .strip_prefix(&request.source_directory)
            .unwrap_or(source)
            .to_string_lossy()
            .into_owned();
        {
            let mut ui = state.lock();
            push_log(
                &mut ui.value,
                format!(
                    "S1-F01 [{}/{}] 开始模板匹配与表结构检查：{}",
                    source_index + 1,
                    request.source_files.len(),
                    source_label
                ),
            );
        }
        current_work_update(
            &work_slot,
            json!({
                "feature": "S1-F01",
                "workbook": source_name.clone(),
                "sheet": Value::Null,
                "stage": "匹配模板并检查表结构",
                "index": source_index + 1,
                "total": request.source_files.len(),
            }),
        );
        current_work_publish(&state, &work_slot);
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
                        let report = json!({
                            "sourceFile": source_name,
                            "status": "skipped_template_unavailable",
                            "template": template_path,
                            "reason": format!("无法读取匹配模板：{error}")
                        });
                        let mut ui = state.lock();
                        push_warn_log(
                            &mut ui.value,
                            format!(
                                "S1-F01 文件={} 阶段=模板预检 已跳过：{}",
                                source_label,
                                report["reason"].as_str().unwrap_or("未提供原因")
                            ),
                        );
                        drop(ui);
                        files.push(report);
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
                            "scanAreaSource": preflight.scan_area_source,
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
                                let scan_area_source = preflight.scan_area_source.clone();
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
                                    "scanAreaSource": scan_area_source,
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
        let status = file_report["status"].as_str().unwrap_or("未知状态");
        let reason = file_report["reason"].as_str().unwrap_or("");
        let message = format!(
            "S1-F01 [{}/{}] 文件={} 阶段=模板匹配与表结构检查 结果={}{}",
            source_index + 1,
            request.source_files.len(),
            source_label,
            status,
            if reason.is_empty() {
                String::new()
            } else {
                format!("；原因：{reason}")
            }
        );
        let mut ui = state.lock();
        if status.starts_with("skipped") {
            push_warn_log(&mut ui.value, message);
        } else {
            push_log(&mut ui.value, message);
        }
        drop(ui);
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
        current_work_finish(&state, &work_slot);
        let end_to_end_total_us = s1_f1_end_to_end_total_us(task_started);
        let end_to_end_total_seconds = end_to_end_total_us as f64 / 1_000_000.0;
        let performance_timings = json!({
            "request": request_timings,
            "preflight": preflight_timings,
            "endToEndTotalUs": end_to_end_total_us
        });
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
            "performanceTimings": performance_timings
        });
        ui.value["status"] = json!(format!(
            "S1-F01 预检结束：结构通过 0 个、跳过 {} 个；未生成审核结果；总用时 {:.1} 秒",
            skipped, end_to_end_total_seconds
        ));
        for file in files {
            let name = file["sourceFile"].as_str().unwrap_or("未知文件");
            let reason = file["reason"].as_str().unwrap_or("未提供原因");
            push_warn_log(
                &mut ui.value,
                format!("S1-F01 文件={name} 阶段=模板/表结构预检 已跳过：{reason}"),
            );
        }
        push_log(
            &mut ui.value,
            format!(
                "S1-F01 未生成审核结果；预期路径：{}",
                request.planned_result_path.display()
            ),
        );
        push_log(
            &mut ui.value,
            format!("S1-F01 总用时：{end_to_end_total_seconds:.1} 秒"),
        );
        publish_run_summary(
            &mut ui.value,
            "S1-F01",
            format!(
                "工作簿总数 {}；结构通过 0；跳过 {}；失败 0；未生成审核副本和结果；输出目录 {}；总耗时 {}。",
                request.source_files.len(),
                skipped,
                request.output_directory.display(),
                format_elapsed(task_started.elapsed())
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

    match execute_s1_f1_workflow(&request, &accepted, Some(work_slot.clone()), Some(&state)) {
        Ok(output) => {
            let end_to_end_total_us = s1_f1_end_to_end_total_us(task_started);
            let end_to_end_total_seconds = end_to_end_total_us as f64 / 1_000_000.0;
            // 最终统计随 currentWork 发布；随后清空"正在处理"对象，
            // 避免任务结束后仍显示旧文件。
            let reviewed = output
                .file_results
                .iter()
                .filter(|file| {
                    matches!(file["status"].as_str(), Some("COMPLETED") | Some("PARTIAL"))
                })
                .count();
            current_work_update(
                &work_slot,
                json!({
                    "stage": "完成",
                    "done": reviewed,
                    "failed": output.failed_files,
                    "partial": output.partial_files,
                    "waitingManualRecalc": output.waiting_manual_recalc,
                }),
            );
            current_work_finish(&state, &work_slot);
            let mut performance_timings = json!({
                "request": request_timings,
                "preflight": preflight_timings,
                "workflow": output.performance_timings
            });
            performance_timings["endToEndTotalUs"] = json!(end_to_end_total_us);
            let recalculation_engine_labels = output
                .recalculation_engines
                .iter()
                .map(|engine| match engine.as_str() {
                    "Excel.Application" => "Microsoft Excel COM",
                    "ket.Application" | "KET.Application" => "WPS 表格 COM",
                    other => other,
                })
                .collect::<Vec<_>>()
                .join("、");
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
                "manualRecalcInstructions": "自动公式重算不可用或失败，已保留待计算的审核副本。请使用 Microsoft Excel 打开审核副本：1. 按 Ctrl+Alt+F9 强制重新计算；2. 等待 Excel 完成计算；3. 按 Ctrl+S 保存；4. 直接在 Excel 中查看「校验区域」工作表的计算结果。如果使用 WPS 或 LibreOffice，请采用对应软件支持的重新计算方式并保存。请勿修改原始报表文件。",
                "externalData": request.external_data,
                "recalculationSidecarFound": request.sidecar_path.is_some(),
                "recalculationSidecar": request.sidecar_path,
                "performanceTimings": performance_timings
            });
            for file in &files {
                let name = file["sourceFile"].as_str().unwrap_or("未知文件");
                let status = file["status"].as_str().unwrap_or("未知状态");
                let reason = file["reason"].as_str().unwrap_or("");
                let message = format!(
                    "S1-F01 {name}：{status}{}",
                    if reason.is_empty() {
                        String::new()
                    } else {
                        format!("；阶段/原因：{reason}")
                    }
                );
                let normalized_status = status.to_ascii_lowercase();
                if normalized_status.contains("fail")
                    || normalized_status.contains("waiting_manual_recalc")
                {
                    push_error_log(&mut ui.value, message);
                } else if normalized_status.starts_with("skipped") {
                    push_warn_log(&mut ui.value, message);
                } else {
                    push_log(&mut ui.value, message);
                }
            }
            ui.value["status"] = json!(if output.waiting_manual_recalc > 0 {
                format!(
                    "S1-F01 部分完成：{} 个待人工重算、{} 个部分审核、{} 个失败；总用时 {:.1} 秒；已生成结果 {}",
                    output.waiting_manual_recalc,
                    output.partial_files,
                    output.failed_files,
                    end_to_end_total_seconds,
                    output
                        .result_path
                        .as_ref()
                        .map(|path| path.display().to_string())
                        .unwrap_or_else(|| "无可发布结果".to_owned())
                )
            } else if output.result_path.is_some() {
                format!(
                    "S1-F01 审核完成：公式异常 {} 条、条件格式异常 {} 条；总耗时 {:.1} 秒；结果已保存",
                    output.formula_issue_count,
                    output.conditional_issue_count,
                    end_to_end_total_seconds
                )
            } else {
                format!(
                    "S1-F01 未生成核查表：待恢复 {} 个、失败 {} 个；总用时 {:.1} 秒；审核副本与状态已保留",
                    output.waiting_manual_recalc, output.failed_files, end_to_end_total_seconds
                )
            });
            let succeeded = output
                .file_results
                .iter()
                .filter(|file| file["status"] == "COMPLETED")
                .count();
            push_log(
                &mut ui.value,
                format!("S1-F01 总用时：{end_to_end_total_seconds:.1} 秒"),
            );
            push_log(
                &mut ui.value,
                format!(
                    "S1-F01 实际重算引擎：{}",
                    if recalculation_engine_labels.is_empty() {
                        "无（没有工作簿完成真实重算）"
                    } else {
                        &recalculation_engine_labels
                    }
                ),
            );
            // 技术路径明细：只在「设置 → 显示详细运行日志」开启时输出，
            // 让人一眼看出本次实际走的是哪条链路，而不是靠推断。
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F01 技术路径：{}",
                    s1_f1_technical_path(&request, &output)
                ),
            );
            if let Some(result_path) = output.result_path.as_ref() {
                push_log(
                    &mut ui.value,
                    format!("S1-F01 结果文件：{}", result_path.display()),
                );
            } else if output.waiting_manual_recalc > 0 {
                push_log(&mut ui.value, "自动公式重算不可用或失败，已保留待计算的审核副本。请用 Excel 打开审核副本：按 Ctrl+Alt+F9 强制重算，等待计算完成后按 Ctrl+S 保存，即可在 Excel 中查看「校验区域」结果。请勿修改原始报表文件。".to_owned());
            }
            push_log(
                &mut ui.value,
                format!("S1-F01 审核副本目录：{}", request.run_directory.display()),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F01 正式审核结束：数据期={}；审核副本={}；公式异常={}；条件格式异常={}；读取历史行={}；本期合计={}；真实重算引擎={}；结果工作表={}；输出={}；待人工重算={}；部分完成={}；失败={}；总用时={:.1} 秒。原始源文件保持只读。",
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
                    output.failed_files,
                    end_to_end_total_seconds
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S1-F01",
                format!(
                    "工作簿总数 {}；成功 {}；部分完成 {}；待人工重算 {}；失败 {}；跳过 {}；生成审核副本 {} 个；结果 {}；输出目录 {}；总耗时 {}。",
                    request.source_files.len(),
                    succeeded,
                    output.partial_files,
                    output.waiting_manual_recalc,
                    output.failed_files,
                    skipped,
                    output.audit_copies.len(),
                    output.result_path.as_ref().map(|path| path.display().to_string()).unwrap_or_else(|| "未生成".to_owned()),
                    request.output_directory.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
        Err(error) => {
            let end_to_end_total_us = s1_f1_end_to_end_total_us(task_started);
            let end_to_end_total_seconds = end_to_end_total_us as f64 / 1_000_000.0;
            current_work_update(
                &work_slot,
                json!({"stage": "失败", "error": format!("{error:.500}")}),
            );
            current_work_finish(&state, &work_slot);
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
                    "preflight": preflight_timings,
                    "endToEndTotalUs": end_to_end_total_us
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
            ui.value["status"] = json!(format!(
                "S1-F01 正式审核失败；未生成审核结果；总用时 {:.1} 秒",
                end_to_end_total_seconds
            ));
            push_log(
                &mut ui.value,
                format!(
                    "S1-F01 未生成审核结果；预期路径：{}",
                    request.planned_result_path.display()
                ),
            );
            push_log(
                &mut ui.value,
                format!("S1-F01 总用时：{end_to_end_total_seconds:.1} 秒"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F01 正式审核失败；本次新建审核目录已清理；原始源文件保持只读；未发布部分结果；总用时={end_to_end_total_seconds:.1} 秒。阻断原因：{reason}"
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S1-F01",
                format!(
                    "工作簿总数 {}；成功 0；部分完成 0；失败 {}；跳过 {}；审核结果未生成；输出目录 {}；总耗时 {}；原因：{}。",
                    request.source_files.len(),
                    structure_passed,
                    skipped,
                    request.output_directory.display(),
                    format_elapsed(task_started.elapsed()),
                    reason
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

fn s1_conditional_format_workbook<'a>(
    source_file: &'a Path,
    audit_copy: &'a Path,
    conditional_format_source: &str,
) -> &'a Path {
    if conditional_format_source == "audit_copy" {
        audit_copy
    } else {
        source_file
    }
}

fn set_conditional_issues_audit_copy(issues: &mut [Issue], audit_copy: &Path) {
    for issue in issues {
        issue.audit_file = audit_copy.display().to_string();
    }
}

fn extract_s1_audit_copy(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    conditional_format_source: &str,
    metadata: &S1FormulaMetadata,
) -> Result<S1ExtractedAuditCopy> {
    extract_s1_audit_copy_with_include_blank(
        item,
        source_file,
        audit_copy,
        conditional_format_source,
        metadata,
        None,
        crate::app_settings::AppSettings::load().conditional_format_include_blank_cells(),
    )
}

fn extract_s1_audit_copy_with_display_observation(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    conditional_format_source: &str,
    metadata: &S1FormulaMetadata,
    display_observation: Option<
        std::result::Result<crate::conditional_format::ConditionalDisplayScan, String>,
    >,
) -> Result<S1ExtractedAuditCopy> {
    extract_s1_audit_copy_with_include_blank(
        item,
        source_file,
        audit_copy,
        conditional_format_source,
        metadata,
        display_observation,
        crate::app_settings::AppSettings::load().conditional_format_include_blank_cells(),
    )
}

/// 空白格开关显式传入的内部入口；S1-F01 执行链从用户设置读取。
fn extract_s1_audit_copy_with_include_blank(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    conditional_format_source: &str,
    metadata: &S1FormulaMetadata,
    display_observation: Option<
        std::result::Result<crate::conditional_format::ConditionalDisplayScan, String>,
    >,
    include_blank_cells: bool,
) -> Result<S1ExtractedAuditCopy> {
    extract_s1_audit_copy_with_display_observation_and_scan_areas(
        item,
        source_file,
        audit_copy,
        conditional_format_source,
        metadata,
        &item.preflight.scan_areas,
        display_observation,
        include_blank_cells,
    )
}

fn extract_s1_audit_copy_with_display_observation_and_scan_areas(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    conditional_format_source: &str,
    metadata: &S1FormulaMetadata,
    scan_areas: &[String],
    display_observation: Option<
        std::result::Result<crate::conditional_format::ConditionalDisplayScan, String>,
    >,
    include_blank_cells: bool,
) -> Result<S1ExtractedAuditCopy> {
    extract_s1_audit_copy_with_conditional_extractor_and_display(
        item,
        source_file,
        audit_copy,
        conditional_format_source,
        metadata,
        scan_areas,
        display_observation,
        include_blank_cells,
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
    conditional_format_source: &str,
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
    // 测试与定向探针入口保持既有口径：空白格命中纳入结果。
    extract_s1_audit_copy_with_conditional_extractor_and_display(
        item,
        source_file,
        audit_copy,
        conditional_format_source,
        metadata,
        &item.preflight.scan_areas,
        None,
        true,
        conditional_extractor,
    )
}

fn extract_s1_audit_copy_with_conditional_extractor_and_display<F>(
    item: &S1F1AcceptedSource,
    source_file: &Path,
    audit_copy: &Path,
    conditional_format_source: &str,
    metadata: &S1FormulaMetadata,
    scan_areas: &[String],
    display_observation: Option<
        std::result::Result<crate::conditional_format::ConditionalDisplayScan, String>,
    >,
    include_blank_cells: bool,
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
    // Formula findings above intentionally come from the recalculated audit
    // copy. Conditional-format values use the user's selected workbook source;
    // the default remains the untouched source to match legacy Python.
    let conditional_workbook =
        s1_conditional_format_workbook(source_file, audit_copy, conditional_format_source);
    let conditional_metadata = S1ConditionalMetadata {
        period: metadata.period.clone(),
        batch_id: metadata.batch_id.clone(),
        audit_time: metadata.audit_time.clone(),
        org_code: metadata.org_code.clone(),
        org_name: metadata.org_name.clone(),
    };
    let conditional_result = match display_observation {
        Some(Ok(scan)) => extract_conditional_issues_with_display_scan(
            conditional_workbook,
            source_file,
            scan_areas,
            &item.preflight.structure_ranges,
            &conditional_metadata,
            &scan,
        ),
        Some(Err(reason)) => extract_conditional_issues_with_display_failure(
            conditional_workbook,
            source_file,
            scan_areas,
            &item.preflight.structure_ranges,
            &conditional_metadata,
            reason,
        ),
        None => conditional_extractor(
            conditional_workbook,
            source_file,
            scan_areas,
            &item.preflight.structure_ranges,
            &conditional_metadata,
        ),
    };
    let conditional_format_extraction_elapsed_us =
        conditional_extraction_started.elapsed().as_micros();
    // 空白格开关（默认开启，与 COM DisplayFormat 实际渲染口径一致）：
    // 关闭时排除目标单元格为空白（值 None 或空文本）的命中，其余口径不变。
    let mut conditional_result = conditional_result;
    if !include_blank_cells {
        if let Ok(result) = conditional_result.as_mut() {
            let before = result.issues.len();
            result.issues.retain(|issue| {
                !(issue.target_value.is_null()
                    || issue
                        .target_value
                        .as_str()
                        .is_some_and(|text| text.is_empty()))
            });
            result
                .evaluation
                .diagnostics
                .push(crate::conditional_format::ConditionalDiagnostic {
                    sheet: String::new(),
                    cell: String::new(),
                    applies_to: String::new(),
                    rule_type: "blank_cells".to_owned(),
                    formula: String::new(),
                    priority: 0,
                    status: crate::conditional_format::EvaluationStatus::Unsupported,
                    reason: format!(
                        "已按设置排除空白单元格命中 {before} → {}",
                        result.issues.len()
                    ),
                });
        }
    }

    let (
        mut conditional_issues,
        mut diagnostics,
        conditional_partial,
        conditional_diagnostic_count,
        unresolved_indicator_count,
        conditional_error,
    ) = match conditional_result {
        Ok(conditional_result) => {
            let mut diagnostics = conditional_result
                .evaluation
                .diagnostics
                .iter()
                .map(|diagnostic| json!({"kind":"conditional_format", "diagnostic":diagnostic}))
                .collect::<Vec<_>>();
            if conditional_result.evaluation.status == "evaluated_com_display" {
                diagnostics.push(json!({
                    "kind": "conditional_display_color_scan",
                    "engine": conditional_result.evaluation.display_engine,
                    "engineProcessId": conditional_result.evaluation.display_engine_process_id,
                    "scannedCells": conditional_result.evaluation.scanned_cells,
                    "changedFills": conditional_result.evaluation.triggered.iter().map(|trigger| json!({
                        "sheet": trigger.sheet,
                        "cell": trigger.cell,
                        "baseFillColor": trigger.base_fill_color,
                        "displayFillColor": trigger.display_fill_color,
                    })).collect::<Vec<_>>(),
                }));
            }
            (
                conditional_result.issues,
                diagnostics,
                conditional_result.status != "completed",
                conditional_result.evaluation.diagnostics.len(),
                conditional_result.context.unresolved_indicator_count,
                None,
            )
        }
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
    // CF evaluation may read the untouched source workbook, but result links
    // must always open the per-source audit copy alongside formula findings.
    set_conditional_issues_audit_copy(&mut conditional_issues, audit_copy);
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

/// S1-F01 失败保留副本的恢复结果路径（仍在 S1-F01 的失败保留流程中使用）。
fn s1_recovery_result_path(manifest: &S1RecoveryManifest, input_sha256: &str) -> PathBuf {
    let run_hash = format!("{:x}", Sha256::digest(manifest.run_id.as_bytes()));
    let run_key = &run_hash[..32];
    let input_key = input_sha256.get(..32).unwrap_or(input_sha256);
    manifest
        .run_directory
        .join(format!("汇总核查表校验_恢复_{run_key}_{input_key}.xlsx"))
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
    let input_bytes = serde_json::to_vec(&prepared.issues).context("序列化 S1-F01 结果输入")?;
    let input_sha256 = format!("{:x}", Sha256::digest(input_bytes));
    let result_path = s1_recovery_result_path(manifest, &input_sha256);
    let result_name = result_path
        .file_name()
        .and_then(|name| name.to_str())
        .context("生成 S1-F01 结果文件名")?;
    let candidate_path = result_path.with_file_name(format!(".{result_name}.candidate.xlsx"));

    if manifest.recovery_result_path.as_deref() != Some(result_path.as_path()) {
        manifest.recovery_result_path = Some(result_path.clone());
        manifest.recovery_result_candidate_path = Some(candidate_path.clone());
        manifest.recovery_result_input_sha256 = Some(input_sha256.clone());
        manifest.recovery_result_sha256 = None;
        store.save(manifest).context("保存 S1-F01 结果发布意图")?;
    }
    ensure!(
        manifest.recovery_result_input_sha256.as_deref() == Some(input_sha256.as_str()),
        "S1-F01 结果发布意图与当前审核结果不一致"
    );

    if result_path.exists() {
        let expected_sha256 = manifest
            .recovery_result_sha256
            .as_deref()
            .context("S1-F01 结果文件已存在，但持久化状态缺少发布摘要；拒绝覆盖或复用")?;
        ensure!(
            file_sha256(&result_path)? == expected_sha256,
            "S1-F01 已发布结果文件摘要变化；拒绝复用或覆盖"
        );
        return summary_for_s1_result(&result_path, &prepared);
    }

    let persisted_candidate = manifest
        .recovery_result_candidate_path
        .as_deref()
        .context("S1-F01 结果发布状态缺少候选文件路径")?;
    ensure!(
        persisted_candidate == candidate_path,
        "S1-F01 候选结果路径与当前恢复结果不一致"
    );
    if candidate_path.exists() {
        match manifest.recovery_result_sha256.as_deref() {
            Some(expected_sha256) => ensure!(
                file_sha256(&candidate_path)? == expected_sha256,
                "S1-F01 候选结果文件摘要变化；拒绝发布"
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
                    "S1-F01 中断候选路径不在本次运行目录，拒绝删除"
                );
                fs::remove_file(&candidate_path).context("清理摘要尚未持久化的 S1-F01 中断候选")?;
            }
        }
    }

    if !candidate_path.exists() {
        manifest.recovery_result_sha256 = None;
        store.save(manifest).context("保存 S1-F01 候选创建状态")?;
        write_s1_result_workbook(&candidate_path, None, &prepared.issues)
            .context("生成 S1-F01 暂存结果工作簿")?;
        let candidate_sha256 = file_sha256(&candidate_path)?;
        manifest.recovery_result_sha256 = Some(candidate_sha256);
        store.save(manifest).context("持久化 S1-F01 候选结果摘要")?;
    }

    ensure!(
        !result_path.exists(),
        "S1-F01 结果路径在发布前被其他文件占用；拒绝覆盖"
    );
    fs::rename(&candidate_path, &result_path).with_context(|| {
        format!(
            "发布 S1-F01 结果工作簿 {} → {}",
            candidate_path.display(),
            result_path.display()
        )
    })?;
    let published_sha256 = file_sha256(&result_path)?;
    ensure!(
        manifest.recovery_result_sha256.as_deref() == Some(published_sha256.as_str()),
        "发布后的 S1-F01 结果摘要与候选摘要不一致"
    );
    summary_for_s1_result(&result_path, &prepared)
}

/// 结果 JSON 中的修复统计；未修复（含无需修复）时为 null。
fn excel_open_repair_json(
    repair: Option<&crate::excel_open_repair::ExcelOpenRepairReport>,
) -> Value {
    match repair {
        Some(repair) => json!({
            "changed": repair.changed,
            "quoted": repair.quoted,
            "paneFixed": repair.pane_fixed,
        }),
        None => Value::Null,
    }
}

/// S1-F01 审核副本的 Excel COM 打开修复完成后向运行日志说明；口径与旧
/// Python Audit 一致：只改审核副本，原始报送文件不动。
fn log_s1_excel_open_repair(
    state: Option<&AppState>,
    source: &Path,
    outcome: &S1FormulaCopyOutcome,
) {
    let repair = match outcome {
        S1FormulaCopyOutcome::Recalculated(report) => report.excel_open_repair.as_ref(),
        S1FormulaCopyOutcome::WaitingManualRecalc(report) => report.excel_open_repair.as_ref(),
    };
    let (Some(repair), Some(app)) = (repair, state) else {
        return;
    };
    let detail = [
        (repair.quoted > 0).then(|| format!("工作表引用补引号 {} 处", repair.quoted)),
        (repair.pane_fixed > 0).then(|| format!("冻结窗格修正 {} 处", repair.pane_fixed)),
    ]
    .into_iter()
    .flatten()
    .collect::<Vec<_>>()
    .join("、");
    let file_name = source
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| source.display().to_string());
    let mut ui = app.lock();
    push_log(
        &mut ui.value,
        format!(
            "{file_name}：Excel 无法打开原样副本，已在审核副本上修复（{detail}）；原始文件未改动。"
        ),
    );
}

fn execute_s1_f1_workflow(
    request: &S1F1PreflightRequest,
    accepted: &[S1F1AcceptedSource],
    work_slot: Option<CurrentWorkSlot>,
    state: Option<&AppState>,
) -> Result<S1F1WorkflowOutput> {
    let adapter_started = Instant::now();
    let session_initialization_started = Instant::now();
    let recalculation_session = std::cell::RefCell::new(if cfg!(windows) {
        Some(S1RecalculationSession::new_native(&request.calc_engine)?)
    } else {
        None
    });
    // Keep one private WPS COM session available for workbooks that Excel
    // cannot open for DisplayFormat inspection. It is started lazily and then
    // reused for the remaining files in this batch.
    let conditional_display_wps_fallback = std::cell::RefCell::new(None::<S1RecalculationSession>);
    let use_linux_lo_batch = !cfg!(windows)
        && matches!(request.calc_engine.as_str(), "auto" | "lo")
        && request.sidecar_path.is_some();
    let mut linux_batch_results: Option<VecDeque<Result<S1FormulaCopyOutcome>>> = None;
    let mut linux_batch_index = 0usize;
    let session_initialization_us = session_initialization_started.elapsed().as_micros();
    let mut display_probe = |workbook_path: &Path,
                             scan_areas: &[String]|
     -> Result<crate::conditional_format::ConditionalDisplayScan> {
        #[cfg(windows)]
        {
            let primary_scan = recalculation_session
                .borrow_mut()
                .as_mut()
                .context("S1-F01 原生 Office COM session 不可用")?
                .scan_conditional_display_colors(workbook_path, scan_areas);
            let retry_with_wps = match &primary_scan {
                Ok(scan) => !is_wps_com_engine(&scan.engine_name),
                Err(_) => true,
            };
            let scan = match primary_scan.and_then(require_complete_com_display_scan) {
                Ok(scan) => scan,
                Err(primary_error)
                    if retry_with_wps && matches!(request.calc_engine.as_str(), "auto" | "lo") =>
                {
                    let wps_scan = {
                        let mut fallback = conditional_display_wps_fallback.borrow_mut();
                        if fallback.is_none() {
                            *fallback = Some(
                                S1RecalculationSession::new_native("wps")
                                    .context("启动 S1-F01 专用 WPS 显示色回退会话")?,
                            );
                        }
                        fallback
                            .as_mut()
                            .context("S1-F01 WPS 显示色回退会话初始化失败")?
                            .scan_conditional_display_colors(workbook_path, scan_areas)
                    };
                    match wps_scan.and_then(require_complete_com_display_scan) {
                        Ok(scan) => {
                            if let Some(app) = state {
                                let mut ui = app.lock();
                                push_warn_log(
                                    &mut ui.value,
                                    format!(
                                        "S1-F01 首选 COM DisplayFormat 扫描失败，已使用复用的 WPS COM 会话完成扫描：{}；首选引擎原因：{}",
                                        workbook_path.file_name().map(|name| name.to_string_lossy()).unwrap_or_default(),
                                        primary_error
                                    ),
                                );
                            }
                            scan
                        }
                        Err(wps_error) => {
                            let error = anyhow::anyhow!(
                                "首选 Office COM 显示色扫描失败：{primary_error:#}；WPS COM 回退也失败：{wps_error:#}"
                            );
                            if let Some(app) = state {
                                let mut ui = app.lock();
                                push_warn_log(
                                    &mut ui.value,
                                    format!(
                                        "S1-F01 条件格式 COM 显示色扫描失败：{}；{}",
                                        workbook_path
                                            .file_name()
                                            .map(|name| name.to_string_lossy())
                                            .unwrap_or_default(),
                                        error
                                    ),
                                );
                            }
                            return Err(error);
                        }
                    }
                }
                Err(error) => {
                    if let Some(app) = state {
                        let mut ui = app.lock();
                        push_warn_log(
                            &mut ui.value,
                            format!(
                                "S1-F01 条件格式显示色扫描失败：{}；{}",
                                workbook_path
                                    .file_name()
                                    .map(|name| name.to_string_lossy())
                                    .unwrap_or_default(),
                                error
                            ),
                        );
                    }
                    return Err(error);
                }
            };
            if let Some(app) = state {
                let mut ui = app.lock();
                push_log(
                    &mut ui.value,
                    format!(
                        "S1-F01 条件格式显示色扫描：{}；引擎={}；扫描 {} 格；实际颜色变化 {} 处；诊断 {} 条",
                        workbook_path.file_name().map(|name| name.to_string_lossy()).unwrap_or_default(),
                        scan.engine_name.as_str(),
                        scan.scanned_cells,
                        scan.changes.len(),
                        scan.diagnostics.len()
                    ),
                );
                for diagnostic in scan.diagnostics.iter().take(3) {
                    push_warn_log(&mut ui.value, diagnostic.clone());
                }
            }
            Ok(scan)
        }
        #[cfg(not(windows))]
        {
            let _ = (workbook_path, scan_areas);
            anyhow::bail!("当前平台没有 Excel COM DisplayFormat 能力，条件格式显示色未能实测")
        }
    };
    // 条件格式判定方式：PYTHON=纯 OOXML 规则求值（不启动 COM 显示扫描）；
    // COM_DISPLAY=显示扫描作为触发事实（Windows 渲染基线）。COM_EVALUATE
    // 未实现，任务开始前明确失败。
    let conditional_format_evaluator = crate::app_settings::AppSettings::load()
        .conditional_format_evaluator()
        .to_owned();
    anyhow::ensure!(
        conditional_format_evaluator != "COM_EVALUATE",
        "条件格式判定方式「逐条公式求值」尚未在本产品实现；请在高级设置改用「规则求值（内置 Rust）」或 COM_DISPLAY"
    );
    let display_probe_enabled = conditional_format_evaluator == "COM_DISPLAY";
    let mut display_probe_opt: Option<
        &mut dyn FnMut(
            &Path,
            &[String],
        ) -> Result<crate::conditional_format::ConditionalDisplayScan>,
    > = if display_probe_enabled {
        Some(&mut display_probe)
    } else {
        None
    };
    let mut output = execute_s1_f1_workflow_with_progress_and_display(
        request,
        accepted,
        None,
        |item, request, audit_copy, modifier| {
            let outcome = if let Some(session) = recalculation_session.borrow_mut().as_mut() {
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
            } else if use_linux_lo_batch {
                if linux_batch_results.is_none() {
                    ensure!(
                        accepted.first().is_some_and(|first| {
                            web_path_identity(&first.source) == web_path_identity(&item.source)
                        }),
                        "Linux S1-F1 LibreOffice 批次没有从首个输入开始"
                    );
                    let sidecar = request
                        .sidecar_path
                        .as_deref()
                        .context("Linux S1-F1 LibreOffice 批次缺少重算 sidecar")?;
                    let batch_items = accepted
                        .iter()
                        .enumerate()
                        .map(|(index, source)| {
                            Ok((
                                source.source.clone(),
                                source.template.clone(),
                                request
                                    .run_directory
                                    .join(s1_recalculated_filename(index + 1, &source.source)?),
                            ))
                        })
                        .collect::<Result<Vec<_>>>()?;
                    ensure!(
                        batch_items
                            .first()
                            .is_some_and(|first| first.2.as_path() == audit_copy),
                        "Linux S1-F1 批次输出路径与恢复清单不一致"
                    );
                    // Formula-copy editing is independent per source. Prepare
                    // all candidates first, then reuse one LibreOffice process
                    // for the full batch. A batch failure returns retained
                    // cache-cleared copies for the existing S1 recovery flow.
                    linux_batch_results = Some(VecDeque::from(
                        execute_s1_formula_copy_batch_with_manual_recovery(
                            &batch_items,
                            request.mapping_config.as_deref(),
                            request.external_data.as_deref(),
                            sidecar,
                            modifier,
                        ),
                    ));
                }
                let expected = accepted
                    .get(linux_batch_index)
                    .context("Linux S1-F1 批次结果顺序超出输入列表")?;
                ensure!(
                    web_path_identity(&expected.source) == web_path_identity(&item.source),
                    "Linux S1-F1 批次结果与当前源文件顺序不一致"
                );
                linux_batch_index += 1;
                linux_batch_results
                    .as_mut()
                    .and_then(VecDeque::pop_front)
                    .context("Linux S1-F1 LibreOffice 批次结果缺失")?
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
            }?;
            log_s1_excel_open_repair(state, &item.source, &outcome);
            Ok(outcome)
        },
        |item, index, stage| {
            let Some(slot) = work_slot.as_ref() else {
                return;
            };
            let sheet = item.and_then(s1_f1_current_sheet);
            current_work_update(
                slot,
                json!({
                    "feature": "S1-F01",
                    "workbook": item.and_then(|item| item.source.file_name().map(|name| name.to_string_lossy().into_owned())),
                    "sheet": sheet.clone(),
                    "stage": stage,
                    "index": index,
                    "total": accepted.len(),
                }),
            );
            if let Some(app) = state {
                current_work_publish(app, slot);
                let workbook = item.map(|item| {
                    item.source
                        .strip_prefix(&request.source_directory)
                        .unwrap_or(&item.source)
                        .to_string_lossy()
                        .into_owned()
                });
                let sheet = sheet.as_deref().map(|name| format!(" 工作表={name}"));
                let mut ui = app.lock();
                push_log(
                    &mut ui.value,
                    format!(
                        "S1-F01 [{} / {}] 文件={}{} 阶段={}",
                        index.map_or_else(|| "?".to_owned(), |value| value.to_string()),
                        accepted.len(),
                        workbook.as_deref().unwrap_or("未知文件"),
                        sheet.unwrap_or_default(),
                        stage
                    ),
                );
            }
        },
        display_probe_opt,
    )?;
    let session_shutdown_started = Instant::now();
    drop(conditional_display_wps_fallback.into_inner());
    drop(recalculation_session.into_inner());
    let session_shutdown_us = session_shutdown_started.elapsed().as_micros();
    output.performance_timings["recalculationSessionInitializationUs"] =
        json!(session_initialization_us);
    output.performance_timings["recalculationSessionShutdownUs"] = json!(session_shutdown_us);
    output.performance_timings["workflowAdapterTotalUs"] =
        json!(adapter_started.elapsed().as_micros());
    Ok(output)
}

fn is_wps_com_engine(engine_name: &str) -> bool {
    engine_name.eq_ignore_ascii_case("ket.Application")
        || engine_name.eq_ignore_ascii_case("KET.Application")
        || engine_name.to_ascii_lowercase().contains("wps")
}

fn require_complete_com_display_scan(
    scan: crate::conditional_format::ConditionalDisplayScan,
) -> Result<crate::conditional_format::ConditionalDisplayScan> {
    anyhow::ensure!(
        scan.engine_name.eq_ignore_ascii_case("Excel.Application")
            || is_wps_com_engine(&scan.engine_name),
        "条件格式显示色需要 Excel/WPS COM DisplayFormat；当前引擎 {} 不作为颜色判定基准，改用静态求值并标记部分完成",
        scan.engine_name
    );
    if let Some(diagnostic) = scan.diagnostics.first() {
        anyhow::bail!(
            "Excel/WPS COM 显示色扫描不完整（{} 条诊断），改用静态求值并标记部分完成：{}",
            scan.diagnostics.len(),
            diagnostic
        );
    }
    Ok(scan)
}

fn execute_s1_f1_workflow_with_progress<F, P>(
    request: &S1F1PreflightRequest,
    accepted: &[S1F1AcceptedSource],
    recovery_store: Option<S1RecoveryStore>,
    copy_executor: F,
    publish_stage: P,
) -> Result<S1F1WorkflowOutput>
where
    F: FnMut(
        &S1F1AcceptedSource,
        &S1F1PreflightRequest,
        &Path,
        &ExcelModifier,
    ) -> Result<S1FormulaCopyOutcome>,
    P: FnMut(Option<&S1F1AcceptedSource>, Option<usize>, &str),
{
    execute_s1_f1_workflow_with_progress_and_display(
        request,
        accepted,
        recovery_store,
        copy_executor,
        publish_stage,
        None,
    )
}

fn execute_s1_f1_workflow_with_progress_and_display<F, P>(
    request: &S1F1PreflightRequest,
    accepted: &[S1F1AcceptedSource],
    recovery_store: Option<S1RecoveryStore>,
    mut copy_executor: F,
    mut publish_stage: P,
    mut display_probe: Option<
        &mut dyn FnMut(
            &Path,
            &[String],
        ) -> Result<crate::conditional_format::ConditionalDisplayScan>,
    >,
) -> Result<S1F1WorkflowOutput>
where
    F: FnMut(
        &S1F1AcceptedSource,
        &S1F1PreflightRequest,
        &Path,
        &ExcelModifier,
    ) -> Result<S1FormulaCopyOutcome>,
    P: FnMut(Option<&S1F1AcceptedSource>, Option<usize>, &str),
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
        None => S1RecoveryStore::user_default().context("初始化 S1 审核副本恢复目录")?,
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
    manifest.conditional_format_source = request.conditional_format_source.clone();
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

    for (accepted_index, item) in accepted.iter().enumerate() {
        publish_stage(Some(item), Some(accepted_index + 1), "复制公式并重算");
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
                publish_stage(Some(item), Some(accepted_index + 1), "检查待人工重算副本");
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
                    "excelOpenRepair": excel_open_repair_json(copy.excel_open_repair.as_ref()),
                    "reason": copy.reason
                }));
                file_results.push(json!({
                    "sourceFile": item.source,
                    "status": manifest.files[file_index].status.as_str(),
                    "reason": manifest.files[file_index].reason,
                    "auditWorkbook": copy.audit_copy,
                    "formulaCells": formula_count,
                    "excelOpenRepair": excel_open_repair_json(copy.excel_open_repair.as_ref())
                }));
                store.save(&mut manifest)?;
            }
            Ok(S1FormulaCopyOutcome::Recalculated(execution)) => {
                publish_stage(Some(item), Some(accepted_index + 1), "验证公式重算结果");
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
                                "重算输出中未发现公式，已保留审核副本；可用 Excel 打开后按 Ctrl+Alt+F9 强制重算查看。".to_owned(),
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
                publish_stage(
                    Some(item),
                    Some(accepted_index + 1),
                    "读取条件格式区域实际显示填充色",
                );
                let extraction_started = Instant::now();
                let conditional_workbook = s1_conditional_format_workbook(
                    &item.source,
                    &audit_workbook,
                    &request.conditional_format_source,
                );
                let (conditional_scan_areas, conditional_scan_area_source, display_observation) =
                    match s1_f1_conditional_scan_area_plan(item, conditional_workbook) {
                        Ok(plan) if plan.no_rules => (
                            plan.areas,
                            plan.source,
                            Some(Ok(crate::conditional_format::ConditionalDisplayScan {
                                engine_name: "no_conditional_rules".to_owned(),
                                ..Default::default()
                            })),
                        ),
                        Ok(plan) => {
                            let observation = display_probe.as_mut().map(|probe| {
                                (**probe)(conditional_workbook, &plan.areas)
                                    .map_err(|error| format!("{error:#}"))
                            });
                            (plan.areas, plan.source, observation)
                        }
                        Err(error) => (
                            item.preflight.scan_areas.clone(),
                            "selected_workbook_ooxml_rule_targets_unavailable",
                            Some(Err(format!("{error:#}"))),
                        ),
                    };
                publish_stage(
                    Some(item),
                    Some(accepted_index + 1),
                    "提取公式与条件格式异常",
                );
                let extracted = extract_s1_audit_copy_with_display_observation_and_scan_areas(
                    item,
                    &item.source,
                    &audit_workbook,
                    &request.conditional_format_source,
                    &metadata,
                    &conditional_scan_areas,
                    display_observation,
                    crate::app_settings::AppSettings::load()
                        .conditional_format_include_blank_cells(),
                );
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
                        .saturating_sub(file_conditional_extraction_us),
                    "conditionalScanAreaSource": conditional_scan_area_source,
                    "conditionalScanAreaCount": conditional_scan_areas.len()
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
                        manifest.files[file_index].reason =
                            Some(format!("自动重算完成，但审核异常提取失败：{error:#}"));
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
                    "excelOpenRepair": excel_open_repair_json(execution.excel_open_repair.as_ref()),
                    "formulaCells": formula_count,
                    "conditionalScanAreaSource": conditional_scan_area_source,
                    "conditionalScanAreaCount": conditional_scan_areas.len()
                }));
                file_results.push(json!({
                    "sourceFile": item.source,
                    "status": manifest.files[file_index].status.as_str(),
                    "reason": manifest.files[file_index].reason,
                    "auditWorkbook": audit_workbook,
                    "conditionalScanAreaSource": conditional_scan_area_source,
                    "conditionalScanAreaCount": conditional_scan_areas.len(),
                    "excelOpenRepair": excel_open_repair_json(execution.excel_open_repair.as_ref())
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
        publish_stage(None, None, "汇总并写入审核结果");
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

/// The S1 formula-copy edit plan builds template ImportRange target_sheet
/// values directly from these same cached copy-range definitions. Reuse the
/// preflight plan here instead of opening a workbook again for the status bar.
fn s1_f1_current_sheet(item: &S1F1AcceptedSource) -> Option<String> {
    item.preflight
        .rule_plan
        .definition
        .copy_ranges
        .first()
        .map(|range| range.sheet_name.clone())
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
    let (scan_areas, scan_area_source) = if area_plan.scan_areas.is_empty() {
        // Some production templates store conditional formatting directly in
        // worksheet XML but omit a matching `条件格式区域` defined name. In
        // that case, use only the explicit OOXML sqref targets; never expand
        // the scan to whole worksheets or the workbook used range.
        let rule_areas = excel_reader::read_conditional_formatting_areas(template_path)
            .context("读取模板 OOXML 条件格式规则目标区域")?;
        let scan_areas = rule_areas
            .into_iter()
            .map(|area| format!("'{}'!{}", area.sheet_name.replace('\'', "''"), area.address))
            .collect::<Vec<_>>();
        let source = if scan_areas.is_empty() {
            "missing"
        } else {
            "ooxml_rule_targets"
        };
        (scan_areas, source.to_owned())
    } else {
        (area_plan.scan_areas, "defined_names".to_owned())
    };
    let scan_area_count = scan_areas.len();
    let structure_area_count = area_plan.structure_ranges.len();
    Ok(S1F1TemplatePreflight {
        structure,
        rule_plan,
        conditional_mappings: mappings,
        scan_areas,
        scan_area_source,
        structure_ranges: area_plan.structure_ranges,
        scan_area_count,
        structure_area_count,
        health_passed: health.passed,
        health_errors,
    })
}

#[derive(Clone, Debug)]
struct S1F1ConditionalScanAreaPlan {
    areas: Vec<String>,
    source: &'static str,
    no_rules: bool,
}

fn resolve_s1_f1_conditional_scan_area_plan(
    selected_workbook_named_areas: &[String],
    workbook_rule_areas: &[excel_reader::ConditionalFormattingArea],
) -> Result<S1F1ConditionalScanAreaPlan> {
    if workbook_rule_areas.is_empty() {
        return Ok(S1F1ConditionalScanAreaPlan {
            areas: Vec::new(),
            source: "selected_workbook_no_conditional_rules",
            no_rules: true,
        });
    }

    // 扫描范围以 OOXML 规则 sqref 为准（Excel 真实口径：条件格式按规则
    // AppliesTo 生效）。命名区域只在规则 sqref 解析失败时兜底，避免按业务
    // 命名区域收窄而漏掉 sqref 内的触发（对齐基线 257 条含 sqref 内全部格）。
    let areas = workbook_rule_areas
        .iter()
        .map(|area| format!("'{}'!{}", area.sheet_name.replace('\'', "''"), area.address))
        .collect::<Vec<_>>();
    anyhow::ensure!(
        !areas.is_empty(),
        "条件格式规则存在，但未能从所选工作簿解析任何 sqref 扫描区域"
    );
    Ok(S1F1ConditionalScanAreaPlan {
        areas,
        source: "selected_workbook_ooxml_rule_targets",
        no_rules: false,
    })
}

fn s1_f1_conditional_scan_area_plan(
    item: &S1F1AcceptedSource,
    conditional_workbook: &Path,
) -> Result<S1F1ConditionalScanAreaPlan> {
    let rule_areas = excel_reader::read_conditional_formatting_areas(conditional_workbook)
        .map(|areas| {
            if std::env::var_os("AUDIT_DEBUG_SCAN_PLAN").is_some() {
                eprintln!(
                    "S1F01-SCAN-PLAN file={} areas={:?}",
                    item.source
                        .file_name()
                        .map(|name| name.to_string_lossy().into_owned())
                        .unwrap_or_default(),
                    areas
                        .iter()
                        .map(|area| format!("{}!{}", area.sheet_name, area.address))
                        .collect::<Vec<_>>()
                );
            }
            areas
        })
        .with_context(|| {
            format!(
                "读取所选条件格式工作簿的 OOXML 规则范围失败：{}",
                conditional_workbook.display()
            )
        })?;
    if rule_areas.is_empty() {
        return resolve_s1_f1_conditional_scan_area_plan(&[], &rule_areas);
    }
    let defined_names =
        excel_reader::read_defined_names(conditional_workbook).with_context(|| {
            format!(
                "读取所选条件格式工作簿的命名区域失败：{}",
                conditional_workbook.display()
            )
        })?;
    let sheet_names = excel_reader::read_sheet_names(conditional_workbook).with_context(|| {
        format!(
            "读取所选条件格式工作簿的工作表目录失败：{}",
            conditional_workbook.display()
        )
    })?;
    let named_areas = named_areas_from_defined_names(&defined_names);
    let name_plan = resolve_s1_conditional_area_plan(
        &item.preflight.conditional_mappings,
        None,
        &named_areas,
        &sheet_names,
    )
    .context("按所选条件格式工作簿的命名区域解析条件格式扫描范围")?;
    resolve_s1_f1_conditional_scan_area_plan(&name_plan.scan_areas, &rule_areas)
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
            format!("当前审核工具尚未接入该工作流：{workflow}"),
        ));
    }
    let task_started = Instant::now();
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
    let source_files =
        s1_region_summary_source_workbooks_with_depth(&input_dir, current_scan_depth(state))
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
        clear_run_summary(&mut ui.value);
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
                task_started,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S1-F02 汇总启动失败");
        push_error_log(&mut ui.value, format!("启动 S1-F02 后台任务失败：{error}"));
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
    task_started: Instant,
) {
    let started = Instant::now();
    let feature_names = ["任意行汇总".to_owned()];
    let progress_state = state.clone();
    let result = summarize_s1_named_regions_selected_with_progress(
        &template_path,
        &input_dir,
        &output_workbook,
        Some(&history_workbook),
        Some(&history_workbook),
        &feature_names,
        true,
        &selected_files,
        |stage| {
            let mut ui = progress_state.lock();
            ui.value["currentWork"] = json!({"feature":"S1-F02","stage":stage});
            push_log(&mut ui.value, format!("S1-F02 阶段：{stage}"));
        },
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
            publish_run_summary(
                &mut ui.value,
                "S1-F02",
                format!(
                    "工作簿总数 {}；结构通过 {}；部分继续 {}；跳过 {}；输出工作表 {} 张、数据 {} 行；结果文件 {}；总耗时 {}。",
                    summary.source_workbooks,
                    summary.structure_preflight_passed_workbooks,
                    summary.structure_preflight_partial_workbooks,
                    summary.structure_preflight_skipped_workbooks,
                    summary.output_sheets,
                    summary.data_rows,
                    summary.output_path.display(),
                    format_elapsed(task_started.elapsed())
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
            publish_run_summary(
                &mut ui.value,
                "S1-F02",
                format!(
                    "输入工作簿 {}；任务失败；结果文件未确认发布（预期路径 {}）；输出目录 {}；总耗时 {}；原因：{error:#}。",
                    selected_files.len(),
                    output_workbook.display(),
                    output_workbook.parent().unwrap_or(Path::new(".")).display(),
                    format_elapsed(task_started.elapsed())
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
        "centralOutput" => false,
        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("当前审核工具暂不支持选择该大集中路径：{field}"),
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
        }
        "centralOutput" => ui.value["centralOutputAuto"] = json!(false),
        _ => {}
    }
    drop(ui);
    push_log_for_state(state, &format!("S3 路径已选择：{field}"));
    persist_state_preferences(state)?;
    Ok(state.lock().value[field].clone())
}

fn start_central_comparison(state: &AppState) -> std::result::Result<Value, ApiError> {
    let task_started = Instant::now();
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
        clear_run_summary(&mut ui.value);
        ui.value["status"] = json!("正在执行大集中比较，请勿关闭窗口……");
        ui.value["currentWork"] = json!({
            "feature": "S3-F01",
            "workbook": current_file.display().to_string(),
            "stage": "准备执行比较",
        });
        push_log(&mut ui.value, "S3-F01 开始大集中执行比较".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F01 开始：input_reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=内置求值（{schema_name}）；本期 {}；上期 {}；3.0配置 {}；3.1配置 {}；输出 {}",
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
                task_started,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["currentWork"] = Value::Null;
        ui.value["status"] = json!("大集中执行比较失败");
        push_error_log(&mut ui.value, format!("启动 S3-F01 比较线程失败：{error}"));
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
    task_started: Instant,
) {
    let started = Instant::now();
    let current_file_display = current_file.display().to_string();
    let previous_file_display = previous_file.display().to_string();
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
            ui.value["currentWork"] = json!({
                "feature": "S3-F01",
                "workbook": current_file_display.clone(),
                "stage": step,
            });
            push_log(
                &mut ui.value,
                format!("S3-F01 阶段：文件={}；{step}", current_file_display),
            );
            push_detail_log(
                &mut ui.value,
                format!("S3-F01 阶段：{step}；expression={schema_name}"),
            );
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.value["currentWork"] = Value::Null;
    match result {
        Ok(report) => {
            let output = report.output_path.to_string_lossy().into_owned();
            ui.value["centralCompareFile"] = json!(output);
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
            publish_run_summary(
                &mut ui.value,
                "S3-F01",
                format!(
                    "本期文件 {}；上期文件 {}；比较行 {}；表达式错误 {}；结果文件 {}；总耗时 {}。",
                    current_file_display,
                    previous_file_display,
                    report.rows,
                    report.expression_errors,
                    report.output_path.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
            // 高级设置：完成后用系统默认程序自动打开比较结果工作簿（默认关闭）。
            if crate::app_settings::AppSettings::load().central_auto_open_result() {
                match open_path_with_default_app(&report.output_path) {
                    Ok(()) => push_log(
                        &mut ui.value,
                        format!("已自动打开比较结果：{}", report.output_path.display()),
                    ),
                    Err(error) => push_log(
                        &mut ui.value,
                        format!(
                            "自动打开比较结果失败（{}）：{error}",
                            report.output_path.display()
                        ),
                    ),
                }
            }
        }
        Err(error) => {
            ui.value["status"] = json!("大集中执行比较失败");
            push_error_log(
                &mut ui.value,
                format!(
                    "S3-F01 文件={} 阶段=执行比较 失败，未报告成功结果：{error:#}",
                    current_file_display
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F01 失败：reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=内置求值（{schema_name}），任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S3-F01",
                format!(
                    "本期文件 {}；上期文件 {}；任务失败；比较结果未生成；输出文件 {}；总耗时 {}；原因：{error:#}。",
                    current_file_display,
                    previous_file_display,
                    output_workbook.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
    }
}

fn start_central_explanation_export(state: &AppState) -> std::result::Result<Value, ApiError> {
    let task_started = Instant::now();
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
            .map(PathBuf::from)
            .or_else(|| comparison_file.parent().map(Path::to_path_buf))
            .map(|path| {
                if ui.value["centralOutput"]
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
        clear_run_summary(&mut ui.value);
        ui.value["status"] = json!("正在导出选定指标说明……");
        ui.value["currentWork"] = json!({
            "feature": "S3-F02",
            "workbook": comparison_file.file_name().map(|name| name.to_string_lossy().into_owned()),
            "stage": "导出说明",
        });
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
                task_started,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["currentWork"] = Value::Null;
        ui.value["status"] = json!("指标说明导出失败");
        push_error_log(&mut ui.value, format!("启动 S3-F02 导出线程失败：{error}"));
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
    task_started: Instant,
) {
    let started = Instant::now();
    let progress_state = state.clone();
    let result = export_selected_explanations_with_progress(
        &comparison_file,
        &output_workbook,
        config_workbook.as_deref(),
        |stage| {
            let mut ui = progress_state.lock();
            ui.value["currentWork"] = json!({
                "feature": "S3-F02",
                "workbook": comparison_file.to_string_lossy(),
                "stage": stage,
            });
            push_log(&mut ui.value, format!("S3-F02 阶段：{stage}"));
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.value["currentWork"] = Value::Null;
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
            publish_run_summary(
                &mut ui.value,
                "S3-F02",
                format!(
                    "输入工作簿 {}；选中说明行 {} 条；结果文件 {}；总耗时 {}。",
                    comparison_file.display(),
                    report.selected_rows,
                    report.output_workbook.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("指标说明导出失败");
            let file_name = comparison_file
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| comparison_file.display().to_string());
            push_error_log(
                &mut ui.value,
                format!("S3-F02 文件={file_name} 阶段=导出说明 失败：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F02 失败：reader=Calamine，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S3-F02",
                format!(
                    "输入工作簿 {}；任务失败；结果未生成；输出目录 {}；总耗时 {}；原因：{error:#}。",
                    comparison_file.display(),
                    output_workbook.parent().unwrap_or(Path::new(".")).display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
    }
}

fn start_central_explanation_feedback_import(
    state: &AppState,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    let task_started = Instant::now();
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
        clear_run_summary(&mut ui.value);
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
                task_started,
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
    task_started: Instant,
) {
    let started = Instant::now();
    let progress_state = state.clone();
    let result = execute_explanation_feedback_import_with_progress(
        &target_workbook,
        &feedback_workbooks,
        &modifier,
        |stage| {
            let mut ui = progress_state.lock();
            ui.value["currentWork"] = json!({
                "feature": "S3-F03",
                "workbook": target_workbook.to_string_lossy(),
                "stage": stage,
            });
            push_log(&mut ui.value, format!("S3-F03 阶段：{stage}"));
        },
    );
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
            publish_run_summary(
                &mut ui.value,
                "S3-F03",
                format!(
                    "目标说明文件 {}；反馈文件 {} 个；写入 {} 条；未匹配 {} 条；冲突 {} 条；目标已原地更新；备份 {}；总耗时 {}。",
                    report.output_workbook.display(),
                    feedback_workbooks.len(),
                    report.imported,
                    report.unmatched,
                    report.conflicts,
                    report.backup_workbook.as_ref().map(|path| path.display().to_string()).unwrap_or_else(|| "未生成".to_owned()),
                    format_elapsed(task_started.elapsed())
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
            publish_run_summary(
                &mut ui.value,
                "S3-F03",
                format!(
                    "目标说明文件 {}；反馈文件 {} 个；原目标未报告更新成功；总耗时 {}；错误：{error:#}。",
                    target_workbook.display(),
                    feedback_workbooks.len(),
                    format_elapsed(task_started.elapsed())
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
                .unwrap_or("auto")
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
    crate::app_settings::AppSettings::store_preferences(&calc_engine, &edit_engine, preferences)
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
    let task_started = Instant::now();
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
        // 金融表单输出目录 = 主输出目录（用户确认 2026-11-09：不再单设）。
        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| comparison_file.parent().map(Path::to_path_buf))
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择输出目录"))?;
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
        clear_run_summary(&mut ui.value);
        ui.value["status"] = json!("正在生成 S3-F05 金融表单……");
        ui.value["currentWork"] = json!({
            "feature": "S3-F05",
            "workbook": comparison_file.file_name().map(|name| name.to_string_lossy().into_owned()),
            "stage": "生成金融表单",
        });
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
                task_started,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["currentWork"] = Value::Null;
        ui.value["status"] = json!("S3-F05 任务启动失败");
        push_error_log(&mut ui.value, format!("启动 S3-F05 任务线程失败：{error}"));
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
    task_started: Instant,
) {
    let started = Instant::now();
    let progress_state = state.clone();
    let result = execute_indicator_comparison_split_with_progress(
        &comparison_file,
        &forms_config,
        &common_config,
        &output_directory,
        &modifier,
        fast_render_mode(),
        |stage| {
            let mut ui = progress_state.lock();
            ui.value["currentWork"] = json!({
                "feature": "S3-F05",
                "workbook": comparison_file.to_string_lossy(),
                "stage": stage,
            });
            push_log(&mut ui.value, format!("S3-F05 阶段：{stage}"));
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.value["currentWork"] = Value::Null;
    match result {
        Ok(report) => {
            let output = report.output_directory.to_string_lossy().into_owned();
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
            publish_run_summary(
                &mut ui.value,
                "S3-F05",
                format!(
                    "比较行 {}；机构 {} 家；匹配指标 {}；生成工作簿 {} 个；输出目录 {}；总耗时 {}。",
                    report.comparison_rows,
                    report.institution_count,
                    report.matched_indicators,
                    report.output_files.len(),
                    report.output_directory.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("S3-F05 失败");
            let file_name = comparison_file
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| comparison_file.display().to_string());
            push_error_log(
                &mut ui.value,
                format!("S3-F05 文件={file_name} 阶段=拆分金融表单 失败：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F05 失败；未发布输出；任务总计={} ms；错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S3-F05",
                format!(
                    "输入比较文件 {}；批次失败；未发布输出；输出根目录 {}；总耗时 {}；原因：{error:#}。",
                    comparison_file.display(),
                    output_directory.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
    }
}

fn start_central_cross_period(state: &AppState) -> std::result::Result<Value, ApiError> {
    let task_started = Instant::now();
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
        clear_run_summary(&mut ui.value);
        ui.value["status"] = json!("正在执行本期数值核对，请勿关闭窗口……");
        ui.value["currentWork"] = json!({
            "feature": "S3-F04",
            "workbook": current_file.file_name().map(|name| name.to_string_lossy().into_owned()),
            "stage": "数值核对",
        });
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
                task_started,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["currentWork"] = Value::Null;
        ui.value["status"] = json!("本期数值核对失败");
        push_error_log(&mut ui.value, format!("启动 S3-F04 任务线程失败：{error}"));
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
    task_started: Instant,
) {
    let started = Instant::now();
    let input_reader = input_reader_for_paths(std::slice::from_ref(&current_file));
    let progress_state = state.clone();
    let result = run_current_value_check_files_with_progress(
        std::slice::from_ref(&current_file),
        &[],
        &config_workbooks,
        &output_workbook,
        log_workbook.as_deref(),
        python_regex_scope_matcher(),
        |stage| {
            let mut ui = progress_state.lock();
            ui.value["currentWork"] = json!({
                "feature": "S3-F04",
                "workbook": current_file.file_name().map(|name| name.to_string_lossy().into_owned()),
                "stage": stage,
            });
            push_log(
                &mut ui.value,
                format!("S3-F04 阶段：文件={}；{}", current_file.display(), stage),
            );
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.value["currentWork"] = Value::Null;
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
            publish_run_summary(
                &mut ui.value,
                "S3-F04",
                format!(
                    "输入文件 {}；规则 {} 条（启用 {}）；输入记录 {}；结果行 {}；输出文件 {}；总耗时 {}。",
                    current_file.display(),
                    report.total_rules,
                    report.enabled_rules,
                    report.current_records,
                    report.rows,
                    report.output_workbook.display(),
                    format_elapsed(task_started.elapsed())
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("本期数值核对失败");
            let file_name = current_file
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| current_file.display().to_string());
            push_error_log(
                &mut ui.value,
                format!("S3-F04 文件={file_name} 阶段=本期数值核对 失败：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F04 失败：reader={input_reader}，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
            publish_run_summary(
                &mut ui.value,
                "S3-F04",
                format!(
                    "输入文件 {}；任务失败；结果未生成；输出目录 {}；总耗时 {}；原因：{error:#}。",
                    current_file.display(),
                    output_workbook.parent().unwrap_or(Path::new(".")).display(),
                    format_elapsed(task_started.elapsed())
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

        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("当前审核工具暂不支持固定该输出目录：{kind}"),
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
        | "centralOutput"),
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
            format!("当前审核工具暂不支持清除该路径：{kind}"),
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
    let task_started = Instant::now();
    let (request, scan_depth) = {
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
        let scan_depth = scan_depth_from_value(&ui.value["recursiveDepth"]);
        ui.value["busy"] = json!(true);
        clear_run_summary(&mut ui.value);
        ui.value["cancelRequested"] = json!(false);
        ui.value["periodSummary"] = json!({});
        ui.value["status"] = json!("正在检查审核配置……");
        ui.value["periodProgress"] = json!({
            "status": "running", "stage": "config", "label": "检查审核配置",
            "percent": 0, "current": 0, "total": 0,
            "message": "准备执行 S2 审核", "elapsedSeconds": 0
        });
        ui.started_at = Some(task_started);
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
                "S2 开始：本期 {}；上期 {}；输出 {}",
                request.current_directory.display(),
                request.previous_directory.display(),
                request.output_directory.display()
            ),
        );
        (request, scan_depth)
    };

    let worker_state = state.clone();
    thread::Builder::new()
        .name("audit-s2-service".to_owned())
        .spawn(move || run_s2_worker(worker_state, request, scan_depth, task_started))
        .map_err(|error| {
            let mut ui = state.lock();
            ui.value["busy"] = json!(false);
            ui.value["currentWork"] = Value::Null;
            ui.value["status"] = json!("无法启动 S2 审核线程");
            publish_run_summary(
                &mut ui.value,
                "S2-F01",
                format!(
                    "任务启动失败；结果未生成；总耗时 {}；原因：{error}。",
                    format_elapsed(task_started.elapsed())
                ),
            );
            ApiError::internal(format!("启动审核线程失败：{error}"))
        })?;
    Ok(json!(true))
}

fn run_s2_worker(
    state: AppState,
    request: S2RunRequest,
    scan_depth: ScanDepth,
    task_started: Instant,
) {
    let mut active_stage: Option<(String, String, Instant)> = None;
    let mut active_workbook: Option<String> = None;
    let mut last_logged_workbook: Option<String> = None;
    let mut on_progress = |event: S2ProgressEvent| {
        let event_workbook = if matches!(event.stage.as_str(), "current" | "previous")
            && Path::new(&event.message)
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xls") || extension.eq_ignore_ascii_case("xlsx")
                }) {
            Some(event.message.clone())
        } else {
            None
        };
        active_workbook = event_workbook.clone();
        let stage_changed = active_stage
            .as_ref()
            .map(|(stage, _, _)| stage != &event.stage)
            .unwrap_or(true);
        let mut ui = state.lock();
        if stage_changed {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段完成：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let engine = s2_engine_for_stage(&event.stage).unwrap_or("未映射");
            push_log(
                &mut ui.value,
                format!(
                    "S2 阶段开始：{}；stage={}；engine={}；进度={}/{}；{}",
                    event.label, event.stage, engine, event.current, event.total, event.message
                ),
            );
            active_stage = Some((event.stage.clone(), engine.to_owned(), Instant::now()));
        }
        if let Some(workbook) = event_workbook.as_ref() {
            if last_logged_workbook.as_deref() != Some(workbook.as_str()) {
                push_log(
                    &mut ui.value,
                    format!(
                        "S2 工作簿 {}/{}：{}；当前阶段={}",
                        event.current, event.total, workbook, event.label
                    ),
                );
                last_logged_workbook = Some(workbook.clone());
            }
        }
        let stage_elapsed = active_stage
            .as_ref()
            .map(|(_, _, started)| started.elapsed().as_secs_f64())
            .unwrap_or_default();
        ui.value["status"] = json!(format!("{}……", event.label));
        ui.value["currentWork"] = json!({
            "feature": "S2-F01",
            "stage": event.label,
            "stageKey": event.stage,
            "workbook": event_workbook,
            "detail": event.message,
            "index": event.current,
            "total": event.total,
        });
        ui.value["periodProgress"] = progress_value(
            &event,
            "running",
            task_started.elapsed().as_secs_f64(),
            stage_elapsed,
        );
    };
    let result = run_s2_audit_with_scan_depth(&request, scan_depth, Some(&mut on_progress));
    drop(on_progress);
    let failed_file = active_workbook.clone();
    let failed_stage = active_stage.as_ref().map(|(stage, _, _)| stage.clone());
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.value["currentWork"] = Value::Null;
    ui.started_at = None;
    match result {
        Ok(result) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_log(
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
                "elapsedSeconds": task_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较完成：已生成审核结果");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务完成：reader=Calamine writer=rust_xlsxwriter findings={} total_ms={}",
                    summary.total_findings,
                    task_started.elapsed().as_millis()
                ),
            );
            push_log(&mut ui.value, format!("审核完成：{}", report_path));
            let final_summary = format!(
                "机构 {} 家；表单 {} 张；指标 {} 个；异常 {} 条（总计 {} 条发现）；结果文件 {}；总耗时 {}。",
                summary.institutions,
                summary.forms,
                summary.indicators,
                summary.abnormal_findings,
                summary.total_findings,
                report_path,
                format_elapsed(task_started.elapsed())
            );
            ui.value["status"] = json!(format!(
                "跨期比较完成；总耗时 {}",
                format_elapsed(task_started.elapsed())
            ));
            publish_run_summary(&mut ui.value, "S2-F01", final_summary);
        }
        Err(error) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段失败：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let message = format!(
                "S2-F01{}{} 跨期比较失败：{error:#}",
                failed_file
                    .as_deref()
                    .map(|file| format!(" 文件={file}"))
                    .unwrap_or_default(),
                failed_stage
                    .as_deref()
                    .map(|stage| format!(" 阶段={stage}"))
                    .unwrap_or_default()
            );
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "failed", "label": "审核失败",
                "percent": 0, "current": 0, "total": 0,
                "message": message,
                "elapsedSeconds": task_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较失败");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务失败：total_ms={} error={error:#}",
                    task_started.elapsed().as_millis()
                ),
            );
            push_error_log(&mut ui.value, message);
            publish_run_summary(
                &mut ui.value,
                "S2-F01",
                format!(
                    "任务失败；位置={}；结果未生成；总耗时 {}；原因：{error:#}。",
                    failed_file
                        .as_deref()
                        .map(|file| format!(
                            "文件 {file}，阶段 {}",
                            failed_stage.as_deref().unwrap_or("未知阶段")
                        ))
                        .unwrap_or_else(|| "未能定位到工作簿".to_owned()),
                    format_elapsed(task_started.elapsed())
                ),
            );
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
    let mut recursive_depth_changed = false;
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
            "recursiveDepth" => {
                let Some(depth) = value.as_i64().filter(|depth| (-1..=10).contains(depth)) else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "递归扫描深度必须是 -1 到 10 的整数（-1 表示最深层，0 表示当前目录）",
                    ));
                };
                recursive_depth_changed = ui.value["recursiveDepth"].as_i64() != Some(depth);
                ui.value["recursiveDepth"] = json!(depth);
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
            "conditionalFormatIncludeBlankCells"
            | "centralTrimIndicatorNames"
            | "centralAutoOpenResult" => {
                if !value.is_boolean() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是布尔值"),
                    ));
                }
                ui.value[key] = value.clone();
            }
            "conditionalFormatEvaluator" => {
                let text = value.as_str().unwrap_or_default();
                if text == "COM_EVALUATE" {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "COM_EVALUATE（Excel/WPS 公式引擎逐条求值）尚未在本产品实现；可用「规则求值（内置 Rust）」或 COM_DISPLAY（实际显示填充）",
                    ));
                }
                let validated =
                    crate::app_settings::validate_preference(key, value).map_err(|error| {
                        ApiError::new(StatusCode::BAD_REQUEST, format!("{error:#}"))
                    })?;
                ui.value[key] = validated;
                push_log(
                    &mut ui.value,
                    format!(
                        "条件格式判定方式已更新：{}",
                        if text == "COM_DISPLAY" {
                            "COM_DISPLAY（Excel/WPS 实际显示填充）"
                        } else {
                            "规则求值（内置 Rust）"
                        }
                    ),
                );
            }
            "centralTolerance" => {
                let validated =
                    crate::app_settings::validate_preference(key, value).map_err(|error| {
                        ApiError::new(StatusCode::BAD_REQUEST, format!("{error:#}"))
                    })?;
                ui.value[key] = validated;
            }
            "s1ConditionalFormatSource" => {
                let source =
                    crate::app_settings::validate_preference(key, value).map_err(|error| {
                        ApiError::new(StatusCode::BAD_REQUEST, format!("{error:#}"))
                    })?;
                let selected_source = source.as_str().unwrap_or("source").to_owned();
                ui.value[key] = source;
                push_log(
                    &mut ui.value,
                    format!(
                        "S1-F01 条件格式提取来源已更新：{}",
                        if selected_source == "audit_copy" {
                            "复制校验区域并重算后的审核副本"
                        } else {
                            "原始源文件"
                        }
                    ),
                );
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
                            .unwrap_or("auto"),
                    );
                if edit == "openpyxl" {
                    let (available, detail) = crate::app_settings::openpyxl_selection_status();
                    if !available {
                        return Err(ApiError::new(
                            StatusCode::BAD_REQUEST,
                            format!("openpyxl 修改引擎不可用：{detail}"),
                        ));
                    }
                }
                let settings =
                    crate::app_settings::AppSettings::store(calc, edit).map_err(|error| {
                        ApiError::new(StatusCode::BAD_REQUEST, format!("{error:#}"))
                    })?;
                ui.value["advancedSettings"]["calcEngine"] = json!(settings.calc_engine.as_str());
                ui.value["advancedSettings"]["editEngine"] = json!(settings.edit_engine.as_str());
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
                    format!("当前审核工具暂不支持修改该设置：{key}"),
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
            .unwrap_or("auto")
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
        .map_err(|error| ApiError::internal(format!("保存用户设置失败：{error:#}")))?;
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
    let updated = ui.value.clone();
    drop(ui);
    if recursive_depth_changed {
        refresh_s1_sources_for_scan_depth(state);
        refresh_period_pairs(state);
        return Ok(state.get_state());
    }
    Ok(updated)
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
    let (current, previous, old_warning, scan_depth) = {
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
            scan_depth_from_value(&ui.value["recursiveDepth"]),
        )
    };
    let warning =
        report_source_warning_with_scan_depth(current.as_deref(), previous.as_deref(), scan_depth);
    let pairs =
        list_period_pairs_with_scan_depth(current.as_deref(), previous.as_deref(), scan_depth);
    let mut ui = state.lock();
    ui.value["periodSourceWarning"] = json!(warning);
    match pairs {
        Ok(pairs) => {
            ui.value["periodPairs"] = serde_json::to_value(pairs).unwrap_or_else(|_| json!([]));
        }
        Err(error) => {
            ui.value["periodPairs"] = json!([]);
            push_error_log(&mut ui.value, format!("两期文件配对失败：{error:#}"));
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

/// 运行日志环形缓冲上限。2000 条覆盖 S1 批量任务（33 文件 × 每文件多条）
/// 加 S3 阶段明细，内存占用约 2MB 以内；旧的最早条目被挤出属预期行为。
const LOG_CAPACITY: usize = 2000;

/// 日志条目的显示时间（本地时区 HH:MM:SS）。
fn log_timestamp() -> String {
    chrono::Local::now().format("%H:%M:%S").to_string()
}

/// 统一的日志写入入口：`level` 取 "info"|"warn"|"error"；`detail` 条目
/// 是开发诊断行（默认隐藏），级别照记以便筛选。
fn push_log_entry(state: &mut Value, message: String, level: &str, detail: bool) {
    let is_repeated_tail = state
        .get("log")
        .and_then(Value::as_array)
        .and_then(|entries| entries.last())
        .is_some_and(|entry| {
            entry["text"].as_str() == Some(message.as_str())
                && entry["level"].as_str() == Some(level)
                && entry["detail"].as_bool() == Some(detail)
        });
    if is_repeated_tail {
        return;
    }
    let sequence = state["logVersion"].as_u64().unwrap_or_default() + 1;
    state["logVersion"] = json!(sequence);
    let Some(log) = state.get_mut("log").and_then(Value::as_array_mut) else {
        return;
    };
    log.push(json!({
        "seq": sequence,
        "text": message,
        "detail": detail,
        "level": level,
        "time": log_timestamp(),
    }));
    if log.len() > LOG_CAPACITY {
        let excess = log.len() - LOG_CAPACITY;
        log.drain(0..excess);
    }
}

/// S1-F01 本次实际走过的技术路径（详细日志）。
///
/// 把「哪个编辑后端、哪个重算引擎、条件格式从哪读、是否发生回退」拼成一行，
/// 避免用户只能从"慢"推断"是不是没用上 betteroffice"。设置里的
/// 「显示详细运行日志」关闭时不会输出。
fn s1_f1_technical_path(request: &S1F1PreflightRequest, output: &S1F1WorkflowOutput) -> String {
    let mut parts: Vec<String> = Vec::new();

    // 1) 编辑后端：SOURCE 预检之后生成审核副本所用的引擎。
    let edit_engine = request.edit_engine.as_str();
    let edit_route = match edit_engine {
        "auto" => "OOXML 直改 → openpyxl（可用时）→ Umya",
        "ooxml" => "OOXML 直改（后端不支持时拒绝，不回退）",
        "openpyxl" => "openpyxl worker",
        "umya" => "Umya",
        other => other,
    };
    parts.push(format!("审核副本编辑={edit_route}"));

    // 2) 重算引擎：实际生效的那个，以及有没有从 betteroffice 回退。
    let engine = if output.recalculation_engines.is_empty() {
        "无（没有工作簿完成真实重算）".to_owned()
    } else {
        output.recalculation_engines.join("、")
    };
    parts.push(format!("校验区域重算={engine}"));
    if output
        .recalculation_engines
        .iter()
        .any(|label| label.contains("betteroffice") && label.contains("回退"))
    {
        parts.push("重算回退=betteroffice 被拒绝后改走 COM/LibreOffice".to_owned());
    }

    // 3) 条件格式：显示色来自重算后的审核副本，还是源文件/静态规则。
    let cf_source = match request.conditional_format_source.as_str() {
        "audit_copy" => "复制校验区域并重算后的审核副本",
        "source" => "原始源文件",
        other => other,
    };
    parts.push(format!("条件格式来源={cf_source}"));

    // 4) 成本结构：让"慢"能定位到具体阶段。
    let timings = &output.performance_timings;
    let ms = |key: &str| -> u128 {
        timings
            .get(key)
            .and_then(serde_json::Value::as_u64)
            .map(|value| value as u128 / 1000)
            .unwrap_or(0)
    };
    let copy_ms = ms("copyAndRecalculationPreparationTotalUs");
    let recalc_ms = ms("externalExcelRecalculationUs");
    let extract_ms = ms("issueExtractionTotalUs");
    let total_ms = ms("workflowTotalUs");
    if total_ms > 0 {
        parts.push(format!(
            "耗时(ms)=合计 {total_ms}｜副本准备 {copy_ms}｜重算 {recalc_ms}｜异常提取 {extract_ms}"
        ));
    }

    parts.push(format!(
        "源文件 {} 个；审核副本 {} 个",
        output.file_results.len(),
        output.audit_copies.len()
    ));
    parts.join("；")
}

fn push_detail_log(state: &mut Value, message: String) {
    push_log_entry(state, message, "info", true);
}

fn push_log(state: &mut Value, message: String) {
    push_log_entry(state, message, "info", false);
}

/// 非致命异常、跳过文件、兼容性问题：醒目但不中断。
fn push_warn_log(state: &mut Value, message: String) {
    push_log_entry(state, message, "warn", false);
}

/// 任务失败、文件处理失败、关键操作异常。
fn push_error_log(state: &mut Value, message: String) {
    push_log_entry(state, message, "error", false);
}

fn clear_run_summary(state: &mut Value) {
    state["lastRunSummary"] = Value::Null;
}

fn publish_run_summary(state: &mut Value, feature: &str, summary: String) {
    state["lastRunSummary"] = json!({
        "feature": feature,
        "text": summary,
        "time": log_timestamp(),
    });
    push_log(state, format!("{feature} 最终摘要：{summary}"));
}

fn format_elapsed(elapsed: Duration) -> String {
    let total_ms = elapsed.as_millis();
    if total_ms >= 60_000 {
        format!(
            "{}分{:.1}秒",
            total_ms / 60_000,
            (total_ms % 60_000) as f64 / 1_000.0
        )
    } else if total_ms < 1_000 {
        let total_ms = (elapsed.as_secs_f64() * 1_000.0).round() as u64;
        if total_ms == 0 {
            "小于 1 毫秒".to_owned()
        } else {
            format!("{total_ms} 毫秒")
        }
    } else {
        format!("{:.1}秒", total_ms as f64 / 1_000.0)
    }
}

/// "当前处理对象"共享槽：worker 在业务线程直接更新，800ms 轮询把
/// `currentWork` 带给前端；与历史日志分离——正常切换不写日志。
/// 结构：{feature, workbook, sheet, stage, index, total, done, failed, skipped}
type CurrentWorkSlot = std::sync::Arc<std::sync::Mutex<Value>>;

fn current_work_update(slot: &CurrentWorkSlot, patch: Value) {
    let mut work = slot.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    if !work.is_object() {
        *work = json!({});
    }
    if let Some(object) = patch.as_object() {
        for (key, value) in object {
            work[key.as_str()] = value.clone();
        }
    }
}

/// 把共享槽的内容合并进 GUI 状态（轮询返回前调用）。
fn current_work_publish(state: &AppState, slot: &CurrentWorkSlot) {
    let snapshot = {
        let work = slot.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
        work.clone()
    };
    let mut ui = state.lock();
    ui.value["currentWork"] = snapshot;
}

/// 任务结束：清空当前对象并发布最终计数（避免遗留"正在处理"误导）。
fn current_work_finish(state: &AppState, slot: &CurrentWorkSlot) {
    {
        let mut work = slot.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
        *work = json!({});
    }
    current_work_publish(state, slot);
}

/// 终止请求：S1/S2/S3 的业务 worker 内部没有取消检查点（Excel 重算、
/// 批量写入均为不可中断的外部调用），无法安全中途终止。此处如实告知，
/// 不伪造取消成功，也不把失败掩盖为正常结束。
fn cancel_execution(state: &AppState) -> std::result::Result<Value, ApiError> {
    let busy = state.lock().value["busy"] == true;
    if !busy {
        return Ok(json!({
            "requested": false,
            "message": "当前没有正在运行的任务。"
        }));
    }
    {
        let mut ui = state.lock();
        push_warn_log(
            &mut ui.value,
            "已收到终止请求；当前版本的任务在文件处理中途不可安全终止，请等待当前文件完成。任务结束后会给出完整结果摘要。"
                .to_owned(),
        );
    }
    Ok(json!({
        "requested": false,
        "message": "当前版本不支持中途终止：文件处理中断会导致输出不完整。请等待当前文件处理完成。"
    }))
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
        browse_dir, cached_s1_f1_template_preflight, execute_s1_f1_workflow_with_progress,
        extract_s1_audit_copy_with_conditional_extractor, format_elapsed, progress_value,
        publish_run_summary, push_log, render_index, require_complete_com_display_scan,
        run_s1_f1_for_acceptance, s1_f1_current_sheet, s1_f1_template_preflight,
        write_s1_recovery_manifest_result, AppState, S1F1AcceptedSource, S1F1OutputDirectoryGuard,
        S1F1PreflightRequest, LOG_CAPACITY,
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

    #[test]
    fn conditional_scan_areas_fall_back_to_selected_workbook_rule_targets() {
        let source_rules = [excel_reader::ConditionalFormattingArea {
            sheet_name: "指标'sheet".to_owned(),
            address: "A2:C9".to_owned(),
        }];
        let plan = super::resolve_s1_f1_conditional_scan_area_plan(&[], &source_rules)
            .expect("derive scan target from selected workbook OOXML rules");

        assert_eq!(plan.source, "selected_workbook_ooxml_rule_targets");
        assert!(!plan.no_rules);
        assert_eq!(plan.areas, ["'指标''sheet'!A2:C9"]);
    }

    #[test]
    // 扫描范围已改为 OOXML sqref 优先（Excel 真实口径）；命名区域不再收窄。
    #[test]
    fn conditional_scan_areas_prefer_ooxml_rule_targets_over_defined_names() {
        let selected_workbook_names = ["'指标'!B5:B20".to_owned()];
        let source_rules = [excel_reader::ConditionalFormattingArea {
            sheet_name: "指标".to_owned(),
            address: "D1:D50".to_owned(),
        }];
        let plan = super::resolve_s1_f1_conditional_scan_area_plan(
            &selected_workbook_names,
            &source_rules,
        )
        .expect("use selected workbook named range");

        assert_eq!(plan.source, "selected_workbook_ooxml_rule_targets");
        assert_eq!(plan.areas, vec!["'指标'!D1:D50".to_owned()]);
        assert!(!plan.no_rules);
    }

    #[test]
    fn conditional_scan_areas_report_selected_workbook_without_rules() {
        let plan = super::resolve_s1_f1_conditional_scan_area_plan(&[], &[])
            .expect("no rules is a valid empty plan");

        assert_eq!(plan.source, "selected_workbook_no_conditional_rules");
        assert!(plan.areas.is_empty());
        assert!(plan.no_rules);
    }

    #[test]
    fn conditional_display_scan_accepts_complete_wps_com_observation() {
        let scan = crate::conditional_format::ConditionalDisplayScan {
            engine_name: "ket.Application".to_owned(),
            scanned_cells: 1,
            ..Default::default()
        };

        assert!(require_complete_com_display_scan(scan).is_ok());
    }

    #[test]
    fn conditional_display_scan_rejects_incomplete_com_observation() {
        let scan = crate::conditional_format::ConditionalDisplayScan {
            engine_name: "ket.Application".to_owned(),
            scanned_cells: 1,
            diagnostics: vec!["Sheet1!A1 颜色不可读".to_owned()],
            ..Default::default()
        };

        let error = require_complete_com_display_scan(scan)
            .expect_err("incomplete COM observation must use static fallback");
        assert!(error.to_string().contains("扫描不完整"));
        assert!(error.to_string().contains("颜色不可读"));
    }

    #[test]
    fn conditional_display_scan_accepts_complete_excel_observation() {
        let scan = crate::conditional_format::ConditionalDisplayScan {
            engine_name: "Excel.Application".to_owned(),
            scanned_cells: 1,
            ..Default::default()
        };

        assert!(require_complete_com_display_scan(scan).is_ok());
    }

    #[test]
    fn conditional_display_scan_rejects_non_office_engine() {
        let scan = crate::conditional_format::ConditionalDisplayScan {
            engine_name: "LibreOffice".to_owned(),
            ..Default::default()
        };

        let error = require_complete_com_display_scan(scan)
            .expect_err("non-COM engine must not be treated as DisplayFormat observation");
        assert!(error.to_string().contains("Excel/WPS COM DisplayFormat"));
    }

    #[test]
    fn run_log_keeps_append_sequence_and_final_summary_beyond_ring_buffer() {
        let mut state = json!({"log": [], "logVersion": 0, "lastRunSummary": null});
        push_log(&mut state, "同一阶段".to_owned());
        push_log(&mut state, "同一阶段".to_owned());
        assert_eq!(state["log"].as_array().unwrap().len(), 1);
        assert_eq!(state["logVersion"], 1);

        for index in 0..LOG_CAPACITY + 1 {
            push_log(&mut state, format!("阶段 {index}"));
        }
        publish_run_summary(
            &mut state,
            "S3-F05",
            "生成工作簿 2 个；总耗时 31 毫秒。".to_owned(),
        );

        let entries = state["log"].as_array().unwrap();
        assert_eq!(entries.len(), LOG_CAPACITY);
        assert!(entries
            .windows(2)
            .all(|pair| pair[0]["seq"].as_u64() < pair[1]["seq"].as_u64()));
        assert!(entries.last().unwrap()["text"]
            .as_str()
            .unwrap()
            .contains("S3-F05 最终摘要"));
        assert_eq!(
            state["lastRunSummary"]["text"],
            "生成工作簿 2 个；总耗时 31 毫秒。"
        );
    }

    #[test]
    fn elapsed_time_format_does_not_round_short_runs_to_zero_seconds() {
        assert_eq!(format_elapsed(Duration::from_millis(37)), "37 毫秒");
        assert_eq!(format_elapsed(Duration::ZERO), "小于 1 毫秒");
        assert_eq!(format_elapsed(Duration::from_millis(1_234)), "1.2秒");
        assert_eq!(format_elapsed(Duration::from_millis(61_200)), "1分1.2秒");
    }

    // 把用户级设置文件（settings.json）隔离到测试专用目录。隔离目录
    // 在进程内只设定一次；所有触发偏好持久化的测试及修改环境变量的
    // 测试通过共享锁串行化，避免写入其他测试刚删除的目录。
    thread_local! {
        static AUDIT_STATE_ENV_LOCK_HELD: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    }

    struct AuditStateEnvTestGuard {
        guard: Option<std::sync::MutexGuard<'static, ()>>,
    }

    impl Drop for AuditStateEnvTestGuard {
        fn drop(&mut self) {
            if self.guard.is_some() {
                AUDIT_STATE_ENV_LOCK_HELD.with(|held| held.set(false));
            }
        }
    }

    fn audit_state_env_guard() -> AuditStateEnvTestGuard {
        if AUDIT_STATE_ENV_LOCK_HELD.with(std::cell::Cell::get) {
            return AuditStateEnvTestGuard { guard: None };
        }
        let guard = crate::app_settings::AUDIT_STATE_ENV_LOCK
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        AUDIT_STATE_ENV_LOCK_HELD.with(|held| held.set(true));
        AuditStateEnvTestGuard { guard: Some(guard) }
    }

    /// Settings access uses process-global AUDIT_STATE_DIR. Serialize test
    /// dispatches that persist preferences or initialize an engine from the
    /// settings file, not only tests that explicitly replace the environment.
    fn dispatch_api(
        state: &AppState,
        method: &str,
        args: &[Value],
    ) -> std::result::Result<Value, super::ApiError> {
        let persists_preferences = matches!(
            method,
            "choose_central"
                | "choose_history_config"
                | "choose_period_compare"
                | "start_dag_flow"
                | "import_central_explanation_feedback"
                | "start_central_indicator_split"
        ) || (method == "update"
            && args
                .first()
                .and_then(Value::as_object)
                .is_some_and(|values| {
                    values.contains_key("advancedSettings")
                        || values
                            .keys()
                            .any(|key| crate::app_settings::preference_spec(key).is_some())
                }));
        let _settings_env_guard = persists_preferences.then(audit_state_env_guard);
        super::dispatch_api(state, method, args)
    }

    pub(crate) fn isolate_settings_state_dir() {
        static ISOLATED: std::sync::Once = std::sync::Once::new();
        // 首个测试状态创建时才设定隔离目录；已自行设置 AUDIT_STATE_DIR 的
        // 测试（如高级设置持久化测试）拥有自己的目录，本助手绝不覆盖——
        // 否则该测试写入自己的 state_dir 而 AppSettings::store 落到别处，
        // 持久化断言随机失败。并发写安全由 app_settings 的唯一 tmp 文件名
        // 与原子改名保证。
        ISOLATED.call_once(|| {
            if std::env::var_os("AUDIT_STATE_DIR").is_some() {
                return;
            }
            let nonce = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|value| value.as_nanos())
                .unwrap_or_default();
            let path = std::env::temp_dir().join(format!(
                "audit-web-bridge-settings-{}-{nonce}",
                std::process::id()
            ));
            std::env::set_var("AUDIT_STATE_DIR", &path);
        });
    }

    fn test_state() -> AppState {
        let _settings_env_guard = audit_state_env_guard();
        isolate_settings_state_dir();
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

    #[test]
    fn s1_f1_template_preflight_cache_reuses_one_immutable_plan_per_path() -> anyhow::Result<()> {
        let directory = S1RecoveryProbeDirectory::create()?;
        let template = directory.0.join("template.xlsx");
        write_s1_recovery_fixture(&template, None)?;
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
        let directory = S1RecoveryProbeDirectory::create()?;
        let input = directory.0.join("input");
        let templates = directory.0.join("templates");
        let output = directory.0.join("output");
        fs::create_dir_all(&input)?;
        fs::create_dir_all(&templates)?;
        let template = templates.join("template.xlsx");
        write_s1_recovery_fixture(&template, None)?;
        let mut source_hashes = Vec::new();
        for index in 1..=3 {
            let source = input.join(format!("机构甲_2026-07_{index}.xlsx"));
            write_s1_recovery_fixture(&source, None)?;
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

    struct S1RecoveryProbeDirectory(PathBuf);

    impl S1RecoveryProbeDirectory {
        fn create() -> anyhow::Result<Self> {
            let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
            let path = std::env::temp_dir().join(format!("audit_s1_recovery_real_{suffix}"));
            fs::create_dir_all(&path)?;
            Ok(Self(path))
        }
    }

    impl Drop for S1RecoveryProbeDirectory {
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
    fn conditional_extraction_failure_keeps_formula_findings_as_partial_result(
    ) -> anyhow::Result<()> {
        let suffix = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!("audit_s1_conditional_partial_{suffix}"));
        fs::create_dir_all(&root)?;
        let source = root.join("机构甲_部分审核.xlsx");
        let template = root.join("template.xlsx");
        let audit_copy = root.join("audit-copy.xlsx");
        write_s1_recovery_fixture(&source, None)?;
        write_s1_recovery_fixture(&template, None)?;
        write_s1_recovery_fixture(&audit_copy, Some("错误|公式核查|测试指标|测试异常"))?;
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
            "source",
            &S1FormulaMetadata {
                period: "2026-07".to_owned(),
                batch_id: "conditional-partial-test".to_owned(),
                audit_time: "2026-10-29 12:00:00".to_owned(),
                org_code: "1001".to_owned(),
                org_name: "测试机构".to_owned(),
            },
            |workbook_path, source_path, _, _, _| {
                assert_eq!(workbook_path, source.as_path());
                assert_eq!(source_path, source.as_path());
                assert_ne!(workbook_path, audit_copy.as_path());
                anyhow::bail!("合成条件格式局部读取失败")
            },
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
    fn conditional_findings_link_to_audit_copy_when_read_from_source() {
        let source = Path::new("source.xlsx");
        let audit_copy = Path::new("audit-copy.xlsx");
        let mut issue = recovery_result_issue(source, source);

        super::set_conditional_issues_audit_copy(std::slice::from_mut(&mut issue), audit_copy);

        assert_eq!(issue.source_file, source.display().to_string());
        assert_eq!(issue.audit_file, audit_copy.display().to_string());
    }

    #[test]
    fn s1_conditional_format_value_source_selects_source_or_audit_copy() {
        let source = Path::new("source.xlsx");
        let audit_copy = Path::new("audit-copy.xlsx");
        assert_eq!(
            super::s1_conditional_format_workbook(source, audit_copy, "source"),
            source
        );
        assert_eq!(
            super::s1_conditional_format_workbook(source, audit_copy, "audit_copy"),
            audit_copy
        );
        assert_eq!(
            super::s1_conditional_format_workbook(source, audit_copy, "unknown"),
            source
        );
    }

    /// 生成 S1-F01 测试样本（含校验区域/条件格式区域/表结构区域三个命名区域）。
    fn write_s1_recovery_fixture(path: &Path, cached_result: Option<&str>) -> anyhow::Result<()> {
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
        write_s1_recovery_fixture(&calculated_source, None)?;
        write_s1_recovery_fixture(&pending_source, None)?;
        write_s1_recovery_fixture(&template, None)?;

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
            conditional_format_source: "audit_copy".to_owned(),
            request_preparation_us: 0,
            source_directory_scan_us: 0,
        };
        let store = S1RecoveryStore::new(state_dir);
        let mut current_work_stages = Vec::new();
        // 本测试验证「有 sidecar / 无 sidecar」的混合语义；betteroffice 是
        // 纯 Rust 独立引擎，会绕开该前提，故显式关闭以保留原测试契约。
        let _env_guard = crate::app_settings::AUDIT_STATE_ENV_LOCK
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let previous_bo = std::env::var_os("PYTOOLS_BO_CALC");
        std::env::set_var("PYTOOLS_BO_CALC", "off");
        let output = execute_s1_f1_workflow_with_progress(
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
                        write_s1_recovery_fixture(
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
                                excel_open_repair: None,
                            },
                        ))
                    }
                    waiting @ S1FormulaCopyOutcome::WaitingManualRecalc(_) => Ok(waiting),
                    S1FormulaCopyOutcome::Recalculated(_) => {
                        anyhow::bail!("没有 sidecar 的合成路径意外报告了重算成功")
                    }
                }
            },
            |item, index, stage| {
                current_work_stages.push((
                    item.map(|item| {
                        item.source
                            .file_name()
                            .unwrap()
                            .to_string_lossy()
                            .into_owned()
                    }),
                    item.and_then(s1_f1_current_sheet),
                    index,
                    stage.to_owned(),
                ));
            },
        )?;

        assert!(current_work_stages
            .iter()
            .any(|(workbook, sheet, index, stage)| {
                workbook.as_deref() == Some("机构甲_2026-07_自动重算.xlsx")
                    && sheet.is_some()
                    && *index == Some(1)
                    && stage == "复制公式并重算"
            }));
        assert!(current_work_stages
            .iter()
            .any(|(_, _, _, stage)| stage == "验证公式重算结果"));
        assert!(current_work_stages
            .iter()
            .any(|(_, _, _, stage)| stage == "提取公式与条件格式异常"));
        assert!(current_work_stages
            .iter()
            .any(|(_, _, _, stage)| stage == "汇总并写入审核结果"));
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
        assert_eq!(manifest.conditional_format_source, "audit_copy");
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

        match previous_bo {
            Some(value) => std::env::set_var("PYTOOLS_BO_CALC", value),
            None => std::env::remove_var("PYTOOLS_BO_CALC"),
        }
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
            .replace("__AUDIT_APP_TITLE__", super::app_title())
            .replace("__AUDIT_ICON_URL__", "/assets/app-icon.png");
        assert_eq!(served_markup, expected_markup);
        assert!(!served.contains("__AUDIT_APP_TITLE__"));
        assert!(!served.contains("__AUDIT_ICON_URL__"));
        assert!(served_markup.contains("centralCrossCur:'central_data'"));
        assert!(served_markup.contains("本期数值核对文件（CSV/XLSX/XLS）"));
        assert!(!served_markup.contains("XLS/XLSX 尚未接通"));
        // S1-F3/F4/F5 已按 AGENTS.md 标记为 deferred 并从界面移除：
        // 按钮节点不再输出，恢复入口的 JS 保留（有 null 保护，不再可达）。
        assert!(!served_markup.contains("id=\"s1F3RecoveryButton\""));
        assert!(!served_markup.contains("id=\"dagExtrasCard\""));
        assert!(!served_markup.contains("data-tab=\"combine\""));
        // S1-F3 恢复功能已整体删除（多重兜底下不再需要）：界面与 API 均不可达。
        assert!(!served.contains("s1F3LostCopyTasks"));
        assert!(!served.contains("start_s1_f3_workflow"));
        assert!(!served.contains("list_s1_f3_tasks"));
        assert!(served.contains("window.bridge = { api: api }"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_comparison()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_explanation_export()"));
        assert!(super::ORIGINAL_HTML
            .contains("bridge.api.import_central_explanation_feedback(picked.paths)"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_indicator_split()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_cross_period()"));
        assert!(super::ORIGINAL_HTML.contains("onclick=\"runCentralForms()\""));
        assert!(super::ORIGINAL_HTML.contains("centralFormsConfig:'config'"));
        // 金融表单输出目录已并入输出目录：独立目录字段与其回显行不再输出。
        assert!(!super::ORIGINAL_HTML.contains("最近输出目录："));
        assert!(!super::ORIGINAL_HTML.contains("centralFormOutput"));
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

    #[test]
    fn advanced_settings_update_persists_and_logs_selection() {
        let _settings_env_guard = audit_state_env_guard();
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
        let previous_state_dir = std::env::var("AUDIT_STATE_DIR").ok();
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

        let updated = dispatch_api(
            &state,
            "update",
            &[json!({"s1ConditionalFormatSource": "audit_copy"})],
        )
        .expect("valid conditional-format source must persist");
        assert_eq!(updated["s1ConditionalFormatSource"], "audit_copy");
        let persisted: Value =
            serde_json::from_slice(&fs::read(&settings_file).expect("read updated settings file"))
                .expect("valid updated settings JSON");
        assert_eq!(
            persisted["preferences"]["s1ConditionalFormatSource"],
            "audit_copy"
        );
        assert_eq!(
            crate::app_settings::AppSettings::load().s1_conditional_format_source(),
            "audit_copy"
        );

        // The next state build reflects the persisted selection (worker threads
        // read the file, not the GUI state).
        assert_eq!(
            crate::app_settings::AppSettings::load()
                .calc_engine
                .as_str(),
            "wps"
        );

        std::env::remove_var("AUDIT_STATE_DIR");
        if let Some(previous) = previous_state_dir {
            std::env::set_var("AUDIT_STATE_DIR", previous);
        }
        fs::remove_dir_all(&state_dir).expect("remove isolated settings directory");
    }

    #[test]
    fn extra_files_append_remove_and_persist() {
        let _settings_env_guard = audit_state_env_guard();
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let state_dir = std::env::temp_dir().join(format!(
            "audit_web_extra_files_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&state_dir).expect("create isolated settings directory");
        let previous_state_dir = std::env::var("AUDIT_STATE_DIR").ok();
        std::env::set_var("AUDIT_STATE_DIR", &state_dir);

        let work = std::env::temp_dir().join(format!(
            "audit_web_extra_files_work_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&work).expect("create work directory");
        let workbook = work.join("追加工作簿.xlsx");
        fs::write(&workbook, b"placeholder").expect("write workbook placeholder");
        let missing = work.join("不存在.xlsx");

        let state = test_state();
        let appended = dispatch_api(
            &state,
            "append_extra_files",
            &[json!([
                workbook.display().to_string(),
                workbook.display().to_string(),
                missing.display().to_string()
            ])],
        )
        .expect("append_extra_files must succeed");
        let listed = appended["extraFiles"]
            .as_array()
            .expect("extraFiles must be an array")
            .iter()
            .filter_map(Value::as_str)
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(listed.len(), 1, "重复路径与不存在的文件都不得入列");
        assert!(
            listed[0].replace('/', "\\").ends_with("追加工作簿.xlsx"),
            "实际入列：{listed:?}"
        );

        // 持久化：重新加载设置能看到追加结果。
        let settings = crate::app_settings::AppSettings::load();
        assert_eq!(settings.extra_files().len(), 1);

        let removed = dispatch_api(&state, "remove_extra_file", &[json!(listed[0])])
            .expect("remove_extra_file must succeed");
        assert_eq!(
            removed["extraFiles"].as_array().map(Vec::len),
            Some(0),
            "移除后清单应为空"
        );
        assert!(crate::app_settings::AppSettings::load()
            .extra_files()
            .is_empty());

        std::env::remove_var("AUDIT_STATE_DIR");
        if let Some(previous) = previous_state_dir {
            std::env::set_var("AUDIT_STATE_DIR", previous);
        }
        let _ = fs::remove_dir_all(&state_dir);
        let _ = fs::remove_dir_all(&work);
    }

    #[test]
    fn advanced_settings_rejects_unknown_engines_and_allows_openpyxl_trial_selection() {
        let _settings_env_guard = audit_state_env_guard();
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
        let previous_state_dir = std::env::var("AUDIT_STATE_DIR").ok();
        std::env::set_var("AUDIT_STATE_DIR", &state_dir);

        let state = test_state();
        let engine_options = state.lock().value["advancedSettings"]["editEngineOptions"]
            .as_array()
            .expect("GUI exposes edit engine choices")
            .iter()
            .filter_map(|option| option["value"].as_str())
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(engine_options, ["auto", "ooxml", "openpyxl", "umya"]);
        let invalid = dispatch_api(
            &state,
            "update",
            &[json!({"advancedSettings": {"calcEngine": "pywin32"}})],
        )
        .expect_err("unknown calc engine must be rejected");
        assert_eq!(invalid.status, StatusCode::BAD_REQUEST);
        assert!(invalid.message.contains("auto、excel、wps 或 libreoffice"));

        for engine in ["auto", "ooxml", "umya"] {
            let updated = dispatch_api(
                &state,
                "update",
                &[json!({"advancedSettings": {"editEngine": engine}})],
            )
            .expect("each non-worker GUI engine choice is accepted");
            assert_eq!(updated["advancedSettings"]["editEngine"], engine);
        }

        let openpyxl_status = crate::app_settings::openpyxl_selection_status();
        let openpyxl_update = dispatch_api(
            &state,
            "update",
            &[json!({"advancedSettings": {"editEngine": "openpyxl"}})],
        );
        if openpyxl_status.0 {
            let openpyxl_update = openpyxl_update
                .expect("openpyxl trial is selectable when its runtime is available");
            assert_eq!(
                openpyxl_update["advancedSettings"]["editEngine"],
                "openpyxl"
            );
            assert_eq!(
                crate::app_settings::AppSettings::load()
                    .edit_engine
                    .as_str(),
                "openpyxl"
            );
        } else {
            let error = openpyxl_update.expect_err("missing worker runtime must reject selection");
            assert_eq!(error.status, StatusCode::BAD_REQUEST);
            assert!(error.message.contains(&openpyxl_status.1));
        }

        std::env::remove_var("AUDIT_STATE_DIR");
        if let Some(previous) = previous_state_dir {
            std::env::set_var("AUDIT_STATE_DIR", previous);
        }
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
        let _settings_env_guard = audit_state_env_guard();
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
                    entry["text"].as_str().is_some_and(|text| {
                        text.contains("S2 阶段开始：")
                            && text.contains(&format!("stage={stage}"))
                            && text.contains(&format!("engine={engine}"))
                    })
                }),
                "missing start detail for S2 stage {stage}"
            );
            assert!(
                logs.iter().any(|entry| {
                    entry["text"].as_str().is_some_and(|text| {
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
    fn recursive_depth_update_validates_range_and_persists() {
        let state = test_state();
        let saved = dispatch_api(&state, "update", &[json!({"recursiveDepth": 2})])
            .expect("valid depth must be accepted");
        assert_eq!(saved["recursiveDepth"], 2);
        assert_eq!(
            crate::app_settings::AppSettings::load()
                .preferences
                .get("recursiveDepth")
                .and_then(Value::as_i64),
            Some(2),
            "depth must persist to the user settings file"
        );

        let rejected = dispatch_api(&state, "update", &[json!({"recursiveDepth": 11})]);
        let Err(error) = rejected else {
            panic!("depth 11 must be rejected");
        };
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(
            error.message.contains("-1 到 10"),
            "error names the valid range"
        );
        assert_eq!(
            state.get_state()["recursiveDepth"],
            2,
            "rejected input must not change state"
        );
    }

    #[test]
    fn state_reports_s2_paths_and_s3_enabled_by_default() {
        let state = test_state();
        let value = state.get_state();
        assert_eq!(value["distributionProfile"], "s1s2s3");
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
        let _settings_env_guard = audit_state_env_guard();
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
            value["centralOutput"],
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
        // AUDIT_STATE_DIR 是进程级变量：与本测试并行运行的持久化测试可能
        // 切换甚至删除该目录，串行化以避免写入已删除路径。
        let _settings_env_guard = audit_state_env_guard();
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
