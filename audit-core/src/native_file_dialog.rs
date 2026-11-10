//! 系统原生文件/目录选择器（可选能力，页面内目录浏览器为固定兜底）。
//!
//! 平台后端：
//! - Windows：`rfd` 同步 `FileDialog`（IFileOpenDialog，COM），在
//!   [`tokio::task::spawn_blocking`] 中执行，不阻塞 Axum/Tokio 工作线程。
//!   rfd 的 Windows 后端只依赖系统自带 COM，不引入新的运行时库。
//! - Linux：`ashpd` 的 XDG Desktop Portal FileChooser（zbus 纯 Rust D-Bus
//!   实现，运行时才连接会话总线；程序启动与二进制加载不依赖 libdbus 或
//!   GTK）。桌面门户缺失/无会话总线时返回可读诊断。
//!
//! 结果契约（[`PickOutcome`]）：
//! - `canceled = true`：用户在原生对话框中取消，调用方直接结束，不再弹页面
//!   内选择器；
//! - `canceled = false, paths 非空`：选择成功，调用方仍须走既有
//!   `choose_file`/`choose_folder`/`choose_central` 等后端校验；
//! - `canceled = false, paths 为空`：原生能力不可用/初始化失败/超时，
//!   `message` 携带可读诊断，调用方回退页面内选择器。
//!
//! 已知限制（如实记录）：rfd 的 Windows 后端把“对话框初始化失败”与“用户
//! 取消”都报告为 `None`，本模块按取消处理；Windows 上 COM 对话框基本总是
//! 可用，失败极罕见，最坏表现是多弹一次页面内选择器，不产生数据影响。

use serde_json::{json, Value};
use std::path::{Path, PathBuf};

/// 原生选择器整体安全超时。对话框由用户逐步操作，正常使用远低于该值；
/// 超时按“不可用”返回诊断并回退，Windows 上已弹出的对话框线程会保留到
/// 用户关闭为止（无法安全终止阻塞中的 COM 调用）。
const PICK_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(600);

/// 一次原生选择的请求参数。`extensions` 复用页面内选择器的
/// `browse_extensions` 模式映射（带点小写，如 `.xlsx`；空表示不过滤）。
#[derive(Debug, Clone)]
pub struct PickRequest {
    pub title: String,
    pub kind: PickKind,
    pub extensions: &'static [&'static str],
    pub start: Option<PathBuf>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PickKind {
    File,
    Directory,
    MultiFile,
}

#[derive(Debug)]
pub struct PickOutcome {
    pub canceled: bool,
    pub paths: Vec<PathBuf>,
    pub backend: &'static str,
    pub message: Option<String>,
}

impl PickOutcome {
    fn picked(backend: &'static str, paths: Vec<PathBuf>) -> Self {
        Self {
            canceled: false,
            paths,
            backend,
            message: None,
        }
    }
    fn canceled(backend: &'static str) -> Self {
        Self {
            canceled: true,
            paths: Vec::new(),
            backend,
            message: None,
        }
    }
    fn unavailable(backend: &'static str, message: String) -> Self {
        Self {
            canceled: false,
            paths: Vec::new(),
            backend,
            message: Some(message),
        }
    }
}

impl PickRequest {
    /// 过滤器显示名，同时用于原生对话框与诊断信息。
    pub fn filter_label(&self) -> String {
        if self.extensions.is_empty() {
            return "所有文件".to_owned();
        }
        format!("支持的文件类型（{}）", self.extensions.join("、"))
    }
}

/// 扩展名对应的 MIME 类型（带点或不带点都接受）。
///
/// XDG Desktop Portal 允许过滤器同时写 glob 与 MIME；部分国产化实现
/// （麒麟 V10 的 portal 组件）只认 MIME，glob 会在内部转换失败后被丢弃，
/// 导致对话框里出现"标签写着支持 csv、文件却是灰的"。这里为常用类型补上
/// MIME，glob 与 MIME 并存，两类后端都能命中。
fn extension_mime_type(extension: &str) -> Option<&'static str> {
    let normalized = extension.trim_start_matches('.').to_ascii_lowercase();
    Some(match normalized.as_str() {
        "csv" => "text/csv",
        "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
        "xls" => "application/vnd.ms-excel",
        "zip" => "application/zip",
        "deb" => "application/vnd.debian.binary-package",
        "gz" => "application/gzip",
        "tgz" => "application/gzip",
        "txt" => "text/plain",
        "json" => "application/json",
        _ => return None,
    })
}

/// 按当前平台调用原生选择器；统一超时与结果契约。
pub async fn pick(request: PickRequest) -> PickOutcome {
    let future = platform_pick(request);
    match tokio::time::timeout(PICK_TIMEOUT, future).await {
        Ok(outcome) => outcome,
        Err(_) => PickOutcome::unavailable(
            "native",
            "打开系统文件选择器超时；已交回页面内选择器。".to_owned(),
        ),
    }
}

/// 将结果序列化为 `pick_paths` API 的 JSON 载荷。
pub fn outcome_to_value(outcome: &PickOutcome) -> Value {
    json!({
        "canceled": outcome.canceled,
        "paths": outcome.paths.iter().map(|path| path.to_string_lossy()).collect::<Vec<_>>(),
        "backend": outcome.backend,
        "message": outcome.message,
    })
}

/// 把请求的起始路径归一化为“存在的目录”：空/不存在 → None（系统默认位置），
/// 文件 → 其父目录。保证 `.current_folder`/`set_directory` 不会因无效路径失败。
pub fn normalize_start(start: Option<&str>) -> (Option<PathBuf>, Option<String>) {
    let Some(raw) = start.map(str::trim).filter(|value| !value.is_empty()) else {
        return (None, None);
    };
    let candidate = PathBuf::from(raw);
    if candidate.is_dir() {
        return (Some(candidate), None);
    }
    if candidate.is_file() {
        return (candidate.parent().map(Path::to_path_buf), None);
    }
    (
        None,
        Some(format!("起始路径不存在，已使用系统默认位置：{raw}")),
    )
}

// ---------------------------------------------------------------------------
// Windows：rfd 同步对话框 + spawn_blocking
// ---------------------------------------------------------------------------
#[cfg(windows)]
async fn platform_pick(request: PickRequest) -> PickOutcome {
    let backend = "windows-com";
    let label = request.filter_label();
    let joined = tokio::task::spawn_blocking(move || blocking_pick(request, label)).await;
    match joined {
        Ok(outcome) => outcome,
        Err(error) => PickOutcome::unavailable(backend, format!("原生选择器任务执行失败：{error}")),
    }
}

#[cfg(windows)]
fn blocking_pick(request: PickRequest, label: String) -> PickOutcome {
    let backend = "windows-com";
    let mut dialog = rfd::FileDialog::new().set_title(request.title);
    if !request.extensions.is_empty() {
        let names: Vec<String> = request
            .extensions
            .iter()
            .map(|extension| extension.trim_start_matches('.').to_owned())
            .collect();
        dialog = dialog.add_filter(label, &names);
    }
    if let Some(start) = &request.start {
        dialog = dialog.set_directory(start);
    }
    let selection = match request.kind {
        PickKind::Directory => dialog.pick_folder().map(|path| vec![path]),
        PickKind::File => dialog.pick_file().map(|path| vec![path]),
        PickKind::MultiFile => dialog.pick_files(),
    };
    match selection {
        Some(paths) if !paths.is_empty() => PickOutcome::picked(backend, paths),
        _ => PickOutcome::canceled(backend),
    }
}

// ---------------------------------------------------------------------------
// Linux：XDG Desktop Portal FileChooser（ashpd/zbus，纯 Rust D-Bus）
// ---------------------------------------------------------------------------
#[cfg(target_os = "linux")]
async fn platform_pick(request: PickRequest) -> PickOutcome {
    portal_pick(request).await
}

#[cfg(target_os = "linux")]
async fn portal_pick(request: PickRequest) -> PickOutcome {
    use ashpd::desktop::file_chooser::{FileFilter, SelectedFiles};
    use ashpd::desktop::ResponseError;
    use ashpd::Error as PortalError;

    const BACKEND: &str = "xdg-portal";
    let mut builder = SelectedFiles::open_file()
        .title(request.title.as_str())
        .modal(false)
        .multiple(matches!(request.kind, PickKind::MultiFile))
        .directory(matches!(request.kind, PickKind::Directory));
    if !request.extensions.is_empty() {
        let mut filter = FileFilter::new(request.filter_label().as_str());
        for extension in &request.extensions {
            // 同时给大小写两种 glob：部分后端按字节比较，*.CSV 不会命中 *.csv。
            filter = filter.glob(format!("*{extension}").as_str());
            filter = filter.glob(format!("*{}", extension.to_uppercase()).as_str());
            // 同时给 MIME 类型：麒麟等国产化 terminal 的 xdg-desktop-portal
            // 实现会把 glob 转成 MIME，转换失败时**静默丢弃该条规则**——
            // 表现为筛选器标签里写着 .csv、但 CSV 文件在对话框里是灰的
            // （2026-11 现场问题）。补上 MIME 后两端都能命中。
            if let Some(mime) = extension_mime_type(extension) {
                filter = filter.mimetype(mime);
            }
        }
        builder = builder.filter(filter);
        // 兜底：再给一个"所有文件"过滤器。若某个桌面环境的 portal 仍然
        // 丢规则，用户可以手动切换筛选器选到任意文件，不至于被卡死。
        builder = builder.filter(FileFilter::new("所有文件").glob("*"));
    }
    if let Some(start) = &request.start {
        builder = match builder.current_folder(start) {
            Ok(builder) => builder,
            Err(error) => return PickOutcome::unavailable(BACKEND, portal_diagnostic(&error)),
        };
    }
    let pending = match builder.send().await {
        Ok(pending) => pending,
        Err(PortalError::Response(ResponseError::Cancelled)) => {
            return PickOutcome::canceled(BACKEND)
        }
        Err(error) => return PickOutcome::unavailable(BACKEND, portal_diagnostic(&error)),
    };
    let selected = match pending.response() {
        Ok(selected) => selected,
        Err(PortalError::Response(ResponseError::Cancelled)) => {
            return PickOutcome::canceled(BACKEND)
        }
        Err(error) => return PickOutcome::unavailable(BACKEND, portal_diagnostic(&error)),
    };
    let mut paths = Vec::with_capacity(selected.uris().len());
    for uri in selected.uris() {
        match file_uri_to_path(uri.as_str()) {
            Some(path) => paths.push(path),
            None => {
                return PickOutcome::unavailable(
                    BACKEND,
                    format!("系统文件选择器返回了无法解析的地址：{}", uri.as_str()),
                )
            }
        }
    }
    PickOutcome::picked(BACKEND, paths)
}

/// 门户错误转可读诊断；不把内部错误细节当作用户失败。
#[cfg(target_os = "linux")]
fn portal_diagnostic(error: &ashpd::Error) -> String {
    use ashpd::Error as PortalError;
    match error {
        PortalError::Zbus(inner) => {
            format!("无法连接桌面门户（D-Bus 会话）：{inner}")
        }
        PortalError::PortalNotFound(interface) => {
            format!("当前桌面环境未提供文件选择门户服务：{interface}")
        }
        PortalError::Portal(inner) => format!("桌面门户拒绝了文件选择请求：{inner}"),
        other => format!("系统文件选择器调用失败：{other}"),
    }
}

/// 把门户返回的 `file:///带%20空格` URI 解码为本地路径。
#[cfg(target_os = "linux")]
fn file_uri_to_path(uri: &str) -> Option<PathBuf> {
    let rest = uri.strip_prefix("file://")?;
    // 个别实现可能返回 file://localhost/path；去掉主机名段。
    let rest = match rest.strip_prefix("localhost") {
        Some(after_host) => after_host,
        None => rest,
    };
    let decoded = percent_decode(rest)?;
    if decoded.is_empty() {
        return None;
    }
    Some(PathBuf::from(decoded))
}

#[cfg(target_os = "linux")]
fn percent_decode(input: &str) -> Option<String> {
    let bytes = input.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' && index + 2 < bytes.len() {
            let high = (bytes[index + 1] as char).to_digit(16)?;
            let low = (bytes[index + 2] as char).to_digit(16)?;
            decoded.push(((high * 16) + low) as u8);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(decoded).ok()
}

// ---------------------------------------------------------------------------
// 其他平台：仅保留可编译占位（产品目标只有 Windows 与 x86_64 Linux）。
// ---------------------------------------------------------------------------
#[cfg(not(any(windows, target_os = "linux")))]
async fn platform_pick(request: PickRequest) -> PickOutcome {
    let _ = request;
    PickOutcome::unavailable("none", "当前平台未实现原生文件选择器。".to_owned())
}

#[cfg(test)]
mod tests {
    use super::{extension_mime_type, PickRequest};

    /// 过滤器标签必须如实列出全部支持的扩展名（它会显示在系统对话框里）。
    #[test]
    fn filter_label_lists_every_supported_extension() {
        let request = PickRequest {
            title: "选择本期数据文件".to_owned(),
            kind: super::PickKind::File,
            extensions: &[".csv", ".xlsx", ".xls"],
            start: None,
        };
        assert_eq!(
            request.filter_label(),
            "支持的文件类型（.csv、.xlsx、.xls）"
        );
        let empty = PickRequest {
            extensions: &[],
            ..request
        };
        assert_eq!(empty.filter_label(), "所有文件");
    }

    /// 关键扩展名都要有 MIME：麒麟 portal 只认 MIME 时会用它命中文档。
    #[test]
    fn spreadsheet_and_csv_extensions_have_mime_types() {
        assert_eq!(extension_mime_type(".csv"), Some("text/csv"));
        assert_eq!(extension_mime_type("csv"), Some("text/csv"));
        assert_eq!(extension_mime_type(".CSV"), Some("text/csv"));
        assert_eq!(
            extension_mime_type(".xlsx"),
            Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        );
        assert_eq!(
            extension_mime_type(".xlsm"),
            Some("application/vnd.ms-excel.sheet.macroEnabled.12")
        );
        assert_eq!(
            extension_mime_type(".xls"),
            Some("application/vnd.ms-excel")
        );
        // 未知扩展名不返回 MIME（只靠 glob 匹配）。
        assert_eq!(extension_mime_type(".unknown-ext"), None);
    }
}
