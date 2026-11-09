//! Windows 可执行文件资源：应用图标与版本信息。
//!
//! Rust 默认不嵌入 `.ico`，Explorer/任务栏会显示通用图标。这里用 Windows SDK
//! 的 `rc.exe` 把 `frontend/assets/app-icon.ico` 与 `VERSION_MANIFEST.json` 的
//! 产品版本编译成 `.res`，再交给链接器。
//!
//! 约束：
//! - 仅 Windows 目标生效；Linux/其他目标直接跳过（构建机没有 rc.exe）。
//! - 找不到 rc.exe 时只提示一次并继续构建，不阻断交付（图标缺失不影响功能）。
//! - 清单路径与版本号读取失败时用 crate 的内置版本，保证资源段始终可编。

use std::path::{Path, PathBuf};
use std::process::Command;

fn main() {
    println!("cargo:rerun-if-changed=frontend/assets/app-icon.ico");
    println!("cargo:rerun-if-changed=../VERSION_MANIFEST.json");
    println!("cargo:rerun-if-changed=build.rs");

    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") {
        return;
    }
    let manifest_dir = PathBuf::from(std::env::var("CARGO_MANIFEST_DIR").unwrap_or_default());
    let icon = manifest_dir.join("frontend/assets/app-icon.ico");
    if !icon.is_file() {
        println!("cargo:warning=应用图标缺失：{}", icon.display());
        return;
    }

    let Some(rc) = locate_rc() else {
        println!("cargo:warning=未找到 Windows SDK rc.exe，跳过 EXE 图标与版本资源");
        return;
    };

    let version = read_product_version(&manifest_dir).unwrap_or_else(|| "0.0.0.0".to_owned());
    let (a, b, c, d) = version_tuple(&version).unwrap_or((0, 0, 0, 0));
    let out_dir = PathBuf::from(std::env::var("OUT_DIR").unwrap_or_default());

    let rc_path = out_dir.join("audittools_app.rc");
    let res_path = out_dir.join("audittools_app.res");
    // rc.exe 的字符串字面量需要转义反斜杠；统一用绝对路径避免受工作目录影响。
    let icon_literal = icon.to_string_lossy().replace('\\', "\\\\");
    let script = format!(
        r#"#pragma code_page(65001)
1 ICON "{icon_literal}"

1 VERSIONINFO
 FILEVERSION {a},{b},{c},{d}
 PRODUCTVERSION {a},{b},{c},{d}
 FILEFLAGSMASK 0x3fL
 FILEFLAGS 0x0L
 FILEOS 0x40004L
 FILETYPE 0x1L
 FILESUBTYPE 0x0L
BEGIN
  BLOCK "StringFileInfo"
  BEGIN
    BLOCK "080404b0"
    BEGIN
      VALUE "FileDescription", "审核工具（AuditTools）"
      VALUE "ProductName", "审核工具"
      VALUE "InternalName", "AuditTools"
      VALUE "OriginalFilename", "AuditTools.exe"
      VALUE "FileVersion", "{version}"
      VALUE "ProductVersion", "{version}"
    END
  END
  BLOCK "VarFileInfo"
  BEGIN
    VALUE "Translation", 0x804, 1200
  END
END
"#
    );
    if std::fs::write(&rc_path, script.as_bytes()).is_err() {
        println!("cargo:warning=写入临时 .rc 失败，跳过 EXE 图标");
        return;
    }

    let output = Command::new(&rc)
        .arg("/nologo")
        .arg("/c")
        .arg("65001")
        .arg(format!("/fo{}", res_path.display()))
        .arg(&rc_path)
        .output();
    match output {
        Ok(output) if output.status.success() && res_path.is_file() => {
            // 资源文件作为链接器输入（link.exe 按扩展名识别 .res）。
            println!("cargo:rustc-link-arg-bins={}", res_path.display());
        }
        Ok(output) => {
            let stderr = String::from_utf8_lossy(&output.stderr);
            println!(
                "cargo:warning=rc.exe 编译资源失败（{}），跳过 EXE 图标：{}",
                output.status,
                stderr.trim()
            );
        }
        Err(error) => println!("cargo:warning=无法运行 rc.exe，跳过 EXE 图标：{error}"),
    }
}

/// 从 `VERSION_MANIFEST.json` 读取 `product.version`（保持与运行时同源）。
fn read_product_version(manifest_dir: &Path) -> Option<String> {
    let text = std::fs::read_to_string(manifest_dir.join("../VERSION_MANIFEST.json")).ok()?;
    for line in text.lines() {
        let line = line.trim();
        if let Some(rest) = line.strip_prefix("\"version\":") {
            let value = rest.trim().trim_end_matches(',').trim().trim_matches('"');
            if version_tuple(value).is_some() {
                return Some(value.to_owned());
            }
        }
    }
    None
}

fn version_tuple(text: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<u32> = text
        .split('.')
        .map(|part| part.trim().parse::<u32>().ok())
        .collect::<Option<Vec<_>>>()?;
    match parts.as_slice() {
        [a, b, c, d] if *a <= 65535 && *b <= 65535 && *c <= 65535 && *d <= 65535 => {
            Some((*a, *b, *c, *d))
        }
        _ => None,
    }
}

/// 在标准 Windows SDK 安装位置里找最新的 x64 rc.exe。
fn locate_rc() -> Option<PathBuf> {
    let mut roots = vec![
        PathBuf::from(r"C:\Program Files (x86)\Windows Kits\10\bin"),
        PathBuf::from(r"C:\Program Files\Windows Kits\10\bin"),
    ];
    if let Some(sdk) = std::env::var_os("WindowsSdkDir") {
        roots.insert(0, PathBuf::from(sdk).join("bin"));
    }
    for root in roots {
        let Ok(entries) = std::fs::read_dir(&root) else {
            continue;
        };
        let mut versions: Vec<PathBuf> = entries
            .filter_map(Result::ok)
            .map(|entry| entry.path())
            .filter(|path| path.is_dir())
            .collect();
        // 目录名形如 10.0.26100.0，按版本倒序取最新的一个。
        versions.sort_by(|left, right| right.file_name().cmp(&left.file_name()));
        for version_dir in versions {
            let candidate = version_dir.join("x64/rc.exe");
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}
