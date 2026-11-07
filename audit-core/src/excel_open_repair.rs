//! Excel COM 打开失败的审核副本修复。
//!
//! 背景（依据 129 份真实核查表根因分析，与旧 Python Audit 的
//! `core/src/base_audit/excel_open_repair.py` 逐条一致）：
//!
//! 1. 部分报送文件的公式直接引用含全角括号的表名（如
//!    `个人贷款发生额信息（隐藏发放）!A1`）而没有加单引号。Excel 要求含特殊
//!    字符的表名必须加引号，存储的公式解析不了 → 拒绝打开整个工作簿；
//!    WPS 解析宽松能打开。
//! 2. 部分文件的冻结窗格 `<pane ySplit="6" topLeftCell="C6" state="frozen"/>`
//!    自相矛盾（topLeftCell 的行落在冻结区内），Excel 强制要求冻结态下
//!    topLeftCell 位于分割线右下方 → 拒绝打开。
//!
//! 修复只做两类无害规范化：给强制引号表名的裸引用补单引号（公式语义不变）；
//! 冻结窗格 topLeftCell 行号改为 ySplit+1。**只允许作用于审核副本**，调用方
//! 必须保证传入路径不是原始报送文件；本模块不做来源检查之外的文件写保护，
//! 但绝不主动打开或改写用户原始数据。
//!
//! 无问题的文件不做任何重写（先廉价预检）；异常一律吞掉并返回未修复，
//! 不打断审核流程——WPS 引擎本来就能打开这些文件，修复只是 Excel 引擎的
//! 前置保障。

use std::fs;
use std::io::{Read as _, Write as _};
use std::path::Path;
use std::sync::LazyLock;

use anyhow::{Context, Result};
use regex::{Captures, Regex};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// 修复统计；`changed=false` 时 `reason` 说明原因，文件保持原样。
#[derive(Clone, Debug, Default, PartialEq, Eq, serde::Serialize)]
pub struct ExcelOpenRepairReport {
    pub changed: bool,
    pub quoted: usize,
    pub pane_fixed: usize,
    pub reason: String,
}

/// 需要改写的包部件：工作表/图表/绘图 XML 与工作簿定义名称。
const XML_PART_PREFIXES: [&str; 3] = ["xl/worksheets/", "xl/charts/", "xl/drawings/"];

static F_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)(<f(?:\s[^>]*)?>)(.*?)(</f>)").unwrap());
static FORMULA_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)(<formula(?:\s[^>]*)?>)(.*?)(</formula>)").unwrap());
static CHART_F_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"(<c:f>)(.*?)(</c:f>)").unwrap());
static DEFINED_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?s)(<definedNames>)(.*?)(</definedNames>)").unwrap());
static PANE_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"<pane[^>]*/>").unwrap());
static PANE_YSPLIT_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r#"ySplit="(\d+)""#).unwrap());
static PANE_TOPLEFT_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"topLeftCell="([A-Z]+)(\d+)""#).unwrap());

/// Excel 强制引号字符：括号（全/半角）、空格（含全角）、连字符等。
/// 纯中文/字母数字表名不需要引号，不在此列。
fn sheet_name_needs_quotes(name: &str) -> bool {
    name.chars().any(|c| {
        c.is_whitespace()
            || matches!(
                c,
                '（' | '）'
                    | '('
                    | ')'
                    | '-'
                    | '+'
                    | '['
                    | ']'
                    | '{'
                    | '}'
                    | '\\'
                    | '/'
                    | ':'
                    | ';'
                    | '"'
                    | '\''
                    | '<'
                    | '>'
                    | ','
                    | '?'
                    | '~'
                    | '！'
                    | '？'
                    | '＠'
                    | '＃'
                    | '￥'
                    | '％'
                    | '…'
                    | '&'
                    | '*'
                    | '='
                    | '|'
            )
    })
}

/// 仅当本次 S1 重算会把工作簿交给 Excel COM 打开时才需要修复。
/// `auto` 在 Windows 上以 Excel 优先；WPS/LibreOffice 解析宽松，
/// 本来就能打开这些文件，无需修复。
pub fn excel_com_engine_requested(engine: &str) -> bool {
    cfg!(windows) && matches!(engine, "auto" | "excel")
}

/// 从 `xl/workbook.xml` 读取需要强制引号的表名（与 ElementTree 一样解析
/// XML 实体）。工作簿部件缺失或解析失败时返回空列表，不影响窗格检查。
fn special_sheet_names(archive: &mut ZipArchive<fs::File>) -> Result<Vec<String>> {
    let Ok(mut entry) = archive.by_name("xl/workbook.xml") else {
        return Ok(Vec::new());
    };
    let mut data = Vec::new();
    if entry.read_to_end(&mut data).is_err() {
        return Ok(Vec::new());
    }
    let Ok(text) = std::str::from_utf8(&data) else {
        return Ok(Vec::new());
    };
    let Ok(document) = roxmltree::Document::parse(text) else {
        return Ok(Vec::new());
    };
    Ok(document
        .descendants()
        .filter(|node| node.tag_name().name() == "sheet")
        .filter_map(|node| node.attribute("name"))
        .filter(|name| !name.is_empty() && sheet_name_needs_quotes(name))
        .map(str::to_owned)
        .collect())
}

fn part_is_repairable(name: &str) -> bool {
    name.ends_with(".xml")
        && (XML_PART_PREFIXES
            .iter()
            .any(|prefix| name.starts_with(prefix))
            || name == "xl/workbook.xml")
}

/// 匹配 `(?<!')NAME!` 的裸引用：前一个字符不是单引号。
fn bare_ref_occurrences(text: &str, name: &str) -> Vec<(usize, usize)> {
    let needle = format!("{name}!");
    let mut ranges = Vec::new();
    let mut search_from = 0usize;
    while let Some(found) = text[search_from..].find(&needle) {
        let start = search_from + found;
        let end = start + needle.len();
        if text[..start].ends_with('\'') {
            search_from = end;
        } else {
            ranges.push((start, end));
            search_from = end;
        }
    }
    ranges
}

fn has_bare_ref(text: &str, special: &[String]) -> bool {
    special
        .iter()
        .any(|name| !bare_ref_occurrences(text, name).is_empty())
}

/// 给强制引号表名的裸引用补单引号；已引号引用与原实现一样跳过。
fn quote_bare_refs(text: &str, special: &[String]) -> (String, usize) {
    let mut result = text.to_owned();
    let mut total = 0usize;
    for name in special {
        let replacement = format!("'{}'!", name.replace('\'', "''"));
        for (start, end) in bare_ref_occurrences(&result, name).into_iter().rev() {
            result.replace_range(start..end, &replacement);
            total += 1;
        }
    }
    (result, total)
}

/// 冻结窗格坐标矛盾：返回需要替换的 `topLeftCell="..."` 属性原文与修正值。
fn contradictory_pane_attribute(tag: &str) -> Option<(String, String)> {
    let frozen = tag.contains("state=\"frozen\"") || tag.contains("state=\"frozenSplit\"");
    if !frozen {
        return None;
    }
    let ysplit = PANE_YSPLIT_RE.captures(tag)?;
    let topleft = PANE_TOPLEFT_RE.captures(tag)?;
    let row: usize = topleft[2].parse().ok()?;
    let split: usize = ysplit[1].parse().ok()?;
    if row > split {
        return None;
    }
    let original = topleft.get(0)?.as_str().to_owned();
    let corrected = format!("topLeftCell=\"{}{}\"", &topleft[1], split + 1);
    Some((original, corrected))
}

/// 处理一个可修复部件；文本不是有效 UTF-8 时保持原始字节不动。
fn repair_part_bytes(
    data: &[u8],
    special: &[String],
    quoted: &mut usize,
    pane_fixed: &mut usize,
) -> Vec<u8> {
    let Ok(text) = std::str::from_utf8(data) else {
        return data.to_vec();
    };
    let mut current = text.to_owned();
    if !special.is_empty() {
        let replaced = F_RE.replace_all(&current, |caps: &Captures| {
            let (segment, count) = quote_bare_refs(&caps[2], special);
            *quoted += count;
            format!("{}{}{}", &caps[1], segment, &caps[3])
        });
        current = replaced.into_owned();
        let replaced = FORMULA_RE.replace_all(&current, |caps: &Captures| {
            let (segment, count) = quote_bare_refs(&caps[2], special);
            *quoted += count;
            format!("{}{}{}", &caps[1], segment, &caps[3])
        });
        current = replaced.into_owned();
        let replaced = CHART_F_RE.replace_all(&current, |caps: &Captures| {
            let (segment, count) = quote_bare_refs(&caps[2], special);
            *quoted += count;
            format!("{}{}{}", &caps[1], segment, &caps[3])
        });
        current = replaced.into_owned();
        let replaced = DEFINED_RE.replace_all(&current, |caps: &Captures| {
            let (segment, count) = quote_bare_refs(&caps[2], special);
            *quoted += count;
            format!("{}{}{}", &caps[1], segment, &caps[3])
        });
        current = replaced.into_owned();
    }
    let replaced =
        PANE_RE.replace_all(
            &current,
            |caps: &Captures| match contradictory_pane_attribute(&caps[0]) {
                Some((original, corrected)) => {
                    *pane_fixed += 1;
                    caps[0].replace(&original, &corrected)
                }
                None => caps[0].to_owned(),
            },
        );
    replaced.into_owned().into_bytes()
}

fn needs_repair(archive: &mut ZipArchive<fs::File>, special: &[String]) -> Result<bool> {
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        if !part_is_repairable(entry.name()) {
            continue;
        }
        let mut data = Vec::new();
        if entry.read_to_end(&mut data).is_err() {
            continue;
        }
        let text = String::from_utf8_lossy(&data);
        if has_bare_ref(&text, special) {
            return Ok(true);
        }
        if entry.name().starts_with("xl/worksheets/")
            && text.contains("<pane ")
            && PANE_RE
                .find_iter(&text)
                .any(|pane| contradictory_pane_attribute(pane.as_str()).is_some())
        {
            return Ok(true);
        }
    }
    Ok(false)
}

/// 修复审核副本使其可被 Excel 打开；返回统计。
///
/// 只处理 OOXML（ZIP 容器）工作簿；其余格式原样跳过。任何异常都返回
/// `changed=false` 且保持文件原样（先写临时文件再原子替换）。
pub fn repair_audit_copy_for_excel(path: &Path) -> ExcelOpenRepairReport {
    let mut report = ExcelOpenRepairReport::default();
    let outcome = repair_audit_copy_inner(path, &mut report);
    if let Err(error) = outcome {
        report.changed = false;
        report.reason = format!("{error:#}");
    }
    report
}

fn repair_audit_copy_inner(path: &Path, report: &mut ExcelOpenRepairReport) -> Result<()> {
    let temporary = excel_repair_temporary_path(path);
    let result = (|| -> Result<()> {
        let file =
            fs::File::open(path).with_context(|| format!("打开审核副本：{}", path.display()))?;
        let mut archive =
            ZipArchive::new(file).with_context(|| format!("解析 ZIP 容器：{}", path.display()))?;
        let names: Vec<String> = (0..archive.len())
            .filter_map(|index| {
                archive
                    .by_index(index)
                    .ok()
                    .map(|entry| entry.name().to_owned())
            })
            .collect();
        if !names.iter().any(|name| name == "xl/workbook.xml") {
            report.reason = "非 OOXML 工作簿".to_owned();
            return Ok(());
        }
        let special = special_sheet_names(&mut archive)?;
        if !needs_repair(&mut archive, &special)? {
            report.reason = "未发现需要修复的内容".to_owned();
            return Ok(());
        }
        let out_file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)?;
        let mut writer = ZipWriter::new(out_file);
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let name = entry.name().to_owned();
            let mut options = SimpleFileOptions::default().compression_method(entry.compression());
            if let Some(modified) = entry.last_modified() {
                options = options.last_modified_time(modified);
            }
            if entry.is_dir() {
                writer.add_directory(name, options)?;
                continue;
            }
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            let data = if part_is_repairable(&name) {
                repair_part_bytes(&data, &special, &mut report.quoted, &mut report.pane_fixed)
            } else {
                data
            };
            writer.start_file(name, options)?;
            writer.write_all(&data)?;
        }
        writer.finish()?;
        drop(archive);
        fs::rename(&temporary, path)
            .with_context(|| format!("原子替换修复副本：{}", path.display()))?;
        report.changed = true;
        Ok(())
    })();
    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}

fn excel_repair_temporary_path(path: &Path) -> std::path::PathBuf {
    let file_name = path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_default();
    let stem = Path::new(&file_name)
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_else(|| file_name.clone());
    let suffix = Path::new(&file_name)
        .extension()
        .map(|extension| format!(".{}", extension.to_string_lossy()))
        .unwrap_or_default();
    path.with_file_name(format!("{stem}.excel-repair.tmp{suffix}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicU32, Ordering};

    fn temp_directory(label: &str) -> std::path::PathBuf {
        static SEQUENCE: AtomicU32 = AtomicU32::new(0);
        let directory = std::env::temp_dir().join(format!(
            "excel-open-repair-test-{}-{}-{}",
            label,
            std::process::id(),
            SEQUENCE.fetch_add(1, Ordering::Relaxed)
        ));
        let _ = fs::remove_dir_all(&directory);
        fs::create_dir_all(&directory).unwrap();
        directory
    }

    fn write_workbook(path: &Path, parts: &[(&str, &str)]) {
        let file = fs::File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        for (name, content) in parts {
            writer
                .start_file(
                    (*name).to_owned(),
                    SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
                )
                .unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    fn read_part(path: &Path, name: &str) -> String {
        let file = fs::File::open(path).unwrap();
        let mut archive = ZipArchive::new(file).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut text = String::new();
        entry.read_to_string(&mut text).unwrap();
        text
    }

    fn workbook_xml(sheet_names: &[&str]) -> String {
        let sheets = sheet_names
            .iter()
            .enumerate()
            .map(|(index, name)| {
                format!(
                    r#"<sheet name="{name}" sheetId="{}" r:id="rId{}"/>"#,
                    index + 1,
                    index + 1
                )
            })
            .collect::<Vec<_>>()
            .join("");
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#
        )
    }

    const SHEET: &str = "个人贷款发生额信息（隐藏发放）";
    const PLAIN_SHEET: &str = "存量个人贷款信息";

    #[test]
    fn quotes_bare_fullwidth_sheet_references_everywhere() {
        let directory = temp_directory("quotes_bare");
        let path = directory.join("copy.xlsx");
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[SHEET])),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<worksheet><sheetData><row r="1"><c r="A1"><f>个人贷款发生额信息（隐藏发放）!A1+1</f><v>2</v></c></row></sheetData>
<conditionalFormatting sqref="B2"><cfRule type="expression"><formula>个人贷款发生额信息（隐藏发放）!$A$1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#,
                ),
                (
                    "xl/charts/chart1.xml",
                    r#"<chartSpace><c:f>个人贷款发生额信息（隐藏发放）!$B$2:$B$9</c:f></chartSpace>"#,
                ),
            ],
        );
        let report = repair_audit_copy_for_excel(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 3, "{report:?}");
        assert_eq!(report.pane_fixed, 0);
        let worksheet = read_part(&path, "xl/worksheets/sheet1.xml");
        assert!(
            worksheet.contains(&format!("'{SHEET}'!A1+1")),
            "{worksheet}"
        );
        assert!(
            worksheet.contains(&format!("'{SHEET}'!$A$1&gt;0")),
            "{worksheet}"
        );
        let chart = read_part(&path, "xl/charts/chart1.xml");
        assert!(chart.contains(&format!("'{SHEET}'!$B$2:$B$9")), "{chart}");
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn defined_names_block_is_quoted_and_already_quoted_refs_are_kept() {
        let directory = temp_directory("defined_names");
        let path = directory.join("copy.xlsx");
        let workbook = workbook_xml(&[SHEET]).replace(
            "</sheets>",
            &format!(
                "</sheets><definedNames><definedName name=\"X\">{SHEET}!$A$1:'{SHEET}'!$A$2</definedName></definedNames>"
            ),
        );
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", workbook.as_str()),
                ("xl/worksheets/sheet1.xml", ""),
            ],
        );
        let report = repair_audit_copy_for_excel(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 1, "{report:?}");
        let text = read_part(&path, "xl/workbook.xml");
        assert!(
            text.contains(&format!(
                "<definedName name=\"X\">'{SHEET}'!$A$1:'{SHEET}'!$A$2</definedName>"
            )),
            "{text}"
        );
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn clean_files_are_not_rewritten() {
        let directory = temp_directory("clean");
        let path = directory.join("copy.xlsx");
        let worksheet =
            format!(r#"<worksheet><f>'{SHEET}'!A1</f><f>{PLAIN_SHEET}!B2</f></worksheet>"#);
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[SHEET, PLAIN_SHEET])),
                ("xl/worksheets/sheet1.xml", &worksheet),
            ],
        );
        let before = fs::read(&path).unwrap();
        let report = repair_audit_copy_for_excel(&path);
        assert!(!report.changed, "{report:?}");
        assert_eq!(report.reason, "未发现需要修复的内容");
        assert_eq!(fs::read(&path).unwrap(), before, "无问题文件不允许任何重写");
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn fixes_contradictory_frozen_pane_only() {
        let directory = temp_directory("pane");
        let path = directory.join("copy.xlsx");
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[PLAIN_SHEET])),
                (
                    "xl/worksheets/sheet2.xml",
                    r#"<worksheet><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="6" topLeftCell="C6" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="I9"/></sheetView></sheetViews></worksheet>"#,
                ),
            ],
        );
        let report = repair_audit_copy_for_excel(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 0);
        assert_eq!(report.pane_fixed, 1);
        let text = read_part(&path, "xl/worksheets/sheet2.xml");
        assert!(text.contains("topLeftCell=\"C7\""), "{text}");
        assert!(text.contains("ySplit=\"6\""), "{text}");
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn compliant_and_unfrozen_panes_are_kept() {
        let directory = temp_directory("pane_ok");
        let path = directory.join("copy.xlsx");
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[PLAIN_SHEET])),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<worksheet><sheetViews><sheetView workbookViewId="0"><pane ySplit="6" topLeftCell="C7" state="frozen"/></sheetView></sheetViews></worksheet>"#,
                ),
                (
                    "xl/worksheets/sheet2.xml",
                    r#"<worksheet><sheetViews><sheetView workbookViewId="0"><pane ySplit="6" topLeftCell="C6" state="split"/></sheetView></sheetViews></worksheet>"#,
                ),
            ],
        );
        let before = fs::read(&path).unwrap();
        let report = repair_audit_copy_for_excel(&path);
        assert!(!report.changed, "{report:?}");
        assert_eq!(fs::read(&path).unwrap(), before);
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn apostrophes_in_names_are_doubled_inside_quotes() {
        let directory = temp_directory("apostrophe");
        let path = directory.join("copy.xlsx");
        let name = "O'Brien（1）";
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[name])),
                ("xl/worksheets/sheet1.xml", &format!("<f>{name}!A1</f>")),
            ],
        );
        let report = repair_audit_copy_for_excel(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 1);
        let text = read_part(&path, "xl/worksheets/sheet1.xml");
        assert!(text.contains("<f>'O''Brien（1）'!A1</f>"), "{text}");
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn non_ooxml_files_are_reported_unchanged() {
        let directory = temp_directory("non_ooxml");
        let path = directory.join("legacy.xlsx");
        fs::write(&path, b"not a zip container").unwrap();
        let report = repair_audit_copy_for_excel(&path);
        assert!(!report.changed, "{report:?}");
        assert!(!report.reason.is_empty(), "{report:?}");
        assert_eq!(fs::read(&path).unwrap(), b"not a zip container");
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn quotes_and_pane_fixes_combine() {
        let directory = temp_directory("combined");
        let path = directory.join("copy.xlsx");
        write_workbook(
            &path,
            &[
                ("xl/workbook.xml", &workbook_xml(&[SHEET])),
                (
                    "xl/worksheets/sheet1.xml",
                    format!(
                        r#"<worksheet><f>{SHEET}!A1</f><pane ySplit="6" topLeftCell="C6" state="frozen"/></worksheet>"#
                    )
                    .as_str(),
                ),
            ],
        );
        let report = repair_audit_copy_for_excel(&path);
        assert!(report.changed, "{report:?}");
        assert_eq!(report.quoted, 1);
        assert_eq!(report.pane_fixed, 1);
        let _ = fs::remove_dir_all(&directory);
    }

    #[test]
    fn excel_com_gate_only_on_windows_auto_or_excel() {
        if cfg!(windows) {
            assert!(excel_com_engine_requested("auto"));
            assert!(excel_com_engine_requested("excel"));
        } else {
            assert!(!excel_com_engine_requested("auto"));
            assert!(!excel_com_engine_requested("excel"));
        }
        assert!(!excel_com_engine_requested("wps"));
        assert!(!excel_com_engine_requested("lo"));
    }

    #[test]
    fn needs_quotes_matches_reference_character_set() {
        assert!(sheet_name_needs_quotes("个人贷款发生额信息（隐藏发放）"));
        assert!(sheet_name_needs_quotes("票据融资跨期差额（隐藏直贴买入）"));
        assert!(sheet_name_needs_quotes(
            "单位贷款置换旧债极端数据核查信息（元）"
        ));
        assert!(sheet_name_needs_quotes("O'Brien"));
        assert!(sheet_name_needs_quotes("Sheet 1"));
        assert!(sheet_name_needs_quotes("盈亏&汇总"));
        assert!(!sheet_name_needs_quotes("存量个人贷款信息"));
        assert!(!sheet_name_needs_quotes("Sheet1"));
        assert!(!sheet_name_needs_quotes(""));
    }

    #[test]
    fn bare_ref_detection_skips_already_quoted_names() {
        assert_eq!(
            bare_ref_occurrences(&format!("{SHEET}!A1"), SHEET),
            vec![(0, SHEET.len() + 1)]
        );
        assert!(bare_ref_occurrences(&format!("'{SHEET}'!A1"), SHEET).is_empty());
        assert!(has_bare_ref(&format!("{SHEET}!A1"), &[SHEET.to_owned()]));
        assert!(!has_bare_ref(&format!("'{SHEET}'!A1"), &[SHEET.to_owned()]));
    }
}
