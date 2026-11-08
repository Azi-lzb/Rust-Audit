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
    /// 外部工作簿引用前缀（索引式/路径式，四种形态）剥离次数。
    pub external_ref_fixed: usize,
    /// 剥离后确认无残留时移除的外链相关部件数（externalLink 等）。
    pub external_link_parts_removed: usize,
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

/// 本地全部工作表名（与 special 检测同一来源，含普通表名）。
fn all_local_sheet_names(archive: &mut ZipArchive<fs::File>) -> Result<Vec<String>> {
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
        .map(str::to_owned)
        .collect())
}

fn has_bare_ref(text: &str, special: &[String]) -> bool {
    special
        .iter()
        .any(|name| !bare_ref_occurrences(text, name).is_empty())
}

/// 给强制引号表名的裸引用补单引号；已引号引用与原实现一样跳过。
/// 表名的"编码无关"正则片段：每个非 ASCII 字符同时匹配字面、十进制 NCR 与
/// 十六进制 NCR。openpyxl 保存时把中文写成 `&#38598;` 实体，WPS/Excel 生成
/// 的文件是字面 UTF-8；两种都必须命中，否则经 openpyxl 处理过的副本会漏修。
fn name_flex_regex(name: &str) -> String {
    let mut parts = String::new();
    for ch in name.chars() {
        if (ch as u32) < 128 {
            parts.push_str(&regex::escape(&ch.to_string()));
        } else {
            let code = ch as u32;
            parts.push_str(&format!(
                "(?:{}|&#{code};|&#[xX]{:x};|&#[xX]{:X};)",
                regex::escape(&ch.to_string()),
                code,
                code
            ));
        }
    }
    parts
}

/// 剥离公式里的外部工作簿引用前缀（业务口径：副本只允许工作簿内引用）。
///
/// 覆盖四种形态（与 Python `_strip_external_prefixes` 逐条一致）：
/// ① `'C:\路径\[簿.xlsx]表名'!` ② `'[簿.xlsx]表名'!`
/// ③ `[簿.xlsx]表名!` ④ `[1]表名!`（或 `[1]'表名'!`）
///
/// 索引式与路径式**一律无条件剥离**（不看本地是否已有同名表）：本地有同名表
/// 则公式改指本地表；本地没有则公式自然成 `#REF!`，按审核异常计入——这优于
/// 整簿被 Excel 拒绝打开。`local_sheets` 仅用于日志判定缺失表名。
///
/// 返回 (新文本, 剥离次数)。
fn strip_external_refs(text: &str, local_sheets: &[String], stripped: &mut usize) -> String {
    let mut result = text.to_owned();
    let mut total = 0usize;

    // ① ② 路径式 + 引号：'...\[簿.xlsx]表名'! → '表名'!
    if let Ok(re) = Regex::new(r"'[^']*\[[^\]/]+\.(?:xlsx|xlsm|xlsb|xls)\]((?:[^']|'')*)'!") {
        let replaced = re.replace_all(&result, |caps: &Captures| format!("'{}'!", &caps[1]));
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // ③ 路径式 + 裸名：[簿.xlsx]表名! → 表名!
    if let Ok(re) = Regex::new(r"\[[^\]/]+\.(?:xlsx|xlsm|xlsb|xls)\]([^\s!'\[\]<>()]+)!") {
        let replaced = re.replace_all(&result, |caps: &Captures| format!("{}!", &caps[1]));
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // ④ 索引式 + 引号：[1]'表名'! → '表名'!
    if let Ok(re) = Regex::new(r"\[\d+\]'((?:[^']|'')*)'!") {
        let replaced = re.replace_all(&result, |caps: &Captures| format!("'{}'!", &caps[1]));
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // 索引式 + NCR 编码表名：[1]&#38598;...! → 表名!
    for name in local_sheets {
        let flex = name_flex_regex(name);
        let escaped = name_flex_regex(&name.replace('\'', "''"));
        if let Ok(re) = Regex::new(&format!(r"\[\d+\]'{escaped}'!")) {
            let replacement = format!("'{}'!", name.replace('\'', "''"));
            let replaced = re.replace_all(&result, replacement.as_str());
            if replaced != result {
                total += re.find_iter(&result).count();
                result = replaced.into_owned();
            }
        }
        if let Ok(re) = Regex::new(&format!(r"\[\d+\]{flex}!")) {
            let replacement = format!("{name}!");
            let replaced = re.replace_all(&result, replacement.as_str());
            if replaced != result {
                total += re.find_iter(&result).count();
                result = replaced.into_owned();
            }
        }
    }

    // ④ 索引式 + 裸名：[1]表名! → 表名!
    if let Ok(re) = Regex::new(r"\[\d+\]([^\s!'\[\]<>()]+)!") {
        let replaced = re.replace_all(&result, |caps: &Captures| format!("{}!", &caps[1]));
        if replaced != result {
            total += re.find_iter(&result).count();
            result = replaced.into_owned();
        }
    }
    // 剥离后若仍有 NCR 实体形态的表名（openpyxl 保存的文件），解码为字面
    // UTF-8，使公式与 WPS/Excel 生成的形态一致。
    if let Ok(re) = Regex::new(r"&#(?:[xX]([0-9a-fA-F]+)|([0-9]+));") {
        let decoded = re.replace_all(&result, |caps: &Captures| {
            let code = caps
                .get(1)
                .and_then(|hex| u32::from_str_radix(hex.as_str(), 16).ok())
                .or_else(|| caps.get(2).and_then(|dec| dec.as_str().parse::<u32>().ok()));
            code.and_then(char::from_u32)
                .map(|ch| ch.to_string())
                .unwrap_or_else(|| caps[0].to_owned())
        });
        result = decoded.into_owned();
    }
    *stripped += total;
    result
}

/// 是否仍存在任何形态的外部引用（索引式或路径式）。用于决定 externalLink
/// 定义能否安全清除——**检测与剥离必须同源**，只查 `[N]` 会漏判路径式，
/// 误删外链定义后原本存活的引用变成新的悬空引用。
fn has_any_external_ref(text: &str) -> bool {
    if Regex::new(r"\[\d+\]")
        .ok()
        .is_some_and(|re| re.is_match(text))
    {
        return true;
    }
    Regex::new(r"\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls)\]")
        .ok()
        .is_some_and(|re| re.is_match(text))
}

/// 剥离后仍指向"本地不存在表"的引用名（用于运行日志提示，不阻断）。
/// 只识别合法表引用：`'表名'!` 或中文/字母/数字/下划线/空格组成的 `表名!`。
fn dangling_after_strip(formula_text: &str, local_sheets: &[String]) -> Vec<String> {
    let mut missing: Vec<String> = Vec::new();
    let Ok(re) = Regex::new(
        r"(?:'((?:[^']|'')*)'|([A-Za-z_\u{4e00}-\u{9fff}][\w\u{4e00}-\u{9fff} .\u{3000}]*))!",
    ) else {
        return missing;
    };
    for caps in re.captures_iter(formula_text) {
        let name = caps
            .get(1)
            .or_else(|| caps.get(2))
            .map(|m| m.as_str().replace("''", "'"))
            .unwrap_or_default();
        let name = name.trim();
        if name.is_empty() || local_sheets.iter().any(|sheet| sheet == name) {
            continue;
        }
        // 排除单元格地址（如 A1、AB12）
        if Regex::new(r"^[A-Za-z]{1,3}\d+$")
            .ok()
            .is_some_and(|re| re.is_match(name))
        {
            continue;
        }
        if !missing.iter().any(|existing| existing == name) {
            missing.push(name.to_owned());
        }
    }
    missing
}

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
    local_sheet_names: &[String],
    quoted: &mut usize,
    external_ref_fixed: &mut usize,
    pane_fixed: &mut usize,
) -> Vec<u8> {
    let Ok(text) = std::str::from_utf8(data) else {
        return data.to_vec();
    };
    let mut current = text.to_owned();
    // 第三类缺陷：导入外部表后公式残留 `[N]表名!` 外部工作簿引用前缀，
    // 目标表已在本地 → 剥离 `[N]` 改为本地引用（悬空外链使 Excel 拒开）。
    if !local_sheet_names.is_empty() {
        let replaced = strip_external_refs(&current, local_sheet_names, external_ref_fixed);
        current = replaced;
    }
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

fn needs_repair(
    archive: &mut ZipArchive<fs::File>,
    special: &[String],
    local_sheet_names: &[String],
) -> Result<bool> {
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
        if has_any_external_ref(&text) {
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
        let local_sheet_names = all_local_sheet_names(&mut archive)?;
        if !needs_repair(&mut archive, &special, &local_sheet_names)? {
            report.reason = "未发现需要修复的内容".to_owned();
            return Ok(());
        }
        let out_file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)?;
        let mut writer = ZipWriter::new(out_file);
        // 预扫描：在不写出的情况下模拟剥离，确定剥离后是否仍有任何外部引用
        // 残留（索引式或路径式）。必须先于写出完成，否则 workbook.xml 可能
        // 先于工作表被写出，导致外链声明残留。检测与剥离同源。
        let mut residual_external_ref = false;
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            if !part_is_repairable(entry.name()) {
                continue;
            }
            let mut probe = Vec::new();
            if entry.read_to_end(&mut probe).is_err() {
                continue;
            }
            let mut scratch_total = 0usize;
            let stripped = repair_part_bytes(
                &probe,
                &special,
                &local_sheet_names,
                &mut 0,
                &mut scratch_total,
                &mut 0,
            );
            if has_any_external_ref(&String::from_utf8_lossy(&stripped)) {
                if std::env::var_os("AUDIT_DEBUG_RESIDUAL").is_some() {
                    let text = String::from_utf8_lossy(&stripped);
                    let sample: String = text.chars().take(400).collect();
                    eprintln!("RESIDUAL part={} sample={sample}", entry.name());
                }
                residual_external_ref = true;
                break;
            }
        }
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
            if !residual_external_ref && name.starts_with("xl/externalLinks/") {
                continue;
            }
            let data = if part_is_repairable(&name) {
                let repaired = repair_part_bytes(
                    &data,
                    &special,
                    &local_sheet_names,
                    &mut report.quoted,
                    &mut report.external_ref_fixed,
                    &mut report.pane_fixed,
                );
                if has_any_external_ref(&String::from_utf8_lossy(&repaired)) {
                    residual_external_ref = true;
                }
                // workbook.xml 既是可修复部件（definedNames 补引号），也要在
                // 确认无残留时清除 <externalReferences> 声明。
                if !residual_external_ref && name == "xl/workbook.xml" {
                    strip_external_link_declarations(&name, &repaired)
                } else {
                    repaired
                }
            } else if !residual_external_ref
                && matches!(
                    name.as_str(),
                    "xl/_rels/workbook.xml.rels" | "[Content_Types].xml"
                )
            {
                strip_external_link_declarations(&name, &data)
            } else {
                data
            };
            writer.start_file(name, options)?;
            writer.write_all(&data)?;
        }
        // 第二遍：确认全包无任何外部引用后，才断开外链定义（移除
        // externalReferences / externalLink 关系 / Content_Types 覆盖 / 部件）。
        // 仍有残留时保持原样，交由重算报 #REF! 或人工处理。
        if !residual_external_ref {
            report.external_link_parts_removed = external_link_part_count(&mut archive)?;
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

/// 统计 externalLinks 部件数量（供报告）。
fn external_link_part_count(archive: &mut ZipArchive<fs::File>) -> Result<usize> {
    let mut count = 0usize;
    for index in 0..archive.len() {
        if let Ok(entry) = archive.by_index(index) {
            if entry.name().starts_with("xl/externalLinks/") {
                count += 1;
            }
        }
    }
    Ok(count)
}

/// 清除工作簿级外链声明：`<externalReferences>` 元素、workbook rels 中
/// Type 含 externalLink 的 Relationship、Content_Types 中 externalLinks 的 Override。
/// 仅在确认全包已无任何外部引用残留时调用（检测与剥离同源）。
fn strip_external_link_declarations(part: &str, data: &[u8]) -> Vec<u8> {
    let text = String::from_utf8_lossy(data).into_owned();
    let mut current = text;
    if part == "xl/workbook.xml" {
        if let Ok(re) = Regex::new(r"(?s)<externalReferences[^>]*>.*?</externalReferences>") {
            current = re.replace_all(&current, "").into_owned();
        }
        if let Ok(re) = Regex::new(r"<externalReferences[^>]*/>") {
            current = re.replace_all(&current, "").into_owned();
        }
    } else if part == "xl/_rels/workbook.xml.rels" {
        if let Ok(re) = Regex::new(r#"<Relationship[^>]*Type="[^"]*externalLink"[^>]*/>"#) {
            current = re.replace_all(&current, "").into_owned();
        }
    } else if part == "[Content_Types].xml" {
        if let Ok(re) = Regex::new(r#"<Override[^>]*PartName="/xl/externalLinks/[^"]*"[^>]*/>"#) {
            current = re.replace_all(&current, "").into_owned();
        }
    }
    current.into_bytes()
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
    #[test]
    fn strips_all_four_external_reference_forms() {
        let locals = vec!["集中系统数据".to_owned(), "参照表".to_owned()];
        // ① 路径式 + 引号 + 完整路径
        let mut n1 = 0usize;
        let t1 = strip_external_refs(
            r"INDEX('C:\外部文件夹\[工作簿.xlsx]集中系统数据'!$A:$A,1)",
            &locals,
            &mut n1,
        );
        assert_eq!(n1, 1, "{t1}");
        assert_eq!(t1, "INDEX('集中系统数据'!$A:$A,1)");

        // ② 路径式 + 引号 + 仅簿名
        let mut n2 = 0usize;
        let t2 = strip_external_refs("'[工作簿.xlsm]参照表'!$A$1", &locals, &mut n2);
        assert_eq!(n2, 1, "{t2}");
        assert_eq!(t2, "'参照表'!$A$1");

        // ③ 路径式 + 无引号
        let mut n3 = 0usize;
        let t3 = strip_external_refs("[工作簿.xlsb]集中系统数据!$A$1", &locals, &mut n3);
        assert_eq!(n3, 1, "{t3}");
        assert_eq!(t3, "集中系统数据!$A$1");

        // ④ 索引式（引号与裸名两种）
        let mut n4 = 0usize;
        let t4 = strip_external_refs("[1]集中系统数据!$A$1+[2]'参照表'!$B$2", &locals, &mut n4);
        assert_eq!(n4, 2, "{t4}");
        // 引号式引用保留引号（与 Python 替换口径一致：'表名'!）。
        assert_eq!(t4, "集中系统数据!$A$1+'参照表'!$B$2");
    }

    #[test]
    fn strips_external_refs_unconditionally_even_for_unknown_sheets() {
        // 业务口径：副本只允许工作簿内引用，因此任何外链前缀都剥离——
        // 本地没有同名表时公式自然成 #REF!，按审核异常计入。
        let locals = vec!["集中系统数据".to_owned()];
        let mut n = 0usize;
        let text = strip_external_refs("SUM([3]其他簿!A1)+[9]未知表!B2", &locals, &mut n);
        assert_eq!(n, 2, "{text}");
        assert_eq!(text, "SUM(其他簿!A1)+未知表!B2");
    }

    #[test]
    fn leaves_in_workbook_references_untouched() {
        let locals = vec!["存量债券投资".to_owned()];
        let mut n = 0usize;
        let original = "SUM(存量债券投资!A1:B9)+'集中系统数据'!C3";
        let text = strip_external_refs(original, &locals, &mut n);
        assert_eq!(n, 0, "{text}");
        assert_eq!(text, original);
    }

    #[test]
    fn strips_ncr_encoded_sheet_names_and_spaced_names() {
        // openpyxl 保存的中文表名会写成 NCR 实体；空格表名需引号包裹。
        let locals = vec!["集中系统数据".to_owned(), "某 表".to_owned()];
        let mut n = 0usize;
        let text = strip_external_refs(
            "INDEX([1]&#38598;&#20013;&#31995;&#32479;&#25968;&#25454;!$A:$A,1)+[2]'某 表'!B1",
            &locals,
            &mut n,
        );
        assert_eq!(n, 2, "{text}");
        assert!(text.contains("集中系统数据!$A:$A"), "{text}");
        assert!(text.contains("'某 表'!B1"), "{text}");
        assert!(!text.contains("&#38598;"), "{text}");
    }

    #[test]
    fn reports_missing_sheets_after_strip_without_blocking() {
        let locals = vec!["集中系统数据".to_owned()];
        let mut n = 0usize;
        let text = strip_external_refs("[1]集中系统数据!A1+[1]参照表!B1", &locals, &mut n);
        assert_eq!(n, 2, "{text}");
        // 剥离后公式为 `表名!`（无前缀）；缺失表名由日志侧收集。
        assert_eq!(text, "集中系统数据!A1+参照表!B1");
        let missing = dangling_after_strip(&text, &locals);
        assert_eq!(missing, vec!["参照表".to_owned()], "{missing:?}");
    }

    #[test]
    fn external_ref_detection_covers_index_and_path_forms() {
        // 检测与剥离同源：索引式与路径式都算残留。
        assert!(has_any_external_ref("[1]集中系统数据!A1"));
        assert!(has_any_external_ref("'[簿.xlsx]表'!A1"));
        assert!(has_any_external_ref("[簿.xlsx]表!A1"));
        assert!(!has_any_external_ref("集中系统数据!A1"));
        assert!(!has_any_external_ref("SUM(A1:B2)"));
    }

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
