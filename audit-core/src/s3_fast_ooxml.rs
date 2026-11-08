//! S3-F05 FAST_OOXML 渲染器——移植旧 audit `fast_ooxml.py` 语义。
//!
//! 与 EditPlan 逐格管线不同：模板**编译一次**（行结构/样式缓存），每机构
//! 纯字符串重建 sheetData + SheetPlan 删表/降级隐藏一次落盘 + 轻量完整性
//! 校验。对齐目标：旧 audit FAST_OOXML（实测 5.4s/8 表）的输出与速度。
use std::collections::{BTreeMap, HashMap, HashSet};

use anyhow::{bail, Context, Result};

const HEADER_DROP_MARKERS: [&str; 5] = ["剥离", "退出", "上期", "增减额", "环比"];

// ---------------------------------------------------------------------------
// 编译产物
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
pub struct CompiledRow {
    pub attrs: String,
    pub row_index: u32,
    /// 列号 → 原始 `<c>` XML（保留列内容原样复用）
    pub cells: BTreeMap<u32, String>,
    pub style_by_col: HashMap<u32, String>,
}

#[derive(Clone, Debug, Default)]
pub struct CompiledSheet {
    pub name: String,
    pub part: String,
    pub prefix: String,
    pub suffix: String,
    pub rows: Vec<CompiledRow>,
    pub base_headers: Vec<String>,
    pub keep_cols: Vec<u32>,
    pub biz_class: String,
    /// 行号 → [(列, 回填键)]；键与比较结果匹配由调用方完成
    pub indicator_rows: Vec<u32>,
    pub header_style: String,
    pub base_style_by_col: HashMap<u32, String>,
    pub new_dimension: String,
}

#[derive(Clone, Debug)]
pub struct SheetEntry {
    pub name: String,
    pub rid: String,
    pub part: String,
    pub state: String,
    pub sheet_id: String,
    pub old_index: usize,
}

#[derive(Clone, Debug, Default)]
pub struct CompiledTemplate {
    pub part_order: Vec<String>,
    pub parts: BTreeMap<String, Vec<u8>>,
    pub workbook_xml: String,
    pub workbook_rels_xml: String,
    pub content_types_xml: String,
    pub styles_xml: String,
    pub sheet_registry: Vec<SheetEntry>,
    pub internal_config_sheets: Vec<String>,
    pub sheets: BTreeMap<String, CompiledSheet>,
    pub shared_strings: Vec<String>,
}

// ---------------------------------------------------------------------------
// 模板编译
// ---------------------------------------------------------------------------

pub fn xml_attr(text: &str, name: &str) -> String {
    let pattern = format!(r#"{name}="([^"]*)""#);
    if let Ok(re) = regex::Regex::new(&pattern) {
        if let Some(caps) = re.captures(text) {
            return caps[1].to_owned();
        }
    }
    String::new()
}

pub fn col_number(cell_ref: &str) -> u32 {
    let mut number = 0u32;
    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            number = number * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else {
            break;
        }
    }
    number
}

pub fn col_letter(mut number: u32) -> String {
    let mut out = String::new();
    while number > 0 {
        let rem = ((number - 1) % 26) as u8;
        out.insert(0, (b'A' + rem) as char);
        number = (number - 1) / 26;
    }
    out
}

/// 单元格文本（t="s" 共享串 / inlineStr / 数值原样），实体统一解码。
fn cell_text(cell_xml: Option<&String>, shared: &[String]) -> String {
    let Some(cell_xml) = cell_xml else {
        return String::new();
    };
    let cell_type = xml_attr(cell_xml, "t");
    let value = regex::Regex::new(r"<v>(.*?)</v>")
        .ok()
        .and_then(|re| re.captures(cell_xml))
        .map(|caps| caps[1].to_owned());
    let text = match cell_type.as_str() {
        "s" => value
            .and_then(|v| v.parse::<usize>().ok())
            .and_then(|index| shared.get(index))
            .cloned()
            .unwrap_or_default(),
        "inlineStr" => regex::Regex::new(r"<t[^>]*>(.*?)</t>")
            .ok()
            .map(|re| {
                re.captures_iter(cell_xml)
                    .map(|caps| caps[1].to_owned())
                    .collect::<String>()
            })
            .unwrap_or_default(),
        _ => value.unwrap_or_default(),
    };
    decode_xml_entities(&text)
}

/// 覆盖数值实体与命名实体（与 Python html.unescape 对齐的常用子集）。
fn decode_xml_entities(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(pos) = rest.find('&') {
        output.push_str(&rest[..pos]);
        rest = &rest[pos..];
        if let Some(semi) = rest.find(';') {
            let entity = &rest[1..semi];
            let decoded = if let Some(hex) = entity
                .strip_prefix("#x")
                .or_else(|| entity.strip_prefix("#X"))
            {
                u32::from_str_radix(hex, 16).ok().and_then(char::from_u32)
            } else if let Some(dec) = entity.strip_prefix('#') {
                dec.parse::<u32>().ok().and_then(char::from_u32)
            } else {
                match entity {
                    "amp" => Some('&'),
                    "lt" => Some('<'),
                    "gt" => Some('>'),
                    "quot" => Some('"'),
                    "apos" => Some('\''),
                    _ => None,
                }
            };
            match decoded {
                Some(ch) => {
                    output.push(ch);
                    rest = &rest[semi + 1..];
                    continue;
                }
                None => {
                    output.push('&');
                    rest = &rest[1..];
                }
            }
        } else {
            output.push('&');
            rest = &rest[1..];
        }
    }
    output.push_str(rest);
    output
}

fn parse_rows(xml: &str) -> Vec<CompiledRow> {
    let row_re = regex::Regex::new(r"<row\b[^>]*(?:/>|>.*?</row>)").expect("row regex");
    let attrs_re = regex::Regex::new(r"<row\b([^>]*?)(?:/>|>)").expect("row attrs regex");
    let cell_re = regex::Regex::new(r"<c\b[^>]*/>|<c\b[^>]*>.*?</c>").expect("cell regex");
    let mut rows = Vec::new();
    for row_match in row_re.captures_iter(xml) {
        let row_xml = row_match[0].to_owned();
        let attrs = attrs_re
            .captures(&row_xml)
            .map(|caps| caps[1].to_owned())
            .unwrap_or_default();
        let row_index = xml_attr(&attrs, "r").parse().unwrap_or(0);
        let mut cells = BTreeMap::new();
        let mut style_by_col = HashMap::new();
        for cell_match in cell_re.captures_iter(&row_xml) {
            let cell_xml = cell_match[0].to_owned();
            let reference = xml_attr(&cell_xml, "r");
            let column = col_number(&reference);
            if column != 0 {
                style_by_col.insert(column, xml_attr(&cell_xml, "s"));
                cells.insert(column, cell_xml);
            }
        }
        rows.push(CompiledRow {
            attrs,
            row_index,
            cells,
            style_by_col,
        });
    }
    rows
}

fn inner_sheet_data(sheet_data: &str) -> String {
    let body = regex::Regex::new(r"^<sheetData\b[^>]*>")
        .expect("sheetData open regex")
        .replace(sheet_data, "");
    regex::Regex::new(r"</sheetData>$")
        .expect("sheetData close regex")
        .replace(&body, "")
        .into_owned()
}

fn sheet_biz_class(sheet_name: &str) -> String {
    // 与 Python _sheet_biz_class 对齐：表名前缀（如 A1411 → "A"）之后的业务
    // 分类取首字母段；旧实现按表名首段字母判定，这里保守取首个字母段。
    let letters: String = sheet_name
        .chars()
        .take_while(|ch| ch.is_ascii_alphabetic())
        .collect();
    letters.to_ascii_uppercase()
}

/// 编译模板（一次）：解析 workbook 结构、共享串、报表清单、各表单行结构。
pub fn compile_template(template_bytes: &[u8]) -> Result<CompiledTemplate> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(template_bytes))
        .context("打开金融表单模板 ZIP 失败")?;
    let part_order: Vec<String> = (0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|file| file.name().to_owned())
        })
        .collect();
    let mut parts = BTreeMap::new();
    for name in &part_order {
        let mut file = archive
            .by_name(name)
            .with_context(|| format!("读取模板部件 {name}"))?;
        let mut data = Vec::with_capacity(file.size() as usize);
        std::io::Read::read_to_end(&mut file, &mut data)?;
        parts.insert(name.clone(), data);
    }

    let text = |name: &str| -> String {
        parts
            .get(name)
            .map(|data| String::from_utf8_lossy(data).into_owned())
            .unwrap_or_default()
    };
    let workbook_xml = text("xl/workbook.xml");
    let workbook_rels_xml = text("xl/_rels/workbook.xml.rels");
    let content_types_xml = text("[Content_Types].xml");
    let styles_xml = text("xl/styles.xml");

    // 工作表注册表：workbook <sheet> + rels 的 rId→Target（namespace 无关）
    let sheet_re = regex::Regex::new(r#"<(?:[\w.]+:)?sheet\b[^>]*/>"#).expect("sheet regex");
    let rel_re = regex::Regex::new(r#"<(?:[\w.]+:)?Relationship\b[^>]*/>"#).expect("rel regex");
    let mut rid_to_target = HashMap::new();
    for caps in rel_re.captures_iter(&workbook_rels_xml) {
        let entry = &caps[0];
        let id = xml_attr(entry, "Id");
        let target = xml_attr(entry, "Target");
        if !id.is_empty() {
            rid_to_target.insert(id, target);
        }
    }
    let mut sheet_registry = Vec::new();
    for (position, caps) in sheet_re.captures_iter(&workbook_xml).enumerate() {
        let entry = &caps[0];
        let name = xml_attr(entry, "name");
        let rid = xml_attr(entry, "r:id");
        let raw_target = rid_to_target.get(&rid).cloned().unwrap_or_default();
        let part = normalize_part_target(&raw_target);
        sheet_registry.push(SheetEntry {
            name,
            rid,
            part,
            state: xml_attr(entry, "state"),
            sheet_id: xml_attr(entry, "sheetId"),
            old_index: position,
        });
    }

    let shared_strings = read_shared_strings(parts.get("xl/sharedStrings.xml"));

    let mut compiled = CompiledTemplate {
        part_order,
        parts,
        workbook_xml,
        workbook_rels_xml,
        content_types_xml,
        styles_xml,
        sheet_registry,
        internal_config_sheets: Vec::new(),
        sheets: Default::default(),
        shared_strings: shared_strings.clone(),
    };

    // 内嵌配置页（与 Python EMBEDDED_CONFIG_SHEETS 对齐）
    const EMBEDDED_CONFIG_SHEETS: [&str; 2] = ["使用说明", "转表设置"];
    let registry_names: HashSet<&str> = compiled
        .sheet_registry
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    if EMBEDDED_CONFIG_SHEETS
        .iter()
        .all(|name| registry_names.contains(name))
    {
        compiled.internal_config_sheets = EMBEDDED_CONFIG_SHEETS
            .iter()
            .map(|name| (*name).to_owned())
            .collect();
    }

    // 表单工作表编译（报表清单之外的 Axxxx/Axxxxx 表）。
    // 与 Python 对齐：只编译带表头行的表单表；无表头的配置/汇总表
    // （如 月报1批 空表）不参与回填，部件按原样保留。
    for entry in &compiled.sheet_registry {
        if entry.name == "报表清单" || compiled.internal_config_sheets.contains(&entry.name) {
            continue;
        }
        if let Some(sheet) =
            compile_sheet(&entry.name, &entry.part, &compiled.parts, &shared_strings)?
        {
            compiled.sheets.insert(entry.name.clone(), sheet);
        }
    }
    Ok(compiled)
}

fn normalize_part_target(target: &str) -> String {
    let trimmed = target.trim_start_matches('/');
    if trimmed.starts_with("xl/") {
        trimmed.to_owned()
    } else {
        format!("xl/{trimmed}")
    }
}

fn read_shared_strings(raw: Option<&Vec<u8>>) -> Vec<String> {
    let Some(raw) = raw else {
        return Vec::new();
    };
    let text = String::from_utf8_lossy(raw);
    let mut values = Vec::new();
    if let Ok(document) = roxmltree::Document::parse(&text) {
        for node in document.descendants() {
            if node.tag_name().name() == "si" {
                let content: String = node
                    .descendants()
                    .filter(|child| child.tag_name().name() == "t")
                    .map(|child| child.text().unwrap_or(""))
                    .collect();
                values.push(content);
            }
        }
    }
    values
}

fn compile_sheet(
    name: &str,
    part: &str,
    parts: &BTreeMap<String, Vec<u8>>,
    shared: &[String],
) -> Result<Option<CompiledSheet>> {
    let raw = parts
        .get(part)
        .with_context(|| format!("模板缺少工作表部件：{name}"))?;
    let xml = String::from_utf8_lossy(raw).into_owned();
    let data_re = regex::Regex::new(r"(?s)<sheetData\b[^>]*(?:/>|>.*?</sheetData>)")
        .expect("sheetData regex");
    let data_match = data_re
        .captures(&xml)
        .with_context(|| format!("工作表「{name}」缺少 sheetData 段"))?;
    let mut compiled = CompiledSheet {
        name: name.to_owned(),
        part: part.to_owned(),
        prefix: xml[..data_match.get(0).unwrap().start()].to_owned(),
        suffix: xml[data_match.get(0).unwrap().end()..].to_owned(),
        biz_class: sheet_biz_class(name),
        ..Default::default()
    };

    let mut base_headers: Vec<String> = Vec::new();
    let mut keep_cols: Vec<u32> = Vec::new();
    let mut header_style = String::new();
    let mut header_seen = false;
    for parsed in parse_rows(&inner_sheet_data(data_match.get(0).unwrap().as_str())) {
        let mut row = parsed;
        if !header_seen {
            let max_col = row.cells.keys().copied().max().unwrap_or(0);
            if max_col <= 3 {
                // 空占位表（月报/季报批次壳表）：无数据列不参与回填
                return Ok(None);
            }
            for column in 4..=max_col {
                let header = cell_text(row.cells.get(&column), shared);
                if column >= 5
                    && (header.is_empty()
                        || HEADER_DROP_MARKERS
                            .iter()
                            .any(|marker| header.contains(marker)))
                {
                    continue;
                }
                base_headers.push(header);
                keep_cols.push(column);
                if header_style.is_empty() {
                    header_style = row.style_by_col.get(&column).cloned().unwrap_or_default();
                }
            }
            if base_headers.is_empty() {
                return Ok(None); // 表头全部命中剔除标记：无可用数据列
            }
            header_seen = true;
            compiled.rows.push(row);
            continue;
        }
        // 数据行：去掉被清理列的单元格；有指标代码的行登记为回填行
        let keys: Vec<u32> = row.cells.keys().copied().collect();
        for column in keys {
            if column >= 4 && !keep_cols.contains(&column) {
                row.cells.remove(&column);
            }
        }
        let indicator = cell_text(row.cells.get(&1), shared)
            .replace('\'', "")
            .trim()
            .to_owned();
        if !indicator.is_empty() {
            compiled.indicator_rows.push(row.row_index);
        }
        compiled.rows.push(row);
    }
    if !header_seen {
        return Ok(None); // 无表头行的非表单表（月报1批 等）：跳过编译
    }
    compiled.base_headers = base_headers;
    compiled.keep_cols = keep_cols.clone();
    let col_count = compiled.base_headers.len() as u32;
    let block_start = compiled.keep_cols.iter().copied().max().unwrap_or(3) + 1;
    let first_data = compiled.rows.iter().find(|row| row.row_index > 1);
    for (offset, column) in keep_cols.iter().enumerate() {
        compiled.base_style_by_col.insert(
            *column,
            first_data
                .and_then(|row| row.style_by_col.get(column))
                .cloned()
                .unwrap_or_default(),
        );
        let offset = offset as u32;
        compiled
            .base_style_by_col
            .insert(block_start + col_count + offset, String::new());
        compiled
            .base_style_by_col
            .insert(block_start + 2 * col_count + offset, String::new());
        compiled
            .base_style_by_col
            .insert(block_start + 3 * col_count + offset, String::new());
    }
    compiled.header_style = header_style;
    let max_row = compiled
        .rows
        .iter()
        .map(|row| row.row_index)
        .max()
        .unwrap_or(1);
    compiled.new_dimension = format!("A1:{}{max_row}", col_letter(3 + 4 * col_count));
    Ok(Some(compiled))
}

// ---------------------------------------------------------------------------
// 渲染期样式表（克隆基础 xf + 内置 numFmt + 警戒色 fill；按 (base,fmt,fill) 去重）
// ---------------------------------------------------------------------------

pub struct StyleBank {
    pub styles_xml: String,
    map: HashMap<String, String>,
    xf_list: Vec<String>,
    xf_count: u32,
    fill_count: u32,
    new_fills: Vec<String>,
    new_xfs: Vec<String>,
    fill_index_by_color: HashMap<String, String>,
}

impl StyleBank {
    pub fn new(
        styles_xml: &str,
        base_indexes: &HashSet<String>,
        colors: &HashSet<u32>,
    ) -> Result<Self> {
        let xf_block_re =
            regex::Regex::new(r"(?s)<cellXfs\b[^>]*>(.*?)</cellXfs>").expect("cellXfs regex");
        let fills_block_re =
            regex::Regex::new(r"(?s)<fills\b[^>]*>(.*?)</fills>").expect("fills regex");
        let xfs_match = xf_block_re
            .captures(styles_xml)
            .with_context(|| "styles.xml 缺少 cellXfs 定义".to_owned())?;
        let fills_match = fills_block_re
            .captures(styles_xml)
            .with_context(|| "styles.xml 缺少 fills 定义".to_owned())?;
        let xf_item_re =
            regex::Regex::new(r"(?s)<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>").expect("xf regex");
        let xf_list: Vec<String> = xf_item_re
            .captures_iter(&xfs_match[1])
            .map(|caps| caps[0].to_owned())
            .collect();
        let fill_count = fills_match[1].matches("<fill>").count() as u32;
        let xf_count = xf_list.len() as u32;
        let mut bank = Self {
            styles_xml: styles_xml.to_owned(),
            map: Default::default(),
            xf_list,
            xf_count,
            fill_count,
            new_fills: Vec::new(),
            new_xfs: Vec::new(),
            fill_index_by_color: Default::default(),
        };
        let mut bases: Vec<&String> = base_indexes.iter().collect();
        bases.sort_by_key(|value| value.parse::<u32>().unwrap_or(0));
        let mut colors: Vec<u32> = colors.iter().copied().collect();
        colors.sort();
        for base in &bases {
            bank.ensure_variant(base, "value", None);
            bank.ensure_variant(base, "percent", None);
        }
        for base in &bases {
            for color in &colors {
                bank.ensure_variant(base, "value", Some(*color));
                bank.ensure_variant(base, "percent", Some(*color));
            }
        }
        if !bank.new_fills.is_empty() {
            bank.styles_xml = bank
                .styles_xml
                .replacen(
                    &format!(r#"<fills count="{fill_count}">"#),
                    &format!(
                        r#"<fills count="{}">"#,
                        fill_count + bank.new_fills.len() as u32
                    ),
                    1,
                )
                .replace(
                    "</fills>",
                    &format!("{}{}", bank.new_fills.concat(), "</fills>"),
                );
        }
        if !bank.new_xfs.is_empty() {
            bank.styles_xml = bank
                .styles_xml
                .replacen(
                    &format!(r#"<cellXfs count="{xf_count}">"#),
                    &format!(
                        r#"<cellXfs count="{}">"#,
                        xf_count + bank.new_xfs.len() as u32
                    ),
                    1,
                )
                .replace(
                    "</cellXfs>",
                    &format!("{}{}", bank.new_xfs.concat(), "</cellXfs>"),
                );
        }
        Ok(bank)
    }

    fn clone_xf(base_xml: &str, num_fmt: Option<&str>, fill_id: Option<&str>) -> String {
        let mut element = base_xml.to_owned();
        if let Some(num_fmt) = num_fmt {
            if element.contains("numFmtId=\"") {
                element = regex::Regex::new(r#"numFmtId="[^"]*""#)
                    .expect("numFmtId regex")
                    .replace_all(&element, format!(r#"numFmtId="{num_fmt}""#))
                    .into_owned();
            } else {
                element = element.replacen("<xf ", &format!(r#"<xf numFmtId="{num_fmt}" "#), 1);
            }
            if !element.contains("applyNumberFormat=") {
                element = element.replacen("<xf ", r#"<xf applyNumberFormat="1" "#, 1);
            }
        }
        if let Some(fill_id) = fill_id {
            if element.contains("fillId=\"") {
                element = regex::Regex::new(r#"fillId="[^"]*""#)
                    .expect("fillId regex")
                    .replace_all(&element, format!(r#"fillId="{fill_id}""#))
                    .into_owned();
            } else {
                element = element.replacen("<xf ", &format!(r#"<xf fillId="{fill_id}" "#), 1);
            }
            if !element.contains("applyFill=") {
                element = element.replacen("<xf ", r#"<xf applyFill="1" "#, 1);
            }
        }
        element
    }

    fn ensure_variant(&mut self, base_idx: &str, kind: &str, color: Option<u32>) {
        let key = match color {
            Some(color) => format!("{base_idx}|{kind}|{color}"),
            None => format!("{base_idx}|{kind}"),
        };
        if self.map.contains_key(&key) {
            return;
        }
        let base_position = base_idx.parse::<usize>().unwrap_or(0);
        let base_xml = self
            .xf_list
            .get(base_position)
            .cloned()
            .unwrap_or_else(|| "<xf/>".to_owned());
        let num_fmt = if kind == "value" { "2" } else { "10" };
        let fill_id = color.map(|color| {
            let hex = format!("{color:06X}");
            if let Some(existing) = self.fill_index_by_color.get(&hex) {
                return existing.clone();
            }
            let fill = format!(
                r#"<fill><patternFill patternType="solid"><fgColor rgb="FF{hex}"/><bgColor indexed="64"/></patternFill></fill>"#
            );
            let id = self.fill_count + self.new_fills.len() as u32;
            self.new_fills.push(fill);
            self.fill_index_by_color.insert(hex.clone(), id.to_string());
            id.to_string()
        });
        let element = Self::clone_xf(&base_xml, Some(num_fmt), fill_id.as_deref());
        let new_index = self.xf_count + self.new_xfs.len() as u32;
        self.new_xfs.push(element);
        self.map.insert(key, new_index.to_string());
    }

    pub fn style_for(&self, base: &str, kind: &str, color: Option<u32>) -> String {
        match color {
            Some(color) => self
                .map
                .get(&format!("{base}|{kind}|{color}"))
                .cloned()
                .unwrap_or_else(|| {
                    self.map
                        .get(&format!("{base}|{kind}"))
                        .cloned()
                        .unwrap_or_else(|| "0".to_owned())
                }),
            None => self
                .map
                .get(&format!("{base}|{kind}"))
                .cloned()
                .unwrap_or_else(|| "0".to_owned()),
        }
    }
}
