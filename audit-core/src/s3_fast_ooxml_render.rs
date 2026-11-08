//! S3-F05 FAST 渲染半场：SheetPlan 删表/降级、回填、definedNames 归一、
//! bookViews 收敛、轻量完整性校验与 ZIP 写出。数据结构见 `s3_fast_ooxml`。
use std::collections::{BTreeMap, HashMap, HashSet};

use anyhow::{bail, Context, Result};

use crate::s3_fast_ooxml::{
    col_letter, xml_attr, CompiledSheet, CompiledTemplate, SheetEntry, StyleBank,
};

/// 单格回填目标：值 + 可选数字格式/填充色（来自编辑计划的同格合并）。
#[derive(Clone, Debug, Default)]
pub struct CellFill {
    pub value: Option<excel_editor::CellValue>,
    pub num_fmt: Option<String>,
    /// 0xRRGGBB（警戒色/填充）
    pub fill_rgb: Option<u32>,
}

#[derive(Default)]
pub struct SheetRenderInput {
    pub cells: BTreeMap<(u32, u32), CellFill>,
    pub hidden_rows: HashSet<u32>,
}

#[derive(Default)]
pub struct OrgRenderInput {
    /// 有回填内容的表
    pub sheets: BTreeMap<String, SheetRenderInput>,
    /// 删表候选（内嵌配置页 + 禁用表单 + 空表；最终可能降级隐藏）
    pub delete_candidates: Vec<String>,
    pub hide_empty_rows: bool,
    pub delete_empty_sheets: bool,
}

pub struct OrgRenderOutput {
    pub parts: BTreeMap<String, Vec<u8>>,
    pub part_order: Vec<String>,
    pub workbook_xml: String,
    pub workbook_rels_xml: String,
    pub content_types_xml: String,
    pub filled: usize,
    pub deleted_sheets: Vec<String>,
    pub hidden_sheets: Vec<(String, String)>,
}

struct SheetPlan {
    sheets: Vec<SheetEntry>,
    by_name: HashMap<String, usize>,
    deleted: HashSet<String>,
    hidden: Vec<(String, String)>,
}

impl SheetPlan {
    fn new(registry: &[SheetEntry]) -> Self {
        Self {
            sheets: registry.to_vec(),
            by_name: registry
                .iter()
                .enumerate()
                .map(|(index, entry)| (entry.name.clone(), index))
                .collect(),
            deleted: Default::default(),
            hidden: Vec::new(),
        }
    }
    fn mark_deleted(&mut self, name: &str) {
        self.deleted.insert(name.to_owned());
        self.hidden.retain(|(hidden, _)| hidden != name);
    }
    #[allow(dead_code)]
    fn mark_hidden(&mut self, name: &str, reason: &str) {
        if !self.deleted.contains(name) && !self.hidden.iter().any(|(n, _)| n == name) {
            self.hidden.push((name.to_owned(), reason.to_owned()));
        }
    }
    fn survivors(&self) -> Vec<&SheetEntry> {
        self.sheets
            .iter()
            .filter(|entry| !self.deleted.contains(&entry.name))
            .collect()
    }
    fn old_to_new(&self) -> HashMap<usize, usize> {
        self.survivors()
            .iter()
            .enumerate()
            .map(|(new, entry)| (entry.old_index, new))
            .collect()
    }
    fn is_empty(&self) -> bool {
        self.deleted.is_empty() && self.hidden.is_empty()
    }
}

fn sheet_referenced_by_formulas(
    sheet_name: &str,
    parts: &BTreeMap<String, Vec<u8>>,
    registry: &[SheetEntry],
    exclude: &HashSet<String>,
) -> bool {
    let quoted = format!("'{sheet_name}'!");
    let bare = format!("{sheet_name}!");
    for info in registry {
        if info.name == sheet_name || exclude.contains(&info.name) {
            continue;
        }
        let Some(data) = parts.get(&info.part) else {
            continue;
        };
        let text = String::from_utf8_lossy(data);
        if text.contains(&quoted) {
            return true;
        }
        let mut search = 0usize;
        while let Some(found) = text[search..].find(&bare) {
            let absolute = search + found;
            let before = if absolute == 0 {
                ' '
            } else {
                text[..absolute].chars().next_back().unwrap_or(' ')
            };
            if !before.is_ascii_alphanumeric() && before != '_' && before != '\'' {
                return true;
            }
            search = absolute + bare.len();
        }
    }
    false
}

fn finalize(
    workbook_xml: &mut String,
    rels_xml: &mut String,
    content_types: &mut String,
    parts: &mut BTreeMap<String, Vec<u8>>,
    plan: &mut SheetPlan,
) {
    let sheet_entry_re = regex::Regex::new(r#"<(?:[\w.]+:)?sheet\b[^>]*/>"#).expect("sheet entry");
    let mut deleted_sorted: Vec<&String> = plan.deleted.iter().collect();
    deleted_sorted.sort_by_key(|name| {
        plan.sheets
            .iter()
            .find(|entry| &entry.name == *name)
            .map(|entry| entry.old_index)
            .unwrap_or(usize::MAX)
    });
    for name in deleted_sorted {
        let Some(index) = plan.by_name.get(name) else {
            continue;
        };
        let info = &plan.sheets[*index];
        if let Some(entry) = sheet_entry_re
            .find_iter(workbook_xml)
            .find(|entry| xml_attr(entry.as_str(), "name") == *name)
            .map(|entry| entry.as_str().to_owned())
        {
            *workbook_xml = workbook_xml.replace(&entry, "");
        }
        let rel_entry = regex::Regex::new(r#"<(?:[\w.]+:)?Relationship\b[^>]*/>"#)
            .expect("rel entry")
            .find_iter(&rels_xml)
            .find(|entry| xml_attr(entry.as_str(), "Id") == info.rid)
            .map(|entry| entry.as_str().to_owned());
        if let Some(entry) = rel_entry {
            *rels_xml = rels_xml.replace(&entry, "");
        }
        let override_entry = regex::Regex::new(r#"<Override\b[^>]*/>"#)
            .expect("override entry")
            .find_iter(&content_types)
            .find(|entry| xml_attr(entry.as_str(), "PartName").trim_start_matches('/') == info.part)
            .map(|entry| entry.as_str().to_owned());
        if let Some(entry) = override_entry {
            *content_types = content_types.replace(&entry, "");
        }
        parts.remove(&info.part);
        let file_name = info.part.rsplit('/').next().unwrap_or_default();
        parts.remove(&format!("xl/worksheets/_rels/{file_name}.rels"));
    }
    if !plan.deleted.is_empty() {
        if parts.remove("xl/calcChain.xml").is_some() {
            *content_types =
                regex::Regex::new(r#"<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>"#)
                    .expect("calcChain override")
                    .replace_all(content_types.as_str(), "")
                    .into_owned();
            *rels_xml = regex::Regex::new(
                r#"<(?:[\w.]+:)?Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>"#,
            )
            .expect("calcChain rel")
            .replace_all(rels_xml.as_str(), "")
            .into_owned();
        }
    }
    for (name, _) in &plan.hidden {
        if let Some(entry) = sheet_entry_re
            .find_iter(workbook_xml)
            .find(|candidate| xml_attr(candidate.as_str(), "name") == *name)
            .map(|entry| entry.as_str().to_owned())
        {
            let replacement = if entry.contains("state=\"") {
                regex::Regex::new(r#"state="[^"]*""#)
                    .expect("state regex")
                    .replace_all(&entry, r#"state="hidden""#)
                    .into_owned()
            } else if let Some(stripped) = entry.strip_suffix("/>") {
                format!("{stripped} state=\"hidden\"/>")
            } else {
                entry.clone()
            };
            *workbook_xml = workbook_xml.replacen(&entry, &replacement, 1);
        }
    }
    let mapping = plan.old_to_new();
    let sheet_count = sheet_entry_re.find_iter(workbook_xml).count();
    if sheet_count == 0 {
        return;
    }
    let views_re = regex::Regex::new(r"(?s)<bookViews>.*?</bookViews>").expect("bookViews regex");
    if let Some(views) = views_re.captures(workbook_xml) {
        let mut patched = views[0].to_owned();
        for attribute in ["activeTab", "firstSheet"] {
            let attr_re =
                regex::Regex::new(&format!(r#"(<workbookView\b[^>]*?){attribute}="(-?\d+)""#))
                    .expect("view attr regex");
            patched = attr_re
                .replace_all(&patched, |caps: &regex::Captures| {
                    let value: i64 = caps[2].parse().unwrap_or(0);
                    let mapped = if value >= 0 && mapping.contains_key(&(value as usize)) {
                        mapping[&(value as usize)]
                    } else {
                        (value.max(0) as usize).min(sheet_count.saturating_sub(1))
                    };
                    format!(r#"{}{}="{}""#, &caps[1], attribute, mapped)
                })
                .into_owned();
        }
        *workbook_xml = workbook_xml.replacen(&views[0], &patched, 1);
    }
}

fn resolve_defined_names_and_apply(
    workbook_xml: &str,
    mut rels_xml: String,
    mut content_types: String,
    parts: &mut BTreeMap<String, Vec<u8>>,
    plan: &mut SheetPlan,
    deletable: &HashSet<String>,
    registry: &[SheetEntry],
) -> (String, String, String) {
    let mut workbook_xml = workbook_xml.to_owned();
    let mut rels_xml = rels_xml;
    let mut content_types = content_types;
    let deleted_now: HashSet<String> = plan.deleted.union(deletable).cloned().collect();

    let block_re =
        regex::Regex::new(r"(?s)<(?:[\w.]+:)?definedNames\b[^>]*>.*?</(?:[\w.]+:)?definedNames>")
            .expect("definedNames block regex");
    let item_re = regex::Regex::new(
        r"(?s)<((?:[\w.]+:)?definedName)\b([^>]*)>(.*?)</(?:[\w.]+:)?definedName>",
    )
    .expect("definedName regex");
    let Some(block) = block_re.captures(&workbook_xml) else {
        finalize(
            &mut workbook_xml,
            &mut rels_xml,
            &mut content_types,
            parts,
            plan,
        );
        return (workbook_xml, rels_xml, content_types);
    };
    let block_text = block[0].to_owned();

    // 收集每个名称的处置：owner 属硬删表→drop；owner 属软删候选且名称仍被
    // 保留表公式以名称引用→救回（降级隐藏）；越界→drop；owner 保留→remap。
    let mut vetoes: HashSet<String> = HashSet::new();
    for caps in item_re.captures_iter(&block_text) {
        let attributes = caps[2].to_owned();
        let body = caps[3].to_owned();
        let name = xml_attr(&format!("<dn {attributes}"), "name");
        let local_sheet_id = regex::Regex::new(r#"localSheetId="(\d+)""#)
            .expect("localSheetId regex")
            .captures(&attributes)
            .and_then(|caps| caps[1].parse::<usize>().ok());
        let Some(owner_old) = local_sheet_id else {
            continue;
        };
        let Some(owner) = registry.iter().find(|entry| entry.old_index == owner_old) else {
            continue;
        };
        if !deletable.contains(&owner.name) {
            continue;
        }
        let mut soft_refs = Vec::new();
        for candidate in deletable {
            if body.contains(&format!("'{}'!", candidate))
                || body.contains(&format!("{}!", candidate))
            {
                soft_refs.push(candidate.clone());
            }
        }
        if soft_refs.is_empty() {
            continue;
        }
        let name_pattern = format!("{name}!");
        let mut used = false;
        for entry in registry {
            if deleted_now.contains(&entry.name) || deletable.contains(&entry.name) {
                continue;
            }
            if let Some(data) = parts.get(&entry.part) {
                if String::from_utf8_lossy(data).contains(&name_pattern) {
                    used = true;
                    break;
                }
            }
        }
        if used {
            vetoes.insert(owner.name.clone());
        }
    }
    for veto in &vetoes {
        plan.mark_hidden(veto, "defined_name_formula_reference");
    }
    for candidate in deletable {
        if !vetoes.contains(candidate) {
            plan.mark_deleted(candidate);
        }
    }

    // 组装保留名称：drop / remap（localSheetId 按 old→new 重算）
    let mapping = plan.old_to_new();
    let mut kept: Vec<String> = Vec::new();
    for caps in item_re.captures_iter(&block_text) {
        let whole = caps[0].to_owned();
        let attributes = caps[2].to_owned();
        let local_sheet_id = regex::Regex::new(r#"localSheetId="(\d+)""#)
            .expect("localSheetId regex 2")
            .captures(&attributes)
            .and_then(|caps| caps[1].parse::<usize>().ok());
        let owner = local_sheet_id
            .and_then(|old| registry.iter().find(|entry| entry.old_index == old))
            .map(|entry| entry.name.clone());
        let drop = match (&owner, local_sheet_id) {
            (Some(owner), _) => plan.deleted.contains(owner),
            (None, Some(_)) => true, // 越界丢弃
            _ => false,
        };
        if drop {
            continue;
        }
        if let Some(old_id) = local_sheet_id {
            let new_id = mapping.get(&old_id).copied().unwrap_or(0);
            let patched = regex::Regex::new(r#"localSheetId="[^"]*""#)
                .expect("remap regex")
                .replace_all(&whole, format!(r#"localSheetId="{new_id}""#))
                .into_owned();
            kept.push(patched);
        } else {
            kept.push(whole);
        }
    }
    let open_tag = regex::Regex::new(r"(?s)^<(?:[\w.]+:)?definedNames\b[^>]*>")
        .expect("dn block open")
        .captures(&block_text)
        .map(|caps| caps[0].to_owned())
        .unwrap_or_else(|| "<definedNames>".to_owned());
    let replacement = if kept.is_empty() {
        String::new()
    } else {
        format!("{open_tag}{}{}</definedNames>", kept.concat(), "")
            .replace("</definedNames></definedNames>", "</definedNames>")
    };
    workbook_xml = workbook_xml.replacen(&block_text, &replacement, 1);

    finalize(
        &mut workbook_xml,
        &mut rels_xml,
        &mut content_types,
        parts,
        plan,
    );
    (workbook_xml, rels_xml, content_types)
}

/// 单表重建 sheetData（纯字符串拼接：编译行内容复用 + 计划目标覆盖）。
fn render_sheet(
    sheet: &CompiledSheet,
    input: &SheetRenderInput,
    styles: &StyleBank,
    hide_empty_rows: bool,
) -> (String, usize) {
    let header_style_base = if sheet.header_style.is_empty() {
        "0".to_owned()
    } else {
        sheet.header_style.clone()
    };
    let header_style = styles.style_for(&header_style_base, "value", None);
    let mut filled = 0usize;
    let mut out_rows: Vec<String> = Vec::new();
    for row in &sheet.rows {
        let mut cells: BTreeMap<u32, String> = row
            .cells
            .iter()
            .filter(|(column, _)| **column < 4 || sheet.keep_cols.contains(column))
            .map(|(column, xml)| (*column, xml.clone()))
            .collect();
        let mut attrs = row.attrs.clone();
        let row_targets: Vec<(&(u32, u32), &CellFill)> = input
            .cells
            .iter()
            .filter(|((target_row, _), _)| *target_row == row.row_index)
            .collect();
        let has_hits = !row_targets.is_empty() || input.hidden_rows.contains(&row.row_index);
        for ((_, target_col), fill) in &row_targets {
            let target_col = *target_col;
            let reference = format!("{}{}", col_letter(target_col), row.row_index);
            let kind = if fill.num_fmt.as_deref() == Some("0.00%") {
                "percent"
            } else {
                "value"
            };
            let base = if row.row_index == 1 {
                header_style_base.clone()
            } else {
                sheet
                    .base_style_by_col
                    .get(&target_col)
                    .cloned()
                    .unwrap_or_default()
            };
            let style_index = styles.style_for(
                if base.is_empty() { "0" } else { &base },
                kind,
                fill.fill_rgb,
            );
            let style_attribute = format!(" s=\"{style_index}\"");
            match &fill.value {
                Some(value) => match value {
                    excel_editor::CellValue::Text(text) => cells.insert(
                        target_col,
                        format!(
                            r#"<c r="{reference}"{style_attribute} t="inlineStr"><is><t>{}</t></is></c>"#,
                            escape_xml_text(text)
                        ),
                    ),
                    excel_editor::CellValue::Decimal(number)
                    | excel_editor::CellValue::DateTime {
                        serial: number, ..
                    } => cells.insert(
                        target_col,
                        format!(r#"<c r="{reference}"{style_attribute}><v>{number}</v></c>"#),
                    ),
                    excel_editor::CellValue::Boolean(flag) => cells.insert(
                        target_col,
                        format!(
                            r#"<c r="{reference}"{style_attribute} t="b"><v>{}</v></c>"#,
                            if *flag { 1 } else { 0 }
                        ),
                    ),
                    excel_editor::CellValue::Error(token) => cells.insert(
                        target_col,
                        format!(
                            r#"<c r="{reference}"{style_attribute} t="e"><v>{}</v></c>"#,
                            escape_xml_text(token)
                        ),
                    ),
                    excel_editor::CellValue::Blank => cells.insert(
                        target_col,
                        format!(r#"<c r="{reference}"{style_attribute}/>"#),
                    ),
                },
                None => cells.insert(
                    target_col,
                    format!(r#"<c r="{reference}"{style_attribute}/>"#),
                ),
            };
            filled += 1;
        }
        if hide_empty_rows
            && !has_hits
            && row.row_index > 1
            && sheet.indicator_rows.contains(&row.row_index)
            && !attrs.contains("hidden=\"")
        {
            attrs.push_str(" hidden=\"1\"");
        }
        if cells.is_empty() {
            out_rows.push(format!("<row{attrs}/>"));
        } else {
            out_rows.push(format!(
                "<row{attrs}>{}</row>",
                cells.values().cloned().collect::<String>()
            ));
        }
    }
    let dimension_re = regex::Regex::new(r#"(<dimension ref=")[^"]*(")"#).expect("dimension regex");
    let prefix = dimension_re
        .replace(
            &sheet.prefix,
            format!(r#"${{1}}{}${{2}}"#, sheet.new_dimension),
        )
        .into_owned();
    (
        format!(
            "{prefix}<sheetData>{}</sheetData>{}",
            out_rows.concat(),
            sheet.suffix
        ),
        filled,
    )
}

fn escape_xml_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

/// 每机构渲染：回填 + 删表/降级 + 落盘 + 轻量完整性校验。
pub fn render_org(
    compiled: &CompiledTemplate,
    styles: &StyleBank,
    input: &OrgRenderInput,
) -> Result<OrgRenderOutput> {
    let mut parts = compiled.parts.clone();
    let mut workbook_xml = compiled.workbook_xml.clone();
    let mut plan = SheetPlan::new(&compiled.sheet_registry);
    for name in &compiled.internal_config_sheets {
        plan.mark_deleted(name);
    }
    for name in &input.delete_candidates {
        plan.mark_deleted(name);
    }
    let mut filled_total = 0usize;
    let mut candidates: HashSet<String> = HashSet::new();
    for (name, sheet) in &compiled.sheets {
        let Some(input_sheet) = input.sheets.get(name) else {
            if input.delete_empty_sheets {
                candidates.insert(name.clone());
            }
            continue;
        };
        let (generated, filled) = render_sheet(sheet, input_sheet, styles, input.hide_empty_rows);
        parts.insert(sheet.part.clone(), generated.into_bytes());
        filled_total += filled;
    }
    // 公式引用收敛：被保留表公式引用的候选 → 降级隐藏（迭代至收敛）
    let mut doomed: HashSet<String> = candidates;
    let mut changed = true;
    while changed && !doomed.is_empty() {
        changed = false;
        let ordered: Vec<String> = {
            let mut names: Vec<String> = doomed.iter().cloned().collect();
            names.sort_by_key(|name| {
                compiled
                    .sheet_registry
                    .iter()
                    .find(|entry| &entry.name == name)
                    .map(|entry| entry.old_index)
                    .unwrap_or(usize::MAX)
            });
            names
        };
        for name in ordered {
            let excluded: HashSet<String> = doomed
                .iter()
                .filter(|candidate| **candidate != name)
                .cloned()
                .chain(plan.deleted.iter().cloned())
                .collect();
            if sheet_referenced_by_formulas(&name, &parts, &compiled.sheet_registry, &excluded) {
                doomed.remove(&name);
                plan.mark_hidden(&name, "formula_reference");
                changed = true;
                break;
            }
        }
    }
    let (mut workbook_xml, mut rels_xml, mut content_types) = resolve_defined_names_and_apply(
        &workbook_xml,
        compiled.workbook_rels_xml.clone(),
        compiled.content_types_xml.clone(),
        &mut parts,
        &mut plan,
        &doomed,
        &compiled.sheet_registry,
    );
    let deleted_sheets = plan.deleted.iter().cloned().collect::<Vec<_>>();
    let hidden_sheets = plan.hidden.clone();
    validate_core(&workbook_xml, &rels_xml, &content_types, &parts)?;
    Ok(OrgRenderOutput {
        parts,
        part_order: compiled.part_order.clone(),
        workbook_xml,
        workbook_rels_xml: rels_xml,
        content_types_xml: content_types,
        filled: filled_total,
        deleted_sheets,
        hidden_sheets,
    })
}

fn validate_core(
    workbook_xml: &str,
    rels_xml: &str,
    content_types: &str,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    roxmltree::Document::parse(workbook_xml).context("输出 workbook.xml 无法解析")?;
    roxmltree::Document::parse(rels_xml).context("输出 workbook rels 无法解析")?;
    roxmltree::Document::parse(content_types).context("输出 Content_Types 无法解析")?;
    // sheet ↔ relationship ↔ part 无悬空
    let rel_re = regex::Regex::new(r#"<(?:[\w.]+:)?Relationship\b[^>]*/>"#).expect("rel validate");
    let mut relationships: HashMap<String, String> = HashMap::new();
    for caps in rel_re.captures_iter(rels_xml) {
        let entry = caps.get(0).unwrap().as_str();
        relationships.insert(
            xml_attr(entry, "Id"),
            normalize_part_target(&xml_attr(entry, "Target")),
        );
    }
    let sheet_re = regex::Regex::new(r#"<(?:[\w.]+:)?sheet\b[^>]*/>"#).expect("sheet validate");
    for caps in sheet_re.captures_iter(workbook_xml) {
        let entry = caps.get(0).unwrap().as_str();
        let rid = xml_attr(entry, "r:id");
        let Some(target) = relationships.get(&rid) else {
            bail!("输出工作表关系悬空：{rid}");
        };
        if !parts.contains_key(target) {
            bail!("输出工作表部件缺失：{target}");
        }
    }
    Ok(())
}

fn normalize_part_target(target: &str) -> String {
    let trimmed = target.trim_start_matches('/');
    if trimmed.starts_with("xl/") {
        trimmed.to_owned()
    } else {
        format!("xl/{trimmed}")
    }
}

/// 渲染产物写 ZIP（part_order 顺序，跳过已删除部件）。
pub fn write_zip(
    output_path: &std::path::Path,
    part_order: &[String],
    parts: &BTreeMap<String, Vec<u8>>,
    workbook_xml: &str,
    rels_xml: &str,
    content_types: &str,
) -> Result<()> {
    use zip::write::SimpleFileOptions;
    if let Some(parent) = output_path.parent() {
        std::fs::create_dir_all(parent)
            .with_context(|| format!("创建输出目录：{}", parent.display()))?;
    }
    let file = std::fs::File::create(output_path)
        .with_context(|| format!("创建输出文件：{}", output_path.display()))?;
    let mut writer = zip::ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for name in part_order {
        let content: Vec<u8> = match name.as_str() {
            "xl/workbook.xml" => workbook_xml.as_bytes().to_vec(),
            "xl/_rels/workbook.xml.rels" => rels_xml.as_bytes().to_vec(),
            "[Content_Types].xml" => content_types.as_bytes().to_vec(),
            other => match parts.get(other) {
                Some(data) => data.clone(),
                None => continue,
            },
        };
        writer
            .start_file(name.as_str(), options)
            .with_context(|| format!("写入 ZIP 条目 {name}"))?;
        std::io::Write::write_all(&mut writer, &content)?;
    }
    writer.finish().context("收尾 ZIP 输出")?;
    Ok(())
}
