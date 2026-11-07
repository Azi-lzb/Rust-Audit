//! OOXML 直改后端：把模板/外部工作簿的区域写入审核副本（S1-F01）。
//!
//! 参考旧 Python Audit 的 `formula_region_writer.py`（DIRECT_OOXML）技术路径，
//! 但为独立 Rust 实现：只重写目标工作表的 `sheetData` 与 `mergeCells`，其余
//! package 部件按字节原样拷贝，因此不会像整簿 load→save 那样改写关系列表、
//! 列尺寸或未触及的对象。
//!
//! 与旧实现的关键差异（均为更保守的选择）：
//! - 只支持 S1-F01 实际使用的操作组合：CreateSheet / DeleteSheet / ImportRange
//!   / ImportRangeWithFormulaTranslation / SetWorkbookCalculationProperties；
//!   任何其他操作在写入前拒绝（`unsupported_before_write`）。
//! - 样式复制（`style_mode != None`）暂不支持：S1 计划使用 `None`，需要样式的
//!   计划直接拒绝而不是静默降级。
//! - 共享公式（`<f t="shared">`）在读取模板区域时即展开为普通公式，避免写出
//!   空 `<f>` 导致 Excel 判定文件损坏。
//! - `calcChain` 在写公式后删除并置 `fullCalcOnLoad`，由重算重建缓存。

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, RangeStyleMode, WorkbookFormat,
};
use anyhow::{Context, Result, bail, ensure};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{Read, Write};
use std::path::Path;
use zip::{CompressionMethod, ZipArchive, ZipWriter, write::SimpleFileOptions};

/// 本后端支持的操作集合。
pub const SUPPORTED_OPERATIONS: &[EditOperationKind] = &[
    EditOperationKind::CreateSheet,
    EditOperationKind::DeleteSheet,
    EditOperationKind::ImportRange,
    EditOperationKind::ImportRangeWithFormulaTranslation,
    EditOperationKind::SetWorkbookCalculationProperties,
];

pub struct OoxmlImportBackend {
    capability: BackendCapability,
}

impl OoxmlImportBackend {
    /// 调用方提供能力证据；本构造函数不会把未验证的操作提升为 passed。
    pub fn new(capability: BackendCapability) -> Result<Self> {
        if capability.engine != crate::EngineId::OoxmlPatch {
            bail!("OOXML 导入后端必须登记为 ooxml_patch");
        }
        for kind in SUPPORTED_OPERATIONS {
            ensure!(
                capability.operations.contains(kind),
                "OOXML 导入后端能力证据缺少操作 {kind:?}"
            );
        }
        Ok(Self { capability })
    }
}

impl EditBackend for OoxmlImportBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "OOXML 导入后端当前只接受 XLSX",
            ));
        }
        // 写前拒绝：任何超出能力集的操作都不允许产生部分写入。
        for operation in &plan.operations {
            let kind = operation.kind();
            if !SUPPORTED_OPERATIONS.contains(&kind) {
                return Err(BackendFailure::unsupported_before_write(format!(
                    "OOXML 导入后端不支持操作 {kind:?}；该计划必须改走其他已验收后端"
                )));
            }
            if let EditOperation::ImportRange { style_mode, .. }
            | EditOperation::ImportRangeWithFormulaTranslation { style_mode, .. } = operation
            {
                if *style_mode != RangeStyleMode::None {
                    return Err(BackendFailure::unsupported_before_write(
                        "OOXML 导入后端暂不支持样式复制；计划必须使用 style_mode=None",
                    ));
                }
            }
        }
        apply_plan(plan).map_err(|error| BackendFailure::failed(format!("{error:#}")))
    }
}

// ---------------------------------------------------------------------------
// 实现
// ---------------------------------------------------------------------------

struct Package {
    order: Vec<String>,
    parts: BTreeMap<String, Vec<u8>>,
}

impl Package {
    fn read(path: &Path) -> Result<Self> {
        let mut archive = ZipArchive::new(fs::File::open(path)?)
            .with_context(|| format!("打开 XLSX 包：{}", path.display()))?;
        let mut order = Vec::with_capacity(archive.len());
        let mut parts = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            if entry.is_dir() {
                continue;
            }
            let name = entry.name().to_owned();
            let mut data = Vec::new();
            entry.read_to_end(&mut data)?;
            order.push(name.clone());
            parts.insert(name, data);
        }
        Ok(Self { order, parts })
    }

    fn text(&self, part: &str) -> Result<String> {
        let bytes = self
            .parts
            .get(part)
            .with_context(|| format!("XLSX 包缺少部件 {part}"))?;
        Ok(String::from_utf8(bytes.clone()).with_context(|| format!("部件 {part} 不是 UTF-8"))?)
    }

    fn set_text(&mut self, part: &str, text: String) {
        if !self.parts.contains_key(part) {
            self.order.push(part.to_owned());
        }
        self.parts.insert(part.to_owned(), text.into_bytes());
    }

    fn remove(&mut self, part: &str) {
        self.order.retain(|name| name != part);
        self.parts.remove(part);
    }

    /// 写出：部件顺序与压缩设置尽量沿用输入，未改动的部件按字节原样写回。
    fn write(&self, output: &Path) -> Result<()> {
        if let Some(parent) = output.parent() {
            fs::create_dir_all(parent)?;
        }
        let mut writer = ZipWriter::new(fs::File::create(output)?);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for name in &self.order {
            let Some(data) = self.parts.get(name) else {
                continue;
            };
            writer.start_file(name.as_str(), options)?;
            writer.write_all(data)?;
        }
        writer.finish()?;
        Ok(())
    }
}

fn apply_plan(plan: &ExcelEditPlan) -> Result<()> {
    let mut package = Package::read(&plan.input)?;
    let mut sheet_parts = SheetParts::resolve(&package)?;
    let mut workbook_xml = package.text("xl/workbook.xml")?;

    for operation in &plan.operations {
        match operation {
            EditOperation::CreateSheet { name } => {
                sheet_parts.create_sheet(&mut package, &mut workbook_xml, name, None)?;
            }
            EditOperation::CreateSheetAt { name, index } => {
                // S1 不使用该操作；保留明确错误而不是静默忽略。
                let _ = (name, index);
                bail!("OOXML 导入后端未实现 CreateSheetAt");
            }
            EditOperation::DeleteSheet { name } => {
                sheet_parts.delete_sheet(&mut package, &mut workbook_xml, name)?;
            }
            EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                copy_merges,
                materialize_blank_cells,
                ..
            }
            | EditOperation::ImportRangeWithFormulaTranslation {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                copy_merges,
                materialize_blank_cells,
                ..
            } => {
                import_range(
                    &mut package,
                    &mut sheet_parts,
                    source_workbook.as_path(),
                    source_sheet,
                    source_range,
                    target_sheet,
                    target_start,
                    ImportFlags {
                        copy_values: *copy_values,
                        copy_formulas: *copy_formulas,
                        copy_comments: *copy_comments,
                        copy_merges: *copy_merges,
                        materialize_blank_cells: *materialize_blank_cells,
                    },
                )?;
            }
            EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load,
                force_full_calc,
            } => {
                workbook_xml = patch_calculation_properties(
                    &workbook_xml,
                    *full_calc_on_load,
                    *force_full_calc,
                )?;
            }
            other => bail!("OOXML 导入后端到达未预期操作：{other:?}"),
        }
    }

    // 写入过公式的计划：丢弃 calcChain（缓存已失效，由重算重建）。
    let wrote_formulas = plan.operations.iter().any(|operation| {
        matches!(
            operation,
            EditOperation::ImportRange {
                copy_formulas: true,
                ..
            } | EditOperation::ImportRangeWithFormulaTranslation {
                copy_formulas: true,
                ..
            }
        )
    });
    if wrote_formulas && package.parts.contains_key("xl/calcChain.xml") {
        package.remove("xl/calcChain.xml");
        strip_part_from_content_types(&mut package, "/xl/calcChain.xml")?;
        strip_relationship_to(&mut package, "calcChain.xml")?;
    }

    package.set_text("xl/workbook.xml", workbook_xml);
    package.write(&plan.output)
}

fn strip_part_from_content_types(package: &mut Package, part_name: &str) -> Result<()> {
    let path = "[Content_Types].xml";
    let text = package.text(path)?;
    let needle = format!("PartName=\"{part_name}\"");
    let mut removed = false;
    let mut result = String::with_capacity(text.len());
    for segment in text.split_inclusive('>') {
        let trimmed = segment.trim_start();
        if trimmed.starts_with("<Override") && segment.contains(&needle) {
            removed = true;
            continue;
        }
        result.push_str(segment);
    }
    if removed {
        package.set_text(path, result);
    }
    Ok(())
}

fn strip_relationship_to(package: &mut Package, target_fragment: &str) -> Result<()> {
    let path = "xl/_rels/workbook.xml.rels";
    if !package.parts.contains_key(path) {
        return Ok(());
    }
    let text = package.text(path)?;
    let mut removed = false;
    let mut result = String::with_capacity(text.len());
    for segment in text.split_inclusive('>') {
        let trimmed = segment.trim_start();
        if trimmed.starts_with("<Relationship") && segment.contains(target_fragment) {
            removed = true;
            continue;
        }
        result.push_str(segment);
    }
    if removed {
        package.set_text(path, result);
    }
    Ok(())
}

fn patch_calculation_properties(
    workbook_xml: &str,
    full_calc_on_load: Option<bool>,
    force_full_calc: Option<bool>,
) -> Result<String> {
    let mut desired: Vec<(&str, bool)> = Vec::new();
    if let Some(value) = full_calc_on_load {
        desired.push(("fullCalcOnLoad", value));
    }
    if let Some(value) = force_full_calc {
        desired.push(("forceFullCalc", value));
    }
    if desired.is_empty() {
        return Ok(workbook_xml.to_owned());
    }

    let Some(start) = workbook_xml.find("<calcPr") else {
        // 没有 calcPr：在 </sheets> 之后插入一个完整的自闭合元素。
        let mut insert_at = workbook_xml
            .find("</sheets>")
            .map(|position| position + "</sheets>".len());
        if insert_at.is_none() {
            insert_at = workbook_xml
                .find("<sheets/>")
                .map(|position| position + "<sheets/>".len());
        }
        let insert_at = insert_at.context("workbook.xml 缺少 sheets 元素")?;
        let attributes = desired
            .iter()
            .map(|(name, value)| format!(" {}=\"{}\"", name, if *value { 1 } else { 0 }))
            .collect::<String>();
        let mut result = String::with_capacity(workbook_xml.len() + attributes.len() + 32);
        result.push_str(&workbook_xml[..insert_at]);
        result.push_str(&format!("<calcPr calcId=\"0\"{attributes}/>"));
        result.push_str(&workbook_xml[insert_at..]);
        return Ok(result);
    };

    let rest = &workbook_xml[start..];
    let element_end = rest
        .find('>')
        .context("workbook.xml 的 calcPr 元素未闭合")?;
    let element = &rest[..=element_end];
    let self_closing = element.trim_end().ends_with("/>");
    // 元素内容去掉结尾的 `>` 或 `/>`，得到可追加属性的位置。
    let body_end = if self_closing {
        element.len() - 2
    } else {
        element.len() - 1
    };
    let mut updated = element[..body_end].trim_end().to_owned();
    for (name, value) in &desired {
        let rendered = format!("{}=\"{}\"", name, if *value { 1 } else { 0 });
        if let Some(position) = updated.find(&format!("{}=\"", name)) {
            // 已有该属性：替换其值。
            let value_start = position + name.len() + 2;
            let value_end = updated[value_start..]
                .find('"')
                .map(|offset| value_start + offset)
                .unwrap_or(updated.len());
            updated.replace_range(position..=value_end, &rendered);
        } else {
            updated.push(' ');
            updated.push_str(&rendered);
        }
    }
    if self_closing {
        updated.push_str("/>");
    } else {
        updated.push('>');
    }

    let mut result = String::with_capacity(workbook_xml.len() + updated.len());
    result.push_str(&workbook_xml[..start]);
    result.push_str(&updated);
    result.push_str(&rest[element_end + 1..]);
    Ok(result)
}

// ---------------------------------------------------------------------------
// 工作表部件定位与增删
// ---------------------------------------------------------------------------

struct SheetParts {
    /// 工作表名 → 部件路径（如 `xl/worksheets/sheet1.xml`）。
    by_name: BTreeMap<String, String>,
}

impl SheetParts {
    fn resolve(package: &Package) -> Result<Self> {
        let workbook = package.text("xl/workbook.xml")?;
        let rels = package.text("xl/_rels/workbook.xml.rels")?;
        let rel_targets = relationship_targets(&rels)?;
        let mut by_name = BTreeMap::new();
        for (name, rel_id) in worksheet_entries(&workbook)? {
            let Some(target) = rel_targets.get(&rel_id) else {
                continue;
            };
            let part = normalize_target(target);
            by_name.insert(name, part);
        }
        ensure!(!by_name.is_empty(), "workbook.xml 中没有工作表条目");
        Ok(Self { by_name })
    }

    fn part_for(&self, sheet: &str) -> Result<&str> {
        self.by_name
            .get(sheet)
            .map(String::as_str)
            .with_context(|| format!("审核副本缺少工作表：{sheet}"))
    }

    fn create_sheet(
        &mut self,
        package: &mut Package,
        workbook_xml: &mut String,
        name: &str,
        index: Option<usize>,
    ) -> Result<()> {
        ensure!(
            !self
                .by_name
                .keys()
                .any(|existing| existing.eq_ignore_ascii_case(name)),
            "工作表已存在：{name}"
        );
        let next_id = next_worksheet_part_number(package)?;
        let part = format!("xl/worksheets/sheet{next_id}.xml");
        let relationships = package.text("xl/_rels/workbook.xml.rels")?;
        let rel_id = next_relationship_id(&relationships)?;
        package.set_text(&part, EMPTY_WORKSHEET_XML.to_owned());
        append_relationship(package, &rel_id, &format!("worksheets/sheet{next_id}.xml"))?;
        append_content_type_override(
            package,
            &format!("/{part}"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        )?;
        *workbook_xml = append_sheet_entry(workbook_xml, name, &rel_id, index)?;
        self.by_name.insert(name.to_owned(), part);
        Ok(())
    }

    fn delete_sheet(
        &mut self,
        package: &mut Package,
        workbook_xml: &mut String,
        name: &str,
    ) -> Result<()> {
        let part = self.part_for(name)?.to_owned();
        let rel_id = sheet_rel_id(workbook_xml, name)?
            .with_context(|| format!("workbook.xml 缺少工作表 {name} 的关系 ID"))?;
        *workbook_xml = remove_sheet_entry(workbook_xml, name)?;
        remove_relationship(package, &rel_id)?;
        strip_part_from_content_types(package, &format!("/{part}"))?;
        package.remove(&part);
        // 工作表自己的 .rels 一并移除，避免孤立部件。
        let rels_part = format!(
            "xl/worksheets/_rels/{}",
            part.rsplit('/').next().unwrap_or_default()
        );
        let rels_part = format!("{rels_part}.rels");
        if package.parts.contains_key(&rels_part) {
            package.remove(&rels_part);
        }
        self.by_name.remove(name);
        Ok(())
    }
}

const EMPTY_WORKSHEET_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    r#"<sheetData/></worksheet>"#,
);

fn relationship_targets(rels_xml: &str) -> Result<BTreeMap<String, String>> {
    let document = roxmltree::Document::parse(rels_xml).context("解析 workbook.xml.rels")?;
    let mut result = BTreeMap::new();
    for relation in document.descendants().filter(|node| node.is_element()) {
        if relation.tag_name().name() != "Relationship" {
            continue;
        }
        let (Some(id), Some(target)) = (relation.attribute("Id"), relation.attribute("Target"))
        else {
            continue;
        };
        if relation
            .attribute("TargetMode")
            .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            continue;
        }
        result.insert(id.to_owned(), target.to_owned());
    }
    Ok(result)
}

fn next_relationship_id(rels_xml: &str) -> Result<String> {
    let document = roxmltree::Document::parse(rels_xml).context("解析 workbook.xml.rels")?;
    let used = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
        .filter_map(|node| node.attribute("Id"))
        .collect::<BTreeSet<_>>();
    let mut index = 1u64;
    loop {
        let candidate = format!("rId{index}");
        if !used.contains(candidate.as_str()) {
            return Ok(candidate);
        }
        index = index
            .checked_add(1)
            .context("无法为新增工作表分配唯一关系 ID")?;
    }
}

fn next_worksheet_part_number(package: &Package) -> Result<u32> {
    (1..=u32::MAX)
        .find(|index| {
            !package
                .parts
                .contains_key(&format!("xl/worksheets/sheet{index}.xml"))
        })
        .context("无法为新增工作表分配唯一 OOXML 部件名")
}

fn worksheet_entries(workbook_xml: &str) -> Result<Vec<(String, String)>> {
    let document = roxmltree::Document::parse(workbook_xml).context("解析 workbook.xml")?;
    let mut result = Vec::new();
    for sheet in document.descendants().filter(|node| node.is_element()) {
        if sheet.tag_name().name() != "sheet" {
            continue;
        }
        let (Some(name), Some(rel_id)) = (sheet.attribute("name"), relationship_id(sheet)) else {
            continue;
        };
        result.push((name.to_owned(), rel_id.to_owned()));
    }
    Ok(result)
}

fn relationship_id<'a>(sheet: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    sheet
        .attributes()
        .find(|attribute| attribute.name() == "id")
        .map(|attribute| attribute.value())
}

fn normalize_target(target: &str) -> String {
    let trimmed = target.trim_start_matches('/');
    if trimmed.starts_with("xl/") {
        trimmed.to_owned()
    } else {
        format!("xl/{trimmed}")
    }
}

fn sheet_rel_id(workbook_xml: &str, name: &str) -> Result<Option<String>> {
    Ok(worksheet_entries(workbook_xml)?
        .into_iter()
        .find(|(sheet_name, _)| sheet_name == name)
        .map(|(_, rel_id)| rel_id))
}

fn append_relationship(package: &mut Package, rel_id: &str, target: &str) -> Result<()> {
    let path = "xl/_rels/workbook.xml.rels";
    let text = package.text(path)?;
    let element = format!(
        "<Relationship Id=\"{rel_id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"{target}\"/>"
    );
    let insertion = text
        .rfind("</Relationships>")
        .context("workbook.xml.rels 缺少 Relationships 结束标签")?;
    let mut result = text;
    result.insert_str(insertion, &element);
    package.set_text(path, result);
    Ok(())
}

fn remove_relationship(package: &mut Package, rel_id: &str) -> Result<()> {
    let path = "xl/_rels/workbook.xml.rels";
    let text = package.text(path)?;
    let needle = format!("Id=\"{rel_id}\"");
    let mut removed = false;
    let mut result = String::with_capacity(text.len());
    for segment in text.split_inclusive('>') {
        if segment.trim_start().starts_with("<Relationship") && segment.contains(&needle) {
            removed = true;
            continue;
        }
        result.push_str(segment);
    }
    ensure!(removed, "workbook.xml.rels 缺少关系 {rel_id}");
    package.set_text(path, result);
    Ok(())
}

fn append_content_type_override(
    package: &mut Package,
    part_name: &str,
    content_type: &str,
) -> Result<()> {
    let path = "[Content_Types].xml";
    let text = package.text(path)?;
    let document = roxmltree::Document::parse(&text).context("解析 [Content_Types].xml")?;
    for override_node in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Override")
    {
        if override_node.attribute("PartName") == Some(part_name) {
            ensure!(
                override_node.attribute("ContentType") == Some(content_type),
                "部件 {part_name} 已登记了不同的 ContentType"
            );
            return Ok(());
        }
    }
    let element = format!("<Override PartName=\"{part_name}\" ContentType=\"{content_type}\"/>");
    let insertion = text
        .rfind("</Types>")
        .context("[Content_Types].xml 缺少 Types 结束标签")?;
    let mut result = text;
    result.insert_str(insertion, &element);
    package.set_text(path, result);
    Ok(())
}

fn append_sheet_entry(
    workbook_xml: &str,
    name: &str,
    rel_id: &str,
    index: Option<usize>,
) -> Result<String> {
    let element = format!(
        "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{rel_id}\"/>",
        escape_xml(name),
        next_sheet_id(workbook_xml)?
    );
    if let Some(index) = index {
        let positions = sheet_element_positions(workbook_xml);
        if index < positions.len() {
            let mut result = workbook_xml.to_owned();
            result.insert_str(positions[index].0, &element);
            return Ok(result);
        }
    }
    let mut result = workbook_xml.to_owned();
    if let Some(insertion) = result.rfind("</sheets>") {
        result.insert_str(insertion, &element);
        return Ok(result);
    }
    // 自闭合 `<sheets/>`：展开为成对标签后再插入。
    if let Some(position) = result.find("<sheets/>") {
        result.replace_range(
            position..position + "<sheets/>".len(),
            &format!("<sheets>{element}</sheets>"),
        );
        return Ok(result);
    }
    bail!("workbook.xml 缺少 sheets 元素")
}

fn remove_sheet_entry(workbook_xml: &str, name: &str) -> Result<String> {
    let needle = format!("name=\"{}\"", escape_xml(name));
    let positions = sheet_element_positions(workbook_xml);
    let mut result = workbook_xml.to_owned();
    for (start, end) in positions.into_iter().rev() {
        let slice = &result[start..end];
        if slice.contains(&needle) {
            result.replace_range(start..end, "");
            return Ok(result);
        }
    }
    bail!("workbook.xml 中找不到工作表 {name}")
}

/// 返回每个 `<sheet .../>` 元素的字节区间，按出现顺序。
fn sheet_element_positions(workbook_xml: &str) -> Vec<(usize, usize)> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = workbook_xml[cursor..].find("<sheet ") {
        let start = cursor + relative;
        let Some(end_relative) = workbook_xml[start..].find('>') else {
            break;
        };
        let end = start + end_relative + 1;
        result.push((start, end));
        cursor = end;
    }
    result
}

fn next_sheet_id(workbook_xml: &str) -> Result<u32> {
    let document = roxmltree::Document::parse(workbook_xml)?;
    let max = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "sheet")
        .filter_map(|node| node.attribute("sheetId"))
        .filter_map(|value| value.parse::<u32>().ok())
        .max()
        .unwrap_or(0);
    Ok(max.saturating_add(1))
}

/// 还原 XML 实体（源部件的公式/文本在解析时保留了原始实体）。
fn unescape_xml(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

// ---------------------------------------------------------------------------
// 区域导入
// ---------------------------------------------------------------------------

struct ImportFlags {
    copy_values: bool,
    copy_formulas: bool,
    copy_comments: bool,
    copy_merges: bool,
    materialize_blank_cells: bool,
}

#[allow(clippy::too_many_arguments)]
fn import_range(
    package: &mut Package,
    sheet_parts: &mut SheetParts,
    source_workbook: &Path,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
    target_start: &str,
    flags: ImportFlags,
) -> Result<()> {
    let source_bounds = parse_range(source_range)?;
    let (source_lo_col, source_lo_row, source_hi_col, source_hi_row) = source_bounds;
    let (row_offset, column_offset) = target_offsets(source_bounds, target_start)?;

    let source_package = Package::read(source_workbook)?;
    let source_parts = SheetParts::resolve(&source_package)?;
    let source_part = source_parts.part_for(source_sheet)?;
    let source_xml = source_package.text(source_part)?;
    let source_cells = read_region_cells(
        &source_package,
        &source_xml,
        (source_lo_col, source_lo_row, source_hi_col, source_hi_row),
    )?;

    let target_part = sheet_parts.part_for(target_sheet)?.to_owned();
    let mut target_xml = package.text(&target_part)?;
    target_xml = write_region(
        &target_xml,
        &source_cells,
        (source_lo_col, source_lo_row, source_hi_col, source_hi_row),
        row_offset,
        column_offset,
        &flags,
    )?;

    if flags.copy_merges {
        let merges = source_merges(
            &source_xml,
            (source_lo_col, source_lo_row, source_hi_col, source_hi_row),
        )?;
        if !merges.is_empty() {
            target_xml = merge_into_target(&target_xml, &merges, row_offset, column_offset)?;
        }
    }

    package.set_text(&target_part, target_xml);

    if flags.copy_comments {
        copy_region_comments(
            package,
            &source_package,
            source_part,
            &target_part,
            (source_lo_col, source_lo_row, source_hi_col, source_hi_row),
            row_offset,
            column_offset,
        )?;
    }
    Ok(())
}

/// 模板区域的一个单元格（已展开共享公式）。
struct RegionCell {
    row: u32,
    column: u32,
    /// 单元格 XML 的 `<c>` 片段内容（不含 `s` 属性；样式由计划控制）。
    formula: Option<String>,
    /// 值片段（`<v>` 或 `<is>`）。
    value_xml: Option<String>,
    /// 文本类型标记（`t="s"` 等），写入时需保留。
    type_attr: Option<String>,
}

fn read_region_cells(
    package: &Package,
    sheet_xml: &str,
    bounds: (u32, u32, u32, u32),
) -> Result<Vec<RegionCell>> {
    let (lo_col, lo_row, hi_col, hi_row) = bounds;
    let shared_strings = read_shared_strings(package)?;
    let shared_formulas = shared_formula_masters(sheet_xml)?;
    let mut cells = Vec::new();
    for cell in cell_elements(sheet_xml) {
        let Some((row, column)) = cell_reference(&cell.xml) else {
            continue;
        };
        if row < lo_row || row > hi_row || column < lo_col || column > hi_col {
            continue;
        }
        let formula = match cell.formula_body.as_ref() {
            Some((attrs, body)) if attrs.contains("t=\"shared\"") && body.trim().is_empty() => {
                Some(expand_shared_formula(
                    &shared_formulas,
                    attrs,
                    body,
                    &cell.xml,
                )?)
            }
            Some((_attrs, body)) => Some(body.clone()),
            None => None,
        };
        let value_xml = if cell.has_inline_string {
            cell.xml
                .find("<is>")
                .map(|start| strip_trailing_tag(&cell.xml[start..], "is"))
                .transpose()?
        } else {
            cell.xml
                .find("<v>")
                .map(|start| {
                    let rest = &cell.xml[start..];
                    rest.find("</v>")
                        .map(|end| cell.xml[start..start + end + "</v>".len()].to_owned())
                        .with_context(|| "单元格 <v> 未闭合")
                })
                .transpose()?
        };
        // 共享字符串（t="s"）需要把索引换成 inlineStr 值，否则跨簿索引错位。
        let (value_xml, type_attr) = match (value_xml, cell.type_attr.as_deref()) {
            (Some(value), Some("s")) => {
                let index = value
                    .trim_start_matches("<v>")
                    .trim_end_matches("</v>")
                    .trim();
                let text = shared_strings
                    .get(index.parse::<usize>().unwrap_or(usize::MAX))
                    .cloned()
                    .with_context(|| format!("共享字符串索引 {index} 越界"))?;
                (
                    Some(format!(
                        "<is><t xml:space=\"preserve\">{}</t></is>",
                        escape_xml(&text)
                    )),
                    Some("inlineStr".to_owned()),
                )
            }
            (Some(value), Some(other)) => (Some(value), Some(other.to_owned())),
            (Some(value), None) => (Some(value), None),
            (None, _) => (None, None),
        };
        cells.push(RegionCell {
            row,
            column,
            formula,
            value_xml,
            type_attr,
        });
    }
    Ok(cells)
}

fn strip_trailing_tag(text: &str, tag: &str) -> Result<String> {
    let closing = format!("</{tag}>");
    let end = text.find(&closing).context("内联字符串未闭合")?;
    Ok(text[..end + closing.len()].to_owned())
}

struct CellElement {
    xml: String,
    formula_body: Option<(String, String)>,
    type_attr: Option<String>,
    has_inline_string: bool,
}

fn cell_elements(sheet_xml: &str) -> Vec<CellElement> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = sheet_xml[cursor..].find("<c ") {
        let start = cursor + relative;
        let rest = &sheet_xml[start..];
        let Some(open_end_relative) = rest.find('>') else {
            break;
        };
        let open_end = open_end_relative + 1;
        let opening_tag = &rest[..open_end];
        let end = if opening_tag.trim_end().ends_with("/>") {
            // 只看 `<c .../>` 起始标签是否自闭合；不能被后续单元格的
            // `</c>` 或公式里的 `<f .../>` 误导。
            start + open_end
        } else {
            let Some(close_relative) = rest[open_end..].find("</c>") else {
                break;
            };
            start + open_end + close_relative + "</c>".len()
        };
        result.push(build_cell_element(&sheet_xml[start..end]));
        cursor = end;
    }
    result
}

fn build_cell_element(xml: &str) -> CellElement {
    let formula_body = xml.find("<f").and_then(|start| {
        let rest = &xml[start..];
        let open_end = rest.find('>')?;
        let attrs = rest[..open_end]
            .trim_start_matches("<f")
            .trim_end_matches('/')
            .to_owned();
        if rest[open_end - 1..open_end] == *"/" {
            return Some((attrs, String::new()));
        }
        let body_start = open_end + 1;
        let close = rest[body_start..].find("</f>")?;
        Some((attrs, rest[body_start..body_start + close].to_owned()))
    });
    // 只从 `<c ...>` 单元格起始标签读取类型；公式自己的 `<f t="shared">`
    // 也有 `t` 属性，不能误写成 `<c t="shared">`。
    let cell_tag_end = xml.find('>').unwrap_or(xml.len());
    let cell_tag = &xml[..cell_tag_end];
    // ` t="s"`：跳过 ` t="` 四个字节，取到闭合引号。
    let type_attr = cell_tag.find(" t=\"").and_then(|start| {
        let rest = &cell_tag[start + 4..];
        let end = rest.find('"')?;
        let value = rest[..end].to_owned();
        (!value.trim().is_empty()).then_some(value)
    });
    CellElement {
        has_inline_string: xml.contains("<is>"),
        formula_body,
        type_attr,
        xml: xml.to_owned(),
    }
}

fn cell_reference(cell_xml: &str) -> Option<(u32, u32)> {
    let start = cell_xml.find(" r=\"")? + 4;
    let rest = &cell_xml[start..];
    let end = rest.find('"')?;
    parse_cell(&rest[..end]).ok()
}

/// 主格公式表：si → (公式文本, 覆盖范围)。
fn shared_formula_masters(sheet_xml: &str) -> Result<BTreeMap<String, (String, String)>> {
    let mut masters = BTreeMap::new();
    for cell in cell_elements(sheet_xml) {
        let Some((attrs, body)) = cell.formula_body else {
            continue;
        };
        if !attrs.contains("t=\"shared\"") || body.trim().is_empty() {
            continue;
        }
        masters.insert(
            attribute_value(&attrs, "si").unwrap_or_default(),
            (body, attribute_value(&attrs, "ref").unwrap_or_default()),
        );
    }
    Ok(masters)
}

fn expand_shared_formula(
    masters: &BTreeMap<String, (String, String)>,
    attrs: &str,
    body: &str,
    cell_xml: &str,
) -> Result<String> {
    if !body.trim().is_empty() {
        return Ok(body.to_owned());
    }
    let si = attribute_value(attrs, "si").unwrap_or_default();
    let Some((master_body, master_ref)) = masters.get(&si) else {
        bail!("共享公式依赖格找不到主格公式（si={si}）");
    };
    let master_cell = master_ref
        .split(':')
        .next()
        .context("共享公式主格缺少 ref")?;
    let (master_row, master_column) = parse_cell(master_cell)?;
    let (dest_row, dest_column) = cell_reference(cell_xml).context("共享公式依赖格缺少坐标")?;
    let row_offset = dest_row as i64 - master_row as i64;
    let column_offset = dest_column as i64 - master_column as i64;
    shift_formula_references(master_body, row_offset, column_offset)
}

/// 平移公式中的相对引用；字符串字面量内（含中文表名）不平移。
/// 支持 `$` 绝对引用与跨表 `Sheet!A1` 形式（`!` 前的引用不动）。
fn shift_formula_references(formula: &str, row_offset: i64, column_offset: i64) -> Result<String> {
    if row_offset == 0 && column_offset == 0 {
        return Ok(formula.to_owned());
    }
    let characters: Vec<char> = formula.chars().collect();
    let mut result = String::with_capacity(formula.len());
    let mut index = 0;
    let mut in_string = false;
    while index < characters.len() {
        let current = characters[index];
        if current == '"' {
            in_string = !in_string;
            result.push(current);
            index += 1;
            continue;
        }
        if in_string {
            result.push(current);
            index += 1;
            continue;
        }
        // 只在 ASCII 位置尝试匹配引用；多字节字符直接输出。
        if current.is_ascii_alphabetic() || current == '$' {
            let rest: String = characters[index..].iter().collect();
            if let Some((consumed, replaced)) =
                try_shift_reference(&rest, row_offset, column_offset, &result)?
            {
                result.push_str(&replaced);
                index += consumed;
                continue;
            }
        }
        result.push(current);
        index += 1;
    }
    Ok(result)
}

fn try_shift_reference(
    rest: &str,
    row_offset: i64,
    column_offset: i64,
    prefix: &str,
) -> Result<Option<(usize, String)>> {
    let bytes = rest.as_bytes();
    let mut cursor = 0;
    let column_locked = bytes.first() == Some(&b'$');
    if column_locked {
        cursor += 1;
    }
    let letters_start = cursor;
    while cursor < bytes.len() && (bytes[cursor] as char).is_ascii_alphabetic() {
        cursor += 1;
    }
    if cursor == letters_start || cursor - letters_start > 3 {
        return Ok(None);
    }
    let letters = &rest[letters_start..cursor];
    if !letters.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Ok(None);
    }
    let row_locked = bytes.get(cursor) == Some(&b'$');
    if row_locked {
        cursor += 1;
    }
    let digits_start = cursor;
    while cursor < bytes.len() && (bytes[cursor] as char).is_ascii_digit() {
        cursor += 1;
    }
    if cursor == digits_start {
        return Ok(None);
    }
    // 引用后面紧跟字母/数字/下划线时不是独立引用（可能是名称的一部分）。
    if let Some(next) = bytes.get(cursor) {
        let next = *next as char;
        if next.is_ascii_alphanumeric() || next == '_' || next == '.' {
            return Ok(None);
        }
    }
    // 前面是 `!`（跨表引用）或标识符字符时不平移。
    if prefix.ends_with('!') || prefix.ends_with('$') {
        return Ok(None);
    }
    if let Some(last) = prefix.chars().last() {
        if last.is_ascii_alphanumeric() || last == '_' || last == '.' {
            return Ok(None);
        }
    }
    let column = column_index(letters);
    let row: i64 = rest[digits_start..cursor]
        .parse()
        .with_context(|| format!("解析公式引用行号失败：{rest}"))?;
    let shifted_column = if column_locked {
        column as i64
    } else {
        column as i64 + column_offset
    };
    let shifted_row = if row_locked { row } else { row + row_offset };
    if shifted_column < 1 || shifted_row < 1 {
        return Ok(Some((cursor, "#REF!".to_owned())));
    }
    let mut replacement = String::new();
    if column_locked {
        replacement.push('$');
    }
    replacement.push_str(&column_letters(shifted_column as u32));
    if row_locked {
        replacement.push('$');
    }
    replacement.push_str(&shifted_row.to_string());
    // cursor 是字节偏移；换算成字符数（引用只含 ASCII，前缀可能含多字节）。
    let consumed_chars = rest[..cursor].chars().count();
    Ok(Some((consumed_chars, replacement)))
}

fn column_index(letters: &str) -> u32 {
    letters.chars().fold(0u32, |accumulator, character| {
        let value = character.to_ascii_uppercase() as u32 - 'A' as u32 + 1;
        accumulator * 26 + value
    })
}

fn column_letters(mut column: u32) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        letters.push((b'A' + remainder) as char);
        column = (column - 1) / 26;
    }
    letters.iter().rev().collect()
}

fn attribute_value(attrs: &str, name: &str) -> Option<String> {
    let needle = format!("{name}=\"");
    let start = attrs.find(&needle)? + needle.len();
    let rest = &attrs[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_owned())
}

fn read_shared_strings(package: &Package) -> Result<Vec<String>> {
    let Some(bytes) = package.parts.get("xl/sharedStrings.xml") else {
        return Ok(Vec::new());
    };
    let text = String::from_utf8(bytes.clone()).context("sharedStrings.xml 不是 UTF-8")?;
    let document = roxmltree::Document::parse(&text).context("解析 sharedStrings.xml")?;
    let mut result = Vec::new();
    for item in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "si")
    {
        let mut value = String::new();
        for text_node in item
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "t")
        {
            value.push_str(text_node.text().unwrap_or_default());
        }
        result.push(value);
    }
    Ok(result)
}

fn write_region(
    target_xml: &str,
    cells: &[RegionCell],
    source_bounds: (u32, u32, u32, u32),
    row_offset: i64,
    column_offset: i64,
    flags: &ImportFlags,
) -> Result<String> {
    let (source_lo_col, source_lo_row, source_hi_col, source_hi_row) = source_bounds;
    let mut by_position: BTreeMap<(u32, u32), &RegionCell> = BTreeMap::new();
    for cell in cells {
        by_position.insert((cell.row, cell.column), cell);
    }
    let mut rows: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
    if flags.materialize_blank_cells {
        let existing_targets = cell_elements(target_xml)
            .into_iter()
            .filter_map(|cell| cell_reference(&cell.xml))
            .collect::<BTreeSet<_>>();
        for source_row in source_lo_row..=source_hi_row {
            let target_row = (source_row as i64 + row_offset) as u32;
            for source_column in source_lo_col..=source_hi_col {
                if by_position.contains_key(&(source_row, source_column)) {
                    continue;
                }
                let target_column = (source_column as i64 + column_offset) as u32;
                if existing_targets.contains(&(target_row, target_column)) {
                    let reference = format!("{}{target_row}", column_letters(target_column));
                    rows.entry(target_row)
                        .or_default()
                        .insert(target_column, format!("<c r=\"{reference}\"/>"));
                }
            }
        }
    }
    for ((row, column), cell) in &by_position {
        let target_row = (*row as i64 + row_offset) as u32;
        let target_column = (*column as i64 + column_offset) as u32;
        let reference = format!("{}{target_row}", column_letters(target_column));
        let mut content = String::new();
        if flags.copy_formulas {
            if let Some(formula) = &cell.formula {
                content.push_str(&format!("<f>{}</f>", escape_xml(&unescape_xml(formula))));
            }
        }
        if flags.copy_values || (flags.copy_formulas && cell.formula.is_none()) {
            if let Some(value) = &cell.value_xml {
                content.push_str(value);
            }
        }
        if content.is_empty() && !flags.materialize_blank_cells {
            continue;
        }
        let type_attr = cell
            .type_attr
            .as_deref()
            .filter(|value| !value.trim().is_empty())
            .map(|value| format!(" t=\"{value}\""))
            .unwrap_or_default();
        let cell_xml = if content.is_empty() {
            format!("<c r=\"{reference}\"{type_attr}/>")
        } else {
            format!("<c r=\"{reference}\"{type_attr}>{content}</c>")
        };
        rows.entry(target_row)
            .or_default()
            .insert(target_column, cell_xml);
    }
    write_rows(target_xml, rows)
}

fn write_rows(target_xml: &str, rows: BTreeMap<u32, BTreeMap<u32, String>>) -> Result<String> {
    // 自闭合 `<sheetData/>`（新建工作表）：先展开为成对标签再写入。
    let expanded;
    let target_xml = if let Some(position) = target_xml.find("<sheetData/>") {
        let mut owned = target_xml.to_owned();
        owned.replace_range(
            position..position + "<sheetData/>".len(),
            "<sheetData></sheetData>",
        );
        expanded = owned;
        expanded.as_str()
    } else {
        target_xml
    };
    let (start, end) = sheet_data_bounds(target_xml)?;
    let existing = &target_xml[start..end];
    let mut merged: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
    for row in row_elements(existing) {
        let Some(row_index) =
            attribute_value(&row.attrs, "r").and_then(|value| value.parse::<u32>().ok())
        else {
            continue;
        };
        let mut cells = BTreeMap::new();
        for cell in cell_elements(&row.xml) {
            if let Some((_row, column)) = cell_reference(&cell.xml) {
                cells.insert(column, cell.xml.clone());
            }
        }
        merged.insert(row_index, cells);
    }
    for (row, cells) in rows {
        let entry = merged.entry(row).or_default();
        for (column, cell_xml) in cells {
            entry.insert(column, cell_xml);
        }
    }
    let mut body = String::new();
    for (row, cells) in merged {
        if cells.is_empty() {
            continue;
        }
        body.push_str(&format!("<row r=\"{row}\">"));
        for cell_xml in cells.into_values() {
            body.push_str(&cell_xml);
        }
        body.push_str("</row>");
    }
    let mut result = String::with_capacity(target_xml.len() + body.len());
    result.push_str(&target_xml[..start]);
    result.push_str(&body);
    result.push_str(&target_xml[end..]);
    Ok(result)
}

/// 目标簿 sharedStrings 里追加一个字符串，返回其索引。
/// 区域内单元格若来自模板的共享字符串，必须转成目标簿的索引或 inlineStr；
/// 当前实现统一转 inlineStr，因此本函数只用于将来需要保留 shared 形式时。
fn target_shared_string_index(package: &Package, text: &str) -> Option<usize> {
    let bytes = package.parts.get("xl/sharedStrings.xml")?;
    let xml = String::from_utf8(bytes.clone()).ok()?;
    let document = roxmltree::Document::parse(&xml).ok()?;
    let mut index = 0;
    for item in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "si")
    {
        let mut value = String::new();
        for text_node in item
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "t")
        {
            value.push_str(text_node.text().unwrap_or_default());
        }
        if value == text {
            return Some(index);
        }
        index += 1;
    }
    None
}

struct RowElement {
    attrs: String,
    xml: String,
}

fn row_elements(sheet_xml: &str) -> Vec<RowElement> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = sheet_xml[cursor..].find("<row ") {
        let start = cursor + relative;
        if let Some(close) = sheet_xml[start..].find("</row>") {
            let end = start + close + "</row>".len();
            let xml = &sheet_xml[start..end];
            let attrs_end = xml.find('>').unwrap_or(xml.len());
            result.push(RowElement {
                attrs: xml[5..attrs_end].to_owned(),
                xml: xml.to_owned(),
            });
            cursor = end;
        } else if let Some(self_close) = sheet_xml[start..].find("/>") {
            let end = start + self_close + 2;
            let xml = &sheet_xml[start..end];
            let attrs_end = xml.find("/>").unwrap_or(xml.len());
            result.push(RowElement {
                attrs: xml[5..attrs_end].to_owned(),
                xml: xml.to_owned(),
            });
            cursor = end;
        } else {
            break;
        }
    }
    result
}

/// 返回 `<sheetData>` 内容区间（不含开/闭标签本身）。
/// 自闭合 `<sheetData/>` 的内容区间为空（插入点为标签之前）。
fn sheet_data_bounds(sheet_xml: &str) -> Result<(usize, usize)> {
    let open_tag = sheet_xml
        .find("<sheetData")
        .context("工作表缺少 sheetData")?;
    let open_end = sheet_xml[open_tag..]
        .find('>')
        .map(|offset| open_tag + offset + 1)
        .context("工作表 sheetData 开标签未闭合")?;
    if sheet_xml[open_tag..open_end].trim_end().ends_with("/>") {
        // 自闭合：内容区间为空；调用方需要先展开成 <sheetData></sheetData>。
        return Ok((open_tag, open_tag));
    }
    let close = sheet_xml
        .find("</sheetData>")
        .context("工作表 sheetData 未闭合")?;
    ensure!(open_end <= close, "工作表 sheetData 标签顺序异常");
    Ok((open_end, close))
}

fn parse_cell(reference: &str) -> Result<(u32, u32)> {
    let trimmed = reference.trim();
    let letters_len = trimmed
        .chars()
        .take_while(|character| character.is_ascii_alphabetic())
        .count();
    ensure!(letters_len > 0, "单元格坐标缺少列字母：{reference}");
    let letters = &trimmed[..letters_len];
    let row = trimmed[letters_len..]
        .parse::<u32>()
        .with_context(|| format!("单元格坐标缺少行号：{reference}"))?;
    Ok((row, column_index(letters)))
}

fn parse_range(range: &str) -> Result<(u32, u32, u32, u32)> {
    let (start, end) = match range.split_once(':') {
        Some((start, end)) => (start, end),
        None => (range, range),
    };
    let (start_row, start_column) = parse_cell(start)?;
    let (end_row, end_column) = parse_cell(end)?;
    Ok((
        start_column.min(end_column),
        start_row.min(end_row),
        start_column.max(end_column),
        start_row.max(end_row),
    ))
}

fn target_offsets(source_bounds: (u32, u32, u32, u32), target_start: &str) -> Result<(i64, i64)> {
    let (source_lo_col, source_lo_row, _, _) = source_bounds;
    let (target_row, target_column) = parse_cell(target_start)?;
    Ok((
        target_row as i64 - source_lo_row as i64,
        target_column as i64 - source_lo_col as i64,
    ))
}

fn source_merges(
    sheet_xml: &str,
    bounds: (u32, u32, u32, u32),
) -> Result<Vec<(u32, u32, u32, u32)>> {
    let (lo_col, lo_row, hi_col, hi_row) = bounds;
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = sheet_xml[cursor..].find("<mergeCell ") {
        let start = cursor + relative;
        let Some(end_relative) = sheet_xml[start..].find("/>") else {
            break;
        };
        let end = start + end_relative + 2;
        let element = &sheet_xml[start..end];
        if let Some(reference) = attribute_value(element, "ref") {
            if let Ok((a_col, a_row, b_col, b_row)) = parse_range(&reference) {
                if a_row >= lo_row && b_row <= hi_row && a_col >= lo_col && b_col <= hi_col {
                    result.push((a_col, a_row, b_col, b_row));
                }
            }
        }
        cursor = end;
    }
    Ok(result)
}

fn merge_into_target(
    target_xml: &str,
    merges: &[(u32, u32, u32, u32)],
    row_offset: i64,
    column_offset: i64,
) -> Result<String> {
    let mut additions = Vec::new();
    for (a_col, a_row, b_col, b_row) in merges {
        let shifted_a_col = (*a_col as i64 + column_offset) as u32;
        let shifted_a_row = (*a_row as i64 + row_offset) as u32;
        let shifted_b_col = (*b_col as i64 + column_offset) as u32;
        let shifted_b_row = (*b_row as i64 + row_offset) as u32;
        additions.push(format!(
            "{}{}:{}{}",
            column_letters(shifted_a_col),
            shifted_a_row,
            column_letters(shifted_b_col),
            shifted_b_row
        ));
    }
    let mut xml = target_xml.to_owned();
    for reference in additions {
        if xml.contains(&format!("ref=\"{reference}\"")) {
            continue;
        }
        if let Some(end) = xml.find("</mergeCells>") {
            xml.insert_str(end, &format!("<mergeCell ref=\"{reference}\"/>"));
        } else {
            // 插入到 </sheetData> 之后（sheet_data_bounds 返回的是内容区间）。
            let insertion = sheet_data_bounds(&xml)?.1;
            xml.insert_str(
                insertion,
                &format!("<mergeCells count=\"1\"><mergeCell ref=\"{reference}\"/></mergeCells>"),
            );
        }
    }
    Ok(xml)
}

/// 复制区域内的传统 Excel 批注对象，包括作者、正文、位置和 VML 批注框。
fn copy_region_comments(
    package: &mut Package,
    source_package: &Package,
    source_part: &str,
    target_part: &str,
    bounds: (u32, u32, u32, u32),
    row_offset: i64,
    column_offset: i64,
) -> Result<()> {
    let (lo_col, lo_row, hi_col, hi_row) = bounds;
    let target_lo_row = checked_shift(lo_row, row_offset)?;
    let target_lo_col = checked_shift(lo_col, column_offset)?;
    let target_hi_row = checked_shift(hi_row, row_offset)?;
    let target_hi_col = checked_shift(hi_col, column_offset)?;

    let source_comments = worksheet_relationship_part(source_package, source_part, "/comments")?
        .map(|part| source_package.text(&part))
        .transpose()?;
    let source_model = source_comments
        .as_deref()
        .map(parse_legacy_comments)
        .transpose()?;
    let source_records = source_model
        .as_ref()
        .map(|model| &model.comments[..])
        .unwrap_or_default();
    let mut imported = Vec::new();
    for comment in source_records {
        let (row, column) = parse_cell(&comment.reference)?;
        if row < lo_row || row > hi_row || column < lo_col || column > hi_col {
            continue;
        }
        let shifted_row = checked_shift(row, row_offset)?;
        let shifted_column = checked_shift(column, column_offset)?;
        imported.push((
            format!("{}{shifted_row}", column_letters(shifted_column)),
            comment,
        ));
    }

    let target_comments_part = worksheet_relationship_part(package, target_part, "/comments")?;
    let target_model = target_comments_part
        .as_deref()
        .map(|part| package.text(part))
        .transpose()?
        .as_deref()
        .map(parse_legacy_comments)
        .transpose()?;
    let mut target_comments = target_model
        .as_ref()
        .map(|model| model.comments.clone())
        .unwrap_or_default();
    target_comments.retain(|comment| {
        let Ok((row, column)) = parse_cell(&comment.reference) else {
            return true;
        };
        row < target_lo_row
            || row > target_hi_row
            || column < target_lo_col
            || column > target_hi_col
    });

    let mut authors = target_model
        .as_ref()
        .map(|model| model.authors.clone())
        .unwrap_or_default();
    let mut output_comments = target_comments;
    for (reference, comment) in &imported {
        let source_author = source_model
            .as_ref()
            .and_then(|model| model.authors.get(comment.author_id))
            .with_context(|| {
                format!(
                    "源批注 {} 的 authorId {} 没有对应作者",
                    comment.reference, comment.author_id
                )
            })?;
        let author_id =
            if let Some(existing) = authors.iter().position(|author| author == source_author) {
                existing
            } else {
                authors.push(source_author.clone());
                authors.len() - 1
            };
        let mut xml = set_xml_attribute(&comment.xml, "ref", reference)?;
        xml = set_xml_attribute(&xml, "authorId", &author_id.to_string())?;
        output_comments.push(LegacyComment {
            reference: reference.clone(),
            author_id,
            xml,
        });
    }
    output_comments.sort_by(|left, right| {
        parse_cell(&left.reference)
            .unwrap_or_default()
            .cmp(&parse_cell(&right.reference).unwrap_or_default())
    });
    let mut seen_references = BTreeSet::new();
    for comment in &output_comments {
        ensure!(
            seen_references.insert(comment.reference.to_ascii_uppercase()),
            "复制批注后出现重复位置 {}!{}",
            target_part,
            comment.reference
        );
    }

    let should_write_comments = target_comments_part.is_some() || !output_comments.is_empty();
    if should_write_comments {
        let root_open = target_model
            .as_ref()
            .map(|model| model.root_open.as_str())
            .or_else(|| source_model.as_ref().map(|model| model.root_open.as_str()))
            .unwrap_or(
                "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
            );
        let source_root = source_model.as_ref().map(|model| model.root_open.as_str());
        let comments_xml =
            build_legacy_comments_xml(root_open, source_root, &authors, &output_comments)?;
        let part = match target_comments_part {
            Some(part) => part,
            None => {
                let part = next_comments_part(package)?;
                ensure_worksheet_relationship(package, target_part, "/comments", &part)?;
                append_content_type_override(
                    package,
                    &format!("/{part}"),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                )?;
                part
            }
        };
        package.set_text(&part, comments_xml);
    }

    let source_vml_part = worksheet_relationship_part(source_package, source_part, "/vmlDrawing")?;
    let source_vml = source_vml_part
        .as_deref()
        .map(|part| source_package.text(part))
        .transpose()?;
    let source_shapes = source_vml
        .as_deref()
        .map(parse_vml_note_shapes)
        .transpose()?
        .unwrap_or_default();
    let mut imported_shapes = Vec::new();
    for (reference, _) in &imported {
        let (row, column) = parse_cell(reference)?;
        let source_row = (row as i64 - row_offset) as u32;
        let source_column = (column as i64 - column_offset) as u32;
        let source_reference = format!("{}{source_row}", column_letters(source_column));
        let shape = source_shapes
            .iter()
            .find(|shape| shape.reference.eq_ignore_ascii_case(&source_reference))
            .with_context(|| format!("源批注 {source_reference} 缺少对应 VML 批注框"))?;
        imported_shapes.push(shape);
    }
    let selected_source_refs = imported
        .iter()
        .map(|(_, comment)| comment.reference.to_ascii_uppercase())
        .collect::<BTreeSet<_>>();
    let selected_shape_refs = imported_shapes
        .iter()
        .map(|shape| shape.reference.to_ascii_uppercase())
        .collect::<BTreeSet<_>>();
    ensure!(
        selected_source_refs == selected_shape_refs,
        "源工作表 comments.xml 与 VML 批注框位置不一致"
    );

    let target_vml_part = worksheet_relationship_part(package, target_part, "/vmlDrawing")?;
    let target_vml = target_vml_part
        .as_deref()
        .map(|part| package.text(part))
        .transpose()?;
    let final_refs = output_comments
        .iter()
        .map(|comment| comment.reference.to_ascii_uppercase())
        .collect::<BTreeSet<_>>();
    if target_vml.is_some() || !imported_shapes.is_empty() {
        let mut existing_shapes = target_vml
            .as_deref()
            .map(parse_vml_note_shapes)
            .transpose()?
            .unwrap_or_default();
        existing_shapes.retain(|shape| {
            let Ok((row, column)) = parse_cell(&shape.reference) else {
                return true;
            };
            row < target_lo_row
                || row > target_hi_row
                || column < target_lo_col
                || column > target_hi_col
        });
        let preserved_shape_refs = existing_shapes
            .iter()
            .map(|shape| shape.reference.to_ascii_uppercase())
            .collect::<BTreeSet<_>>();
        let preserved_comment_refs = output_comments
            .iter()
            .filter_map(|comment| {
                let Ok((row, column)) = parse_cell(&comment.reference) else {
                    return None;
                };
                (row < target_lo_row
                    || row > target_hi_row
                    || column < target_lo_col
                    || column > target_hi_col)
                    .then(|| comment.reference.to_ascii_uppercase())
            })
            .collect::<BTreeSet<_>>();
        ensure!(
            preserved_shape_refs == preserved_comment_refs,
            "目标工作表保留区 comments.xml 与 VML 批注框位置不一致"
        );
        let source_root = source_vml
            .as_deref()
            .map(vml_root_opening_tag)
            .transpose()?;
        let mut shape_xml = target_vml
            .as_deref()
            .map(|xml| {
                build_updated_vml(
                    xml,
                    source_vml.as_deref(),
                    &existing_shapes,
                    &imported_shapes,
                    row_offset,
                    column_offset,
                )
            })
            .transpose()?;
        if target_vml.is_none() {
            shape_xml = Some(build_new_vml(
                source_vml.as_deref().context("源批注缺少 VML 绘图")?,
                source_root.as_deref().context("源批注缺少 VML 根节点")?,
                &imported_shapes,
                row_offset,
                column_offset,
            )?);
        }
        let shape_xml = shape_xml.context("构造 S1 审核副本的 VML 批注框失败")?;
        let part = match target_vml_part {
            Some(part) => part,
            None => {
                let part = next_vml_part(package)?;
                let relationship_id =
                    ensure_worksheet_relationship(package, target_part, "/vmlDrawing", &part)?;
                let sheet_xml = package.text(target_part)?;
                package.set_text(
                    target_part,
                    append_legacy_drawing(&sheet_xml, &relationship_id)?,
                );
                ensure_default_content_type(
                    package,
                    "vml",
                    "application/vnd.openxmlformats-officedocument.vmlDrawing",
                )?;
                part
            }
        };
        package.set_text(&part, shape_xml);
        let resulting_shapes = parse_vml_note_shapes(package.text(&part)?.as_str())?;
        let resulting_refs = resulting_shapes
            .iter()
            .map(|shape| shape.reference.to_ascii_uppercase())
            .collect::<BTreeSet<_>>();
        ensure!(
            resulting_refs == final_refs,
            "写入后的目标 VML 批注框位置与 comments.xml 不一致"
        );
    } else {
        ensure!(
            final_refs.is_empty(),
            "目标批注存在但没有可写入的 VML 批注框"
        );
    }
    Ok(())
}

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const VML_NS: &str = "urn:schemas-microsoft-com:vml";
const EXCEL_VML_NS: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const EXCEL_VML_NS_LEGACY: &str = "http://schemas.microsoft.com/office/excel";
const EXCEL_VML_NS_URN: &str = "urn:schemas-microsoft-com:office:excel";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn is_excel_vml_namespace(namespace: Option<&str>) -> bool {
    matches!(
        namespace,
        Some(EXCEL_VML_NS | EXCEL_VML_NS_LEGACY | EXCEL_VML_NS_URN)
    )
}

#[derive(Clone, Debug)]
struct LegacyComment {
    reference: String,
    author_id: usize,
    xml: String,
}

#[derive(Clone, Debug)]
struct LegacyCommentsModel {
    root_open: String,
    authors: Vec<String>,
    comments: Vec<LegacyComment>,
}

#[derive(Clone, Debug)]
struct VmlNoteShape {
    reference: String,
    xml: String,
}

fn parse_legacy_comments(xml: &str) -> Result<LegacyCommentsModel> {
    let document = roxmltree::Document::parse(xml).context("解析传统 Excel comments.xml")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(MAIN_NS) && root.tag_name().name() == "comments",
        "comments.xml 根节点不是 SpreadsheetML comments"
    );
    let root_open = element_opening_tag(xml, root.range().start)?;
    let authors = root
        .children()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "authors"
        })
        .into_iter()
        .flat_map(|authors| authors.children())
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "author"
        })
        .map(|author| author.text().unwrap_or_default().to_owned())
        .collect::<Vec<_>>();
    let comment_list = root
        .children()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "commentList"
        })
        .context("comments.xml 缺少 commentList")?;
    let mut comments = Vec::new();
    for node in comment_list.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "comment"
    }) {
        let reference = node
            .attribute("ref")
            .context("comments.xml 中 comment 缺少 ref")?
            .to_owned();
        parse_cell(&reference).context("comments.xml 中 comment 的 ref 无效")?;
        let author_id = node
            .attribute("authorId")
            .context("comments.xml 中 comment 缺少 authorId")?
            .parse::<usize>()
            .with_context(|| format!("批注 {reference} 的 authorId 无效"))?;
        ensure!(
            authors.get(author_id).is_some(),
            "批注 {reference} 的 authorId {author_id} 超出 authors 表"
        );
        let range = node.range();
        comments.push(LegacyComment {
            reference,
            author_id,
            xml: xml[range].to_owned(),
        });
    }
    Ok(LegacyCommentsModel {
        root_open,
        authors,
        comments,
    })
}

fn build_legacy_comments_xml(
    target_root_open: &str,
    source_root_open: Option<&str>,
    authors: &[String],
    comments: &[LegacyComment],
) -> Result<String> {
    let mut root_open = target_root_open.to_owned();
    if let Some(source_root_open) = source_root_open {
        root_open = merge_namespace_declarations(&root_open, source_root_open)?;
    }
    let root_qname = tag_qualified_name(&root_open)?;
    let prefix = root_qname
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default();
    let authors_tag = format!("{prefix}authors");
    let author_tag = format!("{prefix}author");
    let comment_list_tag = format!("{prefix}commentList");
    let root_tag = root_qname;
    let mut output = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.push_str(&root_open);
    output.push_str(&format!("<{authors_tag}>"));
    for author in authors {
        output.push_str(&format!(
            "<{author_tag}>{}</{author_tag}>",
            escape_xml(author)
        ));
    }
    output.push_str(&format!("</{authors_tag}><{comment_list_tag}>"));
    for comment in comments {
        output.push_str(&comment.xml);
    }
    output.push_str(&format!("</{comment_list_tag}></{root_tag}>"));
    Ok(output)
}

fn parse_vml_note_shapes(xml: &str) -> Result<Vec<VmlNoteShape>> {
    let document = roxmltree::Document::parse(xml).context("解析工作表 VML 批注绘图")?;
    let mut shapes = Vec::new();
    let mut references = BTreeSet::new();
    for shape in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(VML_NS)
            && node.tag_name().name() == "shape"
    }) {
        let Some(client_data) = shape.descendants().find(|node| {
            node.is_element()
                && is_excel_vml_namespace(node.tag_name().namespace())
                && node.tag_name().name() == "ClientData"
                && node.attribute("ObjectType") == Some("Note")
        }) else {
            continue;
        };
        let row = vml_client_data_number(client_data, "Row")?;
        let column = vml_client_data_number(client_data, "Column")?;
        let row = row + 1;
        let reference = format!("{}{row}", column_letters(column + 1));
        ensure!(
            references.insert(reference.to_ascii_uppercase()),
            "VML 中存在重复的 Note 位置 {reference}"
        );
        shapes.push(VmlNoteShape {
            reference,
            xml: xml[shape.range()].to_owned(),
        });
    }
    Ok(shapes)
}

fn vml_client_data_number(client_data: roxmltree::Node<'_, '_>, name: &str) -> Result<u32> {
    client_data
        .children()
        .find(|node| {
            node.is_element()
                && is_excel_vml_namespace(node.tag_name().namespace())
                && node.tag_name().name() == name
        })
        .and_then(|node| node.text())
        .with_context(|| format!("VML ClientData 缺少 {name}"))?
        .trim()
        .parse::<u32>()
        .with_context(|| format!("VML ClientData {name} 不是有效坐标"))
}

fn build_new_vml(
    source_xml: &str,
    root_open: &str,
    imported_shapes: &[&VmlNoteShape],
    row_offset: i64,
    column_offset: i64,
) -> Result<String> {
    let document = roxmltree::Document::parse(source_xml).context("解析源工作表 VML")?;
    let root = document.root_element();
    let root_qname = tag_qualified_name(&element_opening_tag(source_xml, root.range().start)?)?;
    let mut children = String::new();
    for child in root.children().filter(|node| node.is_element()) {
        let is_layout = child.tag_name().name() == "shapelayout";
        let is_shape_type =
            child.tag_name().namespace() == Some(VML_NS) && child.tag_name().name() == "shapetype";
        if is_layout || is_shape_type {
            children.push_str(&source_xml[child.range()]);
        }
    }
    let mut next_id = 1025;
    for shape in imported_shapes {
        children.push_str(&transform_vml_note_shape(
            &shape.xml,
            root_open,
            row_offset,
            column_offset,
            &format!("_x0000_s{next_id}"),
        )?);
        next_id += 1;
    }
    Ok(format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{root_open}{children}</{root_qname}>"
    ))
}

fn build_updated_vml(
    target_xml: &str,
    source_xml: Option<&str>,
    preserved_shapes: &[VmlNoteShape],
    imported_shapes: &[&VmlNoteShape],
    row_offset: i64,
    column_offset: i64,
) -> Result<String> {
    let target_document = roxmltree::Document::parse(target_xml).context("解析目标工作表 VML")?;
    let target_root = target_document.root_element();
    let mut output = target_xml.to_owned();
    let mut ranges = target_root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(VML_NS)
                && node.tag_name().name() == "shape"
                && node.descendants().any(|client_data| {
                    client_data.is_element()
                        && is_excel_vml_namespace(client_data.tag_name().namespace())
                        && client_data.tag_name().name() == "ClientData"
                        && client_data.attribute("ObjectType") == Some("Note")
                })
        })
        .filter_map(|node| {
            let cell_data = node.descendants().find(|client_data| {
                client_data.is_element()
                    && is_excel_vml_namespace(client_data.tag_name().namespace())
                    && client_data.tag_name().name() == "ClientData"
                    && client_data.attribute("ObjectType") == Some("Note")
            })?;
            let row = vml_client_data_number(cell_data, "Row")
                .ok()?
                .checked_add(1)?;
            let column = vml_client_data_number(cell_data, "Column")
                .ok()?
                .checked_add(1)?;
            let address = format!("{}{row}", column_letters(column));
            (!preserved_shapes
                .iter()
                .any(|shape| shape.reference.eq_ignore_ascii_case(&address)))
            .then_some(node.range())
        })
        .collect::<Vec<_>>();
    ranges.sort_by_key(|range| std::cmp::Reverse(range.start));
    for range in ranges {
        output.replace_range(range, "");
    }

    let mut missing_shape_types = String::new();
    let target_type_ids = roxmltree::Document::parse(&output)?
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(VML_NS)
                && node.tag_name().name() == "shapetype"
        })
        .filter_map(|node| node.attribute("id"))
        .map(str::to_owned)
        .collect::<BTreeSet<_>>();
    if let Some(source_xml) = source_xml {
        let source_document = roxmltree::Document::parse(source_xml).context("解析源工作表 VML")?;
        let source_root = source_document.root_element();
        for shape_type in source_root.children().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(VML_NS)
                && node.tag_name().name() == "shapetype"
        }) {
            let id = shape_type.attribute("id").unwrap_or_default();
            if !target_type_ids.contains(id) {
                missing_shape_types.push_str(&source_xml[shape_type.range()]);
            }
        }
    }

    let max_shape_id = roxmltree::Document::parse(&output)?
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(VML_NS)
                && node.tag_name().name() == "shape"
        })
        .filter_map(|node| node.attribute("id"))
        .filter_map(|id| {
            id.rsplit_once('s')
                .and_then(|(_, digits)| digits.parse::<u64>().ok())
        })
        .max()
        .unwrap_or(1024);
    let mut next_id = max_shape_id.saturating_add(1);
    let mut imported_xml = String::new();
    if !imported_shapes.is_empty() {
        let source_xml = source_xml.context("复制 VML 批注时缺少源绘图")?;
        let source_root_open = source_root_opening_tag(source_xml)?;
        for shape in imported_shapes {
            imported_xml.push_str(&transform_vml_note_shape(
                &shape.xml,
                &source_root_open,
                row_offset,
                column_offset,
                &format!("_x0000_s{next_id}"),
            )?);
            next_id = next_id.saturating_add(1);
        }
    }

    let current_root_open = vml_root_opening_tag(&output)?;
    let merged_root_open = if let Some(source_xml) = source_xml {
        merge_namespace_declarations(&current_root_open, &source_root_opening_tag(source_xml)?)?
    } else {
        current_root_open
    };
    output = replace_root_opening_tag(&output, &merged_root_open)?;
    append_before_root_close(&mut output, &(missing_shape_types + &imported_xml))?;
    Ok(output)
}

fn transform_vml_note_shape(
    shape_xml: &str,
    namespace_context: &str,
    row_offset: i64,
    column_offset: i64,
    new_id: &str,
) -> Result<String> {
    let shape = set_xml_attribute(shape_xml, "id", new_id)?;
    let namespace_attributes = xml_namespace_attributes(namespace_context)?;
    let wrapped = format!("<wrapper{namespace_attributes}>{shape}</wrapper>");
    let document = roxmltree::Document::parse(&wrapped).context("解析单个 VML 批注框")?;
    let mut edits = Vec::new();
    let client_data = document
        .descendants()
        .find(|node| {
            node.is_element()
                && is_excel_vml_namespace(node.tag_name().namespace())
                && node.tag_name().name() == "ClientData"
                && node.attribute("ObjectType") == Some("Note")
        })
        .context("VML shape 缺少 Note ClientData")?;
    let row_node = client_data
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "Row")
        .context("VML Note 缺少 Row")?;
    let column_node = client_data
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "Column")
        .context("VML Note 缺少 Column")?;
    for (node, offset, name) in [
        (row_node, row_offset, "Row"),
        (column_node, column_offset, "Column"),
    ] {
        let current = node
            .text()
            .context("VML Note 坐标缺少内容")?
            .trim()
            .parse::<i64>()
            .with_context(|| format!("VML Note {name} 坐标无效"))?;
        let shifted = current
            .checked_add(offset)
            .filter(|value| *value >= 0)
            .context("平移后的 VML Note 坐标超出工作表范围")?;
        edits.push(
            xml_node_text_range(&wrapped, node.range())?.map(|range| (range, shifted.to_string())),
        );
    }
    if let Some(anchor) = client_data.children().find(|node| {
        node.is_element()
            && is_excel_vml_namespace(node.tag_name().namespace())
            && node.tag_name().name() == "Anchor"
    }) {
        let fields = anchor
            .text()
            .unwrap_or_default()
            .split(',')
            .map(str::trim)
            .map(str::to_owned)
            .collect::<Vec<_>>();
        if fields.len() == 8 {
            let mut shifted = fields.clone();
            for index in [0usize, 4] {
                shifted[index] =
                    checked_shift_zero_based(fields[index].parse::<u32>()?, column_offset)?
                        .to_string();
            }
            for index in [2usize, 6] {
                shifted[index] =
                    checked_shift_zero_based(fields[index].parse::<u32>()?, row_offset)?
                        .to_string();
            }
            edits.push(
                xml_node_text_range(&wrapped, anchor.range())?
                    .map(|range| (range, shifted.join(", "))),
            );
        } else {
            bail!("VML Note Anchor 必须包含 8 个坐标字段");
        }
    }
    let edits = edits.into_iter().flatten().collect::<Vec<_>>();
    let mut updated = wrapped;
    let mut edits = edits;
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        updated.replace_range(range, &replacement);
    }
    let open_end = updated.find('>').context("VML 包装根节点未闭合")? + 1;
    let close_start = updated
        .rfind("</wrapper>")
        .context("VML 包装根节点缺少结束标签")?;
    Ok(updated[open_end..close_start].to_owned())
}

fn xml_node_text_range(
    xml: &str,
    range: std::ops::Range<usize>,
) -> Result<Option<std::ops::Range<usize>>> {
    let element = &xml[range.clone()];
    let open_end = element.find('>').context("XML 节点开始标签未闭合")? + range.start + 1;
    let close_start = element.rfind("</").map(|index| range.start + index);
    Ok(close_start.map(|close_start| open_end..close_start))
}

fn source_root_opening_tag(xml: &str) -> Result<String> {
    let document = roxmltree::Document::parse(xml)?;
    element_opening_tag(xml, document.root_element().range().start)
}

fn vml_root_opening_tag(xml: &str) -> Result<String> {
    let document = roxmltree::Document::parse(xml).context("解析 VML 根节点")?;
    ensure!(
        document.root_element().tag_name().name() == "xml",
        "VML 绘图根节点必须为 xml"
    );
    element_opening_tag(xml, document.root_element().range().start)
}

fn append_before_root_close(xml: &mut String, content: &str) -> Result<()> {
    if content.is_empty() {
        return Ok(());
    }
    let document = roxmltree::Document::parse(xml)?;
    let root = document.root_element();
    let qname = tag_qualified_name(&element_opening_tag(xml, root.range().start)?)?;
    let closing_tag = format!("</{qname}>");
    let close_start = xml.rfind(&closing_tag).context("VML 根节点缺少结束标签")?;
    xml.insert_str(close_start, content);
    Ok(())
}

fn replace_root_opening_tag(xml: &str, replacement: &str) -> Result<String> {
    let document = roxmltree::Document::parse(xml)?;
    let start = document.root_element().range().start;
    let old = element_opening_tag(xml, start)?;
    Ok(format!(
        "{}{replacement}{}",
        &xml[..start],
        &xml[start + old.len()..]
    ))
}

fn merge_namespace_declarations(target: &str, source: &str) -> Result<String> {
    let target_attributes = xml_start_tag_attributes(target)?;
    let source_attributes = xml_start_tag_attributes(source)?;
    let mut merged = target.to_owned();
    for (name, value) in source_attributes {
        if name != "xmlns" && !name.starts_with("xmlns:") {
            continue;
        }
        if let Some(existing) = target_attributes.get(&name) {
            ensure!(
                existing == &value,
                "批注 XML 前缀 {name} 在源与目标中指向不同命名空间"
            );
        } else {
            merged = set_xml_attribute(&merged, &name, &value)?;
        }
    }
    Ok(merged)
}

fn xml_namespace_attributes(root_open: &str) -> Result<String> {
    Ok(xml_start_tag_attributes(root_open)?
        .into_iter()
        .filter(|(name, _)| name == "xmlns" || name.starts_with("xmlns:"))
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_xml(&value)))
        .collect())
}

fn xml_start_tag_attributes(tag: &str) -> Result<BTreeMap<String, String>> {
    let open_end = find_start_tag_end(tag, 0)?;
    let bytes = tag.as_bytes();
    let mut index = tag[..open_end].find('<').context("XML 开始标签缺少 <")? + 1;
    while index < open_end
        && !bytes[index].is_ascii_whitespace()
        && bytes[index] != b'>'
        && bytes[index] != b'/'
    {
        index += 1;
    }
    let mut attributes = BTreeMap::new();
    while index < open_end {
        while index < open_end && (bytes[index].is_ascii_whitespace() || bytes[index] == b'/') {
            index += 1;
        }
        if index >= open_end || bytes[index] == b'>' {
            break;
        }
        let name_start = index;
        while index < open_end
            && !bytes[index].is_ascii_whitespace()
            && bytes[index] != b'='
            && bytes[index] != b'>'
            && bytes[index] != b'/'
        {
            index += 1;
        }
        if name_start == index {
            bail!("无法解析 XML 开始标签属性");
        }
        let name = &tag[name_start..index];
        while index < open_end && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        ensure!(
            index < open_end && bytes[index] == b'=',
            "XML 属性 {name} 缺少等号"
        );
        index += 1;
        while index < open_end && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        ensure!(
            index < open_end && (bytes[index] == b'\"' || bytes[index] == b'\''),
            "XML 属性 {name} 未加引号"
        );
        let quote = bytes[index];
        index += 1;
        let value_start = index;
        while index < open_end && bytes[index] != quote {
            index += 1;
        }
        ensure!(index < open_end, "XML 属性 {name} 引号未闭合");
        let value = unescape_xml(&tag[value_start..index]);
        attributes.insert(name.to_owned(), value);
        index += 1;
    }
    Ok(attributes)
}

fn find_start_tag_end(xml: &str, start: usize) -> Result<usize> {
    let bytes = xml.as_bytes();
    let mut quote = None;
    for index in start..bytes.len() {
        match (quote, bytes[index]) {
            (Some(current), byte) if byte == current => quote = None,
            (Some(_), _) => {}
            (None, b'\"' | b'\'') => quote = Some(bytes[index]),
            (None, b'>') => return Ok(index),
            _ => {}
        }
    }
    bail!("XML 开始标签缺少 >")
}

fn element_opening_tag(xml: &str, start: usize) -> Result<String> {
    let end = find_start_tag_end(xml, start)?;
    Ok(xml[start..=end].to_owned())
}

fn tag_qualified_name(opening_tag: &str) -> Result<String> {
    let start = opening_tag.find('<').context("XML 开始标签缺少 <")? + 1;
    let end = opening_tag[start..]
        .find(|character: char| character.is_whitespace() || character == '/' || character == '>')
        .map(|offset| start + offset)
        .unwrap_or(opening_tag.len());
    ensure!(end > start, "XML 开始标签缺少名称");
    Ok(opening_tag[start..end].to_owned())
}

fn set_xml_attribute(xml: &str, name: &str, value: &str) -> Result<String> {
    let end = find_start_tag_end(xml, 0)?;
    let start_tag = &xml[..=end];
    let attributes = xml_start_tag_attribute_ranges(start_tag)?;
    if let Some((value_start, value_end, quote)) = attributes.get(name) {
        let mut result = xml.to_owned();
        result.replace_range(
            *value_start..*value_end,
            &escape_xml_attribute(value, *quote),
        );
        return Ok(result);
    }
    let insertion = if start_tag[..end].trim_end().ends_with('/') {
        end - 1
    } else {
        end
    };
    let mut result = xml.to_owned();
    result.insert_str(
        insertion,
        &format!(" {name}=\"{}\"", escape_xml_attribute(value, b'\"')),
    );
    Ok(result)
}

fn xml_start_tag_attribute_ranges(tag: &str) -> Result<BTreeMap<String, (usize, usize, u8)>> {
    let end = find_start_tag_end(tag, 0)?;
    let bytes = tag.as_bytes();
    let mut index = tag.find('<').context("XML 开始标签缺少 <")? + 1;
    while index < end
        && !bytes[index].is_ascii_whitespace()
        && bytes[index] != b'>'
        && bytes[index] != b'/'
    {
        index += 1;
    }
    let mut attributes = BTreeMap::new();
    while index < end {
        while index < end && (bytes[index].is_ascii_whitespace() || bytes[index] == b'/') {
            index += 1;
        }
        if index >= end || bytes[index] == b'>' {
            break;
        }
        let name_start = index;
        while index < end
            && !bytes[index].is_ascii_whitespace()
            && bytes[index] != b'='
            && bytes[index] != b'>'
            && bytes[index] != b'/'
        {
            index += 1;
        }
        let name = tag[name_start..index].to_owned();
        while index < end && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        ensure!(
            index < end && bytes[index] == b'=',
            "XML 属性 {name} 缺少等号"
        );
        index += 1;
        while index < end && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        ensure!(
            index < end && (bytes[index] == b'\"' || bytes[index] == b'\''),
            "XML 属性 {name} 未加引号"
        );
        let quote = bytes[index];
        index += 1;
        let value_start = index;
        while index < end && bytes[index] != quote {
            index += 1;
        }
        ensure!(index < end, "XML 属性 {name} 引号未闭合");
        attributes.insert(name, (value_start, index, quote));
        index += 1;
    }
    Ok(attributes)
}

fn escape_xml_attribute(value: &str, quote: u8) -> String {
    let value = value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");
    if quote == b'\'' {
        value.replace('\'', "&apos;")
    } else {
        value.replace('\"', "&quot;")
    }
}

fn checked_shift(value: u32, offset: i64) -> Result<u32> {
    (value as i64)
        .checked_add(offset)
        .filter(|shifted| *shifted > 0 && *shifted <= 1_048_576)
        .and_then(|shifted| u32::try_from(shifted).ok())
        .context("平移后的批注位置超出工作表坐标范围")
}

fn checked_shift_zero_based(value: u32, offset: i64) -> Result<u32> {
    (value as i64)
        .checked_add(offset)
        .filter(|shifted| *shifted >= 0 && *shifted < 1_048_576)
        .and_then(|shifted| u32::try_from(shifted).ok())
        .context("平移后的 VML Anchor 坐标超出工作表范围")
}

fn worksheet_relationship_part(
    package: &Package,
    sheet_part: &str,
    suffix: &str,
) -> Result<Option<String>> {
    let rels = worksheet_rels_part(sheet_part);
    let Some(bytes) = package.parts.get(&rels) else {
        return Ok(None);
    };
    let xml =
        std::str::from_utf8(bytes).with_context(|| format!("工作表 rels {rels} 不是 UTF-8"))?;
    relationship_id_and_target_by_type(xml, suffix)?
        .map(|(_, target)| resolve_worksheet_target(sheet_part, &target))
        .map(Some)
        .map(Ok)
        .unwrap_or(Ok(None))
}

fn ensure_worksheet_relationship(
    package: &mut Package,
    sheet_part: &str,
    suffix: &str,
    target_part: &str,
) -> Result<String> {
    let rels_part = worksheet_rels_part(sheet_part);
    let mut rels_xml = package.parts.get(&rels_part)
        .map(|bytes| String::from_utf8(bytes.clone()).context("目标工作表 rels 不是 UTF-8"))
        .transpose()?
        .unwrap_or_else(|| "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_owned());
    if let Some((id, target)) = relationship_id_and_target_by_type(&rels_xml, suffix)? {
        ensure!(
            resolve_worksheet_target(sheet_part, &target) == target_part,
            "目标工作表已有 {suffix} 关系，但指向的部件不一致"
        );
        return Ok(id);
    }
    let id = next_relationship_id(&rels_xml)?;
    let relationship_type = if suffix.starts_with('/') {
        format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships{suffix}")
    } else {
        suffix.to_owned()
    };
    let relative_target = relative_from_worksheets(target_part);
    let element = format!(
        "<Relationship Id=\"{id}\" Type=\"{relationship_type}\" Target=\"{relative_target}\"/>"
    );
    let insertion = rels_xml
        .rfind("</Relationships>")
        .context("工作表 rels 缺少结束标签")?;
    rels_xml.insert_str(insertion, &element);
    package.set_text(&rels_part, rels_xml);
    Ok(id)
}

fn relationship_id_and_target_by_type(xml: &str, suffix: &str) -> Result<Option<(String, String)>> {
    let document = roxmltree::Document::parse(xml).context("解析工作表 rels")?;
    let matches = document
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "Relationship"
                && node
                    .attribute("Type")
                    .is_some_and(|value| value.ends_with(suffix))
        })
        .collect::<Vec<_>>();
    ensure!(matches.len() <= 1, "工作表存在多个 {suffix} 关系");
    matches
        .first()
        .map(|node| {
            let id = node
                .attribute("Id")
                .context("Relationship 缺少 Id")?
                .to_owned();
            let target = node
                .attribute("Target")
                .context("Relationship 缺少 Target")?
                .to_owned();
            Ok((id, target))
        })
        .transpose()
}

fn next_comments_part(package: &Package) -> Result<String> {
    (1..=u32::MAX)
        .map(|index| format!("xl/comments{index}.xml"))
        .find(|part| !package.parts.contains_key(part))
        .context("无法分配唯一 comments.xml 部件名")
}

fn next_vml_part(package: &Package) -> Result<String> {
    (1..=u32::MAX)
        .map(|index| format!("xl/drawings/vmlDrawing{index}.vml"))
        .find(|part| !package.parts.contains_key(part))
        .context("无法分配唯一 VML 部件名")
}

fn ensure_default_content_type(
    package: &mut Package,
    extension: &str,
    content_type: &str,
) -> Result<()> {
    let path = "[Content_Types].xml";
    let xml = package.text(path)?;
    let document = roxmltree::Document::parse(&xml).context("解析 [Content_Types].xml")?;
    for default in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Default")
    {
        if default
            .attribute("Extension")
            .is_some_and(|value| value.eq_ignore_ascii_case(extension))
        {
            ensure!(
                default.attribute("ContentType") == Some(content_type),
                "扩展名 {extension} 已登记了不同的 ContentType"
            );
            return Ok(());
        }
    }
    let insertion = xml
        .rfind("</Types>")
        .context("[Content_Types].xml 缺少 Types 结束标签")?;
    let mut updated = xml;
    updated.insert_str(
        insertion,
        &format!("<Default Extension=\"{extension}\" ContentType=\"{content_type}\"/>"),
    );
    package.set_text(path, updated);
    Ok(())
}

fn append_legacy_drawing(sheet_xml: &str, relationship_id: &str) -> Result<String> {
    let document = roxmltree::Document::parse(sheet_xml).context("解析目标工作表 XML")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(MAIN_NS),
        "目标部件不是 SpreadsheetML 工作表"
    );
    ensure!(
        !root.descendants().any(|node| node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "legacyDrawing"),
        "工作表已有 legacyDrawing 元素，但缺少对应 VML 关系"
    );
    let root_open = element_opening_tag(sheet_xml, root.range().start)?;
    let root_qname = tag_qualified_name(&root_open)?;
    let prefix = root_qname
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default();
    let mut updated = if xml_start_tag_attributes(&root_open)?.contains_key("xmlns:r") {
        sheet_xml.to_owned()
    } else {
        replace_root_opening_tag(
            sheet_xml,
            &set_xml_attribute(&root_open, "xmlns:r", OFFICE_REL_NS)?,
        )?
    };
    let document = roxmltree::Document::parse(&updated)?;
    let root = document.root_element();
    let ext_list = root
        .children()
        .find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "extLst"
        })
        .map(|node| node.range().start);
    let insertion = if let Some(insertion) = ext_list {
        insertion
    } else {
        let root_open = element_opening_tag(&updated, root.range().start)?;
        let root_qname = tag_qualified_name(&root_open)?;
        updated
            .rfind(&format!("</{root_qname}>"))
            .context("无法定位 worksheet legacyDrawing 插入点")?
    };
    updated.insert_str(
        insertion,
        &format!("<{prefix}legacyDrawing r:id=\"{relationship_id}\"/>"),
    );
    Ok(updated)
}

fn worksheet_rels_part(sheet_part: &str) -> String {
    let (directory, file) = sheet_part
        .rsplit_once('/')
        .unwrap_or(("xl/worksheets", sheet_part));
    format!("{directory}/_rels/{file}.rels")
}

fn resolve_worksheet_target(sheet_part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_owned();
    }
    let directory = sheet_part
        .rsplit_once('/')
        .map(|(directory, _)| directory)
        .unwrap_or("xl/worksheets");
    let mut segments: Vec<&str> = directory.split('/').collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                segments.pop();
            }
            "." | "" => {}
            other => segments.push(other),
        }
    }
    segments.join("/")
}

fn relative_from_worksheets(part: &str) -> String {
    part.strip_prefix("xl/")
        .map(|rest| format!("../{rest}"))
        .unwrap_or_else(|| part.to_owned())
}

// ---------------------------------------------------------------------------
// 测试
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    #[test]
    fn cell_reference_parsing_handles_absolute_and_relative_forms() {
        assert_eq!(parse_cell("A1").unwrap(), (1, 1));
        assert_eq!(parse_cell("C4").unwrap(), (4, 3));
        assert_eq!(parse_cell("AA10").unwrap(), (10, 27));
        assert_eq!(
            parse_range("B2:D5").unwrap(),
            (2, 2, 4, 5),
            "range returns (lo_col, lo_row, hi_col, hi_row)"
        );
    }

    #[test]
    fn target_offsets_keep_row_and_column_coordinates_in_order() {
        let source = parse_range("K7:V133").unwrap();
        assert_eq!(target_offsets(source, "K7").unwrap(), (0, 0));
        assert_eq!(target_offsets(source, "D10").unwrap(), (3, -7));
    }

    #[test]
    fn missing_blank_template_cells_clear_existing_target_content() {
        let target = concat!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
            r#"<row r="1"><c r="A1"><v>old</v></c><c r="B1"><f>SUM(A1)</f><v>99</v></c></row>"#,
            "</sheetData></worksheet>"
        );
        let source_cells = vec![RegionCell {
            row: 1,
            column: 1,
            formula: None,
            value_xml: Some("<v>new</v>".to_owned()),
            type_attr: None,
        }];
        let output = write_region(
            target,
            &source_cells,
            (1, 1, 2, 1),
            0,
            0,
            &ImportFlags {
                copy_values: true,
                copy_formulas: true,
                copy_comments: false,
                copy_merges: false,
                materialize_blank_cells: true,
            },
        )
        .unwrap();
        let cells = cell_elements(&output);
        let copied = cells
            .iter()
            .find(|cell| cell_reference(&cell.xml) == Some((1, 1)))
            .unwrap();
        assert_eq!(copied.xml, r#"<c r="A1"><v>new</v></c>"#);
        let cleared = cells
            .iter()
            .find(|cell| cell_reference(&cell.xml) == Some((1, 2)))
            .unwrap();
        assert!(cleared.formula_body.is_none());
        assert!(!cleared.xml.contains("<v>"));
        assert_eq!(cleared.xml, r#"<c r="B1"/>"#);
    }

    #[test]
    fn copies_comment_author_text_position_and_vml_anchor_while_preserving_outside_notes() {
        let mut target = Package {
            order: vec![
                "[Content_Types].xml".to_owned(),
                "xl/worksheets/sheet1.xml".to_owned(),
                "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                "xl/comments1.xml".to_owned(),
                "xl/drawings/vmlDrawing1.vml".to_owned(),
            ],
            parts: BTreeMap::from([
                (
                    "[Content_Types].xml".to_owned(),
                    br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#.to_vec(),
                ),
                (
                    "xl/worksheets/sheet1.xml".to_owned(),
                    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rId2"/></worksheet>"#.to_vec(),
                ),
                (
                    "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid" TargetMode="External"/></Relationships>"#.to_vec(),
                ),
                (
                    "xl/comments1.xml".to_owned(),
                    br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Source User</author></authors><commentList><comment ref="A2" authorId="0"><text><t>old target note</t></text></comment><comment ref="C1" authorId="0"><text><t>keep this note</t></text></comment></commentList></comments>"#.to_vec(),
                ),
                (
                    "xl/drawings/vmlDrawing1.vml".to_owned(),
                    test_vml(&[
                        ("_x0000_s1025", "A2", 1, 0, "0, 15, 1, 2, 2, 15, 5, 2"),
                        ("_x0000_s1026", "C1", 0, 2, "2, 15, 0, 2, 4, 15, 3, 2"),
                    ])
                    .into_bytes(),
                ),
            ]),
        };
        let source = Package {
            order: vec![
                "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                "xl/comments1.xml".to_owned(),
                "xl/drawings/vmlDrawing1.vml".to_owned(),
            ],
            parts: BTreeMap::from([
                (
                    "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>"#.to_vec(),
                ),
                (
                    "xl/comments1.xml".to_owned(),
                    br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Template Author</author></authors><commentList><comment ref="B1" authorId="0"><text><t>template note</t></text></comment></commentList></comments>"#.to_vec(),
                ),
                (
                    "xl/drawings/vmlDrawing1.vml".to_owned(),
                    test_vml(&[("_x0000_s1025", "B1", 0, 1, "1, 15, 0, 2, 3, 15, 4, 2")])
                        .into_bytes(),
                ),
            ]),
        };

        copy_region_comments(
            &mut target,
            &source,
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet1.xml",
            (2, 1, 3, 1),
            1,
            -1,
        )
        .expect("copy notes from B1:C1 into A2:B2");

        let comments = parse_legacy_comments(&target.text("xl/comments1.xml").unwrap()).unwrap();
        assert_eq!(comments.authors, vec!["Source User", "Template Author"]);
        let by_reference = comments
            .comments
            .iter()
            .map(|comment| (comment.reference.as_str(), comment))
            .collect::<BTreeMap<_, _>>();
        assert!(by_reference.contains_key("A2"));
        assert!(by_reference.contains_key("C1"));
        assert_eq!(by_reference["A2"].author_id, 1);
        assert!(by_reference["A2"].xml.contains("template note"));
        assert!(!by_reference["A2"].xml.contains("old target note"));
        assert!(by_reference["C1"].xml.contains("keep this note"));

        let vml = target.text("xl/drawings/vmlDrawing1.vml").unwrap();
        let shapes = parse_vml_note_shapes(&vml).unwrap();
        assert_eq!(
            shapes
                .iter()
                .map(|shape| shape.reference.as_str())
                .collect::<BTreeSet<_>>(),
            BTreeSet::from(["A2", "C1"])
        );
        assert!(vml.contains("id=\"_x0000_s1027\""));
        assert!(vml.contains("<x:Anchor>0, 15, 1, 2, 2, 15, 5, 2</x:Anchor>"));
        assert!(
            target
                .text("xl/worksheets/_rels/sheet1.xml.rels")
                .unwrap()
                .contains("https://example.invalid"),
            "copying notes must preserve non-comment worksheet relationships"
        );
    }

    #[test]
    fn creates_comment_parts_relationships_and_legacy_drawing_for_clean_target_sheet() {
        let mut target = Package {
            order: vec![
                "[Content_Types].xml".to_owned(),
                "xl/worksheets/sheet1.xml".to_owned(),
            ],
            parts: BTreeMap::from([
                (
                    "[Content_Types].xml".to_owned(),
                    br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#.to_vec(),
                ),
                (
                    "xl/worksheets/sheet1.xml".to_owned(),
                    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
                ),
            ]),
        };
        let source = Package {
            order: vec![
                "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                "xl/comments1.xml".to_owned(),
                "xl/drawings/vmlDrawing1.vml".to_owned(),
            ],
            parts: BTreeMap::from([
                (
                    "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>"#.to_vec(),
                ),
                (
                    "xl/comments1.xml".to_owned(),
                    br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Template Author</author></authors><commentList><comment ref="A1" authorId="0"><text><t>template note</t></text></comment></commentList></comments>"#.to_vec(),
                ),
                (
                    "xl/drawings/vmlDrawing1.vml".to_owned(),
                    test_vml(&[("_x0000_s1025", "A1", 0, 0, "0, 15, 0, 2, 2, 15, 3, 2")])
                        .into_bytes(),
                ),
            ]),
        };

        copy_region_comments(
            &mut target,
            &source,
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet1.xml",
            (1, 1, 1, 1),
            0,
            0,
        )
        .expect("create comment metadata parts on the target sheet");

        let comments_part =
            worksheet_relationship_part(&target, "xl/worksheets/sheet1.xml", "/comments")
                .unwrap()
                .expect("new comments relationship");
        let vml_part =
            worksheet_relationship_part(&target, "xl/worksheets/sheet1.xml", "/vmlDrawing")
                .unwrap()
                .expect("new VML relationship");
        assert!(target.parts.contains_key(&comments_part));
        assert!(target.parts.contains_key(&vml_part));
        assert!(
            target
                .text("xl/worksheets/sheet1.xml")
                .unwrap()
                .contains("legacyDrawing")
        );
        assert!(
            target
                .text("xl/worksheets/sheet1.xml")
                .unwrap()
                .contains("xmlns:r=")
        );
        assert!(
            target
                .text("[Content_Types].xml")
                .unwrap()
                .contains("Extension=\"vml\"")
        );
        assert_eq!(
            parse_vml_note_shapes(&target.text(&vml_part).unwrap())
                .unwrap()
                .first()
                .map(|shape| shape.reference.as_str()),
            Some("A1")
        );
    }

    fn test_vml(shapes: &[(&str, &str, u32, u32, &str)]) -> String {
        let mut xml = String::from(
            r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="http://schemas.microsoft.com/office/excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"#,
        );
        for (id, _, row, column, anchor) in shapes {
            xml.push_str(&format!(
                "<v:shape id=\"{id}\" type=\"#_x0000_t202\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\"/><v:path o:connecttype=\"none\"/><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{column}</x:Column></x:ClientData></v:shape>"
            ));
        }
        xml.push_str("</xml>");
        xml
    }

    #[test]
    fn new_sheet_uses_relationship_and_part_ids_not_already_in_use() {
        let relationships = concat!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            r#"<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>"#,
            r#"<Relationship Id="rId2" Type="styles" Target="styles.xml"/>"#,
            r#"<Relationship Id="rId6" Type="theme" Target="theme/theme1.xml"/>"#,
            r#"<Relationship Id="rIdComments" Type="comments" Target="comments1.xml"/>"#,
            "</Relationships>"
        );
        assert_eq!(next_relationship_id(relationships).unwrap(), "rId3");

        let mut package = Package {
            order: vec![
                "xl/worksheets/sheet1.xml".to_owned(),
                "xl/worksheets/sheet3.xml".to_owned(),
            ],
            parts: BTreeMap::from([
                ("xl/worksheets/sheet1.xml".to_owned(), Vec::new()),
                ("xl/worksheets/sheet3.xml".to_owned(), Vec::new()),
            ]),
        };
        assert_eq!(next_worksheet_part_number(&package).unwrap(), 2);
        package
            .parts
            .insert("xl/worksheets/sheet2.xml".to_owned(), Vec::new());
        assert_eq!(next_worksheet_part_number(&package).unwrap(), 4);
    }

    #[test]
    fn formula_shifting_moves_relative_refs_only() {
        // 相对引用平移；绝对引用不动。
        assert_eq!(
            shift_formula_references("=A1+$B$2", 1, 1).unwrap(),
            "=B2+$B$2"
        );
        // 字符串字面量内不平移。
        assert_eq!(
            shift_formula_references("=IF(A1>0,\"A1 是正数\",\"负\")", 1, 0).unwrap(),
            "=IF(A2>0,\"A1 是正数\",\"负\")"
        );
        // 跨表引用（`!` 前）不平移。
        assert_eq!(
            shift_formula_references("=数据!A1+A1", 1, 0).unwrap(),
            "=数据!A1+A2"
        );
        // 越界产生 #REF!。
        assert_eq!(shift_formula_references("=A1", -5, 0).unwrap(), "=#REF!");
    }

    #[test]
    fn shared_formula_dependency_expands_from_master() {
        let sheet_xml = concat!(
            "<sheetData>",
            "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"0\" ref=\"A1:A2\">B1*2</f><v>2</v></c></row>",
            "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>",
            "</sheetData>"
        );
        let masters = shared_formula_masters(sheet_xml).unwrap();
        assert_eq!(
            expand_shared_formula(&masters, " t=\"shared\" si=\"0\"", "", "<c r=\"A2\">").unwrap(),
            "B2*2",
            "dependency cell formula is shifted from the master"
        );
    }

    #[test]
    fn shared_formula_type_is_not_mistaken_for_cell_type() {
        let cell = build_cell_element(r#"<c r="A2"><f t="shared" si="0"/><v>4</v></c>"#);
        assert_eq!(cell.type_attr, None);
        assert_eq!(
            cell.formula_body,
            Some((" t=\"shared\" si=\"0\"".to_owned(), String::new()))
        );
    }

    #[test]
    fn self_closing_blank_cell_does_not_consume_the_next_cell() {
        let cells = cell_elements(
            r#"<row r="211"><c r="A211" s="2"/><c r="B211" t="s"><v>604</v></c></row>"#,
        );
        assert_eq!(cells.len(), 2);
        assert_eq!(cell_reference(&cells[0].xml), Some((211, 1)));
        assert!(!cells[0].xml.contains("<v>"));
        assert!(!cells[0].xml.contains("<is>"));
        assert_eq!(cell_reference(&cells[1].xml), Some((211, 2)));
        assert_eq!(cells[1].type_attr.as_deref(), Some("s"));
        assert_eq!(cells[1].formula_body, None);
        assert_eq!(cells[1].xml, r#"<c r="B211" t="s"><v>604</v></c>"#);
    }

    #[test]
    fn calculation_properties_patch_inserts_and_updates() {
        let with_calc = r#"<workbook><sheets/><calcPr calcId="123"/></workbook>"#;
        let patched = patch_calculation_properties(with_calc, Some(true), Some(true)).unwrap();
        assert!(patched.contains("fullCalcOnLoad=\"1\""));
        assert!(patched.contains("forceFullCalc=\"1\""));

        let without_calc = "<workbook><sheets/></workbook>";
        let patched = patch_calculation_properties(without_calc, Some(true), None).unwrap();
        assert!(patched.contains("<calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/>"));

        // 已存在的属性被替换而不是重复追加。
        let existing =
            r#"<workbook><sheets/><calcPr fullCalcOnLoad="0" forceFullCalc="0"/></workbook>"#;
        let patched = patch_calculation_properties(existing, Some(true), Some(true)).unwrap();
        assert_eq!(patched.matches("fullCalcOnLoad").count(), 1);
        assert!(patched.contains("fullCalcOnLoad=\"1\""));
    }

    #[test]
    fn unsupported_operation_is_rejected_before_write() {
        let capability = BackendCapability {
            engine: crate::EngineId::OoxmlPatch,
            status: crate::CapabilityStatus::Passed,
            format: WorkbookFormat::Xlsx,
            operations: SUPPORTED_OPERATIONS.iter().copied().collect(),
            supported_features: Default::default(),
            preserved: Default::default(),
            platform: "test".to_owned(),
            evidence: None,
        };
        let backend = OoxmlImportBackend::new(capability).unwrap();
        let plan = ExcelEditPlan {
            request_id: "test".to_owned(),
            input: PathBuf::from("missing.xlsx"),
            output: PathBuf::from("out.xlsx"),
            format: WorkbookFormat::Xlsx,
            workbook: crate::WorkbookProfile {
                format: WorkbookFormat::Xlsx,
                inventory_complete: true,
                features_present: Default::default(),
            },
            operations: vec![EditOperation::DeleteColumn {
                sheet: "S".to_owned(),
                column_index: 1,
            }],
            preserve: Default::default(),
            drop_allowed: Default::default(),
        };
        let failure = backend.apply(&plan).unwrap_err();
        assert_eq!(
            failure.kind,
            crate::BackendErrorKind::UnsupportedBeforeWrite
        );
        assert!(
            !plan.output.exists(),
            "rejected plan must not create an output"
        );
    }
}
