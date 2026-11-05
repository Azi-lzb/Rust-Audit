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
use anyhow::{bail, ensure, Context, Result};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

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
        Ok(String::from_utf8(bytes.clone())
            .with_context(|| format!("部件 {part} 不是 UTF-8"))?)
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
            EditOperation::ImportRange { copy_formulas: true, .. }
                | EditOperation::ImportRangeWithFormulaTranslation { copy_formulas: true, .. }
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
            insert_at = workbook_xml.find("<sheets/>").map(|position| position + "<sheets/>".len());
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
    let element_end = rest.find('>').context("workbook.xml 的 calcPr 元素未闭合")?;
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
        let next_id = self.by_name.len() + 1;
        let part = format!("xl/worksheets/sheet{next_id}.xml");
        package.set_text(&part, EMPTY_WORKSHEET_XML.to_owned());
        let rel_id = format!("rId{next_id}");
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
        let part = self
            .part_for(name)?
            .to_owned();
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

fn worksheet_entries(workbook_xml: &str) -> Result<Vec<(String, String)>> {
    let document = roxmltree::Document::parse(workbook_xml).context("解析 workbook.xml")?;
    let mut result = Vec::new();
    for sheet in document.descendants().filter(|node| node.is_element()) {
        if sheet.tag_name().name() != "sheet" {
            continue;
        }
        let (Some(name), Some(rel_id)) = (sheet.attribute("name"), relationship_id(sheet))
        else {
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
    let element = format!("<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{rel_id}\"/>", escape_xml(name), next_sheet_id(workbook_xml)?);
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
        result.replace_range(position..position + "<sheets/>".len(), &format!("<sheets>{element}</sheets>"));
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
    let (source_lo_col, source_lo_row, source_hi_col, source_hi_row) =
        parse_range(source_range)?;
    let (target_lo_col, target_lo_row) = parse_cell(target_start)?;
    let row_offset = target_lo_row as i64 - source_lo_row as i64;
    let column_offset = target_lo_col as i64 - source_lo_col as i64;

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
    is_blank: bool,
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
                    Some(format!("<is><t xml:space=\"preserve\">{}</t></is>", escape_xml(&text))),
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
            is_blank: cell.is_blank,
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
    is_blank: bool,
    has_inline_string: bool,
}

fn cell_elements(sheet_xml: &str) -> Vec<CellElement> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = sheet_xml[cursor..].find("<c ") {
        let start = cursor + relative;
        let Some(end_relative) = sheet_xml[start..].find("</c>") else {
            // 自闭合单元格
            if let Some(self_close) = sheet_xml[start..].find("/>") {
                let end = start + self_close + 2;
                result.push(build_cell_element(&sheet_xml[start..end]));
                cursor = end;
                continue;
            }
            break;
        };
        let end = start + end_relative + "</c>".len();
        result.push(build_cell_element(&sheet_xml[start..end]));
        cursor = end;
    }
    result
}

fn build_cell_element(xml: &str) -> CellElement {
    let formula_body = xml.find("<f").and_then(|start| {
        let rest = &xml[start..];
        let open_end = rest.find('>')?;
        let attrs = rest[..open_end].trim_start_matches("<f").to_owned();
        if rest[open_end - 1..open_end] == *"/" {
            return Some((attrs, String::new()));
        }
        let body_start = open_end + 1;
        let close = rest[body_start..].find("</f>")?;
        Some((
            attrs,
            rest[body_start..body_start + close].to_owned(),
        ))
    });
    // ` t="s"`：跳过 ` t="` 四个字节，取到闭合引号。
    let type_attr = xml.find(" t=\"").and_then(|start| {
        let rest = &xml[start + 4..];
        let end = rest.find('"')?;
        let value = rest[..end].to_owned();
        (!value.trim().is_empty()).then_some(value)
    });
    CellElement {
        is_blank: formula_body.is_none()
            && !xml.contains("<v>")
            && !xml.contains("<is>"),
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
    let (dest_row, dest_column) = cell_reference(cell_xml)
        .context("共享公式依赖格缺少坐标")?;
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
            if let Some((consumed, replaced)) = try_shift_reference(
                &rest,
                row_offset,
                column_offset,
                &result,
            )? {
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
    _source_bounds: (u32, u32, u32, u32),
    row_offset: i64,
    column_offset: i64,
    flags: &ImportFlags,
) -> Result<String> {
    let mut by_position: BTreeMap<(u32, u32), &RegionCell> = BTreeMap::new();
    for cell in cells {
        by_position.insert((cell.row, cell.column), cell);
    }
    let mut rows: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
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

fn write_rows(
    target_xml: &str,
    rows: BTreeMap<u32, BTreeMap<u32, String>>,
) -> Result<String> {
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
        let Some(row_index) = attribute_value(&row.attrs, "r")
            .and_then(|value| value.parse::<u32>().ok())
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
    let open_tag = sheet_xml.find("<sheetData").context("工作表缺少 sheetData")?;
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

/// 复制区域内批注：把源表 `comments{N}.xml` 里落在区域内的批注搬进目标表。
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
    let source_rels = worksheet_rels_part(source_part);
    let Some(rels_xml) = source_package.parts.get(&source_rels) else {
        return Ok(());
    };
    let rels_text = String::from_utf8(rels_xml.clone()).context("工作表 rels 不是 UTF-8")?;
    let Some(comments_target) = relationship_target_by_type(&rels_text, "/comments")? else {
        return Ok(());
    };
    let comments_part = resolve_worksheet_target(source_part, &comments_target);
    let Some(bytes) = source_package.parts.get(&comments_part) else {
        return Ok(());
    };
    let comments_text = String::from_utf8(bytes.clone()).context("comments.xml 不是 UTF-8")?;
    let mut selected = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = comments_text[cursor..].find("<comment ") {
        let start = cursor + relative;
        let Some(close) = comments_text[start..].find("</comment>") else {
            break;
        };
        let end = start + close + "</comment>".len();
        let element = &comments_text[start..end];
        if let Some(reference) = attribute_value(element, "ref") {
            if let Ok((row, column)) = parse_cell(&reference) {
                if row >= lo_row && row <= hi_row && column >= lo_col && column <= hi_col {
                    let shifted_row = (row as i64 + row_offset) as u32;
                    let shifted_column = (column as i64 + column_offset) as u32;
                    let new_reference = format!("{}{shifted_row}", column_letters(shifted_column));
                    selected.push((new_reference, element.to_owned()));
                }
            }
        }
        cursor = end;
    }
    if selected.is_empty() {
        return Ok(());
    }
    // 目标表：追加到既有 comments 部件，或新建一个并登记关系与内容类型。
    let target_rels = worksheet_rels_part(target_part);
    let target_rels_text = package
        .parts
        .get(&target_rels)
        .map(|bytes| String::from_utf8(bytes.clone()))
        .transpose()
        .context("目标工作表 rels 不是 UTF-8")?;
    let existing = target_rels_text
        .as_deref()
        .map(|text| relationship_target_by_type(text, "/comments"))
        .transpose()?
        .flatten();
    let (target_comments_part, target_xml) = match existing {
        Some(target) => {
            let part = resolve_worksheet_target(target_part, &target);
            let text = package.text(&part)?;
            (part, text)
        }
        None => {
            let sheet_file = target_part.rsplit('/').next().unwrap_or("sheet.xml");
            let stem = sheet_file.trim_end_matches(".xml");
            let part = format!("xl/comments{}.xml", comment_part_suffix(stem));
            let skeleton = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
<authors><author>Audit</author></authors><commentList></commentList></comments>"
                .to_owned();
            register_comment_relationship(package, target_part, &part)?;
            append_content_type_override(
                package,
                &format!("/{part}"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            )?;
            (part, skeleton)
        }
    };
    let mut updated = target_xml;
    let insertion = updated
        .find("</commentList>")
        .context("comments.xml 缺少 commentList")?;
    let mut block = String::new();
    for (reference, element) in selected {
        let mut element = element;
        if let Some(start) = element.find("ref=\"") {
            let value_start = start + 5;
            let value_end = element[value_start..]
                .find('"')
                .map(|offset| value_start + offset)
                .unwrap_or(element.len());
            element.replace_range(value_start..value_end, &reference);
        }
        block.push_str(&element);
    }
    updated.insert_str(insertion, &block);
    package.set_text(&target_comments_part, updated);
    Ok(())
}

fn comment_part_suffix(sheet_stem: &str) -> String {
    let digits: String = sheet_stem
        .chars()
        .filter(|character| character.is_ascii_digit())
        .collect();
    if digits.is_empty() {
        format!("-{sheet_stem}")
    } else {
        digits
    }
}

fn worksheet_rels_part(sheet_part: &str) -> String {
    let (directory, file) = sheet_part
        .rsplit_once('/')
        .unwrap_or(("xl/worksheets", sheet_part));
    format!("{directory}/_rels/{file}.rels")
}

fn relationship_target_by_type(rels_xml: &str, type_suffix: &str) -> Result<Option<String>> {
    let document = roxmltree::Document::parse(rels_xml).context("解析工作表 rels")?;
    for relation in document.descendants().filter(|node| node.is_element()) {
        if relation.tag_name().name() != "Relationship" {
            continue;
        }
        let matches_type = relation
            .attribute("Type")
            .is_some_and(|value| value.ends_with(type_suffix));
        if matches_type {
            return Ok(relation.attribute("Target").map(str::to_owned));
        }
    }
    Ok(None)
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

fn register_comment_relationship(
    package: &mut Package,
    sheet_part: &str,
    comments_part: &str,
) -> Result<()> {
    let rels_part = worksheet_rels_part(sheet_part);
    let directory = rels_part
        .rsplit_once('/')
        .map(|(directory, _)| directory)
        .unwrap_or("xl/worksheets/_rels");
    let target = comments_part
        .rsplit_once('/')
        .map(|(_, file)| format!("../comments/{file}"))
        .unwrap_or_else(|| comments_part.to_owned());
    let _ = directory;
    // 目标路径按 OOXML 约定放在 worksheets 同级，用相对路径表达。
    let relative = relative_from_worksheets(comments_part);
    let element = format!(
        "<Relationship Id=\"rIdComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"{relative}\"/>"
    );
    let _ = target;
    match package.parts.get(&rels_part) {
        Some(bytes) => {
            let text = String::from_utf8(bytes.clone()).context("工作表 rels 不是 UTF-8")?;
            if text.contains("relationships/comments") {
                return Ok(());
            }
            let insertion = text
                .rfind("</Relationships>")
                .context("工作表 rels 缺少 Relationships 结束标签")?;
            let mut updated = text;
            updated.insert_str(insertion, &element);
            package.set_text(&rels_part, updated);
        }
        None => {
            let skeleton = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{element}</Relationships>"
            );
            package.set_text(&rels_part, skeleton);
        }
    }
    Ok(())
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
        assert_eq!(failure.kind, crate::BackendErrorKind::UnsupportedBeforeWrite);
        assert!(
            !plan.output.exists(),
            "rejected plan must not create an output"
        );
    }
}
