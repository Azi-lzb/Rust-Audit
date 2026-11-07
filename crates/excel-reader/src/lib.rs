use anyhow::{bail, ensure, Context, Result};
use calamine::{open_workbook_auto, Data, Reader, Sheets};
use roxmltree::Document;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    io::{Read, Write},
    ops::Range,
    path::{Path, PathBuf},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const SPREADSHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Engine-independent representation of a cell value returned by the reader.
/// The tagged JSON form is part of the existing WorkbookDump probe contract.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value")]
pub enum CellValue {
    #[serde(rename = "empty")]
    Empty,
    #[serde(rename = "string")]
    Text(String),
    #[serde(rename = "number")]
    Number(f64),
    #[serde(rename = "integer")]
    Integer(i64),
    #[serde(rename = "boolean")]
    Boolean(bool),
    #[serde(rename = "datetime_serial")]
    DateTimeSerial(f64),
    #[serde(rename = "datetime_iso")]
    DateTimeIso(String),
    #[serde(rename = "duration_iso")]
    DurationIso(String),
    #[serde(rename = "error")]
    Error(String),
}

impl From<&Data> for CellValue {
    fn from(data: &Data) -> Self {
        match data {
            Data::Empty => Self::Empty,
            Data::String(value) => Self::Text(value.clone()),
            Data::Float(value) => Self::Number(*value),
            Data::Int(value) => Self::Integer(*value),
            Data::Bool(value) => Self::Boolean(*value),
            Data::DateTime(value) => Self::DateTimeSerial(value.as_f64()),
            Data::DateTimeIso(value) => Self::DateTimeIso(value.clone()),
            Data::DurationIso(value) => Self::DurationIso(value.clone()),
            Data::Error(value) => Self::Error(format!("{value:?}")),
        }
    }
}

impl CellValue {
    /// Decode the tagged JSON representation emitted by `WorkbookDump`.
    /// Non-finite numbers are emitted by `json!` with a null value; decode those
    /// as NaN so callers can consistently classify them as non-finite.
    pub fn from_probe_json(value: &Value) -> Result<Self> {
        if value.get("type").and_then(Value::as_str) == Some("number")
            && value.get("value").is_some_and(Value::is_null)
        {
            return Ok(Self::Number(f64::NAN));
        }
        serde_json::from_value(value.clone()).context("decode workbook cell JSON")
    }

    fn to_probe_json(self) -> Option<Value> {
        Some(match self {
            Self::Empty => return None,
            Self::Text(value) => json!({"type":"string", "value":value}),
            Self::Number(value) => json!({"type":"number", "value":value}),
            Self::Integer(value) => json!({"type":"integer", "value":value}),
            Self::Boolean(value) => json!({"type":"boolean", "value":value}),
            Self::DateTimeSerial(value) => json!({"type":"datetime_serial", "value":value}),
            Self::DateTimeIso(value) => json!({"type":"datetime_iso", "value":value}),
            Self::DurationIso(value) => json!({"type":"duration_iso", "value":value}),
            Self::Error(value) => json!({"type":"error", "value":value}),
        })
    }
}

#[derive(Debug, Serialize)]
pub struct SheetDump {
    pub name: String,
    pub values: BTreeMap<String, Value>,
    pub formulas: BTreeMap<String, String>,
    pub formula_status: String,
    pub merges: Vec<String>,
    pub merge_status: String,
    pub comments: BTreeMap<String, String>,
    pub comments_status: String,
    pub conditional_formatting: Vec<Value>,
    pub conditional_formatting_status: String,
    pub data_validations: Vec<Value>,
    pub data_validations_status: String,
}

/// A traditional OOXML cell comment with its author resolved from the
/// worksheet's comments part. Calamine does not expose this metadata.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct WorkbookComment {
    pub text: String,
    pub author: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct WorkbookDump {
    pub schema_version: u32,
    pub reader: &'static str,
    pub elapsed_ms: u128,
    pub sheets: Vec<SheetDump>,
}

/// A one-based worksheet area needed by a structure-only read.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CellProjectionRegion {
    pub sheet_name: String,
    pub start_row: usize,
    pub start_column: usize,
    pub end_row: usize,
    pub end_column: usize,
}

/// Cell values, formula text, and merged ranges selected for one worksheet.
#[derive(Clone, Debug, PartialEq)]
pub struct ProjectedSheetCells {
    pub name: String,
    pub values: BTreeMap<String, CellValue>,
    pub formulas: BTreeMap<String, String>,
    pub merges: Vec<String>,
}

/// Minimal workbook data needed to compare fixed S1 structure labels.
#[derive(Clone, Debug, PartialEq)]
pub struct WorkbookCellProjection {
    pub sheet_names: Vec<String>,
    pub sheets: Vec<ProjectedSheetCells>,
    pub elapsed_ms: u128,
    /// True when the existing OOXML error-token compatibility path was used.
    pub fallback_used: bool,
}

/// A local unsupported error token identified in OOXML before the normalized
/// temporary package was retried through Calamine.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct UnknownErrorTokenCell {
    pub sheet_name: String,
    pub coordinate: String,
    pub token: String,
    pub formula_cell: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkbookReadFailureClass {
    FileMissing,
    PermissionDenied,
    UnsupportedFormat,
    UnsupportedCellErrorToken,
    WorksheetRead,
    CorruptWorkbook,
    Unknown,
}

impl WorkbookReadFailureClass {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::FileMissing => "file_missing",
            Self::PermissionDenied => "permission_denied",
            Self::UnsupportedFormat => "unsupported_format",
            Self::UnsupportedCellErrorToken => "unsupported_cell_error_token",
            Self::WorksheetRead => "worksheet_read_error",
            Self::CorruptWorkbook => "corrupt_workbook",
            Self::Unknown => "unknown_read_error",
        }
    }
}

/// Diagnostics for Calamine-first reading and the narrowly supported OOXML
/// error-token compatibility path.
#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
pub struct WorkbookReadDiagnostics {
    pub primary_reader: String,
    pub primary_error: Option<String>,
    pub failure_class: Option<WorkbookReadFailureClass>,
    pub fallback_reader: Option<String>,
    pub fallback_status: String,
    pub normalized_error_cells: Vec<UnknownErrorTokenCell>,
    pub elapsed_ms: u128,
}

#[derive(Debug)]
pub struct WorkbookReadReport {
    pub workbook: WorkbookDump,
    pub diagnostics: WorkbookReadDiagnostics,
}

#[derive(Debug)]
pub struct ValueColumnsReadReport {
    pub sheets: Vec<SheetValues>,
    pub diagnostics: WorkbookReadDiagnostics,
}

/// Diagnostics returned by the explicit unsupported error-token compatibility API.
#[derive(Debug)]
pub struct UnknownErrorTokenRead {
    pub workbook: WorkbookDump,
    pub compatibility_normalization_used: bool,
    pub normalized_cell_count: usize,
    pub normalized_cells: Vec<UnknownErrorTokenCell>,
}

/// A lightweight projection of stored values from a sheet.
///
/// Values use absolute A1 coordinates and preserve the same typed `CellValue`
/// interpretation as `WorkbookDump`. The calamine `Range` is still materialized
/// for the used sheet, but callers can avoid serializing unrelated columns and
/// reading OOXML metadata they do not need.
#[derive(Clone, Debug, PartialEq)]
pub struct SheetValues {
    pub name: String,
    pub values: BTreeMap<String, CellValue>,
    /// Absolute worksheet extent reported by Calamine, in one-based rows and columns.
    /// Empty worksheets report `(0, 0)`; callers that need application-style defaults
    /// should apply those at their own contract boundary.
    pub max_row: usize,
    pub max_column: usize,
}

/// Cell and merged-range bounds stored in one XLSX worksheet part. This is a
/// small OOXML metadata projection used when Calamine's value range does not
/// include styled blank cells or merge-only coordinates. Zero min/max values
/// mean that the worksheet contains no stored cells and no merged ranges.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct WorksheetExtent {
    pub name: String,
    pub min_row: usize,
    pub min_column: usize,
    pub max_row: usize,
    pub max_column: usize,
}

/// One formula cell whose cache node is absent from an OOXML worksheet.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct FormulaCacheLocation {
    pub sheet_name: String,
    pub cell: String,
}

/// Package-level inventory of formula cells and their cached-value nodes.
/// An empty `<v/>` counts as a cache node; this only proves cache presence,
/// not that a value is fresh or mathematically correct.
#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
pub struct FormulaCacheInventory {
    pub formula_cell_count: usize,
    pub cached_formula_cell_count: usize,
    pub missing_cache_cells: Vec<FormulaCacheLocation>,
}

/// One formula whose cached result is absent, empty, or has no recognized
/// value type in the Calamine readback.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct FormulaResultValueLocation {
    pub sheet_name: String,
    pub cell: String,
}

/// Coverage of formula cells with readable, typed cached results.
///
/// A formula cached as an empty string remains typed as `string` and counts as
/// a result. Missing values or `empty` values do not count.
#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
pub struct FormulaResultValueCoverage {
    pub formula_cell_count: usize,
    pub typed_result_cell_count: usize,
    pub missing_result_cells: Vec<FormulaResultValueLocation>,
}

/// Count formulas that Calamine can read back with a typed cached result.
/// Formula text and cache-node presence alone do not prove that a formula has
/// a usable result for Audit extraction.
pub fn formula_result_value_coverage(workbook: &WorkbookDump) -> FormulaResultValueCoverage {
    let mut coverage = FormulaResultValueCoverage::default();
    for sheet in &workbook.sheets {
        for cell in sheet.formulas.keys() {
            coverage.formula_cell_count += 1;
            let has_typed_result = sheet
                .values
                .get(cell)
                .and_then(|value| value.get("type"))
                .and_then(Value::as_str)
                .is_some_and(|value_type| value_type != "empty");
            if has_typed_result {
                coverage.typed_result_cell_count += 1;
            } else {
                coverage
                    .missing_result_cells
                    .push(FormulaResultValueLocation {
                        sheet_name: sheet.name.clone(),
                        cell: cell.clone(),
                    });
            }
        }
    }
    coverage
}

/// Compare formula locations and text while allowing calculation engines to
/// normalize boolean literals and add or remove optional quotes around known
/// internal sheet references. Cell references, ranges, functions and other
/// formula text must remain identical.
pub fn formula_texts_equivalent_after_recalculation(
    source: &WorkbookDump,
    result: &WorkbookDump,
) -> bool {
    if source.sheets.len() != result.sheets.len() {
        return false;
    }
    let sheet_names = source
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    source
        .sheets
        .iter()
        .zip(&result.sheets)
        .all(|(before, after)| {
            before.name == after.name
                && before.formulas.len() == after.formulas.len()
                && before.formulas.iter().all(|(cell, formula)| {
                    after.formulas.get(cell).is_some_and(|result_formula| {
                        formula_text_equivalent_after_recalculation(
                            formula,
                            result_formula,
                            &sheet_names,
                        )
                    })
                })
        })
}

/// Compare a formula while allowing only known calculation-engine rewrites:
/// boolean literals with or without empty parentheses, and optional quoting
/// around references to known internal worksheets. External workbook links,
/// cell references, functions and all other formula text remain significant.
pub fn formula_text_equivalent_after_recalculation(
    source_formula: &str,
    result_formula: &str,
    sheet_names: &[&str],
) -> bool {
    canonicalize_formula_for_recalculation(source_formula, sheet_names)
        == canonicalize_formula_for_recalculation(result_formula, sheet_names)
}

fn canonicalize_formula_for_recalculation(formula: &str, sheet_names: &[&str]) -> String {
    let formula = canonicalize_formula_boolean_literals(formula);
    canonicalize_formula_sheet_reference_quoting(&formula, sheet_names)
}

fn canonicalize_formula_sheet_reference_quoting(formula: &str, sheet_names: &[&str]) -> String {
    let mut candidates = sheet_names
        .iter()
        .copied()
        .map(|name| (name, name.chars().collect::<Vec<_>>()))
        .collect::<Vec<_>>();
    candidates.sort_by_key(|(_, characters)| std::cmp::Reverse(characters.len()));

    let characters = formula.chars().collect::<Vec<_>>();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    let mut in_string = false;
    while index < characters.len() {
        let character = characters[index];
        if character == '"' {
            output.push(character);
            if in_string && characters.get(index + 1) == Some(&'"') {
                output.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if in_string {
            output.push(character);
            index += 1;
            continue;
        }

        if character == '\'' {
            let mut quoted_name = Vec::new();
            let mut cursor = index + 1;
            let mut closing_quote = None;
            while cursor < characters.len() {
                if characters[cursor] == '\'' {
                    if characters.get(cursor + 1) == Some(&'\'') {
                        quoted_name.push('\'');
                        cursor += 2;
                    } else {
                        closing_quote = Some(cursor);
                        break;
                    }
                } else {
                    quoted_name.push(characters[cursor]);
                    cursor += 1;
                }
            }
            if let Some(closing_quote) = closing_quote {
                if characters.get(closing_quote + 1) == Some(&'!') {
                    let name = quoted_name.iter().collect::<String>();
                    if let Some((canonical, _)) = candidates
                        .iter()
                        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(&name))
                    {
                        output.push('\'');
                        output.push_str(&canonical.replace('\'', "''"));
                        output.push_str("'!");
                        index = closing_quote + 2;
                        continue;
                    }
                }
            }
            output.push(character);
            index += 1;
            continue;
        }

        let external_reference_boundary = index > 0 && characters[index - 1] == ']';
        let identifier_boundary = index == 0
            || !(characters[index - 1].is_alphanumeric()
                || matches!(characters[index - 1], '_' | '.'));
        let matched_name = if !external_reference_boundary && identifier_boundary {
            candidates.iter().find(|(_, name)| {
                index + name.len() < characters.len()
                    && characters[index..index + name.len()]
                        .iter()
                        .zip(name)
                        .all(|(left, right)| left.eq_ignore_ascii_case(right))
                    && characters[index + name.len()] == '!'
            })
        } else {
            None
        };
        if let Some((name, _)) = matched_name {
            output.push('\'');
            output.push_str(&name.replace('\'', "''"));
            output.push_str("'!");
            index += name.chars().count() + 1;
        } else {
            output.push(character);
            index += 1;
        }
    }
    output
}

/// Canonicalize only the boolean literal spelling known to differ after
/// LibreOffice saves an XLSX. Quoted formula strings are copied byte-for-byte.
pub fn canonicalize_formula_boolean_literals(formula: &str) -> String {
    let chars = formula.chars().collect::<Vec<_>>();
    let mut normalized = String::with_capacity(formula.len());
    let mut index = 0usize;
    let mut in_string = false;
    let mut in_sheet_name = false;

    while index < chars.len() {
        let current = chars[index];
        if current == '"' && !in_sheet_name {
            normalized.push(current);
            if in_string && chars.get(index + 1) == Some(&'"') {
                normalized.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }

        if current == '\'' && !in_string {
            normalized.push(current);
            if in_sheet_name && chars.get(index + 1) == Some(&'\'') {
                normalized.push('\'');
                index += 2;
                continue;
            }
            in_sheet_name = !in_sheet_name;
            index += 1;
            continue;
        }

        if !in_string && !in_sheet_name {
            let remaining = &chars[index..];
            let token_len = if remaining.len() >= 4
                && remaining[..4]
                    .iter()
                    .copied()
                    .collect::<String>()
                    .eq_ignore_ascii_case("TRUE")
            {
                4
            } else if remaining.len() >= 5
                && remaining[..5]
                    .iter()
                    .copied()
                    .collect::<String>()
                    .eq_ignore_ascii_case("FALSE")
            {
                5
            } else {
                0
            };
            let previous_is_identifier = index > 0
                && (chars[index - 1].is_alphanumeric() || matches!(chars[index - 1], '_' | '.'));
            let next_is_identifier = chars.get(index + token_len).is_some_and(|character| {
                character.is_alphanumeric() || matches!(character, '_' | '.')
            });
            let followed_by_empty_call = token_len > 0
                && chars.get(index + token_len) == Some(&'(')
                && chars.get(index + token_len + 1) == Some(&')');
            if token_len > 0
                && !previous_is_identifier
                && !next_is_identifier
                && followed_by_empty_call
            {
                normalized.extend(chars[index..index + token_len].iter());
                index += token_len + 2;
                continue;
            }
        }

        normalized.push(current);
        index += 1;
    }
    normalized
}

#[derive(Default)]
struct OoxmlSheetMetadata {
    comments: BTreeMap<String, String>,
    comment_records: BTreeMap<String, WorkbookComment>,
    conditional_formatting: Vec<Value>,
    data_validations: Vec<Value>,
}

/// The three SpreadsheetML worksheet header variants.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd)]
pub enum HeaderVariant {
    Odd,
    Even,
    First,
}

/// Plain text from one worksheet header variant, including Excel control codes.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderText {
    pub sheet_name: String,
    pub variant: HeaderVariant,
    pub text: String,
}

/// Header text inventory for one worksheet, in workbook sheet order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderInventory {
    pub sheet_name: String,
    pub headers: Vec<WorksheetHeaderText>,
}

/// Scope of a workbook defined name. `local_sheet_id` follows the zero-based
/// index in the workbook's `<sheets>` collection, including non-worksheet tabs.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum DefinedNameScope {
    Workbook,
    Worksheet {
        local_sheet_id: usize,
        sheet_name: Option<String>,
        sheet_kind: Option<String>,
    },
    UnresolvedWorksheet {
        local_sheet_id: usize,
    },
    InvalidWorksheetScope {
        raw_local_sheet_id: String,
    },
}

/// A resolved A1 area in a defined-name formula.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct DefinedNameArea {
    pub sheet_name: String,
    pub sheet_kind: String,
    pub reference: String,
}

/// How much of the defined-name formula this reader understands.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DefinedNameParseStatus {
    Areas,
    NonAreaExpression,
    UnsupportedAreaExpression,
    InvalidAreaExpression,
}

/// Namespace-aware inventory of workbook defined names. `raw_reference` is
/// preserved exactly as stored (apart from XML entity decoding); `areas` is
/// populated only when every union member is a supported A1 area reference.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct DefinedNameRecord {
    pub name: String,
    pub scope: DefinedNameScope,
    pub raw_reference: String,
    pub parse_status: DefinedNameParseStatus,
    pub areas: Vec<DefinedNameArea>,
}

/// One `sqref` area covered by an OOXML conditional-formatting block.
/// Multiple rule blocks may cover the same area; the reader returns each
/// sheet/address pair once so callers can inspect only those target ranges.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct ConditionalFormattingArea {
    pub sheet_name: String,
    pub address: String,
}

#[derive(Clone, Debug)]
struct WorkbookSheetIdentity {
    name: String,
    kind: String,
}

/// Read workbook and worksheet-scoped defined names from XLSX/XLSM OOXML.
/// This is a metadata adapter and does not read or infer cell values. Supported
/// area expressions are returned with sheet identity; formulas, constants,
/// external references and unsupported reference forms remain available in
/// `raw_reference` with an explicit parse status.
pub fn read_defined_names(path: impl AsRef<Path>) -> Result<Vec<DefinedNameRecord>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "defined-name reader only supports .xlsx and .xlsm"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open workbook package")?)
        .context("open workbook ZIP package")?;
    let (workbook_part, sheets) = workbook_sheet_identities(&mut zip)?;
    let workbook_xml = zip_read(&mut zip, &workbook_part).context("read workbook part")?;
    let document = Document::parse(&workbook_xml).context("parse workbook part")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "workbook part has an unexpected root element"
    );

    let Some(container) = root.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "definedNames"
    }) else {
        return Ok(Vec::new());
    };
    let mut records = Vec::new();
    for node in container.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "definedName"
    }) {
        let name = node.attribute("name").context("definedName has no name")?;
        let raw_reference = node.text().unwrap_or_default().to_owned();
        let scope = match node.attribute("localSheetId") {
            None => DefinedNameScope::Workbook,
            Some(raw) => match raw.parse::<usize>() {
                Ok(local_sheet_id) => match sheets.get(local_sheet_id) {
                    Some(sheet) => DefinedNameScope::Worksheet {
                        local_sheet_id,
                        sheet_name: Some(sheet.name.clone()),
                        sheet_kind: Some(sheet.kind.clone()),
                    },
                    None => DefinedNameScope::UnresolvedWorksheet { local_sheet_id },
                },
                Err(_) => DefinedNameScope::InvalidWorksheetScope {
                    raw_local_sheet_id: raw.to_owned(),
                },
            },
        };
        let (parse_status, areas) = parse_defined_name_areas(&raw_reference, &sheets);
        records.push(DefinedNameRecord {
            name: name.to_owned(),
            scope,
            raw_reference,
            parse_status,
            areas,
        });
    }
    Ok(records)
}

/// Read the exact OOXML conditional-formatting target ranges without loading
/// cell values. This is a targeted fallback for templates that store rules on
/// worksheets but omit a `条件格式区域` defined name.
pub fn read_conditional_formatting_areas(
    path: impl AsRef<Path>,
) -> Result<Vec<ConditionalFormattingArea>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "conditional-formatting area reader only supports .xlsx and .xlsm"
    );
    let metadata = extract_ooxml_metadata(path)
        .with_context(|| format!("read conditional-formatting areas from {}", path.display()))?;
    let mut seen = BTreeSet::new();
    let mut areas = Vec::new();
    for (sheet_name, sheet_metadata) in metadata {
        for node in sheet_metadata.conditional_formatting {
            if node.get("name").and_then(Value::as_str) != Some("conditionalFormatting") {
                continue;
            }
            let Some(sqref) = node
                .get("attributes")
                .and_then(|attributes| attributes.get("sqref"))
                .and_then(Value::as_str)
            else {
                continue;
            };
            for address in sqref.split_whitespace().filter(|address| !address.is_empty()) {
                let key = (sheet_name.clone(), address.to_ascii_uppercase());
                if seen.insert(key) {
                    areas.push(ConditionalFormattingArea {
                        sheet_name: sheet_name.clone(),
                        address: address.to_owned(),
                    });
                }
            }
        }
    }
    Ok(areas)
}

fn workbook_sheet_identities(
    zip: &mut ZipArchive<fs::File>,
) -> Result<(String, Vec<WorkbookSheetIdentity>)> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            (workbook && !external)
                .then(|| relationship.attribute("Target"))
                .flatten()
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");
    let workbook_xml = zip_read(zip, &workbook_part)?;
    let workbook_doc = Document::parse(&workbook_xml)?;
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let mut rels = BTreeMap::new();
    for relationship in rels_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if let (Some(id), Some(kind), Some(target)) = (
            relationship.attribute("Id"),
            relationship.attribute("Type"),
            relationship.attribute("Target"),
        ) {
            let external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            rels.insert(
                id.to_owned(),
                (kind.to_owned(), target.to_owned(), external),
            );
        }
    }
    let mut sheets = Vec::new();
    for sheet in workbook_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let rel_id = sheet
            .attribute((OFFICE_REL_NS, "id"))
            .context("workbook sheet has no relationship id")?;
        let (kind, target, external) = rels
            .get(rel_id)
            .with_context(|| format!("unresolved relationship for sheet {name:?}"))?;
        ensure!(
            !*external,
            "workbook sheet {name:?} has an external relationship"
        );
        let part = norm_part(workbook_parent, target);
        ensure!(
            !part.is_empty(),
            "workbook sheet {name:?} has an empty part path"
        );
        sheets.push(WorkbookSheetIdentity {
            name: name.to_owned(),
            kind: kind.rsplit('/').next().unwrap_or(kind).to_owned(),
        });
    }
    Ok((workbook_part, sheets))
}

fn parse_defined_name_areas(
    expression: &str,
    sheets: &[WorkbookSheetIdentity],
) -> (DefinedNameParseStatus, Vec<DefinedNameArea>) {
    if expression.contains('(') || expression.contains(')') {
        return (DefinedNameParseStatus::NonAreaExpression, Vec::new());
    }
    let parts = split_formula_union(expression);
    if parts.is_empty() {
        return (DefinedNameParseStatus::NonAreaExpression, Vec::new());
    }
    let mut areas = Vec::new();
    for part in parts {
        let Some((sheet_token, reference)) = part.rsplit_once('!') else {
            let status = if is_a1_area(part) {
                DefinedNameParseStatus::UnsupportedAreaExpression
            } else {
                DefinedNameParseStatus::NonAreaExpression
            };
            return (status, Vec::new());
        };
        let Some(sheet_name) = decode_sheet_token(sheet_token) else {
            return (
                DefinedNameParseStatus::UnsupportedAreaExpression,
                Vec::new(),
            );
        };
        if !is_a1_area(reference) {
            return (
                DefinedNameParseStatus::UnsupportedAreaExpression,
                Vec::new(),
            );
        }
        let Some(sheet) = sheets.iter().find(|sheet| sheet.name == sheet_name) else {
            return (DefinedNameParseStatus::InvalidAreaExpression, Vec::new());
        };
        areas.push(DefinedNameArea {
            sheet_name: sheet.name.clone(),
            sheet_kind: sheet.kind.clone(),
            reference: reference.to_owned(),
        });
    }
    (DefinedNameParseStatus::Areas, areas)
}

fn split_formula_union(expression: &str) -> Vec<&str> {
    split_formula_union_bytes(expression)
}

fn split_formula_union_bytes(expression: &str) -> Vec<&str> {
    let (mut start, mut quoted) = (0, false);
    let mut parts = Vec::new();
    let chars: Vec<(usize, char)> = expression.char_indices().collect();
    let mut i = 0;
    while i < chars.len() {
        let (byte, ch) = chars[i];
        if ch == '\'' {
            if quoted && chars.get(i + 1).is_some_and(|(_, next)| *next == '\'') {
                i += 2;
                continue;
            }
            quoted = !quoted;
        } else if ch == ',' && !quoted {
            parts.push(expression[start..byte].trim());
            start = byte + 1;
        }
        i += 1;
    }
    if quoted {
        return Vec::new();
    }
    parts.push(expression[start..].trim());
    parts
}

fn decode_sheet_token(token: &str) -> Option<String> {
    let token = token.strip_prefix('=').unwrap_or(token);
    let decoded = if token.starts_with('\'') && token.ends_with('\'') && token.len() >= 2 {
        token[1..token.len() - 1].replace("''", "'")
    } else if token
        .chars()
        .all(|ch| ch.is_alphanumeric() || ch == '_' || ch == '.')
    {
        token.to_owned()
    } else {
        return None;
    };
    // External workbook qualifiers and 3-D references are deliberately unsupported.
    (!decoded.contains('[') && !decoded.contains(']') && !decoded.contains(':')).then_some(decoded)
}

fn is_a1_area(reference: &str) -> bool {
    fn cell(value: &str) -> bool {
        let value = value.strip_prefix('$').unwrap_or(value);
        let letters = value
            .chars()
            .take_while(|ch| ch.is_ascii_alphabetic())
            .count();
        if letters == 0 || letters > 3 {
            return false;
        }
        let column = value[..letters].chars().fold(0_usize, |column, ch| {
            column * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1)
        });
        if column > 16_384 {
            return false;
        }
        let rest = &value[letters..];
        let rest = rest.strip_prefix('$').unwrap_or(rest);
        !rest.is_empty()
            && rest.chars().all(|ch| ch.is_ascii_digit())
            && rest
                .parse::<u32>()
                .is_ok_and(|row| (1..=1_048_576).contains(&row))
    }
    let mut pair = reference.split(':');
    let first = pair.next().unwrap_or_default();
    let second = pair.next();
    pair.next().is_none() && cell(first) && second.map_or(true, cell)
}

/// Read odd/even/first-page header text from an OOXML workbook.
///
/// This is an OOXML metadata read, separate from Calamine cell-value reading.
/// Rich/nested header markup is rejected so callers cannot mistake a partial
/// text projection for the original header contract.
pub fn read_header_inventory(path: impl AsRef<Path>) -> Result<Vec<WorksheetHeaderInventory>> {
    let path = path.as_ref();
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xlam"
            )
        });
    ensure!(
        is_ooxml,
        "header metadata reader only supports XLSX/XLSM/XLAM"
    );

    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut inventory = Vec::with_capacity(worksheets.len());
    for (sheet_name, sheet_part) in worksheets {
        let xml = zip_read(&mut zip, &sheet_part)?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let mut headers = Vec::new();
        if let Some(header_footer) = document.root_element().descendants().find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "headerFooter"
        }) {
            let mut seen = std::collections::BTreeSet::new();
            for header in header_footer.children().filter(|node| node.is_element()) {
                if header.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS) {
                    continue;
                }
                let variant = match header.tag_name().name() {
                    "oddHeader" => HeaderVariant::Odd,
                    "evenHeader" => HeaderVariant::Even,
                    "firstHeader" => HeaderVariant::First,
                    _ => continue,
                };
                ensure!(
                    seen.insert(variant),
                    "worksheet {sheet_name:?} contains duplicate {variant:?} header elements"
                );
                if header.children().any(|child| child.is_element()) {
                    bail!(
                        "worksheet {sheet_name:?} contains unsupported nested markup in {variant:?} header"
                    );
                }
                let text = header
                    .children()
                    .filter(|child| child.is_text())
                    .filter_map(|child| child.text())
                    .collect();
                headers.push(WorksheetHeaderText {
                    sheet_name: sheet_name.clone(),
                    variant,
                    text,
                });
            }
        }
        inventory.push(WorksheetHeaderInventory {
            sheet_name,
            headers,
        });
    }
    Ok(inventory)
}

fn value_json(data: &Data) -> Option<Value> {
    CellValue::from(data).to_probe_json()
}

fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    loop {
        let rem = col % 26;
        out.insert(0, (b'A' + rem as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out
}

fn coord(row: usize, col: usize) -> String {
    format!("{}{}", column_letters(col), row + 1)
}

fn dimensions_to_range(start: (u32, u32), end: (u32, u32)) -> String {
    format!(
        "{}{}:{}{}",
        column_letters(start.1 as usize),
        start.0 + 1,
        column_letters(end.1 as usize),
        end.0 + 1
    )
}

fn merges_for(
    workbook: &mut Sheets<std::io::BufReader<fs::File>>,
    name: &str,
) -> Result<(Vec<String>, &'static str)> {
    match workbook {
        Sheets::Xlsx(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xls(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xlsb(_) => Ok((
            Vec::new(),
            "unsupported_api: Calamine 0.37.0 exposes no xlsb merge reader",
        )),
        Sheets::Ods(_) => Ok((Vec::new(), "out_of_scope_format")),
    }
}

pub fn read_workbook(path: impl AsRef<Path>) -> Result<WorkbookDump> {
    read_workbook_with_diagnostics(path).map(|report| report.workbook)
}

/// Read only the cells covered by the supplied regions for S1 structure
/// matching. Calamine still validates stored values in every worksheet, as
/// the full reader does, but this path avoids building a workbook-wide JSON
/// dump and skips unrelated OOXML metadata such as comments and conditional
/// formatting. `include_merge_anchors` is used for source workbooks so a
/// requested merged cell can be compared through its top-left anchor.
pub fn read_workbook_structure_projection(
    path: impl AsRef<Path>,
    regions: &[CellProjectionRegion],
    include_merge_anchors: bool,
) -> Result<WorkbookCellProjection> {
    let path = path.as_ref();
    let started = std::time::Instant::now();
    match read_workbook_structure_projection_calamine(path, regions, include_merge_anchors) {
        Ok(mut projection) => {
            projection.elapsed_ms = started.elapsed().as_millis();
            Ok(projection)
        }
        Err(primary_error) => {
            let failure_class = classify_read_failure(path, &primary_error);
            if failure_class != WorkbookReadFailureClass::UnsupportedCellErrorToken {
                let primary_context = format!("{primary_error:#}");
                return Err(primary_error).context(format!(
                    "Calamine structure projection failed ({}) for {}; no safe OOXML value fallback applies: {primary_context}",
                    failure_class.as_str(),
                    path.display()
                ));
            }

            // Keep the existing narrow compatibility behavior for uncommon
            // error tokens: normalize only those OOXML cells on a temporary
            // copy, then project from the successful full-reader result.
            let report = read_workbook_with_diagnostics(path).with_context(|| {
                format!(
                    "Calamine structure projection failed ({}) for {}: {:#}; existing error-token fallback failed",
                    failure_class.as_str(),
                    path.display(),
                    primary_error
                )
            })?;
            let fallback_used = report.diagnostics.fallback_reader.is_some();
            let mut projection =
                project_workbook_dump(&report.workbook, regions, include_merge_anchors)?;
            projection.elapsed_ms = started.elapsed().as_millis();
            projection.fallback_used = fallback_used;
            Ok(projection)
        }
    }
}

fn read_workbook_structure_projection_calamine(
    path: &Path,
    regions: &[CellProjectionRegion],
    include_merge_anchors: bool,
) -> Result<WorkbookCellProjection> {
    validate_legacy_xls_signature(path)?;
    let started = std::time::Instant::now();
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let sheet_names = workbook.sheet_names().to_vec();
    let mut regions_by_sheet = BTreeMap::<String, Vec<&CellProjectionRegion>>::new();
    for region in regions {
        regions_by_sheet
            .entry(region.sheet_name.clone())
            .or_default()
            .push(region);
    }

    let mut sheets = Vec::new();
    for name in &sheet_names {
        let range = workbook
            .worksheet_range(name)
            .with_context(|| format!("read values for {name}"))?;
        let merges = merges_for(&mut workbook, name)
            .with_context(|| format!("read merge regions for {name}"))?
            .0;
        let Some(sheet_regions) = regions_by_sheet.get(name) else {
            continue;
        };

        let mut merge_anchors = BTreeSet::new();
        if include_merge_anchors {
            for merge in &merges {
                let Some((first_row, first_column, last_row, last_column)) =
                    parse_ooxml_range_reference(merge).ok()
                else {
                    continue;
                };
                if sheet_regions.iter().any(|region| {
                    region.start_row == region.end_row
                        && region.start_column == region.end_column
                        && coordinate_within(
                            region.start_row,
                            region.start_column,
                            first_row,
                            first_column,
                            last_row,
                            last_column,
                        )
                }) {
                    merge_anchors.insert((first_row, first_column));
                }
            }
        }

        let (start_row, start_column) = range.start().unwrap_or((0, 0));
        let mut values = BTreeMap::new();
        for (local_row, local_column, value) in range.cells() {
            let row = start_row as usize + local_row + 1;
            let column = start_column as usize + local_column + 1;
            if !projection_coordinate_requested(row, column, sheet_regions)
                && !merge_anchors.contains(&(row, column))
            {
                continue;
            }
            let value = CellValue::from(value);
            if !matches!(value, CellValue::Empty) {
                values.insert(coord(row - 1, column - 1), value);
            }
        }

        let mut formulas = BTreeMap::new();
        if let Ok(formula_range) = workbook.worksheet_formula(name) {
            let (formula_start_row, formula_start_column) = formula_range.start().unwrap_or((0, 0));
            for (local_row, local_column, formula) in formula_range.cells() {
                if formula.trim().is_empty() {
                    continue;
                }
                let row = formula_start_row as usize + local_row + 1;
                let column = formula_start_column as usize + local_column + 1;
                if projection_coordinate_requested(row, column, sheet_regions)
                    || merge_anchors.contains(&(row, column))
                {
                    formulas.insert(coord(row - 1, column - 1), formula.to_owned());
                }
            }
        }

        sheets.push(ProjectedSheetCells {
            name: name.clone(),
            values,
            formulas,
            merges,
        });
    }

    Ok(WorkbookCellProjection {
        sheet_names,
        sheets,
        elapsed_ms: started.elapsed().as_millis(),
        fallback_used: false,
    })
}

fn project_workbook_dump(
    workbook: &WorkbookDump,
    regions: &[CellProjectionRegion],
    include_merge_anchors: bool,
) -> Result<WorkbookCellProjection> {
    let mut regions_by_sheet = BTreeMap::<String, Vec<&CellProjectionRegion>>::new();
    for region in regions {
        regions_by_sheet
            .entry(region.sheet_name.clone())
            .or_default()
            .push(region);
    }
    let mut sheets = Vec::new();
    for sheet in &workbook.sheets {
        let Some(sheet_regions) = regions_by_sheet.get(&sheet.name) else {
            continue;
        };
        let mut merge_anchors = BTreeSet::new();
        if include_merge_anchors {
            for merge in &sheet.merges {
                let Some((first_row, first_column, last_row, last_column)) =
                    parse_ooxml_range_reference(merge).ok()
                else {
                    continue;
                };
                if sheet_regions.iter().any(|region| {
                    region.start_row == region.end_row
                        && region.start_column == region.end_column
                        && coordinate_within(
                            region.start_row,
                            region.start_column,
                            first_row,
                            first_column,
                            last_row,
                            last_column,
                        )
                }) {
                    merge_anchors.insert((first_row, first_column));
                }
            }
        }

        let mut values = BTreeMap::new();
        for (address, value) in &sheet.values {
            let Some((row, column)) = parse_ooxml_cell_reference(address).ok() else {
                continue;
            };
            if projection_coordinate_requested(row, column, sheet_regions)
                || merge_anchors.contains(&(row, column))
            {
                values.insert(address.clone(), CellValue::from_probe_json(value)?);
            }
        }
        let formulas = sheet
            .formulas
            .iter()
            .filter_map(|(address, formula)| {
                let (row, column) = parse_ooxml_cell_reference(address).ok()?;
                (projection_coordinate_requested(row, column, sheet_regions)
                    || merge_anchors.contains(&(row, column)))
                .then(|| (address.clone(), formula.clone()))
            })
            .collect();
        sheets.push(ProjectedSheetCells {
            name: sheet.name.clone(),
            values,
            formulas,
            merges: sheet.merges.clone(),
        });
    }

    Ok(WorkbookCellProjection {
        sheet_names: workbook
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect(),
        sheets,
        elapsed_ms: workbook.elapsed_ms,
        fallback_used: false,
    })
}

fn projection_coordinate_requested(
    row: usize,
    column: usize,
    regions: &[&CellProjectionRegion],
) -> bool {
    regions.iter().any(|region| {
        row >= region.start_row
            && row <= region.end_row
            && column >= region.start_column
            && column <= region.end_column
    })
}

fn coordinate_within(
    row: usize,
    column: usize,
    first_row: usize,
    first_column: usize,
    last_row: usize,
    last_column: usize,
) -> bool {
    row >= first_row && row <= last_row && column >= first_column && column <= last_column
}

/// Read the worksheet order and formula counts without materializing stored
/// cell values or OOXML metadata. This is intended for output checks whose
/// contract only requires confirming that a workbook reopens with its expected
/// sheets and formula inventory.
pub fn read_workbook_sheet_formula_counts(path: impl AsRef<Path>) -> Result<Vec<(String, usize)>> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let mut result = Vec::with_capacity(names.len());
    for name in names {
        let formula_count = workbook
            .worksheet_formula(&name)
            .with_context(|| format!("read formulas for {name}"))?
            .cells()
            .filter(|(_, _, formula)| !formula.trim().is_empty())
            .count();
        result.push((name, formula_count));
    }
    Ok(result)
}

/// Read with Calamine first and use a temporary, cell-scoped OOXML
/// compatibility copy only when Calamine rejects an unsupported error token
/// error token. All other failures retain their original error chain.
pub fn read_workbook_with_diagnostics(path: impl AsRef<Path>) -> Result<WorkbookReadReport> {
    let path = path.as_ref();
    let started = std::time::Instant::now();
    match read_workbook_calamine(path) {
        Ok(workbook) => Ok(WorkbookReadReport {
            workbook,
            diagnostics: WorkbookReadDiagnostics {
                primary_reader: "calamine".to_owned(),
                primary_error: None,
                failure_class: None,
                fallback_reader: None,
                fallback_status: "not_needed".to_owned(),
                normalized_error_cells: Vec::new(),
                elapsed_ms: started.elapsed().as_millis(),
            },
        }),
        Err(primary_error) => {
            let failure_class = classify_read_failure(path, &primary_error);
            if failure_class != WorkbookReadFailureClass::UnsupportedCellErrorToken {
                let primary_context = format!("{primary_error:#}");
                return Err(primary_error).context(format!(
                    "Calamine read failed ({}) for {}; no safe OOXML value fallback applies: {primary_context}",
                    failure_class.as_str(),
                    path.display()
                ));
            }
            let (temporary_copy, normalized_error_cells) =
                create_normalized_error_copy(path).with_context(|| {
                    format!(
                        "Calamine read failed ({}) for {}: {:#}; targeted OOXML fallback preparation failed",
                        failure_class.as_str(),
                        path.display(),
                        primary_error
                    )
                })?;
            let workbook = read_workbook_calamine(temporary_copy.path()).with_context(|| {
                format!(
                    "Calamine retry failed after targeted OOXML error-token normalization for {}: primary error: {:#}",
                    path.display(),
                    primary_error
                )
            })?;
            Ok(WorkbookReadReport {
                workbook,
                diagnostics: WorkbookReadDiagnostics {
                    primary_reader: "calamine".to_owned(),
                    primary_error: Some(format!("{primary_error:#}")),
                    failure_class: Some(failure_class),
                    fallback_reader: Some("ooxml_error_token_normalization+calamine".to_owned()),
                    fallback_status: "succeeded".to_owned(),
                    normalized_error_cells,
                    elapsed_ms: started.elapsed().as_millis(),
                },
            })
        }
    }
}

fn read_workbook_calamine(path: &Path) -> Result<WorkbookDump> {
    validate_legacy_xls_signature(path)?;
    let started = std::time::Instant::now();
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let is_ooxml = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| matches!(e.to_ascii_lowercase().as_str(), "xlsx" | "xlsm" | "xlam"))
        .unwrap_or(false);
    let metadata_by_sheet = if is_ooxml {
        extract_ooxml_metadata(path)?
    } else {
        BTreeMap::new()
    };
    let mut sheets = Vec::new();
    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            if let Some(v) = value_json(value) {
                values.insert(
                    coord(
                        start_row as usize + local_row,
                        start_col as usize + local_col,
                    ),
                    v,
                );
            }
        }
        let (formulas, formula_status) = match workbook.worksheet_formula(&name) {
            Ok(range) => {
                let (fr, fc) = range.start().unwrap_or((0, 0));
                (
                    range
                        .cells()
                        .filter_map(|(r, c, v)| {
                            (!v.trim().is_empty())
                                .then(|| (coord(fr as usize + r, fc as usize + c), v.to_owned()))
                        })
                        .collect(),
                    "read_by_calamine".to_owned(),
                )
            }
            Err(error) => (BTreeMap::new(), format!("read_error:{error}")),
        };
        let (merges, merge_status) = merges_for(&mut workbook, &name)
            .with_context(|| format!("read merge regions for {name}"))?;
        let (
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        ) = if is_ooxml {
            let metadata = metadata_by_sheet
                .get(&name)
                .with_context(|| format!("OOXML metadata missing for sheet {name}"))?;
            (
                metadata.comments.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.conditional_formatting.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.data_validations.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
            )
        } else {
            (
                BTreeMap::new(),
                "unsupported_api: comment adapter only supports OOXML xlsx/xlsm".to_owned(),
                Vec::new(),
                "unsupported_api: conditional-formatting adapter only supports OOXML xlsx/xlsm"
                    .to_owned(),
                Vec::new(),
                "unsupported_api: data-validation adapter only supports OOXML xlsx/xlsm".to_owned(),
            )
        };
        sheets.push(SheetDump {
            name,
            values,
            formulas,
            formula_status,
            merges,
            merge_status: merge_status.to_owned(),
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        });
    }
    Ok(WorkbookDump {
        schema_version: 3,
        reader: "calamine",
        elapsed_ms: started.elapsed().as_millis(),
        sheets,
    })
}

fn classify_read_failure(path: &Path, error: &anyhow::Error) -> WorkbookReadFailureClass {
    for cause in error.chain() {
        if let Some(io_error) = cause.downcast_ref::<std::io::Error>() {
            if io_error.kind() == std::io::ErrorKind::PermissionDenied {
                return WorkbookReadFailureClass::PermissionDenied;
            }
            if io_error.kind() == std::io::ErrorKind::NotFound {
                return WorkbookReadFailureClass::FileMissing;
            }
        }
    }
    if !path.exists() {
        return WorkbookReadFailureClass::FileMissing;
    }
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !matches!(extension.as_str(), "xlsx" | "xlsm" | "xls" | "xlsb" | "ods") {
        return WorkbookReadFailureClass::UnsupportedFormat;
    }
    let message = format!("{error:#}").to_ascii_lowercase();
    if message.contains("unsupported cell error value") {
        return WorkbookReadFailureClass::UnsupportedCellErrorToken;
    }
    if message.contains("read values for ") || message.contains("read merge regions for ") {
        return WorkbookReadFailureClass::WorksheetRead;
    }
    if matches!(extension.as_str(), "xlsx" | "xlsm") {
        return WorkbookReadFailureClass::CorruptWorkbook;
    }
    WorkbookReadFailureClass::Unknown
}

struct TemporaryWorkbookCopy(PathBuf);

impl TemporaryWorkbookCopy {
    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TemporaryWorkbookCopy {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn create_normalized_error_copy(
    path: &Path,
) -> Result<(TemporaryWorkbookCopy, Vec<UnknownErrorTokenCell>)> {
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "targeted OOXML error-token fallback supports XLSX/XLSM only"
    );
    let (replacements, normalized_cells) = prepare_unknown_error_token_replacements(path)?;
    ensure!(
        !replacements.is_empty() && !normalized_cells.is_empty(),
        "Calamine reported an unsupported cell error but no safe literal OOXML error token was found"
    );
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)?
        .as_nanos();
    let temp = std::env::temp_dir().join(format!(
        "excel-reader-error-text-{}-{nonce}.xlsx",
        std::process::id()
    ));
    if let Err(error) = write_package_with_replacements(path, &temp, &replacements) {
        let _ = fs::remove_file(&temp);
        return Err(error)
            .context("write temporary workbook with unsupported error tokens normalized");
    }
    Ok((TemporaryWorkbookCopy(temp), normalized_cells))
}

/// Read traditional cell comments and author names from an OOXML workbook.
///
/// The map is keyed by worksheet name and then absolute A1 coordinate. This is
/// a metadata-only supplement; cell values continue to be read through
/// [`read_workbook`] and Calamine.
pub fn read_ooxml_comments(
    path: impl AsRef<Path>,
) -> Result<BTreeMap<String, BTreeMap<String, WorkbookComment>>> {
    let path = path.as_ref();
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    ensure!(
        matches!(extension.as_str(), "xlsx" | "xlsm" | "xlam"),
        "OOXML comment reader only supports .xlsx, .xlsm, and .xlam files"
    );
    let metadata = extract_ooxml_metadata(path)?;
    Ok(metadata
        .into_iter()
        .map(|(sheet, metadata)| (sheet, metadata.comment_records))
        .collect())
}

/// Read a workbook with the same safe error-token compatibility used by the
/// default reader and expose the normalized cell locations to diagnostics.
pub fn read_workbook_with_unknown_error_tokens_as_text(
    path: impl AsRef<Path>,
) -> Result<UnknownErrorTokenRead> {
    let report = read_workbook_with_diagnostics(path)?;
    let normalized_cell_count = report.diagnostics.normalized_error_cells.len();
    let compatibility_normalization_used = report.diagnostics.fallback_reader.is_some();
    Ok(UnknownErrorTokenRead {
        workbook: report.workbook,
        compatibility_normalization_used,
        normalized_cell_count,
        normalized_cells: report.diagnostics.normalized_error_cells,
    })
}

fn prepare_unknown_error_token_replacements(
    path: &Path,
) -> Result<(BTreeMap<String, Vec<u8>>, Vec<UnknownErrorTokenCell>)> {
    let mut zip = ZipArchive::new(fs::File::open(path).context("open workbook package")?)?;
    let mut replacements = BTreeMap::new();
    let mut normalized_cells = Vec::new();
    for (sheet_name, part) in ooxml_worksheet_parts(&mut zip)? {
        let xml = zip_read(&mut zip, &part)?.into_bytes();
        let (patched, cells) = normalize_unknown_error_cells(&xml, &sheet_name)?;
        if patched != xml {
            replacements.insert(part, patched);
            normalized_cells.extend(cells);
        }
    }
    Ok((replacements, normalized_cells))
}

fn normalize_unknown_error_cells(
    xml: &[u8],
    sheet_name: &str,
) -> Result<(Vec<u8>, Vec<UnknownErrorTokenCell>)> {
    let document = Document::parse(std::str::from_utf8(xml)?)?;
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut normalized_cells = Vec::new();
    for cell in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "c"
            && node.attribute("t") == Some("e")
    }) {
        ensure!(
            cell.attribute("r").is_some(),
            "error cell is missing a coordinate"
        );
        let value_nodes = cell
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "v"
            })
            .collect::<Vec<_>>();
        ensure!(
            value_nodes.len() == 1,
            "error cell {:?} must contain exactly one cached value",
            cell.attribute("r")
        );
        ensure!(
            cell.children().all(|child| !child.is_element()
                || (child.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && matches!(child.tag_name().name(), "v" | "f"))),
            "unsupported child element in literal error cell {:?}",
            cell.attribute("r")
        );
        let value = value_nodes[0].text().unwrap_or_default();
        if is_calamine_error_token(value) {
            continue;
        }
        let formula_cell = cell.children().any(|child| {
            child.is_element()
                && child.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && child.tag_name().name() == "f"
        });
        let type_attribute = cell
            .attribute_node("t")
            .context("error cell is missing type attribute")?;
        if formula_cell {
            // A formula cache may legally be stored as a string result. Keep
            // the formula and value intact while allowing Calamine to read
            // the unsupported error token as text; the diagnostic preserves
            // the fact that the OOXML cell was typed as an error.
            edits.push((type_attribute.range_value(), b"str".to_vec()));
        } else {
            edits.push((type_attribute.range_value(), b"inlineStr".to_vec()));
            let value_node = value_nodes[0];
            let prefix = xml_element_qname(xml, cell)?;
            let is_name = child_qname(&prefix, "is");
            let text_name = child_qname(&prefix, "t");
            let text = escape_xml_text(value);
            let replacement = format!("<{is_name}><{text_name}>{text}</{text_name}></{is_name}>");
            edits.push((value_node.range(), replacement.into_bytes()));
        }
        normalized_cells.push(UnknownErrorTokenCell {
            sheet_name: sheet_name.to_owned(),
            coordinate: cell.attribute("r").unwrap_or_default().to_owned(),
            token: value.to_owned(),
            formula_cell,
        });
    }
    Ok((apply_xml_edits(xml, edits)?, normalized_cells))
}

fn is_calamine_error_token(value: &str) -> bool {
    matches!(
        value,
        "#DIV/0!" | "#N/A" | "#NAME?" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!"
    )
}

fn xml_element_qname(xml: &[u8], node: roxmltree::Node<'_, '_>) -> Result<String> {
    let fragment = std::str::from_utf8(&xml[node.range()])?;
    let start = fragment
        .find('<')
        .context("XML element has no opening tag")?
        + 1;
    let end = fragment[start..]
        .find(|ch: char| ch.is_whitespace() || matches!(ch, '>' | '/'))
        .map(|offset| start + offset)
        .context("XML element has no name")?;
    Ok(fragment[start..end].to_owned())
}

fn child_qname(parent_qname: &str, child: &str) -> String {
    parent_qname
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:{child}"))
        .unwrap_or_else(|| child.to_owned())
}

fn escape_xml_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn apply_xml_edits(xml: &[u8], mut edits: Vec<(Range<usize>, Vec<u8>)>) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| std::cmp::Reverse((range.start, range.end)));
    let mut output = xml.to_vec();
    for (range, replacement) in edits {
        ensure!(
            range.start <= range.end && range.end <= output.len(),
            "OOXML edit range is out of bounds"
        );
        output.splice(range, replacement);
    }
    Ok(output)
}

fn write_package_with_replacements(
    source: &Path,
    output: &Path,
    replacements: &BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    let mut input = ZipArchive::new(fs::File::open(source)?)?;
    let file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    let mut writer = ZipWriter::new(file);
    for index in 0..input.len() {
        let mut entry = input.by_index(index)?;
        let name = entry.name().to_owned();
        let mut options = SimpleFileOptions::default().compression_method(entry.compression());
        if let Some(modified) = entry.last_modified() {
            options = options.last_modified_time(modified);
        }
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        writer.start_file(&name, options)?;
        if let Some(bytes) = replacements.get(&name) {
            writer.write_all(bytes)?;
        } else {
            std::io::copy(&mut entry, &mut writer)?;
        }
    }
    writer.finish()?;
    Ok(())
}

/// Read worksheet names without reading worksheet cell values.
///
/// This is intended for workflows that identify workbook/sheet-level items
/// from file metadata alone. The reader and supported file formats follow the
/// same Calamine dispatch and legacy XLS signature checks as the value APIs.
pub fn read_sheet_names(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let workbook = open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    ensure!(
        !names.is_empty(),
        "workbook has no worksheets: {}",
        path.display()
    );
    Ok(names)
}

/// Return the OOXML workbook date epoch used for numeric date cells.
///
/// This is intentionally separate from value decoding so callers that compare
/// date-formatted cells with libraries exposing calendar values can normalize
/// the stored serial correctly. Non-OOXML workbooks use the 1900 system in the
/// current adapter contract and return `false`.
pub fn uses_1904_date_system(path: impl AsRef<Path>) -> Result<bool> {
    let path = path.as_ref();
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xlam"
            )
        });
    if !is_ooxml {
        return Ok(false);
    }
    let mut zip = ZipArchive::new(fs::File::open(path).context("open OOXML workbook")?)
        .context("open OOXML ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read workbook date-system metadata")?;
    let document = Document::parse(&xml).context("parse workbook date-system metadata")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );
    let Some(properties) = root.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "workbookPr"
    }) else {
        return Ok(false);
    };
    Ok(matches!(
        properties
            .attribute("date1904")
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str(),
        "1" | "true" | "on"
    ))
}

/// Read the worksheet directory from the standard XLSX `xl/workbook.xml`
/// part using a namespace-aware OOXML parser. This metadata-only adapter does
/// not read cell values, and it deliberately does not fall back to Calamine
/// when the workbook part is missing or malformed.
pub fn read_xlsx_sheet_names_ooxml(path: impl AsRef<Path>) -> Result<Vec<String>> {
    read_xlsx_sheet_names_ooxml_filtered(path.as_ref(), false)
}

/// Read the worksheet directory while ignoring Excel `veryHidden` sheets.
///
/// S1-F01 template matching follows the source Python rule: calculation and
/// reconciliation sheets marked `veryHidden` are not report tabs and must not
/// prevent a report template match. Ordinary `hidden` sheets remain part of
/// this classification rule.
pub fn read_xlsx_sheet_names_ooxml_excluding_very_hidden(
    path: impl AsRef<Path>,
) -> Result<Vec<String>> {
    read_xlsx_sheet_names_ooxml_filtered(path.as_ref(), true)
}

fn read_xlsx_sheet_names_ooxml_filtered(
    path: &Path,
    exclude_very_hidden: bool,
) -> Result<Vec<String>> {
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "OOXML sheet-name adapter only supports .xlsx and .xlsm"
    );

    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read xl/workbook.xml for sheet names")?;
    let document = Document::parse(&xml).context("parse xl/workbook.xml")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );

    let names = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "sheet"
        })
        .filter(|node| !exclude_very_hidden || node.attribute("state") != Some("veryHidden"))
        .filter_map(|node| node.attribute("name"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    ensure!(!names.is_empty(), "workbook has no worksheets");
    Ok(names)
}

/// Read actual cell and merged-range bounds from an OOXML workbook. The
/// `<dimension>` hint is deliberately ignored because some writers leave it
/// stale or broader than the stored cell set. XLSX and XLSM are accepted for
/// this read-only metadata projection; VBA parts are not opened or changed.
pub fn read_xlsx_worksheet_extents(path: impl AsRef<Path>) -> Result<Vec<WorksheetExtent>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "worksheet extent reader only supports .xlsx and .xlsm"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let worksheet_parts = ooxml_worksheet_parts(&mut zip)?;
    let mut seen = std::collections::BTreeSet::new();
    let mut extents = Vec::with_capacity(worksheet_parts.len());
    for (name, part) in worksheet_parts {
        ensure!(
            seen.insert(name.clone()),
            "duplicate worksheet name {name:?}"
        );
        let xml =
            zip_read(&mut zip, &part).with_context(|| format!("read worksheet part {part}"))?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {part}"))?;
        let root = document.root_element();
        ensure!(
            root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && root.tag_name().name() == "worksheet",
            "worksheet part {part} has an unexpected root element"
        );
        let mut extent = WorksheetExtent {
            name,
            ..WorksheetExtent::default()
        };
        for cell in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "c"
        }) {
            let reference = cell
                .attribute("r")
                .with_context(|| format!("cell in {part} has no A1 reference"))?;
            let (row, column) = parse_ooxml_cell_reference(reference)
                .with_context(|| format!("invalid cell reference {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row, column);
        }
        for merged in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "mergeCell"
        }) {
            let reference = merged
                .attribute("ref")
                .with_context(|| format!("merged range in {part} has no reference"))?;
            let (row1, column1, row2, column2) = parse_ooxml_range_reference(reference)
                .with_context(|| format!("invalid merged range {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row1, column1);
            include_extent_coordinate(&mut extent, row2, column2);
        }
        extents.push(extent);
    }
    Ok(extents)
}

fn include_extent_coordinate(extent: &mut WorksheetExtent, row: usize, column: usize) {
    extent.min_row = if extent.min_row == 0 {
        row
    } else {
        extent.min_row.min(row)
    };
    extent.min_column = if extent.min_column == 0 {
        column
    } else {
        extent.min_column.min(column)
    };
    extent.max_row = extent.max_row.max(row);
    extent.max_column = extent.max_column.max(column);
}

fn parse_ooxml_cell_reference(reference: &str) -> Result<(usize, usize)> {
    let reference = reference
        .rsplit_once('!')
        .map(|(_, cell)| cell)
        .unwrap_or(reference)
        .trim()
        .replace('$', "");
    let split = reference
        .find(|character: char| character.is_ascii_digit())
        .context("A1 reference has no row number")?;
    let (letters, digits) = reference.split_at(split);
    ensure!(
        !letters.is_empty() && letters.bytes().all(|byte| byte.is_ascii_alphabetic()),
        "A1 reference has an invalid column"
    );
    ensure!(
        !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()),
        "A1 reference has an invalid row"
    );
    let mut column = 0usize;
    for letter in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| {
                value.checked_add(usize::from(letter.to_ascii_uppercase() - b'A' + 1))
            })
            .context("A1 column index overflows")?;
    }
    let row = digits.parse::<usize>().context("parse A1 row number")?;
    ensure!(row > 0 && column > 0, "A1 coordinates are 1-based");
    Ok((row, column))
}

fn parse_ooxml_range_reference(reference: &str) -> Result<(usize, usize, usize, usize)> {
    let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
    let (row1, column1) = parse_ooxml_cell_reference(first)?;
    let (row2, column2) = parse_ooxml_cell_reference(last)?;
    Ok((row1, column1, row2, column2))
}

/// Read stored values from the first `max_columns` columns of each worksheet.
///
/// This is a value-only API: it does not read formula text or OOXML metadata.
/// Formula cells are represented by the value available from Calamine's
/// worksheet range (normally the cached result); callers that require formula
/// text must use `read_workbook` and inspect `SheetDump::formulas` separately.
pub fn read_value_columns(path: impl AsRef<Path>, max_columns: usize) -> Result<Vec<SheetValues>> {
    read_value_columns_with_diagnostics(path, max_columns).map(|report| report.sheets)
}

/// Read stored values while representing Calamine date-formatted cells as ISO
/// datetime strings. This preserves the workbook's 1900/1904 date epoch for
/// callers that need calendar dates rather than raw serial values.
pub fn read_value_columns_with_iso_dates(
    path: impl AsRef<Path>,
    max_columns: usize,
) -> Result<Vec<SheetValues>> {
    read_value_columns_with_iso_dates_and_diagnostics(path, max_columns).map(|report| report.sheets)
}

pub fn read_value_columns_with_diagnostics(
    path: impl AsRef<Path>,
    max_columns: usize,
) -> Result<ValueColumnsReadReport> {
    read_value_columns_with_mode_and_diagnostics(path.as_ref(), max_columns, false)
}

pub fn read_value_columns_with_iso_dates_and_diagnostics(
    path: impl AsRef<Path>,
    max_columns: usize,
) -> Result<ValueColumnsReadReport> {
    read_value_columns_with_mode_and_diagnostics(path.as_ref(), max_columns, true)
}

fn read_value_columns_with_mode_and_diagnostics(
    path: &Path,
    max_columns: usize,
    iso_dates: bool,
) -> Result<ValueColumnsReadReport> {
    let started = std::time::Instant::now();
    match read_value_columns_calamine(path, max_columns, iso_dates) {
        Ok(sheets) => Ok(ValueColumnsReadReport {
            sheets,
            diagnostics: WorkbookReadDiagnostics {
                primary_reader: "calamine".to_owned(),
                primary_error: None,
                failure_class: None,
                fallback_reader: None,
                fallback_status: "not_needed".to_owned(),
                normalized_error_cells: Vec::new(),
                elapsed_ms: started.elapsed().as_millis(),
            },
        }),
        Err(primary_error) => {
            let failure_class = classify_read_failure(path, &primary_error);
            if failure_class != WorkbookReadFailureClass::UnsupportedCellErrorToken {
                let primary_context = format!("{primary_error:#}");
                return Err(primary_error).context(format!(
                    "Calamine value read failed ({}) for {}; no safe OOXML value fallback applies: {primary_context}",
                    failure_class.as_str(),
                    path.display()
                ));
            }
            let (temporary_copy, normalized_error_cells) =
                create_normalized_error_copy(path).with_context(|| {
                    format!(
                        "Calamine value read failed ({}) for {}: {:#}; targeted OOXML fallback preparation failed",
                        failure_class.as_str(),
                        path.display(),
                        primary_error
                    )
                })?;
            let sheets = read_value_columns_calamine(
                temporary_copy.path(),
                max_columns,
                iso_dates,
            )
            .with_context(|| {
                format!(
                    "Calamine value retry failed after targeted OOXML error-token normalization for {}: primary error: {:#}",
                    path.display(),
                    primary_error
                )
            })?;
            Ok(ValueColumnsReadReport {
                sheets,
                diagnostics: WorkbookReadDiagnostics {
                    primary_reader: "calamine".to_owned(),
                    primary_error: Some(format!("{primary_error:#}")),
                    failure_class: Some(failure_class),
                    fallback_reader: Some("ooxml_error_token_normalization+calamine".to_owned()),
                    fallback_status: "succeeded".to_owned(),
                    normalized_error_cells,
                    elapsed_ms: started.elapsed().as_millis(),
                },
            })
        }
    }
}

fn read_value_columns_calamine(
    path: &Path,
    max_columns: usize,
    iso_dates: bool,
) -> Result<Vec<SheetValues>> {
    anyhow::ensure!(max_columns > 0, "max_columns must be greater than zero");
    validate_legacy_xls_signature(path)?;
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let mut sheets = Vec::with_capacity(names.len());

    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let (max_row, max_column) = range
            .end()
            .map(|(row, column)| (row as usize + 1, column as usize + 1))
            .unwrap_or_default();
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            let absolute_col = start_col as usize + local_col;
            if absolute_col >= max_columns {
                continue;
            }
            if !matches!(value, Data::Empty) {
                let cell_value = match value {
                    Data::DateTime(datetime) if iso_dates && datetime.is_datetime() => {
                        let (year, month, day, hour, minute, second, millisecond) =
                            datetime.to_ymd_hms_milli();
                        CellValue::DateTimeIso(format!(
                            "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}.{millisecond:03}"
                        ))
                    }
                    other => CellValue::from(other),
                };
                values.insert(
                    coord(start_row as usize + local_row, absolute_col),
                    cell_value,
                );
            }
        }
        sheets.push(SheetValues {
            name,
            values,
            max_row,
            max_column,
        });
    }
    Ok(sheets)
}

fn validate_legacy_xls_signature(path: &Path) -> Result<()> {
    let is_xls = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"));
    if !is_xls {
        return Ok(());
    }

    const CFB_OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut signature = [0_u8; 8];
    if file.read_exact(&mut signature).is_err() || signature != CFB_OLE_SIGNATURE {
        anyhow::bail!(
            "格式不支持：.xls 文件的 CFB/OLE 签名无效（{}）",
            path.display()
        );
    }
    Ok(())
}

fn zip_read(zip: &mut ZipArchive<fs::File>, name: &str) -> Result<String> {
    let mut file = zip
        .by_name(name)
        .with_context(|| format!("missing ZIP part {name}"))?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    Ok(text)
}

fn norm_part(base_dir: &str, target: &str) -> String {
    let combined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if base_dir.is_empty() {
        target.to_owned()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts: Vec<&str> = Vec::new();
    for p in combined.split('/') {
        match p {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(p),
        }
    }
    parts.join("/")
}

fn relationships_part(part: &str) -> String {
    let (parent, file_name) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() {
        format!("_rels/{file_name}.rels")
    } else {
        format!("{parent}/_rels/{file_name}.rels")
    }
}

fn ooxml_worksheet_parts(zip: &mut ZipArchive<fs::File>) -> Result<Vec<(String, String)>> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels_doc = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let is_workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            if is_workbook && !is_external {
                relationship.attribute("Target")
            } else {
                None
            }
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_xml = zip_read(zip, &workbook_part)?;
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let workbook_doc = Document::parse(&workbook_xml)?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");

    let mut rel_targets = BTreeMap::new();
    for relationship in rels_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if let (Some(id), Some(kind), Some(target)) = (
            relationship.attribute("Id"),
            relationship.attribute("Type"),
            relationship.attribute("Target"),
        ) {
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            rel_targets.insert(
                id.to_owned(),
                (kind.to_owned(), target.to_owned(), is_external),
            );
        }
    }

    let mut worksheets = Vec::new();
    let mut names = std::collections::BTreeSet::new();
    let mut parts = std::collections::BTreeSet::new();
    for sheet in workbook_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let id = sheet
            .attribute((OFFICE_REL_NS, "id"))
            .with_context(|| format!("workbook sheet {name:?} has no relationship id"))?;
        let (kind, target, is_external) = rel_targets
            .get(id)
            .with_context(|| format!("workbook sheet {name:?} has an unresolved relationship"))?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        ensure!(
            !*is_external,
            "workbook worksheet {name:?} has an external relationship"
        );
        let part = norm_part(workbook_parent, target);
        ensure!(
            !part.is_empty(),
            "workbook worksheet {name:?} has an empty part path"
        );
        ensure!(
            names.insert(name.to_owned()),
            "duplicate worksheet name {name:?}"
        );
        ensure!(
            parts.insert(part.clone()),
            "multiple worksheet names resolve to part {part}"
        );
        worksheets.push((name.to_owned(), part));
    }
    ensure!(
        !worksheets.is_empty(),
        "workbook contains no resolvable worksheet relationships"
    );
    Ok(worksheets)
}

/// Inspect formula/cache node presence without asking Calamine to interpret
/// individual cached values. This is useful after a real calculation engine
/// or a user manually recalculates a workbook.
pub fn inspect_formula_cache_nodes(path: impl AsRef<Path>) -> Result<FormulaCacheInventory> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "formula cache inventory supports XLSX/XLSM only"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open formula cache workbook")?)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut inventory = FormulaCacheInventory::default();
    for (sheet_name, part) in worksheets {
        let xml = zip_read(&mut zip, &part)?.into_bytes();
        let document = Document::parse(std::str::from_utf8(&xml)?)
            .with_context(|| format!("parse formula cache worksheet {part}"))?;
        for cell in document.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "c"
                && node.children().any(|child| {
                    child.is_element()
                        && child.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && child.tag_name().name() == "f"
                })
        }) {
            inventory.formula_cell_count += 1;
            let coordinate = cell
                .attribute("r")
                .with_context(|| format!("formula cell in {part} has no coordinate"))?;
            let has_cache_node = cell.children().any(|child| {
                child.is_element()
                    && child.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && child.tag_name().name() == "v"
            });
            if has_cache_node {
                inventory.cached_formula_cell_count += 1;
            } else {
                inventory.missing_cache_cells.push(FormulaCacheLocation {
                    sheet_name: sheet_name.clone(),
                    cell: coordinate.to_owned(),
                });
            }
        }
    }
    Ok(inventory)
}

fn xml_element_value(node: roxmltree::Node<'_, '_>) -> Value {
    let mut attributes = BTreeMap::new();
    for attribute in node.attributes() {
        let key = match attribute.namespace() {
            Some(namespace) => format!("{{{namespace}}}{}", attribute.name()),
            None => attribute.name().to_owned(),
        };
        attributes.insert(key, attribute.value().to_owned());
    }

    let content = node
        .children()
        .filter_map(|child| {
            if child.is_element() {
                Some(xml_element_value(child))
            } else {
                child
                    .text()
                    .filter(|text| !text.trim().is_empty())
                    .map(|text| json!({"text": text}))
            }
        })
        .collect::<Vec<_>>();

    json!({
        "namespace": node.tag_name().namespace().unwrap_or(""),
        "name": node.tag_name().name(),
        "attributes": attributes,
        "content": content,
    })
}

// Comments, conditional formatting, and data validations are not exposed by Calamine. This
// namespace-aware supplement follows package relationships and records normalized XML structure.
fn extract_ooxml_metadata(path: &Path) -> Result<BTreeMap<String, OoxmlSheetMetadata>> {
    let file = fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut result = BTreeMap::new();
    for (sheet_name, sheet_part) in worksheets {
        let sheet_xml = zip_read(&mut zip, &sheet_part)?;
        let sheet_doc = Document::parse(&sheet_xml)
            .with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let conditional_formatting = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "conditionalFormatting")
            .map(xml_element_value)
            .collect();
        let data_validations = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "dataValidations")
            .map(xml_element_value)
            .collect();

        let parent = sheet_part
            .rsplit_once('/')
            .map(|(parent, _)| parent)
            .unwrap_or("");
        let rels_part = relationships_part(&sheet_part);
        let mut comment_parts = Vec::new();
        if let Ok(sheet_rels_xml) = zip_read(&mut zip, &rels_part) {
            let sheet_rels_doc = Document::parse(&sheet_rels_xml)?;
            for rel in sheet_rels_doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
            {
                let ty = rel.attribute("Type").unwrap_or_default();
                if ty.ends_with("/comments") {
                    if let Some(target) = rel.attribute("Target") {
                        comment_parts.push(norm_part(parent, target));
                    }
                }
            }
        }
        let mut comments = BTreeMap::new();
        let mut comment_records = BTreeMap::new();
        for part in comment_parts {
            let xml = zip_read(&mut zip, &part)?;
            let doc = Document::parse(&xml)?;
            let authors = doc
                .descendants()
                .filter(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && node.tag_name().name() == "author"
                        && node.parent().is_some_and(|parent| {
                            parent.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                                && parent.tag_name().name() == "authors"
                        })
                })
                .map(|author| author.text().unwrap_or_default().to_owned())
                .collect::<Vec<_>>();
            for comment in doc.descendants().filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                    && node.tag_name().name() == "comment"
            }) {
                let Some(reference) = comment.attribute("ref") else {
                    continue;
                };
                let mut text = String::new();
                for t in comment.descendants().filter(|node| {
                    node.is_element()
                        && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                        && node.tag_name().name() == "t"
                }) {
                    if let Some(s) = t.text() {
                        text.push_str(s);
                    }
                }
                let author = comment
                    .attribute("authorId")
                    .and_then(|index| index.parse::<usize>().ok())
                    .and_then(|index| authors.get(index).cloned());
                comments.insert(reference.to_owned(), text.clone());
                comment_records.insert(reference.to_owned(), WorkbookComment { text, author });
            }
        }
        result.insert(
            sheet_name.to_owned(),
            OoxmlSheetMetadata {
                comments,
                comment_records,
                conditional_formatting,
                data_validations,
            },
        );
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use std::path::PathBuf;
    use zip::{write::SimpleFileOptions, ZipWriter};

    fn nonstandard_header_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-header-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/books/main.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="外汇报表" sheetId="1" r:id="worksheet"/><s:sheet name="图表页" sheetId="2" r:id="chartsheet"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="worksheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/fx.xml"/><Relationship Id="chartsheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="../charts/chart1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/><s:headerFooter differentOddEven="1" differentFirst="1"><s:oddHeader>&amp;L单位：万元</s:oddHeader><s:evenHeader>&amp;C第二页</s:evenHeader><s:firstHeader/></s:headerFooter></s:worksheet>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("custom/books/main.xml", workbook),
            ("custom/books/_rels/main.xml.rels", workbook_relationships),
            ("custom/worksheets/fx.xml", worksheet),
            ("custom/charts/chart1.xml", "<chart/>".to_owned()),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn comment_records_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-comment-records-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet =
            format!(r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/></s:worksheet>"#);
        let worksheet_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comments1.xml"/></Relationships>"#
        );
        let comments = format!(
            r#"<x:comments xmlns:x="{SPREADSHEET_MAIN_NS}"><x:authors><x:author>Data Team</x:author><x:author></x:author></x:authors><x:commentList><x:comment ref="A1" authorId="0"><x:text><x:r><x:t>First </x:t></x:r><x:r><x:t>note</x:t></x:r></x:text></x:comment><x:comment ref="B2" authorId="1"><x:text><x:t>Empty author</x:t></x:text></x:comment><x:comment ref="C3" authorId="9"><x:text><x:t>Unknown author</x:t></x:text></x:comment></x:commentList></x:comments>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_relationships),
            ("xl/worksheets/sheet1.xml", worksheet),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                worksheet_relationships,
            ),
            ("xl/comments/comments1.xml", comments),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn sheet_name_ooxml_fixture(extension: &str) -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}-{}.{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos(),
            extension
        ));
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:x="urn:unrelated"><s:sheets><s:sheet name="单位贷款" sheetId="1"/><s:sheet name="普通隐藏表" sheetId="2" state="hidden"/><s:sheet name="内部计算表" sheetId="3" state="veryHidden"/><x:sheet name="must-be-ignored" sheetId="4"/></s:sheets></s:workbook>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        archive.start_file("xl/workbook.xml", SimpleFileOptions::default())?;
        archive.write_all(workbook.as_bytes())?;
        archive.finish()?;
        Ok(path)
    }

    fn conditional_formatting_areas_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-cf-areas-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="单位存款" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/><s:conditionalFormatting sqref="$A$1:$B$2 D4"><s:cfRule type="expression" priority="1"><s:formula>A1&gt;0</s:formula></s:cfRule></s:conditionalFormatting><s:conditionalFormatting sqref="D4"><s:cfRule type="cellIs" priority="2" operator="greaterThan"><s:formula>5</s:formula></s:cfRule></s:conditionalFormatting></s:worksheet>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn defined_names_ooxml_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-defined-names-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let root_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/book.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Data Set" sheetId="1" r:id="rId1"/><s:sheet name="Chart" sheetId="2" r:id="rId2"/><s:sheet name="存量单位贷款信息" sheetId="3" r:id="rId3"/></s:sheets><s:definedNames><s:definedName name="GlobalRange">'Data Set'!$A$1:$B$8</s:definedName><s:definedName name="LocalUnion" localSheetId="0">'Data Set'!A1:B2,'Data Set'!D4</s:definedName><s:definedName name="Formula">SUM('Data Set'!A:A)</s:definedName><s:definedName name="UnicodeRange">存量单位贷款信息!$A$4:$B$90</s:definedName><s:definedName name="UnicodeLocal" localSheetId="2">存量单位贷款信息!A1:B2</s:definedName><s:definedName name="Broken" localSheetId="9">'Missing'!A1</s:definedName></s:definedNames></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/c1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s3.xml"/></Relationships>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", root_rels),
            ("custom/book.xml", workbook),
            ("custom/_rels/book.xml.rels", workbook_rels),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn worksheet_extent_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-worksheet-extents-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Layout" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:dimension ref="A1:XFD1048576"/><s:sheetData><s:row r="3"><s:c r="B3"><s:v>4</s:v></s:c></s:row><s:row r="8"><s:c r="G8" s="1"/></s:row><s:row r="10"><s:c r="C10" t="inlineStr"><s:is><s:t>merged anchor</s:t></s:is></s:c></s:row></s:sheetData><s:mergeCells count="1"><s:mergeCell ref="C10:D12"/></s:mergeCells></s:worksheet>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", package_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn unknown_error_token_fixture() -> Result<(PathBuf, Vec<u8>)> {
        let source = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let original = fs::read(&source)?;
        let path = std::env::temp_dir().join(format!(
            "excel-reader-error-token-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let mut archive = ZipArchive::new(fs::File::open(&source)?)?;
        let worksheet = zip_read(&mut archive, "xl/worksheets/sheet1.xml")?;
        let marker = "<c r=\"J1\"";
        ensure!(
            worksheet.contains(marker),
            "synthetic worksheet marker missing"
        );
        let patched = worksheet.replacen(
            marker,
            concat!(
                "<c r=\"B1\" t=\"e\"><v>#CUSTOMERR!</v></c>",
                "<c r=\"C1\" t=\"e\"><f>1+1</f><v>#CUSTOMERR!</v></c>",
                "<c r=\"D1\"><f>1+2</f></c>",
                "<c r=\"J1\""
            ),
            1,
        );
        let replacements =
            BTreeMap::from([("xl/worksheets/sheet1.xml".to_owned(), patched.into_bytes())]);
        write_package_with_replacements(&source, &path, &replacements)?;
        Ok((path, original))
    }

    #[test]
    fn cell_value_json_shape_remains_compatible_and_typed() {
        let cases = [
            (
                Data::String("001.20".to_owned()),
                json!({"type":"string", "value":"001.20"}),
            ),
            (Data::Float(1.0), json!({"type":"number", "value":1.0})),
            (Data::Int(42), json!({"type":"integer", "value":42})),
            (Data::Bool(true), json!({"type":"boolean", "value":true})),
        ];

        for (data, expected) in cases {
            assert_eq!(value_json(&data), Some(expected));
        }
        assert_eq!(value_json(&Data::Empty), None);
        assert_eq!(
            value_json(&Data::Float(f64::NAN)),
            Some(json!({"type":"number", "value":null}))
        );

        let decoded = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": 12.5
        }))
        .unwrap();
        assert_eq!(decoded, CellValue::Number(12.5));
        let non_finite = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": null
        }))
        .unwrap();
        assert!(matches!(non_finite, CellValue::Number(value) if value.is_nan()));
    }

    #[test]
    fn ooxml_sheet_name_adapter_matches_namespace_and_workbook_order() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture("xlsx")?;
        let names = read_xlsx_sheet_names_ooxml(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(
            names,
            vec![
                "单位贷款".to_owned(),
                "普通隐藏表".to_owned(),
                "内部计算表".to_owned()
            ]
        );
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_excludes_only_very_hidden_sheets() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture("xlsx")?;
        let names = read_xlsx_sheet_names_ooxml_excluding_very_hidden(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(names, vec!["单位贷款".to_owned(), "普通隐藏表".to_owned()]);
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_accepts_xlsm_metadata() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture("xlsm")?;
        let names = read_xlsx_sheet_names_ooxml(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(
            names,
            vec![
                "单位贷款".to_owned(),
                "普通隐藏表".to_owned(),
                "内部计算表".to_owned()
            ]
        );
        Ok(())
    }

    #[test]
    fn ooxml_conditional_format_area_reader_deduplicates_rule_targets() -> Result<()> {
        let fixture = conditional_formatting_areas_fixture()?;
        let areas = read_conditional_formatting_areas(&fixture)?;
        fs::remove_file(fixture)?;

        assert_eq!(
            areas,
            vec![
                ConditionalFormattingArea {
                    sheet_name: "单位存款".to_owned(),
                    address: "$A$1:$B$2".to_owned(),
                },
                ConditionalFormattingArea {
                    sheet_name: "单位存款".to_owned(),
                    address: "D4".to_owned(),
                },
            ]
        );
        Ok(())
    }

    #[test]
    fn defined_name_reader_preserves_scope_raw_text_and_resolves_area_sheet_identity() -> Result<()>
    {
        let fixture = defined_names_ooxml_fixture()?;
        let names = read_defined_names(&fixture)?;
        fs::remove_file(fixture)?;

        assert_eq!(names.len(), 6);
        assert_eq!(names[0].name, "GlobalRange");
        assert_eq!(names[0].scope, DefinedNameScope::Workbook);
        assert_eq!(names[0].raw_reference, "'Data Set'!$A$1:$B$8");
        assert_eq!(names[0].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(
            names[0].areas,
            vec![DefinedNameArea {
                sheet_name: "Data Set".to_owned(),
                sheet_kind: "worksheet".to_owned(),
                reference: "$A$1:$B$8".to_owned(),
            }]
        );
        assert_eq!(
            names[1].scope,
            DefinedNameScope::Worksheet {
                local_sheet_id: 0,
                sheet_name: Some("Data Set".to_owned()),
                sheet_kind: Some("worksheet".to_owned()),
            }
        );
        assert_eq!(names[1].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(names[1].areas.len(), 2);
        assert_eq!(
            names[2].parse_status,
            DefinedNameParseStatus::NonAreaExpression
        );
        assert_eq!(names[2].raw_reference, "SUM('Data Set'!A:A)");
        assert_eq!(names[3].name, "UnicodeRange");
        assert_eq!(names[3].parse_status, DefinedNameParseStatus::Areas);
        assert_eq!(
            names[3].areas,
            vec![DefinedNameArea {
                sheet_name: "存量单位贷款信息".to_owned(),
                sheet_kind: "worksheet".to_owned(),
                reference: "$A$4:$B$90".to_owned(),
            }]
        );
        assert_eq!(
            names[4].scope,
            DefinedNameScope::Worksheet {
                local_sheet_id: 2,
                sheet_name: Some("存量单位贷款信息".to_owned()),
                sheet_kind: Some("worksheet".to_owned()),
            }
        );
        assert_eq!(
            names[5].scope,
            DefinedNameScope::UnresolvedWorksheet { local_sheet_id: 9 }
        );
        assert_eq!(
            names[5].parse_status,
            DefinedNameParseStatus::InvalidAreaExpression
        );
        assert!(decode_sheet_token("存量单位贷款信息").is_some());
        assert!(decode_sheet_token("[Other.xlsx]Data").is_none());
        assert!(decode_sheet_token("Data:Sheet").is_none());
        Ok(())
    }

    #[test]
    fn worksheet_extents_include_styled_blanks_and_merge_coordinates_not_dimension_hint(
    ) -> Result<()> {
        let fixture = worksheet_extent_fixture()?;
        let extents = read_xlsx_worksheet_extents(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(
            extents,
            vec![WorksheetExtent {
                name: "Layout".to_owned(),
                min_row: 3,
                min_column: 2,
                max_row: 12,
                max_column: 7,
            }]
        );
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_rejects_biff_formats_without_fallback() -> Result<()> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}.xls",
            std::process::id()
        ));
        fs::write(&path, b"not an OOXML package")?;
        let error = read_xlsx_sheet_names_ooxml(&path).expect_err("reject unsupported format");
        fs::remove_file(path)?;
        assert!(error.to_string().contains("only supports .xlsx and .xlsm"));
        Ok(())
    }

    #[test]
    fn value_column_projection_matches_dump_and_keeps_absolute_coordinates() {
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let dump = read_workbook(&fixture).expect("read full probe fixture");
        let projected = read_value_columns(&fixture, 3).expect("read first three columns");

        assert_eq!(projected.len(), dump.sheets.len());
        for (projection, sheet) in projected.iter().zip(&dump.sheets) {
            assert_eq!(projection.name, sheet.name);
            let expected: BTreeMap<_, _> = sheet
                .values
                .iter()
                .filter(|(address, _)| {
                    address
                        .chars()
                        .take_while(|character| character.is_ascii_alphabetic())
                        .fold(0_usize, |column, character| {
                            column * 26
                                + (character.to_ascii_uppercase() as usize - 'A' as usize + 1)
                        })
                        <= 3
                })
                .map(|(address, value)| {
                    (
                        address.clone(),
                        CellValue::from_probe_json(value).expect("decode dumped cell value"),
                    )
                })
                .collect();
            assert_eq!(projection.values, expected);
        }
    }

    #[test]
    fn structure_projection_matches_full_dump_for_values_formulas_and_merges() -> Result<()> {
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let dump = read_workbook(&fixture)?;
        let regions = dump
            .sheets
            .iter()
            .map(|sheet| CellProjectionRegion {
                sheet_name: sheet.name.clone(),
                start_row: 1,
                start_column: 1,
                end_row: 1_048_576,
                end_column: 16_384,
            })
            .collect::<Vec<_>>();
        let projection = read_workbook_structure_projection(&fixture, &regions, false)?;

        assert_eq!(
            projection.sheet_names,
            dump.sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect::<Vec<_>>()
        );
        assert!(!projection.fallback_used);
        for full_sheet in &dump.sheets {
            let projected = projection
                .sheets
                .iter()
                .find(|sheet| sheet.name == full_sheet.name)
                .with_context(|| format!("missing projected sheet {}", full_sheet.name))?;
            let projected_values = projected
                .values
                .iter()
                .map(|(address, value)| {
                    (
                        address.clone(),
                        serde_json::to_value(value).expect("serialize projected value"),
                    )
                })
                .collect::<BTreeMap<_, _>>();
            assert_eq!(
                projected_values, full_sheet.values,
                "values in {}",
                full_sheet.name
            );
            assert_eq!(
                projected.formulas, full_sheet.formulas,
                "formulas in {}",
                full_sheet.name
            );
            assert_eq!(
                projected.merges, full_sheet.merges,
                "merges in {}",
                full_sheet.name
            );
        }
        Ok(())
    }

    #[test]
    fn structure_projection_reads_merge_anchor_for_non_anchor_request() -> Result<()> {
        let fixture = worksheet_extent_fixture()?;
        let projection = read_workbook_structure_projection(
            &fixture,
            &[CellProjectionRegion {
                sheet_name: "Layout".to_owned(),
                start_row: 11,
                start_column: 4,
                end_row: 11,
                end_column: 4,
            }],
            true,
        )?;
        let sheet = projection
            .sheets
            .iter()
            .find(|sheet| sheet.name == "Layout")
            .context("projected Layout sheet missing")?;
        assert_eq!(sheet.merges, vec!["C10:D12"]);
        assert_eq!(
            sheet.values.get("C10"),
            Some(&CellValue::Text("merged anchor".to_owned()))
        );
        assert!(!sheet.values.contains_key("D11"));
        fs::remove_file(fixture)?;
        Ok(())
    }

    #[test]
    fn structure_projection_reuses_unknown_error_token_fallback() -> Result<()> {
        let (fixture, _) = unknown_error_token_fixture()?;
        let projection = read_workbook_structure_projection(
            &fixture,
            &[CellProjectionRegion {
                sheet_name: "Synthetic".to_owned(),
                start_row: 1,
                start_column: 1,
                end_row: 1,
                end_column: 10,
            }],
            false,
        )?;
        let sheet = projection
            .sheets
            .iter()
            .find(|sheet| sheet.name == "Synthetic")
            .context("projected Synthetic sheet missing")?;
        assert!(projection.fallback_used);
        assert_eq!(
            sheet.values.get("B1"),
            Some(&CellValue::Text("#CUSTOMERR!".to_owned()))
        );
        assert_eq!(sheet.formulas.get("C1").map(String::as_str), Some("1+1"));
        fs::remove_file(fixture)?;
        Ok(())
    }

    #[test]
    fn unknown_error_tokens_use_cell_scoped_ooxml_normalization_and_preserve_formula_text(
    ) -> Result<()> {
        let (fixture, original) = unknown_error_token_fixture()?;
        let original_source = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let report = read_workbook_with_diagnostics(&fixture)?;
        assert_eq!(report.diagnostics.primary_reader, "calamine");
        assert_eq!(
            report.diagnostics.failure_class,
            Some(WorkbookReadFailureClass::UnsupportedCellErrorToken)
        );
        assert_eq!(
            report.diagnostics.fallback_reader.as_deref(),
            Some("ooxml_error_token_normalization+calamine")
        );
        assert_eq!(report.diagnostics.fallback_status, "succeeded");
        assert_eq!(report.diagnostics.normalized_error_cells.len(), 2);
        assert!(report
            .diagnostics
            .normalized_error_cells
            .iter()
            .any(|cell| {
                cell.sheet_name == "Synthetic"
                    && cell.coordinate == "B1"
                    && cell.token == "#CUSTOMERR!"
                    && !cell.formula_cell
            }));
        assert!(report
            .diagnostics
            .normalized_error_cells
            .iter()
            .any(|cell| { cell.coordinate == "C1" && cell.formula_cell }));
        let sheet = report
            .workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "Synthetic")
            .context("Synthetic sheet missing")?;
        assert_eq!(
            sheet.values["B1"],
            json!({"type":"string", "value":"#CUSTOMERR!"})
        );
        assert_eq!(
            sheet.values["C1"],
            json!({"type":"string", "value":"#CUSTOMERR!"})
        );
        assert_eq!(sheet.formulas["C1"], "1+1");
        assert_eq!(sheet.formulas["D1"], "1+2");
        assert_eq!(
            sheet.values["A1"]["type"], "number",
            "unrelated ordinary data remains unchanged"
        );
        assert_eq!(fs::read(&original_source)?, original);

        let value_report = read_value_columns_with_diagnostics(&fixture, 4)?;
        assert_eq!(value_report.diagnostics.fallback_status, "succeeded");
        assert_eq!(
            value_report.sheets[0].values["B1"],
            CellValue::Text("#CUSTOMERR!".to_owned())
        );
        let baseline_values = read_value_columns(&original_source, 4)?;
        assert_eq!(
            value_report.sheets[0].values["A1"], baseline_values[0].values["A1"],
            "fallback preserves ordinary stored data exactly"
        );
        let _ = fs::remove_file(fixture);
        Ok(())
    }

    #[test]
    fn formula_cache_inventory_distinguishes_cache_nodes_from_missing_nodes() -> Result<()> {
        let (fixture, _) = unknown_error_token_fixture()?;
        let inventory = inspect_formula_cache_nodes(&fixture)?;
        assert_eq!(inventory.formula_cell_count, 3);
        assert_eq!(inventory.cached_formula_cell_count, 2);
        assert_eq!(
            inventory.missing_cache_cells,
            vec![FormulaCacheLocation {
                sheet_name: "Synthetic".to_owned(),
                cell: "D1".to_owned(),
            }]
        );
        let _ = fs::remove_file(fixture);
        Ok(())
    }

    #[test]
    fn formula_result_coverage_accepts_typed_errors_and_empty_text_only() {
        let workbook = WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "Audit".to_owned(),
                values: BTreeMap::from([
                    ("A1".to_owned(), json!({"type": "error", "value": "Div0"})),
                    ("A2".to_owned(), json!({"type": "string", "value": ""})),
                    ("A3".to_owned(), json!({"type": "empty"})),
                ]),
                formulas: BTreeMap::from([
                    ("A1".to_owned(), "1/0".to_owned()),
                    ("A2".to_owned(), "IF(TRUE,\"\",\"x\")".to_owned()),
                    ("A3".to_owned(), "A4".to_owned()),
                    ("A4".to_owned(), "A5".to_owned()),
                ]),
                formula_status: "ok".to_owned(),
                merges: Vec::new(),
                merge_status: "ok".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "ok".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "ok".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "ok".to_owned(),
            }],
        };

        let coverage = formula_result_value_coverage(&workbook);
        assert_eq!(coverage.formula_cell_count, 4);
        assert_eq!(coverage.typed_result_cell_count, 2);
        assert_eq!(
            coverage.missing_result_cells,
            vec![
                FormulaResultValueLocation {
                    sheet_name: "Audit".to_owned(),
                    cell: "A3".to_owned(),
                },
                FormulaResultValueLocation {
                    sheet_name: "Audit".to_owned(),
                    cell: "A4".to_owned(),
                },
            ]
        );
    }

    #[test]
    fn formula_comparison_allows_only_libreoffice_boolean_literal_normalization() {
        assert_eq!(
            canonicalize_formula_boolean_literals("IF(A1=TRUE(),B1,FALSE())"),
            "IF(A1=TRUE,B1,FALSE)"
        );
        assert_eq!(
            canonicalize_formula_boolean_literals("=\"TRUE()\"&FALSE()"),
            "=\"TRUE()\"&FALSE"
        );
        assert_eq!(
            canonicalize_formula_boolean_literals("=MYTRUE()+FALSE_VALUE"),
            "=MYTRUE()+FALSE_VALUE"
        );
        assert_eq!(
            canonicalize_formula_boolean_literals("'TRUE()'!A1"),
            "'TRUE()'!A1"
        );

        let workbook = |formula: &str| WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "Audit".to_owned(),
                values: BTreeMap::new(),
                formulas: BTreeMap::from([("D2".to_owned(), formula.to_owned())]),
                formula_status: "ok".to_owned(),
                merges: Vec::new(),
                merge_status: "ok".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "ok".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "ok".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "ok".to_owned(),
            }],
        };
        assert!(formula_texts_equivalent_after_recalculation(
            &workbook("VLOOKUP(A1,Ref!A:B,2,FALSE)"),
            &workbook("VLOOKUP(A1,Ref!A:B,2,FALSE())")
        ));
        assert!(!formula_texts_equivalent_after_recalculation(
            &workbook("VLOOKUP(A1,Ref!A:B,2,FALSE)"),
            &workbook("VLOOKUP(A1,Other!A:B,2,FALSE())")
        ));
        assert!(!formula_texts_equivalent_after_recalculation(
            &workbook("'TRUE()'!A1"),
            &workbook("'TRUE'!A1")
        ));

        assert_eq!(
            canonicalize_formula_sheet_reference_quoting(
                "=Data!A1+'Data (prior)'!A1+\"Data!A1\"+MyData!A1",
                &["Data", "Data (prior)"]
            ),
            "='Data'!A1+'Data (prior)'!A1+\"Data!A1\"+MyData!A1"
        );

        let workbook_with_references = |formula: &str| WorkbookDump {
            schema_version: 1,
            reader: "test",
            elapsed_ms: 0,
            sheets: vec![
                SheetDump {
                    name: "Audit".to_owned(),
                    values: BTreeMap::new(),
                    formulas: BTreeMap::from([("D2".to_owned(), formula.to_owned())]),
                    formula_status: "ok".to_owned(),
                    merges: Vec::new(),
                    merge_status: "ok".to_owned(),
                    comments: BTreeMap::new(),
                    comments_status: "ok".to_owned(),
                    conditional_formatting: Vec::new(),
                    conditional_formatting_status: "ok".to_owned(),
                    data_validations: Vec::new(),
                    data_validations_status: "ok".to_owned(),
                },
                SheetDump {
                    name: "Data".to_owned(),
                    values: BTreeMap::new(),
                    formulas: BTreeMap::new(),
                    formula_status: "ok".to_owned(),
                    merges: Vec::new(),
                    merge_status: "ok".to_owned(),
                    comments: BTreeMap::new(),
                    comments_status: "ok".to_owned(),
                    conditional_formatting: Vec::new(),
                    conditional_formatting_status: "ok".to_owned(),
                    data_validations: Vec::new(),
                    data_validations_status: "ok".to_owned(),
                },
                SheetDump {
                    name: "Data (prior)".to_owned(),
                    values: BTreeMap::new(),
                    formulas: BTreeMap::new(),
                    formula_status: "ok".to_owned(),
                    merges: Vec::new(),
                    merge_status: "ok".to_owned(),
                    comments: BTreeMap::new(),
                    comments_status: "ok".to_owned(),
                    conditional_formatting: Vec::new(),
                    conditional_formatting_status: "ok".to_owned(),
                    data_validations: Vec::new(),
                    data_validations_status: "ok".to_owned(),
                },
            ],
        };
        assert!(formula_texts_equivalent_after_recalculation(
            &workbook_with_references("IF(Data!A1=1,'Data (prior)'!A1,Data!B2)"),
            &workbook_with_references("IF('Data'!A1=1,'Data (prior)'!A1,'Data'!B2)")
        ));
        assert!(!formula_texts_equivalent_after_recalculation(
            &workbook_with_references("Data!A1"),
            &workbook_with_references("'Data (prior)'!A1")
        ));
    }

    #[test]
    fn formula_comparison_accepts_excel_quoting_for_parenthesized_sheet_names() {
        let sheet_names = ["存量个人贷款信息", "个人贷款发生额信息（隐藏发放）"];
        let before =
            "=IF(ISBLANK(个人贷款发生额信息（隐藏发放）!C10),0,个人贷款发生额信息（隐藏发放）!C10)";
        let after = "=IF(ISBLANK('个人贷款发生额信息（隐藏发放）'!C10),0,'个人贷款发生额信息（隐藏发放）'!C10)";
        assert!(formula_text_equivalent_after_recalculation(
            before,
            after,
            &sheet_names,
        ));
        assert!(!formula_text_equivalent_after_recalculation(
            before,
            &after.replace("!C10", "!D10"),
            &sheet_names,
        ));
    }

    #[test]
    fn unsafe_unknown_error_token_xml_is_rejected_instead_of_silently_normalized() {
        let xml = format!(
            r#"<worksheet xmlns="{SPREADSHEET_MAIN_NS}"><sheetData><row r="1"><c t="e"><v>#CUSTOMERR!</v></c></row></sheetData></worksheet>"#
        );
        let error = normalize_unknown_error_cells(xml.as_bytes(), "Synthetic")
            .expect_err("an error cell without a coordinate is unsafe to recover");
        assert!(error.to_string().contains("missing a coordinate"));
    }

    #[test]
    fn invalid_legacy_xls_signature_is_rejected_as_unsupported_format() {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-invalid-xls-{}.xls",
            std::process::id()
        ));
        fs::write(&path, b"<html>not an xls</html>").expect("write temporary pseudo XLS");
        let error = read_value_columns(&path, 3).expect_err("reject invalid CFB/OLE signature");
        let _ = fs::remove_file(&path);
        assert!(error.to_string().contains("格式不支持"));
        assert!(error.to_string().contains("CFB/OLE"));
    }

    #[test]
    fn workbook_read_failures_are_classified_without_broadening_fallback() -> Result<()> {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)?
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "excel-reader-failure-classes-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir_all(&root)?;
        let missing = root.join("missing.xlsx");
        let missing_error = anyhow::Error::new(std::io::Error::new(
            std::io::ErrorKind::NotFound,
            "missing workbook",
        ));
        assert_eq!(
            classify_read_failure(&missing, &missing_error),
            WorkbookReadFailureClass::FileMissing
        );
        let permission_error = anyhow::Error::new(std::io::Error::new(
            std::io::ErrorKind::PermissionDenied,
            "workbook is locked",
        ));
        assert_eq!(
            classify_read_failure(&missing, &permission_error),
            WorkbookReadFailureClass::PermissionDenied
        );

        let unsupported = root.join("input.csv");
        fs::write(&unsupported, b"a,b")?;
        assert_eq!(
            classify_read_failure(&unsupported, &anyhow::Error::msg("reader failed")),
            WorkbookReadFailureClass::UnsupportedFormat
        );

        let worksheet = root.join("worksheet.xlsx");
        fs::write(&worksheet, b"fixture")?;
        assert_eq!(
            classify_read_failure(
                &worksheet,
                &anyhow::Error::msg("read values for Data: invalid cell value")
            ),
            WorkbookReadFailureClass::WorksheetRead
        );
        assert_eq!(
            classify_read_failure(&worksheet, &anyhow::Error::msg("malformed ZIP directory")),
            WorkbookReadFailureClass::CorruptWorkbook
        );
        assert_eq!(
            classify_read_failure(
                &worksheet,
                &anyhow::Error::msg("unsupported cell error value: #CUSTOMERR!")
            ),
            WorkbookReadFailureClass::UnsupportedCellErrorToken
        );

        let unsupported_reader = root.join("input.ods");
        fs::write(&unsupported_reader, b"fixture")?;
        assert_eq!(
            classify_read_failure(
                &unsupported_reader,
                &anyhow::Error::msg("reader failed without a known category")
            ),
            WorkbookReadFailureClass::Unknown
        );

        let missing_read = read_workbook_with_diagnostics(&missing)
            .expect_err("a missing workbook must not be sent to the OOXML fallback");
        let diagnostic = format!("{missing_read:#}");
        assert!(diagnostic.contains("file_missing"));
        assert!(diagnostic.contains("no safe OOXML value fallback applies"));
        fs::remove_dir_all(root)?;
        Ok(())
    }

    #[test]
    fn normalized_ooxml_keeps_namespaces_attributes_and_text() {
        let xml = r#"<x:conditionalFormatting xmlns:x="urn:example:sheet" sqref="A1">
            <x:cfRule type="expression"><x:formula>A1&gt;0</x:formula></x:cfRule>
        </x:conditionalFormatting>"#;
        let document = Document::parse(xml).unwrap();
        let value = xml_element_value(document.root_element());

        assert_eq!(value["namespace"], "urn:example:sheet");
        assert_eq!(value["attributes"]["sqref"], "A1");
        assert_eq!(value["content"][0]["name"], "cfRule");
        assert_eq!(
            value["content"][0]["content"][0]["content"][0]["text"],
            "A1>0"
        );
    }

    #[test]
    fn ooxml_comment_reader_returns_text_and_resolved_authors() -> Result<()> {
        let fixture = comment_records_fixture()?;
        let comments = read_ooxml_comments(&fixture)?;
        fs::remove_file(fixture)?;

        let sheet = comments.get("Data").context("worksheet comments missing")?;
        assert_eq!(
            sheet.get("A1"),
            Some(&WorkbookComment {
                text: "First note".to_owned(),
                author: Some("Data Team".to_owned()),
            })
        );
        assert_eq!(
            sheet.get("B2"),
            Some(&WorkbookComment {
                text: "Empty author".to_owned(),
                author: Some(String::new()),
            })
        );
        assert_eq!(
            sheet.get("C3"),
            Some(&WorkbookComment {
                text: "Unknown author".to_owned(),
                author: None,
            })
        );
        Ok(())
    }

    #[test]
    fn ooxml_comment_reader_rejects_non_ooxml_formats() {
        let error = read_ooxml_comments("not-a-workbook.xls")
            .expect_err("the comment supplement does not read BIFF8 workbooks");
        assert!(error.to_string().contains("only supports"));
    }

    #[test]
    fn relationship_targets_are_resolved_from_their_source_part() {
        assert_eq!(
            norm_part("custom/worksheets", "../comments/comments1.xml"),
            "custom/comments/comments1.xml"
        );
        assert_eq!(norm_part("", "/custom/workbook.xml"), "custom/workbook.xml");
        assert_eq!(
            relationships_part("custom/worksheets/sheet1.xml"),
            "custom/worksheets/_rels/sheet1.xml.rels"
        );
    }

    #[test]
    fn header_inventory_resolves_custom_workbook_part_and_excludes_chartsheets() -> Result<()> {
        let fixture = nonstandard_header_fixture()?;
        let inventory = read_header_inventory(&fixture)?;

        assert_eq!(inventory.len(), 1);
        assert_eq!(inventory[0].sheet_name, "外汇报表");
        assert_eq!(
            inventory[0].headers,
            vec![
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Odd,
                    text: "&L单位：万元".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Even,
                    text: "&C第二页".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::First,
                    text: String::new(),
                },
            ]
        );

        fs::remove_file(fixture)?;
        Ok(())
    }
}
