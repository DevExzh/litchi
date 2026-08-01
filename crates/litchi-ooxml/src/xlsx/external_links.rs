//! Read-only SpreadsheetML external-workbook link metadata and cached values.

use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_ooxml_common::external_link::{
    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES, is_external_workbook_relationship,
};
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SHEET_NAMES: usize = 65_536;
const MAX_DEFINED_NAMES: usize = 65_536;
const MAX_CACHED_SHEETS: usize = 65_536;
const MAX_CACHED_ROWS: usize = 1_048_576;
const MAX_CACHED_CELLS: usize = 1_000_000;
const MAX_LINK_ITEMS: usize = 65_536;
const MAX_CACHE_TEXT_BYTES: usize = 64 * 1024 * 1024;
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const TRANSITIONAL_SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_EXTERNAL_TARGET_BYTES: usize = 32 * 1024;
/// Highest column index addressable by a SpreadsheetML cell reference (`XFD`).
const MAX_CELL_COLUMN: u32 = 16_384;
/// Highest row index addressable by a SpreadsheetML cell reference.
const MAX_CELL_ROW: u32 = 1_048_576;
/// Longest column prefix a valid reference can carry (`XFD` is three letters).
const MAX_COLUMN_LETTERS: usize = 3;

/// Namespace conformance used when authoring an external-link part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ExternalLinkConformance {
    #[default]
    Transitional,
    Strict,
}

impl ExternalLinkConformance {
    fn sml(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_SML,
            Self::Strict => STRICT_SML,
        }
    }

    fn rel(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }

    pub(crate) fn external_link_relationship(self) -> &'static str {
        match self {
            Self::Transitional => rt::EXTERNAL_LINK,
            Self::Strict => rt::STRICT_EXTERNAL_LINK,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalLinkEntry {
    pub index: u32,
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub kind: ExternalLinkKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalLinkKind {
    Workbook(ExternalWorkbookLink),
    Dde(ExternalDdeLink),
    Ole(ExternalOleLink),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeLink {
    pub service: String,
    pub topic: String,
    pub items: Vec<ExternalDdeItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleLink {
    pub target: ExternalOleTarget,
    pub program_id: String,
    pub items: Vec<ExternalOleItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleTarget {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeItem {
    pub name: Option<String>,
    pub use_ole: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<ExternalDdeValues>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalOleItemSource {
    SpreadsheetMl,
    Office2010,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleItem {
    pub source: ExternalOleItemSource,
    pub name: String,
    pub icon: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<ExternalDdeValues>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeValues {
    pub rows: u32,
    pub columns: u32,
    pub values: Vec<ExternalDdeValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalDdeValueType {
    Nil,
    Boolean,
    Number,
    Error,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeValue {
    pub value_type: ExternalDdeValueType,
    pub raw_value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalWorkbookLink {
    pub target: ExternalWorkbookTarget,
    pub sheet_names: Vec<String>,
    pub defined_names: Vec<ExternalDefinedName>,
    pub cached_sheets: Vec<ExternalSheetData>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalWorkbookTarget {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDefinedName {
    pub name: String,
    pub refers_to: Option<String>,
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalSheetData {
    pub sheet_id: u32,
    pub refresh_error: bool,
    pub rows: Vec<ExternalRow>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalRow {
    pub row: u32,
    pub cells: Vec<ExternalCell>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalCellType {
    Number,
    Boolean,
    Date,
    Error,
    InlineString,
    SharedString,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalCell {
    pub reference: Option<String>,
    pub cell_type: ExternalCellType,
    pub raw_value: Option<String>,
    pub value_metadata_index: u32,
}

impl ExternalLinkKind {
    /// Serialize this link as a canonical transitional SpreadsheetML external-link part.
    ///
    /// External targets are represented only as OPC relationship metadata and are never opened.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_conformance(ExternalLinkConformance::Transitional)
    }

    /// Serialize this link without dereferencing or executing any target metadata.
    pub fn to_xml_with_conformance(&self, conformance: ExternalLinkConformance) -> Result<Vec<u8>> {
        let has_x14 = matches!(self, Self::Ole(link) if link.items.iter().any(|item| item.source == ExternalOleItemSource::Office2010));
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="{}""#,
            conformance.sml()
        );
        if matches!(self, Self::Workbook(_) | Self::Ole(_)) {
            xml.push_str(r#" xmlns:r="#);
            xml.push('"');
            xml.push_str(conformance.rel());
            xml.push('"');
        }
        if has_x14 {
            xml.push_str(
                r#" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""#,
            );
        }
        xml.push('>');
        match self {
            Self::Workbook(link) => write_external_workbook(&mut xml, link)?,
            Self::Dde(link) => write_dde_link(&mut xml, link)?,
            Self::Ole(link) => write_ole_link(&mut xml, link)?,
        }
        xml.push_str("</externalLink>");
        if xml.len() > MAX_CACHE_TEXT_BYTES {
            return Err(limit("serialized XML"));
        }
        Ok(xml.into_bytes())
    }
}

#[cfg(test)]
pub(crate) fn build_external_link_part(
    part_uri: PackURI,
    kind: &ExternalLinkKind,
) -> Result<BlobPart> {
    build_external_link_part_with_conformance(part_uri, kind, ExternalLinkConformance::Transitional)
}

pub(crate) fn build_external_link_part_with_conformance(
    part_uri: PackURI,
    kind: &ExternalLinkKind,
    conformance: ExternalLinkConformance,
) -> Result<BlobPart> {
    let xml = kind.to_xml_with_conformance(conformance)?;
    let mut part = BlobPart::new(
        part_uri,
        litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
        xml,
    );
    match kind {
        ExternalLinkKind::Workbook(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES,
            "external workbook",
        )?,
        ExternalLinkKind::Ole(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            &[rt::OLE_OBJECT, rt::STRICT_OLE_OBJECT],
            "OLE",
        )?,
        ExternalLinkKind::Dde(_) => {},
    }
    Ok(part)
}

trait ExternalTargetMetadata {
    fn relationship_id(&self) -> &str;
    fn target(&self) -> &str;
    fn relationship_type(&self) -> &str;
}

impl ExternalTargetMetadata for ExternalWorkbookTarget {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    fn target(&self) -> &str {
        &self.target
    }
    fn relationship_type(&self) -> &str {
        &self.relationship_type
    }
}

impl ExternalTargetMetadata for ExternalOleTarget {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    fn target(&self) -> &str {
        &self.target
    }
    fn relationship_type(&self) -> &str {
        &self.relationship_type
    }
}

fn add_external_target_relationship(
    part: &mut BlobPart,
    target: &impl ExternalTargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    validate_external_target(target, allowed_types, description)?;
    part.rels_mut().add_relationship(
        target.relationship_type().to_string(),
        target.target().to_string(),
        target.relationship_id().to_string(),
        true,
    );
    Ok(())
}

fn validate_external_target(
    target: &impl ExternalTargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    if target.relationship_id().is_empty() {
        return Err(invalid(format!(
            "{description} relationship ID must not be empty"
        )));
    }
    if target.target().is_empty() {
        return Err(invalid(format!("{description} target must not be empty")));
    }
    if target.target().len() > MAX_EXTERNAL_TARGET_BYTES {
        return Err(limit(&format!("{description} target URI")));
    }
    if target.target().chars().any(|character| {
        character.is_control() || character == '\u{fffe}' || character == '\u{ffff}'
    }) {
        return Err(invalid(format!(
            "{description} target URI contains an invalid character"
        )));
    }
    if target.relationship_id().len() > 1024
        || target.relationship_id().chars().any(char::is_control)
    {
        return Err(invalid(format!("{description} relationship ID is invalid")));
    }
    if !allowed_types.contains(&target.relationship_type()) {
        return Err(invalid(format!(
            "{description} has invalid relationship type '{}'",
            target.relationship_type()
        )));
    }
    Ok(())
}

fn write_external_workbook(xml: &mut String, link: &ExternalWorkbookLink) -> Result<()> {
    validate_external_target(
        &link.target,
        EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES,
        "external workbook",
    )?;
    xml.push_str("<externalBook");
    push_xml_attr(xml, "r:id", &link.target.relationship_id)?;
    xml.push('>');
    if link.sheet_names.len() > MAX_SHEET_NAMES {
        return Err(limit("sheet names"));
    }
    if !link.sheet_names.is_empty() {
        xml.push_str("<sheetNames>");
        for name in &link.sheet_names {
            xml.push_str("<sheetName");
            push_xml_attr(xml, "val", name)?;
            xml.push_str("/>");
        }
        xml.push_str("</sheetNames>");
    }
    if link.defined_names.len() > MAX_DEFINED_NAMES {
        return Err(limit("defined names"));
    }
    if !link.defined_names.is_empty() {
        xml.push_str("<definedNames>");
        for name in &link.defined_names {
            xml.push_str("<definedName");
            push_xml_attr(xml, "name", &name.name)?;
            if let Some(value) = &name.refers_to {
                push_xml_attr(xml, "refersTo", value)?;
            }
            if let Some(sheet_id) = name.sheet_id {
                push_u32_attr(xml, "sheetId", sheet_id);
            }
            xml.push_str("/>");
        }
        xml.push_str("</definedNames>");
    }
    if link.cached_sheets.len() > MAX_CACHED_SHEETS {
        return Err(limit("cached sheets"));
    }
    if !link.cached_sheets.is_empty() {
        let mut sheet_ids = std::collections::HashSet::with_capacity(link.cached_sheets.len());
        let mut row_count = 0usize;
        let mut cell_count = 0usize;
        xml.push_str("<sheetDataSet>");
        for sheet in &link.cached_sheets {
            if !sheet_ids.insert(sheet.sheet_id) {
                return Err(invalid(format!(
                    "duplicate external cached sheetId {}",
                    sheet.sheet_id
                )));
            }
            xml.push_str("<sheetData");
            push_u32_attr(xml, "sheetId", sheet.sheet_id);
            if sheet.refresh_error {
                xml.push_str(" refreshError=\"1\"");
            }
            xml.push('>');
            let mut rows = std::collections::HashSet::with_capacity(sheet.rows.len());
            for row in &sheet.rows {
                if row.row == 0 || row.row > 1_048_576 {
                    return Err(invalid("external cached row is outside worksheet bounds"));
                }
                if !rows.insert(row.row) {
                    return Err(invalid(format!(
                        "duplicate external cached row {}",
                        row.row
                    )));
                }
                row_count = row_count
                    .checked_add(1)
                    .ok_or_else(|| limit("cached rows"))?;
                if row_count > MAX_CACHED_ROWS {
                    return Err(limit("cached rows"));
                }
                xml.push_str("<row");
                push_u32_attr(xml, "r", row.row);
                xml.push('>');
                let mut references = std::collections::HashSet::with_capacity(row.cells.len());
                for cell in &row.cells {
                    cell_count = cell_count
                        .checked_add(1)
                        .ok_or_else(|| limit("cached cells"))?;
                    if cell_count > MAX_CACHED_CELLS {
                        return Err(limit("cached cells"));
                    }
                    if let Some(reference) = &cell.reference {
                        validate_cell_reference(reference)?;
                        if !references.insert(reference.as_str()) {
                            return Err(invalid(format!(
                                "duplicate external cached cell '{reference}'"
                            )));
                        }
                    }
                    xml.push_str("<cell");
                    if let Some(reference) = &cell.reference {
                        push_xml_attr(xml, "r", reference)?;
                    }
                    if cell.cell_type != ExternalCellType::Number {
                        push_xml_attr(xml, "t", external_cell_type_token(cell.cell_type))?;
                    }
                    if cell.value_metadata_index != 0 {
                        push_u32_attr(xml, "vm", cell.value_metadata_index);
                    }
                    if let Some(value) = &cell.raw_value {
                        xml.push_str("><v>");
                        push_xml_text(xml, value)?;
                        xml.push_str("</v></cell>");
                    } else {
                        xml.push_str("/>");
                    }
                }
                xml.push_str("</row>");
            }
            xml.push_str("</sheetData>");
        }
        xml.push_str("</sheetDataSet>");
    }
    xml.push_str("</externalBook>");
    Ok(())
}

fn write_dde_link(xml: &mut String, link: &ExternalDdeLink) -> Result<()> {
    xml.push_str("<ddeLink");
    push_xml_attr(xml, "ddeService", &link.service)?;
    push_xml_attr(xml, "ddeTopic", &link.topic)?;
    if link.items.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    if link.items.len() > MAX_LINK_ITEMS {
        return Err(limit("DDE items"));
    }
    xml.push_str("><ddeItems>");
    for item in &link.items {
        xml.push_str("<ddeItem");
        if let Some(name) = &item.name {
            push_xml_attr(xml, "name", name)?;
        }
        push_true_attr(xml, "ole", item.use_ole);
        push_true_attr(xml, "advise", item.advise);
        push_true_attr(xml, "preferPic", item.prefer_picture);
        if let Some(values) = &item.values {
            xml.push('>');
            write_dde_values(xml, "values", values)?;
            xml.push_str("</ddeItem>");
        } else {
            xml.push_str("/>");
        }
    }
    xml.push_str("</ddeItems></ddeLink>");
    Ok(())
}

fn write_ole_link(xml: &mut String, link: &ExternalOleLink) -> Result<()> {
    validate_external_target(
        &link.target,
        &[rt::OLE_OBJECT, rt::STRICT_OLE_OBJECT],
        "OLE",
    )?;
    if link.program_id.is_empty() {
        return Err(invalid("OLE program ID must not be empty"));
    }
    xml.push_str("<oleLink");
    push_xml_attr(xml, "r:id", &link.target.relationship_id)?;
    push_xml_attr(xml, "progId", &link.program_id)?;
    if link.items.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    if link.items.len() > MAX_LINK_ITEMS {
        return Err(limit("OLE items"));
    }
    xml.push_str("><oleItems>");
    for item in &link.items {
        if item.name.is_empty() {
            return Err(invalid("OLE item name must not be empty"));
        }
        let element = if item.source == ExternalOleItemSource::Office2010 {
            "x14:oleItem"
        } else {
            "oleItem"
        };
        xml.push('<');
        xml.push_str(element);
        push_xml_attr(xml, "name", &item.name)?;
        push_true_attr(xml, "icon", item.icon);
        push_true_attr(xml, "advise", item.advise);
        push_true_attr(xml, "preferPic", item.prefer_picture);
        match (&item.values, item.source) {
            (Some(values), ExternalOleItemSource::Office2010) => {
                xml.push('>');
                write_dde_values(xml, "x14:values", values)?;
                xml.push_str("</x14:oleItem>");
            },
            (Some(_), ExternalOleItemSource::SpreadsheetMl) => {
                return Err(invalid("cached OLE values require an Office 2010 oleItem"));
            },
            (None, _) => xml.push_str("/>"),
        }
    }
    xml.push_str("</oleItems></oleLink>");
    Ok(())
}

fn write_dde_values(xml: &mut String, element: &str, values: &ExternalDdeValues) -> Result<()> {
    let expected = u64::from(values.rows)
        .checked_mul(u64::from(values.columns))
        .ok_or_else(|| limit("DDE/OLE matrix dimensions"))?;
    if expected == 0 || expected > MAX_CACHED_CELLS as u64 {
        return Err(limit("DDE/OLE matrix dimensions"));
    }
    if expected != values.values.len() as u64 {
        return Err(invalid(format!(
            "DDE/OLE matrix declares {expected} values but contains {}",
            values.values.len()
        )));
    }
    xml.push('<');
    xml.push_str(element);
    push_u32_attr(xml, "rows", values.rows);
    push_u32_attr(xml, "cols", values.columns);
    xml.push('>');
    for value in &values.values {
        xml.push_str("<value");
        if value.value_type != ExternalDdeValueType::Nil {
            push_xml_attr(xml, "t", dde_value_type_token(value.value_type))?;
        }
        xml.push_str("><val>");
        push_xml_text(xml, &value.raw_value)?;
        xml.push_str("</val></value>");
    }
    xml.push_str("</");
    xml.push_str(element);
    xml.push('>');
    Ok(())
}

fn external_cell_type_token(value: ExternalCellType) -> &'static str {
    match value {
        ExternalCellType::Number => "n",
        ExternalCellType::Boolean => "b",
        ExternalCellType::Date => "d",
        ExternalCellType::Error => "e",
        ExternalCellType::InlineString => "inlineStr",
        ExternalCellType::SharedString => "s",
        ExternalCellType::String => "str",
    }
}

fn dde_value_type_token(value: ExternalDdeValueType) -> &'static str {
    match value {
        ExternalDdeValueType::Nil => "nil",
        ExternalDdeValueType::Boolean => "b",
        ExternalDdeValueType::Number => "n",
        ExternalDdeValueType::Error => "e",
        ExternalDdeValueType::String => "str",
    }
}

fn push_true_attr(xml: &mut String, name: &str, value: bool) {
    if value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"1\"");
    }
}

fn push_u32_attr(xml: &mut String, name: &str, value: u32) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&value.to_string());
    xml.push('"');
}

fn push_xml_attr(xml: &mut String, name: &str, value: &str) -> Result<()> {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    push_xml_text(xml, value)?;
    xml.push('"');
    Ok(())
}

fn push_xml_text(xml: &mut String, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            '\t' | '\n' | '\r' => xml.push(character),
            value if value >= '\u{20}' && value != '\u{fffe}' && value != '\u{ffff}' => {
                xml.push(value)
            },
            value => {
                return Err(invalid(format!(
                    "invalid XML character U+{:04X} in external link",
                    value as u32
                )));
            },
        }
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    ExternalBook,
    DdeLink,
    DdeItems,
    DdeItem(usize),
    DdeValues(usize),
    DdeValue(usize, usize),
    DdeValueText(usize, usize),
    OleLink,
    OleItems,
    OleItem(usize),
    OleValues(usize),
    OleValue(usize, usize),
    OleValueText(usize, usize),
    SheetNames,
    DefinedNames,
    SheetDataSet,
    SheetData(usize),
    Row(usize, usize),
    Cell(usize, usize, usize),
    Value(usize, usize, usize),
    Other,
}

enum ParsedKind {
    Workbook(ParsedExternalBook),
    Dde(ParsedDdeLink),
    Ole(ParsedOleLink),
}

struct ParsedDdeLink {
    service: String,
    topic: String,
    items: Vec<ParsedDdeItem>,
    saw_items: bool,
}

struct ParsedOleLink {
    target_relationship_id: String,
    program_id: String,
    items: Vec<ParsedOleItem>,
    saw_items: bool,
}

struct ParsedDdeItem {
    name: Option<String>,
    use_ole: bool,
    advise: bool,
    prefer_picture: bool,
    values: Option<ParsedDdeValues>,
}

struct ParsedOleItem {
    source: ExternalOleItemSource,
    name: String,
    icon: bool,
    advise: bool,
    prefer_picture: bool,
    values: Option<ParsedDdeValues>,
}

struct ParsedDdeValues {
    rows: u32,
    columns: u32,
    values: Vec<ParsedDdeValue>,
}

struct ParsedDdeValue {
    value_type: ExternalDdeValueType,
    raw_value: Option<String>,
}

struct ParsedExternalBook {
    target_relationship_id: String,
    sheet_names: Vec<String>,
    defined_names: Vec<ExternalDefinedName>,
    cached_sheets: Vec<ExternalSheetData>,
    saw_sheet_names: bool,
    saw_defined_names: bool,
    saw_sheet_data_set: bool,
}

struct Parser {
    kind: Option<ParsedKind>,
    cached_rows: usize,
    cached_cells: usize,
    text_bytes: usize,
}

impl Parser {
    fn new() -> Self {
        Self {
            kind: None,
            cached_rows: 0,
            cached_cells: 0,
            text_bytes: 0,
        }
    }

    fn book_mut(&mut self) -> Result<&mut ParsedExternalBook> {
        match self.kind.as_mut() {
            Some(ParsedKind::Workbook(book)) => Ok(book),
            _ => Err(invalid("external-link content is outside externalBook")),
        }
    }

    fn dde_mut(&mut self) -> Result<&mut ParsedDdeLink> {
        match self.kind.as_mut() {
            Some(ParsedKind::Dde(link)) => Ok(link),
            _ => Err(invalid("DDE content is outside ddeLink")),
        }
    }

    fn ole_mut(&mut self) -> Result<&mut ParsedOleLink> {
        match self.kind.as_mut() {
            Some(ParsedKind::Ole(link)) => Ok(link),
            _ => Err(invalid("OLE content is outside oleLink")),
        }
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        if parent == Context::Root
            && is_spreadsheetml_name(namespace, element.name(), b"externalBook")
        {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| invalid("externalBook is missing relationship ID"))?;
            self.kind = Some(ParsedKind::Workbook(ParsedExternalBook {
                target_relationship_id: relationship_id,
                sheet_names: Vec::new(),
                defined_names: Vec::new(),
                cached_sheets: Vec::new(),
                saw_sheet_names: false,
                saw_defined_names: false,
                saw_sheet_data_set: false,
            }));
            return Ok(Context::ExternalBook);
        }
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"ddeLink") {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let service = required_attr(element, b"ddeService", decoder, "DDE service")?;
            let topic = required_attr(element, b"ddeTopic", decoder, "DDE topic")?;
            self.add_text(service.len() + topic.len())?;
            self.kind = Some(ParsedKind::Dde(ParsedDdeLink {
                service,
                topic,
                items: Vec::new(),
                saw_items: false,
            }));
            return Ok(Context::DdeLink);
        }
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"oleLink") {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| invalid("oleLink is missing relationship ID"))?;
            let program_id = required_attr(element, b"progId", decoder, "OLE program ID")?;
            if program_id.is_empty() {
                return Err(invalid("OLE program ID must not be empty"));
            }
            self.add_text(relationship_id.len() + program_id.len())?;
            self.kind = Some(ParsedKind::Ole(ParsedOleLink {
                target_relationship_id: relationship_id,
                program_id,
                items: Vec::new(),
                saw_items: false,
            }));
            return Ok(Context::OleLink);
        }
        if parent == Context::DdeLink
            && is_spreadsheetml_name(namespace, element.name(), b"ddeItems")
        {
            let link = self.dde_mut()?;
            mark_once(&mut link.saw_items, "DDE items")?;
            return Ok(Context::DdeItems);
        }
        if parent == Context::DdeItems
            && is_spreadsheetml_name(namespace, element.name(), b"ddeItem")
        {
            let name = unqualified_attribute_value(element, b"name", decoder)?;
            let use_ole = optional_bool(element, b"ole", decoder, "DDE item ole")?.unwrap_or(false);
            let advise =
                optional_bool(element, b"advise", decoder, "DDE item advise")?.unwrap_or(false);
            let prefer_picture =
                optional_bool(element, b"preferPic", decoder, "DDE item preferPic")?
                    .unwrap_or(false);
            self.add_text(name.as_ref().map_or(0, String::len))?;
            let link = self.dde_mut()?;
            if link.items.len() >= MAX_LINK_ITEMS {
                return Err(limit("DDE items"));
            }
            let index = link.items.len();
            link.items.push(ParsedDdeItem {
                name,
                use_ole,
                advise,
                prefer_picture,
                values: None,
            });
            return Ok(Context::DdeItem(index));
        }
        if let Context::DdeItem(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"values")
        {
            let rows = optional_u32(element, b"rows", decoder, "DDE value rows")?.unwrap_or(1);
            let columns =
                optional_u32(element, b"cols", decoder, "DDE value columns")?.unwrap_or(1);
            let target = &mut self.dde_mut()?.items[item].values;
            if target
                .replace(ParsedDdeValues {
                    rows,
                    columns,
                    values: Vec::new(),
                })
                .is_some()
            {
                return Err(invalid("duplicate DDE item values"));
            }
            return Ok(Context::DdeValues(item));
        }
        if let Context::DdeValues(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"value")
        {
            let value_type = parse_dde_value_type(
                unqualified_attribute_value(element, b"t", decoder)?.as_deref(),
            )?;
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("DDE values"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("DDE values"));
            }
            let values = &mut self.dde_mut()?.items[item]
                .values
                .as_mut()
                .expect("values context")
                .values;
            let index = values.len();
            values.push(ParsedDdeValue {
                value_type,
                raw_value: None,
            });
            return Ok(Context::DdeValue(item, index));
        }
        if let Context::DdeValue(item, value) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"val")
        {
            let raw = &mut self.dde_mut()?.items[item]
                .values
                .as_mut()
                .expect("values context")
                .values[value]
                .raw_value;
            if raw.replace(String::new()).is_some() {
                return Err(invalid("duplicate DDE val element"));
            }
            return Ok(Context::DdeValueText(item, value));
        }
        if parent == Context::OleLink
            && is_spreadsheetml_name(namespace, element.name(), b"oleItems")
        {
            let link = self.ole_mut()?;
            mark_once(&mut link.saw_items, "OLE items")?;
            return Ok(Context::OleItems);
        }
        if parent == Context::OleItems
            && (is_spreadsheetml_name(namespace, element.name(), b"oleItem")
                || is_exact_name(namespace, element.name(), X14, b"oleItem"))
        {
            let source = if is_exact_name(namespace, element.name(), X14, b"oleItem") {
                ExternalOleItemSource::Office2010
            } else {
                ExternalOleItemSource::SpreadsheetMl
            };
            let name = required_attr(element, b"name", decoder, "OLE item name")?;
            if name.is_empty() {
                return Err(invalid("OLE item name must not be empty"));
            }
            let icon = optional_bool(element, b"icon", decoder, "OLE item icon")?.unwrap_or(false);
            let advise =
                optional_bool(element, b"advise", decoder, "OLE item advise")?.unwrap_or(false);
            let prefer_picture =
                optional_bool(element, b"preferPic", decoder, "OLE item preferPic")?
                    .unwrap_or(false);
            self.add_text(name.len())?;
            let link = self.ole_mut()?;
            if link.items.len() >= MAX_LINK_ITEMS {
                return Err(limit("OLE items"));
            }
            let index = link.items.len();
            link.items.push(ParsedOleItem {
                source,
                name,
                icon,
                advise,
                prefer_picture,
                values: None,
            });
            return Ok(Context::OleItem(index));
        }
        if let Context::OleItem(item) = parent
            && is_exact_name(namespace, element.name(), X14, b"values")
        {
            if self.ole_mut()?.items[item].source != ExternalOleItemSource::Office2010 {
                return Err(invalid("cached OLE values require an Office 2010 oleItem"));
            }
            let rows = optional_u32(element, b"rows", decoder, "OLE value rows")?.unwrap_or(1);
            let columns =
                optional_u32(element, b"cols", decoder, "OLE value columns")?.unwrap_or(1);
            let target = &mut self.ole_mut()?.items[item].values;
            if target
                .replace(ParsedDdeValues {
                    rows,
                    columns,
                    values: Vec::new(),
                })
                .is_some()
            {
                return Err(invalid("duplicate OLE item values"));
            }
            return Ok(Context::OleValues(item));
        }
        if let Context::OleValues(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"value")
        {
            let value_type = parse_dde_value_type(
                unqualified_attribute_value(element, b"t", decoder)?.as_deref(),
            )?;
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("OLE values"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("OLE values"));
            }
            let values = &mut self.ole_mut()?.items[item]
                .values
                .as_mut()
                .expect("values context")
                .values;
            let index = values.len();
            values.push(ParsedDdeValue {
                value_type,
                raw_value: None,
            });
            return Ok(Context::OleValue(item, index));
        }
        if let Context::OleValue(item, value) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"val")
        {
            let raw = &mut self.ole_mut()?.items[item]
                .values
                .as_mut()
                .expect("values context")
                .values[value]
                .raw_value;
            if raw.replace(String::new()).is_some() {
                return Err(invalid("duplicate OLE val element"));
            }
            return Ok(Context::OleValueText(item, value));
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"sheetNames")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_names, "external sheetNames")?;
            return Ok(Context::SheetNames);
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_defined_names, "external definedNames")?;
            return Ok(Context::DefinedNames);
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"sheetDataSet")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_data_set, "external sheetDataSet")?;
            return Ok(Context::SheetDataSet);
        }
        if parent == Context::SheetNames
            && is_spreadsheetml_name(namespace, element.name(), b"sheetName")
        {
            let value = unqualified_attribute_value(element, b"val", decoder)?.unwrap_or_default();
            self.add_text(value.len())?;
            let book = self.book_mut()?;
            if book.sheet_names.len() >= MAX_SHEET_NAMES {
                return Err(limit("sheet names"));
            }
            book.sheet_names.push(value);
            return Ok(Context::Other);
        }
        if parent == Context::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            let name = required_attr(element, b"name", decoder, "external defined name")?;
            let refers_to = unqualified_attribute_value(element, b"refersTo", decoder)?;
            let sheet_id = optional_u32(
                element,
                b"sheetId",
                decoder,
                "external defined-name sheetId",
            )?;
            self.add_text(name.len() + refers_to.as_ref().map_or(0, String::len))?;
            let book = self.book_mut()?;
            if book.defined_names.len() >= MAX_DEFINED_NAMES {
                return Err(limit("defined names"));
            }
            book.defined_names.push(ExternalDefinedName {
                name,
                refers_to,
                sheet_id,
            });
            return Ok(Context::Other);
        }
        if parent == Context::SheetDataSet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            let sheet_id = required_u32(element, b"sheetId", decoder, "external cached sheetId")?;
            let refresh_error =
                optional_bool(element, b"refreshError", decoder, "external refreshError")?
                    .unwrap_or(false);
            let book = self.book_mut()?;
            if book.cached_sheets.len() >= MAX_CACHED_SHEETS {
                return Err(limit("cached sheets"));
            }
            if book
                .cached_sheets
                .iter()
                .any(|sheet| sheet.sheet_id == sheet_id)
            {
                return Err(invalid(format!(
                    "duplicate external cached sheetId {sheet_id}"
                )));
            }
            let index = book.cached_sheets.len();
            book.cached_sheets.push(ExternalSheetData {
                sheet_id,
                refresh_error,
                rows: Vec::new(),
            });
            return Ok(Context::SheetData(index));
        }
        if let Context::SheetData(sheet) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            let row = required_u32(element, b"r", decoder, "external cached row")?;
            if row == 0 {
                return Err(invalid("external cached row must be positive"));
            }
            self.cached_rows = self
                .cached_rows
                .checked_add(1)
                .ok_or_else(|| limit("cached rows"))?;
            if self.cached_rows > MAX_CACHED_ROWS {
                return Err(limit("cached rows"));
            }
            let book = self.book_mut()?;
            let rows = &mut book.cached_sheets[sheet].rows;
            if rows.iter().any(|item| item.row == row) {
                return Err(invalid(format!("duplicate external cached row {row}")));
            }
            let index = rows.len();
            rows.push(ExternalRow {
                row,
                cells: Vec::new(),
            });
            return Ok(Context::Row(sheet, index));
        }
        if let Context::Row(sheet, row) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"cell")
        {
            let reference = unqualified_attribute_value(element, b"r", decoder)?;
            if let Some(value) = reference.as_deref() {
                validate_cell_reference(value)?;
            }
            let cell_type =
                parse_cell_type(unqualified_attribute_value(element, b"t", decoder)?.as_deref())?;
            let value_metadata_index =
                optional_u32(element, b"vm", decoder, "external value metadata index")?
                    .unwrap_or(0);
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("cached cells"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("cached cells"));
            }
            let book = self.book_mut()?;
            let cells = &mut book.cached_sheets[sheet].rows[row].cells;
            if let Some(reference) = reference.as_deref()
                && cells
                    .iter()
                    .any(|cell| cell.reference.as_deref() == Some(reference))
            {
                return Err(invalid(format!(
                    "duplicate external cached cell '{reference}'"
                )));
            }
            let index = cells.len();
            cells.push(ExternalCell {
                reference,
                cell_type,
                raw_value: None,
                value_metadata_index,
            });
            return Ok(Context::Cell(sheet, row, index));
        }
        if let Context::Cell(sheet, row, cell) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"v")
        {
            return Ok(Context::Value(sheet, row, cell));
        }
        if matches!(
            parent,
            Context::DdeLink
                | Context::DdeItems
                | Context::DdeItem(_)
                | Context::DdeValues(_)
                | Context::DdeValue(_, _)
                | Context::OleLink
                | Context::OleItems
                | Context::OleItem(_)
                | Context::OleValues(_)
                | Context::OleValue(_, _)
        ) {
            return Err(invalid("unexpected child in DDE/OLE external link"));
        }
        Ok(Context::Other)
    }

    fn push_value(&mut self, sheet: usize, row: usize, cell: usize, value: &str) -> Result<()> {
        self.add_text(value.len())?;
        let target = &mut self.book_mut()?.cached_sheets[sheet].rows[row].cells[cell].raw_value;
        target.get_or_insert_with(String::new).push_str(value);
        Ok(())
    }

    fn push_dde_value(&mut self, item: usize, value: usize, text: &str) -> Result<()> {
        self.add_text(text.len())?;
        self.dde_mut()?.items[item]
            .values
            .as_mut()
            .expect("values context")
            .values[value]
            .raw_value
            .as_mut()
            .expect("val context")
            .push_str(text);
        Ok(())
    }

    fn push_ole_value(&mut self, item: usize, value: usize, text: &str) -> Result<()> {
        self.add_text(text.len())?;
        self.ole_mut()?.items[item]
            .values
            .as_mut()
            .expect("values context")
            .values[value]
            .raw_value
            .as_mut()
            .expect("val context")
            .push_str(text);
        Ok(())
    }

    fn add_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("cache text"))?;
        if self.text_bytes > MAX_CACHE_TEXT_BYTES {
            return Err(limit("cache text"));
        }
        Ok(())
    }
}

pub(crate) fn load_external_link(
    part: &dyn Part,
    workbook_relationship_id: String,
    index: u32,
) -> Result<ExternalLinkEntry> {
    let parsed = parse_external_link(part.blob())?;
    let kind = match parsed {
        ParsedKind::Workbook(book) => {
            let relationship = part
                .rels()
                .get(&book.target_relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "externalBook references missing relationship '{}'",
                        book.target_relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("externalBook target relationship must be external"));
            }
            if !is_external_workbook_relationship(relationship.reltype()) {
                return Err(invalid(format!(
                    "externalBook target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            ExternalLinkKind::Workbook(ExternalWorkbookLink {
                target: ExternalWorkbookTarget {
                    relationship_id: book.target_relationship_id,
                    target: relationship.target_ref().to_string(),
                    relationship_type: relationship.reltype().to_string(),
                },
                sheet_names: book.sheet_names,
                defined_names: book.defined_names,
                cached_sheets: book.cached_sheets,
            })
        },
        ParsedKind::Dde(link) => {
            if link.saw_items && link.items.is_empty() {
                return Err(invalid("ddeItems must contain at least one ddeItem"));
            }
            ExternalLinkKind::Dde(ExternalDdeLink {
                service: link.service,
                topic: link.topic,
                items: link
                    .items
                    .into_iter()
                    .map(finalize_dde_item)
                    .collect::<Result<_>>()?,
            })
        },
        ParsedKind::Ole(link) => {
            if link.saw_items && link.items.is_empty() {
                return Err(invalid("oleItems must contain at least one oleItem"));
            }
            let relationship = part
                .rels()
                .get(&link.target_relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "oleLink references missing relationship '{}'",
                        link.target_relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("oleLink target relationship must be external"));
            }
            if !matches!(
                relationship.reltype(),
                rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT
            ) {
                return Err(invalid(format!(
                    "oleLink target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            ExternalLinkKind::Ole(ExternalOleLink {
                target: ExternalOleTarget {
                    relationship_id: link.target_relationship_id,
                    target: relationship.target_ref().to_string(),
                    relationship_type: relationship.reltype().to_string(),
                },
                program_id: link.program_id,
                items: link
                    .items
                    .into_iter()
                    .map(finalize_ole_item)
                    .collect::<Result<_>>()?,
            })
        },
    };
    Ok(ExternalLinkEntry {
        index,
        relationship_id: workbook_relationship_id,
        part_uri: part.partname().clone(),
        kind,
    })
}

fn parse_external_link(xml: &[u8]) -> Result<ParsedKind> {
    let xml = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut parser = Parser::new();
    let mut stack = Vec::new();
    let mut closed_root = false;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                if closed_root
                    || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink")
                {
                    return Err(invalid(
                        "external-link XML must have one SpreadsheetML externalLink root",
                    ));
                }
                stack.push(Context::Root);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root
                    || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink")
                {
                    return Err(invalid(
                        "external-link XML must have one SpreadsheetML externalLink root",
                    ));
                }
                return Err(invalid("externalLink must contain a link kind"));
            },
            Event::Start(element) => {
                let parent = *stack
                    .last()
                    .ok_or_else(|| invalid("external-link XML is missing its root"))?;
                let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = *stack
                    .last()
                    .ok_or_else(|| invalid("external-link XML is missing its root"))?;
                parser.start(parent, &namespace, &element, decoder, &resolver)?;
            },
            Event::Text(text) => {
                let text = text.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::CData(text) => {
                let text = text.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::GeneralRef(reference) => {
                let text = decode_xml_reference(&reference)?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::End(element) => {
                let context = stack
                    .pop()
                    .ok_or_else(|| invalid("external-link closing element outside root"))?;
                if context == Context::Root {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"externalLink") {
                        return Err(invalid(
                            "external-link XML has invalid root closing element",
                        ));
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err(invalid("external-link XML has an unterminated root"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    parser
        .kind
        .ok_or_else(|| invalid("externalLink must contain a link kind"))
}

fn push_context_text(parser: &mut Parser, context: Option<Context>, text: &str) -> Result<()> {
    match context {
        Some(Context::Value(sheet, row, cell)) => parser.push_value(sheet, row, cell, text),
        Some(Context::DdeValueText(item, value)) => parser.push_dde_value(item, value, text),
        Some(Context::OleValueText(item, value)) => parser.push_ole_value(item, value, text),
        Some(
            Context::DdeLink
            | Context::DdeItems
            | Context::DdeItem(_)
            | Context::DdeValues(_)
            | Context::DdeValue(_, _)
            | Context::OleLink
            | Context::OleItems
            | Context::OleItem(_)
            | Context::OleValues(_)
            | Context::OleValue(_, _),
        ) if !text.trim().is_empty() => Err(invalid("unexpected text in DDE/OLE external link")),
        _ => Ok(()),
    }
}

fn finalize_dde_item(item: ParsedDdeItem) -> Result<ExternalDdeItem> {
    Ok(ExternalDdeItem {
        name: item.name,
        use_ole: item.use_ole,
        advise: item.advise,
        prefer_picture: item.prefer_picture,
        values: item.values.map(finalize_dde_values).transpose()?,
    })
}

fn finalize_ole_item(item: ParsedOleItem) -> Result<ExternalOleItem> {
    Ok(ExternalOleItem {
        source: item.source,
        name: item.name,
        icon: item.icon,
        advise: item.advise,
        prefer_picture: item.prefer_picture,
        values: item.values.map(finalize_dde_values).transpose()?,
    })
}

fn finalize_dde_values(values: ParsedDdeValues) -> Result<ExternalDdeValues> {
    let ParsedDdeValues {
        rows,
        columns,
        values,
    } = values;
    let expected = u64::from(rows)
        .checked_mul(u64::from(columns))
        .ok_or_else(|| limit("DDE/OLE matrix dimensions"))?;
    if expected > MAX_CACHED_CELLS as u64 {
        return Err(limit("DDE/OLE matrix dimensions"));
    }
    if expected != values.len() as u64 {
        return Err(invalid(format!(
            "DDE/OLE matrix declares {expected} values but contains {}",
            values.len()
        )));
    }
    let values = values
        .into_iter()
        .map(|value| {
            Ok(ExternalDdeValue {
                value_type: value.value_type,
                raw_value: value
                    .raw_value
                    .ok_or_else(|| invalid("DDE/OLE value is missing val child"))?,
            })
        })
        .collect::<Result<_>>()?;
    Ok(ExternalDdeValues {
        rows,
        columns,
        values,
    })
}

fn parse_dde_value_type(value: Option<&str>) -> Result<ExternalDdeValueType> {
    match value.unwrap_or("nil") {
        "nil" => Ok(ExternalDdeValueType::Nil),
        "b" => Ok(ExternalDdeValueType::Boolean),
        "n" => Ok(ExternalDdeValueType::Number),
        "e" => Ok(ExternalDdeValueType::Error),
        "str" => Ok(ExternalDdeValueType::String),
        value => Err(invalid(format!("invalid DDE value type '{value}'"))),
    }
}

fn is_exact_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    uri: &[u8],
    local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == uri)
        && name.local_name().as_ref() == local
}

fn parse_cell_type(value: Option<&str>) -> Result<ExternalCellType> {
    match value.unwrap_or("n") {
        "n" => Ok(ExternalCellType::Number),
        "b" => Ok(ExternalCellType::Boolean),
        "d" => Ok(ExternalCellType::Date),
        "e" => Ok(ExternalCellType::Error),
        "inlineStr" => Ok(ExternalCellType::InlineString),
        "s" => Ok(ExternalCellType::SharedString),
        "str" => Ok(ExternalCellType::String),
        value => Err(invalid(format!(
            "invalid external cached cell type '{value}'"
        ))),
    }
}

/// Validate an `A1`-style cached cell reference.
///
/// Accepts the full SpreadsheetML address space, so multi-letter columns
/// (`AA1` through `XFD1048576`) are valid. The column prefix is bounded to
/// [`MAX_COLUMN_LETTERS`] so a long letter run cannot drive the accumulator.
fn validate_cell_reference(value: &str) -> Result<()> {
    let malformed = || invalid(format!("invalid external cached cell reference '{value}'"));

    let mut column = 0u32;
    let mut split = 0usize;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() {
            break;
        }
        if split == MAX_COLUMN_LETTERS {
            return Err(malformed());
        }
        column = column
            .checked_mul(26)
            .and_then(|n| n.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid("external cell column overflow"))?;
        split += 1;
    }
    if split == 0 || split == value.len() || column > MAX_CELL_COLUMN {
        return Err(malformed());
    }
    let row = value[split..].parse::<u32>().map_err(|_| malformed())?;
    if row == 0 || row > MAX_CELL_ROW {
        return Err(malformed());
    }
    Ok(())
}

fn required_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}
fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" | "on" => Ok(true),
            "0" | "false" | "off" => Ok(false),
            _ => Err(invalid(format!("invalid {description} '{value}'"))),
        })
        .transpose()
}
fn mark_once(seen: &mut bool, description: &str) -> Result<()> {
    if std::mem::replace(seen, true) {
        Err(invalid(format!("duplicate {description}")))
    } else {
        Ok(())
    }
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("external-link {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    #[test]
    fn parses_sparse_lexical_cache_without_dereferencing_target() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Data"/></sheetNames><definedNames><definedName name="Rate" refersTo="Data!$A$1" sheetId="1"/></definedNames><sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1" t="str"><v>001.2300</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext(
            "https://127.0.0.1:9/never-open.xlsx",
            rt::EXTERNAL_LINK_PATH,
        );
        let link = load_external_link(&part, "bookRel".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else {
            panic!("expected workbook link")
        };
        assert_eq!(book.target.target, "https://127.0.0.1:9/never-open.xlsx");
        assert_eq!(book.sheet_names, ["Data"]);
        assert_eq!(
            book.cached_sheets[0].rows[0].cells[0].raw_value.as_deref(),
            Some("001.2300")
        );
    }

    #[test]
    fn parses_typed_dde_and_ole_links_without_dereferencing_targets() {
        let dde = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="excel" ddeTopic="[source.xlsx]Sheet1"><ddeItems><ddeItem name="R1C1:R1C2" advise="1"><values cols="2"><value t="n"><val>001.20</val></value><value t="str"><val>A&amp;B</val></value></values></ddeItem></ddeItems></ddeLink></externalLink>"#;
        let part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            dde.to_vec(),
        );
        let ExternalLinkKind::Dde(link) = load_external_link(&part, "rId1".into(), 1).unwrap().kind
        else {
            panic!("expected DDE link")
        };
        assert_eq!(link.service, "excel");
        assert!(link.items[0].advise);
        let values = link.items[0].values.as_ref().unwrap();
        assert_eq!((values.rows, values.columns), (1, 2));
        assert_eq!(values.values[0].raw_value, "001.20");
        assert_eq!(values.values[1].raw_value, "A&B");

        let ole = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><oleLink r:id="rId1" progId="Excel.Sheet.12"><oleItems><oleItem name="Core" icon="1"/><x14:oleItem name="Cached" preferPic="true"><x14:values><value t="b"><val>1</val></value></x14:values></x14:oleItem></oleItems></oleLink></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink2.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            ole.to_vec(),
        );
        part.relate_to_ext("https://127.0.0.1:9/never-open.bin", rt::OLE_OBJECT);
        let ExternalLinkKind::Ole(link) = load_external_link(&part, "rId2".into(), 2).unwrap().kind
        else {
            panic!("expected OLE link")
        };
        assert_eq!(link.target.target, "https://127.0.0.1:9/never-open.bin");
        assert_eq!(link.items[0].source, ExternalOleItemSource::SpreadsheetMl);
        assert_eq!(link.items[1].source, ExternalOleItemSource::Office2010);
        assert_eq!(
            link.items[1].values.as_ref().unwrap().values[0].value_type,
            ExternalDdeValueType::Boolean
        );
    }

    #[test]
    fn rejects_malformed_dde_ole_links_and_matrices() {
        let sml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        for child in [
            "<ddeLink ddeService=\"x\"/>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems/></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values rows=\"2\"><value><val>x</val></value></values></ddeItem></ddeItems></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values><value t=\"future\"><val>x</val></value></values></ddeItem></ddeItems></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values><value/></values></ddeItem></ddeItems></ddeLink>",
            "<oleLink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>",
        ] {
            let xml = format!("<externalLink xmlns=\"{sml}\">{child}</externalLink>");
            let part = BlobPart::new(
                PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
                litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
                xml.into_bytes(),
            );
            assert!(
                load_external_link(&part, "rId1".into(), 1).is_err(),
                "accepted {child}"
            );
        }
    }

    #[test]
    fn canonical_writer_round_trips_typed_dde_and_ole_links() {
        let dde = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "x&y".into(),
            topic: "topic".into(),
            items: vec![ExternalDdeItem {
                name: Some("R1C1".into()),
                use_ole: false,
                advise: true,
                prefer_picture: false,
                values: Some(ExternalDdeValues {
                    rows: 1,
                    columns: 1,
                    values: vec![ExternalDdeValue {
                        value_type: ExternalDdeValueType::String,
                        raw_value: "<&>".into(),
                    }],
                }),
            }],
        });
        let xml = dde.to_xml().unwrap();
        let part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml,
        );
        assert_eq!(
            load_external_link(&part, "rId1".into(), 1).unwrap().kind,
            dde
        );

        let ole = ExternalLinkKind::Ole(ExternalOleLink {
            target: ExternalOleTarget {
                relationship_id: "rId1".into(),
                target: "file.bin".into(),
                relationship_type: rt::OLE_OBJECT.into(),
            },
            program_id: "Excel.Sheet.12".into(),
            items: vec![ExternalOleItem {
                source: ExternalOleItemSource::Office2010,
                name: "Item".into(),
                icon: false,
                advise: false,
                prefer_picture: true,
                values: Some(ExternalDdeValues {
                    rows: 1,
                    columns: 1,
                    values: vec![ExternalDdeValue {
                        value_type: ExternalDdeValueType::Number,
                        raw_value: "1.00".into(),
                    }],
                }),
            }],
        });
        let part = build_external_link_part(
            PackURI::new("/xl/externalLinks/externalLink2.xml").unwrap(),
            &ole,
        )
        .unwrap();
        assert_eq!(
            load_external_link(&part, "rId2".into(), 2).unwrap().kind,
            ole
        );
    }

    #[test]
    fn workbook_add_and_replace_external_link_survive_save() {
        let first = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "one".into(),
            topic: "topic".into(),
            items: Vec::new(),
        });
        let replacement = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "two".into(),
            topic: "updated".into(),
            items: vec![ExternalDdeItem {
                name: Some("R1C1".into()),
                use_ole: false,
                advise: true,
                prefer_picture: false,
                values: None,
            }],
        });
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        assert_eq!(workbook.add_external_link(first).unwrap(), 1);
        workbook
            .replace_external_link(1, replacement.clone())
            .unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("typed-external-link.xlsx");
        workbook.save(&path).unwrap();
        let package = litchi_opc::OpcPackage::from_bytes(&std::fs::read(path).unwrap()).unwrap();
        let reopened = crate::xlsx::Workbook::new(package).unwrap();
        assert_eq!(reopened.external_links().len(), 1);
        assert_eq!(reopened.external_link(1).unwrap().kind, replacement);
    }

    #[test]
    fn mce_fallback_is_semantic_input() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:no"><mc:AlternateContent><mc:Choice Requires="x"><x:no/></mc:Choice><mc:Fallback><externalBook r:id="rId1"/></mc:Fallback></mc:AlternateContent></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext("opaque.xlsx", rt::EXTERNAL_LINK_PATH);
        assert!(matches!(
            load_external_link(&part, "rId1".into(), 1).unwrap().kind,
            ExternalLinkKind::Workbook(_)
        ));
    }

    #[test]
    fn accepts_strict_external_workbook_namespaces_and_relationships() {
        let xml = br#"<externalLink xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Strict"/></sheetNames></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext("strict-target.xlsx", rt::STRICT_EXTERNAL_LINK_PATH);
        let link = load_external_link(&part, "rId1".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else {
            panic!("expected workbook link")
        };
        assert_eq!(book.sheet_names, ["Strict"]);
        assert_eq!(book.target.relationship_type, rt::STRICT_EXTERNAL_LINK_PATH);
    }

    #[test]
    fn rejects_duplicate_and_over_limit_workbook_external_references() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let duplicate = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences><externalReference r:id="link"/><externalReference r:id="link"/></externalReferences></workbook>"#
        );
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&duplicate).is_err());

        let mut oversized = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences>"#
        );
        for index in 0..=4096 {
            oversized.push_str(&format!(r#"<externalReference r:id="link{index}"/>"#));
        }
        oversized.push_str("</externalReferences></workbook>");
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&oversized).is_err());
    }

    #[test]
    fn rejects_internal_targets_and_malformed_caches() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetDataSet><sheetData sheetId="0"/></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to("../workbook.xml", rt::EXTERNAL_LINK_PATH);
        assert!(load_external_link(&part, "rId1".into(), 1).is_err());
    }

    #[test]
    fn loads_poi_ordered_external_workbook_reference() {
        let package = litchi_opc::OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/link-external-workbook-b.xlsx"
        ))
        .unwrap();
        let workbook = crate::xlsx::Workbook::new(package).unwrap();
        assert_eq!(workbook.external_links().len(), 1);
        let link = workbook.external_link(1).unwrap();
        assert_eq!(link.index, 1);
        assert_eq!(link.relationship_id, "rId4");
        let ExternalLinkKind::Workbook(book) = &link.kind else {
            panic!("expected external workbook")
        };
        assert_eq!(book.target.target, "link-external-workbook-a.xlsx");
    }

    #[test]
    fn loads_libreoffice_sparse_external_cache() {
        let package = litchi_opc::OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
        ))
        .unwrap();
        let workbook = crate::xlsx::Workbook::new(package).unwrap();
        let values: Vec<&str> = workbook
            .external_links()
            .iter()
            .filter_map(|link| match &link.kind {
                ExternalLinkKind::Workbook(book) => Some(book),
                _ => None,
            })
            .flat_map(|book| &book.cached_sheets)
            .flat_map(|sheet| &sheet.rows)
            .flat_map(|row| &row.cells)
            .filter_map(|cell| cell.raw_value.as_deref())
            .collect();
        for expected in ["Name", "Andy", "Bruce", "Charlie"] {
            assert!(
                values.contains(&expected),
                "missing cached value {expected}"
            );
        }
    }

    #[test]
    fn unrelated_modified_save_preserves_external_part_and_relationship() {
        const FIXTURE: &[u8] = include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
        );
        let original = litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap();
        let original_workbook =
            crate::xlsx::Workbook::new(litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap())
                .unwrap();
        let original_link = original_workbook.external_links().first().unwrap();
        let external_uri = original_link.part_uri.clone();
        let workbook_relationship_id = original_link.relationship_id.clone();
        let original_external = original.get_part(&external_uri).unwrap().blob().to_vec();
        let original_workbook_part = original.main_document_part().unwrap();
        let original_relationship = original_workbook_part
            .rels()
            .get(&workbook_relationship_id)
            .unwrap();
        let original_target = original_relationship.target_ref().to_string();
        let original_type = original_relationship.reltype().to_string();

        let mut workbook =
            crate::xlsx::Workbook::new(litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap())
                .unwrap();
        workbook.properties_mut().title = Some("Unrelated edit".to_owned());
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("external-link-roundtrip.xlsx");
        workbook.save(&path).unwrap();

        let bytes = std::fs::read(&path).unwrap();
        let saved = litchi_opc::OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            saved.get_part(&external_uri).unwrap().blob(),
            original_external
        );
        let workbook_part = saved.main_document_part().unwrap();
        let relationship = workbook_part.rels().get(&workbook_relationship_id).unwrap();
        assert_eq!(relationship.reltype(), original_type);
        assert_eq!(relationship.target_ref(), original_target);
        let workbook_xml = std::str::from_utf8(workbook_part.blob()).unwrap();
        assert!(workbook_xml.contains("<externalReferences>"));
        assert!(workbook_xml.contains(&format!(
            "<externalReference r:id=\"{workbook_relationship_id}\"/>"
        )));
        assert_eq!(
            crate::xlsx::Workbook::new(saved)
                .unwrap()
                .external_links()
                .len(),
            original_workbook.external_links().len()
        );
    }
}
