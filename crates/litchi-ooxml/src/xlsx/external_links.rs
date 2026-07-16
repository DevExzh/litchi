//! Read-only SpreadsheetML external-workbook link metadata and cached values.

use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_opc::constants::relationship_type as rt;
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
const MAX_CACHE_TEXT_BYTES: usize = 64 * 1024 * 1024;

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
    DdeOpaque,
    OleOpaque,
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    ExternalBook,
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
    Dde,
    Ole,
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
        Self { kind: None, cached_rows: 0, cached_cells: 0, text_bytes: 0 }
    }

    fn book_mut(&mut self) -> Result<&mut ParsedExternalBook> {
        match self.kind.as_mut() {
            Some(ParsedKind::Workbook(book)) => Ok(book),
            _ => Err(invalid("external-link content is outside externalBook")),
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
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"externalBook") {
            if self.kind.is_some() { return Err(invalid("externalLink has multiple link kinds")); }
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
            if self.kind.replace(ParsedKind::Dde).is_some() { return Err(invalid("externalLink has multiple link kinds")); }
            return Ok(Context::Other);
        }
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"oleLink") {
            if self.kind.replace(ParsedKind::Ole).is_some() { return Err(invalid("externalLink has multiple link kinds")); }
            return Ok(Context::Other);
        }
        if parent == Context::ExternalBook && is_spreadsheetml_name(namespace, element.name(), b"sheetNames") {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_names, "external sheetNames")?;
            return Ok(Context::SheetNames);
        }
        if parent == Context::ExternalBook && is_spreadsheetml_name(namespace, element.name(), b"definedNames") {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_defined_names, "external definedNames")?;
            return Ok(Context::DefinedNames);
        }
        if parent == Context::ExternalBook && is_spreadsheetml_name(namespace, element.name(), b"sheetDataSet") {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_data_set, "external sheetDataSet")?;
            return Ok(Context::SheetDataSet);
        }
        if parent == Context::SheetNames && is_spreadsheetml_name(namespace, element.name(), b"sheetName") {
            let value = unqualified_attribute_value(element, b"val", decoder)?.unwrap_or_default();
            self.add_text(value.len())?;
            let book = self.book_mut()?;
            if book.sheet_names.len() >= MAX_SHEET_NAMES { return Err(limit("sheet names")); }
            book.sheet_names.push(value);
            return Ok(Context::Other);
        }
        if parent == Context::DefinedNames && is_spreadsheetml_name(namespace, element.name(), b"definedName") {
            let name = required_attr(element, b"name", decoder, "external defined name")?;
            let refers_to = unqualified_attribute_value(element, b"refersTo", decoder)?;
            let sheet_id = optional_u32(element, b"sheetId", decoder, "external defined-name sheetId")?;
            self.add_text(name.len() + refers_to.as_ref().map_or(0, String::len))?;
            let book = self.book_mut()?;
            if book.defined_names.len() >= MAX_DEFINED_NAMES { return Err(limit("defined names")); }
            book.defined_names.push(ExternalDefinedName { name, refers_to, sheet_id });
            return Ok(Context::Other);
        }
        if parent == Context::SheetDataSet && is_spreadsheetml_name(namespace, element.name(), b"sheetData") {
            let sheet_id = required_u32(element, b"sheetId", decoder, "external cached sheetId")?;
            let refresh_error = optional_bool(element, b"refreshError", decoder, "external refreshError")?.unwrap_or(false);
            let book = self.book_mut()?;
            if book.cached_sheets.len() >= MAX_CACHED_SHEETS { return Err(limit("cached sheets")); }
            if book.cached_sheets.iter().any(|sheet| sheet.sheet_id == sheet_id) {
                return Err(invalid(format!("duplicate external cached sheetId {sheet_id}")));
            }
            let index = book.cached_sheets.len();
            book.cached_sheets.push(ExternalSheetData { sheet_id, refresh_error, rows: Vec::new() });
            return Ok(Context::SheetData(index));
        }
        if let Context::SheetData(sheet) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            let row = required_u32(element, b"r", decoder, "external cached row")?;
            if row == 0 { return Err(invalid("external cached row must be positive")); }
            self.cached_rows = self.cached_rows.checked_add(1).ok_or_else(|| limit("cached rows"))?;
            if self.cached_rows > MAX_CACHED_ROWS { return Err(limit("cached rows")); }
            let book = self.book_mut()?;
            let rows = &mut book.cached_sheets[sheet].rows;
            if rows.iter().any(|item| item.row == row) { return Err(invalid(format!("duplicate external cached row {row}"))); }
            let index = rows.len();
            rows.push(ExternalRow { row, cells: Vec::new() });
            return Ok(Context::Row(sheet, index));
        }
        if let Context::Row(sheet, row) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"cell")
        {
            let reference = unqualified_attribute_value(element, b"r", decoder)?;
            if let Some(value) = reference.as_deref() { validate_cell_reference(value)?; }
            let cell_type = parse_cell_type(unqualified_attribute_value(element, b"t", decoder)?.as_deref())?;
            let value_metadata_index = optional_u32(element, b"vm", decoder, "external value metadata index")?.unwrap_or(0);
            self.cached_cells = self.cached_cells.checked_add(1).ok_or_else(|| limit("cached cells"))?;
            if self.cached_cells > MAX_CACHED_CELLS { return Err(limit("cached cells")); }
            let book = self.book_mut()?;
            let cells = &mut book.cached_sheets[sheet].rows[row].cells;
            if let Some(reference) = reference.as_deref()
                && cells.iter().any(|cell| cell.reference.as_deref() == Some(reference))
            { return Err(invalid(format!("duplicate external cached cell '{reference}'"))); }
            let index = cells.len();
            cells.push(ExternalCell { reference, cell_type, raw_value: None, value_metadata_index });
            return Ok(Context::Cell(sheet, row, index));
        }
        if let Context::Cell(sheet, row, cell) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"v")
        {
            return Ok(Context::Value(sheet, row, cell));
        }
        Ok(Context::Other)
    }

    fn push_value(&mut self, sheet: usize, row: usize, cell: usize, value: &str) -> Result<()> {
        self.add_text(value.len())?;
        let target = &mut self.book_mut()?.cached_sheets[sheet].rows[row].cells[cell].raw_value;
        target.get_or_insert_with(String::new).push_str(value);
        Ok(())
    }

    fn add_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| limit("cache text"))?;
        if self.text_bytes > MAX_CACHE_TEXT_BYTES { return Err(limit("cache text")); }
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
            let relationship = part.rels().get(&book.target_relationship_id).ok_or_else(|| {
                invalid(format!("externalBook references missing relationship '{}'", book.target_relationship_id))
            })?;
            if !relationship.is_external() {
                return Err(invalid("externalBook target relationship must be external"));
            }
            if !matches!(relationship.reltype(), rt::EXTERNAL_LINK_PATH | rt::STRICT_EXTERNAL_LINK_PATH) {
                return Err(invalid(format!("externalBook target has invalid relationship type '{}'", relationship.reltype())));
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
        ParsedKind::Dde => ExternalLinkKind::DdeOpaque,
        ParsedKind::Ole => ExternalLinkKind::OleOpaque,
    };
    Ok(ExternalLinkEntry {
        index,
        relationship_id: workbook_relationship_id,
        part_uri: part.partname().clone(),
        kind,
    })
}

fn parse_external_link(xml: &[u8]) -> Result<ParsedKind> {
    let xml = crate::common::mce::process_ooxml(xml)?;
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
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink") {
                    return Err(invalid("external-link XML must have one SpreadsheetML externalLink root"));
                }
                stack.push(Context::Root);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink") {
                    return Err(invalid("external-link XML must have one SpreadsheetML externalLink root"));
                }
                return Err(invalid("externalLink must contain a link kind"));
            },
            Event::Start(element) => {
                let parent = *stack.last().ok_or_else(|| invalid("external-link XML is missing its root"))?;
                let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = *stack.last().ok_or_else(|| invalid("external-link XML is missing its root"))?;
                parser.start(parent, &namespace, &element, decoder, &resolver)?;
            },
            Event::Text(text) => if let Some(Context::Value(sheet, row, cell)) = stack.last().copied() {
                parser.push_value(sheet, row, cell, &text.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?)?;
            },
            Event::CData(text) => if let Some(Context::Value(sheet, row, cell)) = stack.last().copied() {
                parser.push_value(sheet, row, cell, &text.decode().map_err(|e| OoxmlError::Xml(e.to_string()))?)?;
            },
            Event::GeneralRef(reference) => if let Some(Context::Value(sheet, row, cell)) = stack.last().copied() {
                parser.push_value(sheet, row, cell, &decode_xml_reference(&reference)?)?;
            },
            Event::End(element) => {
                let context = stack.pop().ok_or_else(|| invalid("external-link closing element outside root"))?;
                if context == Context::Root {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"externalLink") {
                        return Err(invalid("external-link XML has invalid root closing element"));
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => return Err(invalid("external-link XML has an unterminated root")),
            Event::Eof => break,
            _ => {},
        }
    }
    parser.kind.ok_or_else(|| invalid("externalLink must contain a link kind"))
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
        value => Err(invalid(format!("invalid external cached cell type '{value}'"))),
    }
}

fn validate_cell_reference(value: &str) -> Result<()> {
    let mut column = 0u32;
    let mut split = 0usize;
    for byte in value.bytes() {
        if byte.is_ascii_alphabetic() && split == 0 {
            column = column.checked_mul(26).and_then(|n| n.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))).ok_or_else(|| invalid("external cell column overflow"))?;
        } else { break; }
        split += 1;
    }
    if split == 0 || split == value.len() || column > 16_384 { return Err(invalid(format!("invalid external cached cell reference '{value}'"))); }
    let row = value[split..].parse::<u32>().map_err(|_| invalid(format!("invalid external cached cell reference '{value}'")))?;
    if row == 0 || row > 1_048_576 { return Err(invalid(format!("invalid external cached cell reference '{value}'"))); }
    Ok(())
}

fn required_attr(element: &BytesStart<'_>, name: &[u8], decoder: Decoder, description: &str) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn required_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder, description: &str) -> Result<u32> {
    optional_u32(element, name, decoder, description)?.ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn optional_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder, description: &str) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?.map(|value| value.parse::<u32>().map_err(|_| invalid(format!("invalid {description} '{value}'")))).transpose()
}
fn optional_bool(element: &BytesStart<'_>, name: &[u8], decoder: Decoder, description: &str) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?.map(|value| match value.as_str() { "1" | "true" | "on" => Ok(true), "0" | "false" | "off" => Ok(false), _ => Err(invalid(format!("invalid {description} '{value}'"))) }).transpose()
}
fn mark_once(seen: &mut bool, description: &str) -> Result<()> { if std::mem::replace(seen, true) { Err(invalid(format!("duplicate {description}"))) } else { Ok(()) } }
fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
fn limit(name: &str) -> OoxmlError { invalid(format!("external-link {name} limit exceeded")) }

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    #[test]
    fn parses_sparse_lexical_cache_without_dereferencing_target() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Data"/></sheetNames><definedNames><definedName name="Rate" refersTo="Data!$A$1" sheetId="1"/></definedNames><sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1" t="str"><v>001.2300</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(), litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(), xml.to_vec());
        part.relate_to_ext("https://127.0.0.1:9/never-open.xlsx", rt::EXTERNAL_LINK_PATH);
        let link = load_external_link(&part, "bookRel".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else { panic!("expected workbook link") };
        assert_eq!(book.target.target, "https://127.0.0.1:9/never-open.xlsx");
        assert_eq!(book.sheet_names, ["Data"]);
        assert_eq!(book.cached_sheets[0].rows[0].cells[0].raw_value.as_deref(), Some("001.2300"));
    }

    #[test]
    fn retains_dde_and_ole_as_opaque_kinds() {
        for (child, expected) in [
            ("<ddeLink ddeService=\"x\" ddeTopic=\"y\"/>", ExternalLinkKind::DdeOpaque),
            ("<oleLink r:id=\"rId1\" progId=\"x\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>", ExternalLinkKind::OleOpaque),
        ] {
            let xml = format!("<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{child}</externalLink>");
            let part = BlobPart::new(PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(), litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(), xml.into_bytes());
            assert_eq!(load_external_link(&part, "rId1".into(), 1).unwrap().kind, expected);
        }
    }

    #[test]
    fn mce_fallback_is_semantic_input() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:no"><mc:AlternateContent><mc:Choice Requires="x"><x:no/></mc:Choice><mc:Fallback><externalBook r:id="rId1"/></mc:Fallback></mc:AlternateContent></externalLink>"#;
        let mut part = BlobPart::new(PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(), litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(), xml.to_vec());
        part.relate_to_ext("opaque.xlsx", rt::EXTERNAL_LINK_PATH);
        assert!(matches!(load_external_link(&part, "rId1".into(), 1).unwrap().kind, ExternalLinkKind::Workbook(_)));
    }

    #[test]
    fn accepts_strict_external_workbook_namespaces_and_relationships() {
        let xml = br#"<externalLink xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Strict"/></sheetNames></externalBook></externalLink>"#;
        let mut part = BlobPart::new(PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(), litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(), xml.to_vec());
        part.relate_to_ext("strict-target.xlsx", rt::STRICT_EXTERNAL_LINK_PATH);
        let link = load_external_link(&part, "rId1".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else { panic!("expected workbook link") };
        assert_eq!(book.sheet_names, ["Strict"]);
        assert_eq!(book.target.relationship_type, rt::STRICT_EXTERNAL_LINK_PATH);
    }

    #[test]
    fn rejects_duplicate_and_over_limit_workbook_external_references() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let duplicate = format!(r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences><externalReference r:id="link"/><externalReference r:id="link"/></externalReferences></workbook>"#);
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&duplicate).is_err());

        let mut oversized = format!(r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences>"#);
        for index in 0..=4096 {
            oversized.push_str(&format!(r#"<externalReference r:id="link{index}"/>"#));
        }
        oversized.push_str("</externalReferences></workbook>");
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&oversized).is_err());
    }

    #[test]
    fn rejects_internal_targets_and_malformed_caches() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetDataSet><sheetData sheetId="0"/></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(), litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(), xml.to_vec());
        part.relate_to("../workbook.xml", rt::EXTERNAL_LINK_PATH);
        assert!(load_external_link(&part, "rId1".into(), 1).is_err());
    }

    #[test]
    fn loads_poi_ordered_external_workbook_reference() {
        let package = litchi_opc::OpcPackage::from_bytes(include_bytes!(
            "../../../../3rdparty/poi/test-data/spreadsheet/link-external-workbook-b.xlsx"
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
            "../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
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
            assert!(values.contains(&expected), "missing cached value {expected}");
        }
    }

    #[test]
    fn unrelated_modified_save_preserves_external_part_and_relationship() {
        const FIXTURE: &[u8] = include_bytes!(
            "../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
        );
        let original = litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap();
        let original_workbook = crate::xlsx::Workbook::new(
            litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap(),
        )
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

        let mut workbook = crate::xlsx::Workbook::new(
            litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap(),
        )
        .unwrap();
        workbook.define_name("Unrelated", "1");
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("external-link-roundtrip.xlsx");
        workbook.save(&path).unwrap();

        let bytes = std::fs::read(&path).unwrap();
        let saved = litchi_opc::OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(saved.get_part(&external_uri).unwrap().blob(), original_external);
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
            crate::xlsx::Workbook::new(saved).unwrap().external_links().len(),
            original_workbook.external_links().len()
        );
    }
}
