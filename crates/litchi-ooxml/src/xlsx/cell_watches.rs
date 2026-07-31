//! Worksheet watch-window entries (`CT_CellWatches`, `CT_CellWatch`).
//!
//! The watch window tracks the values of a small set of cells while the user
//! edits elsewhere. This module parses and serializes the worksheet
//! `cellWatches` collection; it never evaluates the watched cells.

use std::fmt::Write;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_str;

const TRANSITIONAL_MAIN: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_CELL_WATCHES: usize = 65_536;
const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// Namespace form used when serializing a cell-watches fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetCellWatchConformance {
    Transitional,
    Strict,
}

impl WorksheetCellWatchConformance {
    fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => std::str::from_utf8(TRANSITIONAL_MAIN).unwrap(),
            Self::Strict => std::str::from_utf8(STRICT_MAIN).unwrap(),
        }
    }
}

/// A validated A1 cell reference (`ST_CellRef`) naming one watched cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellWatchReference(String);

impl CellWatchReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_cell_reference(&value)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// The worksheet `cellWatches` collection in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetCellWatches {
    references: Vec<CellWatchReference>,
}

impl WorksheetCellWatches {
    pub fn new(references: Vec<CellWatchReference>) -> Result<Self> {
        if references.is_empty() {
            return Err(invalid("cellWatches requires at least one cellWatch"));
        }
        if references.len() > MAX_CELL_WATCHES {
            return Err(invalid(format!(
                "cellWatches exceeds safety limit {MAX_CELL_WATCHES}"
            )));
        }
        Ok(Self { references })
    }

    pub fn references(&self) -> &[CellWatchReference] {
        &self.references
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    CellWatches,
    CellWatch,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Unbound,
    Main,
    Other,
}

/// Parses the direct worksheet `cellWatches` child after applying shared MCE processing.
pub fn parse_worksheet_cell_watches(xml: &[u8]) -> Result<Option<WorksheetCellWatches>> {
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    let processed = process_str(source)?;
    let mut reader = NsReader::from_reader(processed.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = Vec::new();
    let mut references: Option<Vec<CellWatchReference>> = None;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        let namespace = namespace_kind(resolved)?;
        match event {
            Event::Start(element) => {
                let scope = begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut references,
                )?;
                scopes.push(scope);
            },
            Event::Empty(element) => {
                begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut references,
                )?;
            },
            Event::End(_) => {
                scopes
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet end element"))?;
            },
            Event::Text(text)
                if matches!(scopes.last(), Some(Scope::CellWatches | Scope::CellWatch))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("cellWatches family cannot contain text"));
            },
            Event::CData(text)
                if matches!(scopes.last(), Some(Scope::CellWatches | Scope::CellWatch))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("cellWatches family cannot contain CDATA"));
            },
            Event::DocType(_) => {
                return Err(invalid("worksheet XML cannot contain a document type"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !scopes.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    match references {
        None => Ok(None),
        Some(references) => Ok(Some(WorksheetCellWatches::new(references)?)),
    }
}

fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    references: &mut Option<Vec<CellWatchReference>>,
) -> Result<Scope> {
    let local = element.local_name();
    let local = local.as_ref();
    let main = namespace == NamespaceKind::Main;
    match parent {
        None => {
            if !main || local != b"worksheet" {
                return Err(invalid("expected SpreadsheetML worksheet root"));
            }
            Ok(Scope::Worksheet)
        },
        Some(Scope::Worksheet) => {
            if local != b"cellWatches" {
                return Ok(Scope::Other);
            }
            if !main {
                return Err(invalid("spoofed cellWatches element namespace"));
            }
            if references.is_some() {
                return Err(invalid("duplicate worksheet cellWatches element"));
            }
            reject_attributes(element, "cellWatches")?;
            *references = Some(Vec::new());
            Ok(Scope::CellWatches)
        },
        Some(Scope::CellWatches) => {
            if local != b"cellWatch" || !main {
                return Err(invalid(if local == b"cellWatch" {
                    "spoofed cellWatch element namespace"
                } else {
                    "unknown cellWatches child element"
                }));
            }
            let collection = references
                .as_mut()
                .ok_or_else(|| invalid("missing cellWatches state"))?;
            if collection.len() >= MAX_CELL_WATCHES {
                return Err(invalid(format!(
                    "cellWatches exceeds safety limit {MAX_CELL_WATCHES}"
                )));
            }
            collection.push(parse_cell_watch_attributes(reader, element)?);
            Ok(Scope::CellWatch)
        },
        Some(Scope::CellWatch) => Err(invalid("cellWatch must be a leaf element")),
        Some(Scope::Other) => Ok(Scope::Other),
    }
}

fn parse_cell_watch_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<CellWatchReference> {
    let mut reference = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid cellWatch attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound || local.as_ref() != b"r" {
            return Err(invalid("unknown or spoofed cellWatch attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid cellWatch attribute value: {error}")))?
            .into_owned();
        if reference.is_some() {
            return Err(invalid("duplicate cellWatch r attribute"));
        }
        reference = Some(CellWatchReference::new(text)?);
    }
    reference.ok_or_else(|| invalid("cellWatch requires r"))
}

/// Serializes one canonical, namespace-complete `cellWatches` fragment.
pub fn write_worksheet_cell_watches(
    value: &WorksheetCellWatches,
    conformance: WorksheetCellWatchConformance,
) -> Result<String> {
    if value.references.is_empty() {
        return Err(invalid("cellWatches requires at least one cellWatch"));
    }
    if value.references.len() > MAX_CELL_WATCHES {
        return Err(invalid(format!(
            "cellWatches exceeds safety limit {MAX_CELL_WATCHES}"
        )));
    }
    let mut xml = String::new();
    write!(
        xml,
        "<cellWatches xmlns=\"{}\">",
        conformance.main_namespace()
    )
    .unwrap();
    for reference in &value.references {
        xml.push_str("<cellWatch r=\"");
        xml.push_str(reference.as_str());
        xml.push_str("\"/>");
    }
    xml.push_str("</cellWatches>");
    Ok(xml)
}

fn reject_attributes(element: &BytesStart<'_>, name: &str) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid {name} attribute: {error}")))?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "unexpected {name} attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn namespace_kind(result: ResolveResult<'_>) -> Result<NamespaceKind> {
    match result {
        ResolveResult::Unbound => Ok(NamespaceKind::Unbound),
        ResolveResult::Bound(namespace) if is_main_namespace(namespace.as_ref()) => {
            Ok(NamespaceKind::Main)
        },
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix {}",
            String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn is_main_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_MAIN || namespace == STRICT_MAIN
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn validate_cell_reference(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start || index - column_start > 3 {
        return Err(invalid(format!("invalid cellWatch reference '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    if column == 0 || column > MAX_COLUMN {
        return Err(invalid(format!(
            "cellWatch column is out of range in '{value}'"
        )));
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return Err(invalid(format!("invalid cellWatch reference '{value}'")));
    }
    let row = value[row_start..]
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid cellWatch row in '{value}'")))?;
    if row == 0 || row > MAX_ROW {
        return Err(invalid(format!(
            "cellWatch row is out of range in '{value}'"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetCellWatches>> {
        parse_worksheet_cell_watches(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_watches_in_document_order() {
        let value = parse(r#"<cellWatches><cellWatch r="A1"/><cellWatch r="$C$10"/><cellWatch r="XFD1048576"/></cellWatches>"#)
            .unwrap()
            .unwrap();
        assert_eq!(
            value
                .references()
                .iter()
                .map(CellWatchReference::as_str)
                .collect::<Vec<_>>(),
            vec!["A1", "$C$10", "XFD1048576"]
        );
        assert!(parse("<sheetData/>").unwrap().is_none());
    }

    #[test]
    fn supports_strict_namespace() {
        let xml = concat!(
            r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">"#,
            r#"<cellWatches><cellWatch r="B2"/></cellWatches></worksheet>"#,
        );
        let value = parse_worksheet_cell_watches(xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(value.references()[0].as_str(), "B2");
    }

    #[test]
    fn rejects_structure_attributes_and_references() {
        for child in [
            "<cellWatches/>",
            "<cellWatches><cellWatch/></cellWatches>",
            r#"<cellWatches><cellWatch r="A0"/></cellWatches>"#,
            r#"<cellWatches><cellWatch r="XFE1"/></cellWatches>"#,
            r#"<cellWatches><cellWatch r="A1048577"/></cellWatches>"#,
            r#"<cellWatches><cellWatch r="A1:B2"/></cellWatches>"#,
            r#"<cellWatches><cellWatch r="A1" mystery="1"/></cellWatches>"#,
            r#"<cellWatches unexpected="1"><cellWatch r="A1"/></cellWatches>"#,
            r#"<cellWatches><cellWatch r="A1"><child/></cellWatch></cellWatches>"#,
            r#"<cellWatches>text</cellWatches>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
        assert!(parse(
            "<cellWatches><cellWatch r=\"A1\"/></cellWatches><cellWatches><cellWatch r=\"B2\"/></cellWatches>"
        )
        .is_err());
    }

    #[test]
    fn write_round_trips_through_the_reader() {
        let expected = WorksheetCellWatches::new(vec![
            CellWatchReference::new("A1").unwrap(),
            CellWatchReference::new("$D$7").unwrap(),
        ])
        .unwrap();
        for conformance in [
            WorksheetCellWatchConformance::Transitional,
            WorksheetCellWatchConformance::Strict,
        ] {
            let fragment = write_worksheet_cell_watches(&expected, conformance).unwrap();
            let document = format!(r#"<worksheet xmlns="{NS}">{fragment}</worksheet>"#);
            let parsed = parse_worksheet_cell_watches(document.as_bytes())
                .unwrap()
                .unwrap();
            assert_eq!(parsed, expected);
        }
        assert!(WorksheetCellWatches::new(Vec::new()).is_err());
    }
}
