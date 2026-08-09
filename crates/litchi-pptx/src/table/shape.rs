//! Borrowed DrawingML table views embedded in PresentationML graphic frames.
//!
//! The table, row, and cell values retain the caller's XML allocation. Only
//! compact byte spans are allocated while indexing, so reading a large table
//! does not duplicate its XML subtrees.

use litchi_ooxml_common::xml::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, decode_xml_reference,
    unqualified_attribute_value,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_ROWS: usize = 100_000;
const MAX_CELLS: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Span {
    start: usize,
    end: usize,
}

impl Span {
    fn get(self, xml: &[u8]) -> Result<&[u8]> {
        xml.get(self.start..self.end)
            .ok_or_else(|| invalid("table byte span is outside its owner XML"))
    }
}

/// Table style switches and the referenced table-style GUID.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub first_row: bool,
    pub first_col: bool,
    pub last_row: bool,
    pub last_col: bool,
    pub band_row: bool,
    pub band_col: bool,
    pub style_id: Option<String>,
}

/// A borrowed `DrawingML` table (`a:tbl`).
#[derive(Debug, Clone, Copy)]
pub struct Table<'a> {
    xml: &'a [u8],
    span: Span,
}

impl<'a> Table<'a> {
    /// Index a standalone `a:tbl` owner without copying its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(xml: &'a [u8]) -> Result<Self> {
        let span = scan(xml, b"tbl")?
            .into_iter()
            .next()
            .ok_or_else(|| Error::PartNotFound("DrawingML table not found".into()))?;
        Ok(Self { xml, span })
    }

    /// Locate the first `DrawingML` table inside a graphic-frame owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_graphic_frame(xml: &'a [u8]) -> Result<Self> {
        Self::from_xml(xml)
    }

    /// Borrow the exact table element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    /// Count table rows without allocating row values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn row_count(self) -> Result<usize> {
        let table = self.xml()?;
        let count = scan(table, b"tr")?.len();
        if count > MAX_ROWS {
            return Err(limit("table row count", MAX_ROWS));
        }
        Ok(count)
    }

    /// Count columns from the first row, or return zero for an empty table.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn column_count(self) -> Result<usize> {
        Ok(self.rows()?.first().map_or(0, |row| row.cell_count()))
    }

    /// Index rows as borrowed views over the table's source XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn rows(self) -> Result<Vec<Row<'a>>> {
        let table = self.xml()?;
        let spans = scan(table, b"tr")?;
        if spans.len() > MAX_ROWS {
            return Err(limit("table row count", MAX_ROWS));
        }
        spans
            .into_iter()
            .map(|span| Ok(Row { xml: table, span }))
            .collect()
    }

    /// Get one cell by zero-based row and column position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn cell(self, row: usize, column: usize) -> Result<Option<Cell<'a>>> {
        let Some(row) = self.rows()?.get(row).copied() else {
            return Ok(None);
        };
        Ok(row.cells()?.get(column).copied())
    }

    /// Read the optional `a:tblPr` values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn properties(self) -> Result<Option<Properties>> {
        parse_properties(self.xml()?)
    }
}

/// A borrowed table row.
#[derive(Debug, Clone, Copy)]
pub struct Row<'a> {
    xml: &'a [u8],
    span: Span,
}

impl<'a> Row<'a> {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    #[must_use]
    pub fn cell_count(self) -> usize {
        scan(self.xml().unwrap_or_default(), b"tc").map_or(0, |spans| spans.len())
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn cells(self) -> Result<Vec<Cell<'a>>> {
        let row = self.xml()?;
        let spans = scan(row, b"tc")?;
        if spans.len() > MAX_CELLS {
            return Err(limit("table cell count", MAX_CELLS));
        }
        spans
            .into_iter()
            .map(|span| Ok(Cell { xml: row, span }))
            .collect()
    }
}

/// A borrowed table cell.
#[derive(Debug, Clone, Copy)]
pub struct Cell<'a> {
    xml: &'a [u8],
    span: Span,
}

impl<'a> Cell<'a> {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.span.get(self.xml)
    }

    /// Decode `DrawingML` text, preserving paragraph, break, and tab markers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(self) -> Result<String> {
        extract_text(self.xml()?)
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn is_drawingml(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => {
            prefix.as_slice() == b"a"
                || fragment_prefix.as_ref().and_then(Option::as_deref) == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
    }
}

fn scan(xml: &[u8], target: &[u8]) -> Result<Vec<Span>> {
    enum ScanEvent {
        Start { local: Vec<u8>, matched: bool },
        Empty { matched: bool },
        End { local: Vec<u8> },
        Other,
        Forbidden,
        Eof,
    }

    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table XML bytes", MAX_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut stack: Vec<(usize, Vec<u8>, bool)> = Vec::new();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut spans = Vec::new();
    let mut nodes = 0usize;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("table XML offset exceeds usize"))?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix =
                            Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                    }
                    let matched =
                        is_drawingml(&namespace, element.name(), target, &fragment_prefix);
                    ScanEvent::Start {
                        local: element.local_name().as_ref().to_vec(),
                        matched,
                    }
                },
                Event::Empty(element) => {
                    if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix =
                            Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                    }
                    ScanEvent::Empty {
                        matched: is_drawingml(&namespace, element.name(), target, &fragment_prefix),
                    }
                },
                Event::End(element) => ScanEvent::End {
                    local: element.local_name().as_ref().to_vec(),
                },
                Event::DocType(_) | Event::PI(_) => ScanEvent::Forbidden,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("table XML offset exceeds usize"))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("table XML node count overflow"))?;
        if nodes > MAX_NODES {
            return Err(limit("table XML nodes", MAX_NODES));
        }

        match event {
            ScanEvent::Start { local, matched } => {
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("table XML depth", MAX_DEPTH));
                }
                stack.push((start, local, matched));
            },
            ScanEvent::Empty { matched } => {
                if matched {
                    spans.push(Span { start, end });
                }
            },
            ScanEvent::End { local: closing } => {
                let Some((open, local, matched)) = stack.pop() else {
                    return Err(invalid("table XML has an unexpected closing element"));
                };
                if local.as_slice() != closing.as_slice() {
                    return Err(invalid("table XML closing element does not match"));
                }
                if matched {
                    spans.push(Span { start: open, end });
                }
            },
            ScanEvent::Forbidden => return Err(invalid("table XML contains forbidden markup")),
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("table XML is unterminated"));
    }
    Ok(spans)
}

fn parse_properties(xml: &[u8]) -> Result<Option<Properties>> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = None;
    let mut depth = 0usize;
    let mut properties = None;
    let mut property_depth = None;
    let mut style_depth = None;
    let mut style = String::new();

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if depth == 1 && !is_drawingml(&namespace, element.name(), b"tbl", &fragment_prefix)
                {
                    return Ok(None);
                }
                if depth == 2
                    && is_drawingml(&namespace, element.name(), b"tblPr", &fragment_prefix)
                {
                    if properties.is_some() {
                        return Err(invalid("table has multiple tblPr elements"));
                    }
                    properties = Some(Properties {
                        first_row: on_off(&element, b"firstRow", decoder)?,
                        first_col: on_off(&element, b"firstCol", decoder)?,
                        last_row: on_off(&element, b"lastRow", decoder)?,
                        last_col: on_off(&element, b"lastCol", decoder)?,
                        band_row: on_off(&element, b"bandRow", decoder)?,
                        band_col: on_off(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                    property_depth = Some(depth);
                } else if property_depth == Some(depth - 1)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                    )
                    && style_depth.replace(depth).is_some()
                {
                    return Err(invalid("table has multiple tableStyleId elements"));
                }
            },
            Event::Empty(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table XML depth overflow"))?;
                if child_depth == 1
                    && !is_drawingml(&namespace, element.name(), b"tbl", &fragment_prefix)
                {
                    return Ok(None);
                }
                if child_depth == 2
                    && is_drawingml(&namespace, element.name(), b"tblPr", &fragment_prefix)
                {
                    if properties.is_some() {
                        return Err(invalid("table has multiple tblPr elements"));
                    }
                    properties = Some(Properties {
                        first_row: on_off(&element, b"firstRow", decoder)?,
                        first_col: on_off(&element, b"firstCol", decoder)?,
                        last_row: on_off(&element, b"lastRow", decoder)?,
                        last_col: on_off(&element, b"lastCol", decoder)?,
                        band_row: on_off(&element, b"bandRow", decoder)?,
                        band_col: on_off(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                }
            },
            Event::Text(text) if style_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                style.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::End(element) => {
                if style_depth == Some(depth)
                    && is_drawingml(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                    )
                {
                    let value = style.trim();
                    if value.is_empty() {
                        return Err(invalid("tableStyleId must not be empty"));
                    }
                    if let Some(properties) = properties.as_mut() {
                        properties.style_id = Some(value.to_owned());
                    }
                    style_depth = None;
                    style.clear();
                }
                if property_depth == Some(depth)
                    && is_drawingml(&namespace, element.name(), b"tblPr", &fragment_prefix)
                {
                    property_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("table XML nesting underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("table XML contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || property_depth.is_some() || style_depth.is_some() {
        return Err(invalid("table XML is unterminated"));
    }
    Ok(properties)
}

fn on_off(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    match unqualified_attribute_value(element, name, decoder)? {
        None => Ok(false),
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(invalid(format!("invalid table property value '{value}'"))),
        },
    }
}

fn extract_text(xml: &[u8]) -> Result<String> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table cell XML bytes", MAX_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = None;
    let mut depth = 0usize;
    let mut text_depth = None;
    let mut result = String::new();
    let mut paragraph_seen = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("table cell XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("table cell XML depth", MAX_DEPTH));
                }
                if is_drawingml(&namespace, element.name(), b"p", &fragment_prefix) {
                    if paragraph_seen && !result.is_empty() && !result.ends_with(' ') {
                        result.push(' ');
                    }
                    paragraph_seen = true;
                } else if text_depth.is_none()
                    && is_drawingml(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if is_drawingml(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Empty(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix =
                        Some(element.name().prefix().map(|p| p.into_inner().to_vec()));
                }
                if is_drawingml(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::CData(text) if text_depth.is_some() => {
                result.push_str(
                    &text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_drawingml(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("table cell XML nesting underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("table cell XML contains forbidden markup"));
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(invalid("table cell XML is unterminated"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(result.trim().to_owned())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #[test]
    fn table_views_borrow_rows_and_cells() {
        let xml = format!(
            r#"<a:graphicFrame xmlns:a="{DML}"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{{id}}</a:tableStyleId></a:tblPr><a:tr><a:tc><a:txBody><a:p><a:r><a:t>A &amp; B</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicFrame>"#
        );
        let table = Table::from_graphic_frame(xml.as_bytes()).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
        assert_eq!(table.column_count().unwrap(), 1);
        assert_eq!(
            table.rows().unwrap()[0].cells().unwrap()[0].text().unwrap(),
            "A & B"
        );
        assert!(table.properties().unwrap().unwrap().first_row);
    }

    #[test]
    fn inherited_prefixes_and_foreign_names_are_checked() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign"><x:tr/><a:tr><a:tc/></a:tr></a:tbl>"#;
        let table = Table::from_xml(xml).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
    }
}
