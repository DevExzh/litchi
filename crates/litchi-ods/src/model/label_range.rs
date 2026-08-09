//! Spreadsheet row and column label ranges.

use super::structure::validate_cell_range_addresses;
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";

trait HasLocalName {
    fn has_local_name(&self, expected: &[u8]) -> bool;
}

impl HasLocalName for BytesStart<'_> {
    fn has_local_name(&self, expected: &[u8]) -> bool {
        self.local_name().as_ref() == expected
    }
}

impl HasLocalName for BytesEnd<'_> {
    fn has_local_name(&self, expected: &[u8]) -> bool {
        self.local_name().as_ref() == expected
    }
}

/// Whether labels identify spreadsheet rows or columns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Orientation {
    /// Labels identify columns in the associated data range.
    Column,
    /// Labels identify rows in the associated data range.
    Row,
}

impl Orientation {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "column" => Ok(Self::Column),
            "row" => Ok(Self::Row),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:orientation value '{value}'"
            ))),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Column => "column",
            Self::Row => "row",
        }
    }
}

/// A cell range whose values label rows or columns in another cell range.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Range {
    /// ODF address of the cells containing labels.
    pub label_cell_range_address: String,
    /// ODF address of the cells to which the labels apply.
    pub data_cell_range_address: String,
    /// Whether the labels identify rows or columns.
    pub orientation: Orientation,
}

impl Range {
    /// Create a validated label range.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(
        label_cell_range_address: impl Into<String>,
        data_cell_range_address: impl Into<String>,
        orientation: Orientation,
    ) -> Result<Self> {
        let range = Self {
            label_cell_range_address: label_cell_range_address.into(),
            data_cell_range_address: data_cell_range_address.into(),
            orientation,
        };
        range.validate()?;
        Ok(range)
    }

    /// Validate both ODF cell-range address attributes.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        validate_cell_range_addresses(std::slice::from_ref(&self.label_cell_range_address))?;
        validate_cell_range_addresses(std::slice::from_ref(&self.data_cell_range_address))
    }
}

/// # Errors
///
/// Returns an error when the input is malformed or exceeds the parser's resource limits.
pub fn parse(xml: &str) -> Result<Vec<Range>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut ranges = Vec::new();
    let mut seen = false;
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        let consumes_container = matches!(
            &event,
            Event::Start(element)
                if is_table_element(&namespace, element, b"label-ranges")
        );

        if let Event::Start(element) = &event
            && is_namespace(&namespace, OFFICE_NAMESPACE)
            && element.local_name().as_ref() == b"spreadsheet"
        {
            spreadsheet_depth = Some(depth);
        }
        let is_spreadsheet_child = spreadsheet_depth.is_some_and(|value| depth == value + 1);

        match event {
            Event::Start(element) if is_table_element(&namespace, &element, b"label-ranges") => {
                ensure_container_location(is_spreadsheet_child, seen)?;
                seen = true;
                ranges = parse_container(&mut reader)?;
            },
            Event::Empty(element) if is_table_element(&namespace, &element, b"label-ranges") => {
                ensure_container_location(is_spreadsheet_child, seen)?;
                seen = true;
            },
            Event::Start(element) | Event::Empty(element)
                if is_table_element(&namespace, &element, b"label-range") =>
            {
                return Err(Error::InvalidFormat(
                    "table:label-range must be inside table:label-ranges".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }

        // A non-empty container is consumed in full by `parse_container`, so it
        // does not contribute to the outer reader's open-element depth.
        if is_start && !consumes_container {
            depth = depth.saturating_add(1);
        } else if is_end {
            depth = depth.saturating_sub(1);
        }
        buf.clear();
    }

    Ok(ranges)
}

fn ensure_container_location(is_spreadsheet_child: bool, seen: bool) -> Result<()> {
    if !is_spreadsheet_child {
        return Err(Error::InvalidFormat(
            "table:label-ranges must be a direct office:spreadsheet child".to_string(),
        ));
    }
    if seen {
        return Err(Error::InvalidFormat(
            "duplicate table:label-ranges".to_string(),
        ));
    }
    Ok(())
}

fn parse_container(reader: &mut NsReader<&[u8]>) -> Result<Vec<Range>> {
    let mut ranges = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) if is_table_element(&namespace, &element, b"label-range") => {
                let range = parse_range_attributes(reader.resolver(), reader.decoder(), &element)?;
                consume_empty_range(reader)?;
                ranges.push(range);
            },
            Event::Empty(element) if is_table_element(&namespace, &element, b"label-range") => {
                ranges.push(parse_range_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::End(element) if is_table_element(&namespace, &element, b"label-ranges") => {
                break;
            },
            Event::Text(text) => {
                ensure_whitespace(&text.xml_content(XmlVersion::Explicit1_0).map_err(
                    |error| Error::InvalidFormat(format!("invalid label-range text: {error}")),
                )?)?;
            },
            Event::CData(text) => {
                ensure_whitespace(&text.xml_content(XmlVersion::Explicit1_0).map_err(
                    |error| Error::InvalidFormat(format!("invalid label-range CDATA: {error}")),
                )?)?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table:label-ranges".to_string(),
                ));
            },
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(Error::InvalidFormat(
                    "unsupported table:label-ranges child".to_string(),
                ));
            },
        }
        buf.clear();
    }
    Ok(ranges)
}

fn consume_empty_range(reader: &mut NsReader<&[u8]>) -> Result<()> {
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::End(element) if is_table_element(&namespace, &element, b"label-range") => {
                return Ok(());
            },
            Event::Text(text) => {
                ensure_whitespace(&text.xml_content(XmlVersion::Explicit1_0).map_err(
                    |error| Error::InvalidFormat(format!("invalid label-range text: {error}")),
                )?)?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table:label-range".to_string(),
                ));
            },
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(Error::InvalidFormat(
                    "table:label-range must be empty".to_string(),
                ));
            },
        }
        buf.clear();
    }
}

fn parse_range_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<Range> {
    let range = Range {
        label_cell_range_address: required_attribute(
            resolver,
            decoder,
            element,
            b"label-cell-range-address",
        )?,
        data_cell_range_address: required_attribute(
            resolver,
            decoder,
            element,
            b"data-cell-range-address",
        )?,
        orientation: Orientation::parse(&required_attribute(
            resolver,
            decoder,
            element,
            b"orientation",
        )?)?,
    };
    range.validate()?;
    Ok(range)
}

fn required_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid label-range attribute: {error}"))
        })?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, TABLE_NAMESPACE) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(std::borrow::Cow::into_owned)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid label-range attribute value: {error}"))
                });
        }
    }
    Err(Error::InvalidFormat(format!(
        "missing table:{}",
        String::from_utf8_lossy(local_name)
    )))
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
pub fn write(out: &mut String, ranges: &[Range]) -> Result<()> {
    if ranges.is_empty() {
        return Ok(());
    }
    out.push_str("<table:label-ranges>");
    for range in ranges {
        range.validate()?;
        out.push_str("<table:label-range table:label-cell-range-address=\"");
        out.push_str(&escape_xml(&range.label_cell_range_address));
        out.push_str("\" table:data-cell-range-address=\"");
        out.push_str(&escape_xml(&range.data_cell_range_address));
        out.push_str("\" table:orientation=\"");
        out.push_str(range.orientation.as_str());
        out.push_str("\"/>");
    }
    out.push_str("</table:label-ranges>");
    Ok(())
}

fn ensure_whitespace(value: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "table:label-ranges cannot contain text".to_string(),
        ))
    }
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn is_table_element(
    namespace: &ResolveResult<'_>,
    element: &impl HasLocalName,
    local_name: &[u8],
) -> bool {
    is_namespace(namespace, TABLE_NAMESPACE) && element.has_local_name(local_name)
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("XML parsing error: {error}"))
}

// These integration cases require the full sheet authoring migration.
#[cfg(any())]
mod tests {
    use super::*;
    use crate::{Builder, MutableSpreadsheet, Spreadsheet};

    const PREFIX: &str = concat!(
        "<office:document-content ",
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
        "<office:body><office:spreadsheet>"
    );
    const SUFFIX: &str = "</office:spreadsheet></office:body></office:document-content>";

    #[test]
    fn parses_namespace_aliases_and_empty_container() {
        let xml = concat!(
            "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
            "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
            "<o:body><o:spreadsheet><t:label-ranges>",
            "<t:label-range t:label-cell-range-address=\"Sheet1.A1:A2\" ",
            "t:data-cell-range-address=\"Sheet1.B1:C2\" t:orientation=\"row\"></t:label-range>",
            "</t:label-ranges></o:spreadsheet></o:body></o:document-content>"
        );
        let parsed = parse(xml).unwrap();
        assert_eq!(
            parsed,
            vec![Range::new("Sheet1.A1:A2", "Sheet1.B1:C2", Orientation::Row).unwrap()]
        );
        assert!(
            parse(&format!("{PREFIX}<table:label-ranges/>{SUFFIX}"))
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn rejects_invalid_structure_and_attributes() {
        for fragment in [
            "<table:label-range table:label-cell-range-address=\"A1:A2\" table:data-cell-range-address=\"B1:B2\" table:orientation=\"row\"/>",
            "<table:table><table:label-ranges/></table:table>",
            "<table:label-ranges/><table:label-ranges/>",
            "<table:label-ranges><table:label-range table:label-cell-range-address=\"A1:A2\" table:data-cell-range-address=\"B1:B2\" table:orientation=\"diagonal\"/></table:label-ranges>",
            "<table:label-ranges><table:label-range table:label-cell-range-address=\"A1 A2\" table:data-cell-range-address=\"B1:B2\" table:orientation=\"column\"/></table:label-ranges>",
            "<table:label-ranges><table:label-range table:label-cell-range-address=\"A1:A2\" table:orientation=\"column\"/></table:label-ranges>",
            "<table:label-ranges><table:label-range table:label-cell-range-address=\"A1:A2\" table:data-cell-range-address=\"B1:B2\" table:orientation=\"column\"><table:x/></table:label-range></table:label-ranges>",
        ] {
            assert!(
                parse(&format!("{PREFIX}{fragment}{SUFFIX}")).is_err(),
                "accepted {fragment}"
            );
        }
    }

    #[test]
    fn writer_escapes_addresses() {
        let range = Range::new("'A&B'.A1:A2", "'A&B'.B1:B2", Orientation::Column).unwrap();
        let mut xml = String::new();
        write(&mut xml, &[range]).unwrap();
        assert!(xml.contains("&amp;"));
        let parsed = parse(&format!("{PREFIX}{xml}{SUFFIX}")).unwrap();
        assert_eq!(parsed[0].label_cell_range_address, "'A&B'.A1:A2");
    }

    #[test]
    fn round_trips_through_builder_and_mutable_packages() {
        let first = Range::new("Sheet1.A1:A2", "Sheet1.B1:D2", Orientation::Row).unwrap();
        let second = Range::new("Sheet1.A1:D1", "Sheet1.A2:D4", Orientation::Column).unwrap();

        let mut builder = Builder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_label_range(first.clone()).unwrap();
        let bytes = builder.build().unwrap();
        let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
        let xml = spreadsheet.content_xml();
        assert!(
            xml.find("<table:label-ranges>").unwrap() < xml.find("<table:table ").unwrap(),
            "label ranges must precede spreadsheet tables"
        );
        assert_eq!(spreadsheet.label_ranges(), &[first]);

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        mutable.add_label_range(second.clone()).unwrap();
        assert!(mutable.remove_label_range(0).is_some());
        let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.label_ranges(), &[second]);
    }
}
