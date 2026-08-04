//! Spreadsheet consolidation declarations.

use super::structure::{split_cell_range_addresses, validate_cell_range_addresses};
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

/// Which source labels participate in a consolidation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ConsolidationUseLabels {
    /// Do not match source data by labels.
    None,
    /// Match row labels.
    Row,
    /// Match column labels.
    Column,
    /// Match both row and column labels.
    Both,
}

impl ConsolidationUseLabels {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "row" => Ok(Self::Row),
            "column" => Ok(Self::Column),
            "both" => Ok(Self::Both),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:use-labels value '{value}'"
            ))),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Row => "row",
            Self::Column => "column",
            Self::Both => "both",
        }
    }
}

/// An inert spreadsheet consolidation configuration.
///
/// The function is retained as a string because ODF permits both its standard
/// function names and application-defined strings. This crate never evaluates
/// the consolidation or follows links to source data.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Consolidation {
    /// Standard or application-defined consolidation function.
    pub function: String,
    /// Ordered ODF source cell-range addresses.
    pub source_cell_range_addresses: Vec<String>,
    /// ODF target cell address.
    pub target_cell_address: String,
    /// Optional source-label matching policy.
    pub use_labels: Option<ConsolidationUseLabels>,
    /// Whether consumers should link results back to source data.
    pub link_to_source_data: Option<bool>,
}

impl Consolidation {
    /// Create a validated inert consolidation declaration.
    pub fn new(
        function: impl Into<String>,
        source_cell_range_addresses: Vec<String>,
        target_cell_address: impl Into<String>,
    ) -> Result<Self> {
        let consolidation = Self {
            function: function.into(),
            source_cell_range_addresses,
            target_cell_address: target_cell_address.into(),
            use_labels: None,
            link_to_source_data: None,
        };
        consolidation.validate()?;
        Ok(consolidation)
    }

    /// Validate address-list boundaries and the single target address.
    pub fn validate(&self) -> Result<()> {
        validate_cell_range_addresses(&self.source_cell_range_addresses)?;
        validate_target_cell_address(&self.target_cell_address)
    }
}

pub(crate) fn parse_consolidation(xml: &str) -> Result<Option<Consolidation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut result = None;
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        let consumes_element = matches!(
            &event,
            Event::Start(element) if is_table_element(&namespace, element, b"consolidation")
        );

        if let Event::Start(element) = &event
            && is_namespace(&namespace, OFFICE_NAMESPACE)
            && element.local_name().as_ref() == b"spreadsheet"
        {
            spreadsheet_depth = Some(depth);
        }
        let is_spreadsheet_child = spreadsheet_depth.is_some_and(|value| depth == value + 1);

        match event {
            Event::Start(element) if is_table_element(&namespace, &element, b"consolidation") => {
                ensure_location(is_spreadsheet_child, result.is_some())?;
                let consolidation =
                    parse_attributes(reader.resolver(), reader.decoder(), &element)?;
                consume_empty_element(&mut reader)?;
                result = Some(consolidation);
            },
            Event::Empty(element) if is_table_element(&namespace, &element, b"consolidation") => {
                ensure_location(is_spreadsheet_child, result.is_some())?;
                result = Some(parse_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::Eof => break,
            _ => {},
        }

        if is_start && !consumes_element {
            depth = depth.saturating_add(1);
        } else if is_end {
            depth = depth.saturating_sub(1);
        }
        buf.clear();
    }

    Ok(result)
}

fn ensure_location(is_spreadsheet_child: bool, seen: bool) -> Result<()> {
    if !is_spreadsheet_child {
        return Err(Error::InvalidFormat(
            "table:consolidation must be a direct office:spreadsheet child".to_string(),
        ));
    }
    if seen {
        return Err(Error::InvalidFormat(
            "duplicate table:consolidation".to_string(),
        ));
    }
    Ok(())
}

fn consume_empty_element(reader: &mut NsReader<&[u8]>) -> Result<()> {
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::End(element) if is_table_element(&namespace, &element, b"consolidation") => {
                return Ok(());
            },
            Event::Text(text) => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid consolidation text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:consolidation must be empty".to_string(),
                    ));
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table:consolidation".to_string(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "table:consolidation must be empty".to_string(),
                ));
            },
        }
        buf.clear();
    }
}

fn parse_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<Consolidation> {
    let source_value =
        required_attribute(resolver, decoder, element, b"source-cell-range-addresses")?;
    let consolidation = Consolidation {
        function: required_attribute(resolver, decoder, element, b"function")?,
        source_cell_range_addresses: split_cell_range_addresses(&source_value)?,
        target_cell_address: required_attribute(
            resolver,
            decoder,
            element,
            b"target-cell-address",
        )?,
        use_labels: optional_attribute(resolver, decoder, element, b"use-labels")?
            .map(|value| ConsolidationUseLabels::parse(&value))
            .transpose()?,
        link_to_source_data: optional_attribute(
            resolver,
            decoder,
            element,
            b"link-to-source-data",
        )?
        .map(|value| parse_bool("table:link-to-source-data", &value))
        .transpose()?,
    };
    consolidation.validate()?;
    Ok(consolidation)
}

fn required_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    optional_attribute(resolver, decoder, element, local_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "missing table:{}",
            String::from_utf8_lossy(local_name)
        ))
    })
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid consolidation attribute: {error}"))
        })?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, TABLE_NAMESPACE) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid consolidation attribute value: {error}"))
                });
        }
    }
    Ok(None)
}

fn parse_bool(name: &str, value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid {name} boolean '{value}'"
        ))),
    }
}

pub(crate) fn write_consolidation(
    out: &mut String,
    consolidation: Option<&Consolidation>,
) -> Result<()> {
    let Some(consolidation) = consolidation else {
        return Ok(());
    };
    consolidation.validate()?;
    out.push_str("<table:consolidation table:function=\"");
    out.push_str(&escape_xml(&consolidation.function));
    out.push_str("\" table:source-cell-range-addresses=\"");
    out.push_str(&escape_xml(
        &consolidation.source_cell_range_addresses.join(" "),
    ));
    out.push_str("\" table:target-cell-address=\"");
    out.push_str(&escape_xml(&consolidation.target_cell_address));
    out.push('"');
    if let Some(value) = consolidation.use_labels {
        out.push_str(" table:use-labels=\"");
        out.push_str(value.as_str());
        out.push('"');
    }
    if let Some(value) = consolidation.link_to_source_data {
        out.push_str(if value {
            " table:link-to-source-data=\"true\""
        } else {
            " table:link-to-source-data=\"false\""
        });
    }
    out.push_str("/>");
    Ok(())
}

fn contains_unquoted(value: &str, needle: char) -> bool {
    let mut quoted = false;
    let mut chars = value.chars().peekable();
    while let Some(character) = chars.next() {
        if character == '\'' {
            if quoted && chars.peek() == Some(&'\'') {
                chars.next();
            } else {
                quoted = !quoted;
            }
        } else if character == needle && !quoted {
            return true;
        }
    }
    false
}

fn validate_target_cell_address(value: &str) -> Result<()> {
    if value != value.trim() || contains_unquoted(value, ':') {
        return Err(invalid_target(value));
    }

    let mut quoted = false;
    let mut separator = None;
    let mut chars = value.char_indices().peekable();
    while let Some((index, character)) = chars.next() {
        if character == '\'' {
            if quoted && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                chars.next();
            } else {
                quoted = !quoted;
            }
        } else if character == '.' && !quoted && separator.replace(index).is_some() {
            return Err(invalid_target(value));
        }
    }
    if quoted {
        return Err(invalid_target(value));
    }

    let separator = separator.ok_or_else(|| invalid_target(value))?;
    let sheet = &value[..separator];
    let cell = &value[separator + 1..];
    if !valid_sheet_name(sheet) || !valid_cell_reference(cell) {
        return Err(invalid_target(value));
    }
    Ok(())
}

fn valid_sheet_name(value: &str) -> bool {
    let qualified = value.starts_with('$');
    let value = value.strip_prefix('$').unwrap_or(value);
    if value.is_empty() {
        return !qualified;
    }
    if let Some(inner) = value
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
    {
        if inner.is_empty() {
            return false;
        }
        let mut chars = inner.chars().peekable();
        while let Some(character) = chars.next() {
            if character == '\'' && chars.next() != Some('\'') {
                return false;
            }
        }
        true
    } else {
        !value
            .chars()
            .any(|character| character == '.' || character == '\'' || character == ' ')
    }
}

fn valid_cell_reference(value: &str) -> bool {
    let value = value.strip_prefix('$').unwrap_or(value);
    let column_length = value
        .bytes()
        .take_while(|byte| byte.is_ascii_uppercase())
        .count();
    if column_length == 0 {
        return false;
    }
    let row = value[column_length..]
        .strip_prefix('$')
        .unwrap_or(&value[column_length..]);
    !row.is_empty() && row.bytes().all(|byte| byte.is_ascii_digit())
}

fn invalid_target(value: &str) -> Error {
    Error::InvalidFormat(format!(
        "invalid consolidation target cell address '{value}'"
    ))
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
    fn parses_all_attributes_and_namespace_aliases() {
        let xml = concat!(
            "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
            "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
            "<o:body><o:spreadsheet><t:consolidation t:function=\"vendor:median\" ",
            "t:source-cell-range-addresses=\"'Q1 Sales'.A1:B2 Sheet2.C3:D4\" ",
            "t:target-cell-address=\"Summary.A1\" t:use-labels=\"both\" ",
            "t:link-to-source-data=\"1\"></t:consolidation>",
            "</o:spreadsheet></o:body></o:document-content>"
        );
        let parsed = parse_consolidation(xml).unwrap().unwrap();
        assert_eq!(parsed.function, "vendor:median");
        assert_eq!(parsed.source_cell_range_addresses.len(), 2);
        assert_eq!(parsed.use_labels, Some(ConsolidationUseLabels::Both));
        assert_eq!(parsed.link_to_source_data, Some(true));
    }

    #[test]
    fn rejects_invalid_structure_and_values() {
        for fragment in [
            "<table:table><table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\"/></table:table>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\"/><table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\"/>",
            "<table:consolidation table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\"/>",
            "<table:consolidation table:function=\"sum\" table:target-cell-address=\"S.B1\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1:B2\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"B1\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.b1\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\" table:use-labels=\"sideways\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\" table:link-to-source-data=\"yes\"/>",
            "<table:consolidation table:function=\"sum\" table:source-cell-range-addresses=\"S.A1:A2\" table:target-cell-address=\"S.B1\"><table:x/></table:consolidation>",
        ] {
            assert!(
                parse_consolidation(&format!("{PREFIX}{fragment}{SUFFIX}")).is_err(),
                "accepted {fragment}"
            );
        }
    }

    #[test]
    fn writer_round_trips_and_escapes() {
        let mut consolidation = Consolidation::new(
            "vendor:&median",
            vec!["'Q1 & Q2'.A1:B9".to_string(), "Sheet2.C1:D9".to_string()],
            "Summary.A1",
        )
        .unwrap();
        consolidation.use_labels = Some(ConsolidationUseLabels::Column);
        consolidation.link_to_source_data = Some(false);
        let mut xml = String::new();
        write_consolidation(&mut xml, Some(&consolidation)).unwrap();
        assert!(xml.contains("vendor:&amp;median"));
        assert_eq!(
            parse_consolidation(&format!("{PREFIX}{xml}{SUFFIX}"))
                .unwrap()
                .unwrap(),
            consolidation
        );
    }

    #[test]
    fn round_trips_through_builder_and_mutable_packages() {
        let original = Consolidation::new(
            "sum",
            vec!["Sheet1.A1:B2".to_string(), "Sheet1.D1:E2".to_string()],
            "Sheet1.G1",
        )
        .unwrap();
        let replacement =
            Consolidation::new("average", vec!["Sheet1.A1:E2".to_string()], "Sheet1.G3").unwrap();

        let mut builder = Builder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.set_consolidation(Some(original.clone())).unwrap();
        let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let xml = spreadsheet.content_xml();
        assert!(
            xml.find("</table:table>").unwrap() < xml.find("<table:consolidation ").unwrap(),
            "consolidation must follow spreadsheet tables"
        );
        assert_eq!(spreadsheet.consolidation(), Some(&original));

        let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        mutable
            .set_consolidation(Some(replacement.clone()))
            .unwrap();
        let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.consolidation(), Some(&replacement));
    }
}
