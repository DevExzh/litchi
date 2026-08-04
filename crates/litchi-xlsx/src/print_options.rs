//! Immutable worksheet print-option metadata.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;

/// The effective flags from one worksheet `printOptions` element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PrintOptions {
    horizontal_centered: bool,
    vertical_centered: bool,
    print_headings: bool,
    grid_lines: bool,
    grid_lines_set: bool,
}

impl PrintOptions {
    /// Center the printed content horizontally on the page.
    pub fn horizontal_centered(&self) -> bool {
        self.horizontal_centered
    }

    /// Center the printed content vertically on the page.
    pub fn vertical_centered(&self) -> bool {
        self.vertical_centered
    }

    /// Print row and column headings.
    pub fn print_headings(&self) -> bool {
        self.print_headings
    }

    /// Raw `gridLines` flag.
    pub fn grid_lines(&self) -> bool {
        self.grid_lines
    }

    /// Raw `gridLinesSet` flag.
    pub fn grid_lines_set(&self) -> bool {
        self.grid_lines_set
    }

    /// Whether gridlines are actually requested for printing.
    ///
    /// SpreadsheetML requires both gridline flags to be true.
    pub fn prints_grid_lines(&self) -> bool {
        self.grid_lines && self.grid_lines_set
    }
}

/// Parse a worksheet's optional core `printOptions` element.
pub fn parse_print_options(xml: &[u8]) -> Result<Option<PrintOptions>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..MceLimits::default()
    };
    let validated = process_markup_compatibility(xml, &MceCapabilities::default(), &limits)?;
    if validated.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed worksheet XML is too large"));
    }
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    parse_selected(selected)
}

fn parse_selected(xml: &[u8]) -> Result<Option<PrintOptions>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, PrintOptions)> = None;
    let mut declaration_seen = false;
    let mut events = 0usize;
    reader.config_mut().check_end_names = true;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("worksheet XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if depth == 1 {
                    if root_seen
                        || !spreadsheet(&namespace)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("print-options parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"printOptions"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet printOptions element"));
                    }
                    open = Some((depth, parse_options(&element, decoder)?));
                } else if open.is_some() {
                    return Err(invalid("printOptions must not contain child elements"));
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if !root_seen {
                    return Err(invalid("worksheet root cannot be empty"));
                }
                if depth == 0 {
                    return Err(invalid("worksheet XML element is outside root"));
                }
                if depth == 1
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"printOptions"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet printOptions element"));
                    }
                    result = Some(parse_options(&element, decoder)?);
                } else if open.is_some() {
                    return Err(invalid("printOptions must not contain child elements"));
                }
            },
            Event::Text(text) => {
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("printOptions must not contain text"));
                }
                if depth == 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("printOptions must not contain CDATA"));
            },
            Event::CData(_) if depth == 1 => {
                return Err(invalid("worksheet cannot contain direct CDATA"));
            },
            Event::CData(_) if !root_seen || root_closed => {
                return Err(invalid("worksheet XML CDATA is outside root"));
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected XML end element"));
                }
                if let Some((element_depth, _)) = open.as_ref() {
                    if *element_depth != depth {
                        return Err(invalid("invalid printOptions closing state"));
                    }
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"printOptions"
                    {
                        return Err(invalid("invalid printOptions closing element"));
                    }
                    let (_, options) = open
                        .take()
                        .ok_or_else(|| invalid("invalid printOptions closing state"))?;
                    result = Some(options);
                }
                if depth == 1 {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("invalid worksheet closing element"));
                    }
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
            },
            Event::GeneralRef(reference) => {
                if !root_seen || root_closed {
                    return Err(invalid("worksheet XML entity is outside root"));
                }
                if reference.resolve_char_ref().map_err(xml_error)?.is_none()
                    && !matches!(
                        reference.decode().map_err(xml_error)?.as_ref(),
                        "amp" | "lt" | "gt" | "apos" | "quot"
                    )
                {
                    return Err(invalid("custom XML entities are rejected"));
                }
                if open.is_some() || depth == 1 {
                    return Err(invalid("printOptions must not contain entity text"));
                }
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Comment(_) | Event::CData(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || open.is_some() {
        return Err(invalid("incomplete worksheet print-options XML"));
    }
    Ok(result)
}

fn parse_options(element: &BytesStart<'_>, decoder: Decoder) -> Result<PrintOptions> {
    let mut options = PrintOptions::default();
    let mut seen = [false; 5];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let (slot, target) = match attribute.key.local_name().as_ref() {
            b"horizontalCentered" => (0, &mut options.horizontal_centered),
            b"verticalCentered" => (1, &mut options.vertical_centered),
            b"headings" => (2, &mut options.print_headings),
            b"gridLines" => (3, &mut options.grid_lines),
            b"gridLinesSet" => (4, &mut options.grid_lines_set),
            name => {
                return Err(invalid(format!(
                    "unknown printOptions attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if seen[slot] {
            return Err(invalid("duplicate printOptions attribute"));
        }
        seen[slot] = true;
        *target = parse_bool(&value)?;
    }
    Ok(options)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid printOptions boolean")),
    }
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    exact(namespace, CORE) || exact(namespace, STRICT)
}
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet print-options XML: {error}"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const START: &str =
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;
    const CORE_STR: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(body: &str) -> Result<Option<PrintOptions>> {
        parse_print_options(format!("{START}{body}</worksheet>").as_bytes())
    }

    fn parse_fixture(path: &str) -> PrintOptions {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_print_options(package.get_part(&uri).unwrap().blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn parses_flags_defaults_and_derived_gridline_behavior() {
        let options = parse(r#"<printOptions horizontalCentered="1" verticalCentered="true" headings="1" gridLines="true" gridLinesSet="0"/>"#).unwrap().unwrap();
        assert!(options.horizontal_centered());
        assert!(options.vertical_centered());
        assert!(options.print_headings());
        assert!(options.grid_lines());
        assert!(!options.grid_lines_set());
        assert!(!options.prints_grid_lines());
        assert_eq!(
            parse("<printOptions/>").unwrap().unwrap(),
            PrintOptions::default()
        );
    }

    #[test]
    fn accepts_strict_namespace_and_absence() {
        let xml = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><printOptions gridLines="1" gridLinesSet="true"/></worksheet>"#;
        assert!(
            parse_print_options(xml)
                .unwrap()
                .unwrap()
                .prints_grid_lines()
        );
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_duplicate_unknown_and_nested_content() {
        assert!(parse(r#"<printOptions headings="yes"/>"#).is_err());
        assert!(parse(r#"<printOptions headings="1" headings="0"/>"#).is_err());
        assert!(parse(r#"<printOptions mystery="1"/>"#).is_err());
        assert!(parse(r#"<printOptions><x/></printOptions>"#).is_err());
    }

    #[test]
    fn rejects_malformed_closing_state_without_panicking() {
        let xml = format!(r#"{START}<printOptions/></worksheet><printOptions/>"#);
        let parsed = std::panic::catch_unwind(|| parse_print_options(xml.as_bytes()));
        assert!(matches!(parsed, Ok(Err(_))));

        let incomplete = format!(r#"{START}<printOptions>"#);
        assert!(parse_print_options(incomplete.as_bytes()).is_err());
    }

    #[test]
    fn rejects_malformed_document_boundaries_and_excessive_depth() {
        for xml in [
            format!(
                r#"<worksheet xmlns="{}"/><worksheet xmlns="{}"/>"#,
                CORE_STR, CORE_STR
            ),
            format!(r#"text{}"#, START),
            format!(r#"{}text</worksheet>"#, START),
            format!(r#"{}</worksheet>tail"#, START),
            format!(r#"{}<![CDATA[data]]></worksheet>"#, START),
        ] {
            assert!(
                parse_print_options(xml.as_bytes()).is_err(),
                "expected rejection for {xml}"
            );
        }

        let mut xml = START.to_owned();
        for _ in 0..MAX_DEPTH {
            xml.push_str("<extension>");
        }
        for _ in 0..MAX_DEPTH {
            xml.push_str("</extension>");
        }
        xml.push_str("</worksheet>");
        assert!(parse_print_options(xml.as_bytes()).is_err());
    }

    #[test]
    fn loads_poi_centering_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/45540_classic_Header.xlsx"
        );
        let options = parse_fixture(path);
        assert!(options.horizontal_centered());
        assert!(!options.vertical_centered());
    }

    #[test]
    fn loads_libreoffice_gridline_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf100034.xlsx"
        );
        let options = parse_fixture(path);
        assert!(options.grid_lines());
        assert!(options.grid_lines_set());
        assert!(options.prints_grid_lines());
        assert!(!options.print_headings());
    }
}
