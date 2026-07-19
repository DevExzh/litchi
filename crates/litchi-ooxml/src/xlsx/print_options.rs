//! Immutable worksheet print-option metadata.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;

/// The effective flags from one worksheet `printOptions` element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct WorksheetPrintOptions {
    horizontal_centered: bool,
    vertical_centered: bool,
    print_headings: bool,
    grid_lines: bool,
    grid_lines_set: bool,
}

impl WorksheetPrintOptions {
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
pub fn parse_worksheet_print_options(xml: &[u8]) -> Result<Option<WorksheetPrintOptions>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let validated =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    parse_selected(selected)
}

fn parse_selected(xml: &[u8]) -> Result<Option<WorksheetPrintOptions>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, WorksheetPrintOptions)> = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
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
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("printOptions must not contain text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("printOptions must not contain CDATA"));
            },
            Event::End(element) => {
                if open
                    .as_ref()
                    .is_some_and(|(element_depth, _)| *element_depth == depth)
                {
                    let (_, options) = open.take().expect("checked above");
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
                if reference.resolve_char_ref().map_err(xml_error)?.is_none()
                    && !matches!(
                        reference.decode().map_err(xml_error)?.as_ref(),
                        "amp" | "lt" | "gt" | "apos" | "quot"
                    )
                {
                    return Err(invalid("custom XML entities are rejected"));
                }
                if open.is_some() {
                    return Err(invalid("printOptions must not contain entity text"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) | Event::CData(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || open.is_some() {
        return Err(invalid("incomplete worksheet print-options XML"));
    }
    Ok(result)
}

fn parse_options(element: &BytesStart<'_>, decoder: Decoder) -> Result<WorksheetPrintOptions> {
    let mut options = WorksheetPrintOptions::default();
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
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const START: &str =
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

    fn parse(body: &str) -> Result<Option<WorksheetPrintOptions>> {
        parse_worksheet_print_options(format!("{START}{body}</worksheet>").as_bytes())
    }

    fn parse_fixture(path: &str) -> WorksheetPrintOptions {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_worksheet_print_options(package.get_part(&uri).unwrap().blob())
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
            WorksheetPrintOptions::default()
        );
    }

    #[test]
    fn accepts_strict_namespace_and_absence() {
        let xml = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><printOptions gridLines="1" gridLinesSet="true"/></worksheet>"#;
        assert!(
            parse_worksheet_print_options(xml)
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
    fn loads_poi_centering_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/poi/test-data/spreadsheet/45540_classic_Header.xlsx"
        );
        let options = parse_fixture(path);
        assert!(options.horizontal_centered());
        assert!(!options.vertical_centered());
    }

    #[test]
    fn loads_libreoffice_gridline_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf100034.xlsx"
        );
        let options = parse_fixture(path);
        assert!(options.grid_lines());
        assert!(options.grid_lines_set());
        assert!(options.prints_grid_lines());
        assert!(!options.print_headings());
    }
}
