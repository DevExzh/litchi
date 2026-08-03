//! Immutable worksheet page-margin metadata.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MAX_OFFICE_MARGIN_INCHES: f64 = 49.0;

/// A validated physical page margin stored by SpreadsheetML in inches.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct PageMargin(f64);

impl PageMargin {
    /// Margin in the native SpreadsheetML unit, inches.
    pub fn inches(self) -> f64 {
        self.0
    }

    /// Margin converted to typographic points.
    pub fn points(self) -> f64 {
        self.0 * 72.0
    }

    /// Margin converted to millimeters.
    pub fn millimeters(self) -> f64 {
        self.0 * 25.4
    }
}

/// The six required margins from one worksheet `pageMargins` element.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct WorksheetPageMargins {
    left: PageMargin,
    right: PageMargin,
    top: PageMargin,
    bottom: PageMargin,
    header: PageMargin,
    footer: PageMargin,
}

impl WorksheetPageMargins {
    pub fn left(&self) -> PageMargin {
        self.left
    }
    pub fn right(&self) -> PageMargin {
        self.right
    }
    pub fn top(&self) -> PageMargin {
        self.top
    }
    pub fn bottom(&self) -> PageMargin {
        self.bottom
    }
    pub fn header(&self) -> PageMargin {
        self.header
    }
    pub fn footer(&self) -> PageMargin {
        self.footer
    }
}

/// Parse a worksheet's optional core `pageMargins` element.
pub fn parse_worksheet_page_margins(xml: &[u8]) -> Result<Option<WorksheetPageMargins>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let validated =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
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

fn parse_selected(xml: &[u8]) -> Result<Option<WorksheetPageMargins>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, WorksheetPageMargins)> = None;
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
                if depth == 0 && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
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
                        return Err(invalid("page-margin parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"pageMargins"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet pageMargins element"));
                    }
                    open = Some((depth, parse_margins(&element, decoder)?));
                } else if open.is_some() {
                    return Err(invalid("pageMargins must not contain child elements"));
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if depth == 0 {
                    return Err(if root_seen {
                        invalid("worksheet XML contains multiple roots")
                    } else {
                        invalid("worksheet root cannot be empty")
                    });
                }
                if depth == 1
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"pageMargins"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet pageMargins element"));
                    }
                    result = Some(parse_margins(&element, decoder)?);
                } else if open.is_some() {
                    return Err(invalid("pageMargins must not contain child elements"));
                }
            },
            Event::Text(text) => {
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("pageMargins must not contain text"));
                }
                if depth == 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("pageMargins must not contain CDATA"));
            },
            Event::CData(_) if depth == 1 => {
                return Err(invalid("worksheet cannot contain direct CDATA"));
            },
            Event::CData(_) if !root_seen || root_closed => {
                return Err(invalid("worksheet XML CDATA is outside root"));
            },
            Event::End(element) => {
                if open
                    .as_ref()
                    .is_some_and(|(element_depth, _)| *element_depth == depth)
                {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"pageMargins" {
                        return Err(invalid("invalid pageMargins closing element"));
                    }
                    let (_, margins) = open
                        .take()
                        .ok_or_else(|| invalid("invalid pageMargins closing state"))?;
                    result = Some(margins);
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
                    return Err(invalid("pageMargins must not contain entity text"));
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
        return Err(invalid("incomplete worksheet page-margin XML"));
    }
    Ok(result)
}

fn parse_margins(element: &BytesStart<'_>, decoder: Decoder) -> Result<WorksheetPageMargins> {
    let mut values = [None; 6];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') {
            continue;
        }
        let slot = match attribute.key.local_name().as_ref() {
            b"left" => 0,
            b"right" => 1,
            b"top" => 2,
            b"bottom" => 3,
            b"header" => 4,
            b"footer" => 5,
            name => {
                return Err(invalid(format!(
                    "unknown pageMargins attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if values[slot].is_some() {
            return Err(invalid("duplicate pageMargins attribute"));
        }
        let raw = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        values[slot] = Some(parse_margin(&raw, attribute.key.local_name().as_ref())?);
    }
    let [left, right, top, bottom, header, footer] = values;
    Ok(WorksheetPageMargins {
        left: left.ok_or_else(|| missing("left"))?,
        right: right.ok_or_else(|| missing("right"))?,
        top: top.ok_or_else(|| missing("top"))?,
        bottom: bottom.ok_or_else(|| missing("bottom"))?,
        header: header.ok_or_else(|| missing("header"))?,
        footer: footer.ok_or_else(|| missing("footer"))?,
    })
}

fn parse_margin(raw: &str, name: &[u8]) -> Result<PageMargin> {
    let value = raw.parse::<f64>().map_err(|_| {
        invalid(format!(
            "invalid {} page margin",
            String::from_utf8_lossy(name)
        ))
    })?;
    if !value.is_finite() || !(0.0..MAX_OFFICE_MARGIN_INCHES).contains(&value) {
        return Err(invalid(format!(
            "{} page margin is outside Office's [0, 49) range",
            String::from_utf8_lossy(name)
        )));
    }
    Ok(PageMargin(value))
}

fn missing(name: &str) -> OoxmlError {
    invalid(format!(
        "pageMargins is missing required '{name}' attribute"
    ))
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
    const CORE_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(body: &str) -> Result<Option<WorksheetPageMargins>> {
        parse_worksheet_page_margins(format!("{START}{body}</worksheet>").as_bytes())
    }

    fn parse_fixture(path: &str) -> WorksheetPageMargins {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_worksheet_page_margins(package.get_part(&uri).unwrap().blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn parses_required_values_and_converts_units() {
        let margins = parse(
            r#"<pageMargins left="0.7" right="0" top="1" bottom="2" header="0.5" footer="0.25"/>"#,
        )
        .unwrap()
        .unwrap();
        assert_eq!(margins.left().inches(), 0.7);
        assert_eq!(margins.right().inches(), 0.0);
        assert_eq!(margins.top().points(), 72.0);
        assert_eq!(margins.bottom().millimeters(), 50.8);
    }

    #[test]
    fn accepts_strict_namespace_and_absence() {
        let xml = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><pageMargins left="1" right="2" top="3" bottom="4" header="5" footer="6"/></worksheet>"#;
        assert_eq!(
            parse_worksheet_page_margins(xml)
                .unwrap()
                .unwrap()
                .footer()
                .inches(),
            6.0
        );
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn rejects_missing_nonfinite_out_of_range_and_content() {
        assert!(
            parse(r#"<pageMargins left="1" right="2" top="3" bottom="4" header="5"/>"#).is_err()
        );
        assert!(
            parse(
                r#"<pageMargins left="NaN" right="2" top="3" bottom="4" header="5" footer="6"/>"#
            )
            .is_err()
        );
        assert!(
            parse(r#"<pageMargins left="49" right="2" top="3" bottom="4" header="5" footer="6"/>"#)
                .is_err()
        );
        assert!(parse(r#"<pageMargins left="1" right="2" top="3" bottom="4" header="5" footer="6"><x/></pageMargins>"#).is_err());
    }

    #[test]
    fn rejects_malformed_document_boundaries_and_excessive_depth() {
        for xml in [
            format!(
                r#"<worksheet xmlns="{CORE_NAMESPACE}"/><worksheet xmlns="{CORE_NAMESPACE}"/>"#
            ),
            format!(r#"text{}"#, START),
            format!(r#"{}text</worksheet>"#, START),
            format!(r#"{}</worksheet>tail"#, START),
            format!(r#"{}<![CDATA[data]]></worksheet>"#, START),
        ] {
            assert!(
                parse_worksheet_page_margins(xml.as_bytes()).is_err(),
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
        assert!(parse_worksheet_page_margins(xml.as_bytes()).is_err());
    }

    #[test]
    fn loads_poi_custom_margin_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/headerFooterTest.xlsx"
        );
        let margins = parse_fixture(path);
        assert!((margins.left().inches() - 0.75).abs() < 1e-12);
        assert_eq!(margins.top().inches(), 1.0);
        assert_eq!(margins.header().inches(), 0.5);
    }

    #[test]
    fn loads_libreoffice_metric_margin_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf134459_HeaderFooterColor.xlsx"
        );
        let margins = parse_fixture(path);
        assert!((margins.left().millimeters() - 20.0).abs() < 1e-10);
        assert!((margins.top().millimeters() - 27.0).abs() < 1e-10);
    }
}
