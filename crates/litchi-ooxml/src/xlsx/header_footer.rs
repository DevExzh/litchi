//! Immutable worksheet header/footer metadata.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::escape::unescape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_TEXT_BYTES: usize = 64 * 1024;

/// A logical left, center, or right header/footer section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterSectionKind {
    Left,
    Center,
    Right,
}

/// Header/footer text with its unambiguous alignment sections extracted.
///
/// Formatting and field control codes inside each section are intentionally
/// preserved. They can be localized, and OOXML requires their text to remain
/// available even when an application does not interpret the formatting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderFooterText {
    raw: String,
    left: Option<String>,
    center: Option<String>,
    right: Option<String>,
}

impl HeaderFooterText {
    fn new(raw: String) -> Self {
        let (left, center, right) = split_sections(&raw);
        Self { raw, left, center, right }
    }

    /// Complete decoded text, including alignment and formatting controls.
    pub fn raw(&self) -> &str { &self.raw }

    /// Content of a logical alignment section, excluding its alignment marker.
    pub fn section(&self, kind: HeaderFooterSectionKind) -> Option<&str> {
        match kind {
            HeaderFooterSectionKind::Left => self.left.as_deref(),
            HeaderFooterSectionKind::Center => self.center.as_deref(),
            HeaderFooterSectionKind::Right => self.right.as_deref(),
        }
    }

    pub fn left(&self) -> Option<&str> { self.left.as_deref() }
    pub fn center(&self) -> Option<&str> { self.center.as_deref() }
    pub fn right(&self) -> Option<&str> { self.right.as_deref() }
}

/// Complete core `headerFooter` settings for one worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetHeaderFooter {
    different_odd_even: bool,
    different_first: bool,
    scale_with_document: bool,
    align_with_margins: bool,
    odd_header: Option<HeaderFooterText>,
    odd_footer: Option<HeaderFooterText>,
    even_header: Option<HeaderFooterText>,
    even_footer: Option<HeaderFooterText>,
    first_header: Option<HeaderFooterText>,
    first_footer: Option<HeaderFooterText>,
}

impl WorksheetHeaderFooter {
    pub fn different_odd_even(&self) -> bool { self.different_odd_even }
    pub fn different_first(&self) -> bool { self.different_first }
    pub fn scale_with_document(&self) -> bool { self.scale_with_document }
    pub fn align_with_margins(&self) -> bool { self.align_with_margins }
    pub fn odd_header(&self) -> Option<&HeaderFooterText> { self.odd_header.as_ref() }
    pub fn odd_footer(&self) -> Option<&HeaderFooterText> { self.odd_footer.as_ref() }
    pub fn even_header(&self) -> Option<&HeaderFooterText> { self.even_header.as_ref() }
    pub fn even_footer(&self) -> Option<&HeaderFooterText> { self.even_footer.as_ref() }
    pub fn first_header(&self) -> Option<&HeaderFooterText> { self.first_header.as_ref() }
    pub fn first_footer(&self) -> Option<&HeaderFooterText> { self.first_footer.as_ref() }
}

impl Default for WorksheetHeaderFooter {
    fn default() -> Self {
        Self {
            different_odd_even: false,
            different_first: false,
            scale_with_document: true,
            align_with_margins: true,
            odd_header: None,
            odd_footer: None,
            even_header: None,
            even_footer: None,
            first_header: None,
            first_footer: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextKind {
    OddHeader,
    OddFooter,
    EvenHeader,
    EvenFooter,
    FirstHeader,
    FirstFooter,
}

impl TextKind {
    fn from_local_name(name: &[u8]) -> Option<Self> {
        match name {
            b"oddHeader" => Some(Self::OddHeader),
            b"oddFooter" => Some(Self::OddFooter),
            b"evenHeader" => Some(Self::EvenHeader),
            b"evenFooter" => Some(Self::EvenFooter),
            b"firstHeader" => Some(Self::FirstHeader),
            b"firstFooter" => Some(Self::FirstFooter),
            _ => None,
        }
    }

    fn order(self) -> u8 {
        match self {
            Self::OddHeader => 0,
            Self::OddFooter => 1,
            Self::EvenHeader => 2,
            Self::EvenFooter => 3,
            Self::FirstHeader => 4,
            Self::FirstFooter => 5,
        }
    }
}

/// Parse a worksheet's core header/footer settings.
pub fn parse_worksheet_header_footer(xml: &[u8]) -> Result<Option<WorksheetHeaderFooter>> {
    if xml.len() > MAX_XML_BYTES { return Err(invalid("worksheet XML is too large")); }
    let validated = process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let selected = if validated.report.alternate_content_count == 0 { xml } else { validated.xml.as_ref() };
    parse_selected(selected)
}

fn parse_selected(xml: &[u8]) -> Result<Option<WorksheetHeaderFooter>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut header_footer: Option<(usize, WorksheetHeaderFooter, Option<u8>)> = None;
    let mut text: Option<(usize, TextKind, String)> = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH { return Err(invalid("worksheet XML nesting is too deep")); }
                if depth == 1 {
                    if root_seen || !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("header/footer parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if depth == 2 && spreadsheet(&namespace) && element.local_name().as_ref() == b"headerFooter" {
                    if result.is_some() || header_footer.is_some() { return Err(invalid("duplicate worksheet headerFooter element")); }
                    header_footer = Some((depth, parse_settings(&element, decoder)?, None));
                } else if let Some((container_depth, _, last_order)) = header_footer.as_mut() {
                    if depth != *container_depth + 1 || !spreadsheet(&namespace) {
                        return Err(invalid("unexpected nested content in headerFooter"));
                    }
                    let kind = TextKind::from_local_name(element.local_name().as_ref())
                        .ok_or_else(|| invalid("unknown headerFooter child element"))?;
                    validate_child_attributes(&element)?;
                    check_order(*last_order, kind)?;
                    *last_order = Some(kind.order());
                    if text.is_some() { return Err(invalid("nested header/footer text element")); }
                    text = Some((depth, kind, String::new()));
                } else if text.is_some() {
                    return Err(invalid("header/footer text cannot contain child elements"));
                }
            }
            Event::Empty(element) => {
                if depth == 1 && spreadsheet(&namespace) && element.local_name().as_ref() == b"headerFooter" {
                    if result.is_some() || header_footer.is_some() { return Err(invalid("duplicate worksheet headerFooter element")); }
                    result = Some(parse_settings(&element, decoder)?);
                } else if let Some((container_depth, settings, last_order)) = header_footer.as_mut() {
                    if depth + 1 != *container_depth + 1 || !spreadsheet(&namespace) {
                        return Err(invalid("unexpected empty content in headerFooter"));
                    }
                    let kind = TextKind::from_local_name(element.local_name().as_ref())
                        .ok_or_else(|| invalid("unknown headerFooter child element"))?;
                    validate_child_attributes(&element)?;
                    check_order(*last_order, kind)?;
                    *last_order = Some(kind.order());
                    assign_text(settings, kind, String::new())?;
                }
            }
            Event::Text(value) => {
                if let Some((_, _, output)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    let decoded = unescape(&decoded).map_err(xml_error)?;
                    append_bounded(output, &decoded)?;
                } else if header_footer.is_some() && !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("headerFooter cannot contain direct text"));
                }
            }
            Event::CData(value) => {
                if let Some((_, _, output)) = text.as_mut() {
                    append_bounded(output, &value.decode().map_err(xml_error)?)?;
                } else {
                    return Err(invalid("CDATA is outside header/footer text"));
                }
            }
            Event::GeneralRef(reference) => {
                let resolved = if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
                    character.to_string()
                } else {
                    match reference.decode().map_err(xml_error)?.as_ref() {
                        "amp" => "&".to_string(),
                        "lt" => "<".to_string(),
                        "gt" => ">".to_string(),
                        "apos" => "'".to_string(),
                        "quot" => "\"".to_string(),
                        _ => return Err(invalid("custom XML entities are rejected")),
                    }
                };
                if let Some((_, _, output)) = text.as_mut() { append_bounded(output, &resolved)?; }
            }
            Event::End(element) => {
                if text.as_ref().is_some_and(|(text_depth, _, _)| *text_depth == depth) {
                    let (_, kind, value) = text.take().expect("checked above");
                    let (_, settings, _) = header_footer.as_mut().ok_or_else(|| invalid("orphan header/footer text"))?;
                    assign_text(settings, kind, value)?;
                } else if header_footer.as_ref().is_some_and(|(container_depth, _, _)| *container_depth == depth) {
                    let (_, settings, _) = header_footer.take().expect("checked above");
                    result = Some(settings);
                }
                if depth == 1 {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" { return Err(invalid("invalid worksheet closing element")); }
                    root_closed = true;
                }
                depth = depth.checked_sub(1).ok_or_else(|| invalid("unexpected XML end element"))?;
            }
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTD and processing instructions are rejected")),
            Event::Decl(_) | Event::Comment(_) => {}
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || header_footer.is_some() || text.is_some() {
        return Err(invalid("incomplete worksheet header/footer XML"));
    }
    Ok(result)
}

fn parse_settings(element: &BytesStart<'_>, decoder: Decoder) -> Result<WorksheetHeaderFooter> {
    let mut settings = WorksheetHeaderFooter::default();
    let mut seen = [false; 4];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') { continue; }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let (slot, target) = match attribute.key.local_name().as_ref() {
            b"differentOddEven" => (0, &mut settings.different_odd_even),
            b"differentFirst" => (1, &mut settings.different_first),
            b"scaleWithDoc" => (2, &mut settings.scale_with_document),
            b"alignWithMargins" => (3, &mut settings.align_with_margins),
            name => return Err(invalid(format!("unknown headerFooter attribute '{}'", String::from_utf8_lossy(name)))),
        };
        if seen[slot] { return Err(invalid("duplicate headerFooter attribute")); }
        seen[slot] = true;
        *target = parse_bool(&value)?;
    }
    Ok(settings)
}

fn validate_child_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !attribute.key.as_ref().contains(&b':') { return Err(invalid("header/footer text elements cannot have attributes")); }
    }
    Ok(())
}

fn check_order(last_order: Option<u8>, kind: TextKind) -> Result<()> {
    if last_order.is_some_and(|last| kind.order() <= last) {
        return Err(invalid("headerFooter children are duplicated or out of schema order"));
    }
    Ok(())
}

fn assign_text(settings: &mut WorksheetHeaderFooter, kind: TextKind, value: String) -> Result<()> {
    let target = match kind {
        TextKind::OddHeader => &mut settings.odd_header,
        TextKind::OddFooter => &mut settings.odd_footer,
        TextKind::EvenHeader => &mut settings.even_header,
        TextKind::EvenFooter => &mut settings.even_footer,
        TextKind::FirstHeader => &mut settings.first_header,
        TextKind::FirstFooter => &mut settings.first_footer,
    };
    if target.replace(HeaderFooterText::new(value)).is_some() { return Err(invalid("duplicate header/footer text element")); }
    Ok(())
}

fn split_sections(raw: &str) -> (Option<String>, Option<String>, Option<String>) {
    let mut sections = [None::<String>, None::<String>, None::<String>];
    let mut current = 1usize;
    let mut index = 0usize;
    while index < raw.len() {
        let tail = &raw[index..];
        if let Some(marker) = tail.as_bytes().get(1).copied().filter(|_| tail.as_bytes()[0] == b'&') {
            if marker == b'&' {
                sections[current].get_or_insert_with(String::new).push_str("&&");
                index += 2;
                continue;
            }
            if let Some(next) = match marker { b'L' => Some(0), b'C' => Some(1), b'R' => Some(2), _ => None } {
                current = next;
                sections[current].get_or_insert_with(String::new);
                index += 2;
                continue;
            }
        }
        let character = tail.chars().next().expect("non-empty tail");
        sections[current].get_or_insert_with(String::new).push(character);
        index += character.len_utf8();
    }
    if raw.is_empty() { sections[1] = Some(String::new()); }
    (sections[0].take(), sections[1].take(), sections[2].take())
}

fn append_bounded(output: &mut String, value: &str) -> Result<()> {
    if output.len().saturating_add(value.len()) > MAX_TEXT_BYTES { return Err(invalid("header/footer text is too large")); }
    output.push_str(value);
    Ok(())
}

fn parse_bool(value: &str) -> Result<bool> {
    match value { "1" | "true" => Ok(true), "0" | "false" => Ok(false), _ => Err(invalid("invalid headerFooter boolean")) }
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool { exact(namespace, CORE) || exact(namespace, STRICT) }
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool { matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected) }
fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
fn xml_error(error: impl std::fmt::Display) -> OoxmlError { OoxmlError::Xml(error.to_string()) }

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const START: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

    fn parse(body: &str) -> Result<Option<WorksheetHeaderFooter>> {
        parse_worksheet_header_footer(format!("{START}{body}</worksheet>").as_bytes())
    }

    fn parse_fixture(path: &str) -> WorksheetHeaderFooter {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_worksheet_header_footer(package.get_part(&uri).unwrap().blob()).unwrap().unwrap()
    }

    #[test]
    fn parses_all_variants_defaults_and_sections() {
        let settings = parse(r#"<headerFooter differentOddEven="1" differentFirst="true" scaleWithDoc="0" alignWithMargins="false"><oddHeader>&amp;Lleft&amp;Ccenter&amp;Rright</oddHeader><oddFooter>&amp;P</oddFooter><evenHeader>even</evenHeader><evenFooter/><firstHeader>first</firstHeader><firstFooter>last</firstFooter></headerFooter>"#).unwrap().unwrap();
        assert!(settings.different_odd_even());
        assert!(settings.different_first());
        assert!(!settings.scale_with_document());
        assert!(!settings.align_with_margins());
        let header = settings.odd_header().unwrap();
        assert_eq!(header.raw(), "&Lleft&Ccenter&Rright");
        assert_eq!(header.left(), Some("left"));
        assert_eq!(header.center(), Some("center"));
        assert_eq!(header.right(), Some("right"));
        assert_eq!(settings.even_footer().unwrap().center(), Some(""));
    }

    #[test]
    fn preserves_ampersands_and_unrecognized_formatting() {
        let settings = parse(r#"<headerFooter><oddHeader>&amp;Cone &amp;&amp; two &amp;&amp;&amp;&amp;&amp;K01+000</oddHeader></headerFooter>"#).unwrap().unwrap();
        let header = settings.odd_header().unwrap();
        assert_eq!(header.center(), Some("one && two &&&&&K01+000"));
    }

    #[test]
    fn rejects_duplicates_order_and_nested_markup() {
        assert!(parse("<headerFooter><oddFooter/><oddHeader/></headerFooter>").is_err());
        assert!(parse("<headerFooter><oddHeader/><oddHeader/></headerFooter>").is_err());
        assert!(parse("<headerFooter><oddHeader><b/></oddHeader></headerFooter>").is_err());
        assert!(parse("<headerFooter scaleWithDoc=\"yes\"/>").is_err());
    }

    #[test]
    fn loads_poi_ampersand_fixture() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../3rdparty/poi/test-data/spreadsheet/AmpersandHeader.xlsx");
        let settings = parse_fixture(path);
        assert_eq!(settings.odd_header().unwrap().center(), Some("one && two &&&&"));
    }

    #[test]
    fn loads_libreoffice_color_sections_fixture() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf134459_HeaderFooterColor.xlsx");
        let settings = parse_fixture(path);
        let header = settings.odd_header().unwrap();
        assert_eq!(header.left(), Some("&KC06040l"));
        assert_eq!(header.center(), Some("&K4C3789c"));
        assert_eq!(header.right(), Some("r"));
    }
}
