//! Bounded `SpreadsheetML` worksheet header/footer XML decoding.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::escape::unescape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};

use super::model::{Settings, Text};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_TEXT_BYTES: usize = 64 * 1024;

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

/// Parse a worksheet's core headerFooter settings.
pub fn parse_worksheet_header_footer(xml: &[u8]) -> Result<Option<Settings>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let validated =
        process_markup_compatibility(xml, &Capabilities::default(), &Limits::default())?;
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    parse_selected(selected)
}

fn parse_selected(xml: &[u8]) -> Result<Option<Settings>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut header_footer: Option<(usize, Settings, Option<u8>)> = None;
    let mut text: Option<(usize, TextKind, String)> = None;

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
                        return Err(invalid("header/footer parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"headerFooter"
                {
                    if result.is_some() || header_footer.is_some() {
                        return Err(invalid("duplicate worksheet headerFooter element"));
                    }
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
                    if text.is_some() {
                        return Err(invalid("nested header/footer text element"));
                    }
                    text = Some((depth, kind, String::new()));
                } else if text.is_some() {
                    return Err(invalid("header/footer text cannot contain child elements"));
                }
            },
            Event::Empty(element) => {
                if depth == 1
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"headerFooter"
                {
                    if result.is_some() || header_footer.is_some() {
                        return Err(invalid("duplicate worksheet headerFooter element"));
                    }
                    result = Some(parse_settings(&element, decoder)?);
                } else if let Some((container_depth, settings, last_order)) = header_footer.as_mut()
                {
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
            },
            Event::Text(value) => {
                if let Some((_, _, output)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    let decoded = unescape(&decoded).map_err(xml_error)?;
                    append_bounded(output, &decoded)?;
                } else if header_footer.is_some()
                    && !value.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("headerFooter cannot contain direct text"));
                }
            },
            Event::CData(value) => {
                if let Some((_, _, output)) = text.as_mut() {
                    append_bounded(output, &value.decode().map_err(xml_error)?)?;
                } else {
                    return Err(invalid("CDATA is outside header/footer text"));
                }
            },
            Event::GeneralRef(reference) => {
                let resolved =
                    if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
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
                if let Some((_, _, output)) = text.as_mut() {
                    append_bounded(output, &resolved)?;
                }
            },
            Event::End(element) => {
                if text
                    .as_ref()
                    .is_some_and(|(text_depth, _, _)| *text_depth == depth)
                {
                    let (_, kind, value) = text.take().expect("checked above");
                    let (_, settings, _) = header_footer
                        .as_mut()
                        .ok_or_else(|| invalid("orphan header/footer text"))?;
                    assign_text(settings, kind, value)?;
                } else if header_footer
                    .as_ref()
                    .is_some_and(|(container_depth, _, _)| *container_depth == depth)
                {
                    let (_, settings, _) = header_footer.take().expect("checked above");
                    result = Some(settings);
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
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || header_footer.is_some() || text.is_some() {
        return Err(invalid("incomplete worksheet header/footer XML"));
    }
    Ok(result)
}

fn parse_settings(element: &BytesStart<'_>, decoder: Decoder) -> Result<Settings> {
    let mut settings = Settings::default();
    let mut seen = [false; 4];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let (slot, target) = match attribute.key.local_name().as_ref() {
            b"differentOddEven" => (0, &mut settings.different_odd_even),
            b"differentFirst" => (1, &mut settings.different_first),
            b"scaleWithDoc" => (2, &mut settings.scale_with_document),
            b"alignWithMargins" => (3, &mut settings.align_with_margins),
            name => {
                return Err(invalid(format!(
                    "unknown headerFooter attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if seen[slot] {
            return Err(invalid("duplicate headerFooter attribute"));
        }
        seen[slot] = true;
        *target = parse_bool(&value)?;
    }
    Ok(settings)
}

fn validate_child_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !attribute.key.as_ref().contains(&b':') {
            return Err(invalid(
                "header/footer text elements cannot have attributes",
            ));
        }
    }
    Ok(())
}

fn check_order(last_order: Option<u8>, kind: TextKind) -> Result<()> {
    if last_order.is_some_and(|last| kind.order() <= last) {
        return Err(invalid(
            "headerFooter children are duplicated or out of schema order",
        ));
    }
    Ok(())
}

fn assign_text(settings: &mut Settings, kind: TextKind, value: String) -> Result<()> {
    let target = match kind {
        TextKind::OddHeader => &mut settings.odd_header,
        TextKind::OddFooter => &mut settings.odd_footer,
        TextKind::EvenHeader => &mut settings.even_header,
        TextKind::EvenFooter => &mut settings.even_footer,
        TextKind::FirstHeader => &mut settings.first_header,
        TextKind::FirstFooter => &mut settings.first_footer,
    };
    if target.replace(Text::new(value)).is_some() {
        return Err(invalid("duplicate header/footer text element"));
    }
    Ok(())
}

fn append_bounded(output: &mut String, value: &str) -> Result<()> {
    if output.len().saturating_add(value.len()) > MAX_TEXT_BYTES {
        return Err(invalid("header/footer text is too large"));
    }
    output.push_str(value);
    Ok(())
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid headerFooter boolean")),
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
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
