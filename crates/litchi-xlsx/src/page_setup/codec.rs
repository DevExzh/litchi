//! Bounded SpreadsheetML worksheet page-setup XML codec.

use std::fmt::Display;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};

use super::model::{
    Comments, Copies, Dpi, ErrorMode, FirstPage, Fit, MAX_MEASURE_BYTES, Measure, Order,
    Orientation, Paper, RelId, Scale, Setup,
};

pub(super) const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
/// Parse a worksheet's optional typed core `pageSetup` element.
pub fn parse_worksheet_page_setup(xml: &[u8]) -> Result<Option<Setup>> {
    Ok(parse_worksheet_page_setup_parts(xml, Projection::Settings)?.map(|parsed| parsed.setup))
}

/// Parse only the relationship edge owned by the printer-settings layer.
pub fn parse_worksheet_page_setup_relationship_id(xml: &[u8]) -> Result<Option<RelId>> {
    Ok(
        parse_worksheet_page_setup_parts(xml, Projection::Relationship)?
            .and_then(|parsed| parsed.printer_settings),
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Projection {
    Settings,
    Relationship,
}

struct ParsedSetup {
    setup: Setup,
    printer_settings: Option<RelId>,
}

fn parse_worksheet_page_setup_parts(
    xml: &[u8],
    projection: Projection,
) -> Result<Option<ParsedSetup>> {
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
    parse_selected(selected, projection)
}

fn parse_selected(xml: &[u8], projection: Projection) -> Result<Option<ParsedSetup>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut worksheet_namespace = None;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, ParsedSetup)> = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver();
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
                    if worksheet_namespace.is_some()
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("page-setup parser requires a worksheet root"));
                    }
                    worksheet_namespace =
                        Some(spreadsheet_namespace(&namespace).ok_or_else(|| {
                            invalid("page-setup parser requires a worksheet root")
                        })?);
                } else if depth == 2 && element.local_name().as_ref() == b"pageSetup" {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if spreadsheet(&namespace) && !exact(&namespace, expected) {
                        return Err(invalid(
                            "pageSetup namespace does not match worksheet conformance",
                        ));
                    }
                    if exact(&namespace, expected) {
                        if result.is_some() || open.is_some() {
                            return Err(invalid("duplicate worksheet pageSetup element"));
                        }
                        open = Some((
                            depth,
                            parse_setup(
                                &element,
                                decoder,
                                resolver,
                                relationship_namespace(expected)?,
                                projection,
                            )?,
                        ));
                    }
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"pageSetup" {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if spreadsheet(&namespace) && !exact(&namespace, expected) {
                        return Err(invalid(
                            "pageSetup namespace does not match worksheet conformance",
                        ));
                    }
                    if exact(&namespace, expected) {
                        if result.is_some() || open.is_some() {
                            return Err(invalid("duplicate worksheet pageSetup element"));
                        }
                        result = Some(parse_setup(
                            &element,
                            decoder,
                            resolver,
                            relationship_namespace(expected)?,
                            projection,
                        )?);
                    }
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Text(text) => {
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("pageSetup must not contain text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("pageSetup must not contain CDATA"));
            },
            Event::End(element) => {
                let closes_page_setup = open
                    .as_ref()
                    .is_some_and(|(element_depth, _)| *element_depth == depth);
                if closes_page_setup && let Some((_, setup)) = open.take() {
                    result = Some(setup);
                }
                if depth == 1 {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if !exact(&namespace, expected) || element.local_name().as_ref() != b"worksheet"
                    {
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
                    return Err(invalid("pageSetup must not contain entity text"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) | Event::CData(_) => {},
            Event::Eof => break,
        }
    }
    if worksheet_namespace.is_none() || !root_closed || depth != 0 || open.is_some() {
        return Err(invalid("incomplete worksheet page-setup XML"));
    }
    Ok(result)
}

fn parse_setup(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    relationship_namespace: &[u8],
    projection: Projection,
) -> Result<ParsedSetup> {
    let mut setup = Setup::default();
    let mut printer_settings = None;
    let mut seen = [false; 18];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified_name = attribute.key.as_ref();
        if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
            continue;
        }
        let local_name = attribute.key.local_name();
        match resolver.resolve_attribute(attribute.key).0 {
            ResolveResult::Unbound => {},
            ResolveResult::Bound(Namespace(namespace))
                if namespace == relationship_namespace && local_name.as_ref() == b"id" =>
            {
                if printer_settings.is_some() {
                    return Err(invalid("duplicate pageSetup relationship ID"));
                }
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?;
                printer_settings = Some(
                    RelId::new(value.into_owned()).map_err(|error| invalid(error.to_string()))?,
                );
                continue;
            },
            ResolveResult::Bound(_) => {
                return Err(invalid(format!(
                    "unexpected qualified pageSetup attribute '{}'",
                    String::from_utf8_lossy(qualified_name)
                )));
            },
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "undeclared pageSetup attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if projection == Projection::Relationship {
            if is_setup_attribute(local_name.as_ref()) {
                continue;
            }
            return Err(invalid(format!(
                "unknown pageSetup attribute '{}'",
                String::from_utf8_lossy(local_name.as_ref())
            )));
        }
        let slot = match local_name.as_ref() {
            b"paperSize" => {
                setup.paper = Some(
                    Paper::try_from(parse_u32(&value, "paperSize")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                0
            },
            b"paperWidth" => {
                setup.paper_width = Some(parse_measure(&value, "paperWidth")?);
                1
            },
            b"paperHeight" => {
                setup.paper_height = Some(parse_measure(&value, "paperHeight")?);
                2
            },
            b"scale" => {
                setup.scale = Some(
                    Scale::try_from(parse_u32(&value, "scale")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                3
            },
            b"firstPageNumber" => {
                setup.first_page = Some(
                    FirstPage::from_wire(parse_u32(&value, "firstPageNumber")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                4
            },
            b"fitToWidth" => {
                setup.fit_to_width = Some(
                    Fit::try_from(parse_u32(&value, "fitToWidth")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                5
            },
            b"fitToHeight" => {
                setup.fit_to_height = Some(
                    Fit::try_from(parse_u32(&value, "fitToHeight")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                6
            },
            b"pageOrder" => {
                setup.order = Some(parse_order(&value)?);
                7
            },
            b"orientation" => {
                setup.orientation = Some(parse_orientation(&value)?);
                8
            },
            b"usePrinterDefaults" => {
                setup.use_printer_defaults = Some(parse_bool(&value, "usePrinterDefaults")?);
                9
            },
            b"blackAndWhite" => {
                setup.black_and_white = Some(parse_bool(&value, "blackAndWhite")?);
                10
            },
            b"draft" => {
                setup.draft = Some(parse_bool(&value, "draft")?);
                11
            },
            b"cellComments" => {
                setup.comments = Some(parse_comments(&value)?);
                12
            },
            b"useFirstPageNumber" => {
                setup.use_first_page = Some(parse_bool(&value, "useFirstPageNumber")?);
                13
            },
            b"errors" => {
                setup.errors = Some(parse_errors(&value)?);
                14
            },
            b"horizontalDpi" => {
                setup.horizontal_dpi = Some(
                    Dpi::new(parse_u32(&value, "horizontalDpi")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                15
            },
            b"verticalDpi" => {
                setup.vertical_dpi = Some(
                    Dpi::new(parse_u32(&value, "verticalDpi")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                16
            },
            b"copies" => {
                setup.copies = Some(
                    Copies::try_from(parse_u32(&value, "copies")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                17
            },
            name => {
                return Err(invalid(format!(
                    "unknown pageSetup attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if seen[slot] {
            return Err(invalid("duplicate pageSetup attribute"));
        }
        seen[slot] = true;
    }
    Ok(ParsedSetup {
        setup,
        printer_settings,
    })
}

fn is_setup_attribute(name: &[u8]) -> bool {
    matches!(
        name,
        b"paperSize"
            | b"paperWidth"
            | b"paperHeight"
            | b"scale"
            | b"firstPageNumber"
            | b"fitToWidth"
            | b"fitToHeight"
            | b"pageOrder"
            | b"orientation"
            | b"usePrinterDefaults"
            | b"blackAndWhite"
            | b"draft"
            | b"cellComments"
            | b"useFirstPageNumber"
            | b"errors"
            | b"horizontalDpi"
            | b"verticalDpi"
            | b"copies"
    )
}

fn parse_measure(raw: &str, field: &str) -> Result<Measure> {
    if raw.len() > MAX_MEASURE_BYTES {
        return Err(invalid(format!("{field} measure is too long")));
    }
    let boundary = raw
        .len()
        .checked_sub(2)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let number = raw
        .get(..boundary)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let suffix = raw
        .get(boundary..)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let unit = suffix
        .parse()
        .map_err(|_| invalid(format!("invalid {field} unit")))?;
    Measure::new(number, unit).map_err(|_| invalid(format!("invalid {field} measure")))
}

fn parse_u32(raw: &str, field: &str) -> Result<u32> {
    raw.parse()
        .map_err(|_| invalid(format!("invalid pageSetup {field}")))
}
fn parse_bool(raw: &str, field: &str) -> Result<bool> {
    match raw {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid pageSetup {field} boolean"))),
    }
}
fn parse_orientation(raw: &str) -> Result<Orientation> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup orientation"))
}
fn parse_order(raw: &str) -> Result<Order> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup pageOrder"))
}
fn parse_comments(raw: &str) -> Result<Comments> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup cellComments"))
}
fn parse_errors(raw: &str) -> Result<ErrorMode> {
    raw.parse().map_err(|_| invalid("invalid pageSetup errors"))
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    spreadsheet_namespace(namespace).is_some()
}
fn spreadsheet_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == CORE => Some(CORE),
        ResolveResult::Bound(Namespace(value)) if *value == STRICT => Some(STRICT),
        _ => None,
    }
}
fn relationship_namespace(namespace: &[u8]) -> Result<&'static [u8]> {
    match namespace {
        CORE => Ok(REL),
        STRICT => Ok(STRICT_REL),
        _ => Err(invalid(
            "pageSetup namespace does not select a relationship dialect",
        )),
    }
}
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}
pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn xml_error(error: impl Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
