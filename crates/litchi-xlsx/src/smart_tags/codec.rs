//! Bounded `SpreadsheetML` smart-tag XML conversion.

use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_sheet::Cell as Address;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Cell, Collection, Conformance, Property, Tag};
use super::{MAX_DEPTH, MAX_XML_BYTES, STRICT, TRANSITIONAL};
use crate::error::{Error, Result, allocation, invalid};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    Collection,
    Cell,
    Tag,
    Property,
    Other,
}

struct Parser {
    scopes: Vec<Scope>,
    cells: Vec<Cell>,
    seen: bool,
    root_seen: bool,
    root_closed: bool,
}

/// Parse the direct worksheet `smartTags` collection after bounded MCE
/// processing.
pub fn parse(xml: &[u8]) -> Result<Option<Collection>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("smart-tag worksheet XML exceeds the size limit"));
    }
    let processed =
        process_markup_compatibility(xml, &Capabilities::default(), &Limits::default())?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(invalid(
            "processed smart-tag worksheet XML exceeds the size limit",
        ));
    }

    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser {
        scopes: Vec::new(),
        cells: Vec::new(),
        seen: false,
        root_seen: false,
        root_closed: false,
    };
    let mut declaration_seen = false;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                parser.start(&namespace, &element, decoder, &resolver)?;
            },
            Event::Empty(element) => {
                parser.empty(&namespace, &element, decoder, &resolver)?;
            },
            Event::End(element) => parser.end(element.local_name().as_ref())?,
            Event::Text(text) => {
                if parser.modeled() && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("unexpected text inside worksheet smart tags"));
                }
                if parser.scopes.is_empty() && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid("smart-tag XML text is outside the root"));
                }
            },
            Event::CData(_) if parser.modeled() || parser.scopes.is_empty() => {
                return Err(invalid("unexpected CDATA in smart-tag XML"));
            },
            Event::Decl(_) => {
                if declaration_seen || parser.root_seen {
                    return Err(invalid("invalid smart-tag XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !parser.scopes.is_empty() || !parser.root_seen || !parser.root_closed {
        return Err(invalid("unterminated smart-tag worksheet XML"));
    }
    if !parser.seen {
        return Ok(None);
    }
    Collection::new(parser.cells).map(Some)
}

impl Parser {
    fn parent(&self) -> Option<Scope> {
        self.scopes.last().copied()
    }

    fn modeled(&self) -> bool {
        matches!(
            self.parent(),
            Some(Scope::Worksheet | Scope::Collection | Scope::Cell | Scope::Tag | Scope::Property)
        )
    }

    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.scopes.len() >= MAX_DEPTH {
            return Err(invalid("smart-tag XML nesting is too deep"));
        }
        let scope = self.begin(namespace, element, decoder, resolver, false)?;
        self.scopes.push(scope);
        Ok(())
    }

    fn empty(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let scope = self.begin(namespace, element, decoder, resolver, true)?;
        match scope {
            Scope::Worksheet => self.root_closed = true,
            Scope::Collection => {
                return Err(invalid("smartTags requires at least one cellSmartTags"));
            },
            Scope::Cell => {
                return Err(invalid("cellSmartTags requires at least one cellSmartTag"));
            },
            Scope::Tag | Scope::Property | Scope::Other => {},
        }
        Ok(())
    }

    fn begin(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        empty: bool,
    ) -> Result<Scope> {
        let local = element.local_name();
        let core = is_main(namespace);
        if self.scopes.is_empty() {
            if self.root_seen || self.root_closed {
                return Err(invalid("smart-tag XML contains multiple roots"));
            }
            if !core || local.as_ref() != b"worksheet" {
                return Err(invalid("smart-tag parser requires a worksheet root"));
            }
            self.root_seen = true;
            if empty {
                return Ok(Scope::Worksheet);
            }
        }

        match (self.parent(), core, local.as_ref()) {
            (None, true, b"worksheet") => Ok(Scope::Worksheet),
            (Some(Scope::Worksheet), true, b"smartTags") => {
                if self.seen {
                    return Err(invalid("worksheet has multiple smartTags collections"));
                }
                reject_attributes(element, "smartTags")?;
                self.seen = true;
                Ok(Scope::Collection)
            },
            (Some(Scope::Collection), true, b"cellSmartTags") => {
                let address = parse_cell(element, decoder, resolver)?;
                self.cells.push(Cell {
                    address,
                    tags: Vec::new(),
                });
                Ok(Scope::Cell)
            },
            (Some(Scope::Cell), true, b"cellSmartTag") => {
                let tag = parse_tag(element, decoder, resolver)?;
                self.cells
                    .last_mut()
                    .ok_or_else(|| invalid("cellSmartTag is outside cellSmartTags"))?
                    .tags
                    .push(tag);
                Ok(Scope::Tag)
            },
            (Some(Scope::Tag), true, b"cellSmartTagPr") => {
                let property = parse_property(element, decoder, resolver)?;
                self.cells
                    .last_mut()
                    .and_then(|cell| cell.tags.last_mut())
                    .ok_or_else(|| invalid("cellSmartTagPr is outside cellSmartTag"))?
                    .properties
                    .push(property);
                Ok(Scope::Property)
            },
            (Some(Scope::Property), _, _) => {
                Err(invalid("cellSmartTagPr cannot contain child elements"))
            },
            (Some(Scope::Collection | Scope::Cell | Scope::Tag), _, _) => Err(invalid(format!(
                "unexpected smart-tag element '{}'",
                String::from_utf8_lossy(local.as_ref())
            ))),
            _ => Ok(Scope::Other),
        }
    }

    fn end(&mut self, local: &[u8]) -> Result<()> {
        let scope = self
            .scopes
            .pop()
            .ok_or_else(|| invalid("unexpected smart-tag end element"))?;
        match scope {
            Scope::Worksheet if local == b"worksheet" => self.root_closed = true,
            Scope::Collection if local == b"smartTags" => {},
            Scope::Cell if local == b"cellSmartTags" => {},
            Scope::Tag if local == b"cellSmartTag" => {},
            Scope::Property if local == b"cellSmartTagPr" => {},
            Scope::Other => {},
            Scope::Worksheet | Scope::Collection | Scope::Cell | Scope::Tag | Scope::Property => {
                return Err(invalid("mismatched smart-tag end element"));
            },
        }
        Ok(())
    }
}

/// Serialize one canonical, namespace-complete `smartTags` fragment.
pub fn write(value: &Collection, conformance: Conformance) -> Result<Vec<u8>> {
    super::validation::collection(value)?;
    if value.is_empty() {
        return Err(invalid("an empty smart-tag collection has no XML form"));
    }
    let mut output = String::new();
    output
        .try_reserve(value.len().saturating_mul(128).saturating_add(128))
        .map_err(|source| allocation("smart-tag XML", source))?;
    output.push_str("<smartTags xmlns=\"");
    output.push_str(conformance.namespace());
    output.push_str("\">");
    for cell in value.cells() {
        output.push_str("<cellSmartTags r=\"");
        output.push_str(&cell.address().a1());
        output.push_str("\">");
        for tag in cell.tags() {
            output.push_str("<cellSmartTag type=\"");
            output.push_str(&tag.type_id().to_string());
            output.push('"');
            if tag.is_deleted() {
                output.push_str(" deleted=\"1\"");
            }
            if tag.is_xml_based() {
                output.push_str(" xmlBased=\"1\"");
            }
            if tag.properties().is_empty() {
                output.push_str("/>");
                continue;
            }
            output.push('>');
            for property in tag.properties() {
                output.push_str("<cellSmartTagPr key=\"");
                escape_attribute(&mut output, property.key());
                output.push_str("\" val=\"");
                escape_attribute(&mut output, property.value());
                output.push_str("\"/>");
            }
            output.push_str("</cellSmartTag>");
        }
        output.push_str("</cellSmartTags>");
    }
    output.push_str("</smartTags>");
    if output.len() > MAX_XML_BYTES {
        return Err(invalid("smart-tag output exceeds the size limit"));
    }
    Ok(output.into_bytes())
}

/// Replace or remove the direct worksheet `smartTags` child while preserving
/// every unrelated source byte.
pub fn replace_worksheet(xml: &[u8], value: Option<&Collection>) -> Result<Vec<u8>> {
    let selected = parse(xml)?;
    let location = locate(xml)?;
    if selected.is_some() && location.span.is_none() {
        return Err(invalid(
            "MCE-selected smart tags cannot be mutated as a direct worksheet child",
        ));
    }
    let replacement = match value {
        Some(value) if !value.is_empty() => write(value, location.conformance)?,
        Some(_) | None => Vec::new(),
    };
    let (start, end) = location
        .span
        .unwrap_or((location.insertion, location.insertion));
    if start == end && replacement.is_empty() {
        return Ok(xml.to_vec());
    }
    let capacity = xml
        .len()
        .checked_sub(end - start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("smart-tag worksheet output size overflow"))?;
    if capacity > MAX_XML_BYTES {
        return Err(invalid("smart-tag worksheet output exceeds the size limit"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("smart-tag worksheet XML", source))?;
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

struct Location {
    conformance: Conformance,
    span: Option<(usize, usize)>,
    insertion: usize,
}

fn locate(xml: &[u8]) -> Result<Location> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("smart-tag worksheet XML exceeds the size limit"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut conformance = None;
    let mut start = None;
    let mut span = None;
    let mut insertion = None;
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("smart-tag XML offset overflow"))?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("smart-tag XML offset overflow"))?;
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("smart-tag parser requires one worksheet root"));
                    }
                    conformance = Some(conformance_of(&namespace)?);
                    root_seen = true;
                } else if depth == 1 && is_main(&namespace) {
                    let local = element.local_name();
                    if local.as_ref() == b"smartTags" {
                        if start.replace(event_start).is_some() || span.is_some() {
                            return Err(invalid("worksheet has multiple direct smartTags"));
                        }
                    } else if insertion.is_none() && follows_smart_tags(local.as_ref()) {
                        insertion = Some(event_start);
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("smart-tag XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("smart-tag XML nesting is too deep"));
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(invalid("an empty worksheet root cannot own smart tags"));
                } else if depth == 1 && is_main(&namespace) {
                    let local = element.local_name();
                    if local.as_ref() == b"smartTags" {
                        if start.is_some() || span.replace((event_start, event_end)).is_some() {
                            return Err(invalid("worksheet has multiple direct smartTags"));
                        }
                    } else if insertion.is_none() && follows_smart_tags(local.as_ref()) {
                        insertion = Some(event_start);
                    }
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("smart-tag XML depth underflow"))?;
                if depth == 1 && element.local_name().as_ref() == b"smartTags" {
                    let element_start = start
                        .take()
                        .ok_or_else(|| invalid("mismatched smartTags closing element"))?;
                    span = Some((element_start, event_end));
                }
                if depth == 0 {
                    if element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("mismatched worksheet closing element"));
                    }
                    insertion.get_or_insert(event_start);
                    root_closed = true;
                }
            },
            Event::Text(text)
                if depth == 0 && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("smart-tag XML text is outside the root"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 || !root_seen || !root_closed || start.is_some() {
        return Err(invalid("unterminated smart-tag worksheet XML"));
    }
    Ok(Location {
        conformance: conformance.ok_or_else(|| invalid("missing worksheet conformance"))?,
        span,
        insertion: insertion.ok_or_else(|| invalid("missing worksheet insertion point"))?,
    })
}

fn follows_smart_tags(name: &[u8]) -> bool {
    matches!(
        name,
        b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

fn parse_cell(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Address> {
    let mut reference = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) || local.as_ref() != b"r" {
            return Err(invalid(format!(
                "unexpected cellSmartTags attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        if reference.is_some() {
            return Err(invalid("duplicate cellSmartTags r attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        reference = Some(Address::from_a1(&value)?);
    }
    reference.ok_or_else(|| invalid("cellSmartTags requires r"))
}

fn parse_tag(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Tag> {
    let mut type_id = None;
    let mut deleted = None;
    let mut xml_based = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(
                "namespaced cellSmartTag attributes are unsupported",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"type" if type_id.is_none() => {
                type_id =
                    Some(value.parse::<u32>().map_err(|_source| {
                        invalid("cellSmartTag type is not an unsigned integer")
                    })?);
            },
            b"deleted" if deleted.is_none() => deleted = Some(parse_bool(&value, "deleted")?),
            b"xmlBased" if xml_based.is_none() => xml_based = Some(parse_bool(&value, "xmlBased")?),
            b"type" | b"deleted" | b"xmlBased" => {
                return Err(invalid("duplicate cellSmartTag attribute"));
            },
            _ => {
                return Err(invalid(format!(
                    "unexpected cellSmartTag attribute '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
        }
    }
    let mut tag = Tag::new(type_id.ok_or_else(|| invalid("cellSmartTag requires type"))?)?;
    tag.set_deleted(deleted.unwrap_or(false));
    tag.set_xml_based(xml_based.unwrap_or(false));
    Ok(tag)
}

fn parse_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Property> {
    let mut key = None;
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(
                "namespaced cellSmartTagPr attributes are unsupported",
            ));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        match local.as_ref() {
            b"key" if key.is_none() => key = Some(decoded),
            b"val" if value.is_none() => value = Some(decoded),
            b"key" | b"val" => return Err(invalid("duplicate cellSmartTagPr attribute")),
            _ => {
                return Err(invalid(format!(
                    "unexpected cellSmartTagPr attribute '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
        }
    }
    Property::new(
        key.ok_or_else(|| invalid("cellSmartTagPr requires key"))?,
        value.ok_or_else(|| invalid("cellSmartTagPr requires val"))?,
    )
}

fn reject_attributes(element: &BytesStart<'_>, name: &str) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "unexpected {name} attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn conformance_of(namespace: &ResolveResult<'_>) -> Result<Conformance> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == TRANSITIONAL.as_bytes() => {
            Ok(Conformance::Transitional)
        },
        ResolveResult::Bound(value) if value.as_ref() == STRICT.as_bytes() => {
            Ok(Conformance::Strict)
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            Err(invalid("worksheet root uses an unsupported namespace"))
        },
    }
}

fn is_main(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == TRANSITIONAL.as_bytes() || value.as_ref() == STRICT.as_bytes())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid cellSmartTag {name} boolean '{value}'"
        ))),
    }
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet smart-tag XML: {error}"
    )))
}
