//! Bounded, namespace-aware ODF text-web content validation and projection.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the streaming codec is ordered by validation and projection pipeline"
)]

use litchi_core::{Error, Position, Result};
use litchi_odf_common::compact_xml;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use std::collections::BTreeMap;
use std::ops::Range;

const MAX_BLOCK_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = compact_xml::DEFAULT_MAX_DEPTH;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const FORM_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const FO_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Element {
    Body,
    Bookmark,
    BookmarkEnd,
    BookmarkStart,
    DocumentContent,
    Field,
    Form,
    FormControl,
    Heading,
    LineBreak,
    Link,
    List,
    ListItem,
    Other,
    Paragraph,
    Resource,
    Space,
    Span,
    Tab,
    Text,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum BlockOrder {
    Heading(usize),
    Paragraph(usize),
}

pub(crate) struct ParagraphSite {
    pub(crate) replacement: Option<ReplacementSite>,
    pub(crate) value: crate::paragraph::Paragraph,
}

pub(crate) struct HeadingSite {
    pub(crate) replacement: Option<ReplacementSite>,
    pub(crate) value: crate::heading::Heading,
}

#[derive(Clone, Debug)]
pub(crate) struct ReplacementSite {
    pub(crate) prefix: String,
    pub(crate) range: Range<usize>,
    pub(crate) suffix: String,
}

pub(crate) struct Projection {
    pub(crate) bookmarks: Vec<crate::bookmark::Bookmark>,
    pub(crate) forms: Vec<crate::form::Form>,
    pub(crate) headings: Vec<HeadingSite>,
    pub(crate) lists: Vec<crate::list::List>,
    pub(crate) list_sites: Vec<ReplacementSite>,
    pub(crate) order: Vec<BlockOrder>,
    pub(crate) paragraphs: Vec<ParagraphSite>,
    pub(crate) resources: Vec<crate::resource::Resource>,
    pub(crate) text_close: usize,
}

struct ActiveBlock {
    content_start: usize,
    has_children: bool,
    kind: Element,
    level: u8,
    link: Option<ActiveLink>,
    links: Vec<crate::link::Link>,
    fields: Vec<crate::field::Field>,
    open_fields: Vec<ActiveField>,
    open_spans: Vec<ActiveSpan>,
    runs: Vec<crate::formatting::Run>,
    style_name: Option<String>,
    text: String,
}

struct ActiveField {
    fixed: bool,
    kind: crate::field::Kind,
    name: Option<String>,
    start: usize,
    value: Option<String>,
}

struct ActiveSpan {
    start: usize,
    style_name: Option<String>,
}

struct PendingList {
    items: Vec<PendingListItem>,
    level: usize,
    source_start: usize,
    style_name: Option<String>,
}

struct PendingListItem {
    paragraph_positions: Vec<usize>,
    start_value: Option<u32>,
}

struct PendingForm {
    controls: Vec<crate::form::Control>,
    name: Option<String>,
}

struct PendingStyle {
    family: Option<String>,
    name: String,
    parent_name: Option<String>,
    text_properties: Option<crate::style::TextProperties>,
}

struct ActiveLink {
    href: String,
    text_start: usize,
}

/// Validate compact XML supplied to the fresh-package authoring boundary.
pub(crate) fn validate_authored(xml: &str) -> Result<()> {
    compact_xml::validate(xml.as_bytes()).map_err(Error::from)?;
    validate_structure(xml)
}

/// Compacts producer formatting before publishing a changed whole XML part.
/// Character data, comments, CDATA, references, and attribute values remain
/// byte-exact; only whitespace outside quoted markup and formatting-only text
/// nodes is removed.
pub(crate) fn compact_for_publication(xml: &str) -> Result<String> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut output = String::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "OTH compact XML output",
            source,
        })?;
    let mut preserve_stack = Vec::new();
    let mut text_block_depth = 0_usize;
    loop {
        let start_offset = source_offset(reader.buffer_position())?;
        let event = reader.read_event().map_err(|error| xml_error(&error))?;
        let end_offset = source_offset(reader.buffer_position())?;
        match event {
            Event::Start(start) => {
                if matches!(
                    element(&reader, start.name()),
                    Element::Heading | Element::Paragraph
                ) {
                    text_block_depth = text_block_depth.saturating_add(1);
                }
                push_compact_tag(&mut output, start.as_ref(), false)?;
                let inherited = preserve_stack.last().copied().unwrap_or(false);
                reserve(&mut preserve_stack, "OTH xml:space stack")?;
                preserve_stack.push(xml_space_preserve(&reader, &start)?.unwrap_or(inherited));
            },
            Event::Empty(start) => push_compact_tag(&mut output, start.as_ref(), true)?,
            Event::End(end) => {
                let closes_text_block = matches!(
                    element(&reader, end.name()),
                    Element::Heading | Element::Paragraph
                );
                output.push_str("</");
                output.push_str(&String::from_utf8_lossy(end.name().as_ref()));
                output.push('>');
                let _ = preserve_stack.pop();
                if closes_text_block {
                    text_block_depth = text_block_depth.saturating_sub(1);
                }
            },
            Event::Text(text)
                if !preserve_stack.last().copied().unwrap_or(false)
                    && text.as_ref().iter().all(u8::is_ascii_whitespace)
                    && text
                        .as_ref()
                        .iter()
                        .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t')) =>
            {
                if text_block_depth > 0 {
                    return invalid(
                        "OTH changed producer XML has ambiguous whitespace-only block text",
                    );
                }
            },
            Event::Decl(decl) => {
                output.push_str("<?xml version=\"");
                output.push_str(&String::from_utf8_lossy(&decl.version().map_err(
                    |error| Error::InvalidFormat(format!("invalid OTH XML declaration: {error}")),
                )?));
                output.push('"');
                if let Some(encoded_value) = decl.encoding() {
                    let encoding = encoded_value.map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH XML encoding: {error}"))
                    })?;
                    output.push_str(" encoding=\"");
                    output.push_str(&String::from_utf8_lossy(&encoding));
                    output.push('"');
                }
                if let Some(encoded_value) = decl.standalone() {
                    let standalone = encoded_value.map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH XML standalone: {error}"))
                    })?;
                    output.push_str(" standalone=\"");
                    output.push_str(&String::from_utf8_lossy(&standalone));
                    output.push('"');
                }
                output.push_str("?>");
            },
            Event::DocType(_) => return invalid("OTH changed XML cannot contain a DTD"),
            Event::Eof => return Ok(output),
            Event::CData(_)
            | Event::Comment(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {
                output.push_str(xml.get(start_offset..end_offset).ok_or_else(|| {
                    Error::InvalidFormat("OTH compact XML source span is invalid".to_string())
                })?);
            },
        }
    }
}

fn push_compact_tag(output: &mut String, source: &[u8], empty: bool) -> Result<()> {
    let mut compact = Vec::new();
    compact
        .try_reserve_exact(source.len().saturating_add(3))
        .map_err(|allocation_error| Error::Allocation {
            resource: "OTH compact XML tag",
            source: allocation_error,
        })?;
    compact.push(b'<');
    let mut quote = None;
    let mut pending_space = false;
    for byte in source {
        match (quote, *byte) {
            (Some(delimiter), current) => {
                compact.push(current);
                if current == delimiter {
                    quote = None;
                }
            },
            (None, b'\'' | b'"') => {
                if pending_space && !matches!(compact.last(), Some(b'<' | b'=' | b' ')) {
                    compact.push(b' ');
                }
                pending_space = false;
                quote = Some(*byte);
                compact.push(*byte);
            },
            (None, current) if current.is_ascii_whitespace() => pending_space = true,
            (None, current) => {
                if pending_space
                    && current != b'='
                    && !matches!(compact.last(), Some(b'<' | b'=' | b' '))
                {
                    compact.push(b' ');
                }
                pending_space = false;
                if current == b'=' && compact.last() == Some(&b' ') {
                    let _ = compact.pop();
                }
                compact.push(current);
            },
        }
    }
    if empty {
        compact.push(b'/');
    }
    compact.push(b'>');
    output.push_str(std::str::from_utf8(&compact).map_err(|error| {
        Error::InvalidFormat(format!("invalid OTH compact XML tag UTF-8: {error}"))
    })?);
    Ok(())
}

fn xml_space_preserve(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Option<bool>> {
    for raw in start.attributes() {
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid OTH attribute: {error}")))?;
        if attribute.key.as_ref() == b"xml:space" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH xml:space value: {error}"))
                })?;
            return Ok(match value.as_ref() {
                "preserve" => Some(true),
                "default" => Some(false),
                _ => None,
            });
        }
    }
    Ok(None)
}

/// Project inert text semantics and lossless paragraph edit sites.
pub(crate) fn project(xml: &str) -> Result<Projection> {
    validate_source(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut active = None::<ActiveBlock>;
    let mut bookmarks = Vec::new();
    let mut bookmark_starts = BTreeMap::<String, crate::bookmark::Anchor>::new();
    let mut forms = Vec::new();
    let mut form_stack = Vec::<PendingForm>::new();
    let mut headings = Vec::new();
    let mut list_stack = Vec::<PendingList>::new();
    let mut lists = Vec::new();
    let mut list_sites = Vec::new();
    let mut order = Vec::new();
    let mut paragraphs = Vec::new();
    let mut resources = Vec::new();
    let mut stack = Vec::new();
    let mut text_close = None;

    loop {
        let event_start = source_offset(reader.buffer_position())?;
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let current = element(&reader, start.name());
                inventory_resource(&reader, &start, current, &mut resources)?;
                start_form(&reader, &start, current, &mut form_stack)?;
                start_list(&reader, &start, current, event_start, &mut list_stack)?;
                reserve(&mut stack, "OTH XML projection stack")?;
                stack.push(current);
                if matches!(current, Element::Paragraph | Element::Heading) {
                    if active.is_some() {
                        return invalid("OTH text blocks cannot contain text blocks");
                    }
                    let content_start = source_offset(reader.buffer_position())?;
                    active = Some(ActiveBlock::new(&reader, &start, current, content_start)?);
                } else if let Some(block) = active.as_mut() {
                    bookmark_event(
                        &reader,
                        &start,
                        current,
                        block,
                        order.len(),
                        &mut bookmark_starts,
                        &mut bookmarks,
                    )?;
                    block.start_child(&reader, &start, current)?;
                }
            },
            Event::Empty(start) => {
                let current = element(&reader, start.name());
                let event_end = source_offset(reader.buffer_position())?;
                inventory_resource(&reader, &start, current, &mut resources)?;
                empty_form(&reader, &start, current, &mut form_stack, &mut forms)?;
                empty_list(
                    &reader,
                    &start,
                    current,
                    event_start..event_end,
                    &mut list_stack,
                    &paragraphs,
                    &mut lists,
                    &mut list_sites,
                )?;
                if matches!(current, Element::Paragraph | Element::Heading) {
                    if active.is_some() {
                        return invalid("OTH text blocks cannot contain text blocks");
                    }
                    let block = ActiveBlock::new(&reader, &start, current, event_end)?;
                    let replacement = empty_replacement(xml, event_start..event_end, &start)?;
                    publish_block(
                        block,
                        Some(replacement),
                        &mut paragraphs,
                        &mut headings,
                        &mut order,
                        &mut list_stack,
                    )?;
                } else if let Some(block) = active.as_mut() {
                    bookmark_event(
                        &reader,
                        &start,
                        current,
                        block,
                        order.len(),
                        &mut bookmark_starts,
                        &mut bookmarks,
                    )?;
                    block.empty_child(&reader, &start, current)?;
                }
            },
            Event::Text(text) if active.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH character data: {error}"))
                })?;
                active_text(&mut active, &value)?;
            },
            Event::CData(text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::InvalidFormat(format!("invalid OTH CDATA: {error}")))?;
                active_text(&mut active, &value)?;
            },
            Event::GeneralRef(reference) if active.is_some() => {
                let value = reference_value(&reference)?;
                active_text(&mut active, &value)?;
            },
            Event::End(end) => {
                let current = element(&reader, end.name());
                let Some(open) = stack.pop() else {
                    return invalid("OTH XML end tag has no matching start tag");
                };
                if current != open {
                    return invalid("OTH XML end tag does not match its start tag");
                }
                if matches!(current, Element::Paragraph | Element::Heading) {
                    let block = active.take().ok_or_else(|| {
                        Error::InvalidFormat("OTH text block state is missing".to_string())
                    })?;
                    let replacement = (!block.has_children).then_some(ReplacementSite {
                        prefix: String::new(),
                        range: block.content_start..event_start,
                        suffix: String::new(),
                    });
                    publish_block(
                        block,
                        replacement,
                        &mut paragraphs,
                        &mut headings,
                        &mut order,
                        &mut list_stack,
                    )?;
                } else if let Some(block) = active.as_mut() {
                    block.end_child(current)?;
                }
                if current == Element::Text {
                    text_close = Some(event_start);
                }
                let event_end = source_offset(reader.buffer_position())?;
                end_list(
                    current,
                    event_end,
                    &mut list_stack,
                    &paragraphs,
                    &mut lists,
                    &mut list_sites,
                )?;
                end_form(current, &mut form_stack, &mut forms)?;
            },
            Event::Eof => {
                if !bookmark_starts.is_empty() {
                    return invalid("OTH bookmark range is not closed");
                }
                return Ok(Projection {
                    bookmarks,
                    forms,
                    headings,
                    lists,
                    list_sites,
                    order,
                    paragraphs,
                    resources,
                    text_close: text_close.ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH office:text close position is missing".to_string(),
                        )
                    })?,
                });
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
}

/// Inventories named style declarations without resolving inheritance.
pub(crate) fn project_styles(
    xml: &str,
    origin: crate::style::Origin,
) -> Result<Vec<crate::style::Style>> {
    if xml.len() > compact_xml::DEFAULT_MAX_BYTES {
        return invalid("OTH style source exceeds the input byte limit");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut active = None::<PendingStyle>;
    let mut styles = Vec::new();
    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let (namespace, local) = reader.resolver().resolve_element(start.name());
                if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
                    && local.as_ref() == b"style"
                {
                    if active.is_some() {
                        return invalid("OTH named styles cannot be nested");
                    }
                    active = Some(pending_style(&reader, &start)?);
                } else if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
                    && local.as_ref() == b"text-properties"
                    && let Some(style) = active.as_mut()
                {
                    let properties = text_properties(&reader, &start)?;
                    style.text_properties = (!properties.is_empty()).then_some(properties);
                }
            },
            Event::Empty(start) => {
                let (namespace, local) = reader.resolver().resolve_element(start.name());
                if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
                    && local.as_ref() == b"style"
                {
                    let style = pending_style(&reader, &start)?;
                    reserve(&mut styles, "OTH style inventory")?;
                    styles.push(crate::style::Style::projected(
                        style.name,
                        style.family,
                        style.parent_name,
                        origin,
                        None,
                    ));
                } else if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
                    && local.as_ref() == b"text-properties"
                    && let Some(style) = active.as_mut()
                {
                    let properties = text_properties(&reader, &start)?;
                    style.text_properties = (!properties.is_empty()).then_some(properties);
                }
            },
            Event::End(end) => {
                let (namespace, local) = reader.resolver().resolve_element(end.name());
                if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
                    && local.as_ref() == b"style"
                {
                    let style = active.take().ok_or_else(|| {
                        Error::InvalidFormat("OTH named style state is missing".to_string())
                    })?;
                    reserve(&mut styles, "OTH style inventory")?;
                    styles.push(crate::style::Style::projected(
                        style.name,
                        style.family,
                        style.parent_name,
                        origin,
                        style.text_properties,
                    ));
                }
            },
            Event::DocType(_) => return invalid("OTH style source cannot contain a DTD"),
            Event::GeneralRef(reference) => {
                let _ = reference_value(&reference)?;
            },
            Event::Eof => return Ok(styles),
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
}

fn pending_style(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<PendingStyle> {
    Ok(PendingStyle {
        family: optional_attribute(reader, start, STYLE_NAMESPACE, b"family")?,
        name: optional_attribute(reader, start, STYLE_NAMESPACE, b"name")?
            .ok_or_else(|| Error::InvalidFormat("OTH named style has no style:name".to_string()))?,
        parent_name: optional_attribute(reader, start, STYLE_NAMESPACE, b"parent-style-name")?,
        text_properties: None,
    })
}

fn text_properties(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<crate::style::TextProperties> {
    let weight =
        optional_attribute(reader, start, FO_NAMESPACE, b"font-weight")?.and_then(|value| {
            match value.as_str() {
                "bold" => Some(crate::style::Weight::Bold),
                "normal" => Some(crate::style::Weight::Normal),
                _ => None,
            }
        });
    let slant =
        optional_attribute(reader, start, FO_NAMESPACE, b"font-style")?.and_then(
            |value| match value.as_str() {
                "italic" => Some(crate::style::Slant::Italic),
                "normal" => Some(crate::style::Slant::Normal),
                _ => None,
            },
        );
    Ok(crate::style::TextProperties::projected(
        optional_attribute(reader, start, FO_NAMESPACE, b"color")?,
        optional_attribute(reader, start, FO_NAMESPACE, b"background-color")?,
        weight,
        slant,
    ))
}

fn validate_source(xml: &str) -> Result<()> {
    if xml.len() > compact_xml::DEFAULT_MAX_BYTES {
        return invalid("OTH content.xml exceeds the input byte limit");
    }
    validate_structure(xml)
}

fn validate_structure(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut body_seen = false;
    let mut root_closed = false;
    let mut stack = Vec::new();
    let mut text_seen = false;

    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let current = element(&reader, start.name());
                handle_start(
                    &mut stack,
                    current,
                    &mut body_seen,
                    &mut text_seen,
                    root_closed,
                )?;
            },
            Event::Empty(start) => {
                let current = element(&reader, start.name());
                handle_empty(&stack, current, &mut body_seen, &mut text_seen, root_closed)?;
            },
            Event::End(end) => {
                let current = element(&reader, end.name());
                let Some(open) = stack.pop() else {
                    return invalid("OTH XML end tag has no matching start tag");
                };
                if current != open {
                    return invalid("OTH XML end tag does not match its start tag");
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::GeneralRef(reference) => {
                let _ = reference_value(&reference)?;
            },
            Event::Text(text)
                if stack.is_empty() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return invalid("OTH content.xml has character data outside its root");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }

    if !root_closed || !stack.is_empty() {
        return invalid("OTH content.xml has no closed document-content root");
    }
    if !body_seen || !text_seen {
        return invalid("OTH content.xml must contain exactly one office:body/office:text chain");
    }
    Ok(())
}

fn handle_start(
    stack: &mut Vec<Element>,
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
    root_closed: bool,
) -> Result<()> {
    if root_closed || stack.len() >= MAX_DEPTH {
        return invalid("OTH content.xml has an invalid element depth or trailing root");
    }
    validate_text_block_position(stack, current)?;
    validate_position(
        stack.last().copied(),
        stack.is_empty(),
        current,
        body_seen,
        text_seen,
    )?;
    reserve(stack, "OTH XML validation stack")?;
    stack.push(current);
    Ok(())
}

fn handle_empty(
    stack: &[Element],
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
    root_closed: bool,
) -> Result<()> {
    if root_closed || stack.len() >= MAX_DEPTH {
        return invalid("OTH content.xml has an invalid empty element depth or trailing root");
    }
    validate_text_block_position(stack, current)?;
    validate_position(
        stack.last().copied(),
        stack.is_empty(),
        current,
        body_seen,
        text_seen,
    )?;
    match current {
        Element::Body | Element::DocumentContent => {
            invalid("OTH content.xml cannot use an empty root or body element")
        },
        Element::Bookmark
        | Element::BookmarkEnd
        | Element::BookmarkStart
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::List
        | Element::ListItem
        | Element::Other
        | Element::Paragraph
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab
        | Element::Text => Ok(()),
    }
}

fn validate_text_block_position(stack: &[Element], current: Element) -> Result<()> {
    if matches!(current, Element::Paragraph | Element::Heading) && !stack.contains(&Element::Text) {
        return invalid("OTH text blocks must occur inside the office:text body");
    }
    Ok(())
}

fn validate_position(
    parent: Option<Element>,
    root: bool,
    current: Element,
    body_seen: &mut bool,
    text_seen: &mut bool,
) -> Result<()> {
    match current {
        Element::DocumentContent if root => Ok(()),
        Element::DocumentContent => {
            invalid("OTH content.xml has a nested or duplicate document-content root")
        },
        Element::Body if parent == Some(Element::DocumentContent) && !*body_seen => {
            *body_seen = true;
            Ok(())
        },
        Element::Body => {
            invalid("OTH office:body must occur once directly inside document-content")
        },
        Element::Text if parent == Some(Element::Body) && !*text_seen => {
            *text_seen = true;
            Ok(())
        },
        Element::Text => invalid("OTH office:text must occur once directly inside office:body"),
        Element::Bookmark
        | Element::BookmarkEnd
        | Element::BookmarkStart
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::List
        | Element::ListItem
        | Element::Other
        | Element::Paragraph
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab
            if root =>
        {
            invalid("OTH root must be office:document-content in the ODF office namespace")
        },
        Element::Bookmark
        | Element::BookmarkEnd
        | Element::BookmarkStart
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::List
        | Element::ListItem
        | Element::Other
        | Element::Paragraph
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab => Ok(()),
    }
}

impl ActiveBlock {
    fn new(
        reader: &NsReader<&[u8]>,
        start: &BytesStart<'_>,
        kind: Element,
        content_start: usize,
    ) -> Result<Self> {
        let style_name = optional_attribute(reader, start, TEXT_NAMESPACE, b"style-name")?;
        let level = if kind == Element::Heading {
            optional_attribute(reader, start, TEXT_NAMESPACE, b"outline-level")?.map_or(
                Ok(1),
                |value| {
                    value.parse::<u8>().map_err(|_error| {
                        Error::InvalidFormat("invalid OTH heading outline level".to_string())
                    })
                },
            )?
        } else {
            0
        };
        if kind == Element::Heading && level == 0 {
            return invalid("OTH heading outline level must be positive");
        }
        Ok(Self {
            content_start,
            fields: Vec::new(),
            has_children: false,
            kind,
            level,
            link: None,
            links: Vec::new(),
            open_fields: Vec::new(),
            open_spans: Vec::new(),
            runs: Vec::new(),
            style_name,
            text: String::new(),
        })
    }

    fn start_child(
        &mut self,
        reader: &NsReader<&[u8]>,
        start: &BytesStart<'_>,
        current: Element,
    ) -> Result<()> {
        self.has_children = true;
        match current {
            Element::Link => self.start_link(reader, start),
            Element::Span => self.start_span(reader, start),
            Element::Field => self.start_field(reader, start),
            Element::LineBreak => self.append("\n"),
            Element::Space => self.append_spaces(reader, start),
            Element::Tab => self.append("\t"),
            Element::Body
            | Element::Bookmark
            | Element::BookmarkEnd
            | Element::BookmarkStart
            | Element::DocumentContent
            | Element::Form
            | Element::FormControl
            | Element::Heading
            | Element::List
            | Element::ListItem
            | Element::Other
            | Element::Paragraph
            | Element::Resource
            | Element::Text => Ok(()),
        }
    }

    fn empty_child(
        &mut self,
        reader: &NsReader<&[u8]>,
        start: &BytesStart<'_>,
        current: Element,
    ) -> Result<()> {
        self.has_children = true;
        match current {
            Element::Link => {
                self.start_link(reader, start)?;
                self.end_link()
            },
            Element::LineBreak => self.append("\n"),
            Element::Span => {
                self.start_span(reader, start)?;
                self.end_span()
            },
            Element::Field => {
                self.start_field(reader, start)?;
                self.end_field()
            },
            Element::Space => self.append_spaces(reader, start),
            Element::Tab => self.append("\t"),
            Element::Body
            | Element::Bookmark
            | Element::BookmarkEnd
            | Element::BookmarkStart
            | Element::DocumentContent
            | Element::Form
            | Element::FormControl
            | Element::Heading
            | Element::List
            | Element::ListItem
            | Element::Other
            | Element::Paragraph
            | Element::Resource
            | Element::Text => Ok(()),
        }
    }

    fn end_child(&mut self, current: Element) -> Result<()> {
        if current == Element::Link {
            self.end_link()?;
        } else if current == Element::Span {
            self.end_span()?;
        } else if current == Element::Field {
            self.end_field()?;
        }
        Ok(())
    }

    fn start_span(&mut self, reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<()> {
        reserve(&mut self.open_spans, "OTH formatting span stack")?;
        self.open_spans.push(ActiveSpan {
            start: self.text.len(),
            style_name: optional_attribute(reader, start, TEXT_NAMESPACE, b"style-name")?,
        });
        Ok(())
    }

    fn end_span(&mut self) -> Result<()> {
        let span = self.open_spans.pop().ok_or_else(|| {
            Error::InvalidFormat("OTH formatting span state is missing".to_string())
        })?;
        if let Some(style_name) = span.style_name {
            reserve(&mut self.runs, "OTH formatting run projection")?;
            self.runs.push(crate::formatting::Run::projected(
                span.start..self.text.len(),
                style_name,
            ));
        }
        Ok(())
    }

    fn start_field(&mut self, reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<()> {
        let (_, local) = reader.resolver().resolve_element(start.name());
        let name = optional_attribute(reader, start, TEXT_NAMESPACE, b"name")?;
        let fixed = optional_attribute(reader, start, TEXT_NAMESPACE, b"fixed")?
            .is_some_and(|value| value == "true");
        let value = first_attribute_value(
            reader,
            start,
            OFFICE_NAMESPACE,
            &[
                b"string-value",
                b"value",
                b"date-value",
                b"time-value",
                b"boolean-value",
                b"currency",
            ],
        )?;
        reserve(&mut self.open_fields, "OTH field stack")?;
        self.open_fields.push(ActiveField {
            fixed,
            kind: crate::field::Kind::from_local(local.as_ref()),
            name,
            start: self.text.len(),
            value,
        });
        Ok(())
    }

    fn end_field(&mut self) -> Result<()> {
        let field = self
            .open_fields
            .pop()
            .ok_or_else(|| Error::InvalidFormat("OTH field state is missing".to_string()))?;
        reserve(&mut self.fields, "OTH field projection")?;
        self.fields.push(crate::field::Field::projected(
            field.kind,
            field.name,
            field.value,
            field.fixed,
            field.start..self.text.len(),
        ));
        Ok(())
    }

    fn start_link(&mut self, reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<()> {
        if self.link.is_some() {
            return invalid("OTH hyperlinks cannot be nested");
        }
        let href = optional_attribute(reader, start, XLINK_NAMESPACE, b"href")?
            .ok_or_else(|| Error::InvalidFormat("OTH hyperlink has no xlink:href".to_string()))?;
        self.link = Some(ActiveLink {
            href,
            text_start: self.text.len(),
        });
        Ok(())
    }

    fn end_link(&mut self) -> Result<()> {
        let link = self
            .link
            .take()
            .ok_or_else(|| Error::InvalidFormat("OTH hyperlink state is missing".to_string()))?;
        let label = self.text.get(link.text_start..).ok_or_else(|| {
            Error::InvalidFormat("OTH hyperlink text range is invalid".to_string())
        })?;
        reserve(&mut self.links, "OTH hyperlink projection")?;
        self.links
            .push(crate::link::Link::new(link.href, label.to_owned()));
        Ok(())
    }

    fn append_spaces(&mut self, reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<()> {
        let count =
            optional_attribute(reader, start, TEXT_NAMESPACE, b"c")?.map_or(Ok(1), |value| {
                value
                    .parse::<usize>()
                    .map_err(|_error| Error::InvalidFormat("invalid OTH text:s count".to_string()))
            })?;
        if count == 0 || count > MAX_BLOCK_BYTES {
            return invalid("OTH text:s count is outside the supported range");
        }
        let target = self
            .text
            .len()
            .checked_add(count)
            .ok_or_else(|| Error::InvalidFormat("OTH block text size overflow".to_string()))?;
        if target > MAX_BLOCK_BYTES {
            return invalid("OTH block text exceeds the projection limit");
        }
        self.text
            .try_reserve(count)
            .map_err(|source| Error::Allocation {
                resource: "OTH block text",
                source,
            })?;
        self.text.extend(std::iter::repeat_n(' ', count));
        Ok(())
    }

    fn append(&mut self, value: &str) -> Result<()> {
        let target = self
            .text
            .len()
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("OTH block text size overflow".to_string()))?;
        if target > MAX_BLOCK_BYTES {
            return invalid("OTH block text exceeds the projection limit");
        }
        self.text
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH block text",
                source,
            })?;
        self.text.push_str(value);
        Ok(())
    }
}

fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> Element {
    let (namespace, local) = reader.resolver().resolve_element(name);
    match namespace {
        ResolveResult::Bound(Namespace(value)) if value == OFFICE_NAMESPACE => {
            match local.as_ref() {
                b"body" => Element::Body,
                b"document-content" => Element::DocumentContent,
                b"text" => Element::Text,
                _ => Element::Other,
            }
        },
        ResolveResult::Bound(Namespace(value)) if value == TEXT_NAMESPACE => match local.as_ref() {
            b"a" => Element::Link,
            b"bookmark" => Element::Bookmark,
            b"bookmark-end" => Element::BookmarkEnd,
            b"bookmark-start" => Element::BookmarkStart,
            b"h" => Element::Heading,
            b"line-break" => Element::LineBreak,
            b"list" => Element::List,
            b"list-item" | b"list-header" => Element::ListItem,
            b"p" => Element::Paragraph,
            b"s" => Element::Space,
            b"span" => Element::Span,
            b"tab" => Element::Tab,
            _ if is_field(local.as_ref()) => Element::Field,
            _ => Element::Other,
        },
        ResolveResult::Bound(Namespace(value)) if value == DRAW_NAMESPACE => match local.as_ref() {
            b"image" | b"object" | b"object-ole" | b"plugin" | b"floating-frame" => {
                Element::Resource
            },
            _ => Element::Other,
        },
        ResolveResult::Bound(Namespace(value)) if value == FORM_NAMESPACE => {
            if local.as_ref() == b"form" {
                Element::Form
            } else {
                Element::FormControl
            }
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            Element::Other
        },
    }
}

fn is_field(local: &[u8]) -> bool {
    matches!(
        local,
        b"author-initials"
            | b"author-name"
            | b"chapter"
            | b"conditional-text"
            | b"creation-date"
            | b"creation-time"
            | b"creator"
            | b"date"
            | b"description"
            | b"editing-cycles"
            | b"editing-duration"
            | b"expression"
            | b"file-name"
            | b"hidden-paragraph"
            | b"hidden-text"
            | b"initial-creator"
            | b"keywords"
            | b"modification-date"
            | b"modification-time"
            | b"page-continuation"
            | b"page-count"
            | b"page-number"
            | b"paragraph-count"
            | b"placeholder"
            | b"printed-by"
            | b"print-date"
            | b"print-time"
            | b"reference-ref"
            | b"sequence-ref"
            | b"sheet-name"
            | b"subject"
            | b"table-count"
            | b"template-name"
            | b"text-input"
            | b"time"
            | b"title"
            | b"user-defined"
            | b"user-field-get"
            | b"variable-get"
            | b"word-count"
    )
}

fn active_text(active: &mut Option<ActiveBlock>, value: &str) -> Result<()> {
    active
        .as_mut()
        .ok_or_else(|| Error::InvalidFormat("OTH text block state is missing".to_string()))?
        .append(value)
}

fn publish_block(
    block: ActiveBlock,
    replacement: Option<ReplacementSite>,
    paragraphs: &mut Vec<ParagraphSite>,
    headings: &mut Vec<HeadingSite>,
    order: &mut Vec<BlockOrder>,
    list_stack: &mut [PendingList],
) -> Result<()> {
    if block.link.is_some() || !block.open_fields.is_empty() || !block.open_spans.is_empty() {
        return invalid("OTH inline semantic range is not closed");
    }
    match block.kind {
        Element::Paragraph => {
            reserve(paragraphs, "OTH paragraph projection")?;
            reserve(order, "OTH block order")?;
            let index = paragraphs.len();
            paragraphs.push(ParagraphSite {
                replacement,
                value: crate::paragraph::Paragraph::projected(
                    block.text,
                    block.style_name,
                    block.links,
                    block.runs,
                    block.fields,
                ),
            });
            if let Some(item) = list_stack.last_mut().and_then(|list| list.items.last_mut()) {
                reserve(
                    &mut item.paragraph_positions,
                    "OTH list item paragraph positions",
                )?;
                item.paragraph_positions.push(index);
            }
            order.push(BlockOrder::Paragraph(index));
        },
        Element::Heading => {
            reserve(headings, "OTH heading projection")?;
            reserve(order, "OTH block order")?;
            let index = headings.len();
            headings.push(HeadingSite {
                replacement,
                value: crate::heading::Heading::projected(
                    block.level,
                    block.links,
                    block.runs,
                    block.fields,
                    block.style_name,
                    block.text,
                ),
            });
            order.push(BlockOrder::Heading(index));
        },
        Element::Body
        | Element::Bookmark
        | Element::BookmarkEnd
        | Element::BookmarkStart
        | Element::DocumentContent
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::LineBreak
        | Element::Link
        | Element::List
        | Element::ListItem
        | Element::Other
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab
        | Element::Text => return invalid("OTH projected block has the wrong element kind"),
    }
    Ok(())
}

fn empty_replacement(
    xml: &str,
    range: Range<usize>,
    start: &BytesStart<'_>,
) -> Result<ReplacementSite> {
    let source = xml
        .get(range.clone())
        .ok_or_else(|| Error::InvalidFormat("OTH empty block span is invalid".to_string()))?;
    let open = source
        .strip_suffix("/>")
        .ok_or_else(|| Error::InvalidFormat("OTH empty block markup is invalid".to_string()))?
        .trim_end();
    let qualified_name = start.name();
    let name = std::str::from_utf8(qualified_name.as_ref())
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH block name: {error}")))?;
    Ok(ReplacementSite {
        prefix: format!("{open}>"),
        range,
        suffix: format!("</{name}>"),
    })
}

fn bookmark_event(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    block: &ActiveBlock,
    block_position: usize,
    starts: &mut BTreeMap<String, crate::bookmark::Anchor>,
    bookmarks: &mut Vec<crate::bookmark::Bookmark>,
) -> Result<()> {
    if !matches!(
        current,
        Element::Bookmark | Element::BookmarkStart | Element::BookmarkEnd
    ) {
        return Ok(());
    }
    let name = optional_attribute(reader, start, TEXT_NAMESPACE, b"name")?
        .ok_or_else(|| Error::InvalidFormat("OTH bookmark has no text:name".to_string()))?;
    let anchor = crate::bookmark::Anchor::new(Position::new(block_position), block.text.len());
    match current {
        Element::Bookmark => {
            reserve(bookmarks, "OTH bookmark projection")?;
            bookmarks.push(crate::bookmark::Bookmark::Point { name, at: anchor });
        },
        Element::BookmarkStart => {
            if starts.insert(name, anchor).is_some() {
                return invalid("OTH bookmark start name is duplicated while open");
            }
        },
        Element::BookmarkEnd => {
            let range_start = starts.remove(&name).ok_or_else(|| {
                Error::InvalidFormat("OTH bookmark end has no matching start".to_string())
            })?;
            reserve(bookmarks, "OTH bookmark projection")?;
            bookmarks.push(crate::bookmark::Bookmark::Range {
                name,
                start: range_start,
                end: anchor,
            });
        },
        Element::Body
        | Element::DocumentContent
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::List
        | Element::ListItem
        | Element::Other
        | Element::Paragraph
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab
        | Element::Text => {},
    }
    Ok(())
}

fn start_list(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    source_start: usize,
    stack: &mut Vec<PendingList>,
) -> Result<()> {
    match current {
        Element::List => {
            let level = stack.len().saturating_add(1);
            reserve(stack, "OTH list stack")?;
            stack.push(PendingList {
                items: Vec::new(),
                level,
                source_start,
                style_name: optional_attribute(reader, start, TEXT_NAMESPACE, b"style-name")?,
            });
        },
        Element::ListItem => {
            let list = stack.last_mut().ok_or_else(|| {
                Error::InvalidFormat("OTH list item occurs outside a text:list".to_string())
            })?;
            reserve(&mut list.items, "OTH list items")?;
            let start_value = optional_attribute(reader, start, TEXT_NAMESPACE, b"start-value")?
                .map(|value| {
                    value.parse::<u32>().map_err(|_error| {
                        Error::InvalidFormat("invalid OTH list item start value".to_string())
                    })
                })
                .transpose()?;
            list.items.push(PendingListItem {
                paragraph_positions: Vec::new(),
                start_value,
            });
        },
        Element::Body
        | Element::Bookmark
        | Element::BookmarkEnd
        | Element::BookmarkStart
        | Element::DocumentContent
        | Element::Field
        | Element::Form
        | Element::FormControl
        | Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::Other
        | Element::Paragraph
        | Element::Resource
        | Element::Space
        | Element::Span
        | Element::Tab
        | Element::Text => {},
    }
    Ok(())
}

fn empty_list(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    source_range: Range<usize>,
    stack: &mut Vec<PendingList>,
    paragraphs: &[ParagraphSite],
    lists: &mut Vec<crate::list::List>,
    sites: &mut Vec<ReplacementSite>,
) -> Result<()> {
    start_list(reader, start, current, source_range.start, stack)?;
    if current == Element::List {
        publish_list(stack, paragraphs, lists, source_range.end, sites)?;
    }
    Ok(())
}

fn end_list(
    current: Element,
    source_end: usize,
    stack: &mut Vec<PendingList>,
    paragraphs: &[ParagraphSite],
    lists: &mut Vec<crate::list::List>,
    sites: &mut Vec<ReplacementSite>,
) -> Result<()> {
    if current == Element::List {
        publish_list(stack, paragraphs, lists, source_end, sites)?;
    }
    Ok(())
}

fn publish_list(
    stack: &mut Vec<PendingList>,
    paragraphs: &[ParagraphSite],
    lists: &mut Vec<crate::list::List>,
    source_end: usize,
    sites: &mut Vec<ReplacementSite>,
) -> Result<()> {
    let list = stack
        .pop()
        .ok_or_else(|| Error::InvalidFormat("OTH list state is missing".to_string()))?;
    let mut items = Vec::new();
    items
        .try_reserve_exact(list.items.len())
        .map_err(|source| Error::Allocation {
            resource: "OTH projected list items",
            source,
        })?;
    for item in list.items {
        let mut values = Vec::new();
        values
            .try_reserve_exact(item.paragraph_positions.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH projected list paragraphs",
                source,
            })?;
        for position in &item.paragraph_positions {
            values.push(
                paragraphs
                    .get(*position)
                    .ok_or_else(|| {
                        Error::InvalidFormat("OTH list paragraph position is invalid".to_string())
                    })?
                    .value
                    .clone(),
            );
        }
        items.push(crate::list::Item::projected(
            values,
            item.paragraph_positions
                .into_iter()
                .map(Position::new)
                .collect(),
            item.start_value,
        ));
    }
    reserve(lists, "OTH list projection")?;
    reserve(sites, "OTH list edit sites")?;
    sites.push(ReplacementSite {
        prefix: String::new(),
        range: list.source_start..source_end,
        suffix: String::new(),
    });
    lists.push(crate::list::List::projected(
        items,
        list.level,
        list.style_name,
    ));
    Ok(())
}

fn inventory_resource(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    resources: &mut Vec<crate::resource::Resource>,
) -> Result<()> {
    if current != Element::Resource {
        return Ok(());
    }
    let href = optional_attribute(reader, start, XLINK_NAMESPACE, b"href")?
        .ok_or_else(|| Error::InvalidFormat("OTH resource has no xlink:href".to_string()))?;
    let qualified = start.name();
    let (_, local) = reader.resolver().resolve_element(qualified);
    let kind = match local.as_ref() {
        b"image" => crate::resource::Kind::Image,
        b"object" => crate::resource::Kind::Object,
        b"object-ole" => crate::resource::Kind::OleObject,
        b"plugin" => crate::resource::Kind::Plugin,
        b"floating-frame" => crate::resource::Kind::FloatingFrame,
        _ => return invalid("OTH resource kind is invalid"),
    };
    reserve(resources, "OTH resource projection")?;
    resources.push(crate::resource::Resource::projected(kind, href));
    Ok(())
}

fn start_form(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    stack: &mut Vec<PendingForm>,
) -> Result<()> {
    if current == Element::Form {
        reserve(stack, "OTH form stack")?;
        stack.push(PendingForm {
            controls: Vec::new(),
            name: optional_attribute(reader, start, FORM_NAMESPACE, b"name")?,
        });
    } else if current == Element::FormControl {
        push_control(reader, start, stack)?;
    }
    Ok(())
}

fn empty_form(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    current: Element,
    stack: &mut Vec<PendingForm>,
    forms: &mut Vec<crate::form::Form>,
) -> Result<()> {
    start_form(reader, start, current, stack)?;
    if current == Element::Form {
        publish_form(stack, forms)?;
    }
    Ok(())
}

fn end_form(
    current: Element,
    stack: &mut Vec<PendingForm>,
    forms: &mut Vec<crate::form::Form>,
) -> Result<()> {
    if current == Element::Form {
        publish_form(stack, forms)?;
    }
    Ok(())
}

fn push_control(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    stack: &mut [PendingForm],
) -> Result<()> {
    let Some(form) = stack.last_mut() else {
        return Ok(());
    };
    let qualified = start.name();
    let (_, local) = reader.resolver().resolve_element(qualified);
    reserve(&mut form.controls, "OTH form controls")?;
    form.controls.push(crate::form::Control::projected(
        String::from_utf8_lossy(local.as_ref()).into_owned(),
        optional_attribute(reader, start, FORM_NAMESPACE, b"id")?,
        optional_attribute(reader, start, FORM_NAMESPACE, b"name")?,
    ));
    Ok(())
}

fn publish_form(stack: &mut Vec<PendingForm>, forms: &mut Vec<crate::form::Form>) -> Result<()> {
    let form = stack
        .pop()
        .ok_or_else(|| Error::InvalidFormat("OTH form state is missing".to_string()))?;
    reserve(forms, "OTH form projection")?;
    forms.push(crate::form::Form::projected(form.name, form.controls));
    Ok(())
}

fn first_attribute_value(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_locals: &[&[u8]],
) -> Result<Option<String>> {
    for local in expected_locals {
        if let Some(value) = optional_attribute(reader, start, expected_namespace, local)? {
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for raw in start.attributes() {
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid OTH attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected_namespace)
            && local.as_ref() == expected_local
        {
            if attribute.value.len() > MAX_BLOCK_BYTES {
                return invalid("OTH semantic attribute exceeds the byte limit");
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH attribute value: {error}"))
                })?
                .into_owned();
            if found.replace(value).is_some() {
                return invalid("OTH semantic attribute is duplicated");
            }
        }
    }
    Ok(found)
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid OTH character reference: {error}"))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "apos" => Ok("'".to_string()),
        "gt" => Ok(">".to_string()),
        "lt" => Ok("<".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("OTH custom entities are not allowed"),
    }
}

fn reserve<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn source_offset(offset: u64) -> Result<usize> {
    usize::try_from(offset)
        .map_err(|_error| Error::InvalidFormat("OTH XML byte offset overflow".to_string()))
}

fn xml_error(error: &quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid OTH content.xml: {error}"))
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_string()))
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "tests use panic-on-failure assertions")]

    use super::{BlockOrder, project, validate_authored};

    #[test]
    fn accepts_a_prefix_aliased_text_web_envelope() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:version="1.4"><o:body><o:text><t:p>text</t:p></o:text></o:body></o:document-content>"#;
        assert!(validate_authored(xml).is_ok());
    }

    #[test]
    fn refuses_wrong_family_dtd_and_misplaced_blocks() {
        let wrong_namespace = r#"<office:document-content xmlns:office="https://example.test/office"><office:body><office:text/></office:body></office:document-content>"#;
        let dtd = r#"<!DOCTYPE office:document-content [<!ENTITY x "unsafe">]><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#;
        let misplaced = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>outside</text:p><office:body><office:text/></office:body></office:document-content>"#;
        assert!(validate_authored(wrong_namespace).is_err());
        assert!(project(dtd).is_err());
        assert!(validate_authored(misplaced).is_err());
    }

    #[test]
    fn projects_blocks_links_and_odf_whitespace_without_activation() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:h text:outline-level="2">Title</text:h><text:p text:style-name="Body">plain<text:s text:c="2"/><text:a xlink:href="https://example.test">link</text:a><text:tab/><text:line-break/>end</text:p></office:text></office:body></office:document-content>"#;
        let projected = project(xml).unwrap();
        assert_eq!(
            projected.order,
            [BlockOrder::Heading(0), BlockOrder::Paragraph(0)]
        );
        assert_eq!(projected.headings[0].value.level(), 2);
        assert_eq!(projected.paragraphs[0].value.text(), "plain  link\t\nend");
        assert_eq!(projected.paragraphs[0].value.style_name(), Some("Body"));
        assert_eq!(
            projected.paragraphs[0].value.links()[0].href(),
            "https://example.test"
        );
        assert!(projected.paragraphs[0].replacement.is_none());
    }

    #[test]
    fn pretty_printed_sources_and_empty_blocks_are_projected() {
        let xml = "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">\n <office:body>\n  <office:text><text:p/></office:text>\n </office:body>\n</office:document-content>\n";
        let projected = project(xml).unwrap();
        assert_eq!(projected.paragraphs[0].value.text(), "");
        assert!(projected.paragraphs[0].replacement.is_some());
    }
}
