//! Bounded, namespace-aware ODF text-web content validation and projection.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the streaming codec is ordered by validation and projection pipeline"
)]

use litchi_core::{Error, Result};
use litchi_odf_common::compact_xml;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

const MAX_BLOCK_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = compact_xml::DEFAULT_MAX_DEPTH;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Element {
    Body,
    DocumentContent,
    Heading,
    LineBreak,
    Link,
    Other,
    Paragraph,
    Space,
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

#[derive(Clone, Debug)]
pub(crate) struct ReplacementSite {
    pub(crate) prefix: String,
    pub(crate) range: Range<usize>,
    pub(crate) suffix: String,
}

pub(crate) struct Projection {
    pub(crate) headings: Vec<crate::heading::Heading>,
    pub(crate) order: Vec<BlockOrder>,
    pub(crate) paragraphs: Vec<ParagraphSite>,
}

struct ActiveBlock {
    content_start: usize,
    has_children: bool,
    kind: Element,
    level: u8,
    link: Option<ActiveLink>,
    links: Vec<crate::link::Link>,
    style_name: Option<String>,
    text: String,
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

/// Project inert text semantics and lossless paragraph edit sites.
pub(crate) fn project(xml: &str) -> Result<Projection> {
    validate_source(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut active = None::<ActiveBlock>;
    let mut headings = Vec::new();
    let mut order = Vec::new();
    let mut paragraphs = Vec::new();
    let mut stack = Vec::new();

    loop {
        let event_start = source_offset(reader.buffer_position())?;
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let current = element(&reader, start.name());
                reserve(&mut stack, "OTH XML projection stack")?;
                stack.push(current);
                if matches!(current, Element::Paragraph | Element::Heading) {
                    if active.is_some() {
                        return invalid("OTH text blocks cannot contain text blocks");
                    }
                    let content_start = source_offset(reader.buffer_position())?;
                    active = Some(ActiveBlock::new(&reader, &start, current, content_start)?);
                } else if let Some(block) = active.as_mut() {
                    block.start_child(&reader, &start, current)?;
                }
            },
            Event::Empty(start) => {
                let current = element(&reader, start.name());
                if matches!(current, Element::Paragraph | Element::Heading) {
                    if active.is_some() {
                        return invalid("OTH text blocks cannot contain text blocks");
                    }
                    let event_end = source_offset(reader.buffer_position())?;
                    let block = ActiveBlock::new(&reader, &start, current, event_end)?;
                    let replacement = empty_replacement(xml, event_start..event_end, &start)?;
                    publish_block(
                        block,
                        Some(replacement),
                        &mut paragraphs,
                        &mut headings,
                        &mut order,
                    )?;
                } else if let Some(block) = active.as_mut() {
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
                    )?;
                } else if let Some(block) = active.as_mut() {
                    block.end_child(current)?;
                }
            },
            Event::Eof => {
                return Ok(Projection {
                    headings,
                    order,
                    paragraphs,
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
        Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::Other
        | Element::Paragraph
        | Element::Space
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
        Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::Other
        | Element::Paragraph
        | Element::Space
        | Element::Tab
            if root =>
        {
            invalid("OTH root must be office:document-content in the ODF office namespace")
        },
        Element::Heading
        | Element::LineBreak
        | Element::Link
        | Element::Other
        | Element::Paragraph
        | Element::Space
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
            has_children: false,
            kind,
            level,
            link: None,
            links: Vec::new(),
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
            Element::LineBreak => self.append("\n"),
            Element::Space => self.append_spaces(reader, start),
            Element::Tab => self.append("\t"),
            Element::Body
            | Element::DocumentContent
            | Element::Heading
            | Element::Other
            | Element::Paragraph
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
            Element::Space => self.append_spaces(reader, start),
            Element::Tab => self.append("\t"),
            Element::Body
            | Element::DocumentContent
            | Element::Heading
            | Element::Other
            | Element::Paragraph
            | Element::Text => Ok(()),
        }
    }

    fn end_child(&mut self, current: Element) -> Result<()> {
        if current == Element::Link {
            self.end_link()?;
        }
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
            b"h" => Element::Heading,
            b"line-break" => Element::LineBreak,
            b"p" => Element::Paragraph,
            b"s" => Element::Space,
            b"tab" => Element::Tab,
            _ => Element::Other,
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            Element::Other
        },
    }
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
    headings: &mut Vec<crate::heading::Heading>,
    order: &mut Vec<BlockOrder>,
) -> Result<()> {
    if block.link.is_some() {
        return invalid("OTH hyperlink is not closed");
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
                ),
            });
            order.push(BlockOrder::Paragraph(index));
        },
        Element::Heading => {
            reserve(headings, "OTH heading projection")?;
            reserve(order, "OTH block order")?;
            let index = headings.len();
            headings.push(crate::heading::Heading::projected(
                block.level,
                block.links,
                block.style_name,
                block.text,
            ));
            order.push(BlockOrder::Heading(index));
        },
        Element::Body
        | Element::DocumentContent
        | Element::LineBreak
        | Element::Link
        | Element::Other
        | Element::Space
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
        assert_eq!(projected.headings[0].level(), 2);
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
