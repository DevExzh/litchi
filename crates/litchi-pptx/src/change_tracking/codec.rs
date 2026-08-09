//! Bounded, range-preserving XML codec for `MS-PPTX` 2.2.9.

use std::collections::HashSet;
use std::ops::Range;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Id, Shape, State};
use super::{CREATION_EXTENSION_URI, MODIFICATION_EXTENSION_URI, NAMESPACE};
use crate::{Error, Result};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const P14: &[u8] = NAMESPACE.as_bytes();
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Owner {
    Slide,
    Shape,
}

impl Owner {
    const fn parent(self) -> &'static [u8] {
        match self {
            Self::Slide => b"cSld",
            Self::Shape => b"nvPr",
        }
    }

    const fn element(self) -> &'static [u8] {
        match self {
            Self::Slide => b"creationId",
            Self::Shape => b"modId",
        }
    }

    const fn uri(self) -> &'static str {
        match self {
            Self::Slide => CREATION_EXTENSION_URI,
            Self::Shape => MODIFICATION_EXTENSION_URI,
        }
    }
}

#[derive(Debug, Clone)]
struct Element {
    span: Range<usize>,
    close_start: usize,
    empty: bool,
}

#[derive(Debug, Clone)]
struct Extension {
    element: Element,
    id: Option<Element>,
    other_content: bool,
}

#[derive(Debug, Clone)]
struct ExtensionList {
    element: Element,
    child_elements: usize,
    other_content: bool,
}

#[derive(Debug, Clone)]
pub(super) struct Source {
    pml_namespace: &'static str,
    parent: Element,
    list: Option<ExtensionList>,
    extension: Option<Extension>,
    pub(super) id: Option<Id>,
}

#[derive(Debug, Clone, Copy)]
enum Kind {
    Root,
    Parent,
    List,
    Extension { known: bool },
    Id,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Pml,
    StrictPml,
    P14,
    Other,
}

#[derive(Debug)]
struct Frame {
    kind: Kind,
    start: usize,
    child_elements: usize,
    other_content: bool,
}

pub(super) fn read_state(xml: &[u8]) -> Result<State> {
    let creation_id = read(xml, Owner::Slide)?.id;
    let scene = crate::shape::Scene::read(xml)?;
    let mut shapes = Vec::new();
    shapes
        .try_reserve_exact(scene.len())
        .map_err(|source| Error::Allocation {
            resource: "change-tracking shape state",
            source,
        })?;
    for (position, shape) in scene.iter().enumerate() {
        let trackable = !matches!(shape, crate::shape::Shape::Unknown(_));
        let modification_id = if trackable {
            let span =
                crate::tag::shape::selected_raw_span(xml, crate::shape::Key::Index(position))?;
            read_at(xml, Owner::Shape, Some(&span))?.id
        } else {
            None
        };
        shapes.push(Shape::new(
            shape.name().map(str::to_owned),
            modification_id,
            trackable,
        ));
    }
    Ok(State::new(creation_id, shapes))
}

pub(super) fn validate_unique_modification_ids(state: &State) -> Result<()> {
    let mut ids = HashSet::new();
    ids.try_reserve(state.shapes().len())
        .map_err(|source| Error::Allocation {
            resource: "change-tracking identifier uniqueness",
            source,
        })?;
    for shape in state.shapes() {
        if let Some(id) = shape.modification_id()
            && !ids.insert(id)
        {
            return Err(invalid(
                "shape modification identifiers must be unique within one slide",
            ));
        }
    }
    Ok(())
}

pub(super) fn set(xml: &[u8], owner: Owner, id: Id) -> Result<Vec<u8>> {
    let source = read(xml, owner)?;
    set_source(xml, owner, id, &source)
}

pub(super) fn set_shape(xml: &[u8], position: usize, id: Id) -> Result<Vec<u8>> {
    let span = crate::tag::shape::selected_raw_span(xml, crate::shape::Key::Index(position))?;
    let source = read_at(xml, Owner::Shape, Some(&span))?;
    set_source(xml, Owner::Shape, id, &source)
}

fn set_source(xml: &[u8], owner: Owner, id: Id, source: &Source) -> Result<Vec<u8>> {
    if source.id == Some(id) {
        return Ok(xml.to_vec());
    }
    let id_xml = id_xml(owner, id);
    if let Some(extension) = &source.extension {
        if let Some(element) = &extension.id {
            return replace(xml, element.span.clone(), id_xml.as_bytes());
        }
        if extension.element.empty {
            return expand_empty(xml, &extension.element, id_xml.as_bytes());
        }
        return replace(
            xml,
            extension.element.close_start..extension.element.close_start,
            id_xml.as_bytes(),
        );
    }
    let extension_xml = format!(
        "<p:ext xmlns:p=\"{}\" uri=\"{}\">{id_xml}</p:ext>",
        source.pml_namespace,
        owner.uri()
    );
    if let Some(list) = &source.list {
        if list.element.empty {
            return expand_empty(xml, &list.element, extension_xml.as_bytes());
        }
        return replace(
            xml,
            list.element.close_start..list.element.close_start,
            extension_xml.as_bytes(),
        );
    }
    let list_xml = format!(
        "<p:extLst xmlns:p=\"{}\">{extension_xml}</p:extLst>",
        source.pml_namespace
    );
    if source.parent.empty {
        expand_empty(xml, &source.parent, list_xml.as_bytes())
    } else {
        replace(
            xml,
            source.parent.close_start..source.parent.close_start,
            list_xml.as_bytes(),
        )
    }
}

pub(super) fn remove(xml: &[u8], owner: Owner) -> Result<Vec<u8>> {
    let source = read(xml, owner)?;
    remove_source(xml, &source)
}

pub(super) fn remove_shape(xml: &[u8], position: usize) -> Result<Vec<u8>> {
    let span = crate::tag::shape::selected_raw_span(xml, crate::shape::Key::Index(position))?;
    let source = read_at(xml, Owner::Shape, Some(&span))?;
    remove_source(xml, &source)
}

fn remove_source(xml: &[u8], source: &Source) -> Result<Vec<u8>> {
    let Some(extension) = &source.extension else {
        return Ok(xml.to_vec());
    };
    let Some(id) = &extension.id else {
        return Ok(xml.to_vec());
    };
    let range = if !extension.other_content
        && source
            .list
            .as_ref()
            .is_some_and(|list| list.child_elements == 1 && !list.other_content)
    {
        source
            .list
            .as_ref()
            .map(|list| list.element.span.clone())
            .ok_or_else(|| invalid("change-tracking extension list disappeared"))?
    } else if !extension.other_content {
        extension.element.span.clone()
    } else {
        id.span.clone()
    };
    replace(xml, range, &[])
}

fn read(xml: &[u8], owner: Owner) -> Result<Source> {
    read_at(xml, owner, None)
}

fn read_at(xml: &[u8], owner: Owner, root_range: Option<&Range<usize>>) -> Result<Source> {
    if xml.is_empty() {
        return Err(invalid("change-tracking XML is empty"));
    }
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "change-tracking XML bytes",
            limit: MAX_XML_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::<Frame>::new();
    let mut nodes = 0usize;
    let mut parent = None;
    let mut list = None;
    let mut extension = None;
    let mut id_element = None;
    let mut id = None;
    let mut pml_namespace = None;
    let mut root_closed = false;
    let mut active = root_range.is_none();

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let namespace = namespace_kind(&resolved_namespace);
        let event = borrowed_event.into_owned();
        let end = position(&reader)?;
        if !active {
            let enters_root = matches!(event, Event::Start(_) | Event::Empty(_))
                && root_range
                    .as_ref()
                    .is_some_and(|range| range.start == start);
            if !enters_root {
                if matches!(event, Event::Eof) {
                    break;
                }
                continue;
            }
            active = true;
        }
        match event {
            Event::Start(element) => {
                bump_node(&mut nodes)?;
                let kind = classify(
                    namespace,
                    &element,
                    decoder,
                    owner,
                    &mut frames,
                    &mut id,
                    &mut pml_namespace,
                )?;
                if frames.len() >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "change-tracking XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                frames.push(Frame {
                    kind,
                    start,
                    child_elements: 0,
                    other_content: false,
                });
            },
            Event::Empty(element) => {
                bump_node(&mut nodes)?;
                let kind = classify(
                    namespace,
                    &element,
                    decoder,
                    owner,
                    &mut frames,
                    &mut id,
                    &mut pml_namespace,
                )?;
                finish(
                    &Frame {
                        kind,
                        start,
                        child_elements: 0,
                        other_content: false,
                    },
                    end,
                    end,
                    true,
                    &mut parent,
                    &mut list,
                    &mut extension,
                    &mut id_element,
                )?;
                if frames.is_empty() {
                    root_closed = true;
                    if let Some(range) = &root_range {
                        if end != range.end {
                            return Err(invalid(
                                "change-tracking root does not cover the selected XML",
                            ));
                        }
                        break;
                    }
                }
            },
            Event::End(_) => {
                let close_start = start;
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("change-tracking XML stack underflow"))?;
                finish(
                    &frame,
                    close_start,
                    end,
                    false,
                    &mut parent,
                    &mut list,
                    &mut extension,
                    &mut id_element,
                )?;
                if frames.is_empty() {
                    root_closed = true;
                    let expected_end = root_range.as_ref().map_or(xml.len(), |range| range.end);
                    if end != expected_end {
                        return Err(invalid(
                            "change-tracking root does not cover the complete XML",
                        ));
                    }
                    break;
                }
            },
            Event::Text(text) => {
                if !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                    && let Some(frame) = frames.last_mut()
                    && matches!(frame.kind, Kind::List | Kind::Extension { .. } | Kind::Id)
                {
                    frame.other_content = true;
                }
            },
            Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                if let Some(frame) = frames.last_mut()
                    && matches!(frame.kind, Kind::List | Kind::Extension { .. } | Kind::Id)
                {
                    frame.other_content = true;
                }
            },
            Event::Decl(_) if frames.is_empty() => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) if !frames.is_empty() => {
                return Err(invalid("change-tracking owner contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
    }
    if !root_closed || !frames.is_empty() {
        return Err(invalid("change-tracking XML is unterminated"));
    }
    Ok(Source {
        pml_namespace: pml_namespace
            .ok_or_else(|| invalid("change-tracking PresentationML profile is missing"))?,
        parent: parent.ok_or_else(|| invalid("change-tracking owner parent is missing"))?,
        list,
        extension,
        id,
    })
}

#[allow(
    clippy::too_many_arguments,
    reason = "bounded XML scanner passes its small typed output slots explicitly"
)]
fn classify(
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    owner: Owner,
    frames: &mut [Frame],
    id: &mut Option<Id>,
    pml_namespace: &mut Option<&'static str>,
) -> Result<Kind> {
    let local = element.local_name();
    if frames.is_empty() {
        let profile = pml_profile(namespace)
            .ok_or_else(|| invalid("change-tracking root is not PresentationML"))?;
        *pml_namespace = Some(profile);
        let valid = match owner {
            Owner::Slide => local.as_ref() == b"sld",
            Owner::Shape => matches!(
                local.as_ref(),
                b"sp" | b"pic" | b"cxnSp" | b"graphicFrame" | b"grpSp" | b"contentPart"
            ),
        };
        if !valid {
            return Err(invalid("unsupported change-tracking XML owner"));
        }
        return Ok(Kind::Root);
    }

    if let Some(parent) = frames.last_mut() {
        parent.child_elements = parent
            .child_elements
            .checked_add(1)
            .ok_or_else(|| invalid("change-tracking child count overflow"))?;
    }
    let depth = frames.len() + 1;
    let parent_kind = frames.last().map(|frame| frame.kind);
    if is_pml(namespace)
        && local.as_ref() == owner.parent()
        && matches!((owner, depth), (Owner::Slide, 2) | (Owner::Shape, 3))
    {
        return Ok(Kind::Parent);
    }
    if is_pml(namespace) && local.as_ref() == b"extLst" && matches!(parent_kind, Some(Kind::Parent))
    {
        return Ok(Kind::List);
    }
    if is_pml(namespace) && local.as_ref() == b"ext" && matches!(parent_kind, Some(Kind::List)) {
        let uri = unqualified_attribute_value(element, b"uri", decoder)?;
        let known = uri.as_deref() == Some(owner.uri());
        return Ok(Kind::Extension { known });
    }
    if matches!(parent_kind, Some(Kind::Extension { known: true }))
        && namespace == NamespaceKind::P14
        && local.as_ref() == owner.element()
    {
        if id.is_some() {
            return Err(invalid("duplicate change-tracking identifier element"));
        }
        *id = Some(parse_id(element, decoder)?);
        return Ok(Kind::Id);
    }
    if let Some(parent) = frames.last_mut()
        && matches!(
            parent.kind,
            Kind::List | Kind::Extension { known: true } | Kind::Id
        )
    {
        parent.other_content = true;
    }
    Ok(Kind::Other)
}

#[allow(
    clippy::too_many_arguments,
    reason = "bounded XML scanner publishes three explicit structural slots"
)]
fn finish(
    frame: &Frame,
    close_start: usize,
    end: usize,
    empty: bool,
    parent: &mut Option<Element>,
    list: &mut Option<ExtensionList>,
    extension: &mut Option<Extension>,
    id_element: &mut Option<Element>,
) -> Result<()> {
    let element = Element {
        span: frame.start..end,
        close_start,
        empty,
    };
    match frame.kind {
        Kind::Parent => {
            if parent.replace(element).is_some() {
                return Err(invalid("duplicate change-tracking owner parent"));
            }
        },
        Kind::List => {
            if list
                .replace(ExtensionList {
                    element,
                    child_elements: frame.child_elements,
                    other_content: frame.other_content,
                })
                .is_some()
            {
                return Err(invalid("duplicate change-tracking extension list"));
            }
        },
        Kind::Extension { known: true } => {
            if extension
                .replace(Extension {
                    element,
                    id: id_element.take(),
                    other_content: frame.other_content || frame.child_elements > 1,
                })
                .is_some()
            {
                return Err(invalid("duplicate change-tracking extension"));
            }
        },
        Kind::Id => {
            if frame.child_elements != 0 || frame.other_content {
                return Err(invalid(
                    "change-tracking identifier must not contain child content",
                ));
            }
            if id_element.replace(element).is_some() {
                return Err(invalid("duplicate change-tracking identifier"));
            }
        },
        Kind::Root | Kind::Extension { known: false } | Kind::Other => {},
    }
    Ok(())
}

fn parse_id(element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder) -> Result<Id> {
    let mut value = None;
    for attribute_result in element.attributes().with_checks(true) {
        let parsed_attribute = attribute_result.map_err(|error| Error::Xml(error.to_string()))?;
        let name = parsed_attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if name != b"val" {
            return Err(invalid(
                "change-tracking identifier has an unsupported attribute",
            ));
        }
        if value.is_some() {
            return Err(invalid("duplicate change-tracking val attribute"));
        }
        value = Some(
            parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    let text = value.ok_or_else(|| invalid("change-tracking identifier is missing val"))?;
    let parsed_value = text
        .parse::<u32>()
        .map_err(|_err| invalid("change-tracking identifier is not an unsigned integer"))?;
    Ok(Id::new(parsed_value))
}

fn pml_profile(namespace: NamespaceKind) -> Option<&'static str> {
    match namespace {
        NamespaceKind::Pml => Some("http://schemas.openxmlformats.org/presentationml/2006/main"),
        NamespaceKind::StrictPml => Some("http://purl.oclc.org/ooxml/presentationml/main"),
        NamespaceKind::P14 | NamespaceKind::Other => None,
    }
}

fn is_pml(namespace: NamespaceKind) -> bool {
    pml_profile(namespace).is_some()
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == PML => NamespaceKind::Pml,
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_PML => NamespaceKind::StrictPml,
        ResolveResult::Bound(Namespace(value)) if *value == P14 => NamespaceKind::P14,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn id_xml(owner: Owner, id: Id) -> String {
    format!(
        "<p14:{} xmlns:p14=\"{}\" val=\"{}\"/>",
        match owner {
            Owner::Slide => "creationId",
            Owner::Shape => "modId",
        },
        NAMESPACE,
        id.value()
    )
}

fn expand_empty(xml: &[u8], element: &Element, child: &[u8]) -> Result<Vec<u8>> {
    let raw = xml
        .get(element.span.clone())
        .ok_or_else(|| invalid("change-tracking empty element range is invalid"))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("change-tracking empty element has no closing slash"))?;
    let mut open_end = slash;
    while open_end > 0 && raw[open_end - 1].is_ascii_whitespace() {
        open_end -= 1;
    }
    let qname = qualified_name(raw)?;
    let mut value = Vec::new();
    let length = raw
        .len()
        .checked_sub(1)
        .and_then(|length| length.checked_add(child.len()))
        .and_then(|length| length.checked_add(qname.len()))
        .and_then(|length| length.checked_add(3))
        .ok_or_else(|| invalid("change-tracking element expansion overflow"))?;
    value
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "change-tracking XML expansion",
            source,
        })?;
    value.extend_from_slice(&raw[..open_end]);
    value.push(b'>');
    value.extend_from_slice(child);
    value.extend_from_slice(b"</");
    value.extend_from_slice(qname);
    value.push(b'>');
    replace(xml, element.span.clone(), &value)
}

fn qualified_name(element: &[u8]) -> Result<&[u8]> {
    let tail = element
        .strip_prefix(b"<")
        .ok_or_else(|| invalid("change-tracking element does not start with '<'"))?;
    let length = tail
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .ok_or_else(|| invalid("change-tracking element name is unterminated"))?;
    tail.get(..length)
        .filter(|name| !name.is_empty())
        .ok_or_else(|| invalid("change-tracking element name is empty"))
}

fn replace(xml: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > xml.len() {
        return Err(invalid("change-tracking replacement range is invalid"));
    }
    let length = xml
        .len()
        .checked_sub(range.len())
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("change-tracking replacement length overflow"))?;
    if length > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "change-tracking XML bytes",
            limit: MAX_XML_BYTES,
        });
    }
    let mut value = Vec::new();
    value
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "change-tracking XML replacement",
            source,
        })?;
    value.extend_from_slice(&xml[..range.start]);
    value.extend_from_slice(replacement);
    value.extend_from_slice(&xml[range.end..]);
    Ok(value)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("change-tracking XML position exceeds usize"))
}

fn bump_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("change-tracking XML node count overflow"))?;
    if *nodes > MAX_XML_NODES {
        return Err(Error::Limit {
            resource: "change-tracking XML nodes",
            limit: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
