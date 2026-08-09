//! Lossless XML codec for `PresentationML` zoom alternate-content entries.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::ops::Range;

use litchi_opc::{OpcPackage, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

use super::model::{
    ImageType, Item, Layout, Link, Owner, Percentage, Properties, Relationship, Section, Slide,
    Summary, Target, Unknown, Zoom,
};

pub(crate) const MAX_OWNER_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_OWNER_NODES: usize = 1_000_000;
pub(crate) const MAX_ZOOMS: usize = 100_000;
pub(crate) const MAX_RELATIONSHIP_ID_BYTES: usize = 4_096;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_UNKNOWN_BYTES: usize = 16 * 1024 * 1024;
const WRAPPER_NS: &str = "urn:litchi:pptx:zoom:wrapper";

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PML_STRICT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DML_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MAIN: &str = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
const SECTION_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom";
const SLIDE_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom";
const SUMMARY_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom";

const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const IMAGE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

#[derive(Debug, Clone)]
struct Attribute {
    namespace: String,
    local: String,
    value: String,
}

#[derive(Debug, Clone)]
struct Node {
    namespace: String,
    local: String,
    attrs: Vec<Attribute>,
    children: Vec<Node>,
    text: bool,
    range: Range<usize>,
    requires: Vec<String>,
}

#[derive(Debug)]
struct Dom {
    xml: Vec<u8>,
    root: Node,
}

impl Dom {
    fn slice(&self, range: &Range<usize>) -> Result<&[u8]> {
        self.xml
            .get(range.clone())
            .ok_or_else(|| Error::Invalid("zoom XML node span is outside its owner".into()))
    }
}

impl Node {
    fn is(&self, namespace: &str, local: &str) -> bool {
        self.namespace == namespace && self.local == local
    }

    fn attr(&self, namespace: &str, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|value| value.namespace == namespace && value.local == local)
            .map(|value| value.value.as_str())
    }

    fn unqualified(&self, local: &str) -> Option<&str> {
        self.attr("", local)
    }

    fn elements(&self) -> &[Node] {
        &self.children
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Other,
    Owner,
    Alternate,
}

#[derive(Debug)]
struct OwnerFrame {
    start: usize,
    namespace: String,
    local: String,
    kind: FrameKind,
    changes: Vec<(String, Option<String>)>,
    inherited: Option<HashMap<String, String>>,
}

#[derive(Debug)]
struct Profile {
    namespaces: Vec<(String, String)>,
    pml: String,
    dml: String,
    relationship: String,
}

/// Parse a complete slide or shape-tree XML owner while retaining exact spans.
pub(crate) fn read_owner(xml: &[u8]) -> Result<Owner> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "zoom owner XML",
            limit: MAX_OWNER_BYTES,
        });
    }

    let mut reader = NsReader::from_reader(xml);
    let mut namespaces = HashMap::<String, String>::new();
    let mut stack = Vec::<OwnerFrame>::new();
    let mut entries = Vec::<Zoom>::new();
    let mut spans = Vec::<(usize, usize)>::new();
    let mut insert_owner = None;
    let mut insert_at = None;
    let mut pml_namespace = None;
    let mut dml_namespace = None;
    let mut relationship_namespace = None;
    let mut all_namespaces = HashMap::<String, String>::new();
    let mut nodes = 0usize;
    let mut roots = 0usize;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("zoom XML offset exceeds usize".into()))?;
        let decoder = reader.decoder();
        let (namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            (resolved_namespace(namespace)?, event.into_owned())
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("zoom XML offset exceeds usize".into()))?;

        match event {
            Event::Start(element) => {
                if stack.is_empty() {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Invalid("zoom XML has multiple root elements".into()));
                    }
                }
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("zoom XML node count overflow".into()))?;
                if nodes > MAX_OWNER_NODES {
                    return Err(Error::Limit {
                        resource: "zoom owner XML nodes",
                        limit: MAX_OWNER_NODES,
                    });
                }
                let changes = apply_namespace_declarations(&element, decoder, &mut namespaces)?;
                all_namespaces.extend(
                    namespaces
                        .iter()
                        .map(|(key, value)| (key.clone(), value.clone())),
                );
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_err| Error::Invalid("zoom XML element name is not UTF-8".into()))?;
                let parent = stack.last().map(|value| value.kind);
                let kind = classify_owner_frame(&namespace, &local, parent);
                if kind == FrameKind::Owner {
                    if insert_owner.is_none() {
                        pml_namespace = Some(namespace.clone());
                        insert_owner = Some((start, 0, false));
                        insert_at = Some(0);
                    } else if pml_namespace.as_deref() != Some(namespace.as_str()) {
                        return Err(Error::Invalid(
                            "zoom owner contains mixed PresentationML namespaces".into(),
                        ));
                    }
                }
                if is_dml_namespace(&namespace) && dml_namespace.is_none() {
                    dml_namespace = Some(namespace.clone());
                }
                if relationship_namespace.is_none()
                    && namespaces
                        .values()
                        .any(|value| value == REL || value == REL_STRICT)
                {
                    relationship_namespace = namespaces
                        .values()
                        .find(|value| **value == REL || **value == REL_STRICT)
                        .cloned();
                }
                let inherited = if kind == FrameKind::Alternate && parent == Some(FrameKind::Owner)
                {
                    Some(namespaces.clone())
                } else {
                    None
                };
                stack.push(OwnerFrame {
                    start,
                    namespace,
                    local,
                    kind,
                    changes,
                    inherited,
                });
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Invalid("zoom XML has multiple root elements".into()));
                    }
                }
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("zoom XML node count overflow".into()))?;
                if nodes > MAX_OWNER_NODES {
                    return Err(Error::Limit {
                        resource: "zoom owner XML nodes",
                        limit: MAX_OWNER_NODES,
                    });
                }
                let changes = apply_namespace_declarations(&element, decoder, &mut namespaces)?;
                all_namespaces.extend(
                    namespaces
                        .iter()
                        .map(|(key, value)| (key.clone(), value.clone())),
                );
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_err| Error::Invalid("zoom XML element name is not UTF-8".into()))?;
                let parent = stack.last().map(|value| value.kind);
                let kind = classify_owner_frame(&namespace, &local, parent);
                if kind == FrameKind::Owner {
                    if insert_owner.is_none() {
                        pml_namespace = Some(namespace.clone());
                        insert_owner = Some((start, end, true));
                        insert_at = Some(end);
                    } else if pml_namespace.as_deref() != Some(namespace.as_str()) {
                        return Err(Error::Invalid(
                            "zoom owner contains mixed PresentationML namespaces".into(),
                        ));
                    }
                }
                if is_dml_namespace(&namespace) && dml_namespace.is_none() {
                    dml_namespace = Some(namespace.clone());
                }
                if relationship_namespace.is_none()
                    && namespaces
                        .values()
                        .any(|value| value == REL || value == REL_STRICT)
                {
                    relationship_namespace = namespaces
                        .values()
                        .find(|value| **value == REL || **value == REL_STRICT)
                        .cloned();
                }
                if kind == FrameKind::Alternate && parent == Some(FrameKind::Owner) {
                    return Err(Error::Invalid(
                        "zoom AlternateContent must contain Choice and Fallback children".into(),
                    ));
                }
                restore_namespace_declarations(&mut namespaces, changes);
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| Error::Invalid("unexpected zoom XML closing element".into()))?;
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_err| Error::Invalid("zoom XML element name is not UTF-8".into()))?;
                if frame.namespace != namespace || frame.local != local {
                    return Err(Error::Invalid("mismatched zoom XML closing element".into()));
                }
                if frame.kind == FrameKind::Owner
                    && insert_owner.is_some_and(|value| value.1 == 0 && value.0 == frame.start)
                {
                    insert_owner = Some((frame.start, end, false));
                    insert_at = Some(start);
                }
                if frame.kind == FrameKind::Alternate
                    && stack
                        .last()
                        .is_some_and(|value| value.kind == FrameKind::Owner)
                {
                    let inherited = frame.inherited.as_ref().ok_or_else(|| {
                        Error::Invalid("zoom AlternateContent has no namespace context".into())
                    })?;
                    let raw = xml.get(frame.start..end).ok_or_else(|| {
                        Error::Invalid("zoom AlternateContent span is outside owner XML".into())
                    })?;
                    if spans.len() >= MAX_ZOOMS {
                        return Err(Error::Limit {
                            resource: "zoom entry count",
                            limit: MAX_ZOOMS,
                        });
                    }
                    let value = parse_alternate(raw, inherited)?;
                    spans.push((frame.start, end));
                    entries.push(value);
                }
                restore_namespace_declarations(&mut namespaces, frame.changes);
            },
            Event::Text(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) && stack.is_empty() {
                    return Err(Error::Invalid(
                        "non-whitespace text is outside zoom XML root".into(),
                    ));
                }
            },
            Event::CData(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) && stack.is_empty() {
                    return Err(Error::Invalid(
                        "non-whitespace text is outside zoom XML root".into(),
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "zoom XML must not contain DTDs or processing instructions".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !stack.is_empty() {
        return Err(Error::Invalid("unterminated zoom XML owner".into()));
    }
    if roots != 1 {
        return Err(Error::Invalid(
            "zoom XML must contain exactly one root element".into(),
        ));
    }
    let insert_owner = insert_owner
        .ok_or_else(|| Error::Invalid("zoom owner has no PresentationML spTree or grpSp".into()))?;
    let profile = Profile {
        namespaces: all_namespaces.into_iter().collect(),
        pml: pml_namespace.unwrap_or_else(|| PML.to_owned()),
        dml: dml_namespace.unwrap_or_else(|| DML.to_owned()),
        relationship: relationship_namespace.unwrap_or_else(|| REL.to_owned()),
    };
    validate_entries(&entries)?;
    Ok(Owner {
        xml: xml.to_vec(),
        base_xml: xml.to_vec(),
        entries,
        spans,
        insert_at: insert_at.unwrap_or(insert_owner.1),
        insert_owner: Some(insert_owner),
        namespaces: profile.namespaces,
        pml_namespace: profile.pml,
        dml_namespace: profile.dml,
        relationship_namespace: profile.relationship,
    })
}

fn classify_owner_frame(namespace: &str, local: &str, parent: Option<FrameKind>) -> FrameKind {
    if is_pml_namespace(namespace) && matches!(local, "spTree" | "grpSp") {
        return FrameKind::Owner;
    }
    if namespace == MC && local == "AlternateContent" && parent == Some(FrameKind::Owner) {
        return FrameKind::Alternate;
    }
    FrameKind::Other
}

fn is_pml_namespace(namespace: &str) -> bool {
    namespace == PML || namespace == PML_STRICT
}

fn is_dml_namespace(namespace: &str) -> bool {
    namespace == DML || namespace == DML_STRICT
}

fn resolved_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(namespace)) => String::from_utf8(namespace.to_vec())
            .map_err(|_err| Error::Invalid("zoom XML namespace is not UTF-8".into())),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(Error::Invalid(format!(
            "zoom XML prefix '{}' is unbound",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespaces: &mut HashMap<String, String>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            Some(String::new())
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            Some(
                String::from_utf8(prefix.to_vec())
                    .map_err(|_err| Error::Invalid("zoom namespace prefix is not UTF-8".into()))?,
            )
        } else {
            None
        };
        let Some(prefix) = prefix else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if value.is_empty() {
            return Err(Error::Invalid("zoom namespace URI cannot be empty".into()));
        }
        changes.push((prefix.clone(), namespaces.insert(prefix, value)));
    }
    Ok(changes)
}

fn restore_namespace_declarations(
    namespaces: &mut HashMap<String, String>,
    changes: Vec<(String, Option<String>)>,
) {
    for (prefix, old) in changes.into_iter().rev() {
        match old {
            Some(value) => {
                namespaces.insert(prefix, value);
            },
            None => {
                namespaces.remove(&prefix);
            },
        }
    }
}

fn parse_alternate(raw: &[u8], inherited: &HashMap<String, String>) -> Result<Zoom> {
    if raw.len() > MAX_UNKNOWN_BYTES {
        return Err(Error::Limit {
            resource: "zoom alternate-content XML",
            limit: MAX_UNKNOWN_BYTES,
        });
    }
    let dom = parse_dom(raw, inherited)?;
    let alternate = dom
        .root
        .children
        .first()
        .ok_or_else(|| Error::Invalid("zoom alternate-content is empty".into()))?;
    expect(alternate, MC, "AlternateContent")?;

    let mut choices = Vec::new();
    let mut fallback = None;
    let mut saw_fallback = false;
    for child in alternate.elements() {
        if child.is(MC, "Choice") {
            if saw_fallback {
                return Err(Error::Invalid(
                    "zoom Choice appears after its Fallback".into(),
                ));
            }
            if child.requires.is_empty() {
                return Err(Error::Invalid("zoom Choice is missing mc:Requires".into()));
            }
            choices.push(child);
        } else if child.is(MC, "Fallback") {
            if fallback.is_some() {
                return Err(Error::Invalid(
                    "zoom AlternateContent has duplicate Fallback".into(),
                ));
            }
            saw_fallback = true;
            fallback = Some(child);
        } else {
            return Err(Error::Invalid(
                "zoom AlternateContent contains an unsupported direct child".into(),
            ));
        }
    }
    let fallback = fallback
        .ok_or_else(|| Error::Invalid("zoom AlternateContent is missing Fallback".into()))?;
    if choices.is_empty() {
        return Err(Error::Invalid(
            "zoom AlternateContent is missing Choice".into(),
        ));
    }

    let mut known = None;
    let mut unknown_xml = Vec::new();
    for (position, choice) in choices.iter().enumerate() {
        let payload = choice.elements();
        let kind = payload.first().and_then(|value| {
            if value.is(SECTION_NS, "sectionZm") {
                Some(0u8)
            } else if value.is(SLIDE_NS, "sldZm") {
                Some(1u8)
            } else if value.is(SUMMARY_NS, "summaryZm") {
                Some(2u8)
            } else {
                None
            }
        });
        let Some(kind) = kind else {
            unknown_xml.push(dom.slice(&choice.range)?.to_vec());
            continue;
        };
        if payload.len() != 1 {
            return Err(Error::Invalid(
                "typed zoom Choice must contain exactly one payload".into(),
            ));
        }
        let expected = match kind {
            0 => SECTION_NS,
            1 => SLIDE_NS,
            _ => SUMMARY_NS,
        };
        if !choice.requires.iter().any(|value| value == expected) {
            return Err(Error::Invalid(
                "zoom Choice Requires does not name its payload namespace".into(),
            ));
        }
        if known.replace((kind, position, &payload[0])).is_some() {
            return Err(Error::Invalid(
                "zoom AlternateContent has multiple typed choices".into(),
            ));
        }
    }
    let Some((kind, _position, payload)) = known else {
        return Ok(Zoom::Unknown(Unknown { xml: raw.to_vec() }));
    };
    let fallback_children = fallback.elements();
    if fallback_children.len() != 1 {
        return Err(Error::Invalid(
            "zoom Fallback must contain exactly one shape".into(),
        ));
    }
    let fallback_expected = if kind == 2 { "grpSp" } else { "pic" };
    if !fallback_children[0].is(PML, fallback_expected)
        && !fallback_children[0].is(PML_STRICT, fallback_expected)
    {
        return Err(Error::Invalid(format!(
            "zoom Fallback must contain p:{fallback_expected}"
        )));
    }
    let fallback_xml = dom.slice(&fallback_children[0].range)?.to_vec();
    match kind {
        0 => parse_section(&dom, payload, fallback_xml, unknown_xml),
        1 => parse_slide(&dom, payload, fallback_xml, unknown_xml),
        _ => parse_summary(&dom, payload, fallback_xml, unknown_xml),
    }
}

fn parse_section(
    dom: &Dom,
    node: &Node,
    fallback_xml: Vec<u8>,
    unknown_xml: Vec<Vec<u8>>,
) -> Result<Zoom> {
    expect(node, SECTION_NS, "sectionZm")?;
    let children = node.elements();
    if children.len() != 1 && children.len() != 2 {
        return Err(Error::Invalid(
            "sectionZm requires sectionZmObj and optional extLst".into(),
        ));
    }
    let object = children
        .first()
        .ok_or_else(|| Error::Invalid("sectionZm is missing sectionZmObj".into()))?;
    expect(object, SECTION_NS, "sectionZmObj")?;
    let section_id = required(object, "", "sectionId")?.to_owned();
    validate_guid(&section_id)?;
    let (properties, extension_xml) = parse_object(dom, object, SECTION_NS)?;
    if children.len() == 2 {
        expect_pml(&children[1], "extLst")?;
    }
    Ok(Zoom::Section(Section {
        section_id,
        properties,
        fallback_xml,
        extension_xml,
        unknown_xml,
    }))
}

fn parse_slide(
    dom: &Dom,
    node: &Node,
    fallback_xml: Vec<u8>,
    unknown_xml: Vec<Vec<u8>>,
) -> Result<Zoom> {
    expect(node, SLIDE_NS, "sldZm")?;
    let children = node.elements();
    if children.len() != 1 && children.len() != 2 {
        return Err(Error::Invalid(
            "sldZm requires sldZmObj and optional extLst".into(),
        ));
    }
    let object = children
        .first()
        .ok_or_else(|| Error::Invalid("sldZm is missing sldZmObj".into()))?;
    expect(object, SLIDE_NS, "sldZmObj")?;
    let slide_id = required(object, "", "sldId")?
        .parse::<u32>()
        .map_err(|_err| Error::Invalid("slide zoom has an invalid sldId".into()))?;
    validate_slide_id(slide_id)?;
    let creation_id = object
        .unqualified("cId")
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_err| Error::Invalid("slide zoom has an invalid cId".into()))
        })
        .transpose()?;
    let (properties, extension_xml) = parse_object(dom, object, SLIDE_NS)?;
    if children.len() == 2 {
        expect_pml(&children[1], "extLst")?;
    }
    Ok(Zoom::Slide(Slide {
        slide_id,
        creation_id,
        properties,
        fallback_xml,
        extension_xml,
        unknown_xml,
    }))
}

fn parse_summary(
    dom: &Dom,
    node: &Node,
    fallback_xml: Vec<u8>,
    unknown_xml: Vec<Vec<u8>>,
) -> Result<Zoom> {
    expect(node, SUMMARY_NS, "summaryZm")?;
    let children = node.elements();
    let mut items = Vec::new();
    let mut cursor = 0usize;
    while let Some(item) = children.get(cursor) {
        if !item.is(SUMMARY_NS, "summaryZmObj") {
            break;
        }
        let section_id = required(item, "", "sectionId")?.to_owned();
        validate_guid(&section_id)?;
        let title = item.unqualified("title").unwrap_or_default().to_owned();
        let description = item.unqualified("descr").unwrap_or_default().to_owned();
        validate_string(&title, "summary zoom title")?;
        validate_string(&description, "summary zoom description")?;
        let offset_x = parse_percentage(item.unqualified("offsetFactorX"), 0, "offsetFactorX")?;
        let offset_y = parse_percentage(item.unqualified("offsetFactorY"), 0, "offsetFactorY")?;
        let scale_x = parse_percentage(item.unqualified("scaleFactorX"), 100_000, "scaleFactorX")?;
        let scale_y = parse_percentage(item.unqualified("scaleFactorY"), 100_000, "scaleFactorY")?;
        let (properties, extension_xml) = parse_object(dom, item, SUMMARY_NS)?;
        items.push(Item {
            section_id,
            title,
            description,
            offset_x,
            offset_y,
            scale_x,
            scale_y,
            properties,
            extension_xml,
            unknown_xml: Vec::new(),
        });
        cursor += 1;
    }
    let layout = match children.get(cursor) {
        Some(value) if value.is(SUMMARY_NS, "gridLayout") => {
            if !value.attrs.is_empty() || !value.children.is_empty() || value.text {
                return Err(Error::Invalid("gridLayout must be empty".into()));
            }
            Layout::Grid
        },
        Some(value) if value.is(SUMMARY_NS, "fixedLayout") => {
            if !value.attrs.is_empty() || !value.children.is_empty() || value.text {
                return Err(Error::Invalid("fixedLayout must be empty".into()));
            }
            Layout::Fixed
        },
        _ => {
            return Err(Error::Invalid(
                "summaryZm requires one gridLayout or fixedLayout".into(),
            ));
        },
    };
    cursor += 1;
    let extension_xml = if let Some(extension) = children.get(cursor) {
        expect_pml(extension, "extLst")?;
        cursor += 1;
        Some(dom.slice(&extension.range)?.to_vec())
    } else {
        None
    };
    if cursor != children.len() {
        return Err(Error::Invalid(
            "summaryZm contains an unsupported child after its layout".into(),
        ));
    }
    Ok(Zoom::Summary(Summary {
        items,
        layout,
        fallback_xml,
        extension_xml,
        unknown_xml,
    }))
}

fn parse_object(
    dom: &Dom,
    object: &Node,
    object_namespace: &str,
) -> Result<(Properties, Option<Vec<u8>>)> {
    let children = object.elements();
    if children.len() != 1 && children.len() != 2 {
        return Err(Error::Invalid(
            "zoom object requires zmPr and optional extLst".into(),
        ));
    }
    let properties = children
        .first()
        .ok_or_else(|| Error::Invalid("zoom object is missing zmPr".into()))?;
    expect(properties, MAIN, "zmPr")?;
    let properties = parse_properties(dom, properties)?;
    if children.len() == 2 {
        expect_pml(&children[1], "extLst")?;
        if children[1].namespace != PML && children[1].namespace != PML_STRICT {
            return Err(Error::Invalid(format!(
                "{object_namespace} object extension has an invalid namespace"
            )));
        }
    }
    Ok((
        properties,
        children
            .get(1)
            .map(|value| dom.slice(&value.range).map(<[u8]>::to_vec))
            .transpose()?,
    ))
}

fn parse_properties(dom: &Dom, node: &Node) -> Result<Properties> {
    let id = required(node, "", "id")?.to_owned();
    validate_guid(&id)?;
    let return_to_parent = parse_bool(node.unqualified("returnToParent"), true, "returnToParent")?;
    let image_type = match node.unqualified("imageType").unwrap_or("preview") {
        "preview" => ImageType::Preview,
        "cover" => ImageType::Cover,
        value => {
            return Err(Error::Invalid(format!(
                "zoom imageType has an unsupported value '{value}'"
            )));
        },
    };
    let transition = node
        .unqualified("transitionDur")
        .map(crate::time::Offset::parse)
        .transpose()
        .map_err(|error| Error::Invalid(format!("invalid zoom transitionDur: {error}")))?;
    let show_background = parse_bool(node.unqualified("showBg"), true, "showBg")?;
    let children = node.elements();
    if children.len() != 2 {
        return Err(Error::Invalid(
            "zoom zmPr requires blipFill followed by spPr".into(),
        ));
    }
    if !is_property_node(&children[0], "blipFill") || !is_property_node(&children[1], "spPr") {
        return Err(Error::Invalid(
            "zoom zmPr children are not blipFill and spPr".into(),
        ));
    }
    let image = parse_blip_relationship_node(&children[0])?;
    Ok(Properties {
        id,
        return_to_parent,
        image_type,
        transition,
        show_background,
        blip_fill_xml: dom.slice(&children[0].range)?.to_vec(),
        shape_properties_xml: dom.slice(&children[1].range)?.to_vec(),
        image,
    })
}

fn parse_dom(raw: &[u8], inherited: &HashMap<String, String>) -> Result<Dom> {
    let mut xml = Vec::with_capacity(raw.len().saturating_add(512));
    xml.extend_from_slice(b"<z:root xmlns:z=\"urn:litchi:pptx:zoom:wrapper\"");
    let mut declarations = inherited.iter().collect::<Vec<_>>();
    declarations.sort_by_key(|(left, _)| *left);
    for (prefix, uri) in declarations {
        if prefix == "xml" || prefix == "xmlns" || prefix == "z" {
            continue;
        }
        let mut writer = StringWriter { bytes: &mut xml };
        if prefix.is_empty() {
            write!(&mut writer, " xmlns=\"{}\"", escape_xml(uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        } else {
            write!(&mut writer, " xmlns:{}=\"{}\"", prefix, escape_xml(uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
    }
    xml.extend_from_slice(b">");
    xml.extend_from_slice(raw);
    xml.extend_from_slice(b"</z:root>");
    let mut reader = NsReader::from_reader(xml.as_slice());
    let mut namespaces = inherited.clone();
    namespaces.insert("z".into(), WRAPPER_NS.into());
    let mut frames = Vec::<DomFrame>::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("zoom DOM offset exceeds usize".into()))?;
        let decoder = reader.decoder();
        let (namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            (resolved_namespace(namespace)?, event.into_owned())
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("zoom DOM offset exceeds usize".into()))?;
        match event {
            Event::Start(element) => {
                nodes += 1;
                if nodes > MAX_OWNER_NODES {
                    return Err(Error::Limit {
                        resource: "zoom alternate-content nodes",
                        limit: MAX_OWNER_NODES,
                    });
                }
                let changes = apply_namespace_declarations(&element, decoder, &mut namespaces)?;
                let node = dom_frame(&element, decoder, namespace, start, changes, &namespaces)?;
                frames.push(node);
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_OWNER_NODES {
                    return Err(Error::Limit {
                        resource: "zoom alternate-content nodes",
                        limit: MAX_OWNER_NODES,
                    });
                }
                let changes = apply_namespace_declarations(&element, decoder, &mut namespaces)?;
                let frame = dom_frame(&element, decoder, namespace, start, changes, &namespaces)?;
                let changes = frame.changes.clone();
                let node = frame.finish(start..end);
                attach_node(&mut frames, &mut root, node);
                restore_namespace_declarations(&mut namespaces, changes);
            },
            Event::End(element) => {
                let frame = frames.pop().ok_or_else(|| {
                    Error::Invalid("zoom DOM has an unexpected closing tag".into())
                })?;
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_err| Error::Invalid("zoom DOM element name is not UTF-8".into()))?;
                if frame.namespace != namespace || frame.local != local {
                    return Err(Error::Invalid(
                        "zoom DOM has mismatched closing tags".into(),
                    ));
                }
                let frame_start = frame.start;
                let changes = frame.changes.clone();
                let node = frame.finish(frame_start..end);
                attach_node(&mut frames, &mut root, node);
                restore_namespace_declarations(&mut namespaces, changes);
            },
            Event::Text(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace)
                    && let Some(frame) = frames.last_mut()
                {
                    frame.text = true;
                }
            },
            Event::CData(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace)
                    && let Some(frame) = frames.last_mut()
                {
                    frame.text = true;
                }
            },
            Event::GeneralRef(_) => {
                if let Some(frame) = frames.last_mut() {
                    frame.text = true;
                }
            },
            Event::Comment(_) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "zoom DOM must not contain DTDs or processing instructions".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !frames.is_empty() {
        return Err(Error::Invalid("zoom DOM is unterminated".into()));
    }
    let root = root.ok_or_else(|| Error::Invalid("zoom DOM has no root".into()))?;
    if !root.is(WRAPPER_NS, "root") {
        return Err(Error::Invalid("zoom DOM wrapper root is missing".into()));
    }
    Ok(Dom { xml, root })
}

#[derive(Debug)]
struct DomFrame {
    start: usize,
    namespace: String,
    local: String,
    attrs: Vec<Attribute>,
    requires: Vec<String>,
    children: Vec<Node>,
    text: bool,
    changes: Vec<(String, Option<String>)>,
}

impl DomFrame {
    fn finish(self, range: Range<usize>) -> Node {
        Node {
            namespace: self.namespace,
            local: self.local,
            attrs: self.attrs,
            children: self.children,
            text: self.text,
            range,
            requires: self.requires,
        }
    }
}

fn dom_frame(
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespace: String,
    start: usize,
    changes: Vec<(String, Option<String>)>,
    namespaces: &HashMap<String, String>,
) -> Result<DomFrame> {
    let local = String::from_utf8(element.local_name().as_ref().to_vec())
        .map_err(|_err| Error::Invalid("zoom DOM element name is not UTF-8".into()))?;
    let mut attrs = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local_name) = resolve_attribute_lexically(attribute.key, namespaces)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        attrs.push(Attribute {
            namespace: resolved,
            local: local_name,
            value,
        });
    }
    let requires = attrs
        .iter()
        .find(|value| value.namespace.is_empty() && value.local == "Requires")
        .map(|value| {
            value
                .value
                .split_whitespace()
                .filter_map(|prefix| namespaces.get(prefix).cloned())
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Ok(DomFrame {
        start,
        namespace,
        local,
        attrs,
        requires,
        children: Vec::new(),
        text: false,
        changes,
    })
}

fn attach_node(frames: &mut Vec<DomFrame>, root: &mut Option<Node>, node: Node) {
    if let Some(parent) = frames.last_mut() {
        parent.children.push(node);
    } else {
        *root = Some(node);
    }
}

fn resolve_attribute_lexically(
    name: QName<'_>,
    namespaces: &HashMap<String, String>,
) -> Result<(String, String)> {
    let raw = name.as_ref();
    let (prefix, local) = match raw.iter().position(|byte| *byte == b':') {
        Some(index) => (&raw[..index], &raw[index + 1..]),
        None => (&[][..], raw),
    };
    let local = String::from_utf8(local.to_vec())
        .map_err(|_err| Error::Invalid("zoom DOM attribute name is not UTF-8".into()))?;
    if prefix.is_empty() {
        return Ok((String::new(), local));
    }
    let prefix = String::from_utf8(prefix.to_vec())
        .map_err(|_err| Error::Invalid("zoom DOM attribute prefix is not UTF-8".into()))?;
    let namespace = namespaces
        .get(&prefix)
        .cloned()
        .or_else(|| (prefix == "xml").then_some(XML.to_owned()))
        .ok_or_else(|| {
            Error::Invalid(format!("zoom DOM attribute prefix '{prefix}' is unbound"))
        })?;
    Ok((namespace, local))
}

fn is_dml_node(node: &Node, local: &str) -> bool {
    (node.is(DML, local) || node.is(DML_STRICT, local)) && !node.text
}

fn is_property_node(node: &Node, local: &str) -> bool {
    (node.is(MAIN, local) || is_dml_node(node, local)) && !node.text
}

fn is_pml_node(node: &Node, local: &str) -> bool {
    node.is(PML, local) || node.is(PML_STRICT, local)
}

fn expect(node: &Node, namespace: &str, local: &str) -> Result<()> {
    if !node.is(namespace, local) {
        return Err(Error::Invalid(format!(
            "expected {{{namespace}}}{local}, found {{{}}}{}",
            node.namespace, node.local
        )));
    }
    Ok(())
}

fn expect_pml(node: &Node, local: &str) -> Result<()> {
    if !is_pml_node(node, local) {
        return Err(Error::Invalid(format!("expected PresentationML {local}")));
    }
    Ok(())
}

fn required<'a>(node: &'a Node, namespace: &str, local: &str) -> Result<&'a str> {
    node.attr(namespace, local)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid(format!("zoom element is missing required {local}")))
}

fn parse_bool(value: Option<&str>, default: bool, field: &str) -> Result<bool> {
    match value.unwrap_or(if default { "true" } else { "false" }) {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        value => Err(Error::Invalid(format!(
            "invalid zoom {field} value '{value}'"
        ))),
    }
}

fn parse_percentage(value: Option<&str>, default: i32, field: &str) -> Result<Percentage> {
    let value = value.unwrap_or({
        // This branch returns a string only for the two schema defaults and
        // avoids allocating a temporary during ordinary parsing.
        if default == 0 { "0" } else { "100000" }
    });
    let value = value
        .parse::<i32>()
        .map_err(|_err| Error::Invalid(format!("invalid zoom {field} value '{value}'")))?;
    Ok(Percentage::new(value))
}

fn validate_string(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES || value.chars().any(|character| character == '\0') {
        return Err(Error::Invalid(format!(
            "{field} exceeds its bounded XML domain"
        )));
    }
    Ok(())
}

pub(crate) fn validate_guid(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let valid = bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && [9, 14, 19, 24].iter().all(|index| bytes[*index] == b'-')
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || byte.is_ascii_hexdigit()
        });
    if valid {
        Ok(())
    } else {
        Err(Error::Invalid("zoom GUID is not in ST_Guid form".into()))
    }
}

pub(crate) fn validate_slide_id(value: u32) -> Result<()> {
    if value < 256 {
        return Err(Error::Invalid(
            "zoom sldId is below the ST_SlideId minimum".into(),
        ));
    }
    Ok(())
}

pub(crate) fn parse_blip_relationship(xml: &[u8]) -> Result<Option<Relationship>> {
    let mut inherited = HashMap::new();
    inherited.insert("p166".into(), MAIN.into());
    inherited.insert("a".into(), DML.into());
    inherited.insert("r".into(), REL.into());
    let dom = parse_dom(xml, &inherited)?;
    let root = dom
        .root
        .children
        .first()
        .ok_or_else(|| Error::Invalid("zoom blipFill XML has no root".into()))?;
    if !is_property_node(root, "blipFill") {
        return Err(Error::Invalid(
            "zoom image payload is not a blipFill".into(),
        ));
    }
    parse_blip_relationship_node(root)
}

fn parse_blip_relationship_node(node: &Node) -> Result<Option<Relationship>> {
    let mut found = None;
    walk(node, &mut |value| {
        if !is_dml_node(value, "blip") {
            return Ok(());
        }
        let embed = value
            .attr(REL, "embed")
            .or_else(|| value.attr(REL_STRICT, "embed"));
        let link = value
            .attr(REL, "link")
            .or_else(|| value.attr(REL_STRICT, "link"));
        if embed.is_some() && link.is_some() {
            return Err(Error::Invalid(
                "zoom a:blip cannot contain both r:embed and r:link".into(),
            ));
        }
        let (id, kind) = match (embed, link) {
            (Some(id), None) => (id, Link::Embed),
            (None, Some(id)) => (id, Link::External),
            (None, None) => return Ok(()),
            (Some(_), Some(_)) => unreachable!(),
        };
        let relationship = Relationship::new(id.to_owned(), kind)?;
        if let Some(previous) = &found
            && previous != &relationship
        {
            return Err(Error::Invalid(
                "zoom blipFill contains conflicting image relationships".into(),
            ));
        }
        found = Some(relationship);
        Ok(())
    })?;
    Ok(found)
}

fn walk(node: &Node, visitor: &mut impl FnMut(&Node) -> Result<()>) -> Result<()> {
    visitor(node)?;
    for child in node.elements() {
        walk(child, visitor)?;
    }
    Ok(())
}

fn validate_entries(entries: &[Zoom]) -> Result<()> {
    if entries.len() > MAX_ZOOMS {
        return Err(Error::Limit {
            resource: "zoom entry count",
            limit: MAX_ZOOMS,
        });
    }
    let mut ids = HashSet::new();
    for value in entries {
        validate_zoom(value)?;
        for id in property_ids(value) {
            if !ids.insert(id.to_owned()) {
                return Err(Error::Invalid(
                    "zoom owner contains duplicate object GUIDs".into(),
                ));
            }
        }
    }
    Ok(())
}

fn property_ids(value: &Zoom) -> Vec<&str> {
    match value {
        Zoom::Section(value) => vec![value.properties.id.as_str()],
        Zoom::Slide(value) => vec![value.properties.id.as_str()],
        Zoom::Summary(value) => value
            .items
            .iter()
            .map(|item| item.properties.id.as_str())
            .collect(),
        Zoom::Unknown(_) => Vec::new(),
    }
}

fn validate_zoom(value: &Zoom) -> Result<()> {
    match value {
        Zoom::Section(value) => {
            validate_guid(&value.section_id)?;
            validate_properties(&value.properties)?;
            validate_fallback(&value.fallback_xml, "pic")?;
            validate_optional_extension(value.extension_xml.as_deref())?;
            validate_unknown(value.unknown_xml.iter())?;
        },
        Zoom::Slide(value) => {
            validate_slide_id(value.slide_id)?;
            validate_properties(&value.properties)?;
            validate_fallback(&value.fallback_xml, "pic")?;
            validate_optional_extension(value.extension_xml.as_deref())?;
            validate_unknown(value.unknown_xml.iter())?;
        },
        Zoom::Summary(value) => {
            validate_fallback(&value.fallback_xml, "grpSp")?;
            validate_optional_extension(value.extension_xml.as_deref())?;
            validate_unknown(value.unknown_xml.iter())?;
            let mut sections = HashSet::new();
            let mut ids = HashSet::new();
            for item in &value.items {
                validate_guid(&item.section_id)?;
                validate_string(&item.title, "summary zoom title")?;
                validate_string(&item.description, "summary zoom description")?;
                validate_properties(&item.properties)?;
                validate_optional_extension(item.extension_xml.as_deref())?;
                validate_unknown(item.unknown_xml.iter())?;
                if !sections.insert(&item.section_id) || !ids.insert(&item.properties.id) {
                    return Err(Error::Invalid(
                        "summary zoom contains duplicate section or object IDs".into(),
                    ));
                }
            }
        },
        Zoom::Unknown(value) => {
            if value.xml.len() > MAX_UNKNOWN_BYTES {
                return Err(Error::Limit {
                    resource: "unknown zoom XML",
                    limit: MAX_UNKNOWN_BYTES,
                });
            }
            let mut inherited = HashMap::new();
            inherited.insert("mc".into(), MC.into());
            let dom = parse_dom(&value.xml, &inherited)?;
            let alternate = dom
                .root
                .children
                .first()
                .ok_or_else(|| Error::Invalid("unknown zoom XML has no root".into()))?;
            expect(alternate, MC, "AlternateContent")?;
        },
    }
    Ok(())
}

fn validate_properties(value: &Properties) -> Result<()> {
    validate_guid(&value.id)?;
    if value.blip_fill_xml.len() > MAX_UNKNOWN_BYTES
        || value.shape_properties_xml.len() > MAX_UNKNOWN_BYTES
    {
        return Err(Error::Limit {
            resource: "zoom DrawingML property XML",
            limit: MAX_UNKNOWN_BYTES,
        });
    }
    let parsed = parse_blip_relationship(&value.blip_fill_xml)?;
    if parsed != value.image {
        return Err(Error::Invalid(
            "zoom image relationship metadata diverges from blipFill XML".into(),
        ));
    }
    let mut inherited = HashMap::new();
    inherited.insert("p166".into(), MAIN.into());
    inherited.insert("a".into(), DML.into());
    let dom = parse_dom(&value.shape_properties_xml, &inherited)?;
    let root = dom
        .root
        .children
        .first()
        .ok_or_else(|| Error::Invalid("zoom spPr XML has no root".into()))?;
    if !is_property_node(root, "spPr") {
        return Err(Error::Invalid(
            "zoom shape-properties XML is not a spPr".into(),
        ));
    }
    Ok(())
}

fn validate_fallback(xml: &[u8], local: &str) -> Result<()> {
    if xml.len() > MAX_UNKNOWN_BYTES {
        return Err(Error::Limit {
            resource: "zoom fallback XML",
            limit: MAX_UNKNOWN_BYTES,
        });
    }
    let mut inherited = HashMap::new();
    inherited.insert("p".into(), PML.into());
    let dom = parse_dom(xml, &inherited)?;
    let node = dom
        .root
        .children
        .first()
        .ok_or_else(|| Error::Invalid("zoom fallback XML has no root".into()))?;
    if !is_pml_node(node, local) {
        return Err(Error::Invalid(format!("zoom fallback is not p:{local}")));
    }
    Ok(())
}

fn validate_optional_extension(xml: Option<&[u8]>) -> Result<()> {
    let Some(xml) = xml else {
        return Ok(());
    };
    if xml.len() > MAX_UNKNOWN_BYTES {
        return Err(Error::Limit {
            resource: "zoom extension XML",
            limit: MAX_UNKNOWN_BYTES,
        });
    }
    let mut inherited = HashMap::new();
    inherited.insert("p".into(), PML.into());
    let dom = parse_dom(xml, &inherited)?;
    let node = dom
        .root
        .children
        .first()
        .ok_or_else(|| Error::Invalid("zoom extension XML has no root".into()))?;
    expect_pml(node, "extLst")
}

fn validate_unknown<'a>(values: impl Iterator<Item = &'a Vec<u8>>) -> Result<()> {
    for value in values {
        if value.len() > MAX_UNKNOWN_BYTES {
            return Err(Error::Limit {
                resource: "zoom unknown XML",
                limit: MAX_UNKNOWN_BYTES,
            });
        }
    }
    Ok(())
}

pub(crate) fn write_zoom(value: &Zoom, owner: &Owner) -> Result<Vec<u8>> {
    validate_zoom(value)?;
    let profile = Profile {
        namespaces: owner
            .namespaces
            .iter()
            .map(|(prefix, uri)| (prefix.clone(), uri.clone()))
            .collect(),
        pml: owner.pml_namespace.clone(),
        dml: owner.dml_namespace.clone(),
        relationship: owner.relationship_namespace.clone(),
    };
    let mut xml = String::with_capacity(4096);
    xml.push_str("<mc:AlternateContent");
    write_namespace_profile(&mut xml, &profile)?;
    xml.push('>');
    match value {
        Zoom::Section(value) => {
            write_unknown_choices(&mut xml, &value.unknown_xml)?;
            xml.push_str(
                "<mc:Choice Requires=\"psez\"><psez:sectionZm><psez:sectionZmObj sectionId=\"",
            );
            xml.push_str(&escape_xml(&value.section_id));
            xml.push_str("\">");
            write_properties(&mut xml, &value.properties)?;
            if let Some(extension) = &value.extension_xml {
                push_utf8(&mut xml, extension, "section zoom extension")?;
            }
            xml.push_str("</psez:sectionZmObj></psez:sectionZm></mc:Choice>");
            write_fallback(&mut xml, &value.fallback_xml)?;
        },
        Zoom::Slide(value) => {
            write_unknown_choices(&mut xml, &value.unknown_xml)?;
            write!(
                &mut xml,
                "<mc:Choice Requires=\"pslz\"><pslz:sldZm><pslz:sldZmObj sldId=\"{}\"",
                value.slide_id
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
            if let Some(creation_id) = value.creation_id {
                write!(&mut xml, " cId=\"{creation_id}\"")
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
            xml.push('>');
            write_properties(&mut xml, &value.properties)?;
            if let Some(extension) = &value.extension_xml {
                push_utf8(&mut xml, extension, "slide zoom extension")?;
            }
            xml.push_str("</pslz:sldZmObj></pslz:sldZm></mc:Choice>");
            write_fallback(&mut xml, &value.fallback_xml)?;
        },
        Zoom::Summary(value) => {
            write_unknown_choices(&mut xml, &value.unknown_xml)?;
            xml.push_str("<mc:Choice Requires=\"psuz\"><psuz:summaryZm>");
            for item in &value.items {
                xml.push_str("<psuz:summaryZmObj sectionId=\"");
                xml.push_str(&escape_xml(&item.section_id));
                xml.push('"');
                if !item.title.is_empty() {
                    xml.push_str(" title=\"");
                    xml.push_str(&escape_xml(&item.title));
                    xml.push('"');
                }
                if !item.description.is_empty() {
                    xml.push_str(" descr=\"");
                    xml.push_str(&escape_xml(&item.description));
                    xml.push('"');
                }
                if item.offset_x != Percentage::ZERO {
                    write!(&mut xml, " offsetFactorX=\"{}\"", item.offset_x.value())
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                if item.offset_y != Percentage::ZERO {
                    write!(&mut xml, " offsetFactorY=\"{}\"", item.offset_y.value())
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                if item.scale_x != Percentage::HUNDRED {
                    write!(&mut xml, " scaleFactorX=\"{}\"", item.scale_x.value())
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                if item.scale_y != Percentage::HUNDRED {
                    write!(&mut xml, " scaleFactorY=\"{}\"", item.scale_y.value())
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                xml.push('>');
                write_properties(&mut xml, &item.properties)?;
                if let Some(extension) = &item.extension_xml {
                    push_utf8(&mut xml, extension, "summary zoom item extension")?;
                }
                xml.push_str("</psuz:summaryZmObj>");
            }
            match value.layout {
                Layout::Grid => xml.push_str("<psuz:gridLayout/>"),
                Layout::Fixed => xml.push_str("<psuz:fixedLayout/>"),
            }
            if let Some(extension) = &value.extension_xml {
                push_utf8(&mut xml, extension, "summary zoom extension")?;
            }
            xml.push_str("</psuz:summaryZm></mc:Choice>");
            write_fallback(&mut xml, &value.fallback_xml)?;
        },
        Zoom::Unknown(value) => {
            return Ok(value.xml.clone());
        },
    }
    xml.push_str("</mc:AlternateContent>");
    if xml.len() > MAX_UNKNOWN_BYTES {
        return Err(Error::Limit {
            resource: "serialized zoom XML",
            limit: MAX_UNKNOWN_BYTES,
        });
    }
    Ok(xml.into_bytes())
}

fn write_properties(xml: &mut String, value: &Properties) -> Result<()> {
    xml.push_str("<p166:zmPr id=\"");
    xml.push_str(&escape_xml(&value.id));
    xml.push('"');
    if !value.return_to_parent {
        xml.push_str(" returnToParent=\"false\"");
    }
    if value.image_type == ImageType::Cover {
        xml.push_str(" imageType=\"cover\"");
    }
    if let Some(transition) = &value.transition {
        xml.push_str(" transitionDur=\"");
        xml.push_str(&escape_xml(transition.as_str()));
        xml.push('"');
    }
    if !value.show_background {
        xml.push_str(" showBg=\"false\"");
    }
    xml.push('>');
    push_utf8(xml, &value.blip_fill_xml, "zoom blipFill")?;
    push_utf8(xml, &value.shape_properties_xml, "zoom spPr")?;
    xml.push_str("</p166:zmPr>");
    Ok(())
}

fn write_fallback(xml: &mut String, fallback: &[u8]) -> Result<()> {
    xml.push_str("<mc:Fallback>");
    push_utf8(xml, fallback, "zoom fallback")?;
    xml.push_str("</mc:Fallback>");
    Ok(())
}

fn write_unknown_choices(xml: &mut String, values: &[Vec<u8>]) -> Result<()> {
    for value in values {
        push_utf8(xml, value, "zoom unknown choice")?;
    }
    Ok(())
}

fn write_namespace_profile(xml: &mut String, profile: &Profile) -> Result<()> {
    let mut namespaces = profile
        .namespaces
        .iter()
        .filter(|(prefix, _)| prefix != "xml" && prefix != "xmlns" && prefix != "z")
        .cloned()
        .collect::<HashMap<_, _>>();
    namespaces.insert("mc".into(), MC.into());
    namespaces.insert("p".into(), profile.pml.clone());
    namespaces.insert("a".into(), profile.dml.clone());
    namespaces.insert("r".into(), profile.relationship.clone());
    namespaces.insert("p166".into(), MAIN.into());
    namespaces.insert("psez".into(), SECTION_NS.into());
    namespaces.insert("pslz".into(), SLIDE_NS.into());
    namespaces.insert("psuz".into(), SUMMARY_NS.into());
    let mut values = namespaces.into_iter().collect::<Vec<_>>();
    values.sort_by(|(left, _), (right, _)| left.cmp(right));
    for (prefix, uri) in values {
        if prefix.is_empty() {
            write!(xml, " xmlns=\"{}\"", escape_xml(&uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        } else {
            write!(xml, " xmlns:{prefix}=\"{}\"", escape_xml(&uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
    }
    Ok(())
}

fn push_utf8(xml: &mut String, value: &[u8], field: &str) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|_err| Error::Invalid(format!("{field} is not UTF-8 XML")))?;
    xml.push_str(value);
    Ok(())
}

fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

/// Validate semantic targets and all image relationships against one OPC
/// package. This is intentionally separate from XML parsing: detached owners
/// remain useful for fixture editing, while package-facing owners cannot leak
/// dangling slide, section, or image references.
pub(crate) fn validate_in_package(
    owner: &mut Owner,
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<()> {
    validate_entries(&owner.entries)?;
    let graph = crate::presentation_properties::metadata::structure::load(package)?;
    let slide_ids = graph
        .slides
        .iter()
        .map(|value| value.slide_id)
        .collect::<HashSet<_>>();
    let section_ids = graph
        .sections
        .sections()
        .iter()
        .filter_map(|value| value.id.as_deref())
        .collect::<HashSet<_>>();
    for value in &mut owner.entries {
        match value {
            Zoom::Section(value) => {
                if !section_ids.contains(value.section_id.as_str()) {
                    return Err(Error::Relationship(format!(
                        "zoom section target '{}' is missing from the presentation",
                        value.section_id
                    )));
                }
                validate_properties_relationship(package, source, &mut value.properties)?;
            },
            Zoom::Slide(value) => {
                if !slide_ids.contains(&value.slide_id) {
                    return Err(Error::Relationship(format!(
                        "zoom slide target {} is missing from the presentation",
                        value.slide_id
                    )));
                }
                validate_properties_relationship(package, source, &mut value.properties)?;
            },
            Zoom::Summary(value) => {
                for item in &mut value.items {
                    if !section_ids.contains(item.section_id.as_str()) {
                        return Err(Error::Relationship(format!(
                            "summary zoom section target '{}' is missing from the presentation",
                            item.section_id
                        )));
                    }
                    validate_properties_relationship(package, source, &mut item.properties)?;
                }
            },
            Zoom::Unknown(_) => {},
        }
    }
    Ok(())
}

fn validate_properties_relationship(
    package: &OpcPackage,
    source: &dyn Part,
    properties: &mut Properties,
) -> Result<()> {
    let Some(image) = properties.image_relationship_mut() else {
        if properties.image_type == ImageType::Cover {
            return Err(Error::Relationship(
                "cover zoom is missing its a:blip image relationship".into(),
            ));
        }
        return Ok(());
    };
    let relationship = source.rels().get(image.id()).ok_or_else(|| {
        Error::Relationship(format!(
            "zoom image relationship '{}' is missing",
            image.id()
        ))
    })?;
    let expected_link = match image.link() {
        Link::Embed => false,
        Link::External => true,
    };
    if relationship.is_external() != expected_link {
        return Err(Error::Relationship(format!(
            "zoom image relationship '{}' has the wrong target mode",
            image.id()
        )));
    }
    if relationship.reltype() != IMAGE_REL && relationship.reltype() != IMAGE_REL_STRICT {
        return Err(Error::Relationship(format!(
            "zoom image relationship '{}' has an unexpected type",
            image.id()
        )));
    }
    if relationship.is_external() {
        let uri = relationship.target_ref();
        if uri.is_empty() {
            return Err(Error::Relationship(
                "zoom external image relationship has an empty target".into(),
            ));
        }
        image.target = Some(Target::External {
            uri: uri.to_owned(),
        });
    } else {
        let part_name = relationship.target_partname()?;
        let part = package.get_part(&part_name)?;
        if !part.content_type().starts_with("image/") {
            return Err(Error::ContentType {
                expected: "image/*".into(),
                actual: part.content_type().to_owned(),
            });
        }
        image.target = Some(Target::Internal {
            part_name: part_name.to_string(),
            content_type: part.content_type().to_owned(),
        });
    }
    Ok(())
}

struct StringWriter<'a> {
    bytes: &'a mut Vec<u8>,
}

impl std::fmt::Write for StringWriter<'_> {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        self.bytes.extend_from_slice(value.as_bytes());
        Ok(())
    }
}
