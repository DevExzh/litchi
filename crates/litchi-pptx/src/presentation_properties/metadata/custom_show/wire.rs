//! Source-preserving PresentationML custom-show wire operations.
//!
//! The semantic List deliberately stays small. This layer owns the source
//! layout needed to edit only known custShow and sld nodes while retaining
//! opaque XML siblings, attributes, whitespace, comments, and the owning
//! presentation relationship topology.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

use super::model::{List, Show};
use crate::presentation_properties::metadata::escape_xml;
use crate::{Error, Result};

pub(crate) const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const PML_SLIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const PML_SLIDE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const PML_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PML_NAMESPACE: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) close_start: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ChildSpan {
    pub(crate) span: Span,
    pub(crate) local: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SlideReference {
    pub(crate) slide_id: u32,
    pub(crate) relationship_id: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SlideSlot {
    pub(crate) span: Span,
    pub(crate) relationship_id: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SlideListLayout {
    pub(crate) span: Span,
    pub(crate) qualified_name: String,
    pub(crate) children: Vec<ChildSpan>,
    pub(crate) slots: Vec<SlideSlot>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ShowLayout {
    pub(crate) span: Span,
    pub(crate) qualified_name: String,
    pub(crate) id: u32,
    pub(crate) name: String,
    pub(crate) name_attribute: Range<usize>,
    pub(crate) slide_list: Option<SlideListLayout>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct CustomListLayout {
    pub(crate) span: Span,
    pub(crate) qualified_name: String,
    pub(crate) children: Vec<ChildSpan>,
    pub(crate) shows: Vec<ShowLayout>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Layout {
    pub(crate) root_close: usize,
    pub(crate) root_children: Vec<ChildSpan>,
    pub(crate) custom_list: Option<CustomListLayout>,
    pub(crate) slide_references: Vec<SlideReference>,
    pub(crate) slide_id_to_relationship: HashMap<u32, String>,
    pub(crate) p_prefix: String,
    pub(crate) r_prefix: String,
    pub(crate) p_namespace: String,
    pub(crate) r_namespace: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct RelationshipState {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) external: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Located {
    pub(crate) list: List,
    pub(crate) layout: Layout,
    pub(crate) relationships: Vec<RelationshipState>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FrameKind {
    Root,
    SlideIdList,
    CustomList,
    Show,
    SlideList,
    Slide,
    Other,
}

#[derive(Debug)]
struct Frame {
    kind: FrameKind,
    local: String,
    qualified_name: String,
    start: usize,
    close_start: usize,
    attributes: Vec<(String, String)>,
    children: Vec<ChildSpan>,
    slots: Vec<SlideSlot>,
    slide_list: Option<SlideListLayout>,
    slide_references: Vec<SlideReference>,
    show_layouts: Vec<ShowLayout>,
}

impl Frame {
    fn span(&self, end: usize) -> Span {
        Span {
            start: self.start,
            end,
            close_start: self.close_start,
        }
    }
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

pub(crate) fn locate(package: &OpcPackage) -> Result<Located> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    if presentation.blob().len() > MAX_BYTES {
        return Err(limit("custom-show PresentationML source bytes"));
    }

    let mut layout = scan_layout(presentation.blob())?;
    let mut relationships = relationship_states(presentation);
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    resolve_slides(package, presentation, &mut layout)?;
    let list = list_from_layout(&layout)?;
    validate_list(&list, layout.slide_id_to_relationship.keys())?;
    Ok(Located {
        list,
        layout,
        relationships,
    })
}

pub(crate) fn decode_rewritten(source: &[u8], original: &Layout) -> Result<(List, Layout)> {
    let mut layout = scan_layout(source)?;
    layout.slide_id_to_relationship = layout
        .slide_references
        .iter()
        .map(|reference| (reference.slide_id, reference.relationship_id.clone()))
        .collect();
    if layout.slide_references != original.slide_references
        || layout.slide_id_to_relationship != original.slide_id_to_relationship
    {
        return Err(invalid(
            "custom-show rewrite changed slide reference topology",
        ));
    }
    let list = list_from_layout(&layout)?;
    validate_list(&list, layout.slide_id_to_relationship.keys())?;
    Ok((list, layout))
}

pub(crate) fn rewrite(
    source: &[u8],
    layout: &Layout,
    before: &List,
    after: &List,
) -> Result<Vec<u8>> {
    if source.len() > MAX_BYTES {
        return Err(limit("custom-show PresentationML source bytes"));
    }
    validate_list(after, layout.slide_id_to_relationship.keys())?;
    if before.shows == after.shows {
        return Ok(source.to_vec());
    }

    let Some(custom) = layout.custom_list.as_ref() else {
        if after.is_empty() {
            return Ok(source.to_vec());
        }
        let rendered = render_custom_list(layout, after)?;
        return replace_ranges(
            source,
            &[Replacement {
                range: insertion_offset(layout)..insertion_offset(layout),
                value: rendered.into_bytes(),
            }],
        );
    };

    if custom.span.close_start == custom.span.start {
        let rendered = expand_empty_custom_list(source, layout, custom, after)?;
        return replace_ranges(
            source,
            &[Replacement {
                range: custom.span.start..custom.span.end,
                value: rendered.into_bytes(),
            }],
        );
    }

    let original_by_id = custom
        .shows
        .iter()
        .map(|show| (show.id, show))
        .collect::<HashMap<_, _>>();
    let original_values = before
        .shows
        .iter()
        .map(|show| (show.id, show))
        .collect::<HashMap<_, _>>();
    let slots = custom
        .children
        .iter()
        .filter(|child| child.local == "custShow")
        .collect::<Vec<_>>();
    let mut replacements = Vec::with_capacity(slots.len() + 1);

    for (index, slot) in slots.iter().enumerate() {
        let rendered = after
            .shows
            .get(index)
            .map(|show| {
                render_show(
                    source,
                    layout,
                    original_by_id.get(&show.id).copied(),
                    original_values.get(&show.id).copied(),
                    show,
                )
            })
            .transpose()?
            .unwrap_or_default();
        replacements.push(Replacement {
            range: slot.span.start..slot.span.end,
            value: rendered.into_bytes(),
        });
    }

    if after.shows.len() > slots.len() {
        let mut rendered = String::new();
        for show in &after.shows[slots.len()..] {
            rendered.push_str(&render_show(
                source,
                layout,
                original_by_id.get(&show.id).copied(),
                original_values.get(&show.id).copied(),
                show,
            )?);
        }
        replacements.push(Replacement {
            range: custom.span.close_start..custom.span.close_start,
            value: rendered.into_bytes(),
        });
    }
    replace_ranges(source, &replacements)
}

fn require_presentation(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid("main document is not a PowerPoint presentation"))
    }
}

fn relationship_states(part: &dyn Part) -> Vec<RelationshipState> {
    part.rels()
        .iter()
        .map(|relationship| RelationshipState {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        })
        .collect()
}

fn resolve_slides(
    package: &OpcPackage,
    presentation: &dyn Part,
    layout: &mut Layout,
) -> Result<()> {
    let mut slide_ids = HashSet::new();
    let mut relationship_ids = HashSet::new();
    let mut resolved = HashMap::new();

    for reference in &layout.slide_references {
        if reference.slide_id < 256 || !slide_ids.insert(reference.slide_id) {
            return Err(invalid("invalid or duplicate presentation slide ID"));
        }
        if !relationship_ids.insert(reference.relationship_id.as_str()) {
            return Err(invalid(
                "duplicate presentation slide relationship or target",
            ));
        }
        let relationship = presentation
            .rels()
            .get(&reference.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "orphan slide relationship {}",
                    reference.relationship_id
                ))
            })?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), PML_SLIDE_REL | PML_SLIDE_REL_STRICT)
        {
            return Err(invalid(format!(
                "relationship {} is not an internal slide relationship",
                reference.relationship_id
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != ct::PML_SLIDE {
            return Err(invalid(format!(
                "relationship {} targets a non-slide part",
                reference.relationship_id
            )));
        }
        resolved.insert(reference.slide_id, reference.relationship_id.clone());
    }

    if let Some(custom) = layout.custom_list.as_ref() {
        for show in &custom.shows {
            let mut membership = HashSet::new();
            if let Some(slide_list) = show.slide_list.as_ref() {
                for slot in &slide_list.slots {
                    let slide_id = resolved
                        .iter()
                        .find_map(|(slide_id, relationship_id)| {
                            (relationship_id == &slot.relationship_id).then_some(*slide_id)
                        })
                        .ok_or_else(|| {
                            invalid(format!(
                                "custom show {} has orphan slide relationship {}",
                                show.id, slot.relationship_id
                            ))
                        })?;
                    if !membership.insert(slide_id) {
                        return Err(invalid("duplicate custom-show slide reference"));
                    }
                }
            }
        }
    }

    layout.slide_id_to_relationship = resolved;
    Ok(())
}

fn list_from_layout(layout: &Layout) -> Result<List> {
    let mut list = List::new();
    let Some(custom) = layout.custom_list.as_ref() else {
        return Ok(list);
    };
    for raw in &custom.shows {
        let mut show = Show::new(raw.id, raw.name.clone());
        if let Some(slide_list) = raw.slide_list.as_ref() {
            for slot in &slide_list.slots {
                let slide_id = layout
                    .slide_id_to_relationship
                    .iter()
                    .find_map(|(slide_id, relationship_id)| {
                        (relationship_id == &slot.relationship_id).then_some(*slide_id)
                    })
                    .ok_or_else(|| {
                        invalid(format!(
                            "custom show {} has orphan slide relationship {}",
                            raw.id, slot.relationship_id
                        ))
                    })?;
                show.slide_ids.push(slide_id);
            }
        }
        list.add(show);
    }
    Ok(list)
}

pub(crate) fn validate_list<'a>(
    list: &List,
    slide_ids: impl Iterator<Item = &'a u32>,
) -> Result<()> {
    let available = slide_ids.copied().collect::<HashSet<_>>();
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for show in &list.shows {
        if show.name.is_empty() || !ids.insert(show.id) || !names.insert(show.name.as_str()) {
            return Err(invalid("duplicate or empty custom-show identity"));
        }
        let mut membership = HashSet::new();
        for slide_id in &show.slide_ids {
            if !available.contains(slide_id) || !membership.insert(*slide_id) {
                return Err(invalid("duplicate or orphan custom-show slide reference"));
            }
        }
    }
    Ok(())
}

fn scan_layout(xml: &[u8]) -> Result<Layout> {
    if xml.len() > MAX_BYTES {
        return Err(limit("custom-show PresentationML source bytes"));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<Frame>::new();
    let mut root_close = None;
    let mut root_children = Vec::new();
    let mut custom_list = None;
    let mut slide_references = Vec::new();
    let mut p_prefix = String::new();
    let r_prefix = "r".to_owned();
    let mut p_namespace = String::new();
    let mut r_namespace = PML_RELATIONSHIPS.to_owned();
    let mut nodes = 0usize;
    let mut saw_root = false;

    loop {
        let before = reader.buffer_position() as usize;
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                count_node(&mut nodes, stack.len())?;
                let local = local_name(element.name().as_ref())?;
                let qualified_name = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let attributes = attributes(&element, decoder)?;
                if stack.is_empty() {
                    if local != "presentation" || saw_root {
                        return Err(invalid("custom-show owner is not a presentation root"));
                    }
                    saw_root = true;
                    p_prefix = prefix(&qualified_name).to_owned();
                    if let Some(value) = find_attribute(&attributes, "xmlns:p") {
                        p_namespace = value.to_owned();
                    }
                    if let Some(value) = find_attribute(&attributes, "xmlns:r") {
                        r_namespace = value.to_owned();
                    }
                }
                stack.push(frame(
                    classify(&stack, &local),
                    local,
                    qualified_name,
                    before,
                    attributes,
                ));
            },
            Ok(Event::Empty(element)) => {
                count_node(&mut nodes, stack.len())?;
                let local = local_name(element.name().as_ref())?;
                let qualified_name = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let attributes = attributes(&element, decoder)?;
                if stack.is_empty() {
                    if local != "presentation" || saw_root {
                        return Err(invalid("custom-show owner is not a presentation root"));
                    }
                    saw_root = true;
                    p_prefix = prefix(&qualified_name).to_owned();
                    if let Some(value) = find_attribute(&attributes, "xmlns:p") {
                        p_namespace = value.to_owned();
                    }
                    if let Some(value) = find_attribute(&attributes, "xmlns:r") {
                        r_namespace = value.to_owned();
                    }
                }
                let mut value = frame(
                    classify(&stack, &local),
                    local,
                    qualified_name,
                    before,
                    attributes,
                );
                value.close_start = before;
                attach_frame(
                    value,
                    reader.buffer_position() as usize,
                    &mut stack,
                    xml,
                    &mut root_children,
                    &mut custom_list,
                    &mut slide_references,
                )?;
            },
            Ok(Event::End(element)) => {
                count_node(&mut nodes, stack.len())?;
                let mut value = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected custom-show closing element"))?;
                let local = local_name(element.name().as_ref())?;
                if value.local != local {
                    return Err(invalid("mismatched custom-show XML element"));
                }
                value.close_start = before;
                let was_root = value.kind == FrameKind::Root;
                attach_frame(
                    value,
                    reader.buffer_position() as usize,
                    &mut stack,
                    xml,
                    &mut root_children,
                    &mut custom_list,
                    &mut slide_references,
                )?;
                if was_root {
                    root_close = Some(before);
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid(
                    "custom-show XML must not contain DTDs or processing instructions",
                ));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(Error::Xml(error.to_string())),
        }
    }

    if !saw_root || !stack.is_empty() {
        return Err(invalid("unterminated custom-show PresentationML"));
    }
    Ok(Layout {
        root_close: root_close.ok_or_else(|| invalid("presentation root is not closed"))?,
        root_children,
        custom_list,
        slide_references,
        slide_id_to_relationship: HashMap::new(),
        p_prefix,
        r_prefix,
        p_namespace: if p_namespace.is_empty() {
            PML_NAMESPACE.to_owned()
        } else {
            p_namespace
        },
        r_namespace,
    })
}

fn frame(
    kind: FrameKind,
    local: String,
    qualified_name: String,
    start: usize,
    attributes: Vec<(String, String)>,
) -> Frame {
    Frame {
        kind,
        local,
        qualified_name,
        start,
        close_start: 0,
        attributes,
        children: Vec::new(),
        slots: Vec::new(),
        slide_list: None,
        slide_references: Vec::new(),
        show_layouts: Vec::new(),
    }
}

fn classify(stack: &[Frame], local: &str) -> FrameKind {
    match stack.last().map(|value| value.kind) {
        None if local == "presentation" => FrameKind::Root,
        Some(FrameKind::Root) => match local {
            "sldIdLst" => FrameKind::SlideIdList,
            "custShowLst" => FrameKind::CustomList,
            _ => FrameKind::Other,
        },
        Some(FrameKind::CustomList) if local == "custShow" => FrameKind::Show,
        Some(FrameKind::Show) if local == "sldLst" => FrameKind::SlideList,
        Some(FrameKind::SlideList) if local == "sld" => FrameKind::Slide,
        _ => FrameKind::Other,
    }
}

fn attach_frame(
    value: Frame,
    end: usize,
    stack: &mut [Frame],
    source: &[u8],
    root_children: &mut Vec<ChildSpan>,
    custom_list: &mut Option<CustomListLayout>,
    slide_references: &mut Vec<SlideReference>,
) -> Result<()> {
    let span = value.span(end);
    let child = ChildSpan {
        span: span.clone(),
        local: value.local.clone(),
    };
    let Some(parent) = stack.last_mut() else {
        if value.kind != FrameKind::Root {
            return Err(invalid("custom-show XML has content outside its root"));
        }
        root_children.extend(value.children);
        slide_references.extend(value.slide_references);
        return Ok(());
    };

    match value.kind {
        FrameKind::Root => return Err(invalid("nested presentation roots are rejected")),
        FrameKind::SlideIdList => {
            parent.children.push(child);
            parent.slide_references.extend(value.slide_references);
        },
        FrameKind::CustomList => {
            parent.children.push(child);
            let layout = CustomListLayout {
                span,
                qualified_name: value.qualified_name,
                children: value.children,
                shows: value.show_layouts,
            };
            if parent.kind == FrameKind::Root {
                if custom_list.replace(layout).is_some() {
                    return Err(invalid("presentation contains multiple custom-show lists"));
                }
            }
        },
        FrameKind::Show => {
            let layout = parse_show_layout(source, span.clone(), &value)?;
            parent.children.push(child);
            if parent.kind == FrameKind::CustomList {
                parent.show_layouts.push(layout);
            }
        },
        FrameKind::SlideList => {
            let layout = SlideListLayout {
                span,
                qualified_name: value.qualified_name,
                children: value.children,
                slots: value.slots,
            };
            parent.children.push(child);
            if parent.kind == FrameKind::Show {
                parent.slide_list = Some(layout);
            }
        },
        FrameKind::Slide => {
            let relationship_id = find_relationship_attribute(&value.attributes)
                .ok_or_else(|| invalid("custom-show slide is missing r:id"))?
                .to_owned();
            parent.children.push(child);
            if parent.kind == FrameKind::SlideList {
                parent.slots.push(SlideSlot {
                    span,
                    relationship_id,
                });
            }
        },
        FrameKind::Other => {
            parent.children.push(child);
            if parent.kind == FrameKind::SlideIdList && value.local == "sldId" {
                let id = find_attribute(&value.attributes, "id")
                    .ok_or_else(|| invalid("presentation slide is missing its ID"))?
                    .parse::<u32>()
                    .map_err(|_| invalid("presentation slide ID is not a u32"))?;
                let relationship_id = find_qualified_relationship_attribute(&value.attributes)
                    .ok_or_else(|| invalid("presentation slide is missing its relationship ID"))?
                    .to_owned();
                parent.slide_references.push(SlideReference {
                    slide_id: id,
                    relationship_id,
                });
            }
        },
    }
    Ok(())
}

fn parse_show_layout(source: &[u8], span: Span, value: &Frame) -> Result<ShowLayout> {
    let start_end = source[span.start..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| span.start + offset + 1)
        .ok_or_else(|| invalid("custom-show start tag is unterminated"))?;
    let raw = &source[span.start..start_end];
    let (qualified_name, parsed) = parse_raw_start_tag(raw)?;
    let id = find_attribute(&parsed, "id")
        .ok_or_else(|| invalid("custom show is missing its ID"))?
        .parse::<u32>()
        .map_err(|_| invalid("custom show ID is not a u32"))?;
    let name = find_attribute(&parsed, "name")
        .ok_or_else(|| invalid("custom show is missing its name"))?
        .to_owned();
    if name.is_empty() {
        return Err(invalid("custom-show name cannot be empty"));
    }
    let name_attribute = attribute_range(raw, "name")
        .map(|range| range.start + span.start..range.end + span.start)
        .ok_or_else(|| invalid("custom show name attribute is malformed"))?;
    Ok(ShowLayout {
        span,
        qualified_name,
        id,
        name,
        name_attribute,
        slide_list: value.slide_list.clone(),
    })
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate custom-show XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn parse_raw_start_tag(raw: &[u8]) -> Result<(String, Vec<(String, String)>)> {
    let end = raw
        .iter()
        .position(|byte| *byte == b'>')
        .ok_or_else(|| invalid("custom-show XML start tag is missing '>'"))?;
    let tag = &raw[..=end];
    let name_start = tag
        .iter()
        .position(|byte| *byte == b'<')
        .map(|offset| offset + 1)
        .ok_or_else(|| invalid("custom-show XML start tag is malformed"))?;
    let name_end = tag[name_start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'>' || *byte == b'/')
        .map(|offset| name_start + offset)
        .unwrap_or(end);
    let qualified_name = std::str::from_utf8(&tag[name_start..name_end])
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let mut values = Vec::new();
    let mut cursor = name_end;
    while cursor < end {
        while cursor < end && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || tag[cursor] == b'/' {
            break;
        }
        let key_start = cursor;
        while cursor < end && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'=' {
            cursor += 1;
        }
        let key = std::str::from_utf8(&tag[key_start..cursor])
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        while cursor < end && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || tag[cursor] != b'=' {
            return Err(invalid("custom-show XML attribute is missing '='"));
        }
        cursor += 1;
        while cursor < end && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| invalid("custom-show XML attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("custom-show XML attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < end && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end {
            return Err(invalid("custom-show XML attribute value is unterminated"));
        }
        let value = quick_xml::escape::unescape(
            std::str::from_utf8(&tag[value_start..cursor])
                .map_err(|error| Error::Xml(error.to_string()))?,
        )
        .map_err(|error| Error::Xml(error.to_string()))?
        .into_owned();
        values.push((key, value));
        cursor += 1;
    }
    Ok((qualified_name, values))
}

fn attribute_range(raw: &[u8], wanted: &str) -> Option<Range<usize>> {
    let end = raw.iter().position(|byte| *byte == b'>')?;
    let tag = &raw[..end];
    let name_start = tag.iter().position(|byte| *byte == b'<')? + 1;
    let mut cursor = name_start;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && tag[cursor] != b'>'
        && tag[cursor] != b'/'
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] == b'/' {
            break;
        }
        let key_start = cursor;
        while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'=' {
            cursor += 1;
        }
        let key = std::str::from_utf8(&tag[key_start..cursor]).ok()?;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] != b'=' {
            return None;
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if quote != b'\'' && quote != b'"' {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return None;
        }
        if key == wanted {
            return Some(value_start..cursor);
        }
        cursor += 1;
    }
    None
}

fn find_attribute<'a>(attributes: &'a [(String, String)], wanted: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(name, _)| name == wanted)
        .map(|(_, value)| value.as_str())
}

fn find_relationship_attribute<'a>(attributes: &'a [(String, String)]) -> Option<&'a str> {
    attributes.iter().find_map(|(name, value)| {
        (name.rsplit_once(':').map(|(_, local)| local) == Some("id") || name == "id")
            .then_some(value.as_str())
    })
}

fn find_qualified_relationship_attribute<'a>(
    attributes: &'a [(String, String)],
) -> Option<&'a str> {
    attributes.iter().find_map(|(name, value)| {
        (name.rsplit_once(':').map(|(_, local)| local) == Some("id")).then_some(value.as_str())
    })
}

fn prefix(value: &str) -> &str {
    value.rsplit_once(':').map_or("", |(prefix, _)| prefix)
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(|error| Error::Xml(error.to_string()))?;
    Ok(name
        .rsplit_once(':')
        .map_or(name, |(_, local)| local)
        .to_owned())
}

fn count_node(nodes: &mut usize, depth: usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("custom-show XML node count"))?;
    if *nodes > MAX_NODES || depth >= MAX_DEPTH {
        return Err(limit("custom-show XML resource limit"));
    }
    Ok(())
}

fn insertion_offset(layout: &Layout) -> usize {
    layout
        .root_children
        .iter()
        .find(|child| schema_rank(&child.local).is_some_and(|rank| rank > 9))
        .map_or(layout.root_close, |child| child.span.start)
}

fn schema_rank(local: &str) -> Option<usize> {
    [
        "sldMasterIdLst",
        "notesMasterIdLst",
        "handoutMasterIdLst",
        "sldIdLst",
        "sldSz",
        "notesSz",
        "smartTags",
        "embeddedFontLst",
        "custShowLst",
        "photoAlbum",
        "custDataLst",
        "kinsoku",
        "defaultTextStyle",
        "modifyVerifier",
        "extLst",
    ]
    .iter()
    .position(|value| *value == local)
    .map(|value| value + 1)
}

fn render_custom_list(layout: &Layout, list: &List) -> Result<String> {
    let name = qualified(&layout.p_prefix, "custShowLst");
    let shows = list
        .shows
        .iter()
        .map(|show| render_new_show(layout, show))
        .collect::<Result<Vec<_>>>()?
        .join("");
    Ok(format!(
        "<{name} xmlns:{p}=\"{}\" xmlns:{r}=\"{}\">{shows}</{name}>",
        escape_xml(&layout.p_namespace),
        escape_xml(&layout.r_namespace),
        p = layout.p_prefix,
        r = layout.r_prefix,
        name = name,
        shows = shows,
    ))
}

fn expand_empty_custom_list(
    source: &[u8],
    layout: &Layout,
    custom: &CustomListLayout,
    list: &List,
) -> Result<String> {
    if list.is_empty() {
        return Ok(
            String::from_utf8_lossy(&source[custom.span.start..custom.span.end]).into_owned(),
        );
    }
    let source = &source[custom.span.start..custom.span.end];
    let close = source
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("empty custom-show list has no closing bracket"))?;
    let opening = std::str::from_utf8(&source[..close])
        .map_err(|error| Error::Xml(error.to_string()))?
        .trim_end_matches('/');
    let shows = list
        .shows
        .iter()
        .map(|show| render_new_show(layout, show))
        .collect::<Result<Vec<_>>>()?
        .join("");
    Ok(format!("{opening}>{shows}</{}>", custom.qualified_name))
}

fn render_show(
    source: &[u8],
    layout: &Layout,
    original_layout: Option<&ShowLayout>,
    original: Option<&Show>,
    show: &Show,
) -> Result<String> {
    match (original_layout, original) {
        (Some(original_layout), Some(original)) if original.id == show.id => {
            render_existing_show(source, layout, original_layout, original, show)
        },
        _ => render_new_show(layout, show),
    }
}

fn render_new_show(layout: &Layout, show: &Show) -> Result<String> {
    let show_name = qualified(&layout.p_prefix, "custShow");
    let list_name = qualified(&layout.p_prefix, "sldLst");
    let slides = show
        .slide_ids
        .iter()
        .map(|slide_id| render_new_slide(layout, *slide_id))
        .collect::<Result<Vec<_>>>()?
        .join("");
    Ok(format!(
        "<{show_name} name=\"{}\" id=\"{}\"><{list_name}>{slides}</{list_name}></{show_name}>",
        escape_xml(&show.name),
        show.id,
        show_name = show_name,
        list_name = list_name,
        slides = slides,
    ))
}

fn render_new_slide(layout: &Layout, slide_id: u32) -> Result<String> {
    let relationship_id = layout
        .slide_id_to_relationship
        .get(&slide_id)
        .ok_or_else(|| invalid("custom show references an undeclared slide"))?;
    Ok(format!(
        "<{} {}:id=\"{}\"/>",
        qualified(&layout.p_prefix, "sld"),
        layout.r_prefix,
        escape_xml(relationship_id),
    ))
}

fn render_existing_show(
    source: &[u8],
    layout: &Layout,
    original_layout: &ShowLayout,
    original: &Show,
    show: &Show,
) -> Result<String> {
    let offset = original_layout.span.start;
    let source_show = &source[offset..original_layout.span.end];
    let mut replacements = Vec::new();
    if original.name != show.name {
        replacements.push(Replacement {
            range: original_layout.name_attribute.start - offset
                ..original_layout.name_attribute.end - offset,
            value: escape_xml(&show.name).into_bytes(),
        });
    }
    if original.slide_ids != show.slide_ids {
        if let Some(slide_list) = original_layout.slide_list.as_ref() {
            replacements.push(Replacement {
                range: slide_list.span.start - offset..slide_list.span.end - offset,
                value: render_slide_list(source, layout, slide_list, &show.slide_ids)?.into_bytes(),
            });
        } else if !show.slide_ids.is_empty() {
            replacements.push(Replacement {
                range: original_layout.span.close_start - offset
                    ..original_layout.span.close_start - offset,
                value: render_new_slide_list(layout, &show.slide_ids)?.into_bytes(),
            });
        }
    }
    if replacements.is_empty() {
        return Ok(String::from_utf8_lossy(source_show).into_owned());
    }
    String::from_utf8(replace_ranges(source_show, &replacements)?)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn render_slide_list(
    source: &[u8],
    layout: &Layout,
    slide_list: &SlideListLayout,
    slide_ids: &[u32],
) -> Result<String> {
    let source_list = &source[slide_list.span.start..slide_list.span.end];
    let mut source_by_slide = HashMap::new();
    for slot in &slide_list.slots {
        if let Some(slide_id) =
            layout
                .slide_id_to_relationship
                .iter()
                .find_map(|(slide_id, relationship_id)| {
                    (relationship_id == &slot.relationship_id).then_some(*slide_id)
                })
        {
            source_by_slide.insert(slide_id, &source[slot.span.start..slot.span.end]);
        }
    }
    let slots = slide_list
        .children
        .iter()
        .filter(|child| child.local == "sld")
        .collect::<Vec<_>>();
    let mut replacements = Vec::with_capacity(slots.len() + 1);
    for (index, slot) in slots.iter().enumerate() {
        let rendered = slide_ids
            .get(index)
            .map(|slide_id| {
                source_by_slide.get(slide_id).map_or_else(
                    || render_new_slide(layout, *slide_id),
                    |raw| Ok(String::from_utf8_lossy(raw).into_owned()),
                )
            })
            .transpose()?
            .unwrap_or_default();
        replacements.push(Replacement {
            range: slot.span.start - slide_list.span.start..slot.span.end - slide_list.span.start,
            value: rendered.into_bytes(),
        });
    }
    if slide_ids.len() > slots.len() {
        let mut rendered = String::new();
        for slide_id in &slide_ids[slots.len()..] {
            rendered.push_str(&render_new_slide(layout, *slide_id)?);
        }
        replacements.push(Replacement {
            range: slide_list.span.close_start - slide_list.span.start
                ..slide_list.span.close_start - slide_list.span.start,
            value: rendered.into_bytes(),
        });
    }
    String::from_utf8(replace_ranges(source_list, &replacements)?)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn render_new_slide_list(layout: &Layout, slide_ids: &[u32]) -> Result<String> {
    let name = qualified(&layout.p_prefix, "sldLst");
    let slides = slide_ids
        .iter()
        .map(|slide_id| render_new_slide(layout, *slide_id))
        .collect::<Result<Vec<_>>>()?
        .join("");
    Ok(format!(
        "<{name}>{slides}</{name}>",
        name = name,
        slides = slides
    ))
}

fn qualified(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn replace_ranges(source: &[u8], replacements: &[Replacement]) -> Result<Vec<u8>> {
    let mut replacements = replacements.iter().collect::<Vec<_>>();
    replacements.sort_by_key(|replacement| (replacement.range.start, replacement.range.end));
    for pair in replacements.windows(2) {
        if pair[0].range.end > pair[1].range.start {
            return Err(invalid("overlapping custom-show XML patch ranges"));
        }
    }
    let mut output = source.to_vec();
    for replacement in replacements.into_iter().rev() {
        if replacement.range.end > output.len() {
            return Err(invalid("custom-show patch range is outside the source"));
        }
        output.splice(replacement.range.clone(), replacement.value.iter().copied());
    }
    if output.len() > MAX_BYTES {
        return Err(limit("patched custom-show PresentationML source bytes"));
    }
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &str) -> Error {
    invalid(format!("{resource} exceed the bounded custom-show limit"))
}
