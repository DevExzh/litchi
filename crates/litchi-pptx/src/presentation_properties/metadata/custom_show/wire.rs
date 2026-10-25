//! Source-preserving `PresentationML` custom-show wire operations.
//!
//! The semantic List deliberately stays small. This layer owns the source
//! layout needed to edit only known custShow and sld nodes while retaining
//! opaque XML siblings, attributes, whitespace, comments, and the owning
//! presentation relationship topology.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

use super::super::{BoundedString, Limits, check_string, checked_output_len, escaped_xml_len};
use super::model::{List, Show};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

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

pub(crate) fn locate(package: &OpcPackage, limits: Limits) -> Result<Located> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    if presentation.blob().len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            presentation.blob().len(),
            limits.max_input_bytes(),
        ));
    }

    let mut layout = scan_layout(presentation.blob(), limits)?;
    let mut relationships = relationship_states(presentation);
    if relationships.len() > limits.max_relationships() {
        return Err(limit(
            crate::opened::CodecLimitKind::Relationships,
            relationships.len(),
            limits.max_relationships(),
        ));
    }
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    resolve_slides(package, presentation, &mut layout)?;
    let list = list_from_layout(&layout, limits)?;
    validate_list(&list, layout.slide_id_to_relationship.keys(), limits)?;
    Ok(Located {
        list,
        layout,
        relationships,
    })
}

pub(crate) fn decode_rewritten(
    source: &[u8],
    original: &Layout,
    limits: Limits,
) -> Result<(List, Layout)> {
    let mut layout = scan_layout(source, limits)?;
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
    let list = list_from_layout(&layout, limits)?;
    validate_list(&list, layout.slide_id_to_relationship.keys(), limits)?;
    Ok((list, layout))
}

/// Rewrite a detached custom-show value directly against a presentation XML
/// source.  The structure graph family uses this entry point so that its
/// graph edit shares the custom-show wire layout: unknown attributes,
/// children, comments, and whitespace remain in place while known semantic
/// slots are changed.
pub(crate) fn rewrite_source(
    source: &[u8],
    before: &List,
    after: &List,
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut layout = scan_layout(source, limits)?;
    layout.slide_id_to_relationship = layout
        .slide_references
        .iter()
        .map(|reference| (reference.slide_id, reference.relationship_id.clone()))
        .collect();
    validate_list(after, layout.slide_id_to_relationship.keys(), limits)?;
    rewrite(source, &layout, before, after, limits)
}

pub(crate) fn rewrite(
    source: &[u8],
    layout: &Layout,
    before: &List,
    after: &List,
    limits: Limits,
) -> Result<Vec<u8>> {
    if source.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            source.len(),
            limits.max_input_bytes(),
        ));
    }
    validate_list(after, layout.slide_id_to_relationship.keys(), limits)?;
    if before.shows == after.shows {
        if source.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                source.len(),
                limits.max_output_bytes(),
            ));
        }
        return Ok(source.to_vec());
    }

    let Some(custom) = layout.custom_list.as_ref() else {
        if after.is_empty() {
            if source.len() > limits.max_output_bytes() {
                return Err(limit(
                    crate::opened::CodecLimitKind::OutputBytes,
                    source.len(),
                    limits.max_output_bytes(),
                ));
            }
            return Ok(source.to_vec());
        }
        let rendered_len = render_custom_list_len(layout, after, limits)?;
        checked_output_len(
            source.len(),
            0,
            rendered_len,
            limits,
            "patched custom-show PresentationML source bytes",
        )?;
        let rendered = render_custom_list(layout, after, limits)?;
        return replace_ranges(
            source,
            &[Replacement {
                range: insertion_offset(layout)..insertion_offset(layout),
                value: rendered.into_bytes(),
            }],
            limits,
        );
    };

    if custom.span.close_start == custom.span.start {
        let replacement_len = expand_empty_custom_list_len(source, layout, custom, after, limits)?;
        checked_output_len(
            source.len(),
            custom.span.end - custom.span.start,
            replacement_len,
            limits,
            "patched custom-show PresentationML source bytes",
        )?;
        let rendered = expand_empty_custom_list(source, layout, custom, after, limits)?;
        return replace_ranges(
            source,
            &[Replacement {
                range: custom.span.start..custom.span.end,
                value: rendered.into_bytes(),
            }],
            limits,
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
    let slot_count = custom
        .children
        .iter()
        .filter(|child| child.local == "custShow")
        .count();

    // Calculate the complete replacement delta before constructing any of
    // the owned replacement buffers.  A per-show BoundedString is not enough:
    // several individually valid replacements can still exceed the final
    // presentation output policy together.
    let mut removed = 0usize;
    let mut added = 0usize;
    for (index, slot) in custom
        .children
        .iter()
        .filter(|child| child.local == "custShow")
        .enumerate()
    {
        removed = removed.checked_add(slot.span.end - slot.span.start).ok_or(
            Error::CapacityOverflow {
                resource: "patched custom-show PresentationML source bytes",
            },
        )?;
        if let Some(show) = after.shows.get(index) {
            added = added
                .checked_add(render_show_len(
                    source,
                    layout,
                    original_by_id.get(&show.id).copied(),
                    original_values.get(&show.id).copied(),
                    show,
                    limits,
                )?)
                .ok_or(Error::CapacityOverflow {
                    resource: "patched custom-show PresentationML source bytes",
                })?;
        }
    }
    if after.shows.len() > slot_count {
        for show in &after.shows[slot_count..] {
            added = added
                .checked_add(render_show_len(
                    source,
                    layout,
                    original_by_id.get(&show.id).copied(),
                    original_values.get(&show.id).copied(),
                    show,
                    limits,
                )?)
                .ok_or(Error::CapacityOverflow {
                    resource: "patched custom-show PresentationML source bytes",
                })?;
        }
    }
    checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "patched custom-show PresentationML source bytes",
    )?;

    let slots = custom
        .children
        .iter()
        .filter(|child| child.local == "custShow")
        .collect::<Vec<_>>();
    let mut replacements = Vec::with_capacity(slot_count + 1);

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
                    limits,
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
        let mut rendered = BoundedString::new(limits, "custom-show XML output");
        for show in &after.shows[slots.len()..] {
            rendered.push_str(&render_show(
                source,
                layout,
                original_by_id.get(&show.id).copied(),
                original_values.get(&show.id).copied(),
                show,
                limits,
            )?)?;
        }
        replacements.push(Replacement {
            range: custom.span.close_start..custom.span.close_start,
            value: rendered.finish().into_bytes(),
        });
    }
    replace_ranges(source, &replacements, limits)
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

fn list_from_layout(layout: &Layout, limits: Limits) -> Result<List> {
    let mut list = List::new();
    let Some(custom) = layout.custom_list.as_ref() else {
        return Ok(list);
    };
    for raw in &custom.shows {
        if list.shows.len() >= limits.max_custom_shows() {
            return Err(limit(
                crate::opened::CodecLimitKind::CustomShows,
                list.shows.len() + 1,
                limits.max_custom_shows(),
            ));
        }
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
    limits: Limits,
) -> Result<()> {
    if list.shows.len() > limits.max_custom_shows() {
        return Err(limit(
            crate::opened::CodecLimitKind::CustomShows,
            list.shows.len(),
            limits.max_custom_shows(),
        ));
    }
    let available = slide_ids.copied().collect::<HashSet<_>>();
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for show in &list.shows {
        if show.name.len() > limits.max_string_bytes()
            || show.name.is_empty()
            || !ids.insert(show.id)
            || !names.insert(show.name.as_str())
        {
            return Err(invalid("duplicate or empty custom-show identity"));
        }
        if show.slide_ids.len() > limits.max_slide_references() {
            return Err(limit(
                crate::opened::CodecLimitKind::SlideReferences,
                show.slide_ids.len(),
                limits.max_slide_references(),
            ));
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

fn scan_layout(xml: &[u8], limits: Limits) -> Result<Layout> {
    if xml.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            xml.len(),
            limits.max_input_bytes(),
        ));
    }
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
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
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML source position exceeds usize"))?;
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                count_node(&mut nodes, stack.len(), limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.name().as_ref())?;
                let qualified_name = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let attributes = attributes(&element, decoder, limits)?;
                if stack.is_empty() {
                    if local != "presentation" || saw_root {
                        return Err(invalid("custom-show owner is not a presentation root"));
                    }
                    saw_root = true;
                    prefix(&qualified_name).clone_into(&mut p_prefix);
                    if let Some(value) = find_attribute(&attributes, "xmlns:p") {
                        value.clone_into(&mut p_namespace);
                    }
                    if let Some(value) = find_attribute(&attributes, "xmlns:r") {
                        value.clone_into(&mut r_namespace);
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
                count_node(&mut nodes, stack.len(), limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.name().as_ref())?;
                let qualified_name = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let attributes = attributes(&element, decoder, limits)?;
                if stack.is_empty() {
                    if local != "presentation" || saw_root {
                        return Err(invalid("custom-show owner is not a presentation root"));
                    }
                    saw_root = true;
                    prefix(&qualified_name).clone_into(&mut p_prefix);
                    if let Some(value) = find_attribute(&attributes, "xmlns:p") {
                        value.clone_into(&mut p_namespace);
                    }
                    if let Some(value) = find_attribute(&attributes, "xmlns:r") {
                        value.clone_into(&mut r_namespace);
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
                    origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?,
                    &mut stack,
                    xml,
                    &mut root_children,
                    &mut custom_list,
                    &mut slide_references,
                )?;
            },
            Ok(Event::End(element)) => {
                count_node(&mut nodes, stack.len(), limits)?;
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
                    origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?,
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
            if parent.kind == FrameKind::Root && custom_list.replace(layout).is_some() {
                return Err(invalid("presentation contains multiple custom-show lists"));
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
                    .map_err(|_err| invalid("presentation slide ID is not a u32"))?;
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
    let raw_end = if value.close_start > span.start {
        value.close_start
    } else {
        span.end
    };
    let raw = &source[span.start..raw_end];
    let id = find_attribute(&value.attributes, "id")
        .ok_or_else(|| invalid("custom show is missing its ID"))?
        .parse::<u32>()
        .map_err(|_err| invalid("custom show ID is not a u32"))?;
    let name = find_attribute(&value.attributes, "name")
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
        qualified_name: value.qualified_name.clone(),
        id,
        name,
        name_attribute,
        slide_list: value.slide_list.clone(),
    })
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if values.len() >= limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                values.len() + 1,
                limits.max_attributes(),
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if name.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                name.len(),
                limits.max_string_bytes(),
            ));
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let name = name.to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate custom-show XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
        values.push((name, value.into_owned()));
    }
    Ok(values)
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

fn find_relationship_attribute(attributes: &[(String, String)]) -> Option<&str> {
    attributes.iter().find_map(|(name, value)| {
        (name.rsplit_once(':').map(|(_, local)| local) == Some("id") || name == "id")
            .then_some(value.as_str())
    })
}

fn find_qualified_relationship_attribute(attributes: &[(String, String)]) -> Option<&str> {
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

fn count_node(nodes: &mut usize, depth: usize, limits: Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        limit(
            crate::opened::CodecLimitKind::XmlNodes,
            usize::MAX,
            limits.max_nodes(),
        )
    })?;
    if *nodes > limits.max_nodes() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlNodes,
            *nodes,
            limits.max_nodes(),
        ));
    }
    if depth >= limits.max_depth() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlDepth,
            depth + 1,
            limits.max_depth(),
        ));
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

fn render_custom_list(layout: &Layout, list: &List, limits: Limits) -> Result<String> {
    let name = qualified(&layout.p_prefix, "custShowLst", limits)?;
    let mut value = BoundedString::new(limits, "custom-show XML output");
    value.push_str("<")?;
    value.push_str(&name)?;
    value.push_str(" xmlns:")?;
    value.push_str(&layout.p_prefix)?;
    value.push_str("=\"")?;
    value.push_escaped_xml(&layout.p_namespace)?;
    value.push_str("\" xmlns:")?;
    value.push_str(&layout.r_prefix)?;
    value.push_str("=\"")?;
    value.push_escaped_xml(&layout.r_namespace)?;
    value.push_str("\">")?;
    for show in &list.shows {
        value.push_str(&render_new_show(layout, show, limits)?)?;
    }
    value.push_str("</")?;
    value.push_str(&name)?;
    value.push_str(">")?;
    Ok(value.finish())
}

fn render_custom_list_len(layout: &Layout, list: &List, limits: Limits) -> Result<usize> {
    let name = qualified_len(&layout.p_prefix, "custShowLst", limits)?;
    let p_prefix = layout.p_prefix.len();
    let r_prefix = layout.r_prefix.len();
    let p_namespace = escaped_xml_len(&layout.p_namespace)?;
    let r_namespace = escaped_xml_len(&layout.r_namespace)?;
    let mut length = 1usize
        .checked_add(name)
        .and_then(|value| value.checked_add(b" xmlns:".len()))
        .and_then(|value| value.checked_add(p_prefix))
        .and_then(|value| value.checked_add(b"=\"".len()))
        .and_then(|value| value.checked_add(p_namespace))
        .and_then(|value| value.checked_add(b"\" xmlns:".len()))
        .and_then(|value| value.checked_add(r_prefix))
        .and_then(|value| value.checked_add(b"=\"".len()))
        .and_then(|value| value.checked_add(r_namespace))
        .and_then(|value| value.checked_add(b"\">".len()))
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })?;
    // The prefixes are checked here as well as by qualified(), because they
    // are copied into namespace attributes verbatim.
    if p_prefix > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::NamespacePrefixes,
            p_prefix,
            limits.max_string_bytes(),
        ));
    }
    if r_prefix > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::NamespacePrefixes,
            r_prefix,
            limits.max_string_bytes(),
        ));
    }
    for show in &list.shows {
        length = length
            .checked_add(render_new_show_len(layout, show, limits)?)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
    }
    length
        .checked_add(2 + name + 1)
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })
}

fn expand_empty_custom_list(
    source: &[u8],
    layout: &Layout,
    custom: &CustomListLayout,
    list: &List,
    limits: Limits,
) -> Result<String> {
    if list.is_empty() {
        if custom.span.end.saturating_sub(custom.span.start) > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                custom.span.end.saturating_sub(custom.span.start),
                limits.max_output_bytes(),
            ));
        }
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
    let mut value = BoundedString::new(limits, "custom-show XML output");
    value.push_str(opening)?;
    value.push_str(">")?;
    for show in &list.shows {
        value.push_str(&render_new_show(layout, show, limits)?)?;
    }
    value.push_str("</")?;
    value.push_str(&custom.qualified_name)?;
    value.push_str(">")?;
    Ok(value.finish())
}

fn expand_empty_custom_list_len(
    source: &[u8],
    layout: &Layout,
    custom: &CustomListLayout,
    list: &List,
    limits: Limits,
) -> Result<usize> {
    if list.is_empty() {
        return Ok(custom.span.end - custom.span.start);
    }
    let source = &source[custom.span.start..custom.span.end];
    let close = source
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("empty custom-show list has no closing bracket"))?;
    let opening = std::str::from_utf8(&source[..close])
        .map_err(|error| Error::Xml(error.to_string()))?
        .trim_end_matches('/');
    let mut length = opening
        .len()
        .checked_add(1)
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })?;
    for show in &list.shows {
        length = length
            .checked_add(render_new_show_len(layout, show, limits)?)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
    }
    length
        .checked_add(2 + custom.qualified_name.len() + 1)
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })
}

fn render_show(
    source: &[u8],
    layout: &Layout,
    original_layout: Option<&ShowLayout>,
    original: Option<&Show>,
    show: &Show,
    limits: Limits,
) -> Result<String> {
    match (original_layout, original) {
        (Some(original_layout), Some(original)) if original.id == show.id => {
            render_existing_show(source, layout, original_layout, original, show, limits)
        },
        _ => render_new_show(layout, show, limits),
    }
}

fn render_show_len(
    source: &[u8],
    layout: &Layout,
    original_layout: Option<&ShowLayout>,
    original: Option<&Show>,
    show: &Show,
    limits: Limits,
) -> Result<usize> {
    match (original_layout, original) {
        (Some(original_layout), Some(original)) if original.id == show.id => {
            render_existing_show_len(source, layout, original_layout, original, show, limits)
        },
        _ => render_new_show_len(layout, show, limits),
    }
}

fn render_new_show_len(layout: &Layout, show: &Show, limits: Limits) -> Result<usize> {
    let show_name = qualified_len(&layout.p_prefix, "custShow", limits)?;
    let list_name = qualified_len(&layout.p_prefix, "sldLst", limits)?;
    check_string(&show.name, limits, "custom-show name")?;
    let escaped_name = escaped_xml_len(&show.name)?;
    let mut length = 1usize
        .checked_add(show_name)
        .and_then(|value| value.checked_add(7)) // ` name="`
        .and_then(|value| value.checked_add(escaped_name))
        .and_then(|value| value.checked_add(6)) // `" id="`
        .and_then(|value| value.checked_add(decimal_len(show.id)))
        .and_then(|value| value.checked_add(3)) // `"><` after the id
        .and_then(|value| value.checked_add(list_name))
        .and_then(|value| value.checked_add(1)) // `>` after sldLst
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })?;
    for slide_id in &show.slide_ids {
        length = length
            .checked_add(render_new_slide_len(layout, *slide_id, limits)?)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
    }
    length = length
        .checked_add(2 + list_name + 1) // `</sldLst>`
        .and_then(|value| value.checked_add(2 + show_name + 1)) // `</custShow>`
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })?;
    Ok(length)
}

fn render_new_slide_len(layout: &Layout, slide_id: u32, limits: Limits) -> Result<usize> {
    let relationship_id = layout
        .slide_id_to_relationship
        .get(&slide_id)
        .ok_or_else(|| invalid("custom show references an undeclared slide"))?;
    check_string(relationship_id, limits, "custom-show relationship ID")?;
    let slide_name = qualified_len(&layout.p_prefix, "sld", limits)?;
    let escaped_relationship = escaped_xml_len(relationship_id)?;
    1usize
        .checked_add(slide_name)
        .and_then(|value| value.checked_add(1)) // space
        .and_then(|value| value.checked_add(layout.r_prefix.len()))
        .and_then(|value| value.checked_add(5)) // `:id="`
        .and_then(|value| value.checked_add(escaped_relationship))
        .and_then(|value| value.checked_add(3)) // `"/>`
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })
}

fn render_existing_show_len(
    source: &[u8],
    layout: &Layout,
    original_layout: &ShowLayout,
    original: &Show,
    show: &Show,
    limits: Limits,
) -> Result<usize> {
    let offset = original_layout.span.start;
    let source_show = &source[offset..original_layout.span.end];
    let mut removed = 0usize;
    let mut added = 0usize;
    if original.name != show.name {
        check_string(&show.name, limits, "custom-show name")?;
        removed = removed
            .checked_add(original_layout.name_attribute.end - original_layout.name_attribute.start)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
        added = added
            .checked_add(escaped_xml_len(&show.name)?)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
    }
    if original.slide_ids != show.slide_ids {
        if let Some(slide_list) = original_layout.slide_list.as_ref() {
            removed = removed
                .checked_add(slide_list.span.end - slide_list.span.start)
                .ok_or(Error::CapacityOverflow {
                    resource: "custom-show XML output",
                })?;
            added = added
                .checked_add(render_slide_list_len(
                    source,
                    layout,
                    slide_list,
                    &show.slide_ids,
                    limits,
                )?)
                .ok_or(Error::CapacityOverflow {
                    resource: "custom-show XML output",
                })?;
        } else if !show.slide_ids.is_empty() {
            added = added
                .checked_add(render_new_slide_list_len(layout, &show.slide_ids, limits)?)
                .ok_or(Error::CapacityOverflow {
                    resource: "custom-show XML output",
                })?;
        }
    }
    checked_output_len(
        source_show.len(),
        removed,
        added,
        limits,
        "custom-show XML output",
    )
}

fn render_slide_list_len(
    source: &[u8],
    layout: &Layout,
    slide_list: &SlideListLayout,
    slide_ids: &[u32],
    limits: Limits,
) -> Result<usize> {
    let source_list = &source[slide_list.span.start..slide_list.span.end];
    let slot_count = slide_list
        .children
        .iter()
        .filter(|child| child.local == "sld")
        .count();
    let mut removed = 0usize;
    let mut added = 0usize;
    for (index, slot) in slide_list
        .children
        .iter()
        .filter(|child| child.local == "sld")
        .enumerate()
    {
        removed = removed.checked_add(slot.span.end - slot.span.start).ok_or(
            Error::CapacityOverflow {
                resource: "custom-show XML output",
            },
        )?;
        if let Some(slide_id) = slide_ids.get(index) {
            let existing = layout
                .slide_id_to_relationship
                .get(slide_id)
                .and_then(|relationship_id| {
                    slide_list
                        .slots
                        .iter()
                        .find(|candidate| &candidate.relationship_id == relationship_id)
                })
                .map(|candidate| candidate.span.end - candidate.span.start);
            let replacement_len = match existing {
                Some(length) => length,
                None => render_new_slide_len(layout, *slide_id, limits)?,
            };
            added = added
                .checked_add(replacement_len)
                .ok_or(Error::CapacityOverflow {
                    resource: "custom-show XML output",
                })?;
        }
    }
    if slide_ids.len() > slot_count {
        for slide_id in &slide_ids[slot_count..] {
            added = added
                .checked_add(render_new_slide_len(layout, *slide_id, limits)?)
                .ok_or(Error::CapacityOverflow {
                    resource: "custom-show XML output",
                })?;
        }
    }
    checked_output_len(
        source_list.len(),
        removed,
        added,
        limits,
        "custom-show XML output",
    )
}

fn render_new_slide_list_len(layout: &Layout, slide_ids: &[u32], limits: Limits) -> Result<usize> {
    let name = qualified_len(&layout.p_prefix, "sldLst", limits)?;
    let mut length = 1usize
        .checked_add(name)
        .and_then(|value| value.checked_add(1))
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })?;
    for slide_id in slide_ids {
        length = length
            .checked_add(render_new_slide_len(layout, *slide_id, limits)?)
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML output",
            })?;
    }
    length
        .checked_add(2 + name + 1)
        .ok_or(Error::CapacityOverflow {
            resource: "custom-show XML output",
        })
}

fn render_new_show(layout: &Layout, show: &Show, limits: Limits) -> Result<String> {
    let show_name = qualified(&layout.p_prefix, "custShow", limits)?;
    let list_name = qualified(&layout.p_prefix, "sldLst", limits)?;
    check_string(&show.name, limits, "custom-show name")?;
    let mut value = BoundedString::new(limits, "custom-show XML output");
    value.push_str("<")?;
    value.push_str(&show_name)?;
    value.push_str(" name=\"")?;
    value.push_escaped_xml(&show.name)?;
    value.write_fmt(format_args!(r#"" id="{}"><{}>"#, show.id, list_name))?;
    for slide_id in &show.slide_ids {
        value.push_str(&render_new_slide(layout, *slide_id, limits)?)?;
    }
    value.push_str("</")?;
    value.push_str(&list_name)?;
    value.push_str("></")?;
    value.push_str(&show_name)?;
    value.push_str(">")?;
    Ok(value.finish())
}

fn render_new_slide(layout: &Layout, slide_id: u32, limits: Limits) -> Result<String> {
    let relationship_id = layout
        .slide_id_to_relationship
        .get(&slide_id)
        .ok_or_else(|| invalid("custom show references an undeclared slide"))?;
    check_string(relationship_id, limits, "custom-show relationship ID")?;
    let mut value = BoundedString::new(limits, "custom-show XML output");
    value.push_str("<")?;
    value.push_str(&qualified(&layout.p_prefix, "sld", limits)?)?;
    value.push_str(" ")?;
    value.push_str(&layout.r_prefix)?;
    value.push_str(":id=\"")?;
    value.push_escaped_xml(relationship_id)?;
    value.push_str("\"/>")?;
    Ok(value.finish())
}

fn render_existing_show(
    source: &[u8],
    layout: &Layout,
    original_layout: &ShowLayout,
    original: &Show,
    show: &Show,
    limits: Limits,
) -> Result<String> {
    let offset = original_layout.span.start;
    let source_show = &source[offset..original_layout.span.end];
    let mut replacements = Vec::new();
    if original.name != show.name {
        check_string(&show.name, limits, "custom-show name")?;
        let mut name = BoundedString::new(limits, "custom-show XML output");
        name.push_escaped_xml(&show.name)?;
        replacements.push(Replacement {
            range: original_layout.name_attribute.start - offset
                ..original_layout.name_attribute.end - offset,
            value: name.finish().into_bytes(),
        });
    }
    if original.slide_ids != show.slide_ids {
        if let Some(slide_list) = original_layout.slide_list.as_ref() {
            replacements.push(Replacement {
                range: slide_list.span.start - offset..slide_list.span.end - offset,
                value: render_slide_list(source, layout, slide_list, &show.slide_ids, limits)?
                    .into_bytes(),
            });
        } else if !show.slide_ids.is_empty() {
            replacements.push(Replacement {
                range: original_layout.span.close_start - offset
                    ..original_layout.span.close_start - offset,
                value: render_new_slide_list(layout, &show.slide_ids, limits)?.into_bytes(),
            });
        }
    }
    if replacements.is_empty() {
        if source_show.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                source_show.len(),
                limits.max_output_bytes(),
            ));
        }
        return Ok(String::from_utf8_lossy(source_show).into_owned());
    }
    String::from_utf8(replace_ranges(source_show, &replacements, limits)?)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn render_slide_list(
    source: &[u8],
    layout: &Layout,
    slide_list: &SlideListLayout,
    slide_ids: &[u32],
    limits: Limits,
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
                    || render_new_slide(layout, *slide_id, limits),
                    |raw| {
                        if raw.len() > limits.max_output_bytes() {
                            return Err(limit(
                                crate::opened::CodecLimitKind::OutputBytes,
                                raw.len(),
                                limits.max_output_bytes(),
                            ));
                        }
                        Ok(String::from_utf8_lossy(raw).into_owned())
                    },
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
        let mut rendered = BoundedString::new(limits, "custom-show XML output");
        for slide_id in &slide_ids[slots.len()..] {
            rendered.push_str(&render_new_slide(layout, *slide_id, limits)?)?;
        }
        replacements.push(Replacement {
            range: slide_list.span.close_start - slide_list.span.start
                ..slide_list.span.close_start - slide_list.span.start,
            value: rendered.finish().into_bytes(),
        });
    }
    String::from_utf8(replace_ranges(source_list, &replacements, limits)?)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn render_new_slide_list(layout: &Layout, slide_ids: &[u32], limits: Limits) -> Result<String> {
    let name = qualified(&layout.p_prefix, "sldLst", limits)?;
    let mut value = BoundedString::new(limits, "custom-show XML output");
    value.push_str("<")?;
    value.push_str(&name)?;
    value.push_str(">")?;
    for slide_id in slide_ids {
        value.push_str(&render_new_slide(layout, *slide_id, limits)?)?;
    }
    value.push_str("</")?;
    value.push_str(&name)?;
    value.push_str(">")?;
    Ok(value.finish())
}

fn qualified(prefix: &str, local: &str, limits: Limits) -> Result<String> {
    if prefix.is_empty() {
        check_string(local, limits, "custom-show XML element name")?;
        Ok(local.to_owned())
    } else {
        let length = prefix
            .len()
            .checked_add(1)
            .and_then(|length| length.checked_add(local.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML element name",
            })?;
        if length > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::NameBytes,
                length,
                limits.max_string_bytes(),
            ));
        }
        let mut value = String::new();
        value
            .try_reserve(length)
            .map_err(|source| Error::Allocation {
                resource: "custom-show XML element name",
                source,
            })?;
        value.push_str(prefix);
        value.push(':');
        value.push_str(local);
        Ok(value)
    }
}

fn qualified_len(prefix: &str, local: &str, limits: Limits) -> Result<usize> {
    let length = if prefix.is_empty() {
        local.len()
    } else {
        prefix
            .len()
            .checked_add(1)
            .and_then(|value| value.checked_add(local.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "custom-show XML element name",
            })?
    };
    if length > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::NameBytes,
            length,
            limits.max_string_bytes(),
        ));
    }
    Ok(length)
}

fn decimal_len(value: u32) -> usize {
    if value < 10 {
        1
    } else if value < 100 {
        2
    } else if value < 1_000 {
        3
    } else if value < 10_000 {
        4
    } else if value < 100_000 {
        5
    } else if value < 1_000_000 {
        6
    } else if value < 10_000_000 {
        7
    } else if value < 100_000_000 {
        8
    } else if value < 1_000_000_000 {
        9
    } else {
        10
    }
}

fn replace_ranges(source: &[u8], replacements: &[Replacement], limits: Limits) -> Result<Vec<u8>> {
    let mut replacements = replacements.iter().collect::<Vec<_>>();
    replacements.sort_by_key(|replacement| (replacement.range.start, replacement.range.end));
    let mut removed = 0usize;
    let mut added = 0usize;
    for pair in replacements.windows(2) {
        if pair[0].range.end > pair[1].range.start {
            return Err(invalid("overlapping custom-show XML patch ranges"));
        }
    }
    for replacement in &replacements {
        if replacement.range.start > replacement.range.end || replacement.range.end > source.len() {
            return Err(invalid("custom-show patch range is outside the source"));
        }
        removed = removed
            .checked_add(replacement.range.end - replacement.range.start)
            .ok_or(Error::CapacityOverflow {
                resource: "patched custom-show PresentationML source bytes",
            })?;
        added = added
            .checked_add(replacement.value.len())
            .ok_or(Error::CapacityOverflow {
                resource: "patched custom-show PresentationML source bytes",
            })?;
    }
    let output_len = checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "patched custom-show PresentationML source bytes",
    )?;
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched custom-show PresentationML source bytes",
            source,
        })?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.value);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::opened::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::opened::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
