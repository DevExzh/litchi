#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the host scanner keeps its bounded XML state together"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "reader events are refined after each bounded validation step"
)]

use std::collections::HashSet;
use std::ops::Range;

use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, OffsetLimits, active_offsets};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::Limits;
use super::codec::{self, Anchor, Form};
use crate::package::story::StoryDialect;
use crate::{Error, Result};

const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_WORDPROCESSING_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_WORDPROCESSING_DRAWING: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const WORD_2010_WORDML: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const WORDPROCESSING_INK: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WORDPROCESSING_CANVAS: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WORDPROCESSING_GROUP: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_MCE_MARKED_BYTES: usize = 128 * 1024 * 1024;
const MAX_MCE_OUTPUT_BYTES: usize = 128 * 1024 * 1024;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_RELATIONSHIP_VALUE_BYTES: usize = 1024;
const MAX_MCE_MARKER_BYTES_PER_ANCHOR: usize = 64;
const MIN_COORDINATE: i64 = -27_273_042_329_600;
const MAX_COORDINATE: i64 = 27_273_042_316_900;

/// One source-bound active Ink host in a story part.
///
/// The byte ranges always address the original source buffer. `removable` is
/// deliberately separate from the ranges: an active host can be inventoried
/// while refusing a source splice that would discard a sibling or stale MCE
/// fallback. Callers must check it before using either optional removal range.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Host {
    pub(crate) anchor: Anchor,
    pub(crate) anchor_span: Range<usize>,
    pub(crate) removal_start_tag: Option<Range<usize>>,
    pub(crate) removal_span: Option<Range<usize>>,
    pub(crate) removable: bool,
}

/// Capture active Ink host spans without normalizing or copying story XML.
///
/// The existing codec remains authoritative for active-anchor semantics. This
/// pass only retains source offsets and conservative removal eligibility. Base
/// anchors inside MCE are intentionally non-removable until an entire
/// alternative can be edited atomically.
pub(crate) fn capture(xml: &[u8], dialect: StoryDialect, limits: Limits) -> Result<Vec<Host>> {
    let limits = limits.validate()?;
    let anchors = codec::scan(
        xml,
        dialect,
        limits.max_xml_nodes,
        limits.max_xml_depth,
        limits.max_annotations,
    )?;
    if anchors.is_empty() {
        return Ok(Vec::new());
    }

    let candidates = candidate_offsets(xml, dialect, limits)?;
    let selected = select_active_offsets(xml, &candidates, limits)?;
    if selected.len() != anchors.len() {
        return Err(invalid(
            "DOCX ink host capture diverged from active scanner",
        ));
    }
    capture_selected(xml, dialect, limits, &anchors, &candidates, &selected)
}

/// Return bounded, unique decoded attribute tokens from a complete story.
///
/// Relationship IDs are normally carried by `r:id`, `r:embed`, `r:link`, or
/// VML `o:relid`; collecting every bounded attribute token also protects
/// unknown producer attributes from an unsafe relationship cleanup. A leading
/// `#` is removed for conservative reference matching. The result includes
/// inactive MCE branches because it is a source census, not an active view.
pub(crate) fn reference_values(xml: &[u8], limits: Limits) -> Result<Vec<String>> {
    let limits = limits.validate()?;
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("story XML bytes", xml.len(), MAX_XML_BYTES));
    }

    let mut reader = reader(xml);
    let mut stack_depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_started = false;
    let mut values = HashSet::new();

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_envelope_start(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack_depth,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                )?;
                collect_attributes(&element, resolver, &mut values, limits)?;
            },
            Event::Empty(element) => {
                observe_envelope_empty(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack_depth,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    &mut root_closed,
                    &mut prolog_started,
                )?;
                collect_attributes(&element, resolver, &mut values, limits)?;
            },
            Event::End(element) => {
                if stack_depth == 0 {
                    return Err(invalid("DOCX story has an unexpected end element"));
                }
                super::xml::text(element.name().as_ref())?;
                stack_depth -= 1;
                if stack_depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                super::xml::text(text.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "DOCX story has non-whitespace text outside its root",
                    ));
                }
            },
            Event::CData(data) => {
                super::xml::text(data.as_ref())?;
                if !root_seen || root_closed {
                    return Err(invalid("DOCX story has CDATA outside its root"));
                }
            },
            Event::Comment(comment) => {
                super::xml::text(comment.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::GeneralRef(reference) => {
                super::xml::reference(reference.as_ref())?;
                if !root_seen || root_closed {
                    return Err(invalid("DOCX story has a reference outside its root"));
                }
            },
            Event::Decl(declaration) => {
                super::xml::declaration(&declaration)?;
                if declaration_seen || prolog_started || root_seen {
                    return Err(invalid(
                        "DOCX story has an XML declaration outside its prolog",
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => return Err(invalid("DOCX story rejects DTDs")),
            Event::PI(_) => return Err(invalid("DOCX story rejects processing instructions")),
            Event::Eof => {
                if !root_seen || !root_closed || stack_depth != 0 {
                    return Err(invalid("DOCX story has an unterminated XML root"));
                }
                break;
            },
        }
    }

    let mut result = Vec::new();
    result
        .try_reserve_exact(values.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink relationship reference values",
            source,
        })?;
    result.extend(values);
    result.sort_unstable();
    Ok(result)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum HostKind {
    AlternateContent,
    Container,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum GraphicKind {
    Ink,
    Canvas,
    Group,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Other,
    Run,
    AlternateContent,
    Choice,
    Fallback,
    Drawing,
    Inline,
    Anchor,
    Graphic,
    GraphicData(GraphicKind),
    Canvas,
    Group,
    GroupCnvGrpSpPr,
    GroupSpPr,
    GroupXfrm,
    GroupOff,
    GroupExt,
    GroupChOff,
    GroupChExt,
    Pict,
    VmlShape,
    VmlObject,
    VmlAllowed,
    VmlUnknown,
    ContentPart(Form),
}

#[derive(Clone, Copy, Debug, Default)]
struct Children {
    unmodeled_text: bool,
    metadata_bad: bool,
    total: usize,
    choice: usize,
    fallback: usize,
    drawing: usize,
    inline: usize,
    anchor: usize,
    graphic: usize,
    graphic_data: usize,
    canvas: usize,
    group: usize,
    group_cnv_grp_sp_pr: usize,
    group_sp_pr: usize,
    group_xfrm: usize,
    group_off: usize,
    group_ext: usize,
    group_ch_off: usize,
    group_ch_ext: usize,
    group_order_next: u8,
    group_order_bad: bool,
    group_transform_order_next: u8,
    group_transform_order_bad: bool,
    content_part: usize,
    pict: usize,
    vml_shape: usize,
    vml_object: usize,
    vml_unknown: usize,
    other: usize,
}

#[derive(Clone, Copy, Debug)]
struct Frame {
    kind: Kind,
    start: usize,
    start_tag_end: usize,
    children: Children,
    host_id: Option<usize>,
    selected: Option<usize>,
}

#[derive(Clone, Debug)]
struct Aggregate {
    kind: HostKind,
    start: usize,
    start_tag_end: usize,
    end: Option<usize>,
    children: Children,
    choice: Option<Children>,
    fallback: Option<Children>,
    active_count: usize,
    path_bad: bool,
}

#[derive(Clone, Debug)]
struct Work {
    anchor: Anchor,
    form: Form,
    anchor_start: usize,
    anchor_start_tag: Range<usize>,
    anchor_end: Option<usize>,
    alternate_id: Option<usize>,
    container_id: Option<usize>,
    base_direct: bool,
    has_mce: bool,
    container_count: usize,
}

fn capture_selected(
    xml: &[u8],
    dialect: StoryDialect,
    limits: Limits,
    anchors: &[Anchor],
    candidates: &[u32],
    selected: &[u32],
) -> Result<Vec<Host>> {
    let mut reader = reader(xml);
    let mut stack = Vec::new();
    stack
        .try_reserve(limits.max_xml_depth)
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink host XML frames",
            source,
        })?;
    let mut aggregates = Vec::new();
    aggregates
        .try_reserve(anchors.len().saturating_mul(2))
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink host span aggregates",
            source,
        })?;
    let mut work = Vec::new();
    work.try_reserve_exact(anchors.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink host spans",
            source,
        })?;
    let mut nodes = 0usize;
    let mut candidate_index = 0usize;
    let mut selected_index = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_position = position(reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_node(&mut nodes, limits.max_xml_nodes)?;
                let (start, start_tag_end) = opening_span(event_position, &element, false, xml)?;
                let parent = semantic_parent(&stack);
                let kind = classify_kind(&namespace, &element, parent, resolver, dialect);
                if !validate_group_metadata(&element, kind, resolver)? {
                    note_group_metadata_bad(&mut stack, &mut aggregates);
                }
                note_child(&mut stack, kind);
                if let Some(parent) = stack.last() {
                    observe_path_child(&stack, parent.kind, parent.children, &mut aggregates);
                }
                let selected_anchor = candidate_form(
                    &namespace,
                    &element,
                    parent,
                    dialect,
                    start as u32,
                    candidates,
                    &mut candidate_index,
                    selected,
                    &mut selected_index,
                )?;
                let selected_id = if let Some((index, form)) = selected_anchor {
                    let anchor = anchors
                        .get(index)
                        .ok_or_else(|| invalid("DOCX ink host anchor inventory diverged"))?
                        .clone();
                    if anchor.form != form {
                        return Err(invalid("DOCX ink host form inventory diverged"));
                    }
                    let has_mce = stack.iter().any(|frame| is_mce(frame.kind));
                    let base_direct = form == Form::Base
                        && stack.last().is_some_and(|frame| frame.kind == Kind::Run)
                        && !has_mce;
                    let alternate_count = stack
                        .iter()
                        .filter(|frame| frame.kind == Kind::AlternateContent)
                        .count();
                    let alternate_id = if form == Form::Base {
                        None
                    } else {
                        ensure_host(&mut stack, &mut aggregates, HostKind::AlternateContent)?
                    };
                    let container_count = stack
                        .iter()
                        .filter(|frame| matches!(frame.kind, Kind::Canvas | Kind::Group))
                        .count();
                    let container_id = if form == Form::GenericDrawing {
                        ensure_host(&mut stack, &mut aggregates, HostKind::Container)?
                    } else {
                        None
                    };
                    if let Some(id) = alternate_id {
                        aggregates[id].active_count = aggregates[id].active_count.saturating_add(1);
                        aggregates[id].path_bad |= alternate_count != 1;
                    }
                    if let Some(id) = container_id {
                        aggregates[id].active_count = aggregates[id].active_count.saturating_add(1);
                        aggregates[id].path_bad |= container_count != 1;
                    }
                    retroactive_path_check(&stack, &mut aggregates);
                    let work_index = work.len();
                    work.push(Work {
                        anchor,
                        form,
                        anchor_start: start,
                        anchor_start_tag: start..start_tag_end,
                        anchor_end: None,
                        alternate_id,
                        container_id,
                        base_direct,
                        has_mce,
                        container_count,
                    });
                    Some((work_index, form))
                } else {
                    None
                };
                let frame = Frame {
                    kind: selected_anchor.map_or(kind, |(_, form)| Kind::ContentPart(form)),
                    start,
                    start_tag_end,
                    children: Children::default(),
                    host_id: None,
                    selected: selected_id.map(|(index, _)| index),
                };
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX ink host XML frames",
                    source,
                })?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                observe_node(&mut nodes, limits.max_xml_nodes)?;
                let (start, end) = opening_span(event_position, &element, true, xml)?;
                let parent = semantic_parent(&stack);
                let kind = classify_kind(&namespace, &element, parent, resolver, dialect);
                if !validate_group_metadata(&element, kind, resolver)?
                    || matches!(kind, Kind::GroupSpPr | Kind::GroupXfrm)
                {
                    note_group_metadata_bad(&mut stack, &mut aggregates);
                }
                note_child(&mut stack, kind);
                if let Some(parent) = stack.last() {
                    observe_path_child(&stack, parent.kind, parent.children, &mut aggregates);
                }
                if kind == Kind::Pict {
                    if let Some(host_id) = stack.iter().rev().find_map(|frame| {
                        (frame.kind == Kind::AlternateContent)
                            .then_some(frame.host_id)
                            .flatten()
                    }) {
                        aggregates[host_id].path_bad = true;
                    }
                }
                if let Some((index, form)) = candidate_form(
                    &namespace,
                    &element,
                    parent,
                    dialect,
                    start as u32,
                    candidates,
                    &mut candidate_index,
                    selected,
                    &mut selected_index,
                )? {
                    let anchor = anchors
                        .get(index)
                        .ok_or_else(|| invalid("DOCX ink host anchor inventory diverged"))?
                        .clone();
                    if anchor.form != form {
                        return Err(invalid("DOCX ink host form inventory diverged"));
                    }
                    let has_mce = stack.iter().any(|frame| is_mce(frame.kind));
                    let base_direct = form == Form::Base
                        && stack.last().is_some_and(|frame| frame.kind == Kind::Run)
                        && !has_mce;
                    let alternate_count = stack
                        .iter()
                        .filter(|frame| frame.kind == Kind::AlternateContent)
                        .count();
                    let alternate_id = if form == Form::Base {
                        None
                    } else {
                        ensure_host(&mut stack, &mut aggregates, HostKind::AlternateContent)?
                    };
                    let container_count = stack
                        .iter()
                        .filter(|frame| matches!(frame.kind, Kind::Canvas | Kind::Group))
                        .count();
                    let container_id = if form == Form::GenericDrawing {
                        ensure_host(&mut stack, &mut aggregates, HostKind::Container)?
                    } else {
                        None
                    };
                    if let Some(id) = alternate_id {
                        aggregates[id].active_count = aggregates[id].active_count.saturating_add(1);
                        aggregates[id].path_bad |= alternate_count != 1;
                    }
                    if let Some(id) = container_id {
                        aggregates[id].active_count = aggregates[id].active_count.saturating_add(1);
                        aggregates[id].path_bad |= container_count != 1;
                    }
                    retroactive_path_check(&stack, &mut aggregates);
                    work.push(Work {
                        anchor,
                        form,
                        anchor_start: start,
                        anchor_start_tag: start..end,
                        anchor_end: Some(end),
                        alternate_id,
                        container_id,
                        base_direct,
                        has_mce,
                        container_count,
                    });
                }
            },
            Event::End(element) => {
                let end = end_position(event_position, &element, xml)?;
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("DOCX ink host has an unexpected end element"))?;
                if let Some(index) = frame.selected {
                    work[index].anchor_end = Some(end);
                }
                if let Some(host_id) = frame.host_id {
                    aggregates[host_id].end = Some(end);
                    aggregates[host_id].children = frame.children;
                }
                let frame_bad = path_frame_bad(frame.kind, frame.children);
                if frame_bad {
                    if let Some(parent) = stack.last_mut() {
                        if propagates_path_bad_to_parent(parent.kind) {
                            parent.children.metadata_bad = true;
                        }
                    }
                    if let Some(host_id) = stack.iter().rev().find_map(|parent| parent.host_id) {
                        aggregates[host_id].path_bad = true;
                    }
                    if let Some(host_id) = frame.host_id {
                        aggregates[host_id].path_bad = true;
                    }
                }
                if frame.kind == Kind::Pict && pict_children_bad(frame.children) {
                    if let Some(host_id) = stack.iter().rev().find_map(|parent| {
                        (parent.kind == Kind::AlternateContent)
                            .then_some(parent.host_id)
                            .flatten()
                    }) {
                        aggregates[host_id].path_bad = true;
                    }
                }
                if matches!(frame.kind, Kind::Choice | Kind::Fallback) {
                    if let Some(host_id) = stack.iter().rev().find_map(|parent| {
                        (parent.kind == Kind::AlternateContent)
                            .then_some(parent.host_id)
                            .flatten()
                    }) {
                        match frame.kind {
                            Kind::Choice => aggregates[host_id].choice = Some(frame.children),
                            Kind::Fallback => aggregates[host_id].fallback = Some(frame.children),
                            _ => {},
                        }
                    }
                }
            },
            Event::Text(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    note_unmodeled_text(&mut stack, &mut aggregates);
                }
            },
            Event::CData(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    note_unmodeled_text(&mut stack, &mut aggregates);
                }
            },
            Event::GeneralRef(_) => note_unmodeled_text(&mut stack, &mut aggregates),
            Event::Eof => break,
            _ => {},
        }
    }

    if candidate_index != candidates.len() || selected_index != anchors.len() || !stack.is_empty() {
        return Err(invalid("DOCX ink host span inventory diverged"));
    }

    let mut result = Vec::new();
    result
        .try_reserve_exact(work.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink host results",
            source,
        })?;
    for item in work {
        let anchor_end = item
            .anchor_end
            .ok_or_else(|| invalid("DOCX ink host anchor has no closing span"))?;
        let anchor_span = item.anchor_start..anchor_end;
        let mut removable = item.form == Form::Base && item.base_direct && !item.has_mce;
        let mut removal_start_tag = None;
        let mut removal_span = None;
        if item.form != Form::Base {
            removable = item.container_count <= 1;
            if let Some(alternate_id) = item.alternate_id {
                let aggregate = aggregates
                    .get(alternate_id)
                    .ok_or_else(|| invalid("DOCX ink host alternate aggregate is missing"))?;
                let valid = aggregate.kind == HostKind::AlternateContent
                    && aggregate.active_count == 1
                    && !aggregate.path_bad
                    && aggregate.children.total == 2
                    && aggregate.children.choice == 1
                    && aggregate.children.fallback == 1
                    && aggregate
                        .choice
                        .is_some_and(|children| children.total == 1 && children.drawing == 1)
                    && aggregate
                        .fallback
                        .is_some_and(|children| children.total == 1 && children.pict == 1);
                removable &= valid;
                if valid {
                    let end = aggregate
                        .end
                        .ok_or_else(|| invalid("DOCX ink host alternate has no end span"))?;
                    removal_start_tag = Some(aggregate.start..aggregate.start_tag_end);
                    removal_span = Some(aggregate.start..end);
                }
            } else {
                removable = false;
            }
            if item.form == Form::GenericDrawing {
                let container_id = item
                    .container_id
                    .ok_or_else(|| invalid("DOCX generic ink host has no container"))?;
                let aggregate = aggregates
                    .get(container_id)
                    .ok_or_else(|| invalid("DOCX generic ink host container is missing"))?;
                let valid = aggregate.kind == HostKind::Container
                    && aggregate.active_count == 1
                    && !aggregate.path_bad
                    && (is_single_canvas_slot(aggregate.children)
                        || group_children_valid(aggregate.children));
                removable &= valid;
                if !valid {
                    removal_start_tag = None;
                    removal_span = None;
                }
            }
        }
        if item.form == Form::Base && removable {
            removal_start_tag = Some(item.anchor_start_tag.clone());
            removal_span = Some(anchor_span.clone());
        }
        if !removable {
            removal_start_tag = None;
            removal_span = None;
        }
        result.push(Host {
            anchor: item.anchor,
            anchor_span,
            removal_start_tag,
            removal_span,
            removable,
        });
    }
    Ok(result)
}

fn candidate_form(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    parent: Option<Kind>,
    dialect: StoryDialect,
    position: u32,
    candidates: &[u32],
    candidate_index: &mut usize,
    selected: &[u32],
    selected_index: &mut usize,
) -> Result<Option<(usize, Form)>> {
    let Some(form) = classify_form(namespace, element, parent, dialect) else {
        return Ok(None);
    };
    let expected = candidates
        .get(*candidate_index)
        .copied()
        .ok_or_else(|| invalid("DOCX ink host candidate inventory diverged"))?;
    if expected != position {
        return Err(invalid("DOCX ink host candidate offset diverged"));
    }
    *candidate_index = (*candidate_index).saturating_add(1);
    if selected.get(*selected_index).copied() == Some(position) {
        let active_index = *selected_index;
        *selected_index = (*selected_index).saturating_add(1);
        return Ok(Some((active_index, form)));
    }
    Ok(None)
}

fn candidate_offsets(xml: &[u8], dialect: StoryDialect, limits: Limits) -> Result<Vec<u32>> {
    let mut reader = reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut offsets = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_position = position(reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_node(&mut nodes, limits.max_xml_nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("XML depth", usize::MAX, limits.max_xml_depth))?;
                if depth > limits.max_xml_depth {
                    return Err(limit("XML depth", depth, limits.max_xml_depth));
                }
                if is_candidate(&namespace, &element, dialect) {
                    push_offset(
                        &mut offsets,
                        opening_span(event_position, &element, false, xml)?.0,
                        limits.max_xml_nodes,
                    )?;
                }
            },
            Event::Empty(element) => {
                observe_node(&mut nodes, limits.max_xml_nodes)?;
                if is_candidate(&namespace, &element, dialect) {
                    push_offset(
                        &mut offsets,
                        opening_span(event_position, &element, true, xml)?.0,
                        limits.max_xml_nodes,
                    )?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("DOCX ink host XML depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(offsets)
}

/// Select source offsets that survive the same MCE capability profile as the
/// active Ink host scanner.
///
/// This remains crate-private because the offsets are source identities, not
/// part of the public annotation model.  Placement selectors use the helper
/// for paragraph ordering so MCE branch selection cannot drift between host
/// and authoring scans.
pub(crate) fn select_active_offsets(
    xml: &[u8],
    offsets: &[u32],
    limits: Limits,
) -> Result<Vec<u32>> {
    if offsets.is_empty() {
        return Ok(Vec::new());
    }
    let marked_extra = offsets
        .len()
        .checked_mul(MAX_MCE_MARKER_BYTES_PER_ANCHOR)
        .ok_or_else(|| limit("MCE marked XML bytes", usize::MAX, MAX_MCE_MARKED_BYTES))?;
    let marked_bytes = xml
        .len()
        .checked_add(marked_extra)
        .ok_or_else(|| limit("MCE marked XML bytes", usize::MAX, MAX_MCE_MARKED_BYTES))?;
    if marked_bytes > MAX_MCE_MARKED_BYTES {
        return Err(limit(
            "MCE marked XML bytes",
            marked_bytes,
            MAX_MCE_MARKED_BYTES,
        ));
    }
    let mut capabilities = Capabilities::default();
    for namespace in [
        WORD_2010_WORDML,
        WORDPROCESSING_INK,
        WORDPROCESSING_CANVAS,
        WORDPROCESSING_GROUP,
    ] {
        let namespace = std::str::from_utf8(namespace)
            .map_err(|error| Error::Invalid(format!("invalid fixed MCE namespace: {error}")))?;
        capabilities.understand_namespace(namespace);
    }
    let growth_factor = limits.max_xml_nodes.max(limits.max_xml_depth).max(8);
    let max_output_bytes = marked_bytes
        .checked_mul(growth_factor)
        .ok_or_else(|| limit("MCE output bytes", usize::MAX, MAX_MCE_OUTPUT_BYTES))?
        .min(MAX_MCE_OUTPUT_BYTES);
    let processing = MceLimits {
        max_input_bytes: marked_bytes,
        max_output_bytes,
        max_depth: limits.max_xml_depth,
        max_namespace_bindings: MAX_NAMESPACE_DECLARATIONS,
        max_directive_tokens: limits.max_xml_nodes,
        max_choices_per_alternate: limits.max_xml_nodes.max(1),
    };
    active_offsets(
        xml,
        offsets,
        &capabilities,
        &OffsetLimits {
            max_source_bytes: MAX_XML_BYTES,
            max_offsets: offsets.len(),
            max_marked_bytes: MAX_MCE_MARKED_BYTES,
            processing,
        },
    )
    .map_err(Error::from)
}

fn classify_form(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    parent: Option<Kind>,
    dialect: StoryDialect,
) -> Option<Form> {
    if element.local_name().as_ref() != b"contentPart" {
        return None;
    }
    if is_namespace(namespace, dialect_word_namespace(dialect)) {
        return Some(Form::Base);
    }
    if is_namespace(namespace, WORD_2010_WORDML) {
        if matches!(parent, Some(Kind::Canvas | Kind::Group)) {
            Some(Form::GenericDrawing)
        } else {
            Some(Form::Drawing)
        }
    } else {
        None
    }
}

fn classify_kind(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    parent: Option<Kind>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
) -> Kind {
    let local = element.local_name();
    if is_namespace(namespace, MCE_NAMESPACE) {
        return match local.as_ref() {
            b"AlternateContent" => Kind::AlternateContent,
            b"Choice" => Kind::Choice,
            b"Fallback" => Kind::Fallback,
            _ => Kind::Other,
        };
    }
    if is_namespace(namespace, dialect_word_namespace(dialect)) {
        if local.as_ref() == b"r" {
            return Kind::Run;
        }
        if local.as_ref() == b"drawing" {
            return Kind::Drawing;
        }
        if local.as_ref() == b"pict" {
            return Kind::Pict;
        }
        if local.as_ref() == b"contentPart" {
            return Kind::ContentPart(Form::Base);
        }
    }
    if is_wordprocessing_drawing_namespace(namespace, dialect) {
        if local.as_ref() == b"inline" && parent == Some(Kind::Drawing) {
            return Kind::Inline;
        }
        if local.as_ref() == b"anchor" && parent == Some(Kind::Drawing) {
            return Kind::Anchor;
        }
    }
    if is_drawing_namespace(namespace, dialect) {
        if local.as_ref() == b"graphic" && matches!(parent, Some(Kind::Inline | Kind::Anchor)) {
            return Kind::Graphic;
        }
        if local.as_ref() == b"graphicData" {
            return graphic_data_kind(element, resolver)
                .map(Kind::GraphicData)
                .unwrap_or(Kind::Other);
        }
    }
    if is_namespace(namespace, WORDPROCESSING_CANVAS)
        && local.as_ref() == b"wpc"
        && parent == Some(Kind::GraphicData(GraphicKind::Canvas))
    {
        return Kind::Canvas;
    }
    if is_namespace(namespace, WORDPROCESSING_GROUP) {
        return match (parent, local.as_ref()) {
            (
                Some(Kind::GraphicData(GraphicKind::Group) | Kind::Canvas | Kind::Group),
                b"wgp" | b"grpSp",
            ) => Kind::Group,
            (Some(Kind::Group), b"cNvGrpSpPr") => Kind::GroupCnvGrpSpPr,
            (Some(Kind::Group), b"grpSpPr") => Kind::GroupSpPr,
            _ => Kind::Other,
        };
    }
    if is_drawing_namespace(namespace, dialect) {
        return match (parent, local.as_ref()) {
            (Some(Kind::Group), b"cNvGrpSpPr") => Kind::GroupCnvGrpSpPr,
            (Some(Kind::Group), b"grpSpPr") => Kind::GroupSpPr,
            (Some(Kind::GroupSpPr), b"xfrm") => Kind::GroupXfrm,
            (Some(Kind::GroupXfrm), b"off") => Kind::GroupOff,
            (Some(Kind::GroupXfrm), b"ext") => Kind::GroupExt,
            (Some(Kind::GroupXfrm), b"chOff") => Kind::GroupChOff,
            (Some(Kind::GroupXfrm), b"chExt") => Kind::GroupChExt,
            _ => Kind::Other,
        };
    }
    if is_namespace(namespace, VML_NAMESPACE) {
        return match local.as_ref() {
            b"shape" => Kind::VmlShape,
            b"group" | b"rect" | b"roundrect" | b"oval" | b"ellipse" | b"arc" | b"line"
            | b"polyline" | b"polygon" | b"curve" | b"image" | b"textbox" | b"textpath"
            | b"shapetype" => Kind::VmlObject,
            b"imagedata" | b"fill" | b"stroke" | b"shadow" | b"path" | b"formulas" | b"f"
            | b"handles" | b"h" | b"lock" | b"wrap" | b"extrusion" | b"skew" | b"clippath" => {
                Kind::VmlAllowed
            },
            _ => Kind::VmlUnknown,
        };
    }
    if is_namespace(namespace, WORD_2010_WORDML) && local.as_ref() == b"contentPart" {
        return if matches!(parent, Some(Kind::Canvas | Kind::Group)) {
            Kind::ContentPart(Form::GenericDrawing)
        } else {
            Kind::ContentPart(Form::Drawing)
        };
    }
    Kind::Other
}

fn graphic_data_kind(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Option<GraphicKind> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.ok()?;
        if attribute.key.as_ref() != b"uri"
            || !matches!(
                resolver.resolve_attribute(attribute.key).0,
                ResolveResult::Unbound
            )
        {
            continue;
        }
        if result.is_some() {
            return None;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .ok()?;
        result = match value.as_bytes() {
            value if value == WORDPROCESSING_INK => Some(GraphicKind::Ink),
            value if value == WORDPROCESSING_CANVAS => Some(GraphicKind::Canvas),
            value if value == WORDPROCESSING_GROUP => Some(GraphicKind::Group),
            _ => None,
        };
    }
    result
}

fn semantic_parent(stack: &[Frame]) -> Option<Kind> {
    stack
        .iter()
        .rev()
        .map(|frame| frame.kind)
        .find(|kind| !is_mce(*kind))
}

fn is_mce(kind: Kind) -> bool {
    matches!(kind, Kind::AlternateContent | Kind::Choice | Kind::Fallback)
}

fn ensure_host(
    stack: &mut [Frame],
    aggregates: &mut Vec<Aggregate>,
    kind: HostKind,
) -> Result<Option<usize>> {
    let frame_index = stack.iter().rposition(|frame| match kind {
        HostKind::AlternateContent => frame.kind == Kind::AlternateContent,
        HostKind::Container => matches!(frame.kind, Kind::Canvas | Kind::Group),
    });
    let Some(frame_index) = frame_index else {
        return Ok(None);
    };
    if let Some(id) = stack[frame_index].host_id {
        return Ok(Some(id));
    }
    let frame = stack[frame_index];
    aggregates
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink host span aggregates",
            source,
        })?;
    let id = aggregates.len();
    aggregates.push(Aggregate {
        kind,
        start: frame.start,
        start_tag_end: frame.start_tag_end,
        end: None,
        children: Children::default(),
        choice: None,
        fallback: None,
        active_count: 0,
        path_bad: false,
    });
    stack[frame_index].host_id = Some(id);
    Ok(Some(id))
}

// Text in a structural host wrapper cannot be proven to belong to the selected
// object. Count it before an anchor too, so retroactive path validation sees it.
fn note_unmodeled_text(stack: &mut [Frame], aggregates: &mut [Aggregate]) {
    let Some(parent) = stack.last_mut() else {
        return;
    };
    if matches!(
        parent.kind,
        Kind::AlternateContent
            | Kind::Choice
            | Kind::Fallback
            | Kind::Drawing
            | Kind::Inline
            | Kind::Anchor
            | Kind::Graphic
            | Kind::GraphicData(_)
            | Kind::Canvas
            | Kind::Group
            | Kind::GroupCnvGrpSpPr
            | Kind::GroupSpPr
            | Kind::GroupXfrm
            | Kind::GroupOff
            | Kind::GroupExt
            | Kind::GroupChOff
            | Kind::GroupChExt
            | Kind::Pict
            | Kind::VmlShape
    ) {
        parent.children.unmodeled_text = true;
        note_kind(&mut parent.children, Kind::Other);
        if let Some(host_id) = stack.iter().rev().find_map(|frame| frame.host_id) {
            aggregates[host_id].path_bad = true;
        }
    }
}

fn note_child(stack: &mut [Frame], child: Kind) {
    let Some(parent_index) = stack.len().checked_sub(1) else {
        return;
    };
    let parent_kind = stack[parent_index].kind;
    note_kind(&mut stack[parent_index].children, child);
    match parent_kind {
        Kind::Group => note_group_child_order(&mut stack[parent_index].children, child),
        Kind::GroupSpPr => {
            note_group_metadata_order(&mut stack[parent_index].children, child, [Kind::GroupXfrm]);
        },
        Kind::GroupXfrm => note_group_metadata_order(
            &mut stack[parent_index].children,
            child,
            [
                Kind::GroupOff,
                Kind::GroupExt,
                Kind::GroupChOff,
                Kind::GroupChExt,
            ],
        ),
        _ => {},
    }
    if parent_kind != Kind::Pict {
        if let Some(pict) = stack[..parent_index]
            .iter_mut()
            .rev()
            .find(|frame| frame.kind == Kind::Pict)
        {
            note_vml_descendant(&mut pict.children, child);
        }
    }
}

fn note_group_metadata_bad(stack: &mut [Frame], aggregates: &mut [Aggregate]) {
    if let Some(parent) = stack.last_mut() {
        parent.children.metadata_bad = true;
    }
    if let Some(host_id) = stack.iter().rev().find_map(|frame| frame.host_id) {
        aggregates[host_id].path_bad = true;
    }
}

fn note_group_child_order(children: &mut Children, child: Kind) {
    let expected = match children.group_order_next {
        0 => Kind::GroupCnvGrpSpPr,
        1 => Kind::GroupSpPr,
        2 => Kind::ContentPart(Form::GenericDrawing),
        _ => Kind::Other,
    };
    if expected != child {
        children.group_order_bad = true;
    } else {
        children.group_order_next = children.group_order_next.saturating_add(1);
    }
}

fn note_group_metadata_order<const N: usize>(
    children: &mut Children,
    child: Kind,
    expected: [Kind; N],
) {
    let index = usize::from(children.group_transform_order_next);
    if expected.get(index).copied() != Some(child) {
        children.group_transform_order_bad = true;
    } else {
        children.group_transform_order_next = children.group_transform_order_next.saturating_add(1);
    }
}

fn note_vml_descendant(children: &mut Children, child: Kind) {
    match child {
        Kind::VmlShape => {
            children.vml_shape = children.vml_shape.saturating_add(1);
            children.vml_object = children.vml_object.saturating_add(1);
        },
        Kind::VmlObject => children.vml_object = children.vml_object.saturating_add(1),
        Kind::VmlUnknown => children.vml_unknown = children.vml_unknown.saturating_add(1),
        Kind::VmlAllowed => {},
        _ => children.other = children.other.saturating_add(1),
    }
}

fn note_kind(children: &mut Children, child: Kind) {
    children.total = children.total.saturating_add(1);
    match child {
        Kind::Choice => children.choice = children.choice.saturating_add(1),
        Kind::Fallback => children.fallback = children.fallback.saturating_add(1),
        Kind::Drawing => children.drawing = children.drawing.saturating_add(1),
        Kind::Inline => children.inline = children.inline.saturating_add(1),
        Kind::Anchor => children.anchor = children.anchor.saturating_add(1),
        Kind::Graphic => children.graphic = children.graphic.saturating_add(1),
        Kind::GraphicData(_) => children.graphic_data = children.graphic_data.saturating_add(1),
        Kind::Canvas => children.canvas = children.canvas.saturating_add(1),
        Kind::Group => children.group = children.group.saturating_add(1),
        Kind::GroupCnvGrpSpPr => {
            children.group_cnv_grp_sp_pr = children.group_cnv_grp_sp_pr.saturating_add(1)
        },
        Kind::GroupSpPr => children.group_sp_pr = children.group_sp_pr.saturating_add(1),
        Kind::GroupXfrm => children.group_xfrm = children.group_xfrm.saturating_add(1),
        Kind::GroupOff => children.group_off = children.group_off.saturating_add(1),
        Kind::GroupExt => children.group_ext = children.group_ext.saturating_add(1),
        Kind::GroupChOff => children.group_ch_off = children.group_ch_off.saturating_add(1),
        Kind::GroupChExt => children.group_ch_ext = children.group_ch_ext.saturating_add(1),
        Kind::ContentPart(_) => children.content_part = children.content_part.saturating_add(1),
        Kind::Pict => children.pict = children.pict.saturating_add(1),
        Kind::VmlShape => {
            children.vml_shape = children.vml_shape.saturating_add(1);
            children.vml_object = children.vml_object.saturating_add(1);
        },
        Kind::VmlObject => children.vml_object = children.vml_object.saturating_add(1),
        Kind::VmlUnknown => children.vml_unknown = children.vml_unknown.saturating_add(1),
        Kind::VmlAllowed => {},
        Kind::Other | Kind::Run | Kind::AlternateContent => {
            children.other = children.other.saturating_add(1)
        },
    }
}

fn retroactive_path_check(stack: &[Frame], aggregates: &mut [Aggregate]) {
    let Some(host_id) = stack.iter().rev().find_map(|frame| {
        (frame.kind == Kind::AlternateContent)
            .then_some(frame.host_id)
            .flatten()
    }) else {
        return;
    };
    if stack
        .iter()
        .any(|frame| path_frame_bad(frame.kind, frame.children))
    {
        aggregates[host_id].path_bad = true;
    }
}

fn observe_path_child(
    stack: &[Frame],
    parent_kind: Kind,
    children: Children,
    aggregates: &mut [Aggregate],
) {
    let Some(host_id) = stack.iter().rev().find_map(|frame| {
        (frame.kind == Kind::AlternateContent)
            .then_some(frame.host_id)
            .flatten()
    }) else {
        return;
    };
    let bad = children.unmodeled_text
        || children.metadata_bad
        || match parent_kind {
            Kind::AlternateContent => {
                children.total != children.choice + children.fallback
                    || !matches!(children.total, 0..=2) // counts are checked again at host close
            },
            Kind::Choice => children.total != children.drawing,
            Kind::Drawing => children.total != children.inline + children.anchor,
            Kind::Inline | Kind::Anchor => {
                children.graphic > 1 || children.inline > 0 || children.anchor > 0
            },
            Kind::Graphic => children.total != children.graphic_data,
            Kind::GraphicData(_) => {
                children.total != children.content_part + children.canvas + children.group
            },
            Kind::Canvas => children.total != children.content_part,
            Kind::Group => !group_children_valid(children),
            Kind::GroupSpPr => !group_properties_children_valid(children),
            Kind::GroupXfrm => !group_transform_children_valid(children),
            Kind::GroupCnvGrpSpPr
            | Kind::GroupOff
            | Kind::GroupExt
            | Kind::GroupChOff
            | Kind::GroupChExt => children.total != 0,
            Kind::Pict => pict_children_bad(children),
            _ => false,
        }
        || stack
            .iter()
            .any(|frame| frame.kind == Kind::Pict && pict_children_bad(frame.children));
    if bad {
        aggregates[host_id].path_bad = true;
    }
}

fn path_frame_bad(kind: Kind, children: Children) -> bool {
    children.unmodeled_text
        || children.metadata_bad
        || match kind {
            Kind::AlternateContent => children.total != children.choice + children.fallback,
            Kind::Choice => children.total != children.drawing,
            Kind::Drawing => children.total != children.inline + children.anchor,
            Kind::Inline | Kind::Anchor => {
                children.graphic > 1 || children.inline > 0 || children.anchor > 0
            },
            Kind::Graphic => children.total != children.graphic_data,
            Kind::GraphicData(GraphicKind::Ink) => children.total != children.content_part,
            Kind::GraphicData(GraphicKind::Canvas | GraphicKind::Group) => {
                children.total != children.canvas + children.group
            },
            Kind::Canvas => children.total != children.content_part,
            Kind::Group => !group_children_valid(children),
            Kind::GroupSpPr => !group_properties_children_valid(children),
            Kind::GroupXfrm => !group_transform_children_valid(children),
            Kind::GroupCnvGrpSpPr
            | Kind::GroupOff
            | Kind::GroupExt
            | Kind::GroupChOff
            | Kind::GroupChExt => children.total != 0,
            Kind::Pict => pict_children_bad(children),
            _ => false,
        }
}

fn propagates_path_bad_to_parent(kind: Kind) -> bool {
    matches!(
        kind,
        Kind::AlternateContent
            | Kind::Choice
            | Kind::Fallback
            | Kind::Drawing
            | Kind::Inline
            | Kind::Anchor
            | Kind::Graphic
            | Kind::GraphicData(_)
            | Kind::Canvas
            | Kind::Group
            | Kind::GroupCnvGrpSpPr
            | Kind::GroupSpPr
            | Kind::GroupXfrm
            | Kind::GroupOff
            | Kind::GroupExt
            | Kind::GroupChOff
            | Kind::GroupChExt
            | Kind::Pict
            | Kind::VmlShape
    )
}

fn pict_children_bad(children: Children) -> bool {
    children.total != 1
        || children.vml_shape != 1
        || children.vml_object != 1
        || children.vml_unknown != 0
        || children.other != 0
}

fn group_children_valid(children: Children) -> bool {
    !children.metadata_bad
        && !children.group_order_bad
        && children.group_order_next == 3
        && children.total == 3
        && children.group_cnv_grp_sp_pr == 1
        && children.group_sp_pr == 1
        && children.content_part == 1
}

fn group_properties_children_valid(children: Children) -> bool {
    !children.metadata_bad
        && !children.group_transform_order_bad
        && children.group_transform_order_next == 1
        && children.total == 1
        && children.group_xfrm == 1
}

fn group_transform_children_valid(children: Children) -> bool {
    !children.metadata_bad
        && !children.group_transform_order_bad
        && children.group_transform_order_next == 4
        && children.total == 4
        && children.group_off == 1
        && children.group_ext == 1
        && children.group_ch_off == 1
        && children.group_ch_ext == 1
}

const fn is_single_canvas_slot(children: Children) -> bool {
    children.total == 1 && children.content_part == 1
}

fn validate_group_metadata(
    element: &BytesStart<'_>,
    kind: Kind,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    let metadata = match kind {
        Kind::GroupCnvGrpSpPr
        | Kind::GroupSpPr
        | Kind::GroupXfrm
        | Kind::GroupOff
        | Kind::GroupExt
        | Kind::GroupChOff
        | Kind::GroupChExt => kind,
        _ => return Ok(true),
    };
    super::xml::element(element, resolver)?;
    let mut attributes = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        attributes = attributes.saturating_add(1);
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Ok(false);
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        match (metadata, local.as_ref()) {
            (Kind::GroupOff | Kind::GroupChOff, b"x" | b"y") => {
                if !validate_group_coordinate(&value, false)? {
                    return Ok(false);
                }
            },
            (Kind::GroupExt | Kind::GroupChExt, b"cx" | b"cy") => {
                if !validate_group_coordinate(&value, true)? {
                    return Ok(false);
                }
            },
            _ => {
                return Ok(false);
            },
        }
    }
    let expected = match metadata {
        Kind::GroupOff | Kind::GroupChOff | Kind::GroupExt | Kind::GroupChExt => 2,
        Kind::GroupCnvGrpSpPr | Kind::GroupSpPr | Kind::GroupXfrm => 0,
        _ => unreachable!("metadata kind checked above"),
    };
    if attributes != expected {
        return Ok(false);
    }
    Ok(true)
}

fn validate_group_coordinate(value: &str, positive: bool) -> Result<bool> {
    if positive {
        let Ok(value) = value.parse::<u64>() else {
            return Ok(false);
        };
        if value == 0 || value > MAX_COORDINATE as u64 {
            return Ok(false);
        }
    } else {
        let Ok(value) = value.parse::<i64>() else {
            return Ok(false);
        };
        if !(MIN_COORDINATE..=MAX_COORDINATE).contains(&value) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn collect_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    values: &mut HashSet<String>,
    limits: Limits,
) -> Result<()> {
    super::xml::element(element, resolver)?;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        for token in value.split_whitespace() {
            let token = token.strip_prefix('#').unwrap_or(token);
            if token.is_empty() || token.len() > MAX_RELATIONSHIP_VALUE_BYTES {
                continue;
            }
            if values.contains(token) {
                continue;
            }
            if values.len() >= limits.max_relationships {
                return Err(limit(
                    "DOCX ink relationship reference values",
                    values.len().saturating_add(1),
                    limits.max_relationships,
                ));
            }
            let mut owned = String::new();
            owned
                .try_reserve(token.len())
                .map_err(|source| Error::Allocation {
                    resource: "DOCX ink relationship reference value",
                    source,
                })?;
            owned.push_str(token);
            values.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "DOCX ink relationship reference values",
                source,
            })?;
            values.insert(owned);
        }
    }
    Ok(())
}

fn observe_envelope_start(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    depth: &mut usize,
    nodes: &mut usize,
    limits: Limits,
    root_seen: &mut bool,
    root_closed: bool,
    prolog_started: &mut bool,
) -> Result<()> {
    super::xml::element(element, resolver)?;
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "DOCX story element uses an unknown namespace prefix",
        ));
    }
    observe_node(nodes, limits.max_xml_nodes)?;
    *depth = depth
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, limits.max_xml_depth))?;
    if *depth > limits.max_xml_depth {
        return Err(limit("XML depth", *depth, limits.max_xml_depth));
    }
    if *root_seen && root_closed {
        return Err(invalid("DOCX story has more than one XML root"));
    }
    if *depth == 1 {
        if *root_seen {
            return Err(invalid("DOCX story has more than one XML root"));
        }
        *root_seen = true;
    }
    *prolog_started = true;
    Ok(())
}

fn observe_envelope_empty(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    depth: &mut usize,
    nodes: &mut usize,
    limits: Limits,
    root_seen: &mut bool,
    root_closed: &mut bool,
    prolog_started: &mut bool,
) -> Result<()> {
    super::xml::element(element, resolver)?;
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "DOCX story element uses an unknown namespace prefix",
        ));
    }
    observe_node(nodes, limits.max_xml_nodes)?;
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, limits.max_xml_depth))?;
    if next_depth > limits.max_xml_depth {
        return Err(limit("XML depth", next_depth, limits.max_xml_depth));
    }
    if *depth == 0 {
        if *root_seen {
            return Err(invalid("DOCX story has more than one XML root"));
        }
        *root_seen = true;
        *root_closed = true;
    } else if *root_closed {
        return Err(invalid("DOCX story has markup after its XML root"));
    }
    *prolog_started = true;
    Ok(())
}

fn push_offset(offsets: &mut Vec<u32>, offset: usize, maximum: usize) -> Result<()> {
    if offsets.len() >= maximum {
        return Err(limit(
            "DOCX ink host candidate anchors",
            offsets.len().saturating_add(1),
            maximum,
        ));
    }
    offsets.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX ink host candidate anchors",
        source,
    })?;
    offsets
        .push(u32::try_from(offset).map_err(|error| {
            Error::Invalid(format!("DOCX story position exceeds u32: {error}"))
        })?);
    Ok(())
}

fn observe_node(nodes: &mut usize, maximum: usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("XML nodes", usize::MAX, maximum))?;
    if *nodes > maximum {
        return Err(limit("XML nodes", *nodes, maximum));
    }
    Ok(())
}

fn reader(xml: &[u8]) -> NsReader<&[u8]> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    reader
}

fn is_candidate(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    dialect: StoryDialect,
) -> bool {
    element.local_name().as_ref() == b"contentPart"
        && (is_namespace(namespace, dialect_word_namespace(dialect))
            || is_namespace(namespace, WORD_2010_WORDML))
}

fn is_drawing_namespace(namespace: &ResolveResult<'_>, dialect: StoryDialect) -> bool {
    is_namespace(
        namespace,
        match dialect {
            StoryDialect::Transitional => TRANSITIONAL_DRAWINGML,
            StoryDialect::Strict => STRICT_DRAWINGML,
        },
    )
}

fn is_wordprocessing_drawing_namespace(
    namespace: &ResolveResult<'_>,
    dialect: StoryDialect,
) -> bool {
    is_namespace(
        namespace,
        match dialect {
            StoryDialect::Transitional => TRANSITIONAL_WORDPROCESSING_DRAWING,
            StoryDialect::Strict => STRICT_WORDPROCESSING_DRAWING,
        },
    )
}

fn dialect_word_namespace(dialect: StoryDialect) -> &'static [u8] {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORD,
        StoryDialect::Strict => STRICT_WORD,
    }
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn opening_span(
    event_position: usize,
    element: &BytesStart<'_>,
    empty: bool,
    xml: &[u8],
) -> Result<(usize, usize)> {
    let suffix = if empty { 3 } else { 2 };
    let consumed = element
        .as_ref()
        .len()
        .checked_add(suffix)
        .ok_or_else(|| invalid("DOCX story element position overflowed"))?;
    let start = event_position
        .checked_sub(consumed)
        .ok_or_else(|| invalid("DOCX story element position underflowed"))?;
    if xml.get(start).copied() != Some(b'<') {
        return Err(invalid("DOCX story element position is not an opening tag"));
    }
    Ok((start, event_position))
}

fn end_position(
    event_position: usize,
    element: &quick_xml::events::BytesEnd<'_>,
    xml: &[u8],
) -> Result<usize> {
    let consumed = element
        .as_ref()
        .len()
        .checked_add(3)
        .ok_or_else(|| invalid("DOCX story end position overflowed"))?;
    let start = event_position
        .checked_sub(consumed)
        .ok_or_else(|| invalid("DOCX story end position underflowed"))?;
    if xml.get(start..start.saturating_add(2)) != Some(b"</".as_slice())
        || event_position > xml.len()
    {
        return Err(invalid("DOCX story end position is not a closing tag"));
    }
    Ok(event_position)
}

fn position(position: u64) -> Result<usize> {
    usize::try_from(position)
        .map_err(|error| Error::Invalid(format!("DOCX story position does not fit usize: {error}")))
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

#[cfg(test)]
#[path = "host_tests.rs"]
mod tests;
