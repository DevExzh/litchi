//! Shape-owned PresentationML programmable-tag attachments.
//!
//! A shape tag list is anchored below the selected shape's application
//! non-visual properties. The relationship itself remains owned by the
//! containing slide, layout, master, notes, or handout part:
//!
//! ```text
//! p:{sp,pic,cxnSp,graphicFrame,grpSp}
//!   / p:nv*Pr / p:nvPr / p:custDataLst / p:tags@r:id
//! ```
//!
//! Name lookup is the ordinary entry point. A checked depth-first position is
//! retained by [`crate::shape::Key`] for duplicate-name repair and source-order
//! workflows. Relationship IDs and tag-part names never participate in the
//! safe selector.

use super::{
    CONTENT_TYPE, Conformance, List, Source, allocation, available_part_name,
    available_relationship_id, has_other_inbound, invalid, pml, relationship_namespace,
    replace_xml, staged_xml, validate_relative_target, validate_selected_relationship,
};
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, OffsetLimits, active_offsets};
use litchi_opc::{OpcPackage, PackURI, Part as OpcPart, XmlPart};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;

const MAX_OWNER_BYTES: usize = 64 * 1024 * 1024;
const MAX_OWNER_NODES: usize = 1_000_000;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4_096;
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";

/// Load the optional programmable-tag list attached to one semantic shape.
///
/// `owner` is the containing slide-like PresentationML part. The selected
/// shape may be a regular shape, picture, connector, graphic frame, group, or
/// a child at any bounded nested-group depth. The shape-tree root itself is not
/// a user shape and is never selected.
pub fn load<'k>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<Option<Source>> {
    let owner = package.get_part(owner)?;
    let layout = selected_layout(owner, key.into())?;
    layout
        .anchor
        .as_ref()
        .map(|anchor| resolve(owner, package, anchor, layout.conformance))
        .transpose()
}

/// Create or replace the optional programmable-tag list on one semantic shape.
///
/// The owned list moves into a fully staged tag part. The shape XML anchor,
/// containing-part relationship, and target part commit together only after
/// bounded re-validation. A byte-identical replacement is a true no-op. When
/// another preserved raw anchor uses the same relationship ID, or any other
/// package edge reaches the target part, replacement forks the selected
/// attachment.
pub fn put<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    list: List,
) -> Result<Option<List>> {
    let key = key.into();
    let (owner_name, layout, attached, anchor_uses) = {
        let owner = package.get_part(owner)?;
        let layout = selected_layout(owner, key)?;
        let attached = layout
            .anchor
            .as_ref()
            .map(|anchor| {
                let source = resolve(owner, package, anchor, layout.conformance)?;
                let relationship = owner.rels().get(source.rel()).ok_or_else(|| {
                    invalid("shape tag relationship disappeared during preflight")
                })?;
                Ok::<_, Error>(Attached {
                    relationship_type: relationship.reltype().into(),
                    source,
                })
            })
            .transpose()?;
        let anchor_uses = layout
            .anchor
            .as_ref()
            .map(|anchor| preserved_anchor_uses(owner.blob(), &anchor.id))
            .transpose()?
            .unwrap_or(0);
        (owner.partname().clone(), layout, attached, anchor_uses)
    };

    let conformance = attached
        .as_ref()
        .map_or(layout.conformance, |value| value.source.conformance());
    let tag_xml = staged_xml(&list, conformance)?;

    if let Some(attached) = attached {
        if package.get_part(attached.source.part())?.blob() == tag_xml {
            return Ok(Some(attached.source.into_list()));
        }

        let shared_target = has_other_inbound(
            package,
            attached.source.part(),
            &owner_name,
            attached.source.rel(),
        )?;
        let shared_anchor = anchor_uses > 1;
        let fork = shared_target || shared_anchor;
        let part_name = if fork {
            available_part_name(package)?
        } else {
            attached.source.part().clone()
        };
        let target_ref = part_name.relative_ref(owner_name.base_uri());
        validate_relative_target(&owner_name, &target_ref, &part_name)?;
        if fork {
            package.validate_new_part_name(&part_name)?;
        }

        let (relationship_id, owner_xml) = if shared_anchor {
            let owner = package.get_part(&owner_name)?;
            let relationship_id = available_relationship_id(owner)?;
            let owner_xml = replace_anchor_id(owner.blob(), &layout, &relationship_id)?;
            validate_staged_anchor(owner, &owner_xml, key, Some(&relationship_id))?;
            (relationship_id, Some(owner_xml))
        } else {
            (attached.source.rel().to_owned(), None)
        };

        {
            let owner = package.get_part_mut(&owner_name)?;
            validate_selected_relationship(
                owner,
                attached.source.rel(),
                &attached.relationship_type,
                attached.source.part(),
            )?;
            if let Some(owner_xml) = owner_xml {
                owner.set_blob(owner_xml);
                owner.rels_mut().add_relationship(
                    attached.relationship_type,
                    target_ref,
                    relationship_id,
                    false,
                );
            } else if shared_target {
                let _ = owner.rels_mut().remove(attached.source.rel());
                owner.rels_mut().add_relationship(
                    attached.relationship_type,
                    target_ref,
                    relationship_id,
                    false,
                );
            }
        }
        package.add_part(Box::new(XmlPart::new(
            part_name,
            CONTENT_TYPE.into(),
            tag_xml,
        )));
        package.unsign();
        return Ok(Some(attached.source.into_list()));
    }

    let (relationship_id, part_name, target_ref, owner_xml) = {
        let owner = package.get_part(&owner_name)?;
        let relationship_id = available_relationship_id(owner)?;
        let part_name = available_part_name(package)?;
        let target_ref = part_name.relative_ref(owner_name.base_uri());
        validate_relative_target(&owner_name, &target_ref, &part_name)?;
        let owner_xml = add_anchor(owner.blob(), &layout, &relationship_id)?;
        validate_staged_anchor(owner, &owner_xml, key, Some(&relationship_id))?;
        (relationship_id, part_name, target_ref, owner_xml)
    };
    package.validate_new_part_name(&part_name)?;

    {
        let owner = package.get_part_mut(&owner_name)?;
        let current = selected_layout(owner, key)?;
        if current.anchor.is_some() || owner.rels().get(&relationship_id).is_some() {
            return Err(invalid("shape tag graph changed during preflight"));
        }
        owner.set_blob(owner_xml);
        owner.rels_mut().add_relationship(
            layout.conformance.relationship().into(),
            target_ref,
            relationship_id,
            false,
        );
    }
    package.add_part(Box::new(XmlPart::new(
        part_name,
        CONTENT_TYPE.into(),
        tag_xml,
    )));
    package.unsign();
    Ok(None)
}

/// Remove the optional programmable-tag list from one semantic shape.
///
/// An absent attachment is an idempotent, signature-preserving `Ok(None)`.
/// Customer data, inactive MCE branches, comments, and extension bytes remain
/// untouched. The relationship remains while another preserved raw anchor
/// uses its ID, and the target part is collected only after the package graph
/// proves it orphaned.
pub fn remove<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<Option<List>> {
    let key = key.into();
    let (owner_name, layout, attached, owner_xml, retain_relationship, orphan) = {
        let owner = package.get_part(owner)?;
        let layout = selected_layout(owner, key)?;
        let Some(anchor) = layout.anchor.as_ref() else {
            return Ok(None);
        };
        let source = resolve(owner, package, anchor, layout.conformance)?;
        let relationship = owner
            .rels()
            .get(source.rel())
            .ok_or_else(|| invalid("shape tag relationship disappeared during preflight"))?;
        let relationship_type = relationship.reltype().to_owned();
        let owner_name = owner.partname().clone();
        let retain_relationship = preserved_anchor_uses(owner.blob(), &anchor.id)? > 1;
        let owner_xml = remove_anchor(owner.blob(), &layout)?;
        validate_staged_anchor(owner, &owner_xml, key, None)?;
        let orphan = !retain_relationship
            && !has_other_inbound(package, source.part(), &owner_name, source.rel())?;
        (
            owner_name,
            layout,
            Attached {
                relationship_type,
                source,
            },
            owner_xml,
            retain_relationship,
            orphan,
        )
    };

    {
        let owner = package.get_part_mut(&owner_name)?;
        let current = selected_layout(owner, key)?;
        if current.anchor.as_ref().map(|anchor| anchor.id.as_str())
            != layout.anchor.as_ref().map(|anchor| anchor.id.as_str())
        {
            return Err(invalid("shape tag anchor changed during preflight"));
        }
        validate_selected_relationship(
            owner,
            attached.source.rel(),
            &attached.relationship_type,
            attached.source.part(),
        )?;
        owner.set_blob(owner_xml);
        if !retain_relationship {
            let _ = owner.rels_mut().remove(attached.source.rel());
        }
    }
    if orphan {
        let _ = package.remove_part(attached.source.part());
    }
    package.unsign();
    Ok(Some(attached.source.into_list()))
}

struct Attached {
    relationship_type: String,
    source: Source,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Family {
    Shape,
    Picture,
    Connector,
    GraphicFrame,
    Group,
}

impl Family {
    fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"sp" => Some(Self::Shape),
            b"pic" => Some(Self::Picture),
            b"cxnSp" => Some(Self::Connector),
            b"graphicFrame" => Some(Self::GraphicFrame),
            b"grpSp" => Some(Self::Group),
            _ => None,
        }
    }

    fn non_visual(self) -> &'static [u8] {
        match self {
            Self::Shape => b"nvSpPr",
            Self::Picture => b"nvPicPr",
            Self::Connector => b"nvCxnSpPr",
            Self::GraphicFrame => b"nvGraphicFramePr",
            Self::Group => b"nvGrpSpPr",
        }
    }

    fn permits_placeholder(self) -> bool {
        !matches!(self, Self::Connector | Self::Group)
    }
}

#[derive(Debug, Clone)]
struct Element {
    span: Range<usize>,
    open_end: usize,
    close_start: usize,
    empty: bool,
}

#[derive(Debug, Clone)]
struct Container {
    element: Element,
    child_elements: usize,
    preserve_when_empty: bool,
}

#[derive(Debug, Clone)]
struct Anchor {
    id: String,
    span: Range<usize>,
    id_value: Range<usize>,
}

#[derive(Debug)]
struct Layout {
    conformance: Conformance,
    nv_pr: Element,
    insertion: usize,
    container: Option<Container>,
    anchor: Option<Anchor>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NvPhase {
    Start,
    Placeholder,
    Media,
    CustomerData,
    Extensions,
}

#[derive(Debug)]
enum NodeKind {
    Root(Family),
    NonVisual,
    NvPr {
        phase: NvPhase,
        ext_start: Option<usize>,
    },
    Container {
        children: usize,
        tags_seen: bool,
        preserve_when_empty: bool,
    },
    Anchor {
        id: String,
        id_value: Range<usize>,
    },
    Opaque,
}

#[derive(Debug)]
struct Node {
    kind: NodeKind,
    start: usize,
    open_end: usize,
}

#[derive(Debug, Clone, Copy)]
struct RawFrame {
    semantic: bool,
}

fn selected_layout<'k>(owner: &dyn OpcPart, key: crate::shape::Key<'k>) -> Result<Layout> {
    validate_owner_content_type(owner.content_type())?;
    let span = selected_raw_span(owner.blob(), key)?;
    scan_layout(owner.blob(), span)
}

fn selected_raw_span<'k>(xml: &[u8], key: crate::shape::Key<'k>) -> Result<Range<usize>> {
    let scene = crate::shape::read(xml)?;
    let shape = scene.shape(key)?;
    let family = selected_family(shape)?;
    raw_shape_span(xml, shape.common().index(), scene.len(), family)
}

fn selected_family(shape: crate::shape::Shape<'_>) -> Result<Family> {
    match shape {
        crate::shape::Shape::Auto(_) => Ok(Family::Shape),
        crate::shape::Shape::Picture(_) => Ok(Family::Picture),
        crate::shape::Shape::Connector(_) => Ok(Family::Connector),
        crate::shape::Shape::Group(_) => Ok(Family::Group),
        crate::shape::Shape::Table(_)
        | crate::shape::Shape::Chart(_)
        | crate::shape::Shape::Diagram(_)
        | crate::shape::Shape::Ole(_)
        | crate::shape::Shape::Frame(_) => Ok(Family::GraphicFrame),
        crate::shape::Shape::Content(_) | crate::shape::Shape::Unknown(_) => Err(invalid(
            "programmable tags are not defined for this extension shape",
        )),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RawCandidateKind {
    Supported(Family),
    Unsupported,
}

impl RawCandidateKind {
    fn is_group(self) -> bool {
        matches!(self, Self::Supported(Family::Group))
    }

    fn family(self) -> Option<Family> {
        match self {
            Self::Supported(family) => Some(family),
            Self::Unsupported => None,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct RawMapFrame {
    active_tree: bool,
    active_candidate: Option<RawCandidateKind>,
    selected_start: Option<usize>,
}

fn raw_shape_span(
    xml: &[u8],
    selected_index: usize,
    scene_len: usize,
    expected_family: Family,
) -> Result<Range<usize>> {
    let offsets = raw_shape_offsets(xml)?;
    let active = active_offsets(
        xml,
        &offsets,
        &shape_mce_capabilities(),
        &OffsetLimits::default(),
    )?;
    let mut active = active.into_iter().peekable();
    let mut reader = NsReader::from_reader(xml);
    let mut frames = Vec::new();
    let mut nodes = 0usize;
    let mut mapped = 0usize;
    let mut active_trees = 0usize;
    let mut selected = None;

    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let classification = match &event {
            Event::Start(element) | Event::Empty(element) => Some((
                is_shape_tree(&namespace, element),
                raw_candidate(&namespace, element),
            )),
            _ => None,
        };
        drop(namespace);
        let end = xml_position(&reader)?;
        match event {
            Event::Start(_) => {
                bump_nodes(&mut nodes)?;
                let (shape_tree, candidate) = classification
                    .ok_or_else(|| invalid("raw shape start lost its classification"))?;
                let active_event =
                    (shape_tree || candidate.is_some()) && take_active(&mut active, start)?;
                let active_tree = shape_tree && active_event;
                let active_candidate = candidate.filter(|_| active_event);

                if active_tree {
                    active_trees = active_trees.checked_add(1).ok_or(Error::Limit {
                        resource: "active PresentationML shape trees",
                        limit: 1,
                    })?;
                    if active_trees > 1 {
                        return Err(invalid(
                            "shape owner contains more than one active shape tree",
                        ));
                    }
                }

                let selected_start = if active_candidate.is_some() {
                    let in_tree = frames.iter().any(|frame: &RawMapFrame| frame.active_tree);
                    let parent = frames.iter().rev().find_map(|frame| frame.active_candidate);
                    if in_tree && parent.is_none_or(RawCandidateKind::is_group) {
                        let index = mapped;
                        mapped = mapped.checked_add(1).ok_or(Error::Limit {
                            resource: "mapped active shapes",
                            limit: MAX_OWNER_NODES,
                        })?;
                        (index == selected_index).then_some(start)
                    } else {
                        None
                    }
                } else {
                    None
                };
                try_push(
                    &mut frames,
                    RawMapFrame {
                        active_tree,
                        active_candidate,
                        selected_start,
                    },
                    "raw shape-map stack",
                )?;
            },
            Event::Empty(_) => {
                bump_nodes(&mut nodes)?;
                let (shape_tree, candidate) = classification
                    .ok_or_else(|| invalid("raw empty shape lost its classification"))?;
                let active_event =
                    (shape_tree || candidate.is_some()) && take_active(&mut active, start)?;
                let active_tree = shape_tree && active_event;
                let active_candidate = candidate.filter(|_| active_event);
                if active_tree {
                    active_trees = active_trees.checked_add(1).ok_or(Error::Limit {
                        resource: "active PresentationML shape trees",
                        limit: 1,
                    })?;
                    if active_trees > 1 {
                        return Err(invalid(
                            "shape owner contains more than one active shape tree",
                        ));
                    }
                }
                if let Some(kind) = active_candidate {
                    let in_tree = frames.iter().any(|frame: &RawMapFrame| frame.active_tree);
                    let parent = frames.iter().rev().find_map(|frame| frame.active_candidate);
                    if in_tree && parent.is_none_or(RawCandidateKind::is_group) {
                        let index = mapped;
                        mapped = mapped.checked_add(1).ok_or(Error::Limit {
                            resource: "mapped active shapes",
                            limit: MAX_OWNER_NODES,
                        })?;
                        if index == selected_index {
                            selected = Some((start..end, kind));
                        }
                    }
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("raw shape-map stack underflow"))?;
                if let Some(start) = frame.selected_start {
                    selected = Some((
                        start..end,
                        frame
                            .active_candidate
                            .ok_or_else(|| invalid("selected raw shape lost its candidate kind"))?,
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("shape owner contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !frames.is_empty() {
        return Err(invalid("raw shape-map XML is unterminated"));
    }
    if active.peek().is_some() {
        return Err(invalid("raw shape-map active offsets were not consumed"));
    }
    if mapped != scene_len {
        return Err(invalid(
            "raw and MCE-processed shape indexes do not have the same length",
        ));
    }
    let (span, kind) = selected.ok_or_else(|| invalid("selected raw shape was not mapped"))?;
    if kind.family() != Some(expected_family) {
        return Err(invalid(
            "raw and MCE-processed shape kinds do not describe the same shape",
        ));
    }
    Ok(span)
}

fn raw_shape_offsets(xml: &[u8]) -> Result<Vec<u32>> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "shape-tag owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let mut offsets = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                bump_nodes(&mut nodes)?;
                if is_shape_tree(&namespace, &element)
                    || raw_candidate(&namespace, &element).is_some()
                {
                    try_push(
                        &mut offsets,
                        offset_u32(start)?,
                        "raw shape-map MCE offsets",
                    )?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(offsets)
}

fn is_shape_tree(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    element.local_name().as_ref() == b"spTree" && pml(namespace).is_some()
}

fn raw_candidate(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Option<RawCandidateKind> {
    let local = element.local_name();
    if let Some(family) = Family::from_local(local.as_ref()) {
        return Some(if pml(namespace).is_some() {
            RawCandidateKind::Supported(family)
        } else {
            RawCandidateKind::Unsupported
        });
    }
    (local.as_ref() == b"contentPart").then_some(RawCandidateKind::Unsupported)
}

fn validate_owner_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
    ) {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "PresentationML shape-tag owner".into(),
            actual: content_type.into(),
        })
    }
}

fn scan_layout(xml: &[u8], shape_span: Range<usize>) -> Result<Layout> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "shape-tag owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    if shape_span.start >= shape_span.end || shape_span.end > xml.len() {
        return Err(invalid("selected shape span is outside its owner XML"));
    }

    let active = active_pml_offsets(xml, &shape_span)?;
    let mut active = active.into_iter().peekable();
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut raw = Vec::new();
    let mut semantic = Vec::new();
    let mut nodes = 0usize;
    let mut conformance = None;
    let mut root_family = None;
    let mut non_visual_seen = false;
    let mut nv_pr = None;
    let mut insertion = None;
    let mut container = None;
    let mut anchor = None;

    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        let end = xml_position(&reader)?;
        if start >= shape_span.end && raw.is_empty() && conformance.is_some() {
            break;
        }

        match event {
            Event::Start(element) => {
                bump_nodes(&mut nodes)?;
                let in_shape = start >= shape_span.start && start < shape_span.end;
                let is_active = in_shape && profile.is_some() && take_active(&mut active, start)?;
                let semantic_node = if is_active {
                    let profile =
                        profile.ok_or_else(|| invalid("active shape element lost profile"))?;
                    if let Some(expected) = conformance
                        && profile != expected
                    {
                        return Err(invalid(
                            "shape tag owner mixes Strict and Transitional PresentationML",
                        ));
                    }
                    let kind = start_node(
                        &reader,
                        xml,
                        &element,
                        start,
                        end,
                        profile,
                        &mut semantic,
                        &mut conformance,
                        &mut root_family,
                        &mut non_visual_seen,
                    )?;
                    try_push(
                        &mut semantic,
                        Node {
                            kind,
                            start,
                            open_end: end,
                        },
                        "shape semantic stack",
                    )?;
                    true
                } else {
                    false
                };
                if in_shape || !raw.is_empty() {
                    try_push(
                        &mut raw,
                        RawFrame {
                            semantic: semantic_node,
                        },
                        "shape raw stack",
                    )?;
                }
            },
            Event::Empty(element) => {
                bump_nodes(&mut nodes)?;
                if start < shape_span.start || start >= shape_span.end || profile.is_none() {
                    continue;
                }
                if !take_active(&mut active, start)? {
                    continue;
                }
                let profile =
                    profile.ok_or_else(|| invalid("active shape element lost profile"))?;
                if let Some(expected) = conformance
                    && profile != expected
                {
                    return Err(invalid(
                        "shape tag owner mixes Strict and Transitional PresentationML",
                    ));
                }
                let kind = start_node(
                    &reader,
                    xml,
                    &element,
                    start,
                    end,
                    profile,
                    &mut semantic,
                    &mut conformance,
                    &mut root_family,
                    &mut non_visual_seen,
                )?;
                finish_node(
                    xml,
                    Node {
                        kind,
                        start,
                        open_end: end,
                    },
                    start,
                    end,
                    true,
                    &mut nv_pr,
                    &mut insertion,
                    &mut container,
                    &mut anchor,
                )?;
            },
            Event::End(_) => {
                if raw.is_empty() {
                    continue;
                }
                let frame = raw
                    .pop()
                    .ok_or_else(|| invalid("shape raw stack underflow"))?;
                if frame.semantic {
                    let node = semantic
                        .pop()
                        .ok_or_else(|| invalid("shape semantic stack underflow"))?;
                    finish_node(
                        xml,
                        node,
                        start,
                        end,
                        false,
                        &mut nv_pr,
                        &mut insertion,
                        &mut container,
                        &mut anchor,
                    )?;
                }
                if raw.is_empty() && start < shape_span.end {
                    if end != shape_span.end {
                        return Err(invalid("selected shape span does not end on its root"));
                    }
                    break;
                }
            },
            Event::Text(text) => {
                if semantic.last().is_some_and(structural_text_parent)
                    && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid("shape structural metadata cannot contain text"));
                }
            },
            Event::CData(text) => {
                if semantic.last().is_some_and(structural_text_parent)
                    && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid("shape structural metadata cannot contain CDATA"));
                }
            },
            Event::GeneralRef(_) if semantic.last().is_some_and(structural_text_parent) => {
                return Err(invalid(
                    "shape structural metadata cannot contain entity text",
                ));
            },
            Event::GeneralRef(_) => {},
            Event::Decl(_) if raw.is_empty() => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) if !raw.is_empty() => {
                return Err(invalid("shape tag owner contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !raw.is_empty() || !semantic.is_empty() {
        return Err(invalid("selected shape XML is unterminated"));
    }
    let conformance = conformance.ok_or_else(|| invalid("selected shape profile is missing"))?;
    let _ = root_family.ok_or_else(|| invalid("selected shape root is missing"))?;
    if !non_visual_seen {
        return Err(invalid(
            "selected shape has no application non-visual container",
        ));
    }
    let nv_pr = nv_pr.ok_or_else(|| invalid("selected shape has no direct p:nvPr"))?;
    let insertion = insertion.unwrap_or(nv_pr.close_start);
    if let (Some(container), Some(anchor)) = (&container, &anchor)
        && (!container.element.span.contains(&anchor.span.start)
            || anchor.span.end > container.element.span.end)
    {
        return Err(invalid("shape p:tags anchor is outside p:custDataLst"));
    }
    Ok(Layout {
        conformance,
        nv_pr,
        insertion,
        container,
        anchor,
    })
}

fn structural_text_parent(node: &Node) -> bool {
    matches!(
        node.kind,
        NodeKind::Root(_)
            | NodeKind::NonVisual
            | NodeKind::NvPr { .. }
            | NodeKind::Container { .. }
            | NodeKind::Anchor { .. }
    )
}

#[allow(clippy::too_many_arguments)]
fn start_node(
    reader: &NsReader<&[u8]>,
    xml: &[u8],
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    profile: Conformance,
    semantic: &mut [Node],
    conformance: &mut Option<Conformance>,
    root_family: &mut Option<Family>,
    non_visual_seen: &mut bool,
) -> Result<NodeKind> {
    let local = element.local_name();
    if semantic.is_empty() {
        if start == 0 || start > xml.len() {
            return Err(invalid("selected shape root offset is invalid"));
        }
        let family = Family::from_local(local.as_ref())
            .ok_or_else(|| invalid("selected object is not a supported PresentationML shape"))?;
        if root_family.replace(family).is_some() {
            return Err(invalid("selected shape has multiple roots"));
        }
        *conformance = Some(profile);
        return Ok(NodeKind::Root(family));
    }

    let parent = semantic
        .last_mut()
        .ok_or_else(|| invalid("shape semantic parent is missing"))?;
    match &mut parent.kind {
        NodeKind::Root(family) if local.as_ref() == family.non_visual() => {
            if *non_visual_seen {
                return Err(invalid("selected shape has multiple non-visual containers"));
            }
            *non_visual_seen = true;
            Ok(NodeKind::NonVisual)
        },
        NodeKind::NonVisual if local.as_ref() == b"nvPr" => Ok(NodeKind::NvPr {
            phase: NvPhase::Start,
            ext_start: None,
        }),
        NodeKind::NvPr { phase, ext_start } => {
            observe_nv_child(
                local.as_ref(),
                start,
                *root_family
                    .as_ref()
                    .ok_or_else(|| invalid("selected shape family is missing"))?,
                phase,
                ext_start,
            )?;
            if local.as_ref() == b"custDataLst" {
                Ok(NodeKind::Container {
                    children: 0,
                    tags_seen: false,
                    preserve_when_empty: has_non_namespace_attrs(element)?,
                })
            } else {
                Ok(NodeKind::Opaque)
            }
        },
        NodeKind::Container {
            children,
            tags_seen,
            ..
        } => {
            observe_customer_child(local.as_ref(), tags_seen)?;
            *children = children.checked_add(1).ok_or(Error::Limit {
                resource: "shape customer-data children",
                limit: MAX_OWNER_NODES,
            })?;
            if local.as_ref() == b"tags" {
                let (id, id_value) = anchor_id(reader, xml, element, start..end, profile)?;
                Ok(NodeKind::Anchor { id, id_value })
            } else {
                Ok(NodeKind::Opaque)
            }
        },
        NodeKind::Anchor { .. } => Err(invalid("shape p:tags cannot contain child elements")),
        _ => Ok(NodeKind::Opaque),
    }
}

#[allow(clippy::too_many_arguments)]
fn finish_node(
    xml: &[u8],
    node: Node,
    close_start: usize,
    end: usize,
    empty: bool,
    nv_pr: &mut Option<Element>,
    insertion: &mut Option<usize>,
    container: &mut Option<Container>,
    anchor: &mut Option<Anchor>,
) -> Result<()> {
    match node.kind {
        NodeKind::NvPr { ext_start, .. } => {
            if nv_pr.is_some() {
                return Err(invalid(
                    "selected shape has multiple direct p:nvPr elements",
                ));
            }
            let element = Element {
                span: node.start..end,
                open_end: node.open_end,
                close_start: if empty { end } else { close_start },
                empty,
            };
            *insertion = Some(ext_start.unwrap_or(element.close_start));
            *nv_pr = Some(element);
        },
        NodeKind::Container {
            children,
            preserve_when_empty,
            ..
        } => {
            if container.is_some() {
                return Err(invalid(
                    "selected shape has multiple p:custDataLst elements",
                ));
            }
            *container = Some(Container {
                element: Element {
                    span: node.start..end,
                    open_end: node.open_end,
                    close_start: if empty { end } else { close_start },
                    empty,
                },
                child_elements: children,
                preserve_when_empty,
            });
        },
        NodeKind::Anchor { id, id_value } => {
            if anchor.is_some() {
                return Err(invalid("selected shape has multiple p:tags anchors"));
            }
            *anchor = Some(Anchor {
                id,
                span: node.start..end,
                id_value,
            });
        },
        _ => {},
    }
    if end > xml.len() {
        return Err(invalid("shape element span exceeds owner XML"));
    }
    Ok(())
}

fn observe_nv_child(
    local: &[u8],
    start: usize,
    family: Family,
    phase: &mut NvPhase,
    ext_start: &mut Option<usize>,
) -> Result<()> {
    *phase = match (local, *phase) {
        (b"ph", NvPhase::Start) if family.permits_placeholder() => NvPhase::Placeholder,
        (b"ph", _) if !family.permits_placeholder() => {
            return Err(invalid(
                "PowerPoint forbids p:ph on connector and group non-visual properties",
            ));
        },
        (
            b"audioCd" | b"wavAudioFile" | b"audioFile" | b"videoFile" | b"quickTimeFile",
            NvPhase::Start | NvPhase::Placeholder,
        ) => NvPhase::Media,
        (b"custDataLst", NvPhase::Start | NvPhase::Placeholder | NvPhase::Media) => {
            NvPhase::CustomerData
        },
        (
            b"extLst",
            NvPhase::Start | NvPhase::Placeholder | NvPhase::Media | NvPhase::CustomerData,
        ) => {
            *ext_start = Some(start);
            NvPhase::Extensions
        },
        (b"ph", _) => return Err(invalid("p:ph is duplicated or out of schema order")),
        (b"audioCd" | b"wavAudioFile" | b"audioFile" | b"videoFile" | b"quickTimeFile", _) => {
            return Err(invalid("shape media is duplicated or out of schema order"));
        },
        (b"custDataLst", _) => {
            return Err(invalid(
                "shape p:custDataLst is duplicated or out of schema order",
            ));
        },
        (b"extLst", _) => return Err(invalid("shape p:extLst is duplicated or out of order")),
        _ => return Err(invalid("unsupported direct p:nvPr child")),
    };
    Ok(())
}

fn observe_customer_child(local: &[u8], tags_seen: &mut bool) -> Result<()> {
    if *tags_seen {
        return Err(invalid("shape p:tags must be the last p:custDataLst child"));
    }
    match local {
        b"custData" => Ok(()),
        b"tags" => {
            *tags_seen = true;
            Ok(())
        },
        _ => Err(invalid("unsupported direct shape p:custDataLst child")),
    }
}

fn resolve(
    owner: &dyn OpcPart,
    package: &OpcPackage,
    anchor: &Anchor,
    conformance: Conformance,
) -> Result<Source> {
    super::resolve_anchor(
        owner,
        package,
        &super::Anchor {
            id: anchor.id.clone(),
            span: anchor.span.clone(),
            id_value: anchor.id_value.clone(),
        },
        conformance,
    )
}

fn add_anchor(xml: &[u8], layout: &Layout, relationship_id: &str) -> Result<Vec<u8>> {
    if layout.anchor.is_some() {
        return Err(invalid("selected shape already has a p:tags anchor"));
    }
    let anchor = format!(
        "<p:tags xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"/>",
        layout.conformance.namespace(),
        layout.conformance.relationship_namespace(),
        relationship_id
    );

    if let Some(container) = &layout.container {
        if !container.element.empty {
            return replace_xml(
                xml,
                container.element.close_start..container.element.close_start,
                anchor.as_bytes(),
            );
        }
        return expand_empty(xml, &container.element, anchor.as_bytes());
    }

    let container = format!(
        "<p:custDataLst xmlns:p=\"{}\">{anchor}</p:custDataLst>",
        layout.conformance.namespace()
    );
    if layout.nv_pr.empty {
        expand_empty(xml, &layout.nv_pr, container.as_bytes())
    } else {
        replace_xml(
            xml,
            layout.insertion..layout.insertion,
            container.as_bytes(),
        )
    }
}

fn replace_anchor_id(xml: &[u8], layout: &Layout, relationship_id: &str) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("selected shape has no p:tags anchor"))?;
    replace_xml(xml, anchor.id_value.clone(), relationship_id.as_bytes())
}

fn remove_anchor(xml: &[u8], layout: &Layout) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("selected shape has no p:tags anchor"))?;
    let container = layout
        .container
        .as_ref()
        .ok_or_else(|| invalid("shape p:tags has no p:custDataLst parent"))?;
    if container.child_elements == 1
        && !container.preserve_when_empty
        && container_contains_only_anchor(xml, container, anchor)?
    {
        replace_xml(xml, container.element.span.clone(), &[])
    } else {
        replace_xml(xml, anchor.span.clone(), &[])
    }
}

fn expand_empty(xml: &[u8], element: &Element, child: &[u8]) -> Result<Vec<u8>> {
    let raw = xml
        .get(element.span.clone())
        .ok_or_else(|| invalid("empty shape element span is outside owner XML"))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("empty shape element has no closing slash"))?;
    let qualified_name = element_qname(raw)?;
    let len = raw
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(child.len()))
        .and_then(|value| value.checked_add(qualified_name.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or_else(|| invalid("expanded shape element length overflow"))?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(len)
        .map_err(|source| allocation("expanded shape-tag element", source))?;
    replacement.extend_from_slice(&raw[..slash]);
    replacement.extend_from_slice(&raw[slash + 1..]);
    replacement.extend_from_slice(child);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(qualified_name);
    replacement.push(b'>');
    replace_xml(xml, element.span.clone(), &replacement)
}

fn element_qname(element: &[u8]) -> Result<&[u8]> {
    if element.first() != Some(&b'<') {
        return Err(invalid("shape element does not start with '<'"));
    }
    let start = 1usize;
    let end = element[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .map(|offset| start + offset)
        .ok_or_else(|| invalid("shape element name is unterminated"))?;
    element
        .get(start..end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("shape element name is empty"))
}

fn container_contains_only_anchor(
    xml: &[u8],
    container: &Container,
    anchor: &Anchor,
) -> Result<bool> {
    if container.element.empty {
        return Ok(false);
    }
    let before = xml
        .get(container.element.open_end..anchor.span.start)
        .ok_or_else(|| invalid("shape customer-data prefix range is invalid"))?;
    let after = xml
        .get(anchor.span.end..container.element.close_start)
        .ok_or_else(|| invalid("shape customer-data suffix range is invalid"))?;
    Ok(before.iter().all(u8::is_ascii_whitespace) && after.iter().all(u8::is_ascii_whitespace))
}

fn validate_staged_anchor<'k>(
    owner: &dyn OpcPart,
    xml: &[u8],
    key: crate::shape::Key<'k>,
    expected_id: Option<&str>,
) -> Result<()> {
    validate_owner_content_type(owner.content_type())?;
    let staged = scan_layout(xml, selected_raw_span(xml, key)?)?;
    if staged.anchor.as_ref().map(|anchor| anchor.id.as_str()) != expected_id {
        return Err(invalid("staged shape tag anchor did not round-trip"));
    }
    Ok(())
}

fn active_pml_offsets(xml: &[u8], span: &Range<usize>) -> Result<Vec<u32>> {
    let mut reader = NsReader::from_reader(xml);
    let mut offsets = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(_) | Event::Empty(_)
                if start >= span.start && start < span.end && profile.is_some() =>
            {
                bump_nodes(&mut nodes)?;
                try_push(
                    &mut offsets,
                    offset_u32(start)?,
                    "shape MCE candidate offsets",
                )?;
            },
            Event::Eof => break,
            _ => {},
        }
        if start >= span.end {
            break;
        }
    }
    active_offsets(
        xml,
        &offsets,
        &shape_mce_capabilities(),
        &OffsetLimits::default(),
    )
    .map_err(Into::into)
}

fn preserved_anchor_uses(xml: &[u8], relationship_id: &str) -> Result<usize> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "shape-tag owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let mut nodes = 0usize;
    let mut uses = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(element) | Event::Empty(element)
                if profile.is_some() && element.local_name().as_ref() == b"tags" =>
            {
                bump_nodes(&mut nodes)?;
                let profile =
                    profile.ok_or_else(|| invalid("preserved p:tags profile is missing"))?;
                let (id, _) = anchor_id(
                    &reader,
                    xml,
                    &element,
                    start..xml_position(&reader)?,
                    profile,
                )?;
                if id == relationship_id {
                    uses = uses.checked_add(1).ok_or(Error::Limit {
                        resource: "preserved shape tag-anchor references",
                        limit: MAX_OWNER_NODES,
                    })?;
                }
            },
            Event::Start(_) | Event::Empty(_) => bump_nodes(&mut nodes)?,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(uses)
}

fn shape_mce_capabilities() -> Capabilities {
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14);
    capabilities.understand_namespace(P15);
    capabilities
}

fn anchor_id(
    reader: &NsReader<&[u8]>,
    xml: &[u8],
    element: &BytesStart<'_>,
    element_span: Range<usize>,
    conformance: Conformance,
) -> Result<(String, Range<usize>)> {
    let mut relationship_id = None;
    let mut qualified_name = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id" {
            continue;
        }
        let is_any_relationship_namespace = matches!(
            &namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == super::REL_TEXT.as_bytes()
                    || *value == super::STRICT_REL_TEXT.as_bytes()
        );
        if !is_any_relationship_namespace {
            continue;
        }
        if !relationship_namespace(&namespace, conformance) {
            return Err(invalid(
                "shape p:tags uses the wrong relationship namespace profile",
            ));
        }
        if relationship_id.is_some() {
            return Err(invalid("shape p:tags has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid("shape p:tags has an invalid relationship ID"));
        }
        relationship_id = Some(value);
        qualified_name = Some(attribute.key.as_ref().to_vec());
    }
    let relationship_id =
        relationship_id.ok_or_else(|| invalid("shape p:tags is missing required r:id"))?;
    let qualified_name = qualified_name
        .ok_or_else(|| invalid("shape p:tags relationship attribute name is missing"))?;
    let raw = xml
        .get(element_span.clone())
        .ok_or_else(|| invalid("shape p:tags start tag is outside owner XML"))?;
    let relative = attribute_value_span(raw, &qualified_name)?;
    let start = element_span
        .start
        .checked_add(relative.start)
        .ok_or_else(|| invalid("shape relationship attribute offset overflow"))?;
    let end = element_span
        .start
        .checked_add(relative.end)
        .ok_or_else(|| invalid("shape relationship attribute offset overflow"))?;
    Ok((relationship_id, start..end))
}

pub(super) fn attribute_value_span(element: &[u8], selected: &[u8]) -> Result<Range<usize>> {
    if element.first() != Some(&b'<') {
        return Err(invalid("shape p:tags start tag is malformed"));
    }
    let mut cursor = 1usize;
    while cursor < element.len()
        && !element[cursor].is_ascii_whitespace()
        && !matches!(element[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    let mut found = None;
    loop {
        while cursor < element.len() && element[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= element.len() || matches!(element[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < element.len()
            && !element[cursor].is_ascii_whitespace()
            && !matches!(element[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < element.len() && element[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if element.get(cursor) != Some(&b'=') {
            return Err(invalid("shape p:tags attribute has no '='"));
        }
        cursor += 1;
        while cursor < element.len() && element[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *element
            .get(cursor)
            .ok_or_else(|| invalid("shape p:tags attribute value is missing"))?;
        if !matches!(quote, b'\'' | b'\"') {
            return Err(invalid("shape p:tags attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < element.len() && element[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= element.len() {
            return Err(invalid("shape p:tags attribute value is unterminated"));
        }
        cursor += 1;
        if element.get(name_start..name_end) == Some(selected) {
            if found.is_some() {
                return Err(invalid("shape p:tags relationship attribute is duplicated"));
            }
            found = Some(value_start..value_end);
        }
    }
    found.ok_or_else(|| invalid("shape p:tags relationship attribute span is missing"))
}

fn has_non_namespace_attrs(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

fn take_active<I>(active: &mut std::iter::Peekable<I>, start: usize) -> Result<bool>
where
    I: Iterator<Item = u32>,
{
    let start = offset_u32(start)?;
    if active.peek().copied() == Some(start) {
        let _ = active.next();
        Ok(true)
    } else if active.peek().is_some_and(|offset| *offset < start) {
        Err(invalid("MCE-active shape offsets are out of source order"))
    } else {
        Ok(false)
    }
}

fn try_push<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    if values.len() == values.capacity() {
        values
            .try_reserve(1)
            .map_err(|source| allocation(resource, source))?;
    }
    values.push(value);
    Ok(())
}

fn bump_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "shape-tag owner XML nodes",
        limit: MAX_OWNER_NODES,
    })?;
    if *nodes > MAX_OWNER_NODES {
        Err(Error::Limit {
            resource: "shape-tag owner XML nodes",
            limit: MAX_OWNER_NODES,
        })
    } else {
        Ok(())
    }
}

fn offset_u32(offset: usize) -> Result<u32> {
    u32::try_from(offset).map_err(|_| invalid("shape-tag XML offset does not fit u32"))
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("shape-tag XML offset does not fit usize"))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tag::Tag;
    use std::sync::Arc;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn list(name: &str, value: &str) -> List {
        let mut list = List::new();
        list.add(Tag::new(name, value).expect("valid tag"))
            .expect("unique tag");
        list
    }

    fn owner_package(xml: Vec<u8>) -> (OpcPackage, PackURI) {
        let owner = PackURI::new("/ppt/slides/slide1.xml").expect("owner URI");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            xml,
        )));
        (package, owner)
    }

    fn shape_xml(name: &str, id: u32) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
        )
    }

    fn mce_shared_anchor_package() -> (OpcPackage, PackURI, PackURI) {
        let anchor = r#"<p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>"#;
        let active = shape_xml("Active", 2).replace("<p:ph/>", anchor);
        let inactive = shape_xml("Inactive", 3).replace("<p:ph/>", anchor);
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p14="{P14}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="p14">{active}</mc:Choice><mc:Fallback>{inactive}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
        let (mut package, owner) = owner_package(xml);
        package
            .get_part_mut(&owner)
            .expect("owner")
            .rels_mut()
            .add_relationship(
                crate::tag::TAG_REL.into(),
                "../tags/tag1.xml".into(),
                "rId1".into(),
                false,
            );
        let part = PackURI::new("/ppt/tags/tag1.xml").expect("part URI");
        package.add_part(Box::new(XmlPart::new(
            part.clone(),
            CONTENT_TYPE.into(),
            format!(r#"<p:tagLst xmlns:p="{PML}"><p:tag name="Owner" val="Alice"/></p:tagLst>"#)
                .into_bytes(),
        )));
        (package, owner, part)
    }

    #[test]
    fn maps_all_five_families_and_nested_groups_to_raw_source() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree>
                <p:nvGrpSpPr/><p:grpSpPr/>
                {}
                <p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>
                <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/></p:graphicFrame>
                <p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}</p:grpSp>
            </p:spTree></p:cSld></p:sld>"#,
            shape_xml("Auto", 2),
            shape_xml("Nested", 7),
        )
        .into_bytes();

        let expected = [
            ("Auto", b"<p:sp>".as_slice()),
            ("Picture", b"<p:pic>".as_slice()),
            ("Connector", b"<p:cxnSp>".as_slice()),
            ("Frame", b"<p:graphicFrame>".as_slice()),
            ("Group", b"<p:grpSp>".as_slice()),
            ("Nested", b"<p:sp>".as_slice()),
        ];
        for (index, (name, opening)) in expected.iter().enumerate() {
            let by_name = selected_raw_span(&xml, crate::shape::Key::Name(name))
                .expect("name maps to raw span");
            let by_index = selected_raw_span(&xml, crate::shape::Key::Index(index))
                .expect("index maps to raw span");
            assert_eq!(by_name, by_index);
            assert!(xml[by_name].starts_with(opening));

            let layout = scan_layout(&xml, by_index).expect("shape layout");
            let staged = add_anchor(&xml, &layout, "rId9").expect("anchor insertion");
            let staged_span = selected_raw_span(&staged, crate::shape::Key::Name(name))
                .expect("staged shape maps");
            let staged_layout = scan_layout(&staged, staged_span).expect("staged layout");
            assert_eq!(
                staged_layout
                    .anchor
                    .as_ref()
                    .map(|anchor| anchor.id.as_str()),
                Some("rId9")
            );
            let removed = remove_anchor(&staged, &staged_layout).expect("anchor removal");
            let removed_span = selected_raw_span(&removed, crate::shape::Key::Name(name))
                .expect("removed shape maps");
            assert!(
                scan_layout(&removed, removed_span)
                    .expect("removed layout")
                    .anchor
                    .is_none()
            );
        }
    }

    #[test]
    fn preserves_customer_data_and_inserts_before_extensions() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Ordered"/><p:cNvSpPr/><p:nvPr><p:ph/><p:audioFile r:link="rIdAudio"/><p:custDataLst keep="yes"><p:custData r:id="rIdData"/></p:custDataLst><!--keep--><p:extLst/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
            </p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
        let span = selected_raw_span(&xml, crate::shape::Key::Name("Ordered")).expect("shape");
        let layout = scan_layout(&xml, span).expect("layout");
        let staged = add_anchor(&xml, &layout, "rIdTags").expect("insert");
        let text = std::str::from_utf8(&staged).expect("UTF-8 fixture");
        assert!(text.find("<p:custData ").unwrap() < text.find("<p:tags ").unwrap());
        assert!(text.find("<p:tags ").unwrap() < text.find("</p:custDataLst>").unwrap());
        assert!(text.find("</p:custDataLst>").unwrap() < text.find("<!--keep-->").unwrap());
        assert!(text.find("<!--keep-->").unwrap() < text.find("<p:extLst").unwrap());

        let staged_span =
            selected_raw_span(&staged, crate::shape::Key::Name("Ordered")).expect("shape");
        let staged_layout = scan_layout(&staged, staged_span).expect("layout");
        let removed = remove_anchor(&staged, &staged_layout).expect("remove");
        assert_eq!(removed, xml);
    }

    #[test]
    fn mce_mapping_edits_only_the_active_raw_branch() {
        let p14 = P14;
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p14="{p14}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <mc:AlternateContent><mc:Choice Requires="p14">{}</mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="3" name="Inactive"/><p:cNvPicPr/><p:nvPr><p:custDataLst><p:tags r:id="rIdInactive"/></p:custDataLst></p:nvPr></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></mc:Fallback></mc:AlternateContent>
            {}
            </p:spTree></p:cSld></p:sld>"#,
            shape_xml("Active", 2),
            shape_xml("Second", 4),
        )
        .into_bytes();

        let active_span =
            selected_raw_span(&xml, crate::shape::Key::Index(0)).expect("active span");
        assert!(
            xml[active_span.clone()]
                .windows(13)
                .any(|bytes| bytes == b"name=\"Active\"")
        );
        assert!(
            !xml[active_span.clone()]
                .windows(15)
                .any(|bytes| bytes == b"name=\"Inactive\"")
        );
        let layout = scan_layout(&xml, active_span).expect("active layout");
        let staged = add_anchor(&xml, &layout, "rIdTags").expect("active insertion");
        let text = std::str::from_utf8(&staged).expect("UTF-8 fixture");
        assert_eq!(text.matches("<p:tags ").count(), 2);
        assert!(text.contains(r#"<p:tags r:id="rIdInactive"/>"#));
        let inactive = text.find("name=\"Inactive\"").expect("inactive branch");
        let inactive_end = text[inactive..].find("</p:pic>").expect("picture end") + inactive;
        assert!(!text[inactive..inactive_end].contains("rIdTags"));
    }

    #[test]
    fn inactive_mce_anchor_forces_replacement_fork_and_removal_retention() {
        let (mut package, owner, original_part) = mce_shared_anchor_package();

        let old = put(&mut package, &owner, "Active", list("Reviewer", "Bob"))
            .expect("replace active attachment")
            .expect("old list");
        assert_eq!(old.get("owner").expect("old tag").value(), "Alice");
        let active = load(&package, &owner, "Active")
            .expect("load active attachment")
            .expect("active attachment");
        assert_ne!(active.rel(), "rId1");
        assert_ne!(active.part(), &original_part);
        assert_eq!(
            active.list().get("reviewer").expect("new tag").value(),
            "Bob"
        );
        let owner_part = package.get_part(&owner).expect("owner");
        assert!(owner_part.rels().get("rId1").is_some());
        assert!(package.get_part(&original_part).is_ok());
        let owner_xml = std::str::from_utf8(owner_part.blob()).expect("UTF-8 fixture");
        assert_eq!(owner_xml.matches(r#"r:id="rId1""#).count(), 1);
        let inactive = owner_xml
            .find("name=\"Inactive\"")
            .expect("inactive branch");
        let inactive_end = owner_xml[inactive..]
            .find("</p:sp>")
            .expect("inactive shape end")
            + inactive;
        assert!(owner_xml[inactive..inactive_end].contains(r#"r:id="rId1""#));

        let (mut package, owner, original_part) = mce_shared_anchor_package();
        let removed = remove(&mut package, &owner, "Active")
            .expect("remove active attachment")
            .expect("old list");
        assert_eq!(removed.get("owner").expect("old tag").value(), "Alice");
        assert!(
            load(&package, &owner, "Active")
                .expect("load active")
                .is_none()
        );
        let owner_part = package.get_part(&owner).expect("owner");
        assert!(owner_part.rels().get("rId1").is_some());
        assert!(package.get_part(&original_part).is_ok());
        let owner_xml = std::str::from_utf8(owner_part.blob()).expect("UTF-8 fixture");
        assert_eq!(owner_xml.matches(r#"r:id="rId1""#).count(), 1);
        let inactive = owner_xml
            .find("name=\"Inactive\"")
            .expect("inactive branch");
        let inactive_end = owner_xml[inactive..]
            .find("</p:sp>")
            .expect("inactive shape end")
            + inactive;
        assert!(owner_xml[inactive..inactive_end].contains(r#"r:id="rId1""#));
    }

    #[test]
    fn shape_crud_is_atomic_noop_safe_and_move_based() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>"#,
            shape_xml("Title", 2),
        )
        .into_bytes();
        let (mut package, owner) = owner_package(xml);

        assert!(load(&package, &owner, "Title").expect("load").is_none());
        assert_eq!(
            put(&mut package, &owner, "Title", list("Owner", "Alice")).expect("create"),
            None
        );
        let source = load(&package, &owner, 0_usize)
            .expect("load")
            .expect("attachment");
        assert_eq!(source.list().get("owner").expect("tag").value(), "Alice");

        let owner_before = package.get_part(&owner).expect("owner").blob_arc();
        let part_before = package
            .get_part(source.part())
            .expect("tag part")
            .blob_arc();
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        let old = put(&mut package, &owner, "Title", list("Owner", "Alice"))
            .expect("no-op")
            .expect("old list");
        assert_eq!(old.get("OWNER").expect("old tag").value(), "Alice");
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).expect("owner").blob_arc()
        ));
        assert!(Arc::ptr_eq(
            &part_before,
            &package
                .get_part(source.part())
                .expect("tag part")
                .blob_arc()
        ));

        assert!(put(&mut package, &owner, "Missing", List::new()).is_err());
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).expect("owner").blob_arc()
        ));

        let removed = remove(&mut package, &owner, "Title")
            .expect("remove")
            .expect("old list");
        assert_eq!(removed.get("owner").expect("tag").value(), "Alice");
        assert!(!package.is_signed());
        assert!(load(&package, &owner, "Title").expect("load").is_none());
        assert!(
            remove(&mut package, &owner, "Title")
                .expect("idempotent remove")
                .is_none()
        );
    }

    #[test]
    fn strict_shape_crud_uses_strict_namespaces_and_relationship_type() {
        let xml = format!(
            r#"<p:sld xmlns:p="{STRICT_PML}" xmlns:r="{STRICT_REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>"#,
            shape_xml("Strict", 2),
        )
        .into_bytes();
        let (mut package, owner) = owner_package(xml);
        assert_eq!(
            put(&mut package, &owner, "Strict", List::new()).expect("strict create"),
            None
        );
        let source = load(&package, &owner, "Strict")
            .expect("strict load")
            .expect("strict attachment");
        assert_eq!(source.conformance(), Conformance::Strict);
        assert_eq!(
            package
                .get_part(&owner)
                .expect("owner")
                .rels()
                .get(source.rel())
                .expect("relationship")
                .reltype(),
            crate::tag::STRICT_TAG_REL
        );
        let owner_xml = std::str::from_utf8(package.get_part(&owner).expect("owner").blob())
            .expect("UTF-8 fixture");
        assert!(owner_xml.contains(STRICT_PML));
        assert!(owner_xml.contains(STRICT_REL));
        let part_xml =
            std::str::from_utf8(package.get_part(source.part()).expect("tag part").blob())
                .expect("UTF-8 tag part");
        assert!(part_xml.contains(STRICT_PML));
    }

    #[test]
    fn shared_shape_anchor_forks_then_collects_each_orphan() {
        let anchor = r#"<p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>"#;
        let first = shape_xml("First", 2).replace("<p:ph/>", anchor);
        let second = shape_xml("Second", 3).replace("<p:ph/>", anchor);
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{first}{second}</p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
        let (mut package, owner) = owner_package(xml);
        package
            .get_part_mut(&owner)
            .expect("owner")
            .rels_mut()
            .add_relationship(
                crate::tag::TAG_REL.into(),
                "../tags/tag1.xml".into(),
                "rId1".into(),
                false,
            );
        let original_part = PackURI::new("/ppt/tags/tag1.xml").expect("part URI");
        package.add_part(Box::new(XmlPart::new(
            original_part.clone(),
            CONTENT_TYPE.into(),
            format!(r#"<p:tagLst xmlns:p="{PML}"><p:tag name="Owner" val="Alice"/></p:tagLst>"#)
                .into_bytes(),
        )));

        let old = put(&mut package, &owner, "First", list("Reviewer", "Bob"))
            .expect("fork")
            .expect("old list");
        assert_eq!(old.get("owner").expect("old tag").value(), "Alice");
        let first = load(&package, &owner, "First")
            .expect("first load")
            .expect("first attachment");
        let second = load(&package, &owner, "Second")
            .expect("second load")
            .expect("second attachment");
        assert_ne!(first.rel(), second.rel());
        assert_ne!(first.part(), second.part());
        assert_eq!(
            first.list().get("reviewer").expect("new tag").value(),
            "Bob"
        );
        assert_eq!(
            second.list().get("owner").expect("old tag").value(),
            "Alice"
        );

        let forked_part = first.part().clone();
        assert!(
            remove(&mut package, &owner, "First")
                .expect("remove first")
                .is_some()
        );
        assert!(package.get_part(&forked_part).is_err());
        assert!(package.get_part(&original_part).is_ok());
        assert!(
            remove(&mut package, &owner, "Second")
                .expect("remove second")
                .is_some()
        );
        assert!(package.get_part(&original_part).is_err());
    }
}
