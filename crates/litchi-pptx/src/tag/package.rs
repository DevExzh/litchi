use super::codec::{parse_profiled, pml, xml_error};
use super::model::*;
use super::*;
use litchi_ooxml_common::mce::{
    OffsetLimits, Capabilities, Limits, active_offsets, process_markup_compatibility,
};
use litchi_opc::{OpcPackage, PackURI, Part as OpcPart, XmlPart};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::ops::Range;

/// Discover and parse every internal tag-list relationship on one owner part.
///
/// This is deliberately low-level diagnostic inventory: it does not inspect
/// XML anchors, so its results can include shape-owned and unanchored parts.
/// Use [`load`] for the part-level semantic attachment. OPC relationship
/// storage does not retain XML source order, so results are returned in
/// ascending relationship-ID byte order.
pub fn discover(owner: &dyn OpcPart, package: &OpcPackage) -> Result<Vec<Source>> {
    let mut scanned = 0usize;
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(owner.rels().len().min(MAX_TAG_PARTS))
        .map_err(|source| allocation("tag relationship inventory", source))?;
    for relationship in owner.rels().iter() {
        bump_graph_link(&mut scanned)?;
        if !is_relationship(relationship.reltype()) {
            continue;
        }
        if relationships.len() == MAX_TAG_PARTS {
            return Err(Error::Limit {
                resource: "owner tag-list relationships",
                limit: MAX_TAG_PARTS,
            });
        }
        relationships.push(relationship);
    }
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut targets = HashSet::new();
    targets
        .try_reserve(relationships.len())
        .map_err(|source| allocation("tag target index", source))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("tag source inventory", source))?;
    for relationship in relationships {
        if relationship.is_external() {
            return Err(invalid(format!(
                "tag-list relationship '{}' cannot be external",
                relationship.r_id()
            )));
        }
        let requested_target = relationship.target_partname()?;
        let part = package.get_part(&requested_target)?;
        let target = part.partname().clone();
        if !targets.insert(target.as_str().to_ascii_lowercase()) {
            return Err(invalid(format!(
                "duplicate owner tag-list target '{target}'"
            )));
        }
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.into(),
                actual: part.content_type().into(),
            });
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(format!(
                "tag-list part '{target}' has unexpected relationships"
            )));
        }
        let (list, conformance) = parse_profiled(part.blob())?;
        output.push(Source {
            relationship_id: relationship.r_id().into(),
            part_name: target,
            conformance,
            list,
        });
    }
    Ok(output)
}

/// Load the optional tag list attached to the selected part-level object.
///
/// For a presentation this follows only direct
/// `p:presentation/p:custDataLst/p:tags`. For slides, layouts, masters, notes,
/// and handouts it follows only direct `p:cSld/p:custDataLst/p:tags`.
/// Shape-level `p:nvPr` anchors and unanchored tag relationships are ignored.
pub fn load(package: &OpcPackage, owner: &PackURI) -> Result<Option<Source>> {
    let owner = package.get_part(owner)?;
    let processed = process_owner_ooxml(owner.blob())?;
    let layout = scan_owner_xml(processed.as_ref(), owner.content_type())?;
    layout
        .anchor
        .as_ref()
        .map(|anchor| resolve_anchor(owner, package, anchor, layout.conformance))
        .transpose()
}

/// Create or replace the selected part-level object's optional tag list.
///
/// The list is moved into a staged tag part. Adding also stages the owner XML
/// anchor and relationship; replacing preserves the anchor and relationship ID
/// and forks the target only when another package edge shares it. A
/// byte-identical replacement is a signature-preserving no-op. The returned
/// value is the previous list, or `None` when a new attachment was created.
pub fn put(package: &mut OpcPackage, owner: &PackURI, list: List) -> Result<Option<List>> {
    let (owner_name, layout, attached, anchor_uses) = {
        let owner = package.get_part(owner)?;
        let layout = scan_owner_source(owner.blob(), owner.content_type())?;
        let attached = layout
            .anchor
            .as_ref()
            .map(|anchor| {
                let source = resolve_anchor(owner, package, anchor, layout.conformance)?;
                let relationship = owner.rels().get(source.rel()).ok_or_else(|| {
                    invalid("anchored tag-list relationship disappeared during preflight")
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
            .map(|anchor| owner_anchor_uses(owner.blob(), anchor.id.as_str()))
            .transpose()?
            .unwrap_or(0);
        (owner.partname().clone(), layout, attached, anchor_uses)
    };

    let conformance = attached
        .as_ref()
        .map_or(layout.conformance, |value| value.source.conformance);
    let xml = staged_xml(&list, conformance)?;

    if let Some(attached) = attached {
        if package.get_part(attached.source.part())?.blob() == xml {
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
            let owner_xml =
                replace_anchor_relationship_id(owner.blob(), &layout, relationship_id.as_str())?;
            let staged = scan_owner_source(&owner_xml, owner.content_type())?;
            if staged.conformance != layout.conformance
                || staged.anchor.as_ref().map(|anchor| anchor.id.as_str())
                    != Some(relationship_id.as_str())
            {
                return Err(invalid("staged tag-owner anchor did not round-trip"));
            }
            (relationship_id, Some(owner_xml))
        } else {
            (attached.source.relationship_id.clone(), None)
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
        package.add_part(Box::new(XmlPart::new(part_name, CONTENT_TYPE.into(), xml)));
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
        let staged = scan_owner_source(&owner_xml, owner.content_type())?;
        if staged.conformance != layout.conformance
            || staged.anchor.as_ref().map(|anchor| anchor.id.as_str())
                != Some(relationship_id.as_str())
        {
            return Err(invalid("staged tag-owner anchor did not round-trip"));
        }
        (relationship_id, part_name, target_ref, owner_xml)
    };
    package.validate_new_part_name(&part_name)?;
    {
        let owner = package.get_part_mut(&owner_name)?;
        let current = scan_owner_source(owner.blob(), owner.content_type())?;
        if current.anchor.is_some() || owner.rels().get(&relationship_id).is_some() {
            return Err(invalid("tag-owner graph changed during preflight"));
        }
        owner.set_blob(owner_xml);
        owner.rels_mut().add_relationship(
            layout.conformance.relationship().into(),
            target_ref,
            relationship_id,
            false,
        );
    }
    package.add_part(Box::new(XmlPart::new(part_name, CONTENT_TYPE.into(), xml)));
    package.unsign();
    Ok(None)
}

/// Remove the selected part-level object's optional tag list.
///
/// An absent attachment is an idempotent, signature-preserving `Ok(None)`.
/// Other customer-data children remain byte-for-byte intact; a customer-data
/// container is removed only when the tag anchor was its sole content. The tag
/// part is collected only when no other package edge retains it.
pub fn remove(package: &mut OpcPackage, owner: &PackURI) -> Result<Option<List>> {
    let (owner_name, layout, attached, owner_xml, retain_relationship, orphan) = {
        let owner = package.get_part(owner)?;
        let layout = scan_owner_source(owner.blob(), owner.content_type())?;
        let Some(anchor) = layout.anchor.as_ref() else {
            return Ok(None);
        };
        let source = resolve_anchor(owner, package, anchor, layout.conformance)?;
        let relationship = owner.rels().get(source.rel()).ok_or_else(|| {
            invalid("anchored tag-list relationship disappeared during preflight")
        })?;
        let relationship_type = relationship.reltype().to_owned();
        let owner_name = owner.partname().clone();
        let retain_relationship = owner_anchor_uses(owner.blob(), anchor.id.as_str())? > 1;
        let owner_xml = remove_anchor(owner.blob(), &layout)?;
        let staged = scan_owner_source(&owner_xml, owner.content_type())?;
        if staged.conformance != layout.conformance || staged.anchor.is_some() {
            return Err(invalid("staged tag-owner removal did not round-trip"));
        }
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
        let current = scan_owner_source(owner.blob(), owner.content_type())?;
        if current.anchor.as_ref().map(|anchor| anchor.id.as_str())
            != layout.anchor.as_ref().map(|anchor| anchor.id.as_str())
        {
            return Err(invalid("tag-owner anchor changed during preflight"));
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
enum OwnerKind {
    Presentation,
    CommonSlide,
}

#[derive(Debug)]
struct OwnerXml {
    conformance: Conformance,
    insertion: usize,
    container: Option<Container>,
    anchor: Option<Anchor>,
}

#[derive(Debug)]
struct Container {
    span: Range<usize>,
    close_start: usize,
    empty: bool,
    qualified_name: Vec<u8>,
    child_elements: usize,
    other_content: bool,
    preserve_when_empty: bool,
}

#[derive(Debug)]
pub(crate) struct Anchor {
    pub(crate) id: String,
    pub(crate) span: Range<usize>,
    pub(crate) id_value: Range<usize>,
}

struct OpenContainer {
    start: usize,
    depth: usize,
    qualified_name: Vec<u8>,
    child_elements: usize,
    other_content: bool,
    preserve_when_empty: bool,
    tags_seen: bool,
}

struct OpenAnchor {
    start: usize,
    depth: usize,
    id: String,
    id_value: Range<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CommonSlidePhase {
    Start,
    Background,
    Shapes,
    CustomerData,
    Controls,
    Extensions,
}

fn observe_common_slide_child(local: &[u8], phase: &mut CommonSlidePhase) -> Result<()> {
    *phase = match (local, *phase) {
        (b"bg", CommonSlidePhase::Start) => CommonSlidePhase::Background,
        (b"spTree", CommonSlidePhase::Start | CommonSlidePhase::Background) => {
            CommonSlidePhase::Shapes
        },
        (b"custDataLst", CommonSlidePhase::Shapes) => CommonSlidePhase::CustomerData,
        (b"controls", CommonSlidePhase::Shapes | CommonSlidePhase::CustomerData) => {
            CommonSlidePhase::Controls
        },
        (
            b"extLst",
            CommonSlidePhase::Shapes | CommonSlidePhase::CustomerData | CommonSlidePhase::Controls,
        ) => CommonSlidePhase::Extensions,
        (b"spTree", _) => {
            return Err(invalid("direct p:spTree is duplicated or out of order"));
        },
        (b"custDataLst", _) => {
            return Err(invalid(
                "direct p:custDataLst must follow p:spTree and precede later p:cSld children",
            ));
        },
        (b"bg" | b"controls" | b"extLst", _) => {
            return Err(invalid(
                "direct p:cSld children are duplicated or out of order",
            ));
        },
        _ => return Err(invalid("unsupported direct PresentationML p:cSld child")),
    };
    Ok(())
}

fn observe_customer_data_child(is_pml: bool, local: &[u8], tags_seen: &mut bool) -> Result<()> {
    if *tags_seen {
        return Err(invalid("p:tags must be the last p:custDataLst child"));
    }
    if !is_pml {
        return Err(invalid(
            "p:custDataLst contains an unsupported direct child",
        ));
    }
    match local {
        b"custData" => Ok(()),
        b"tags" => {
            *tags_seen = true;
            Ok(())
        },
        _ => Err(invalid(
            "p:custDataLst contains an unsupported direct child",
        )),
    }
}

fn scan_owner_xml(xml: &[u8], content_type: &str) -> Result<OwnerXml> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut kind = None;
    let mut conformance = None;
    let mut root_close = None;
    let mut presentation_insert = None;
    let mut presentation_later_seen = false;
    let mut common_slide_depth = None;
    let mut common_slide_seen = false;
    let mut common_slide_phase = CommonSlidePhase::Start;
    let mut after_shape_tree = None;
    let mut shape_tree_depth = None;
    let mut container = None;
    let mut open_container: Option<OpenContainer> = None;
    let mut anchor = None;
    let mut open_anchor: Option<OpenAnchor> = None;

    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let event_conformance = pml(&namespace);
        drop(namespace);
        let end = xml_position(&reader)?;
        match event {
            Event::Start(element) => {
                bump_owner_node(&mut nodes)?;
                if open_anchor.is_some() {
                    return Err(invalid("p:tags cannot contain child elements"));
                }
                if depth == 0 {
                    let found_kind = owner_kind(element.local_name().as_ref(), content_type)?;
                    let found_conformance = event_conformance.ok_or_else(|| {
                        invalid("tag-owner root has an unsupported namespace profile")
                    })?;
                    kind = Some(found_kind);
                    conformance = Some(found_conformance);
                } else {
                    let active_kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
                    let profile =
                        conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                    let is_pml = event_conformance == Some(profile);
                    let local = element.local_name();

                    if active_kind == OwnerKind::CommonSlide
                        && depth == 1
                        && is_pml
                        && local.as_ref() == b"cSld"
                    {
                        if common_slide_seen {
                            return Err(invalid("tag owner has multiple direct p:cSld elements"));
                        }
                        common_slide_seen = true;
                        common_slide_depth = Some(depth + 1);
                    }
                    if active_kind == OwnerKind::Presentation
                        && depth == 1
                        && is_pml
                        && presentation_later(local.as_ref())
                    {
                        presentation_insert.get_or_insert(start);
                        presentation_later_seen = true;
                    }
                    if active_kind == OwnerKind::Presentation
                        && depth == 1
                        && is_pml
                        && local.as_ref() == b"custDataLst"
                        && presentation_later_seen
                    {
                        return Err(invalid(
                            "direct presentation p:custDataLst must precede later root children",
                        ));
                    }
                    let target_depth = match active_kind {
                        OwnerKind::Presentation => Some(1),
                        OwnerKind::CommonSlide => common_slide_depth,
                    };
                    if active_kind == OwnerKind::CommonSlide
                        && target_depth == Some(depth)
                        && is_pml
                    {
                        observe_common_slide_child(local.as_ref(), &mut common_slide_phase)?;
                    }
                    if target_depth == Some(depth) && is_pml && local.as_ref() == b"custDataLst" {
                        if container.is_some() || open_container.is_some() {
                            return Err(invalid(
                                "tag owner has multiple direct p:custDataLst elements",
                            ));
                        }
                        open_container = Some(OpenContainer {
                            start,
                            depth: depth + 1,
                            qualified_name: element.name().as_ref().to_vec(),
                            child_elements: 0,
                            other_content: false,
                            preserve_when_empty: has_non_namespace_attrs(&element)?,
                            tags_seen: false,
                        });
                    } else if let Some(current) = open_container.as_mut()
                        && depth == current.depth
                    {
                        observe_customer_data_child(
                            is_pml,
                            local.as_ref(),
                            &mut current.tags_seen,
                        )?;
                        current.child_elements = current.child_elements.saturating_add(1);
                        if is_pml && local.as_ref() == b"tags" {
                            if anchor.is_some() || open_anchor.is_some() {
                                return Err(invalid(
                                    "p:custDataLst contains multiple direct p:tags anchors",
                                ));
                            }
                            let identity = anchor_relationship_id(
                                xml, start, end, &reader, &element, profile,
                            )?;
                            open_anchor = Some(OpenAnchor {
                                start,
                                depth: depth + 1,
                                id_value: identity.id_value,
                                id: identity.value,
                            });
                        }
                    }
                    if active_kind == OwnerKind::CommonSlide
                        && common_slide_depth == Some(depth)
                        && is_pml
                        && local.as_ref() == b"spTree"
                    {
                        shape_tree_depth = Some(depth + 1);
                    }
                }
                depth = checked_owner_depth(depth)?;
            },
            Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                if open_anchor.is_some() {
                    return Err(invalid("p:tags cannot contain child elements"));
                }
                if depth == 0 {
                    return Err(invalid("tag-owner root cannot be empty"));
                }
                let active_kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
                let profile = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                let is_pml = event_conformance == Some(profile);
                let local = element.local_name();
                if active_kind == OwnerKind::Presentation
                    && depth == 1
                    && is_pml
                    && presentation_later(local.as_ref())
                {
                    presentation_insert.get_or_insert(start);
                    presentation_later_seen = true;
                }
                if active_kind == OwnerKind::Presentation
                    && depth == 1
                    && is_pml
                    && local.as_ref() == b"custDataLst"
                    && presentation_later_seen
                {
                    return Err(invalid(
                        "direct presentation p:custDataLst must precede later root children",
                    ));
                }
                let target_depth = match active_kind {
                    OwnerKind::Presentation => Some(1),
                    OwnerKind::CommonSlide => common_slide_depth,
                };
                if active_kind == OwnerKind::CommonSlide && target_depth == Some(depth) && is_pml {
                    observe_common_slide_child(local.as_ref(), &mut common_slide_phase)?;
                }
                if target_depth == Some(depth) && is_pml && local.as_ref() == b"custDataLst" {
                    if container.is_some() || open_container.is_some() {
                        return Err(invalid(
                            "tag owner has multiple direct p:custDataLst elements",
                        ));
                    }
                    container = Some(Container {
                        span: start..end,
                        close_start: end,
                        empty: true,
                        qualified_name: element.name().as_ref().to_vec(),
                        child_elements: 0,
                        other_content: false,
                        preserve_when_empty: has_non_namespace_attrs(&element)?,
                    });
                } else if let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    observe_customer_data_child(is_pml, local.as_ref(), &mut current.tags_seen)?;
                    current.child_elements = current.child_elements.saturating_add(1);
                    if is_pml && local.as_ref() == b"tags" {
                        if anchor.is_some() || open_anchor.is_some() {
                            return Err(invalid(
                                "p:custDataLst contains multiple direct p:tags anchors",
                            ));
                        }
                        let identity =
                            anchor_relationship_id(xml, start, end, &reader, &element, profile)?;
                        anchor = Some(Anchor {
                            id_value: identity.id_value,
                            id: identity.value,
                            span: start..end,
                        });
                    }
                }
                if active_kind == OwnerKind::CommonSlide
                    && common_slide_depth == Some(depth)
                    && is_pml
                    && local.as_ref() == b"spTree"
                {
                    after_shape_tree = Some(end);
                }
            },
            Event::End(element) => {
                let profile = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                let is_pml = event_conformance == Some(profile);
                let local = element.local_name();
                if open_anchor
                    .as_ref()
                    .is_some_and(|value| value.depth == depth)
                {
                    if !is_pml || local.as_ref() != b"tags" {
                        return Err(invalid("malformed direct p:tags anchor"));
                    }
                    let value = open_anchor
                        .take()
                        .ok_or_else(|| invalid("tag anchor parser state is inconsistent"))?;
                    anchor = Some(Anchor {
                        id: value.id,
                        id_value: value.id_value,
                        span: value.start..end,
                    });
                }
                if open_container
                    .as_ref()
                    .is_some_and(|value| value.depth == depth)
                {
                    if !is_pml || local.as_ref() != b"custDataLst" {
                        return Err(invalid("malformed direct p:custDataLst"));
                    }
                    let value = open_container
                        .take()
                        .ok_or_else(|| invalid("customer-data parser state is inconsistent"))?;
                    container = Some(Container {
                        span: value.start..end,
                        close_start: start,
                        empty: false,
                        qualified_name: value.qualified_name,
                        child_elements: value.child_elements,
                        other_content: value.other_content,
                        preserve_when_empty: value.preserve_when_empty,
                    });
                }
                if shape_tree_depth == Some(depth) {
                    if !is_pml || local.as_ref() != b"spTree" {
                        return Err(invalid("malformed direct p:spTree"));
                    }
                    after_shape_tree = Some(end);
                    shape_tree_depth = None;
                }
                if common_slide_depth == Some(depth) {
                    if !is_pml || local.as_ref() != b"cSld" {
                        return Err(invalid("malformed direct p:cSld"));
                    }
                    if matches!(
                        common_slide_phase,
                        CommonSlidePhase::Start | CommonSlidePhase::Background
                    ) {
                        return Err(invalid("direct p:cSld has no p:spTree"));
                    }
                    common_slide_depth = None;
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("tag-owner XML depth underflow"))?;
            },
            Event::Text(text) => {
                let nonempty = !text.decode().map_err(xml_error)?.trim().is_empty();
                if open_anchor.is_some() && nonempty {
                    return Err(invalid("p:tags cannot contain text"));
                }
                if nonempty
                    && let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::CData(value) => {
                if open_anchor.is_some() && !value.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("p:tags cannot contain CDATA"));
                }
                if let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::Comment(_) => {
                if open_anchor.is_none()
                    && let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::Decl(_) if depth == 0 => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("tag-owner XML contains forbidden markup"));
            },
            Event::Eof => break,
        }
    }

    if depth != 0 || open_container.is_some() || open_anchor.is_some() {
        return Err(invalid("unterminated tag-owner XML"));
    }
    let kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
    let conformance = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
    let insertion = match kind {
        OwnerKind::Presentation => presentation_insert
            .or(root_close)
            .ok_or_else(|| invalid("presentation root is not closed"))?,
        OwnerKind::CommonSlide => {
            if !common_slide_seen {
                return Err(invalid("tag owner has no direct p:cSld"));
            }
            after_shape_tree.ok_or_else(|| invalid("direct p:cSld has no p:spTree"))?
        },
    };
    Ok(OwnerXml {
        conformance,
        insertion,
        container,
        anchor,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum OwnerMapName {
    Presentation,
    Slide,
    SlideLayout,
    SlideMaster,
    Notes,
    NotesMaster,
    HandoutMaster,
    ShapeTree,
    CustomerDataList,
    Tags,
    Kinsoku,
    DefaultTextStyle,
    ModifyVerifier,
    Extensions,
}

impl OwnerMapName {
    fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"presentation" => Some(Self::Presentation),
            b"sld" => Some(Self::Slide),
            b"sldLayout" => Some(Self::SlideLayout),
            b"sldMaster" => Some(Self::SlideMaster),
            b"notes" => Some(Self::Notes),
            b"notesMaster" => Some(Self::NotesMaster),
            b"handoutMaster" => Some(Self::HandoutMaster),
            b"spTree" => Some(Self::ShapeTree),
            b"custDataLst" => Some(Self::CustomerDataList),
            b"tags" => Some(Self::Tags),
            b"kinsoku" => Some(Self::Kinsoku),
            b"defaultTextStyle" => Some(Self::DefaultTextStyle),
            b"modifyVerifier" => Some(Self::ModifyVerifier),
            b"extLst" => Some(Self::Extensions),
            _ => None,
        }
    }

    fn is_owner_root(self) -> bool {
        matches!(
            self,
            Self::Presentation
                | Self::Slide
                | Self::SlideLayout
                | Self::SlideMaster
                | Self::Notes
                | Self::NotesMaster
                | Self::HandoutMaster
        )
    }

    fn is_presentation_later(self) -> bool {
        matches!(
            self,
            Self::Kinsoku | Self::DefaultTextStyle | Self::ModifyVerifier | Self::Extensions
        )
    }
}

#[derive(Debug)]
struct OwnerMapElement {
    name: OwnerMapName,
    conformance: Conformance,
    span: Range<usize>,
    open_end: usize,
    close_start: usize,
    empty: bool,
    qualified_name: Vec<u8>,
    preserve_when_empty: bool,
}

fn owner_mce_capabilities() -> Capabilities {
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14);
    capabilities.understand_namespace(P15);
    capabilities
}

fn owner_mce_limits(max_bytes: usize) -> Limits {
    Limits {
        max_input_bytes: max_bytes,
        max_output_bytes: max_bytes,
        max_depth: MAX_OWNER_DEPTH,
        ..Limits::default()
    }
}

fn owner_active_offset_limits() -> OffsetLimits {
    OffsetLimits {
        max_source_bytes: MAX_OWNER_BYTES,
        max_offsets: MAX_OWNER_NODES,
        max_marked_bytes: MAX_OWNER_MARKED_BYTES,
        processing: owner_mce_limits(MAX_OWNER_MARKED_BYTES),
    }
}

pub(crate) fn process_owner_ooxml(xml: &[u8]) -> Result<std::borrow::Cow<'_, [u8]>> {
    process_pptx_ooxml(xml, MAX_OWNER_BYTES)
}

pub(crate) fn process_pptx_ooxml(
    xml: &[u8],
    max_bytes: usize,
) -> Result<std::borrow::Cow<'_, [u8]>> {
    Ok(
        process_markup_compatibility(xml, &owner_mce_capabilities(), &owner_mce_limits(max_bytes))?
            .xml,
    )
}

fn scan_owner_source(xml: &[u8], content_type: &str) -> Result<OwnerXml> {
    let processed = process_owner_ooxml(xml)?;
    if matches!(processed, std::borrow::Cow::Borrowed(_)) {
        return scan_owner_xml(xml, content_type);
    }
    let semantic = scan_owner_xml(processed.as_ref(), content_type)?;
    let source_offsets = owner_map_offsets(xml)?;
    let active = active_offsets(
        xml,
        &source_offsets,
        &owner_mce_capabilities(),
        &owner_active_offset_limits(),
    )?;
    let processed_elements = collect_owner_map_elements(processed.as_ref(), None)?;
    let source_elements = collect_owner_map_elements(xml, Some(&active))?;
    map_owner_source(xml, semantic, &processed_elements, source_elements)
}

fn owner_map_offsets(xml: &[u8]) -> Result<Vec<u32>> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let mut offsets = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let conformance = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(element) => {
                bump_owner_node(&mut nodes)?;
                let name = conformance
                    .and_then(|_| OwnerMapName::from_local(element.local_name().as_ref()));
                if name.is_some_and(|name| !(depth == 0 && name.is_owner_root())) {
                    offsets
                        .try_reserve(1)
                        .map_err(|source| allocation("tag-owner MCE map offsets", source))?;
                    offsets.push(u32::try_from(start).map_err(|_| {
                        invalid("tag-owner MCE map offset exceeds the compact u32 domain")
                    })?);
                }
                depth = checked_owner_depth(depth)?;
            },
            Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                let name = conformance
                    .and_then(|_| OwnerMapName::from_local(element.local_name().as_ref()));
                if name.is_some_and(|name| !(depth == 0 && name.is_owner_root())) {
                    offsets
                        .try_reserve(1)
                        .map_err(|source| allocation("tag-owner MCE map offsets", source))?;
                    offsets.push(u32::try_from(start).map_err(|_| {
                        invalid("tag-owner MCE map offset exceeds the compact u32 domain")
                    })?);
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("tag-owner MCE offset depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(offsets)
}

fn collect_owner_map_elements(xml: &[u8], active: Option<&[u32]>) -> Result<Vec<OwnerMapElement>> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut active = active.map(|offsets| offsets.iter().copied().peekable());
    let mut reader = NsReader::from_reader(xml);
    let mut frames = Vec::new();
    let mut elements = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let conformance = pml(&namespace);
        drop(namespace);
        let end = xml_position(&reader)?;
        match event {
            Event::Start(element) => {
                bump_owner_node(&mut nodes)?;
                if frames.len() >= MAX_OWNER_DEPTH {
                    return Err(Error::Limit {
                        resource: "tag-owner XML nesting depth",
                        limit: MAX_OWNER_DEPTH,
                    });
                }
                let name = conformance
                    .and_then(|_| OwnerMapName::from_local(element.local_name().as_ref()));
                let source_root =
                    frames.is_empty() && name.is_some_and(OwnerMapName::is_owner_root);
                let selected = match (&mut active, name) {
                    (_, Some(_)) if source_root => true,
                    (Some(offsets), Some(_)) => take_owner_map_offset(offsets, start)?,
                    (None, Some(_)) => true,
                    _ => false,
                };
                let record =
                    if let (true, Some(name), Some(conformance)) = (selected, name, conformance) {
                        let index = elements.len();
                        elements
                            .try_reserve(1)
                            .map_err(|source| allocation("tag-owner MCE source map", source))?;
                        elements.push(owner_map_element(
                            name,
                            conformance,
                            &element,
                            start,
                            end,
                            false,
                        )?);
                        Some(index)
                    } else {
                        None
                    };
                frames
                    .try_reserve(1)
                    .map_err(|source| allocation("tag-owner MCE source stack", source))?;
                frames.push(record);
            },
            Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                let name = conformance
                    .and_then(|_| OwnerMapName::from_local(element.local_name().as_ref()));
                let selected = match (&mut active, name) {
                    (Some(offsets), Some(_)) => take_owner_map_offset(offsets, start)?,
                    (None, Some(_)) => true,
                    _ => false,
                };
                if let (true, Some(name), Some(conformance)) = (selected, name, conformance) {
                    elements
                        .try_reserve(1)
                        .map_err(|source| allocation("tag-owner MCE source map", source))?;
                    elements.push(owner_map_element(
                        name,
                        conformance,
                        &element,
                        start,
                        end,
                        true,
                    )?);
                }
            },
            Event::End(_) => {
                let record = frames
                    .pop()
                    .ok_or_else(|| invalid("tag-owner MCE source stack underflow"))?;
                if let Some(index) = record {
                    let value = elements
                        .get_mut(index)
                        .ok_or_else(|| invalid("tag-owner MCE source element was lost"))?;
                    value.close_start = start;
                    value.span.end = end;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("tag-owner XML contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !frames.is_empty() {
        return Err(invalid("tag-owner MCE source XML is unterminated"));
    }
    if active
        .as_mut()
        .is_some_and(|offsets| offsets.peek().is_some())
    {
        return Err(invalid(
            "tag-owner MCE active offsets were not consumed in source order",
        ));
    }
    Ok(elements)
}

fn owner_map_element(
    name: OwnerMapName,
    conformance: Conformance,
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    empty: bool,
) -> Result<OwnerMapElement> {
    let mut qualified_name = Vec::new();
    let preserve_when_empty = if name == OwnerMapName::CustomerDataList {
        qualified_name
            .try_reserve_exact(element.name().as_ref().len())
            .map_err(|source| allocation("tag-owner customer-data qualified name", source))?;
        qualified_name.extend_from_slice(element.name().as_ref());
        has_non_namespace_attrs(element)?
    } else {
        false
    };
    Ok(OwnerMapElement {
        name,
        conformance,
        span: start..end,
        open_end: end,
        close_start: end,
        empty,
        qualified_name,
        preserve_when_empty,
    })
}

fn take_owner_map_offset<I>(active: &mut std::iter::Peekable<I>, start: usize) -> Result<bool>
where
    I: Iterator<Item = u32>,
{
    let start = u32::try_from(start)
        .map_err(|_| invalid("tag-owner MCE source offset exceeds the compact u32 domain"))?;
    if active.peek().copied() == Some(start) {
        let _ = active.next();
        Ok(true)
    } else if active.peek().is_some_and(|offset| *offset < start) {
        Err(invalid(
            "tag-owner MCE active offsets are not in source order",
        ))
    } else {
        Ok(false)
    }
}

fn map_owner_source(
    xml: &[u8],
    semantic: OwnerXml,
    processed: &[OwnerMapElement],
    mut source: Vec<OwnerMapElement>,
) -> Result<OwnerXml> {
    if processed.len() != source.len() {
        return Err(invalid(format!(
            "raw and MCE-processed tag-owner element maps have different lengths (raw {}, processed {})",
            source.len(),
            processed.len(),
        )));
    }
    for (processed, source) in processed.iter().zip(&source) {
        if processed.name != source.name || processed.conformance != source.conformance {
            return Err(invalid(
                "raw and MCE-processed tag-owner element maps diverge",
            ));
        }
    }
    let insertion = map_owner_insertion(semantic.insertion, processed, &source)?;
    let anchor = if let Some(value) = semantic.anchor.as_ref() {
        let index = mapped_owner_index(
            processed,
            value.span.start,
            OwnerMapName::Tags,
            "direct p:tags",
        )?;
        let raw = source
            .get(index)
            .ok_or_else(|| invalid("raw direct p:tags mapping is missing"))?;
        let identity = source_anchor_identity(xml, raw, semantic.conformance)?;
        if identity.value != value.id {
            return Err(invalid("raw and MCE-processed direct p:tags IDs diverge"));
        }
        Some(Anchor {
            id: identity.value,
            id_value: identity.id_value,
            span: raw.span.clone(),
        })
    } else {
        None
    };
    let container = if let Some(value) = semantic.container {
        let index = mapped_owner_index(
            processed,
            value.span.start,
            OwnerMapName::CustomerDataList,
            "direct p:custDataLst",
        )?;
        let raw = source
            .get_mut(index)
            .ok_or_else(|| invalid("raw direct p:custDataLst mapping is missing"))?;
        let contains_only_anchor = anchor
            .as_ref()
            .map(|anchor| raw_container_contains_only_anchor(xml, raw, anchor))
            .transpose()?
            .unwrap_or(false);
        let other_content = value.other_content || (anchor.is_some() && !contains_only_anchor);
        Some(Container {
            span: raw.span.clone(),
            close_start: raw.close_start,
            empty: raw.empty,
            qualified_name: std::mem::take(&mut raw.qualified_name),
            child_elements: value.child_elements,
            other_content,
            preserve_when_empty: raw.preserve_when_empty,
        })
    } else {
        None
    };
    Ok(OwnerXml {
        conformance: semantic.conformance,
        insertion,
        container,
        anchor,
    })
}

fn mapped_owner_index(
    processed: &[OwnerMapElement],
    start: usize,
    expected: OwnerMapName,
    resource: &'static str,
) -> Result<usize> {
    processed
        .iter()
        .position(|element| element.span.start == start && element.name == expected)
        .ok_or_else(|| {
            invalid(format!(
                "MCE-processed {resource} has no raw-source mapping"
            ))
        })
}

fn map_owner_insertion(
    insertion: usize,
    processed: &[OwnerMapElement],
    source: &[OwnerMapElement],
) -> Result<usize> {
    for (index, element) in processed.iter().enumerate() {
        let mapped = if element.name == OwnerMapName::ShapeTree && element.span.end == insertion {
            source.get(index).map(|value| value.span.end)
        } else if element.name.is_presentation_later() && element.span.start == insertion {
            source.get(index).map(|value| value.span.start)
        } else if element.name.is_owner_root() && element.close_start == insertion {
            source.get(index).map(|value| value.close_start)
        } else {
            None
        };
        if let Some(mapped) = mapped {
            return Ok(mapped);
        }
    }
    Err(invalid(
        "MCE-processed tag-owner insertion point has no raw-source mapping",
    ))
}

fn source_anchor_identity(
    xml: &[u8],
    selected: &OwnerMapElement,
    conformance: Conformance,
) -> Result<AnchorIdentity> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let is_selected = start == selected.span.start
            && pml(&namespace) == Some(conformance)
            && matches!(&event, Event::Start(element) | Event::Empty(element) if element.local_name().as_ref() == b"tags");
        drop(namespace);
        let end = xml_position(&reader)?;
        if is_selected {
            let element = match event {
                Event::Start(element) | Event::Empty(element) => element,
                _ => return Err(invalid("raw direct p:tags mapping is not an element")),
            };
            if end != selected.open_end {
                return Err(invalid("raw direct p:tags opening span diverged"));
            }
            return anchor_relationship_id(xml, start, end, &reader, &element, conformance);
        }
        if start > selected.span.start || matches!(event, Event::Eof) {
            return Err(invalid("raw direct p:tags mapping was not found"));
        }
    }
}

fn raw_container_contains_only_anchor(
    xml: &[u8],
    container: &OwnerMapElement,
    anchor: &Anchor,
) -> Result<bool> {
    if container.empty {
        return Ok(false);
    }
    let before = xml
        .get(container.open_end..anchor.span.start)
        .ok_or_else(|| invalid("raw customer-data prefix range is invalid"))?;
    let after = xml
        .get(anchor.span.end..container.close_start)
        .ok_or_else(|| invalid("raw customer-data suffix range is invalid"))?;
    Ok(before.iter().all(u8::is_ascii_whitespace) && after.iter().all(u8::is_ascii_whitespace))
}

fn owner_kind(root: &[u8], content_type: &str) -> Result<OwnerKind> {
    let kind = match root {
        b"presentation"
            if matches!(
                content_type,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                    | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
                    | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
                    | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
            ) =>
        {
            OwnerKind::Presentation
        },
        b"sld"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"sldLayout"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"sldMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"notes"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"notesMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"handoutMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        _ => {
            return Err(Error::ContentType {
                expected: "PresentationML programmable-tag owner".into(),
                actual: content_type.into(),
            });
        },
    };
    Ok(kind)
}

fn presentation_later(local: &[u8]) -> bool {
    matches!(
        local,
        b"kinsoku" | b"defaultTextStyle" | b"modifyVerifier" | b"extLst"
    )
}

struct AnchorIdentity {
    value: String,
    id_value: Range<usize>,
}

fn anchor_relationship_id(
    xml: &[u8],
    start: usize,
    end: usize,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
) -> Result<AnchorIdentity> {
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id" || !relationship_namespace(&namespace, conformance) {
            continue;
        }
        if relationship_id.is_some() {
            return Err(invalid("p:tags has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid("p:tags has an invalid relationship ID"));
        }
        let raw = xml
            .get(start..end)
            .ok_or_else(|| invalid("p:tags opening span is outside owner XML"))?;
        let relative = shape::attribute_value_span(raw, attribute.key.as_ref())?;
        let value_start = start
            .checked_add(relative.start)
            .ok_or_else(|| invalid("p:tags relationship value offset overflow"))?;
        let value_end = start
            .checked_add(relative.end)
            .ok_or_else(|| invalid("p:tags relationship value offset overflow"))?;
        relationship_id = Some(AnchorIdentity {
            value,
            id_value: value_start..value_end,
        });
    }
    relationship_id.ok_or_else(|| invalid("p:tags is missing required r:id"))
}

pub(crate) fn relationship_namespace(
    namespace: &ResolveResult<'_>,
    conformance: Conformance,
) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == conformance.relationship_namespace().as_bytes()
    )
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

fn owner_anchor_uses(xml: &[u8], relationship_id: &str) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml);
    let mut nodes = 0usize;
    let mut uses = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let conformance = pml(&namespace);
        drop(namespace);
        let end = xml_position(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                if let Some(conformance) = conformance
                    && element.local_name().as_ref() == b"tags"
                {
                    let candidate =
                        anchor_relationship_id(xml, start, end, &reader, &element, conformance)?;
                    if candidate.value == relationship_id {
                        uses = uses.checked_add(1).ok_or(Error::Limit {
                            resource: "raw tag-owner anchor references",
                            limit: MAX_OWNER_NODES,
                        })?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(uses)
}

pub(crate) fn resolve_anchor(
    owner: &dyn OpcPart,
    package: &OpcPackage,
    anchor: &Anchor,
    expected: Conformance,
) -> Result<Source> {
    let relationship = owner.rels().get(&anchor.id).ok_or_else(|| {
        invalid(format!(
            "p:tags references missing relationship '{}'",
            anchor.id
        ))
    })?;
    if relationship.reltype() != expected.relationship() {
        return Err(invalid(format!(
            "p:tags relationship '{}' has type '{}' instead of the owner profile's '{}'",
            anchor.id,
            relationship.reltype(),
            expected.relationship(),
        )));
    }
    if relationship.is_external() {
        return Err(invalid("p:tags relationship cannot be external"));
    }
    let requested = relationship.target_partname()?;
    let part = package.get_part(&requested)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.into(),
            actual: part.content_type().into(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "tag-list part '{}' has unexpected relationships",
            part.partname()
        )));
    }
    let (list, conformance) = parse_profiled(part.blob())?;
    if conformance != expected {
        return Err(invalid(
            "tag-list namespace profile does not match its PresentationML owner",
        ));
    }
    Ok(Source {
        relationship_id: anchor.id.clone(),
        part_name: part.partname().clone(),
        conformance,
        list,
    })
}

fn add_anchor(xml: &[u8], layout: &OwnerXml, relationship_id: &str) -> Result<Vec<u8>> {
    if layout.anchor.is_some() {
        return Err(invalid("tag owner already has a direct p:tags anchor"));
    }
    let anchor = format!(
        "<p:tags xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"/>",
        layout.conformance.namespace(),
        layout.conformance.relationship_namespace(),
        relationship_id
    );
    if let Some(container) = &layout.container {
        if !container.empty {
            return insert_xml(xml, container.close_start, anchor.as_bytes());
        }
        let raw = xml
            .get(container.span.clone())
            .ok_or_else(|| invalid("customer-data span is outside owner XML"))?;
        let slash = raw
            .iter()
            .rposition(|byte| *byte == b'/')
            .ok_or_else(|| invalid("empty p:custDataLst has no closing slash"))?;
        let mut replacement = Vec::new();
        replacement.extend_from_slice(&raw[..slash]);
        replacement.extend_from_slice(&raw[slash + 1..]);
        replacement.extend_from_slice(anchor.as_bytes());
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&container.qualified_name);
        replacement.push(b'>');
        return replace_xml(xml, container.span.clone(), &replacement);
    }
    let container = format!(
        "<p:custDataLst xmlns:p=\"{}\">{anchor}</p:custDataLst>",
        layout.conformance.namespace()
    );
    insert_xml(xml, layout.insertion, container.as_bytes())
}

fn replace_anchor_relationship_id(
    xml: &[u8],
    layout: &OwnerXml,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("tag owner has no direct p:tags anchor"))?;
    replace_xml(xml, anchor.id_value.clone(), relationship_id.as_bytes())
}

fn remove_anchor(xml: &[u8], layout: &OwnerXml) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("tag owner has no direct p:tags anchor"))?;
    let container = layout
        .container
        .as_ref()
        .ok_or_else(|| invalid("direct p:tags has no p:custDataLst parent"))?;
    if container.child_elements == 1 && !container.other_content && !container.preserve_when_empty {
        replace_xml(xml, container.span.clone(), &[])
    } else {
        replace_xml(xml, anchor.span.clone(), &[])
    }
}

fn insert_xml(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    replace_xml(xml, offset..offset, value)
}

pub(crate) fn replace_xml(xml: &[u8], range: Range<usize>, value: &[u8]) -> Result<Vec<u8>> {
    let before = xml
        .get(..range.start)
        .ok_or_else(|| invalid("XML patch start is outside owner XML"))?;
    let after = xml
        .get(range.end..)
        .ok_or_else(|| invalid("XML patch end is outside owner XML"))?;
    let len = before
        .len()
        .checked_add(value.len())
        .and_then(|len| len.checked_add(after.len()))
        .ok_or_else(|| invalid("patched tag-owner XML length overflow"))?;
    if len > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "patched tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(len)
        .map_err(|source| allocation("tag-owner XML patch", source))?;
    output.extend_from_slice(before);
    output.extend_from_slice(value);
    output.extend_from_slice(after);
    Ok(output)
}

pub(crate) fn staged_xml(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    let xml = write(value, conformance)?;
    let (staged, staged_conformance) = parse_profiled(&xml)?;
    if staged_conformance != conformance || &staged != value {
        return Err(invalid("staged tag-list XML did not round-trip"));
    }
    Ok(xml)
}

pub(crate) fn available_relationship_id(owner: &dyn OpcPart) -> Result<String> {
    if owner.rels().len() >= MAX_SOURCE_RELATIONSHIPS {
        return Err(Error::Limit {
            resource: "tag-owner relationships",
            limit: MAX_SOURCE_RELATIONSHIPS,
        });
    }
    for number in 1..=MAX_SOURCE_RELATIONSHIPS {
        let candidate = format!("rId{number}");
        if owner.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        resource: "tag-list relationship-ID allocation attempts",
        limit: MAX_SOURCE_RELATIONSHIPS,
    })
}

pub(crate) fn available_part_name(package: &OpcPackage) -> Result<PackURI> {
    let existing = sorted_part_names(package)?;
    for number in 1..=PART_NAME_ATTEMPTS {
        let path = format!("/ppt/tags/tag{number}.xml");
        let candidate = PackURI::new(&path)
            .map_err(|error| invalid(format!("invalid generated tag-list part name: {error}")))?;
        if !part_name_conflicts(&existing, &path.to_ascii_lowercase()) {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        resource: "tag-list part-name allocation attempts",
        limit: PART_NAME_ATTEMPTS,
    })
}

fn sorted_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(Error::Limit {
            resource: "tag package part-name scan",
            limit: MAX_GRAPH_PARTS,
        });
    }
    let mut names = Vec::new();
    names
        .try_reserve_exact(package.part_count())
        .map_err(|source| allocation("tag package part-name index", source))?;
    let mut bytes = 0usize;
    for part in package.iter_parts() {
        bytes = bytes
            .checked_add(part.partname().as_str().len())
            .ok_or(Error::Limit {
                resource: "tag package part-name bytes",
                limit: MAX_PART_NAME_BYTES,
            })?;
        if bytes > MAX_PART_NAME_BYTES {
            return Err(Error::Limit {
                resource: "tag package part-name bytes",
                limit: MAX_PART_NAME_BYTES,
            });
        }
        names.push(part.partname().as_str().to_ascii_lowercase());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
    if existing
        .binary_search_by(|name| name.as_str().cmp(candidate))
        .is_ok()
    {
        return true;
    }
    for (index, _) in candidate.match_indices('/').skip(1) {
        if existing
            .binary_search_by(|name| name.as_str().cmp(&candidate[..index]))
            .is_ok()
        {
            return true;
        }
    }
    let descendant = format!("{candidate}/");
    let position = existing.partition_point(|name| name.as_str() < descendant.as_str());
    existing
        .get(position)
        .is_some_and(|name| name.starts_with(&descendant))
}

pub(crate) fn validate_relative_target(
    source: &PackURI,
    reference: &str,
    target: &PackURI,
) -> Result<()> {
    let resolved = PackURI::from_rel_ref(source.base_uri(), reference)
        .map_err(|error| invalid(format!("invalid generated tag-list target: {error}")))?;
    if resolved.is_equivalent_to(target) {
        Ok(())
    } else {
        Err(invalid("generated tag-list target resolves incorrectly"))
    }
}

pub(crate) fn validate_selected_relationship(
    owner: &dyn OpcPart,
    relationship_id: &str,
    relationship_type: &str,
    target: &PackURI,
) -> Result<()> {
    let relationship = owner.rels().get(relationship_id).ok_or_else(|| {
        invalid(format!(
            "tag-list relationship '{relationship_id}' is missing"
        ))
    })?;
    if relationship.is_external()
        || relationship.reltype() != relationship_type
        || !relationship.target_partname()?.is_equivalent_to(target)
    {
        return Err(invalid(
            "anchored tag-list relationship changed during preflight",
        ));
    }
    Ok(())
}

pub(crate) fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    selected_owner: &PackURI,
    selected_relationship: &str,
) -> Result<bool> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(Error::Limit {
            resource: "tag package graph parts",
            limit: MAX_GRAPH_PARTS,
        });
    }
    let mut scanned = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_link(&mut scanned)?;
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            bump_graph_link(&mut scanned)?;
            if source.partname().is_equivalent_to(selected_owner)
                && relationship.r_id() == selected_relationship
            {
                continue;
            }
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn bump_graph_link(scanned: &mut usize) -> Result<()> {
    *scanned = scanned.checked_add(1).ok_or(Error::Limit {
        resource: "tag package graph relationships",
        limit: MAX_GRAPH_LINKS,
    })?;
    if *scanned > MAX_GRAPH_LINKS {
        Err(Error::Limit {
            resource: "tag package graph relationships",
            limit: MAX_GRAPH_LINKS,
        })
    } else {
        Ok(())
    }
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("tag-owner XML offset overflow"))
}

fn bump_owner_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "tag-owner XML nodes",
        limit: MAX_OWNER_NODES,
    })?;
    if *nodes > MAX_OWNER_NODES {
        Err(Error::Limit {
            resource: "tag-owner XML nodes",
            limit: MAX_OWNER_NODES,
        })
    } else {
        Ok(())
    }
}

fn checked_owner_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("tag-owner XML depth overflow"))?;
    if depth > MAX_OWNER_DEPTH {
        Err(Error::Limit {
            resource: "tag-owner XML depth",
            limit: MAX_OWNER_DEPTH,
        })
    } else {
        Ok(depth)
    }
}
