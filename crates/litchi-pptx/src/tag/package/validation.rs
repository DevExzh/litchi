//! Bounded semantic and package-graph validation for package-level tags.

use super::super::codec::{parse_profiled, pml, xml_error};
use super::super::{
    CONTENT_TYPE, MAX_GRAPH_LINKS, MAX_GRAPH_PARTS, MAX_OWNER_DEPTH, MAX_OWNER_NODES,
    MAX_PART_NAME_BYTES, MAX_RELATIONSHIP_ID_BYTES, MAX_SOURCE_RELATIONSHIPS, PART_NAME_ATTEMPTS,
    allocation, invalid,
};
use super::model::{Anchor, AnchorIdentity, CommonSlidePhase, OwnerKind};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part as OpcPart};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) fn observe_common_slide_child(local: &[u8], phase: &mut CommonSlidePhase) -> Result<()> {
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

pub(super) fn observe_customer_data_child(
    is_pml: bool,
    local: &[u8],
    tags_seen: &mut bool,
) -> Result<()> {
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

pub(super) fn owner_kind(root: &[u8], content_type: &str) -> Result<OwnerKind> {
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

pub(super) fn presentation_later(local: &[u8]) -> bool {
    matches!(
        local,
        b"kinsoku" | b"defaultTextStyle" | b"modifyVerifier" | b"extLst"
    )
}

pub(super) fn anchor_relationship_id(
    xml: &[u8],
    start: usize,
    end: usize,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: super::super::Conformance,
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
        let relative = super::super::shape::attribute_value_span(raw, attribute.key.as_ref())?;
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
    conformance: super::super::Conformance,
) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == conformance.relationship_namespace().as_bytes()
    )
}

pub(super) fn has_non_namespace_attrs(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(super) fn owner_anchor_uses(xml: &[u8], relationship_id: &str) -> Result<usize> {
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
    expected: super::super::Conformance,
) -> Result<super::super::Source> {
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
    Ok(super::super::Source {
        relationship_id: anchor.id.clone(),
        part_name: part.partname().clone(),
        conformance,
        list,
    })
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

pub(super) fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
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

pub(super) fn bump_graph_link(scanned: &mut usize) -> Result<()> {
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

pub(super) fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("tag-owner XML offset overflow"))
}

pub(super) fn bump_owner_node(nodes: &mut usize) -> Result<()> {
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

pub(super) fn checked_owner_depth(depth: usize) -> Result<usize> {
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
