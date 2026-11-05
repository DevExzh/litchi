//! Source-native primary slide insertion.
//!
//! Insertion is represented as one added slide image plus a lexical
//! presentation-root splice.  The topology module turns that description into
//! the same prepared OPC revision used by removal, ordering, and text edits.
//! This module performs the eager-compatible authoring and admission work
//! without constructing an eager package or reopening a ZIP.

use std::collections::{HashMap, HashSet};
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, SourceLimits, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::topology_wire::{Change, PartImage};
use super::{Package, allocation, limit};
use crate::presentation::order::{StableSlideKey, stable_slide_key};
use crate::{Error, Result};

const MIN_SLIDE_ID: u32 = 256;
const MAX_SLIDE_ID: u32 = 2_147_483_647;
const INSERTION_WORKSPACE_BYTES: usize = 32 * 1024;

#[derive(Clone)]
struct EffectivePartDelta {
    uri: PackURI,
    before: bool,
    after: bool,
}

/// Build the source-native insertion recipe for one logical position.
pub(super) fn plan(
    package: &Package,
    position: Position,
    layout_name: &str,
    max_patch_bytes: Option<usize>,
) -> Result<(
    Arc<Vec<u8>>,
    Arc<[StableSlideKey]>,
    Change,
    Option<Arc<litchi_core::Reservation>>,
)> {
    package.check()?;
    let position = position.get();
    if position > package.order.len() {
        return Err(Error::SlideIndexOutOfBounds {
            index: position,
            len: package.order.len().saturating_add(1),
        });
    }
    if package.order.len() >= crate::parts::MAX_SLIDES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Slides),
            observed: package.order.len().checked_add(1).unwrap_or(usize::MAX),
            maximum: crate::parts::MAX_SLIDES,
        });
    }
    let selector_limit = crate::opened::Limits::default().max_text_bytes();
    if layout_name.len() > selector_limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::NameBytes),
            observed: layout_name.len(),
            maximum: selector_limit,
        });
    }

    let source = &package.owner.package;
    source.check_execution()?;
    source.validate_topology_source_boundary()?;
    if !source.tolerance_report().is_clean() {
        return Err(unsafe_edit("the source manifests required ingress repairs"));
    }

    let presentation_uri = package.owner._presentation.source_partname();
    let presentation = source.part(presentation_uri)?;
    if presentation.content_type() != ct::PML_PRESENTATION_MAIN {
        return Err(Error::ContentType {
            expected: ct::PML_PRESENTATION_MAIN.to_owned(),
            actual: presentation.content_type().to_owned(),
        });
    }
    let physical_root = package.owner._presentation.source_data().as_bytes();
    crate::source_limits::validate_xml_limits(physical_root, package.owner.limits)?;

    let strict = strict_presentation(physical_root)?;

    // A structural revision can remove a slide and the exclusive dependent
    // Parts it owns.  Count the final catalog with a bounded index over only
    // those changed/dependent URIs; the physical source pass remains the
    // existing metadata-only count and never copies every source name.
    let effective_parts = effective_part_count(
        source,
        package.topology.as_deref(),
        package.owner.limits.max_parts(),
    )?;
    if effective_parts >= package.owner.limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Parts),
            observed: effective_parts.checked_add(1).unwrap_or(usize::MAX),
            maximum: package.owner.limits.max_parts(),
        });
    }

    let layout = registered_layout(source, &presentation, physical_root, layout_name)?;

    let slide_id = next_slide_id(package)?;
    let relationship_id = next_relationship_id(package, package.root.as_bytes())?;
    // Validate the lexical root before creating the new entry. This catches
    // references to a candidate relationship ID that are not represented by
    // the OPC graph (custom-show and other dangling r:id users).
    crate::opened::validate_insertion_owner(package.root.as_bytes(), &relationship_id)
        .map_err(map_error)?;

    let part_uri = next_slide_part_name(package)?;

    let shape_limits =
        crate::open_options::selected(package.owner.semantic_limits.as_ref()).shape();
    let placeholders = {
        let layout_part = source.part(&layout)?;
        let data = layout_part.data()?;
        crate::opened::plan_new_slide_placeholders(data.as_bytes(), shape_limits)?
    };
    let (slide_xml, relationships_xml) = authored_slide(
        package.owner.limits,
        part_uri.clone(),
        layout,
        slide_id,
        strict,
        &placeholders,
        shape_limits,
    )?;
    crate::opened::read_back_new_slide(&slide_xml, &placeholders, shape_limits)?;
    crate::source_limits::validate_xml_limits(&slide_xml, package.owner.limits)?;
    crate::opened::validate_removal_slide_surface(&slide_xml).map_err(map_error)?;
    crate::source_limits::validate_xml_limits(&relationships_xml, package.owner.limits)?;

    let root = crate::opened::insert_slide_binding(
        package.root.as_bytes(),
        package
            .order
            .iter()
            .map(|key| (key.slide_id, key.relationship_id.as_str())),
        position,
        slide_id,
        &relationship_id,
    )?;
    // The new slide joins the section of its neighbour, as the eager
    // editor places it (record 0303).
    let root = if crate::opened::sections::present(&root) {
        let mut ids: Vec<u32> = package.order.iter().map(|key| key.slide_id).collect();
        ids.insert(position, slide_id);
        crate::opened::sections::place(&root, &ids, position)?.unwrap_or(root)
    } else {
        root
    };
    if root.len() > package.owner.limits.max_part_bytes() as usize {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            root.len(),
            package.owner.limits.max_part_bytes() as usize,
        ));
    }

    // Run the shared source security/MCE/root-policy boundary against the
    // authored candidate. The final effective order proof is repeated by the
    // topology builder after it has installed the new stable key.
    crate::presentation::source_removal::validate_primary_root(
        source,
        presentation_uri,
        &root,
        &relationship_id,
    )
    .map_err(map_error)?;

    let key = stable_slide_key(part_uri, slide_id, relationship_id.clone());
    let mut order = Vec::new();
    order
        .try_reserve_exact(package.order.len().saturating_add(1))
        .map_err(allocation)?;
    order.extend(package.order.iter().take(position).cloned());
    order.push(key.clone());
    order.extend(package.order.iter().skip(position).cloned());

    let change = Change {
        key,
        content_type: ct::PML_SLIDE.to_owned(),
        relationship_type: if strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/slide".to_owned()
        } else {
            rt::SLIDE.to_owned()
        },
        before: None,
        after: Some(PartImage {
            xml: Arc::new(slide_xml),
            relationships: Some(Arc::new(relationships_xml)),
        }),
    };
    if let Some(max_patch_bytes) = max_patch_bytes {
        let mut planned = super::topology::planned_current_changes_for_admission(package)?;
        planned.push(super::topology::planned_change_for_admission(&change)?);
        let dependents = super::topology::planned_current_dependents_for_admission(package);
        super::topology::preflight_candidate_for_admission(
            package,
            &root,
            order.clone().into(),
            &planned,
            &dependents,
            max_patch_bytes,
        )?;
    }
    // Identity indexes and the authored slide/root splice allocate before the
    // prepared topology retains its recipe. Compute a checked workspace bound
    // only after the exact retained wire has been admitted.
    let identity_bytes = insertion_workspace_bound(package, source)?;
    let envelope = package
        .root
        .as_bytes()
        .len()
        .checked_add(metadata_bytes(
            source,
            package.owner.limits,
            presentation_uri,
        )?)
        .and_then(|size| size.checked_add(identity_bytes))
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::MemoryBytes,
                usize::MAX,
                usize::MAX,
            )
        })?;
    let managed_reservation = super::topology::reserve_payload(package, envelope)?.map(Arc::new);
    Ok((Arc::new(root), order.into(), change, managed_reservation))
}

fn effective_part_count(
    source: &litchi_opc::SourceBackedPackage,
    revision: Option<&super::topology::Revision>,
    maximum: usize,
) -> Result<usize> {
    let source_parts = source.iter_parts().count();
    let Some(revision) = revision else {
        return Ok(source_parts);
    };
    let entry_count = revision
        .data
        .changes
        .len()
        .checked_add(revision.data.dependents.len())
        .ok_or_else(|| limit(crate::error::CodecLimitKind::Parts, usize::MAX, maximum))?;
    let mut deltas: HashMap<String, EffectivePartDelta> = HashMap::new();
    deltas.try_reserve(entry_count).map_err(allocation)?;

    let mut add_delta = |uri: &PackURI, before: bool, after: bool| -> Result<()> {
        if !before && !after {
            return Ok(());
        }
        let key = uri.as_str().to_ascii_lowercase();
        if let Some(delta) = deltas.get_mut(&key) {
            // A duplicate URI across the two topology namespaces describes
            // one final catalog Part.  Presence wins when one overlapping
            // entry restores it; later topology validation still checks the
            // corresponding images and owners.
            delta.before |= before;
            delta.after |= after;
        } else {
            deltas.insert(
                key,
                EffectivePartDelta {
                    uri: uri.clone(),
                    before,
                    after,
                },
            );
        }
        Ok(())
    };
    for change in revision.data.changes.iter() {
        add_delta(
            &change.key.part_uri,
            change.before.is_some(),
            change.after.is_some(),
        )?;
    }
    for dependent in revision.data.dependents.iter() {
        add_delta(
            &dependent.part_uri,
            dependent.before.is_some(),
            dependent.after.is_some(),
        )?;
    }

    let mut additions = 0usize;
    let mut removals = 0usize;
    for delta in deltas.values() {
        let physical = match source.part(&delta.uri) {
            Ok(_) => true,
            Err(litchi_opc::OpcError::PartNotFound(_)) => false,
            Err(error) => return Err(Error::Opc(error)),
        };
        if delta.before && !physical {
            // Keep the old checked-subtraction refusal observable instead of
            // silently treating a stale dependent as an absent Part.
            return Err(Error::PatchConflict);
        }
        if delta.after && !physical {
            additions = additions
                .checked_add(1)
                .ok_or_else(|| limit(crate::error::CodecLimitKind::Parts, usize::MAX, maximum))?;
        } else if !delta.after && physical {
            removals = removals
                .checked_add(1)
                .ok_or_else(|| limit(crate::error::CodecLimitKind::Parts, usize::MAX, maximum))?;
        }
    }
    source_parts
        .checked_sub(removals)
        .ok_or(Error::PatchConflict)?
        .checked_add(additions)
        .ok_or_else(|| limit(crate::error::CodecLimitKind::Parts, usize::MAX, maximum))
}

/// Map the shared source topology refusal vocabulary to the eager editor's
/// operation-level error while preserving ordinary XML/OPC/limit failures.
pub(super) fn map_error(error: Error) -> Error {
    match error {
        // Invalid XML/OPC validation failures are malformed input, rather
        // than a supported-operation refusal.  Keep those errors intact.
        Error::Refused(_) => unsafe_edit(
            "the source is outside the supported dependency-free slide insertion topology",
        ),
        other => other,
    }
}

fn unsafe_edit(reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "insert_slide",
        reason,
    )
}

fn registered_layout(
    source: &litchi_opc::SourceBackedPackage,
    presentation_view: &litchi_opc::PartView<'_>,
    physical_root: &[u8],
    layout_name: &str,
) -> Result<PackURI> {
    crate::presentation::source_cross_copy::registered_source_layout_by_name(
        source,
        presentation_view,
        physical_root,
        layout_name,
        source.execution_context().as_ref(),
    )
    .map_err(map_error)
}

fn metadata_bytes(
    source: &litchi_opc::SourceBackedPackage,
    limits: SourceLimits,
    presentation_uri: &PackURI,
) -> Result<usize> {
    let relationships = source
        .relationships_data_for_with_limit(presentation_uri, limits.max_relationship_xml_bytes())?
        .ok_or(Error::PatchConflict)?;
    let content_types = source.content_types_data_with_limit(limits.max_content_types_bytes())?;
    relationships
        .as_bytes()
        .len()
        .checked_add(content_types.as_bytes().len())
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::MemoryBytes,
                usize::MAX,
                usize::MAX,
            )
        })
}

fn insertion_workspace_bound(
    package: &Package,
    source: &litchi_opc::SourceBackedPackage,
) -> Result<usize> {
    let mut bytes = package.root.as_bytes().len();
    add_workspace_bytes(&mut bytes, source.rels())?;
    for part in source.iter_parts() {
        add_workspace_text(&mut bytes, part.partname().as_str())?;
        add_workspace_text(&mut bytes, part.content_type())?;
        add_workspace_amount(&mut bytes, size_of::<PackURI>() + 128)?;
        add_workspace_bytes(&mut bytes, part.rels())?;
    }
    for key in package.order.iter() {
        add_workspace_text(&mut bytes, key.part_uri.as_str())?;
        add_workspace_text(&mut bytes, &key.relationship_id)?;
        add_workspace_amount(&mut bytes, size_of::<StableSlideKey>() + 96)?;
    }
    if let Some(revision) = &package.topology {
        for change in revision.data.changes.iter() {
            add_workspace_text(&mut bytes, change.key.part_uri.as_str())?;
            add_workspace_text(&mut bytes, &change.key.relationship_id)?;
            add_workspace_text(&mut bytes, &change.content_type)?;
            add_workspace_text(&mut bytes, &change.relationship_type)?;
            add_workspace_amount(&mut bytes, size_of::<Change>() + 256)?;
            for image in [change.before.as_ref(), change.after.as_ref()]
                .into_iter()
                .flatten()
            {
                add_workspace_amount(&mut bytes, image.xml.len())?;
                if let Some(relationships) = &image.relationships {
                    add_workspace_amount(&mut bytes, relationships.len())?;
                }
            }
        }
    }
    // Hash tables reserve spare buckets and the XML scanners retain one event
    // at a time. Doubling the measured strings/entries covers those capacities
    // while the fixed workspace covers the crate::builder::Slide and root splice.
    bytes = bytes.checked_mul(2).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::MemoryBytes,
            usize::MAX,
            usize::MAX,
        )
    })?;
    add_workspace_amount(&mut bytes, INSERTION_WORKSPACE_BYTES)?;
    Ok(bytes)
}

fn add_workspace_amount(total: &mut usize, amount: usize) -> Result<()> {
    *total = total.checked_add(amount).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::MemoryBytes,
            usize::MAX,
            usize::MAX,
        )
    })?;
    Ok(())
}

fn add_workspace_text(total: &mut usize, text: &str) -> Result<()> {
    add_workspace_amount(total, text.len())?;
    add_workspace_amount(total, size_of::<String>())
}

fn add_workspace_bytes(total: &mut usize, relationships: &litchi_opc::Relationships) -> Result<()> {
    for relationship in relationships.iter() {
        add_workspace_text(total, relationship.r_id())?;
        add_workspace_text(total, relationship.reltype())?;
        add_workspace_text(total, relationship.target_ref())?;
        add_workspace_amount(total, size_of::<litchi_opc::Relationship>() + 128)?;
    }
    Ok(())
}

fn strict_presentation(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(false)
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(true)
                    },
                    _ => Err(unsafe_edit("unsupported presentation namespace")),
                };
            },
            Event::Eof => return Err(Error::Invalid("missing presentation root".into())),
            _ => {},
        }
    }
}

fn authored_slide(
    limits: SourceLimits,
    part_uri: PackURI,
    layout: PackURI,
    slide_id: u32,
    strict: bool,
    placeholders: &[crate::opened::NewSlidePlaceholder],
    shape_limits: crate::shape::Limits,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut slide_xml = crate::opened::new_slide_xml(slide_id, placeholders, shape_limits)?;
    if strict {
        for (old, new) in [
            (
                "http://schemas.openxmlformats.org/presentationml/2006/main",
                "http://purl.oclc.org/ooxml/presentationml/main",
            ),
            (
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://purl.oclc.org/ooxml/drawingml/main",
            ),
            (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
            ),
        ] {
            slide_xml = slide_xml.replace(old, new);
        }
    }
    let mut slide = BlobPart::new(part_uri, ct::PML_SLIDE.to_owned(), slide_xml.into_bytes());
    let layout_type = if strict {
        "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout"
    } else {
        rt::SLIDE_LAYOUT
    };
    let base_uri = slide.partname().base_uri().to_owned();
    let target = layout.relative_ref(&base_uri);
    slide.rels_mut().try_add_relationship(
        layout_type.to_owned(),
        target,
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    let relationships = slide.rels().plan_canonical(limits)?.materialize(limits)?;
    Ok((slide.blob().to_vec(), relationships))
}

fn next_slide_part_name(package: &Package) -> Result<PackURI> {
    let mut used = HashSet::new();
    let part_count = package.owner.package.iter_parts().count();
    let relationship_count = package.owner.package.rels().len()
        + package
            .owner
            .package
            .iter_parts()
            .map(|part| part.rels().len())
            .sum::<usize>();
    let change_count = package
        .topology
        .as_ref()
        .map_or(0, |revision| revision.data.changes.len());
    let mut image_relationship_count = 0usize;
    if let Some(revision) = &package.topology {
        for change in revision.data.changes.iter() {
            for image in [change.before.as_ref(), change.after.as_ref()]
                .into_iter()
                .flatten()
            {
                if let Some(relationships) = &image.relationships {
                    image_relationship_count = image_relationship_count
                        .checked_add(relationship_entry_count(relationships)?)
                        .ok_or_else(|| {
                            Error::Invalid("slide insertion relationship count overflows".into())
                        })?;
                }
            }
        }
    }
    let capacity = part_count
        .checked_add(relationship_count)
        .and_then(|value| value.checked_add(image_relationship_count))
        .and_then(|value| value.checked_add(change_count))
        .and_then(|value| value.checked_add(8))
        .ok_or_else(|| Error::Invalid("slide insertion name reservation overflows".into()))?;
    used.try_reserve(capacity).map_err(allocation)?;
    let reserve = |name: &str, used: &mut HashSet<usize>| -> Result<()> {
        const PREFIX: &str = "/ppt/slides/slide";
        if !name
            .get(..PREFIX.len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(PREFIX))
            || !name
                .get(name.len().saturating_sub(4)..)
                .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".xml"))
        {
            return Ok(());
        }
        let Some(value) = name
            .get(PREFIX.len()..name.len().saturating_sub(4))
            .and_then(|value| value.parse::<usize>().ok())
        else {
            return Ok(());
        };
        used.insert(value);
        Ok(())
    };
    for part in package.owner.package.iter_parts() {
        reserve(part.partname().as_str(), &mut used)?;
    }
    for graph in std::iter::once(package.owner.package.rels())
        .chain(package.owner.package.iter_parts().map(|part| part.rels()))
    {
        for relationship in graph
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            reserve(relationship.target_partname()?.as_str(), &mut used)?;
        }
    }
    if let Some(revision) = &package.topology {
        for change in revision.data.changes.iter() {
            reserve(change.key.part_uri.as_str(), &mut used)?;
            for image in [change.before.as_ref(), change.after.as_ref()]
                .into_iter()
                .flatten()
            {
                if let Some(relationships) = &image.relationships {
                    reserve_relationship_targets(relationships, &change.key.part_uri, &mut used)?;
                }
            }
        }
    }
    let maximum = part_count
        .checked_add(used.len())
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| Error::Invalid("slide insertion name bound overflow".into()))?;
    for index in 1..=maximum {
        if used.contains(&index) {
            continue;
        }
        let candidate =
            PackURI::new(format!("/ppt/slides/slide{index}.xml")).map_err(Error::Invalid)?;
        match package.owner.package.part(&candidate) {
            Ok(_) => continue,
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(candidate),
            Err(error) => return Err(Error::Opc(error)),
        }
    }
    Err(unsafe_edit(
        "no collision-free slide part name is available",
    ))
}

fn reserve_relationship_targets(
    xml: &[u8],
    owner: &PackURI,
    used: &mut HashSet<usize>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let mut target = None;
                let mut external = false;
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    let local = key
                        .iter()
                        .rposition(|byte| *byte == b':')
                        .map_or(key, |index| &key[index + 1..]);
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if local == b"Target" {
                        target = Some(value.into_owned());
                    } else if local == b"TargetMode" {
                        external = value.as_ref() == "External";
                    }
                }
                if !external && let Some(target) = target {
                    let target =
                        PackURI::from_rel_ref(owner.base_uri(), &target).map_err(Error::Invalid)?;
                    reserve_slide_uri(target.as_str(), used)?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn relationship_entry_count(xml: &[u8]) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml);
    let mut count = 0usize;
    loop {
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("relationship count overflows".into()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(count)
}

fn reserve_slide_uri(name: &str, used: &mut HashSet<usize>) -> Result<()> {
    const PREFIX: &str = "/ppt/slides/slide";
    if !name
        .get(..PREFIX.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(PREFIX))
        || !name
            .get(name.len().saturating_sub(4)..)
            .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".xml"))
    {
        return Ok(());
    }
    let Some(value) = name
        .get(PREFIX.len()..name.len().saturating_sub(4))
        .and_then(|value| value.parse::<usize>().ok())
    else {
        return Ok(());
    };
    used.insert(value);
    Ok(())
}

fn next_slide_id(package: &Package) -> Result<u32> {
    let mut used = HashSet::new();
    let capacity = package
        .owner
        .slides
        .len()
        .checked_add(package.order.len())
        .and_then(|value| {
            package.topology.as_ref().map_or(Some(value), |revision| {
                value.checked_add(revision.data.changes.len().saturating_mul(2))
            })
        })
        .ok_or_else(|| Error::Invalid("slide ID reservation overflows".into()))?;
    used.try_reserve(capacity).map_err(allocation)?;
    for slide in package.owner.slides.iter() {
        if !(MIN_SLIDE_ID..=MAX_SLIDE_ID).contains(&slide.slide_id) {
            return Err(unsafe_edit(
                "an existing slide ID is outside the PresentationML range",
            ));
        }
        used.insert(slide.slide_id);
    }
    for key in package.order.iter() {
        if !(MIN_SLIDE_ID..=MAX_SLIDE_ID).contains(&key.slide_id) {
            return Err(unsafe_edit(
                "an effective slide ID is outside the PresentationML range",
            ));
        }
        used.insert(key.slide_id);
    }
    if let Some(revision) = &package.topology {
        for change in revision.data.changes.iter() {
            if !(MIN_SLIDE_ID..=MAX_SLIDE_ID).contains(&change.key.slide_id) {
                return Err(unsafe_edit(
                    "a retained slide ID is outside the PresentationML range",
                ));
            }
            used.insert(change.key.slide_id);
        }
    }
    let mut candidate = MIN_SLIDE_ID;
    while used.contains(&candidate) {
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| unsafe_edit("the PresentationML slide-ID space is exhausted"))?;
    }
    if candidate > MAX_SLIDE_ID {
        return Err(unsafe_edit(
            "the PresentationML slide-ID space is exhausted",
        ));
    }
    Ok(candidate)
}

fn next_relationship_id(package: &Package, root: &[u8]) -> Result<String> {
    let mut used = HashSet::<String>::new();
    let root_ids = root_relationship_ids(root)?;
    let capacity = package
        .owner
        ._presentation
        .source_relationships()
        .len()
        .checked_add(
            package
                .topology
                .as_ref()
                .map_or(0, |revision| revision.data.changes.len()),
        )
        .and_then(|value| value.checked_add(8))
        .and_then(|value| value.checked_add(root_ids.len()))
        .ok_or_else(|| Error::Invalid("relationship ID reservation overflows".into()))?;
    used.try_reserve(capacity).map_err(allocation)?;
    for relationship in package.owner._presentation.source_relationships().iter() {
        used.insert(relationship.r_id().to_owned());
    }
    if let Some(revision) = &package.topology {
        for change in revision.data.changes.iter() {
            used.insert(change.key.relationship_id.clone());
        }
    }
    let mut candidate = 1u32;
    loop {
        let value = format!("rId{candidate}");
        if !used.contains(&value) && !root_ids.contains(&value) {
            return Ok(value);
        }
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| unsafe_edit("the presentation relationship-ID space is exhausted"))?;
    }
}

fn root_relationship_ids(root: &[u8]) -> Result<HashSet<String>> {
    let mut output = HashSet::new();
    let mut reader = NsReader::from_reader(root);
    reader.config_mut().trim_text(false);
    loop {
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    let Some(separator) = key.iter().rposition(|byte| *byte == b':') else {
                        continue;
                    };
                    if &key[separator + 1..] != b"id" {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    // Resolve the namespace first so a producer's custom
                    // relationship prefix is treated like `r:`. Unknown or
                    // unbound prefixes are retained conservatively as well:
                    // activating one by adding a relationship would turn a
                    // previously dangling XML reference into a live edge.
                    let (attribute_namespace, _) =
                        reader.resolver().resolve_attribute(attribute.key);
                    let relationship_namespace = matches!(
                        attribute_namespace,
                        ResolveResult::Bound(Namespace(namespace))
                            if namespace
                                == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                || namespace
                                    == b"http://purl.oclc.org/ooxml/officeDocument/relationships"
                    );
                    if relationship_namespace || separator != 0 {
                        output.insert(value.into_owned());
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(output)
}
