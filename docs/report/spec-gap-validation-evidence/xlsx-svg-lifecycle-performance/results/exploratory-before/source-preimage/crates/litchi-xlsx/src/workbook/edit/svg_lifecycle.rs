//! Ordinary worksheet SVG attach/detach planning.
//!
//! The source drawing scanner owns semantic picture discovery.  This module
//! owns the mutation and the complete OPC dependency closure: one source XML
//! splice, one source relationship token, an optional SVG media part, and the
//! content-types delta required by that part.

use std::collections::HashSet;
use std::collections::HashMap;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OwnedContentTypes, OwnedElementEdit, OwnedElementUpdate, OwnedRelationships,
    OwnedXmlPart, PackURI, Part, Relationship, Relationships, TargetMode,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use super::model::{
    ContentTypesChange, GraphAction, PackageChange, PartChange, RelationshipChange, SvgPartChange,
};
use super::semantic::SvgLifecycleIntent;
use super::svg::PictureSelector;
use super::{Workbook, allocation, invalid};
use crate::drawing::source::{self, ElementRange, SourceDrawing, SvgOwnerState};
use crate::error::{Error, Result};

const MAX_DRAWING_BYTES: usize = 32 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 32 * 1024 * 1024;
const MAX_ID_ATTEMPTS: usize = 100_000;
const SVG_EXTENSION_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SVG_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const SPREADSHEETML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEETML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

fn output_limit(workbook: &Workbook) -> usize {
    let caller = workbook.inner.package.read_limits().max_part_bytes();
    usize::try_from(caller)
        .unwrap_or(usize::MAX)
        .min(MAX_OUTPUT_BYTES)
}

fn relationship_output_limit(workbook: &Workbook) -> usize {
    let caller = workbook.inner.package.read_limits();
    output_limit(workbook).min(caller.max_relationship_xml_bytes())
}

fn content_types_output_limit(workbook: &Workbook) -> usize {
    let caller = workbook.inner.package.read_limits();
    output_limit(workbook).min(caller.max_content_types_bytes())
}

pub(super) fn payload_limit(workbook: &Workbook) -> usize {
    output_limit(workbook)
}

/// Source-backed semantic facts cached by one transaction.  The cache owns
/// only bounded scalar projections; source ranges remain in the immutable
/// package and are rescanned by the commit planner when it builds its proof.
#[derive(Clone, Debug)]
pub(super) struct DrawingPreflight {
    drawing_uri: PackURI,
    pictures: Vec<CachedPicture>,
}

#[derive(Clone, Debug)]
struct CachedPicture {
    owner: ProjectedSvgOwner,
    raster_relationship_id: Box<str>,
}

fn scan_limits(workbook: &Workbook) -> source::ScanLimits {
    let caller = workbook.inner.package.read_limits();
    let defaults = source::ScanLimits::default();
    let part_bytes = usize::try_from(caller.max_part_bytes()).unwrap_or(usize::MAX);
    source::ScanLimits {
        max_xml_bytes: defaults.max_xml_bytes.min(part_bytes),
        max_nodes: defaults.max_nodes.min(caller.max_xml_events()),
        max_depth: defaults.max_depth.min(caller.max_xml_depth()),
        max_pictures: defaults
            .max_pictures
            .min(caller.max_relationships_per_part()),
        max_relationship_references: defaults
            .max_relationship_references
            .min(caller.max_relationships_per_part()),
        max_fragment_bytes: defaults.max_fragment_bytes.min(part_bytes),
    }
}

/// Stage every SVG lifecycle intent against one candidate package closure.
/// Multiple pictures in one drawing share one source XML and relationship
/// transition, while media and content-type deltas remain package-scoped.
pub(super) fn plan(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    parts: &mut Vec<PartChange>,
    relationships: &mut Vec<RelationshipChange>,
    svg_parts: &mut Vec<SvgPartChange>,
    content_types: &mut Option<ContentTypesChange>,
    package_changes: &mut Vec<PackageChange>,
) -> Result<()> {
    if intents.is_empty() {
        return Ok(());
    }
    plan_composed(
        workbook,
        intents,
        parts,
        relationships,
        svg_parts,
        content_types,
        package_changes,
    )
}

/// Preflight a public operation before its borrowed payload is copied.  A
/// detach with no admitted owner is an exact in-memory no-op.
/// Preflight an operation against the source package plus operations already
/// staged in this transaction.  The source scanner remains authoritative for
/// the initial owner state; this small projection prevents a later detach
/// from silently becoming a no-op after an earlier attach (and permits the
/// corresponding detach-then-attach sequence).
pub(super) fn preflight_with_pending(
    workbook: &Workbook,
    position: usize,
    selector: PictureSelector,
    attach: bool,
    pending: &[SvgLifecycleIntent],
    cache: &mut HashMap<(usize, usize), DrawingPreflight>,
) -> Result<bool> {
    let key = (position, selector_drawing(selector));
    if !cache.contains_key(&key) {
        cache
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle preflight cache", source))?;
        let facts = build_preflight_cache(workbook, position, selector)?;
        cache.insert(key, facts);
    }
    let facts = cache
        .get(&key)
        .ok_or_else(|| invalid("SVG lifecycle preflight cache disappeared"))?;
    let part = workbook.inner.package.get_part(&facts.drawing_uri)?;
    let picture = facts
        .pictures
        .get(selector_picture(selector))
        .ok_or_else(|| invalid("picture ordinal is outside worksheet drawing pictures"))?;
    validate_raster_fallback_id(workbook, part, &picture.raster_relationship_id)?;
    let mut owner = picture.owner;
    for prior in pending.iter().filter(|prior| {
        prior.position == position && prior.selector == selector
    }) {
        owner = project_owner(owner, prior.attach)?;
    }
    match (attach, owner) {
        (true, ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque) => Ok(true),
        (true, ProjectedSvgOwner::Embedded) => Err(invalid(
            "selected picture already has an embedded SVG owner",
        )),
        (true, ProjectedSvgOwner::Linked) => Err(invalid(
            "linked SVG owners cannot be replaced by attach_svg",
        )),
        (true, ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused) => Err(invalid(
            "selected picture has an ambiguous or refused SVG owner",
        )),
        (false, ProjectedSvgOwner::Embedded) => Ok(true),
        (false, ProjectedSvgOwner::Linked) => Err(invalid("linked SVG owners cannot be detached")),
        (false, ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused) => Err(invalid(
            "selected picture has an ambiguous or refused SVG owner",
        )),
        (false, ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque) => Ok(false),
    }
}

fn build_preflight_cache(
    workbook: &Workbook,
    position: usize,
    selector: PictureSelector,
) -> Result<DrawingPreflight> {
    let (drawing_uri, _) = resolve_drawing(workbook, position, selector)?;
    let part = workbook.inner.package.get_part(&drawing_uri)?;
    let scanned = SourceDrawing::scan_with_limits(
        part.blob(),
        selector_drawing(selector),
        scan_limits(workbook),
    )?;
    let mut pictures = Vec::new();
    pictures
        .try_reserve(scanned.pictures().len())
        .map_err(|source| allocation("SVG lifecycle preflight pictures", source))?;
    for picture in scanned.pictures() {
        let owner = match picture.svg_owner() {
            SvgOwnerState::None => ProjectedSvgOwner::None,
            SvgOwnerState::Opaque => ProjectedSvgOwner::Opaque,
            SvgOwnerState::Embedded(_) => ProjectedSvgOwner::Embedded,
            SvgOwnerState::Linked(_) => ProjectedSvgOwner::Linked,
            SvgOwnerState::Ambiguous => ProjectedSvgOwner::Ambiguous,
            SvgOwnerState::Refused => ProjectedSvgOwner::Refused,
        };
        pictures.push(CachedPicture {
            owner,
            raster_relationship_id: picture.raster_relationship_id().into(),
        });
    }
    Ok(DrawingPreflight {
        drawing_uri,
        pictures,
    })
}

#[derive(Clone, Copy, Debug)]
enum ProjectedSvgOwner {
    None,
    Opaque,
    Embedded,
    Linked,
    Ambiguous,
    Refused,
}

fn project_owner(owner: ProjectedSvgOwner, attach: bool) -> Result<ProjectedSvgOwner> {
    if attach {
        return match owner {
            ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => {
                Ok(ProjectedSvgOwner::Embedded)
            },
            ProjectedSvgOwner::Embedded => Err(invalid(
                "selected picture already has an embedded SVG owner",
            )),
            ProjectedSvgOwner::Linked => Err(invalid(
                "linked SVG owners cannot be replaced by attach_svg",
            )),
            ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => Err(invalid(
                "selected picture has an ambiguous or refused SVG owner",
            )),
        };
    }
    match owner {
        ProjectedSvgOwner::Embedded => Ok(ProjectedSvgOwner::None),
        ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => Ok(owner),
        ProjectedSvgOwner::Linked => Err(invalid("linked SVG owners cannot be detached")),
        ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => Err(invalid(
            "selected picture has an ambiguous or refused SVG owner",
        )),
    }
}

#[derive(Debug)]
struct DrawingPlanState {
    uri: PackURI,
    before_xml: OwnedXmlPart,
    current_xml: OwnedXmlPart,
    before_rels: OwnedRelationships,
    current_rels: OwnedRelationships,
    before_edges: Relationships,
    current_edges: Relationships,
}

fn new_drawing_plan_state(workbook: &Workbook, uri: PackURI) -> Result<DrawingPlanState> {
    let part = workbook.inner.package.get_part(&uri)?;
    let limits = scan_limits(workbook);
    if part.blob().len() > MAX_DRAWING_BYTES || part.blob().len() > limits.max_xml_bytes {
        return Err(invalid(format!(
            "SVG lifecycle drawing XML exceeds {} bytes",
            limits.max_xml_bytes.min(MAX_DRAWING_BYTES)
        )));
    }
    let before_xml = workbook.inner.package.source_xml_part(&uri)?;
    let before_rels = workbook
        .inner
        .package
        .source_relationships_with_limits(&uri, workbook.inner.package.read_limits())?;
    Ok(DrawingPlanState {
        current_xml: before_xml.clone(),
        current_rels: before_rels.clone(),
        before_xml,
        before_rels,
        before_edges: part.rels().clone(),
        current_edges: part.rels().clone(),
        uri,
    })
}

fn plan_composed(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    parts: &mut Vec<PartChange>,
    relationships: &mut Vec<RelationshipChange>,
    svg_parts: &mut Vec<SvgPartChange>,
    content_types: &mut Option<ContentTypesChange>,
    package_changes: &mut Vec<PackageChange>,
) -> Result<()> {
    let package_changes_start = package_changes.len();
    let parts_start = parts.len();
    let relationships_start = relationships.len();
    let svg_parts_start = svg_parts.len();
    let content_types_was_none = content_types.is_none();
    package_changes
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle package changes", source))?;
    let mut states = Vec::new();
    states
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing states", source))?;
    let mut state_by_uri = HashMap::new();
    state_by_uri
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing lookup", source))?;
    let mut state_groups: Vec<Vec<usize>> = Vec::new();
    state_groups
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing intent groups", source))?;
    let mut pending_parts = Vec::new();
    pending_parts
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle pending media parts", source))?;
    let mut reserved_uris = HashSet::new();
    reserved_uris
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle media identities", source))?;
    let mut cleanup_targets = Vec::new();
    cleanup_targets
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle media cleanup targets", source))?;
    let manifest_before = workbook
        .inner
        .package
        .source_content_types_with_limits(workbook.inner.package.read_limits())?;
    let mut manifest_after = manifest_before.clone();

    for (intent_index, intent) in intents.iter().enumerate() {
        let (drawing_uri, _) = resolve_drawing(workbook, intent.position, intent.selector)?;
        let state_index = if let Some(index) = state_by_uri.get(&drawing_uri).copied() {
            index
        } else {
            state_by_uri.insert(drawing_uri.clone(), states.len());
            states.push(new_drawing_plan_state(workbook, drawing_uri.clone())?);
            state_groups.push(Vec::new());
            states.len() - 1
        };
        state_groups
            .get_mut(state_index)
            .ok_or_else(|| invalid("SVG lifecycle drawing intent group disappeared"))?
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle drawing intent group", source))?;
        state_groups[state_index].push(intent_index);
        package_changes.push(PackageChange::SvgLifecycle {
            sheet: workbook.inner.sheets[intent.position].name.as_str().into(),
            drawing: selector_drawing(intent.selector),
            picture: selector_picture(intent.selector),
            attached: intent.attach,
        });
    }

    for state_index in 0..state_groups.len() {
        let group = &state_groups[state_index];
        let all_attach = group.iter().all(|index| intents[*index].attach);
        let unique_selectors = if all_attach {
            let mut selectors = HashSet::new();
            selectors
                .try_reserve(group.len())
                .map_err(|source| allocation("SVG lifecycle picture selectors", source))?;
            group
                .iter()
                .all(|index| selectors.insert(intents[*index].selector))
        } else {
            false
        };
        if all_attach && unique_selectors {
            plan_composed_attach_batch(
                workbook,
                group,
                intents,
                &mut states[state_index],
                &mut pending_parts,
                &mut reserved_uris,
                &mut manifest_after,
            )?;
        } else {
            for index in group {
                plan_composed_intent(
                    workbook,
                    &intents[*index],
                    &mut states[state_index],
                    &mut pending_parts,
                    &mut reserved_uris,
                    &mut manifest_after,
                    &mut cleanup_targets,
                )?;
            }
        }
    }

    finalize_svg_target_cleanup(
        workbook,
        &states,
        &mut pending_parts,
        &cleanup_targets,
        &mut manifest_after,
        content_types_output_limit(workbook),
    )?;

    for state in states {
        if state.current_xml.bytes() != state.before_xml.bytes() {
            parts.push(PartChange {
                uri: state.uri.clone(),
                before: state.before_xml.shared_bytes(),
                after: state.current_xml.shared_bytes(),
                before_source: Some(state.before_xml),
                after_source: Some(state.current_xml),
            });
        }
        if let Some(id) = changed_relationship_id(&state.before_edges, &state.current_edges)? {
            relationships.push(RelationshipChange {
                owner: state.uri,
                before: state.before_edges.get(&id).cloned(),
                after: state.current_edges.get(&id).cloned(),
                before_source: Some(state.before_rels),
                after_source: Some(state.current_rels),
            });
        }
    }
    svg_parts.extend(pending_parts);
    if manifest_after.bytes() != manifest_before.bytes() {
        *content_types = Some(ContentTypesChange {
            before: manifest_before,
            after: manifest_after,
        });
    }
    if content_types_was_none
        && parts.len() == parts_start
        && relationships.len() == relationships_start
        && svg_parts.len() == svg_parts_start
        && content_types.is_none()
    {
        package_changes.truncate(package_changes_start);
    }
    Ok(())
}

fn plan_composed_intent(
    workbook: &Workbook,
    intent: &SvgLifecycleIntent,
    state: &mut DrawingPlanState,
    pending_parts: &mut Vec<SvgPartChange>,
    reserved_uris: &mut HashSet<PackURI>,
    manifest: &mut OwnedContentTypes,
    cleanup_targets: &mut Vec<PackURI>,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let relationship_cap = relationship_output_limit(workbook);
    let content_types_cap = content_types_output_limit(workbook);
    let scanned = SourceDrawing::scan_with_limits(
        state.current_xml.bytes(),
        selector_drawing(intent.selector),
        limits,
    )?;
    let picture = scanned.picture(selector_picture(intent.selector))?;
    validate_raster_fallback_edges(workbook, &state.current_edges, picture)?;

    if intent.attach {
        let payload = intent
            .payload
            .as_deref()
            .ok_or_else(|| invalid("attach_svg has no staged payload"))?;
        if payload.len() > output_cap {
            return Err(invalid(format!(
                "SVG lifecycle payload exceeds {} bytes",
                output_cap
            )));
        }
        match picture.svg_owner() {
            SvgOwnerState::None | SvgOwnerState::Opaque => {},
            SvgOwnerState::Embedded(_) => {
                return Err(invalid(
                    "selected picture already has an embedded SVG owner",
                ));
            },
            SvgOwnerState::Linked(_) => {
                return Err(invalid(
                    "linked SVG owners cannot be replaced by attach_svg",
                ));
            },
            SvgOwnerState::Ambiguous | SvgOwnerState::Refused => {
                return Err(invalid(
                    "selected picture has an ambiguous or refused SVG owner",
                ));
            },
        }
        let relationship_id = allocate_relationship_id(&state.current_edges)?;
        let part_uri = allocate_part_uri_with_reservations(&workbook.inner.package, reserved_uris)?;
        reserved_uris.insert(part_uri.clone());
        let target_ref = part_uri.relative_ref(state.uri.base_uri());
        let relationship_type =
            if scanned.relationship_dialect() == source::RelationshipDialect::Strict {
                rt::STRICT_IMAGE
            } else {
                rt::IMAGE
            };
        let next_rels = state.current_rels.with_relationship(
            relationship_type,
            &target_ref,
            &relationship_id,
            TargetMode::Internal,
            relationship_cap,
        )?;
        let next_xml = attach_xml(
            &state.current_xml,
            picture,
            &relationship_id,
            scanned.dialect(),
            output_cap,
        )?;
        let relationship = Relationship::new_with_mode(
            relationship_id.clone(),
            relationship_type.to_owned(),
            target_ref,
            state.uri.base_uri().to_owned(),
            TargetMode::Internal,
        );
        if state.current_edges.len()
            >= workbook
                .inner
                .package
                .read_limits()
                .max_relationships_per_part()
        {
            return Err(invalid("SVG relationship count exceeds caller limit"));
        }
        state.current_edges.try_add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            relationship.target_mode(),
        )?;
        state.current_rels = next_rels;
        state.current_xml = next_xml;
        pending_parts.push(SvgPartChange {
            action: GraphAction::Add,
            part: Box::new(BlobPart::new_shared(
                part_uri.clone(),
                "image/svg+xml".to_owned(),
                Arc::clone(
                    intent
                        .payload
                        .as_ref()
                        .ok_or_else(|| invalid("attach_svg has no staged payload"))?,
                ),
            )),
        });
        *manifest = manifest_add_token(manifest, &part_uri, content_types_cap)?;
    } else {
        let owner = match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner,
            SvgOwnerState::None | SvgOwnerState::Opaque => return Ok(()),
            SvgOwnerState::Linked(_) => {
                return Err(invalid("linked SVG owners cannot be detached"));
            },
            SvgOwnerState::Ambiguous | SvgOwnerState::Refused => {
                return Err(invalid(
                    "selected picture has an ambiguous or refused SVG owner",
                ));
            },
        };
        let relationship_id = owner
            .embedded_relationship_id()
            .ok_or_else(|| invalid("embedded SVG owner has no relationship ID"))?;
        let relationship = state
            .current_edges
            .get(relationship_id)
            .cloned()
            .ok_or_else(|| invalid("embedded SVG relationship is missing"))?;
        let target = relationship.target_partname()?;
        validate_svg_target_candidate(workbook, &relationship, &target, pending_parts)?;
        let (after_xml, removed_range) = detach_xml(&state.current_xml, picture, owner)?;
        let remaining = scanned
            .relationship_references()
            .iter()
            .filter(|reference| {
                reference.id() == relationship_id
                    && !ranges_overlap(reference.range(), removed_range)
            })
            .count();
        let remove_edge = remaining == 0;
        if remove_edge {
            state.current_rels = state
                .current_rels
                .without_relationship(relationship_id, relationship_cap)?;
            state.current_edges.remove(relationship_id);
            if !cleanup_targets.iter().any(|candidate| candidate == &target) {
                cleanup_targets.push(target);
            }
        }
        state.current_xml = after_xml;
    }
    if state.current_xml.bytes().len() > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle drawing output exceeds {} bytes",
            output_cap
        )));
    }
    Ok(())
}

/// Plan the common attach-only batch in one source scan and one XML/
/// relationship allocation.  Detach and repeated-selector sequences retain
/// the sequential path below because their source ranges can be created or
/// removed by an earlier operation in the same transaction.
fn plan_composed_attach_batch(
    workbook: &Workbook,
    group: &[usize],
    intents: &[SvgLifecycleIntent],
    state: &mut DrawingPlanState,
    pending_parts: &mut Vec<SvgPartChange>,
    reserved_uris: &mut HashSet<PackURI>,
    manifest: &mut OwnedContentTypes,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let relationship_cap = relationship_output_limit(workbook);
    let content_types_cap = content_types_output_limit(workbook);
    let first_intent = intents
        .get(*group.first().ok_or_else(|| invalid("empty SVG lifecycle group"))?)
        .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
    let scanned = SourceDrawing::scan_with_limits(
        state.before_xml.bytes(),
        selector_drawing(first_intent.selector),
        limits,
    )?;
    let relationship_type = if scanned.relationship_dialect() == source::RelationshipDialect::Strict
    {
        rt::STRICT_IMAGE
    } else {
        rt::IMAGE
    };
    let mut additions: Vec<(String, String, String, TargetMode)> = Vec::new();
    additions
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle relationship batch", source))?;
    let mut fragments: Vec<Vec<u8>> = Vec::new();
    fragments
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle extension batch", source))?;
    let mut append_targets = Vec::new();
    append_targets
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle XML update batch", source))?;
    let mut new_uris = Vec::new();
    new_uris
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle media batch", source))?;

    for index in group {
        let intent = intents
            .get(*index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let payload = intent
            .payload
            .as_deref()
            .ok_or_else(|| invalid("attach_svg has no staged payload"))?;
        if payload.len() > output_cap {
            return Err(invalid(format!(
                "SVG lifecycle payload exceeds {} bytes",
                output_cap
            )));
        }
        let picture = scanned.picture(selector_picture(intent.selector))?;
        validate_raster_fallback_edges(workbook, &state.current_edges, picture)?;
        match picture.svg_owner() {
            SvgOwnerState::None | SvgOwnerState::Opaque => {},
            SvgOwnerState::Embedded(_) => {
                return Err(invalid(
                    "selected picture already has an embedded SVG owner",
                ));
            },
            SvgOwnerState::Linked(_) => {
                return Err(invalid(
                    "linked SVG owners cannot be replaced by attach_svg",
                ));
            },
            SvgOwnerState::Ambiguous | SvgOwnerState::Refused => {
                return Err(invalid(
                    "selected picture has an ambiguous or refused SVG owner",
                ));
            },
        }
        let relationship_id = allocate_relationship_id(&state.current_edges)?;
        let part_uri = allocate_part_uri_with_reservations(&workbook.inner.package, reserved_uris)?;
        reserved_uris.insert(part_uri.clone());
        let target_ref = part_uri.relative_ref(state.uri.base_uri());
        let next_edges_len = state
            .current_edges
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("SVG relationship count overflows"))?;
        let maximum_relationships = workbook
            .inner
            .package
            .read_limits()
            .max_relationships_per_part();
        if next_edges_len > maximum_relationships {
            return Err(invalid("SVG relationship count exceeds caller limit"));
        }
        state.current_edges.try_add_relationship(
            relationship_type.to_owned(),
            target_ref.clone(),
            relationship_id.clone(),
            TargetMode::Internal,
        )?;
        additions.push((
            relationship_type.to_owned(),
            target_ref,
            relationship_id.clone(),
            TargetMode::Internal,
        ));

        let blip = picture.blip_range();
        let extension_prefix = picture
            .ext_list_range()
            .map_or(blip.prefix(), ElementRange::prefix);
        let extension = generated_extension(blip.prefix(), &relationship_id, scanned.dialect())?;
        let (parent, parent_local, fragment) = if let Some(ext_list) = picture.ext_list_range() {
            (ext_list, "extLst", extension)
        } else {
            let list = generated_ext_list(extension_prefix, &extension, scanned.dialect())?;
            (blip, "blip", list)
        };
        let fragment_index = fragments.len();
        fragments.push(fragment);
        append_targets.push(XmlAppendTarget {
            range: parent.range().start..parent.start_end(),
            parent_empty: parent.is_empty(),
            parent_name_len: element_name_len(parent.prefix(), parent_local),
            fragment_index,
        });
        new_uris.push(part_uri.clone());
        pending_parts.push(SvgPartChange {
            action: GraphAction::Add,
            part: Box::new(BlobPart::new_shared(
                part_uri,
                "image/svg+xml".to_owned(),
                Arc::clone(
                    intent
                        .payload
                        .as_ref()
                        .ok_or_else(|| invalid("attach_svg has no staged payload"))?,
                ),
            )),
        });
    }

    let mut relationship_refs = Vec::new();
    relationship_refs
        .try_reserve(additions.len())
        .map_err(|source| allocation("SVG lifecycle relationship references", source))?;
    relationship_refs.extend(
        additions
            .iter()
            .map(|(reltype, target, id, mode)| (reltype.as_str(), target.as_str(), id.as_str(), *mode)),
    );
    let next_rels = state
        .current_rels
        .with_relationships(&relationship_refs, relationship_cap)?;

    append_targets.sort_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then(left.range.end.cmp(&right.range.end))
            .then(left.fragment_index.cmp(&right.fragment_index))
    });
    let mut output_size = state.before_xml.bytes().len();
    let mut expanded_parents: Vec<std::ops::Range<usize>> = Vec::new();
    expanded_parents
        .try_reserve(append_targets.len())
        .map_err(|source| allocation("SVG lifecycle empty XML parents", source))?;
    for target in &append_targets {
        output_size = output_size
            .checked_add(fragments[target.fragment_index].len())
            .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
        if target.parent_empty
            && !expanded_parents.contains(&target.range)
        {
            output_size = output_size
                .checked_add(target.parent_name_len.saturating_add(4))
                .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
            expanded_parents.push(target.range.clone());
        }
    }
    if output_size > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle drawing output exceeds {} bytes",
            output_cap
        )));
    }
    let mut updates = Vec::new();
    updates
        .try_reserve(append_targets.len())
        .map_err(|source| allocation("SVG lifecycle XML updates", source))?;
    for target in &append_targets {
        updates.push(OwnedElementUpdate {
            start_tag: target.range.clone(),
            edit: OwnedElementEdit::AppendChild(&fragments[target.fragment_index]),
        });
    }
    let next_xml = state.before_xml.update_elements(&updates, output_cap)?;
    state.current_rels = next_rels;
    state.current_xml = next_xml;

    *manifest = manifest_add_tokens(manifest, &new_uris, content_types_cap)?;
    Ok(())
}

struct XmlAppendTarget {
    range: std::ops::Range<usize>,
    parent_empty: bool,
    parent_name_len: usize,
    fragment_index: usize,
}

fn element_name_len(prefix: &[u8], local: &str) -> usize {
    local
        .len()
        .saturating_add(if prefix.is_empty() { 0 } else { prefix.len() + 1 })
}

fn finalize_svg_target_cleanup(
    workbook: &Workbook,
    states: &[DrawingPlanState],
    pending_parts: &mut Vec<SvgPartChange>,
    cleanup_targets: &[PackURI],
    manifest: &mut OwnedContentTypes,
    content_types_cap: usize,
) -> Result<()> {
    for target in cleanup_targets {
        if incoming_relationship_exists_with_states(workbook, target, states) {
            continue;
        }
        if let Some(index) = pending_parts.iter().position(|change| {
            change.action == GraphAction::Add && change.part.partname() == target
        }) {
            pending_parts.remove(index);
        } else {
            let part = workbook.inner.package.get_part(target)?;
            pending_parts.push(SvgPartChange {
                action: GraphAction::Remove,
                part: Box::new(BlobPart::new_shared(
                    target.clone(),
                    part.content_type().to_owned(),
                    part.blob_arc(),
                )),
            });
        }
        *manifest = manifest_remove_token(manifest, target, content_types_cap)?;
    }
    Ok(())
}

fn changed_relationship_id(
    before: &Relationships,
    after: &Relationships,
) -> Result<Option<String>> {
    let capacity = before
        .len()
        .checked_add(after.len())
        .ok_or_else(|| invalid("relationship ID count overflows"))?;
    let mut ids = Vec::new();
    ids.try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle relationship IDs", source))?;
    ids.extend(before.iter().map(|relationship| relationship.r_id().to_owned()));
    ids.extend(after.iter().map(|relationship| relationship.r_id().to_owned()));
    ids.sort_unstable();
    ids.dedup();
    Ok(ids.into_iter().find(|id| {
        !same_relationship_option(before.get(id), after.get(id))
    }))
}

fn same_relationship_option(before: Option<&Relationship>, after: Option<&Relationship>) -> bool {
    match (before, after) {
        (None, None) => true,
        (Some(before), Some(after)) => {
            before.r_id() == after.r_id()
                && before.reltype() == after.reltype()
                && before.target_ref() == after.target_ref()
                && before.target_mode() == after.target_mode()
        },
        _ => false,
    }
}

fn validate_raster_fallback_edges(
    workbook: &Workbook,
    relationships: &Relationships,
    picture: &source::PictureSource<'_>,
) -> Result<()> {
    let relationship = relationships
        .get(picture.raster_relationship_id())
        .ok_or_else(|| invalid("picture raster relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Err(invalid(
            "picture raster fallback relationship is unsupported",
        ));
    }
    let target = relationship.target_partname()?;
    if !target.as_str().starts_with("/xl/media/") {
        return Err(invalid("picture raster fallback is outside /xl/media/"));
    }
    let part = workbook.inner.package.get_part(&target)?;
    if part.content_type() != ct::PNG || !part.rels().is_empty() {
        return Err(invalid("picture raster fallback must be an inert PNG part"));
    }
    Ok(())
}

fn validate_svg_target_candidate(
    workbook: &Workbook,
    relationship: &Relationship,
    target: &PackURI,
    pending_parts: &[SvgPartChange],
) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
        || !target.as_str().starts_with("/xl/media/")
    {
        return Err(invalid("SVG relationship target is unsupported"));
    }
    if let Some(part) = pending_parts
        .iter()
        .find(|change| change.action == GraphAction::Add && change.part.partname() == target)
        .map(|change| change.part.as_ref())
    {
        if part.content_type() != "image/svg+xml" || !part.rels().is_empty() {
            return Err(invalid("SVG target must be an inert image/svg+xml part"));
        }
        return Ok(());
    }
    validate_svg_target(workbook, relationship, target)
}

fn incoming_relationship_exists_with_states(
    workbook: &Workbook,
    target: &PackURI,
    states: &[DrawingPlanState],
) -> bool {
    if relationships_target_uri(workbook.inner.package.rels(), target) {
        return true;
    }
    for part in workbook.inner.package.iter_parts() {
        let relationships = states
            .iter()
            .find(|state| state.uri == *part.partname())
            .map_or_else(|| part.rels(), |state| &state.current_edges);
        if relationships_target_uri(relationships, target) {
            return true;
        }
    }
    false
}

fn relationships_target_uri(relationships: &Relationships, target: &PackURI) -> bool {
    relationships.iter().any(|relationship| {
        relationship.target_mode() == TargetMode::Internal
            && relationship
                .target_partname()
                .is_ok_and(|candidate| candidate == *target)
    })
}

fn allocate_part_uri_with_reservations(
    package: &litchi_opc::OpcPackage,
    reserved: &HashSet<PackURI>,
) -> Result<PackURI> {
    for index in 0..MAX_ID_ATTEMPTS {
        let name = if index == 0 {
            "/xl/media/vector.svg".to_owned()
        } else {
            format!("/xl/media/vector{index}.svg")
        };
        let uri = PackURI::new(name).map_err(litchi_opc::OpcError::InvalidPackUri)?;
        if !reserved.contains(&uri) && package.validate_new_part_name(&uri).is_ok() {
            return Ok(uri);
        }
    }
    Err(invalid(format!(
        "generated SVG media Part candidates exceed {}",
        MAX_ID_ATTEMPTS
    )))
}

fn manifest_add_token(
    before: &OwnedContentTypes,
    uri: &PackURI,
    max_output_bytes: usize,
) -> Result<OwnedContentTypes> {
    if manifest_covers_svg(before.bytes(), uri)? {
        return Ok(before.clone());
    }
    before
        .with_part_overrides(&[(uri, "image/svg+xml")], max_output_bytes)
        .map_err(Into::into)
}

fn manifest_add_tokens(
    before: &OwnedContentTypes,
    uris: &[PackURI],
    max_output_bytes: usize,
) -> Result<OwnedContentTypes> {
    if uris.is_empty() || manifest_has_default_svg(before.bytes())? {
        return Ok(before.clone());
    }
    let mut overrides = Vec::new();
    overrides
        .try_reserve(uris.len())
        .map_err(|source| allocation("SVG lifecycle content-type overrides", source))?;
    for uri in uris {
        if !manifest_has_override(before.bytes(), uri)? {
            overrides.push((uri, "image/svg+xml"));
        }
    }
    if overrides.is_empty() {
        return Ok(before.clone());
    }
    before
        .with_part_overrides(&overrides, max_output_bytes)
        .map_err(Into::into)
}

fn manifest_remove_token(
    before: &OwnedContentTypes,
    uri: &PackURI,
    max_output_bytes: usize,
) -> Result<OwnedContentTypes> {
    if !manifest_has_override(before.bytes(), uri)? {
        return Ok(before.clone());
    }
    before
        .without_parts(std::slice::from_ref(uri), max_output_bytes)
        .map_err(Into::into)
}

fn resolve_drawing(
    workbook: &Workbook,
    position: usize,
    selector: PictureSelector,
) -> Result<(PackURI, Relationship)> {
    let sheet = workbook
        .inner
        .sheets
        .get(position)
        .ok_or_else(|| invalid("worksheet selector disappeared during SVG planning"))?;
    let worksheet = workbook.inner.package.get_part(&sheet.part_uri)?;
    let references = worksheet_drawing_references(workbook, worksheet.blob())?;
    let relationship_id = references
        .get(selector_drawing(selector))
        .ok_or_else(|| invalid("drawing ordinal is outside worksheet drawings"))?;
    let relationship = worksheet
        .rels()
        .get(relationship_id)
        .cloned()
        .ok_or_else(|| invalid("worksheet drawing reference has no package relationship"))?;
    if !matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING) {
        return Err(invalid(
            "worksheet drawing reference has a non-drawing relationship type",
        ));
    }
    if relationship.target_mode() != TargetMode::Internal {
        return Err(invalid("worksheet drawing relationship is external"));
    }
    let target = relationship.target_partname()?;
    let drawing = workbook.inner.package.get_part(&target)?;
    if drawing.content_type() != ct::OFC_DRAWING {
        return Err(invalid(format!(
            "worksheet drawing part has content type '{}', expected '{}'",
            drawing.content_type(),
            ct::OFC_DRAWING
        )));
    }
    Ok((target, relationship))
}

fn worksheet_drawing_references(workbook: &Workbook, source: &[u8]) -> Result<Vec<String>> {
    let limits = scan_limits(workbook);
    if source.len() > limits.max_xml_bytes {
        return Err(invalid(format!(
            "worksheet XML exceeds {} bytes",
            limits.max_xml_bytes
        )));
    }
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut mce_depth = 0usize;
    let mut mce_drawing_seen = false;
    let mut nodes = 0usize;
    let mut references = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML event count overflow"))?;
        if nodes > limits.max_nodes {
            return Err(invalid("worksheet XML exceeds caller event limit"));
        }
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen
                        || !namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                            && !namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("worksheet XML has an invalid root"));
                    }
                    root_seen = true;
                }
                if namespace_matches(&namespace, MCE_NAMESPACE)
                    && element.local_name().as_ref() == b"AlternateContent"
                {
                    mce_depth = mce_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("worksheet MCE nesting overflow"))?;
                }
                if mce_depth > 0
                    && (namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                        || namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE))
                    && element.local_name().as_ref() == b"drawing"
                {
                    mce_drawing_seen = true;
                }
                if depth == 1
                    && mce_depth == 0
                    && (namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                        || namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE))
                    && element.local_name().as_ref() == b"drawing"
                {
                    references.try_reserve(1).map_err(|source| {
                        allocation("worksheet drawing references", source)
                    })?;
                    references.push(worksheet_drawing_relationship_id(
                        &reader,
                        &element,
                    )?);
                    if references.len() > limits.max_relationship_references {
                        return Err(invalid(
                            "worksheet drawing references exceed caller limit",
                        ));
                    }
                }
                let next_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if next_depth > limits.max_depth {
                    return Err(invalid("worksheet XML exceeds caller depth limit"));
                }
                depth = next_depth;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if root_seen
                        || (!namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                            && !namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE))
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("worksheet XML has an invalid root"));
                    }
                    root_seen = true;
                    buffer.clear();
                    continue;
                }
                if mce_depth > 0
                    && (namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                        || namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE))
                    && element.local_name().as_ref() == b"drawing"
                {
                    mce_drawing_seen = true;
                }
                if depth == 1
                    && mce_depth == 0
                    && (namespace_matches(&namespace, SPREADSHEETML_NAMESPACE)
                        || namespace_matches(&namespace, STRICT_SPREADSHEETML_NAMESPACE))
                    && element.local_name().as_ref() == b"drawing"
                {
                    references.try_reserve(1).map_err(|source| {
                        allocation("worksheet drawing references", source)
                    })?;
                    references.push(worksheet_drawing_relationship_id(
                        &reader,
                        &element,
                    )?);
                    if references.len() > limits.max_relationship_references {
                        return Err(invalid(
                            "worksheet drawing references exceed caller limit",
                        ));
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("worksheet XML has an unexpected end element"));
                }
                if namespace_matches(&namespace, MCE_NAMESPACE)
                    && element.local_name().as_ref() == b"AlternateContent"
                {
                    mce_depth = mce_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("worksheet MCE nesting underflow"))?;
                }
                depth -= 1;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 || mce_depth != 0 {
        return Err(invalid("worksheet XML is incomplete"));
    }
    if mce_drawing_seen {
        return Err(invalid(
            "worksheet drawing reference under MCE is refused",
        ));
    }
    Ok(references)
}

fn worksheet_drawing_relationship_id(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<String> {
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        if (namespace_matches(&namespace, RELATIONSHIPS_NAMESPACE)
            || namespace_matches(&namespace, STRICT_RELATIONSHIPS_NAMESPACE))
            && attribute.key.local_name().as_ref() == b"id"
        {
            if attribute.value.len() > source::MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4) {
                return Err(invalid("worksheet relationship ID lexical value is too large"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid(error.to_string()))?;
            if value.len() > source::MAX_RELATIONSHIP_ID_BYTES || value.is_empty() {
                return Err(invalid("worksheet relationship ID value is invalid"));
            }
            if relationship_id.is_some() {
                return Err(invalid(
                    "worksheet drawing reference has duplicate relationship IDs",
                ));
            }
            relationship_id = Some(value.into_owned());
        }
    }
    relationship_id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("worksheet drawing reference has no relationship ID"))
}

fn namespace_matches(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn validate_raster_fallback_id(
    workbook: &Workbook,
    drawing_part: &dyn Part,
    relationship_id: &str,
) -> Result<()> {
    let relationship = drawing_part
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid("picture raster relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Err(invalid(
            "picture raster fallback relationship is unsupported",
        ));
    }
    let target = relationship.target_partname()?;
    if !target.as_str().starts_with("/xl/media/") {
        return Err(invalid("picture raster fallback is outside /xl/media/"));
    }
    let part = workbook.inner.package.get_part(&target)?;
    if part.content_type() != ct::PNG || !part.rels().is_empty() {
        return Err(invalid("picture raster fallback must be an inert PNG part"));
    }
    Ok(())
}

fn validate_svg_target(
    workbook: &Workbook,
    relationship: &Relationship,
    target: &PackURI,
) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
        || !target.as_str().starts_with("/xl/media/")
    {
        return Err(invalid("SVG relationship target is unsupported"));
    }
    let part = workbook.inner.package.get_part(target)?;
    if part.content_type() != "image/svg+xml" || !part.rels().is_empty() {
        return Err(invalid("SVG target must be an inert image/svg+xml part"));
    }
    Ok(())
}

fn attach_xml(
    before: &OwnedXmlPart,
    picture: &source::PictureSource<'_>,
    relationship_id: &str,
    drawing_dialect: source::DrawingDialect,
    max_output_bytes: usize,
) -> Result<OwnedXmlPart> {
    let blip = picture.blip_range();
    let drawing_prefix = blip.prefix();
    let extension_prefix = picture
        .ext_list_range()
        .map_or(drawing_prefix, ElementRange::prefix);
    let extension = generated_extension(drawing_prefix, relationship_id, drawing_dialect)?;
    let replacement = if let Some(ext_list) = picture.ext_list_range() {
        let output_len = before
            .bytes()
            .len()
            .checked_add(extension.len())
            .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
        if output_len > max_output_bytes {
            return Err(invalid("SVG lifecycle drawing output exceeds caller limit"));
        }
        before.append_element(ext_list.range().start..ext_list.start_end(), &extension)?
    } else {
        let name_len = if extension_prefix.is_empty() {
            "extLst".len()
        } else {
            extension_prefix.len().saturating_add(7)
        };
        let binding_len = if extension_prefix.is_empty() {
            0
        } else {
            10usize
                .saturating_add(extension_prefix.len())
                .saturating_add(drawing_namespace(drawing_dialect).len())
        };
        let growth = extension
            .len()
            .saturating_add(name_len.saturating_mul(2).saturating_add(5))
            .saturating_add(binding_len);
        let output_len = before
            .bytes()
            .len()
            .checked_add(growth)
            .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
        if output_len > max_output_bytes {
            return Err(invalid("SVG lifecycle drawing output exceeds caller limit"));
        }
        let list = generated_ext_list(extension_prefix, &extension, drawing_dialect)?;
        before.append_element(blip.range().start..blip.start_end(), &list)?
    };
    Ok(replacement)
}

fn detach_xml(
    before: &OwnedXmlPart,
    picture: &source::PictureSource<'_>,
    owner: &source::SvgOwner<'_>,
) -> Result<(OwnedXmlPart, source::ByteRange)> {
    if let Some(ext_list) = picture.ext_list_range() {
        if ext_list_contains_only_owner(before.bytes(), ext_list, owner)? {
            return Ok((
                before.remove_element(ext_list.range().start..ext_list.start_end())?,
                ext_list.range(),
            ));
        }
    }
    let extension = owner.extension_range();
    let opening = opening_tag_range(before.bytes(), extension)?;
    Ok((before.remove_element(opening)?, extension))
}

fn ranges_overlap(left: source::ByteRange, right: source::ByteRange) -> bool {
    left.start < right.end && right.start < left.end
}

fn opening_tag_range(source: &[u8], element: source::ByteRange) -> Result<std::ops::Range<usize>> {
    let bytes = element.slice(source)?;
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    let start = reader.buffer_position() as usize;
    let event = reader
        .read_event_into(&mut buffer)
        .map_err(|error| invalid(error.to_string()))?;
    let end = reader.buffer_position() as usize;
    if !matches!(event, Event::Start(_) | Event::Empty(_)) || end <= start {
        return Err(invalid("SVG owner extension has no complete opening tag"));
    }
    Ok(element.start + start..element.start + end)
}

fn ext_list_contains_only_owner(
    source: &[u8],
    ext_list: &ElementRange,
    owner: &source::SvgOwner<'_>,
) -> Result<bool> {
    let bytes = ext_list.range().slice(source)?;
    let selected = owner.extension_range();
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut child_count = 0usize;
    let mut child_start = None;
    let mut child_end = None;
    let mut buffer = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if depth == 0 && element.attributes().next().is_some() {
                    return Ok(false);
                }
                if depth == 1 {
                    if child_count != 0 {
                        return Ok(false);
                    }
                    let absolute = ext_list.range().start.saturating_add(start);
                    if absolute != selected.start {
                        return Ok(false);
                    }
                    child_start = Some(absolute);
                    child_count = child_count.saturating_add(1);
                }
                depth = depth.saturating_add(1);
            },
            Event::Empty(element) => {
                if depth == 0 && element.attributes().next().is_some() {
                    return Ok(false);
                }
                if depth == 1 {
                    let absolute = ext_list.range().start.saturating_add(start)
                        ..ext_list.range().start.saturating_add(end);
                    if child_count != 0 || absolute != (selected.start..selected.end) {
                        return Ok(false);
                    }
                    child_start = Some(absolute.start);
                    child_end = Some(absolute.end);
                    child_count = child_count.saturating_add(1);
                }
            },
            Event::End(_) => {
                if depth == 2 {
                    child_end = Some(ext_list.range().start.saturating_add(end));
                }
                depth = depth.saturating_sub(1);
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) if depth == 1 => {
                return Ok(false);
            },
            Event::Text(text)
                if depth == 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Ok(false);
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(child_count == 1 && child_start == Some(selected.start) && child_end == Some(selected.end))
}

fn generated_extension(
    prefix: &[u8],
    relationship_id: &str,
    drawing_dialect: source::DrawingDialect,
) -> Result<Vec<u8>> {
    let qname = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{}:{local}", String::from_utf8_lossy(prefix))
        }
    };
    let ext = qname("ext");
    let drawing_namespace = drawing_namespace(drawing_dialect);
    let drawing_binding = if prefix.is_empty() {
        String::new()
    } else {
        format!(
            " xmlns:{}=\"{}\"",
            String::from_utf8_lossy(prefix),
            drawing_namespace
        )
    };
    let value = format!(
        "<{ext}{drawing_binding} uri=\"{SVG_EXTENSION_URI}\"><asvg:svgBlip xmlns:asvg=\"{SVG_NAMESPACE}\" xmlns:r=\"{}\" r:embed=\"{}\"/></{ext}>",
        source::RelationshipDialect::authored_svg_attribute_namespace(),
        relationship_id,
    );
    Ok(value.into_bytes())
}

const fn drawing_namespace(dialect: source::DrawingDialect) -> &'static str {
    match dialect {
        source::DrawingDialect::Transitional => DRAWINGML_NAMESPACE,
        source::DrawingDialect::Strict => STRICT_DRAWINGML_NAMESPACE,
    }
}

fn generated_ext_list(
    prefix: &[u8],
    extension: &[u8],
    drawing_dialect: source::DrawingDialect,
) -> Result<Vec<u8>> {
    let name = if prefix.is_empty() {
        "extLst".to_owned()
    } else {
        format!("{}:extLst", String::from_utf8_lossy(prefix))
    };
    let drawing_binding = if prefix.is_empty() {
        String::new()
    } else {
        format!(
            " xmlns:{}=\"{}\"",
            String::from_utf8_lossy(prefix),
            drawing_namespace(drawing_dialect)
        )
    };
    let mut output = Vec::new();
    output
        .try_reserve(
            name.len()
                .saturating_mul(2)
                .saturating_add(extension.len())
                .saturating_add(drawing_binding.len())
                .saturating_add(5),
        )
        .map_err(|source| allocation("SVG lifecycle extension list", source))?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(drawing_binding.as_bytes());
    output.extend_from_slice(b">");
    output.extend_from_slice(extension);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
    Ok(output)
}

fn allocate_relationship_id(relationships: &Relationships) -> Result<String> {
    for index in 0..MAX_ID_ATTEMPTS {
        let candidate = if index == 0 {
            "rIdSvg".to_owned()
        } else {
            format!("rIdSvg{index}")
        };
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid(format!(
        "generated SVG relationship ID candidates exceed {}",
        MAX_ID_ATTEMPTS
    )))
}

fn manifest_covers_svg(bytes: &[u8], uri: &PackURI) -> Result<bool> {
    manifest_has_override(bytes, uri).or_else(|_| manifest_has_default_svg(bytes))
}

fn manifest_has_override(bytes: &[u8], uri: &PackURI) -> Result<bool> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?
        {
            Event::Empty(element) | Event::Start(element)
                if element.local_name().as_ref() == b"Override" =>
            {
                let mut part = None;
                let mut content = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
                    match attribute.key.local_name().as_ref() {
                        b"PartName" => {
                            part = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| invalid(error.to_string()))?
                                    .into_owned(),
                            );
                        },
                        b"ContentType" => {
                            content = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| invalid(error.to_string()))?
                                    .into_owned(),
                            );
                        },
                        _ => {},
                    }
                }
                if part.as_deref() == Some(uri.as_str())
                    && content.as_deref() == Some("image/svg+xml")
                {
                    return Ok(true);
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
        buffer.clear();
    }
}

fn manifest_has_default_svg(bytes: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?
        {
            Event::Empty(element) | Event::Start(element)
                if element.local_name().as_ref() == b"Default" =>
            {
                let mut extension = None;
                let mut content = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
                    match attribute.key.local_name().as_ref() {
                        b"Extension" => {
                            extension = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| invalid(error.to_string()))?
                                    .into_owned(),
                            );
                        },
                        b"ContentType" => {
                            content = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| invalid(error.to_string()))?
                                    .into_owned(),
                            );
                        },
                        _ => {},
                    }
                }
                if extension
                    .as_deref()
                    .is_some_and(|value| value.eq_ignore_ascii_case("svg"))
                    && content.as_deref() == Some("image/svg+xml")
                {
                    return Ok(true);
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
        buffer.clear();
    }
}

const fn selector_drawing(selector: PictureSelector) -> usize {
    selector.drawing
}

const fn selector_picture(selector: PictureSelector) -> usize {
    selector.picture
}
