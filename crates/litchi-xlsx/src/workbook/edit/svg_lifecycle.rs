//! Ordinary worksheet SVG attach/detach planning.
//!
//! The source drawing scanner owns semantic picture discovery.  This module
//! owns the mutation and the complete OPC dependency closure: one source XML
//! splice, one source relationship token, an optional SVG media part, and the
//! content-types delta required by that part.

use std::borrow::Cow;
use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::package::{ContentTypeEdit, RelationshipEdit};
use litchi_opc::{
    BlobPart, OwnedContentTypes, OwnedElementEdit, OwnedElementUpdate, OwnedRelationships,
    OwnedXmlPart, PackURI, Part, Relationship, Relationships, TargetMode,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::model::{
    ContentTypesChange, GraphAction, PackageChange, PartChange, RelationshipChange, SvgFinalGuard,
    SvgFinalMedia, SvgFinalState, SvgPartChange, SvgReadGuard,
};
use super::semantic::SvgLifecycleIntent;
use super::svg::PictureSelector;
use super::{Workbook, allocation, invalid};
use crate::drawing::source::{self, ElementRange, SourceDrawing, SvgOwnerState};
use crate::error::{Error, Result};

mod relationship_ids;
mod topology;
use relationship_ids::RelationshipIdAllocator;

const MAX_DRAWING_BYTES: usize = 32 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 32 * 1024 * 1024;
const MAX_ID_ATTEMPTS: usize = 100_000;
const MAX_GENERATED_RELATIONSHIP_ID: &str = "rIdSvg18446744073709551615";
pub(super) const MAX_SVG_LIFECYCLE_INTENTS: usize = 65_536;
const SVG_EXTENSION_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SVG_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

fn output_limit(workbook: &Workbook) -> usize {
    let caller = workbook.inner.package.read_limits().max_part_bytes();
    usize::try_from(caller)
        .unwrap_or(usize::MAX)
        .min(MAX_OUTPUT_BYTES)
}

fn content_types_output_limit(workbook: &Workbook) -> usize {
    let caller = workbook.inner.package.read_limits();
    output_limit(workbook).min(caller.max_content_types_bytes())
}

pub(super) fn payload_limit(workbook: &Workbook) -> usize {
    output_limit(workbook)
}

/// Census the aggregate materialized-part bytes of the transaction base.
///
/// The census is charged against inflated bytes, so it decodes every payload
/// (ADR 0030). A transaction therefore takes it only once it stages an SVG
/// lifecycle intent; an ordinary edit never pays for it.
pub(super) fn materialized_part_bytes(workbook: &Workbook) -> Result<usize> {
    workbook
        .inner
        .package
        .try_iter_parts()
        .try_fold(0usize, |total, part| {
            total
                .checked_add(part?.blob().len())
                .ok_or_else(|| invalid("SVG lifecycle package byte count overflows"))
        })
}

/// Check the staged SVG payload pool before a caller's borrowed bytes become
/// owned transaction state.
pub(super) fn check_payload_budget(
    workbook: &Workbook,
    already_staged: usize,
    incoming: usize,
) -> Result<()> {
    // This is an independent bound on the temporary transaction-owned payload
    // pool, rather than a second additive package-byte estimate.  The
    // package-wide candidate budget is checked later from the coalesced final
    // state, so an attach which is subsequently cancelled can still be
    // staged under an exact original package budget.
    let prospective = already_staged
        .checked_add(incoming)
        .ok_or_else(|| invalid("SVG lifecycle staged package byte count overflows"))?;
    let maximum = usize::try_from(workbook.inner.package.read_limits().max_total_part_bytes())
        .unwrap_or(usize::MAX);
    if prospective > maximum {
        return Err(invalid(format!(
            "SVG lifecycle staged package bytes exceed {}",
            maximum
        )));
    }
    Ok(())
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
        max_content_parts: defaults
            .max_content_parts
            .min(caller.max_relationships_per_part()),
        max_relationship_references: defaults
            .max_relationship_references
            .min(caller.max_relationships_per_part()),
        max_fragment_bytes: defaults.max_fragment_bytes.min(part_bytes),
    }
}

struct DrawingGrowth {
    uri: PackURI,
    source_len: usize,
    picture_states: Vec<PictureGrowth>,
    relationship_type: Box<str>,
    base_relationships: usize,
    relationship_reference_counts: HashMap<String, usize>,
    relationship_targets: HashMap<String, PackURI>,
    removed_relationship_ids: Vec<String>,
    planned_add_bytes: usize,
    planned_remove_bytes: usize,
    planned_relationship_additions: usize,
    planned_relationship_removals: usize,
}

#[derive(Default)]
struct RelationshipOwnerDelta {
    additions: usize,
    removals: usize,
}

struct PictureGrowth {
    initial: ProjectedSvgOwner,
    current: ProjectedSvgOwner,
    embedded_relationship_id: Option<Box<str>>,
    embedded_reference_count: usize,
    attach_growth: usize,
    attach_overhead: usize,
    blip_prefix_len: usize,
    dialect: source::DrawingDialect,
    detach_growth: usize,
    owner_extension_len: Option<usize>,
    saw_attach: bool,
    saw_detach: bool,
    last_attach_index: Option<usize>,
}

struct SvgDrawingBudget {
    added_bytes: usize,
    removed_bytes: usize,
    effective_attach_count: usize,
    effective_payload_bytes: usize,
    payload_lengths: Vec<usize>,
    removed_edges: Vec<RemovedSvgEdge>,
    relationship_additions: Vec<PlannedSvgRelationship>,
}

struct RemovedSvgEdge {
    owner: PackURI,
    relationship_id: String,
    target: PackURI,
}

struct PlannedSvgRelationship {
    owner: PackURI,
    relationship_id: String,
    relationship_type: Box<str>,
}

#[derive(Debug, Default)]
struct CleanupTargets {
    values: Vec<PackURI>,
    keys: HashSet<String>,
}

impl CleanupTargets {
    fn with_capacity(capacity: usize) -> Result<Self> {
        let mut values = Vec::new();
        values
            .try_reserve(capacity)
            .map_err(|source| allocation("SVG lifecycle media cleanup targets", source))?;
        let mut keys = HashSet::new();
        keys.try_reserve(capacity)
            .map_err(|source| allocation("SVG lifecycle media cleanup target index", source))?;
        Ok(Self { values, keys })
    }

    fn len(&self) -> usize {
        self.values.len()
    }

    fn push(&mut self, target: PackURI) -> Result<()> {
        let mut key = String::new();
        key.try_reserve(target.as_str().len())
            .map_err(|source| allocation("SVG lifecycle cleanup target key", source))?;
        for byte in target.as_str().bytes() {
            key.push(char::from(byte.to_ascii_lowercase()));
        }
        if self.keys.contains(&key) {
            return Ok(());
        }
        self.keys
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle media cleanup target index", source))?;
        self.values
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle media cleanup targets", source))?;
        self.keys.insert(key);
        self.values.push(target);
        Ok(())
    }
}

fn project_svg_drawing_budget(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
) -> Result<SvgDrawingBudget> {
    let relationship_owner_deltas =
        index_relationship_owner_deltas(workbook, ordinary_graph, relationship_changes)?;
    let mut drawings = HashMap::<String, DrawingGrowth>::new();
    drawings
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing growth index", source))?;
    let mut drawing_order = Vec::new();
    drawing_order
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing growth order", source))?;
    let mut drawing_lookup = HashMap::<(usize, usize), PackURI>::new();
    drawing_lookup
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing lookup", source))?;
    for (intent_index, intent) in intents.iter().enumerate() {
        let drawing_key = (intent.position, selector_drawing(intent.selector));
        let drawing_uri = if let Some(uri) = drawing_lookup.get(&drawing_key) {
            uri.clone()
        } else {
            let (uri, _) = resolve_drawing(workbook, intent.position, intent.selector)?;
            drawing_lookup.insert(drawing_key, uri.clone());
            uri
        };
        let key = drawing_uri.as_str().to_ascii_lowercase();
        if !drawings.contains_key(&key) {
            let drawing = workbook.inner.package.source_xml_part(&drawing_uri)?;
            let scanned = SourceDrawing::scan_with_limits(
                drawing.bytes(),
                selector_drawing(intent.selector),
                scan_limits(workbook),
            )?;
            let mut picture_states = Vec::new();
            picture_states
                .try_reserve(scanned.pictures().len())
                .map_err(|source| allocation("SVG lifecycle drawing growth", source))?;
            let embedded_reference_counts_by_picture = embedded_reference_counts(&scanned)?;
            let mut relationship_reference_counts = HashMap::<String, usize>::new();
            relationship_reference_counts
                .try_reserve(scanned.relationship_references().len())
                .map_err(|source| {
                    allocation("SVG lifecycle relationship reference index", source)
                })?;
            for reference in scanned.relationship_references() {
                let count = relationship_reference_counts
                    .entry(reference.id().to_owned())
                    .or_insert(0);
                *count = (*count).checked_add(1).ok_or_else(|| {
                    invalid("SVG lifecycle relationship reference count overflows")
                })?;
            }
            let drawing_part = workbook.inner.package.get_part(&drawing_uri)?;
            let mut relationship_targets = HashMap::new();
            relationship_targets
                .try_reserve(drawing_part.rels().len())
                .map_err(|source| allocation("SVG lifecycle relationship target index", source))?;
            for relationship in drawing_part.rels().iter() {
                if relationship.target_mode() == TargetMode::Internal {
                    relationship_targets.insert(
                        relationship.r_id().to_owned(),
                        relationship.target_partname()?,
                    );
                }
            }
            for (picture_index, picture) in scanned.pictures().iter().enumerate() {
                let (initial, embedded_relationship_id, embedded_reference_count) = match picture
                    .svg_owner()
                {
                    SvgOwnerState::None => (ProjectedSvgOwner::None, None, 0),
                    SvgOwnerState::Opaque => (ProjectedSvgOwner::Opaque, None, 0),
                    SvgOwnerState::Embedded(owner) => {
                        let relationship_id = owner
                            .embedded_relationship_id()
                            .ok_or_else(|| invalid("embedded SVG owner has no relationship ID"))?;
                        let reference_count = *embedded_reference_counts_by_picture
                            .get(picture_index)
                            .ok_or_else(|| invalid("SVG lifecycle embedded picture disappeared"))?;
                        (
                            ProjectedSvgOwner::Embedded,
                            Some(relationship_id.to_owned().into_boxed_str()),
                            reference_count,
                        )
                    },
                    SvgOwnerState::Linked(_) => (ProjectedSvgOwner::Linked, None, 0),
                    SvgOwnerState::Ambiguous => (ProjectedSvgOwner::Ambiguous, None, 0),
                    SvgOwnerState::Refused => (ProjectedSvgOwner::Refused, None, 0),
                };
                let extension_len = generated_extension_size(
                    picture.blip_range().prefix(),
                    MAX_GENERATED_RELATIONSHIP_ID,
                    scanned.dialect(),
                );
                let attach_growth = if let Some(ext_list) = picture.ext_list_range() {
                    extension_len.saturating_add(if ext_list.is_empty() {
                        element_name_len(ext_list.prefix(), "extLst").saturating_add(2)
                    } else {
                        0
                    })
                } else {
                    generated_ext_list_size(
                        picture.blip_range().prefix(),
                        extension_len,
                        scanned.dialect(),
                    )
                    .saturating_add(if picture.blip_range().is_empty() {
                        element_name_len(picture.blip_range().prefix(), "blip").saturating_add(2)
                    } else {
                        0
                    })
                };
                let attach_overhead = attach_growth
                    .checked_sub(extension_len)
                    .ok_or_else(|| invalid("SVG lifecycle drawing growth underflows"))?;
                let (detach_growth, owner_extension_len) =
                    if let SvgOwnerState::Embedded(owner) = picture.svg_owner() {
                        let (_, removed) = detach_ranges(drawing.bytes(), picture, owner)?;
                        (removed.len()?, Some(owner.extension_range().len()?))
                    } else {
                        (0, None)
                    };
                picture_states.push(PictureGrowth {
                    initial,
                    current: initial,
                    embedded_relationship_id,
                    embedded_reference_count,
                    attach_growth,
                    attach_overhead,
                    blip_prefix_len: picture.blip_range().prefix().len(),
                    dialect: scanned.dialect(),
                    detach_growth,
                    owner_extension_len,
                    saw_attach: false,
                    saw_detach: false,
                    last_attach_index: None,
                });
            }
            let base_relationships = drawing_part.rels().len();
            drawings.insert(
                key.clone(),
                DrawingGrowth {
                    uri: drawing_uri.clone(),
                    source_len: drawing.bytes().len(),
                    picture_states,
                    relationship_type: if scanned.relationship_dialect()
                        == source::RelationshipDialect::Strict
                    {
                        rt::STRICT_IMAGE.to_owned().into_boxed_str()
                    } else {
                        rt::IMAGE.to_owned().into_boxed_str()
                    },
                    base_relationships,
                    relationship_reference_counts,
                    relationship_targets,
                    removed_relationship_ids: Vec::new(),
                    planned_add_bytes: 0,
                    planned_remove_bytes: 0,
                    planned_relationship_additions: 0,
                    planned_relationship_removals: 0,
                },
            );
            drawing_order.push(key.clone());
        }
        let drawing = drawings
            .get_mut(&key)
            .ok_or_else(|| invalid("SVG lifecycle drawing growth index disappeared"))?;
        let picture = drawing
            .picture_states
            .get_mut(selector_picture(intent.selector))
            .ok_or_else(|| invalid("picture ordinal is outside worksheet drawing pictures"))?;
        if intent.attach {
            picture.saw_attach = true;
            picture.last_attach_index = Some(intent_index);
            match picture.current {
                ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => {
                    picture.current = ProjectedSvgOwner::Embedded;
                },
                ProjectedSvgOwner::Embedded => {
                    return Err(invalid(
                        "selected picture already has an embedded SVG owner",
                    ));
                },
                ProjectedSvgOwner::Linked => {
                    return Err(invalid(
                        "linked SVG owners cannot be replaced by attach_svg",
                    ));
                },
                ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => {
                    return Err(invalid(
                        "selected picture has an ambiguous or refused SVG owner",
                    ));
                },
            }
        } else {
            picture.saw_detach = true;
            match picture.current {
                ProjectedSvgOwner::Embedded => {
                    picture.current = if picture.initial == ProjectedSvgOwner::Opaque {
                        ProjectedSvgOwner::Opaque
                    } else {
                        ProjectedSvgOwner::None
                    };
                },
                ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => {},
                ProjectedSvgOwner::Linked => {
                    return Err(invalid("linked SVG owners cannot be detached"));
                },
                ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => {
                    return Err(invalid(
                        "selected picture has an ambiguous or refused SVG owner",
                    ));
                },
            }
        }
    }

    let mut budget = SvgDrawingBudget {
        added_bytes: 0,
        removed_bytes: 0,
        effective_attach_count: 0,
        effective_payload_bytes: 0,
        payload_lengths: Vec::new(),
        removed_edges: Vec::new(),
        relationship_additions: Vec::new(),
    };
    budget
        .payload_lengths
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle attachment payload lengths", source))?;
    budget
        .relationship_additions
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle relationship additions", source))?;
    for drawing_key in drawing_order {
        let drawing = drawings
            .get_mut(&drawing_key)
            .ok_or_else(|| invalid("SVG lifecycle drawing growth order disappeared"))?;
        let mut removed_reference_counts = HashMap::<&str, usize>::new();
        removed_reference_counts
            .try_reserve(drawing.picture_states.len())
            .map_err(|source| {
                allocation("SVG lifecycle drawing relationship references", source)
            })?;
        for picture in &drawing.picture_states {
            match (picture.initial, picture.current) {
                (ProjectedSvgOwner::Embedded, ProjectedSvgOwner::Embedded)
                    if picture.saw_attach && picture.saw_detach =>
                {
                    let old_len = picture
                        .owner_extension_len
                        .ok_or_else(|| invalid("SVG lifecycle owner range disappeared"))?;
                    drawing.planned_remove_bytes = drawing
                        .planned_remove_bytes
                        .checked_add(old_len)
                        .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
                    drawing.planned_add_bytes = drawing
                        .planned_add_bytes
                        .checked_add(picture.attach_growth)
                        .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
                    drawing.planned_relationship_additions = drawing
                        .planned_relationship_additions
                        .checked_add(1)
                        .ok_or_else(|| {
                            invalid("SVG lifecycle drawing relationship count overflows")
                        })?;
                    if let Some(relationship_id) = picture.embedded_relationship_id.as_deref() {
                        let count = removed_reference_counts.entry(relationship_id).or_insert(0);
                        *count = (*count)
                            .checked_add(picture.embedded_reference_count)
                            .ok_or_else(|| {
                                invalid("SVG lifecycle relationship reference count overflows")
                            })?;
                    }
                    budget.effective_attach_count = budget
                        .effective_attach_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG lifecycle attachment count overflows"))?;
                    let payload_index = picture
                        .last_attach_index
                        .ok_or_else(|| invalid("SVG lifecycle replacement payload disappeared"))?;
                    let payload = intents
                        .get(payload_index)
                        .and_then(|intent| intent.payload.as_deref())
                        .ok_or_else(|| invalid("SVG lifecycle replacement payload disappeared"))?;
                    budget.effective_payload_bytes = budget
                        .effective_payload_bytes
                        .checked_add(payload.len())
                        .ok_or_else(|| invalid("SVG lifecycle payload bytes overflow"))?;
                },
                (ProjectedSvgOwner::Embedded, ProjectedSvgOwner::None) => {
                    drawing.planned_remove_bytes = drawing
                        .planned_remove_bytes
                        .checked_add(picture.detach_growth)
                        .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
                    if let Some(relationship_id) = picture.embedded_relationship_id.as_deref() {
                        let count = removed_reference_counts.entry(relationship_id).or_insert(0);
                        *count = (*count)
                            .checked_add(picture.embedded_reference_count)
                            .ok_or_else(|| {
                                invalid("SVG lifecycle relationship reference count overflows")
                            })?;
                    }
                },
                (
                    ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque,
                    ProjectedSvgOwner::Embedded,
                ) => {
                    drawing.planned_add_bytes = drawing
                        .planned_add_bytes
                        .checked_add(picture.attach_growth)
                        .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
                    drawing.planned_relationship_additions = drawing
                        .planned_relationship_additions
                        .checked_add(1)
                        .ok_or_else(|| {
                            invalid("SVG lifecycle drawing relationship count overflows")
                        })?;
                    budget.effective_attach_count = budget
                        .effective_attach_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG lifecycle attachment count overflows"))?;
                    let payload_index = picture
                        .last_attach_index
                        .ok_or_else(|| invalid("SVG lifecycle attach payload disappeared"))?;
                    let payload = intents
                        .get(payload_index)
                        .and_then(|intent| intent.payload.as_deref())
                        .ok_or_else(|| invalid("SVG lifecycle attach payload disappeared"))?;
                    budget.effective_payload_bytes = budget
                        .effective_payload_bytes
                        .checked_add(payload.len())
                        .ok_or_else(|| invalid("SVG lifecycle payload bytes overflow"))?;
                },
                _ => {},
            }
        }
        let mut removed_relationship_ids = Vec::new();
        removed_relationship_ids
            .try_reserve(removed_reference_counts.len())
            .map_err(|source| allocation("SVG lifecycle drawing relationship IDs", source))?;
        for (relationship_id, removed_count) in removed_reference_counts {
            if drawing
                .relationship_reference_counts
                .get(relationship_id)
                .copied()
                == Some(removed_count)
            {
                removed_relationship_ids.push(relationship_id.to_owned());
            }
        }
        drawing.planned_relationship_removals = removed_relationship_ids.len();
        drawing.removed_relationship_ids = removed_relationship_ids;
        let mut final_attachment_indices = Vec::new();
        final_attachment_indices
            .try_reserve(drawing.picture_states.len())
            .map_err(|source| allocation("SVG lifecycle relationship additions", source))?;
        for (picture_index, picture) in drawing.picture_states.iter().enumerate() {
            let final_attach = match (picture.initial, picture.current) {
                (ProjectedSvgOwner::Embedded, ProjectedSvgOwner::Embedded) => {
                    picture.saw_attach && picture.saw_detach
                },
                (
                    ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque,
                    ProjectedSvgOwner::Embedded,
                ) => true,
                _ => false,
            };
            if final_attach {
                final_attachment_indices.push(picture_index);
            }
        }
        final_attachment_indices.sort_by_key(|index| {
            drawing.picture_states[*index]
                .last_attach_index
                .unwrap_or(usize::MAX)
        });
        let drawing_part = workbook.inner.package.get_part(&drawing.uri)?;
        let mut relationship_ids = RelationshipIdAllocator::from_source(drawing_part.rels())?;
        let relationship_addition_start = budget.relationship_additions.len();
        for picture_index in &final_attachment_indices {
            let picture = drawing
                .picture_states
                .get(*picture_index)
                .ok_or_else(|| invalid("SVG lifecycle relationship picture disappeared"))?;
            let reuse_existing = picture.initial == ProjectedSvgOwner::Embedded
                && picture.current == ProjectedSvgOwner::Embedded
                && picture.saw_attach
                && picture.saw_detach
                && picture
                    .embedded_relationship_id
                    .as_deref()
                    .is_some_and(|id| {
                        drawing.relationship_reference_counts.get(id).copied()
                            == Some(picture.embedded_reference_count)
                    });
            let relationship_id = if reuse_existing {
                picture
                    .embedded_relationship_id
                    .as_deref()
                    .ok_or_else(|| invalid("SVG lifecycle replacement relationship disappeared"))?
                    .to_owned()
            } else {
                relationship_ids.allocate(drawing_part.rels())?
            };
            budget.relationship_additions.push(PlannedSvgRelationship {
                owner: drawing.uri.clone(),
                relationship_id,
                relationship_type: drawing.relationship_type.clone(),
            });
            let payload_index = picture
                .last_attach_index
                .ok_or_else(|| invalid("SVG lifecycle attachment payload disappeared"))?;
            let payload = intents
                .get(payload_index)
                .and_then(|intent| intent.payload.as_deref())
                .ok_or_else(|| invalid("SVG lifecycle attachment payload disappeared"))?;
            budget.payload_lengths.push(payload.len());
        }
        let drawing_relationship_additions = budget
            .relationship_additions
            .get(relationship_addition_start..)
            .ok_or_else(|| invalid("SVG lifecycle relationship additions disappeared"))?;
        if drawing_relationship_additions.len() != final_attachment_indices.len() {
            return Err(invalid(
                "SVG lifecycle relationship additions and pictures differ",
            ));
        }
        for (picture_index, addition) in final_attachment_indices
            .iter()
            .copied()
            .zip(drawing_relationship_additions)
        {
            let picture = drawing
                .picture_states
                .get(picture_index)
                .ok_or_else(|| invalid("SVG lifecycle relationship picture disappeared"))?;
            let relationship_id_len = addition.relationship_id.len();
            let actual_extension_len = generated_extension_size_with_lengths(
                picture.blip_prefix_len,
                relationship_id_len,
                picture.dialect,
            );
            let actual_growth = picture
                .attach_overhead
                .checked_add(actual_extension_len)
                .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
            if actual_growth >= picture.attach_growth {
                drawing.planned_add_bytes = drawing
                    .planned_add_bytes
                    .checked_add(actual_growth - picture.attach_growth)
                    .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
            } else {
                drawing.planned_remove_bytes = drawing
                    .planned_remove_bytes
                    .checked_add(picture.attach_growth - actual_growth)
                    .ok_or_else(|| invalid("SVG lifecycle drawing growth underflows"))?;
            }
        }
        let final_size = drawing
            .source_len
            .checked_sub(drawing.planned_remove_bytes)
            .and_then(|size| size.checked_add(drawing.planned_add_bytes))
            .ok_or_else(|| invalid("SVG lifecycle drawing output size underflows"))?;
        if final_size > output_limit(workbook) {
            return Err(invalid("SVG lifecycle drawing output exceeds output limit"));
        }
        budget.added_bytes = budget
            .added_bytes
            .checked_add(drawing.planned_add_bytes)
            .ok_or_else(|| invalid("SVG lifecycle drawing growth overflows"))?;
        budget.removed_bytes = budget
            .removed_bytes
            .checked_add(drawing.planned_remove_bytes)
            .ok_or_else(|| invalid("SVG lifecycle drawing growth underflows"))?;
        budget
            .removed_edges
            .try_reserve(drawing.removed_relationship_ids.len())
            .map_err(|source| allocation("SVG lifecycle removed SVG edges", source))?;
        for relationship_id in &drawing.removed_relationship_ids {
            let Some(target) = drawing.relationship_targets.get(relationship_id).cloned() else {
                return Err(invalid(
                    "SVG lifecycle removed relationship target disappeared",
                ));
            };
            budget.removed_edges.push(RemovedSvgEdge {
                owner: drawing.uri.clone(),
                relationship_id: relationship_id.clone(),
                target,
            });
        }
        let mut relationship_count = drawing
            .base_relationships
            .checked_add(drawing.planned_relationship_additions)
            .and_then(|count| count.checked_sub(drawing.planned_relationship_removals))
            .ok_or_else(|| invalid("SVG lifecycle drawing relationship count underflows"))?;
        if let Some(delta) =
            relationship_owner_deltas.get(&canonical_owner_key(workbook, &drawing.uri)?)
        {
            relationship_count = relationship_count
                .checked_add(delta.additions)
                .and_then(|count| count.checked_sub(delta.removals))
                .ok_or_else(|| invalid("SVG lifecycle drawing relationship count underflows"))?;
        }
        if relationship_count
            > workbook
                .inner
                .package
                .read_limits()
                .max_relationships_per_part()
        {
            return Err(invalid(
                "SVG lifecycle drawing relationships exceed caller limit",
            ));
        }
    }
    Ok(budget)
}

fn embedded_reference_counts(scanned: &SourceDrawing<'_>) -> Result<Vec<usize>> {
    let mut counts = Vec::<usize>::new();
    counts
        .try_reserve_exact(scanned.pictures().len())
        .map_err(|source| allocation("SVG lifecycle embedded reference counts", source))?;
    counts.resize(scanned.pictures().len(), 0);

    let mut owner_ranges = Vec::new();
    owner_ranges
        .try_reserve(scanned.pictures().len())
        .map_err(|source| allocation("SVG lifecycle embedded owner ranges", source))?;
    for (picture_index, picture) in scanned.pictures().iter().enumerate() {
        let SvgOwnerState::Embedded(owner) = picture.svg_owner() else {
            continue;
        };
        let relationship_id = owner
            .embedded_relationship_id()
            .ok_or_else(|| invalid("embedded SVG owner has no relationship ID"))?;
        let range = owner.extension_range();
        owner_ranges.push((range.start, range.end, relationship_id, picture_index));
    }
    owner_ranges.sort_unstable_by_key(|owner| owner.0);
    for reference in scanned.relationship_references() {
        let range = reference.range();
        let mut lower = 0usize;
        let mut upper = owner_ranges.len();
        while lower < upper {
            let middle = lower + (upper - lower) / 2;
            if owner_ranges[middle].1 <= range.start {
                lower = middle + 1;
            } else {
                upper = middle;
            }
        }
        let Some((start, end, relationship_id, picture_index)) = owner_ranges.get(lower) else {
            continue;
        };
        if *start < range.end && range.start < *end && *relationship_id == reference.id() {
            let count = counts
                .get_mut(*picture_index)
                .ok_or_else(|| invalid("SVG lifecycle embedded picture disappeared"))?;
            *count = (*count)
                .checked_add(1)
                .ok_or_else(|| invalid("SVG lifecycle relationship reference count overflows"))?;
        }
    }
    Ok(counts)
}

/// Stage every SVG lifecycle intent against one candidate package closure.
/// Multiple pictures in one drawing share one source XML and relationship
/// transition, while media and content-type deltas remain package-scoped.
pub(super) fn plan(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    base_part_bytes: Option<usize>,
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
    parts: &mut Vec<PartChange>,
    relationships: &mut Vec<RelationshipChange>,
    svg_parts: &mut Vec<SvgPartChange>,
    content_types: &mut Option<ContentTypesChange>,
    package_changes: &mut Vec<PackageChange>,
) -> Result<()> {
    if intents.is_empty() {
        return Ok(());
    }
    if intents.len() > MAX_SVG_LIFECYCLE_INTENTS {
        return Err(invalid(format!(
            "SVG lifecycle intents exceed {MAX_SVG_LIFECYCLE_INTENTS}"
        )));
    }
    // Staging an intent takes the base census; the base is immutable, so a
    // census still absent here is taken from the same base now.
    let base_part_bytes = match base_part_bytes {
        Some(bytes) => bytes,
        None => materialized_part_bytes(workbook)?,
    };
    preflight_candidate_budget(
        workbook,
        intents,
        base_part_bytes,
        ordinary_graph,
        relationship_changes,
        parts,
        content_types.as_ref(),
    )?;
    plan_composed(
        workbook,
        intents,
        ordinary_graph,
        relationship_changes,
        parts,
        relationships,
        svg_parts,
        content_types,
        package_changes,
    )
}

#[derive(Debug)]
struct RelationshipOwnerProjection {
    owner: PackURI,
    final_edges: Option<Relationships>,
    source: Option<OwnedRelationships>,
    /// Exact canonical-member metrics admitted before cloning a source-less
    /// graph owner.  The clone is still needed for the final topology map,
    /// but the owner-local byte/event/count limits are checked first.
    canonical_metrics: Option<(usize, usize, usize)>,
    active: bool,
}

#[derive(Debug, Default)]
struct RelationshipTopologyMetrics {
    parts: usize,
    bytes: usize,
    events: usize,
    relationships: usize,
    graph_nodes: usize,
}

/// Build one final relationship-owner topology before any replacement XML is
/// materialized.  Existing owners retain their source token and use the OPC
/// source-preserving plan; newly added graph owners use the neutral canonical
/// plan.  All ordinary graph/delta edges and lifecycle drawing edges are
/// applied to the same indexed owner state, so aggregate limits observe the
/// final topology rather than an additive approximation.
fn plan_final_relationship_topology(
    workbook: &Workbook,
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
    drawing_budget: &SvgDrawingBudget,
    planned_media_uris: &[PackURI],
    cleanup_targets: &[PackURI],
    limits: litchi_opc::ReadLimits,
) -> Result<RelationshipTopologyMetrics> {
    if drawing_budget.relationship_additions.len() != planned_media_uris.len() {
        return Err(invalid(
            "SVG lifecycle relationship and media plans have different attachment counts",
        ));
    }

    let owner_capacity = workbook
        .inner
        .package
        .part_count()
        .checked_add(ordinary_graph.len())
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("SVG lifecycle relationship owner count overflows"))?;
    let mut owners = Vec::<RelationshipOwnerProjection>::new();
    owners
        .try_reserve(owner_capacity)
        .map_err(|source| allocation("SVG lifecycle relationship owner topology", source))?;
    let mut owner_index = HashMap::<String, usize>::new();
    owner_index
        .try_reserve(owner_capacity)
        .map_err(|source| allocation("SVG lifecycle relationship owner index", source))?;

    let root = PackURI::new("/").map_err(invalid)?;
    let root_source = workbook
        .inner
        .package
        .source_relationships_with_limits(&root, limits)?;
    push_relationship_owner(&mut owners, &mut owner_index, root, Some(root_source), true)?;
    for part in workbook.inner.package.iter_parts() {
        let source = workbook
            .inner
            .package
            .source_relationships_with_limits(part.partname(), limits)?;
        push_relationship_owner(
            &mut owners,
            &mut owner_index,
            part.partname().clone(),
            Some(source),
            true,
        )?;
    }

    // RelationshipChange is applied before ordinary GraphChange in Patch::apply.
    // Keeping that order here catches ID collisions before a graph addition is
    // allowed to publish a new owner or source edge.
    for change in relationship_changes {
        let index = owner_index
            .get(&normalized_uri_key(&change.owner))
            .copied()
            .ok_or_else(|| invalid("relationship change owner is not an existing part"))?;
        let owner = owners
            .get_mut(index)
            .ok_or_else(|| invalid("relationship change owner index disappeared"))?;
        if !owner.active {
            return Err(invalid("relationship change targets a removed owner"));
        }
        let final_edges = ensure_projected_relationships(workbook, owner)?;
        if let Some(before) = &change.before {
            let actual = final_edges
                .get(before.r_id())
                .ok_or_else(|| invalid("relationship change before edge is missing"))?;
            if !same_relationship_values(actual, before) {
                return Err(invalid("relationship change before edge differs"));
            }
            final_edges.remove(before.r_id());
        }
        if let Some(after) = &change.after {
            insert_projected_relationship(final_edges, after)?;
        }
    }

    for edge in &drawing_budget.removed_edges {
        let index = owner_index
            .get(&normalized_uri_key(&edge.owner))
            .copied()
            .ok_or_else(|| invalid("SVG lifecycle drawing owner is missing"))?;
        let owner = owners
            .get_mut(index)
            .ok_or_else(|| invalid("SVG lifecycle drawing owner index disappeared"))?;
        if !owner.active {
            return Err(invalid("SVG lifecycle drawing owner was removed"));
        }
        let final_edges = ensure_projected_relationships(workbook, owner)?;
        let actual = final_edges
            .get(&edge.relationship_id)
            .ok_or_else(|| invalid("SVG lifecycle drawing relationship is missing"))?;
        let target = actual.target_partname()?;
        if !target.is_equivalent_to(&edge.target) {
            return Err(invalid("SVG lifecycle drawing relationship target differs"));
        }
        final_edges.remove(&edge.relationship_id);
    }
    for (addition, media_uri) in drawing_budget
        .relationship_additions
        .iter()
        .zip(planned_media_uris)
    {
        let index = owner_index
            .get(&normalized_uri_key(&addition.owner))
            .copied()
            .ok_or_else(|| invalid("SVG lifecycle drawing owner is missing"))?;
        let owner = owners
            .get_mut(index)
            .ok_or_else(|| invalid("SVG lifecycle drawing owner index disappeared"))?;
        if !owner.active {
            return Err(invalid("SVG lifecycle drawing owner was removed"));
        }
        let target_ref = media_uri.relative_ref(owner.owner.base_uri());
        let relationship = Relationship::new_with_mode(
            addition.relationship_id.clone(),
            addition.relationship_type.to_string(),
            target_ref,
            owner.owner.base_uri().to_owned(),
            TargetMode::Internal,
        );
        insert_projected_relationship(
            ensure_projected_relationships(workbook, owner)?,
            &relationship,
        )?;
    }

    // Graph changes are applied after relationship deltas, matching Patch::apply.
    // A newly added part becomes a source-less owner whose canonical plan is
    // still admitted before any XML buffer is constructed.
    for change in ordinary_graph {
        let source_index = owner_index
            .get(&normalized_uri_key(&change.source))
            .copied()
            .ok_or_else(|| invalid("ordinary graph source owner is missing"))?;
        match change.action {
            GraphAction::Add => {
                let source = owners
                    .get_mut(source_index)
                    .ok_or_else(|| invalid("ordinary graph source index disappeared"))?;
                if !source.active {
                    return Err(invalid("ordinary graph source owner was removed"));
                }
                insert_projected_relationship(
                    ensure_projected_relationships(workbook, source)?,
                    &change.relationship,
                )?;
                let target_key = normalized_uri_key(change.part.partname());
                if owner_index.contains_key(&target_key) {
                    return Err(invalid("ordinary graph target owner already exists"));
                }
                if change.part.rels().len() > limits.max_relationships_per_part() {
                    return Err(invalid(format!(
                        "SVG lifecycle relationships for '{}' exceed {}",
                        change.part.partname(),
                        limits.max_relationships_per_part()
                    )));
                }
                let canonical_metrics = if change.part.rels().is_empty() {
                    None
                } else {
                    let canonical_plan = change.part.rels().plan_canonical(limits)?;
                    Some((
                        canonical_plan.final_len(),
                        canonical_plan.event_count(),
                        canonical_plan.relationship_count(),
                    ))
                };
                let final_edges = clone_relationships_bounded(change.part.rels())?;
                let owner = change.part.partname().clone();
                let index = owners.len();
                owners.push(RelationshipOwnerProjection {
                    owner,
                    final_edges: Some(final_edges),
                    source: None,
                    canonical_metrics,
                    active: true,
                });
                owner_index.insert(target_key, index);
            },
            GraphAction::Remove => {
                let source = owners
                    .get_mut(source_index)
                    .ok_or_else(|| invalid("ordinary graph source index disappeared"))?;
                if !source.active {
                    return Err(invalid("ordinary graph source owner was removed"));
                }
                let final_edges = ensure_projected_relationships(workbook, source)?;
                let actual = final_edges
                    .get(change.relationship.r_id())
                    .ok_or_else(|| invalid("ordinary graph relationship is missing"))?;
                if !same_relationship_values(actual, &change.relationship) {
                    return Err(invalid("ordinary graph relationship differs"));
                }
                final_edges.remove(change.relationship.r_id());
                let target_key = normalized_uri_key(change.part.partname());
                let target_index = owner_index
                    .get(&target_key)
                    .copied()
                    .ok_or_else(|| invalid("ordinary graph target owner is missing"))?;
                let target = owners
                    .get_mut(target_index)
                    .ok_or_else(|| invalid("ordinary graph target index disappeared"))?;
                if !target.active {
                    return Err(invalid("ordinary graph target owner was already removed"));
                }
                target.active = false;
                target.final_edges = None;
            },
        }
    }
    for target in cleanup_targets {
        let Some(index) = owner_index.get(&normalized_uri_key(target)).copied() else {
            continue;
        };
        let owner = owners
            .get_mut(index)
            .ok_or_else(|| invalid("SVG lifecycle cleanup owner index disappeared"))?;
        if owner.active {
            owner.active = false;
            owner.final_edges = None;
        }
    }

    let mut metrics = RelationshipTopologyMetrics::default();
    let mut graph_targets = HashSet::<String>::new();
    let graph_node_limit = limits.max_relationship_graph_nodes();

    for owner in &owners {
        if !owner.active {
            continue;
        }
        let final_edges = match owner.final_edges.as_ref() {
            Some(edges) => edges,
            None if owner.source.is_some() => original_relationships(workbook, &owner.owner)?,
            None => {
                return Err(invalid(
                    "SVG lifecycle source-less relationship owner disappeared",
                ));
            },
        };
        if final_edges.len() > limits.max_relationships_per_part() {
            return Err(invalid(format!(
                "SVG lifecycle relationships for '{}' exceed {}",
                owner.owner,
                limits.max_relationships_per_part()
            )));
        }
        let (member_present, bytes, events, count) = if let Some(source) = &owner.source {
            let base = original_relationships(workbook, &owner.owner)?;
            let same = relationships_equal(base, final_edges);
            if same {
                (
                    source.member_present(),
                    if source.member_present() {
                        source.bytes().len()
                    } else {
                        0
                    },
                    if source.member_present() {
                        relationship_xml_event_count(source.bytes(), limits.max_xml_events())?
                    } else {
                        0
                    },
                    final_edges.len(),
                )
            } else {
                let (removals, additions) = relationship_plan_delta(base, final_edges)?;
                let removal_refs = relationship_removal_refs(&removals)?;
                let edits = relationship_edits(&additions)?;
                let plan = source.plan_edit(&edits, &removal_refs, limits)?;
                (
                    plan.member_present(),
                    if plan.member_present() {
                        plan.final_len()
                    } else {
                        0
                    },
                    if plan.member_present() {
                        plan.event_count()
                    } else {
                        0
                    },
                    plan.relationship_count(),
                )
            }
        } else if let Some((bytes, events, count)) = owner.canonical_metrics {
            (true, bytes, events, count)
        } else if final_edges.is_empty() {
            (false, 0, 0, 0)
        } else {
            let plan = final_edges.plan_canonical(limits)?;
            (
                true,
                plan.final_len(),
                plan.event_count(),
                plan.relationship_count(),
            )
        };
        if member_present {
            metrics.parts = metrics
                .parts
                .checked_add(1)
                .ok_or_else(|| invalid("SVG lifecycle relationship part count overflows"))?;
            metrics.bytes = metrics
                .bytes
                .checked_add(bytes)
                .ok_or_else(|| invalid("SVG lifecycle relationship XML bytes overflows"))?;
            metrics.events = metrics
                .events
                .checked_add(events)
                .ok_or_else(|| invalid("SVG lifecycle relationship XML events overflows"))?;
        }
        metrics.relationships = metrics
            .relationships
            .checked_add(count)
            .ok_or_else(|| invalid("SVG lifecycle relationship count overflows"))?;
        for relationship in final_edges.iter() {
            if relationship.target_mode() != TargetMode::Internal {
                continue;
            }
            let target = relationship.target_partname()?;
            insert_graph_target(workbook, &mut graph_targets, &target, graph_node_limit)?;
        }
    }
    metrics.graph_nodes = graph_targets.len();
    Ok(metrics)
}

fn push_relationship_owner(
    owners: &mut Vec<RelationshipOwnerProjection>,
    owner_index: &mut HashMap<String, usize>,
    owner: PackURI,
    source: Option<OwnedRelationships>,
    active: bool,
) -> Result<()> {
    let key = normalized_uri_key(&owner);
    if owner_index.contains_key(&key) {
        return Err(invalid("duplicate relationship owner"));
    }
    let index = owners.len();
    owners.push(RelationshipOwnerProjection {
        owner,
        final_edges: None,
        source,
        canonical_metrics: None,
        active,
    });
    owner_index.insert(key, index);
    Ok(())
}

fn original_relationships<'a>(
    workbook: &'a Workbook,
    owner: &PackURI,
) -> Result<&'a Relationships> {
    if owner.as_str() == "/" {
        Ok(workbook.inner.package.rels())
    } else {
        Ok(workbook.inner.package.get_part(owner)?.rels())
    }
}

fn ensure_projected_relationships<'a>(
    workbook: &Workbook,
    owner: &'a mut RelationshipOwnerProjection,
) -> Result<&'a mut Relationships> {
    if owner.final_edges.is_none() {
        let source = original_relationships(workbook, &owner.owner)?;
        owner.final_edges = Some(clone_relationships_bounded(source)?);
    }
    // A later graph addition may attach an edge to an owner introduced by an
    // earlier addition. Its admission metrics no longer describe the final
    // map once a caller can mutate it; replan before aggregate admission.
    owner.canonical_metrics = None;
    owner
        .final_edges
        .as_mut()
        .ok_or_else(|| invalid("SVG lifecycle projected relationships disappeared"))
}

fn bounded_string(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    owned.push_str(value);
    Ok(owned)
}

fn empty_relationships_bounded(base_uri: &str) -> Result<Relationships> {
    Ok(Relationships::new(bounded_string(
        base_uri,
        "SVG lifecycle relationship base URI",
    )?))
}

fn clone_relationships_bounded(source: &Relationships) -> Result<Relationships> {
    let mut copy = empty_relationships_bounded(source.base_uri())?;
    for relationship in source.iter() {
        copy.try_add_relationship(
            bounded_string(relationship.reltype(), "SVG lifecycle relationship type")?,
            bounded_string(
                relationship.target_ref(),
                "SVG lifecycle relationship target",
            )?,
            bounded_string(relationship.r_id(), "SVG lifecycle relationship ID")?,
            relationship.target_mode(),
        )?;
    }
    Ok(copy)
}

fn clone_relationship_bounded(relationship: &Relationship) -> Result<Relationship> {
    Ok(Relationship::new_with_mode(
        bounded_string(relationship.r_id(), "SVG lifecycle relationship ID")?,
        bounded_string(relationship.reltype(), "SVG lifecycle relationship type")?,
        bounded_string(
            relationship.target_ref(),
            "SVG lifecycle relationship target",
        )?,
        bounded_string(
            relationship.base_uri(),
            "SVG lifecycle relationship base URI",
        )?,
        relationship.target_mode(),
    ))
}

fn insert_projected_relationship(
    relationships: &mut Relationships,
    relationship: &Relationship,
) -> Result<()> {
    if let Some(existing) = relationships.get(relationship.r_id()) {
        if same_relationship_values(existing, relationship) {
            return Ok(());
        }
        return Err(invalid(
            "relationship ID is changed by multiple transaction lanes",
        ));
    }
    relationships.try_add_relationship(
        relationship.reltype().to_owned(),
        relationship.target_ref().to_owned(),
        relationship.r_id().to_owned(),
        relationship.target_mode(),
    )?;
    Ok(())
}

fn relationships_equal(left: &Relationships, right: &Relationships) -> bool {
    left.len() == right.len()
        && left.iter().all(|relationship| {
            right
                .get(relationship.r_id())
                .is_some_and(|candidate| same_relationship_values(relationship, candidate))
        })
}

fn relationship_plan_delta(
    base: &Relationships,
    final_edges: &Relationships,
) -> Result<(Vec<String>, Vec<Relationship>)> {
    let mut removals = Vec::new();
    removals
        .try_reserve(base.len())
        .map_err(|source| allocation("SVG lifecycle relationship plan removals", source))?;
    let mut additions = Vec::new();
    additions
        .try_reserve(final_edges.len())
        .map_err(|source| allocation("SVG lifecycle relationship plan additions", source))?;
    for relationship in base.iter() {
        match final_edges.get(relationship.r_id()) {
            None => removals.push(relationship.r_id().to_owned()),
            Some(candidate) if !same_relationship_values(relationship, candidate) => {
                removals.push(relationship.r_id().to_owned());
                additions.push(clone_relationship_bounded(candidate)?);
            },
            Some(_) => {},
        }
    }
    for relationship in final_edges.iter() {
        if base.get(relationship.r_id()).is_none() {
            additions.push(clone_relationship_bounded(relationship)?);
        }
    }
    removals.sort_unstable();
    additions.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    Ok((removals, additions))
}

fn relationship_removal_refs(removals: &[String]) -> Result<Vec<&str>> {
    let mut refs = Vec::new();
    refs.try_reserve_exact(removals.len())
        .map_err(|source| allocation("SVG lifecycle relationship removal references", source))?;
    for removal in removals {
        refs.push(removal.as_str());
    }
    Ok(refs)
}

fn relationship_edits<'a>(additions: &'a [Relationship]) -> Result<Vec<RelationshipEdit<'a>>> {
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(additions.len())
        .map_err(|source| allocation("SVG lifecycle relationship edit references", source))?;
    for relationship in additions {
        edits.push(RelationshipEdit {
            id: relationship.r_id(),
            reltype: relationship.reltype(),
            target: relationship.target_ref(),
            mode: relationship.target_mode(),
        });
    }
    Ok(edits)
}

fn preflight_candidate_budget(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    base_part_bytes: usize,
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
    part_changes: &[PartChange],
    existing_content_types: Option<&ContentTypesChange>,
) -> Result<()> {
    let limits = workbook.inner.package.read_limits();
    let drawing_budget =
        project_svg_drawing_budget(workbook, intents, ordinary_graph, relationship_changes)?;
    let attach_count = drawing_budget.effective_attach_count;
    let cleanup_targets = safe_svg_cleanup_targets(
        workbook,
        &drawing_budget,
        ordinary_graph,
        relationship_changes,
    )?;
    let mut cleanup_excluded_parts = HashSet::new();
    cleanup_excluded_parts
        .try_reserve(
            ordinary_graph
                .len()
                .checked_add(part_changes.len())
                .ok_or_else(|| invalid("SVG lifecycle cleanup index overflows"))?,
        )
        .map_err(|source| allocation("SVG lifecycle cleanup exclusion index", source))?;
    for change in ordinary_graph {
        if change.action == GraphAction::Remove {
            cleanup_excluded_parts.insert(canonical_target_key(workbook, change.part.partname()));
        }
    }
    for change in part_changes {
        cleanup_excluded_parts.insert(canonical_target_key(workbook, &change.uri));
    }
    let mut candidate_cleanup_targets = Vec::new();
    candidate_cleanup_targets
        .try_reserve(cleanup_targets.len())
        .map_err(|source| allocation("SVG lifecycle candidate cleanup targets", source))?;
    for target in cleanup_targets {
        // Only `PartNotFound` proves a cleanup target absent; any other
        // refusal propagates rather than dropping the target (ADR 0030).
        let present = match workbook.inner.package.get_part(&target) {
            Ok(_) => true,
            Err(litchi_opc::OpcError::PartNotFound(_)) => false,
            Err(error) => return Err(error.into()),
        };
        if present && !cleanup_excluded_parts.contains(&canonical_target_key(workbook, &target)) {
            candidate_cleanup_targets.push(target);
        }
    }
    let cheap_part_count = cheap_candidate_part_count(
        workbook,
        ordinary_graph,
        &candidate_cleanup_targets,
        attach_count,
    )?;
    if cheap_part_count > limits.max_parts() {
        return Err(invalid(format!(
            "SVG lifecycle candidate parts exceed {}",
            limits.max_parts()
        )));
    }
    let planned_media_uris =
        planned_svg_part_uris(workbook, ordinary_graph, part_changes, attach_count)?;
    let (part_count, candidate_bytes) = topology::admit_parts(
        workbook,
        base_part_bytes,
        ordinary_graph,
        part_changes,
        &candidate_cleanup_targets,
        &planned_media_uris,
        &drawing_budget,
    )?;
    if part_count > limits.max_parts() {
        return Err(invalid(format!(
            "SVG lifecycle candidate parts exceed {}",
            limits.max_parts()
        )));
    }
    for intent in intents {
        let Some(payload) = intent.payload.as_deref() else {
            continue;
        };
        if payload.len() > payload_limit(workbook) {
            return Err(invalid(format!(
                "SVG lifecycle payload exceeds {} bytes",
                payload_limit(workbook)
            )));
        }
    }
    if candidate_bytes > usize::try_from(limits.max_total_part_bytes()).unwrap_or(usize::MAX) {
        return Err(invalid(format!(
            "SVG lifecycle candidate part bytes exceed {}",
            limits.max_total_part_bytes()
        )));
    }

    let relationship_metrics = plan_final_relationship_topology(
        workbook,
        ordinary_graph,
        relationship_changes,
        &drawing_budget,
        &planned_media_uris,
        &candidate_cleanup_targets,
        limits,
    )?;
    if relationship_metrics.relationships > limits.max_total_relationships() {
        return Err(invalid(format!(
            "SVG lifecycle relationships exceed {}",
            limits.max_total_relationships()
        )));
    }
    if relationship_metrics.parts > limits.max_relationship_parts() {
        return Err(invalid(format!(
            "SVG lifecycle relationship parts exceed {}",
            limits.max_relationship_parts()
        )));
    }
    if relationship_metrics.bytes > limits.max_total_relationship_xml_bytes() {
        return Err(invalid(format!(
            "SVG lifecycle relationship XML bytes exceed {}",
            limits.max_total_relationship_xml_bytes()
        )));
    }
    if relationship_metrics.events > limits.max_total_relationship_xml_events() {
        return Err(invalid(format!(
            "SVG lifecycle relationship XML events exceed {}",
            limits.max_total_relationship_xml_events()
        )));
    }
    if relationship_metrics.graph_nodes > limits.max_relationship_graph_nodes() {
        return Err(invalid(format!(
            "SVG lifecycle relationship graph nodes exceed {}",
            limits.max_relationship_graph_nodes()
        )));
    }

    let manifest = workbook
        .inner
        .package
        .source_content_types_with_limits(limits)?;
    let manifest_candidate = existing_content_types.map_or(&manifest, |change| &change.after);
    let (has_default_svg, existing_svg_overrides) =
        manifest_svg_inventory(manifest_candidate.bytes())?;
    let mut manifest_removals = Vec::new();
    manifest_removals
        .try_reserve(candidate_cleanup_targets.len())
        .map_err(|source| allocation("SVG lifecycle content-type plan removals", source))?;
    for target in &candidate_cleanup_targets {
        if existing_svg_overrides.contains(&normalized_uri_key(target)) {
            manifest_removals.push(target.clone());
        }
    }
    let mut manifest_additions = Vec::new();
    if !has_default_svg {
        manifest_additions
            .try_reserve(planned_media_uris.len())
            .map_err(|source| allocation("SVG lifecycle content-type plan additions", source))?;
        for uri in &planned_media_uris {
            if !existing_svg_overrides.contains(&normalized_uri_key(uri)) {
                manifest_additions.push(ContentTypeEdit {
                    part_name: uri,
                    content_type: "image/svg+xml",
                });
            }
        }
    }
    let content_types_plan =
        manifest_candidate.plan_edit(&manifest_additions, &manifest_removals, limits)?;
    let content_type_bytes = content_types_plan.final_len();
    if content_type_bytes > content_types_output_limit(workbook)
        || content_types_plan.mapping_count() > limits.max_content_type_mappings()
    {
        return Err(invalid(format!(
            "SVG lifecycle content-types final state exceeds caller limits ({} bytes, {} mappings)",
            content_type_bytes,
            content_types_plan.mapping_count(),
        )));
    }
    Ok(())
}

fn cheap_candidate_part_count(
    workbook: &Workbook,
    ordinary_graph: &[super::model::GraphChange],
    cleanup_targets: &[PackURI],
    attach_count: usize,
) -> Result<usize> {
    let mut count = workbook.inner.package.part_count();
    for change in ordinary_graph {
        count = match change.action {
            GraphAction::Add => count
                .checked_add(1)
                .ok_or_else(|| invalid("SVG lifecycle candidate part count overflows"))?,
            GraphAction::Remove => count
                .checked_sub(1)
                .ok_or_else(|| invalid("SVG lifecycle candidate part count underflows"))?,
        };
    }
    count = count
        .checked_add(attach_count)
        .ok_or_else(|| invalid("SVG lifecycle candidate part count overflows"))?;
    count = count
        .checked_sub(cleanup_targets.len())
        .ok_or_else(|| invalid("SVG lifecycle candidate part count underflows"))?;
    Ok(count)
}

/// Compute the final part count and aggregate part bytes from one projected
/// part-name map.  Every ordinary add/remove, source part replacement,
/// lifecycle media add, and cleanup removal is applied before the aggregate
/// limit is checked; this keeps exact-fit candidates from being rejected by
/// additive upper bounds.
fn final_part_budget(
    workbook: &Workbook,
    base_part_bytes: usize,
    ordinary_graph: &[super::model::GraphChange],
    part_changes: &[PartChange],
    cleanup_targets: &[PackURI],
    planned_media_uris: &[PackURI],
    drawing_budget: &SvgDrawingBudget,
) -> Result<(usize, usize)> {
    if drawing_budget.payload_lengths.len() != planned_media_uris.len() {
        return Err(invalid(
            "SVG lifecycle payload and media plans have different attachment counts",
        ));
    }
    let mut capacity = workbook
        .inner
        .package
        .part_count()
        .checked_add(ordinary_graph.len())
        .and_then(|count| count.checked_add(part_changes.len()))
        .and_then(|count| count.checked_add(planned_media_uris.len()))
        .ok_or_else(|| invalid("SVG lifecycle candidate part index overflows"))?;
    let mut sizes = HashMap::<String, usize>::new();
    sizes
        .try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle candidate part index", source))?;
    let mut initial_bytes = 0usize;
    // Every part's inflated length enters the budget, so every payload is
    // decoded (ADR 0030); the base census has already forced them.
    for part in workbook.inner.package.try_iter_parts() {
        let part = part?;
        initial_bytes = initial_bytes
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("SVG lifecycle candidate part bytes overflow"))?;
        sizes.insert(normalized_uri_key(part.partname()), part.blob().len());
    }
    if initial_bytes != base_part_bytes {
        return Err(invalid(
            "SVG lifecycle base part byte census differs from transaction snapshot",
        ));
    }
    let mut removed_graph_parts = HashSet::new();
    let mut added_graph_parts = HashSet::new();
    removed_graph_parts
        .try_reserve(ordinary_graph.len())
        .map_err(|source| allocation("SVG lifecycle removed part index", source))?;
    added_graph_parts
        .try_reserve(ordinary_graph.len())
        .map_err(|source| allocation("SVG lifecycle added part index", source))?;
    for change in ordinary_graph {
        if change.action == GraphAction::Remove {
            removed_graph_parts.insert(normalized_uri_key(change.part.partname()));
            sizes.remove(&normalized_uri_key(change.part.partname()));
        }
    }
    for change in ordinary_graph {
        if change.action != GraphAction::Add {
            continue;
        }
        let key = normalized_uri_key(change.part.partname());
        if !removed_graph_parts.contains(&key) && sizes.contains_key(&key) {
            return Err(invalid(format!(
                "SVG lifecycle ordinary graph part '{}' collides with the final package",
                change.part.partname()
            )));
        }
        if !added_graph_parts.insert(key.clone()) {
            return Err(invalid(format!(
                "SVG lifecycle ordinary graph part '{}' is added more than once",
                change.part.partname()
            )));
        }
        sizes.insert(key, change.part.blob().len());
    }
    for change in part_changes {
        let key = normalized_uri_key(&change.uri);
        if removed_graph_parts.contains(&key) {
            return Err(invalid(format!(
                "SVG lifecycle part replacement '{}' conflicts with ordinary graph removal",
                change.uri
            )));
        }
        sizes.insert(key, change.after.len());
    }
    for target in cleanup_targets {
        sizes.remove(&canonical_target_key(workbook, target));
    }
    let mut payload_bytes = 0usize;
    for (uri, payload_len) in planned_media_uris
        .iter()
        .zip(drawing_budget.payload_lengths.iter().copied())
    {
        payload_bytes = payload_bytes
            .checked_add(payload_len)
            .ok_or_else(|| invalid("SVG lifecycle payload bytes overflow"))?;
        if sizes.insert(normalized_uri_key(uri), payload_len).is_some() {
            return Err(invalid(format!(
                "SVG lifecycle media part '{}' collides with the final package",
                uri
            )));
        }
    }
    if payload_bytes != drawing_budget.effective_payload_bytes {
        return Err(invalid(
            "SVG lifecycle payload byte census differs from attachment projection",
        ));
    }
    let mut candidate_bytes = sizes.values().try_fold(0usize, |total, size| {
        total
            .checked_add(*size)
            .ok_or_else(|| invalid("SVG lifecycle candidate part bytes overflow"))
    })?;
    candidate_bytes = candidate_bytes
        .checked_sub(drawing_budget.removed_bytes)
        .and_then(|total| total.checked_add(drawing_budget.added_bytes))
        .ok_or_else(|| invalid("SVG lifecycle candidate drawing bytes underflow"))?;
    let maximum_part_bytes = usize::try_from(workbook.inner.package.read_limits().max_part_bytes())
        .unwrap_or(usize::MAX);
    if sizes.values().any(|size| *size > maximum_part_bytes) {
        return Err(invalid(format!(
            "SVG lifecycle candidate part exceeds {maximum_part_bytes} bytes"
        )));
    }
    capacity = sizes.len();
    Ok((capacity, candidate_bytes))
}

/// Reserve the exact generated media-name identities used by the later
/// planner.  The index includes ordinary graph additions and source part
/// changes before any SVG output is materialized, so a batch cannot discover
/// a case-equivalent or package-prefix collision halfway through planning.
fn planned_svg_part_uris(
    workbook: &Workbook,
    ordinary_graph: &[super::model::GraphChange],
    part_changes: &[PartChange],
    attach_count: usize,
) -> Result<Vec<PackURI>> {
    let package_parts = workbook.inner.package.iter_parts().count();
    let capacity = package_parts
        .checked_add(ordinary_graph.len())
        .and_then(|count| count.checked_add(part_changes.len()))
        .and_then(|count| count.checked_add(attach_count))
        .ok_or_else(|| invalid("SVG lifecycle media identity count overflows"))?;
    let mut occupied_part_names = HashSet::new();
    occupied_part_names
        .try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle media identity index", source))?;
    let mut occupied_part_prefixes = HashSet::new();
    occupied_part_prefixes
        .try_reserve(capacity.saturating_mul(2))
        .map_err(|source| allocation("SVG lifecycle media prefix index", source))?;
    for part in workbook.inner.package.iter_parts() {
        index_part_name(
            part.partname(),
            &mut occupied_part_names,
            &mut occupied_part_prefixes,
        );
    }
    for change in ordinary_graph {
        if change.action == GraphAction::Add {
            index_part_name(
                change.part.partname(),
                &mut occupied_part_names,
                &mut occupied_part_prefixes,
            );
        }
    }
    for change in part_changes {
        index_part_name(
            &change.uri,
            &mut occupied_part_names,
            &mut occupied_part_prefixes,
        );
    }
    let mut reserved_uris = HashSet::new();
    reserved_uris
        .try_reserve(attach_count)
        .map_err(|source| allocation("SVG lifecycle media identities", source))?;
    let mut media_uris = Vec::new();
    media_uris
        .try_reserve(attach_count)
        .map_err(|source| allocation("SVG lifecycle media identities", source))?;
    let mut next_part_index = 0usize;
    for _ in 0..attach_count {
        let uri = allocate_part_uri_with_reservations(
            &reserved_uris,
            &occupied_part_names,
            &occupied_part_prefixes,
            &mut next_part_index,
        )?;
        reserved_uris.insert(uri.clone());
        index_part_name(&uri, &mut occupied_part_names, &mut occupied_part_prefixes);
        media_uris.push(uri);
    }
    Ok(media_uris)
}

fn normalized_uri_key(uri: &PackURI) -> String {
    uri.as_str()
        .bytes()
        .map(|byte| char::from(byte.to_ascii_lowercase()))
        .collect()
}

fn canonical_owner_key(workbook: &Workbook, owner: &PackURI) -> Result<String> {
    let canonical = workbook
        .inner
        .package
        .get_part(owner)
        .map_or_else(|_| owner.clone(), |part| part.partname().clone());
    let candidate = normalized_uri_key_candidate(&canonical)?;
    match candidate {
        Cow::Borrowed(value) => bounded_string(value, "SVG lifecycle relationship owner key"),
        Cow::Owned(value) => Ok(value),
    }
}

fn index_relationship_owner_deltas(
    workbook: &Workbook,
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
) -> Result<HashMap<String, RelationshipOwnerDelta>> {
    let capacity = ordinary_graph
        .len()
        .checked_add(relationship_changes.len())
        .ok_or_else(|| invalid("SVG lifecycle relationship owner index overflows"))?;
    let mut deltas = HashMap::new();
    deltas
        .try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle relationship owner index", source))?;
    for change in relationship_changes {
        let (additions, removals) = match (change.before.is_some(), change.after.is_some()) {
            (false, true) => (1usize, 0usize),
            (true, false) => (0usize, 1usize),
            _ => continue,
        };
        let key = canonical_owner_key(workbook, &change.owner)?;
        let delta = deltas
            .entry(key)
            .or_insert_with(RelationshipOwnerDelta::default);
        delta.additions = delta
            .additions
            .checked_add(additions)
            .ok_or_else(|| invalid("SVG lifecycle relationship additions overflow"))?;
        delta.removals = delta
            .removals
            .checked_add(removals)
            .ok_or_else(|| invalid("SVG lifecycle relationship removals overflow"))?;
    }
    for change in ordinary_graph {
        let key = canonical_owner_key(workbook, &change.source)?;
        let delta = deltas
            .entry(key)
            .or_insert_with(RelationshipOwnerDelta::default);
        match change.action {
            GraphAction::Add => {
                delta.additions = delta
                    .additions
                    .checked_add(1)
                    .ok_or_else(|| invalid("SVG lifecycle relationship additions overflow"))?;
            },
            GraphAction::Remove => {
                delta.removals = delta
                    .removals
                    .checked_add(1)
                    .ok_or_else(|| invalid("SVG lifecycle relationship removals overflow"))?;
            },
        }
    }
    Ok(deltas)
}

fn relationship_key(owner: &PackURI, relationship_id: &str) -> String {
    let mut key = normalized_uri_key(owner);
    key.push('#');
    key.push_str(relationship_id);
    key
}

fn canonical_target_key(workbook: &Workbook, target: &PackURI) -> String {
    workbook.inner.package.get_part(target).map_or_else(
        |_| normalized_uri_key(target),
        |part| normalized_uri_key(part.partname()),
    )
}

fn insert_graph_target(
    workbook: &Workbook,
    graph_targets: &mut HashSet<String>,
    target: &PackURI,
    maximum: usize,
) -> Result<()> {
    let canonical = workbook
        .inner
        .package
        .get_part(target)
        .map_or_else(|_| target.clone(), |part| part.partname().clone());
    let candidate = normalized_uri_key_candidate(&canonical)?;
    if graph_targets.contains(candidate.as_ref()) {
        return Ok(());
    }
    if graph_targets.len() >= maximum {
        return Err(invalid(format!(
            "SVG lifecycle relationship graph nodes exceed {maximum}"
        )));
    }
    graph_targets
        .try_reserve(1)
        .map_err(|source| allocation("SVG lifecycle relationship final graph nodes", source))?;
    let key = match candidate {
        Cow::Borrowed(value) => {
            bounded_string(value, "SVG lifecycle relationship graph target key")?
        },
        Cow::Owned(value) => value,
    };
    graph_targets.insert(key);
    Ok(())
}

fn normalized_uri_key_candidate<'a>(uri: &'a PackURI) -> Result<Cow<'a, str>> {
    if !uri.as_str().bytes().any(|byte| byte.is_ascii_uppercase()) {
        return Ok(Cow::Borrowed(uri.as_str()));
    }
    let mut key = String::new();
    key.try_reserve_exact(uri.as_str().len())
        .map_err(|source| allocation("SVG lifecycle relationship graph target key", source))?;
    for byte in uri.as_str().bytes() {
        key.push(char::from(byte.to_ascii_lowercase()));
    }
    Ok(Cow::Owned(key))
}

/// Determine which SVG media leaves are actually removed by the projected
/// graph.  The source package is walked once, while ordinary and lifecycle
/// edge removals are indexed by owner/id so shared relationships are not
/// mistaken for final target removal.
fn safe_svg_cleanup_targets(
    workbook: &Workbook,
    drawing_budget: &SvgDrawingBudget,
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
) -> Result<Vec<PackURI>> {
    if drawing_budget.removed_edges.is_empty() {
        return Ok(Vec::new());
    }
    let mut target_keys = HashSet::new();
    target_keys
        .try_reserve(drawing_budget.removed_edges.len())
        .map_err(|source| allocation("SVG lifecycle cleanup target census", source))?;
    for edge in &drawing_budget.removed_edges {
        target_keys.insert(canonical_target_key(workbook, &edge.target));
    }
    let mut removed_edges = HashSet::new();
    removed_edges
        .try_reserve(
            drawing_budget
                .removed_edges
                .len()
                .saturating_add(relationship_changes.len())
                .saturating_add(ordinary_graph.len()),
        )
        .map_err(|source| allocation("SVG lifecycle cleanup edge census", source))?;
    let mut removed_sources = HashSet::new();
    removed_sources
        .try_reserve(ordinary_graph.len())
        .map_err(|source| allocation("SVG lifecycle cleanup source census", source))?;
    for edge in &drawing_budget.removed_edges {
        removed_edges.insert(relationship_key(&edge.owner, &edge.relationship_id));
    }
    for change in relationship_changes {
        if let Some(before) = &change.before {
            removed_edges.insert(relationship_key(&change.owner, before.r_id()));
        }
    }
    for change in ordinary_graph {
        if change.action == GraphAction::Remove {
            removed_sources.insert(normalized_uri_key(&change.source));
            removed_edges.insert(relationship_key(&change.source, change.relationship.r_id()));
        }
    }

    let mut incoming = HashSet::new();
    incoming
        .try_reserve(target_keys.len())
        .map_err(|source| allocation("SVG lifecycle cleanup incoming census", source))?;
    let mut visit_relationships = |owner: &PackURI, relationships: &Relationships| -> Result<()> {
        if removed_sources.contains(&normalized_uri_key(owner)) {
            return Ok(());
        }
        for relationship in relationships.iter() {
            if removed_edges.contains(&relationship_key(owner, relationship.r_id())) {
                continue;
            }
            if relationship.target_mode() != TargetMode::Internal {
                continue;
            }
            let target = relationship.target_partname()?;
            let key = canonical_target_key(workbook, &target);
            if target_keys.contains(&key) {
                incoming.insert(key);
            }
        }
        Ok(())
    };
    let root = PackURI::new("/").map_err(invalid)?;
    visit_relationships(&root, workbook.inner.package.rels())?;
    for part in workbook.inner.package.iter_parts() {
        visit_relationships(part.partname(), part.rels())?;
    }
    for change in relationship_changes {
        if let Some(after) = &change.after {
            if after.target_mode() == TargetMode::Internal {
                let target = after.target_partname()?;
                let key = canonical_target_key(workbook, &target);
                if target_keys.contains(&key) {
                    incoming.insert(key);
                }
            }
        }
    }
    for change in ordinary_graph {
        if change.action != GraphAction::Add {
            continue;
        }
        if change.relationship.target_mode() == TargetMode::Internal {
            let target = change.relationship.target_partname()?;
            let key = canonical_target_key(workbook, &target);
            if target_keys.contains(&key) {
                incoming.insert(key);
            }
        }
        for relationship in change.part.rels().iter() {
            if relationship.target_mode() == TargetMode::Internal {
                let target = relationship.target_partname()?;
                let key = canonical_target_key(workbook, &target);
                if target_keys.contains(&key) {
                    incoming.insert(key);
                }
            }
        }
    }
    let mut cleanup = Vec::new();
    cleanup
        .try_reserve(drawing_budget.removed_edges.len())
        .map_err(|source| allocation("SVG lifecycle safe cleanup targets", source))?;
    let mut seen = HashSet::new();
    seen.try_reserve(drawing_budget.removed_edges.len())
        .map_err(|source| allocation("SVG lifecycle safe cleanup target index", source))?;
    for edge in &drawing_budget.removed_edges {
        let key = canonical_target_key(workbook, &edge.target);
        if incoming.contains(&key) || !seen.insert(key) {
            continue;
        }
        cleanup.push(edge.target.clone());
    }
    Ok(cleanup)
}

/// Capture the worksheet and package relationship reads that select an SVG
/// picture.  The resulting guards are attached to the in-memory patch after
/// the candidate workbook has been published, so their `after` snapshots
/// include any ordinary worksheet edits in the same transaction.
pub(super) fn read_guards(
    source: &Workbook,
    target: &Workbook,
    intents: &[SvgLifecycleIntent],
) -> Result<Box<[SvgReadGuard]>> {
    if intents.is_empty() {
        return Ok(Box::new([]));
    }
    let root = PackURI::new("/").map_err(invalid)?;
    let mut owners = Vec::new();
    owners
        .try_reserve(intents.len().saturating_add(1))
        .map_err(|source| allocation("SVG lifecycle read guards", source))?;
    owners.push(root);
    for intent in intents {
        let sheet = source
            .inner
            .sheets
            .get(intent.position)
            .ok_or_else(|| invalid("SVG lifecycle guard worksheet disappeared"))?;
        if !owners.iter().any(|owner| owner == &sheet.part_uri) {
            owners.push(sheet.part_uri.clone());
        }
    }

    let mut guards = Vec::new();
    guards
        .try_reserve(owners.len())
        .map_err(|source| allocation("SVG lifecycle read guards", source))?;
    let limits = source.inner.package.read_limits();
    for owner in owners {
        let before_relationships = source
            .inner
            .package
            .source_relationships_with_limits(&owner, limits)?;
        let after_relationships = target
            .inner
            .package
            .source_relationships_with_limits(&owner, limits)?;
        let before_xml = if owner.as_str() == "/" {
            None
        } else {
            Some(source.inner.package.get_part(&owner)?.blob_arc())
        };
        let after_xml = if owner.as_str() == "/" {
            None
        } else {
            Some(target.inner.package.get_part(&owner)?.blob_arc())
        };
        guards.push(SvgReadGuard {
            owner,
            before_xml,
            after_xml,
            before_relationships,
            after_relationships,
        });
    }
    Ok(guards.into_boxed_slice())
}

/// One bounded source inventory shared by every selected picture in a
/// drawing.  The source scanner and package tokens are immutable, so keeping
/// this short-lived cache avoids rescanning a large drawing once per intent.
struct FinalDrawingInventory<'a> {
    drawing: PackURI,
    worksheet_xml: Arc<Vec<u8>>,
    worksheet_relationships: OwnedRelationships,
    drawing_xml: Arc<Vec<u8>>,
    drawing_relationships: OwnedRelationships,
    content_types: OwnedContentTypes,
    source: Arc<SourceDrawing<'a>>,
}

/// Capture and validate the complete final readback contract for each
/// affected picture.  The mutation planner proves source ranges and graph
/// deltas, but those proofs inspect preimages.  This paired facade read is
/// deliberately performed only after all ordinary and SVG changes have been
/// published into the candidate workbook.
pub(super) fn capture_final_guards(
    source: &Workbook,
    target: &Workbook,
    intents: &[SvgLifecycleIntent],
) -> Result<Box<[SvgFinalGuard]>> {
    if intents.is_empty() {
        return Ok(Box::new([]));
    }
    let mut seen = HashSet::new();
    seen.try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle final selector guards", source))?;
    let mut guards = Vec::new();
    guards
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle final selector guards", source))?;
    let mut source_inventories = HashMap::new();
    source_inventories
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle final drawing inventories", source))?;
    let mut target_inventories = HashMap::new();
    target_inventories
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle final drawing inventories", source))?;
    let source_content_types = source
        .inner
        .package
        .source_content_types_with_limits(source.inner.package.read_limits())?;
    let target_content_types = target
        .inner
        .package
        .source_content_types_with_limits(target.inner.package.read_limits())?;
    let mut expected_by_selector = HashMap::new();
    expected_by_selector
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle final selector expectations", source))?;
    for intent in intents {
        expected_by_selector.insert((intent.position, intent.selector), intent.attach);
    }
    for intent in intents {
        let sheet = source
            .inner
            .sheets
            .get(intent.position)
            .ok_or_else(|| invalid("SVG lifecycle guard worksheet disappeared"))?;
        let key = (
            sheet.part_uri.clone(),
            selector_drawing(intent.selector),
            selector_picture(intent.selector),
        );
        if !seen.insert(key.clone()) {
            continue;
        }
        if let Entry::Vacant(entry) = source_inventories.entry((key.0.clone(), key.1)) {
            let inventory = capture_final_inventory(source, &key.0, key.1, &source_content_types)?;
            entry.insert(inventory);
        }
        if let Entry::Vacant(entry) = target_inventories.entry((key.0.clone(), key.1)) {
            let inventory = capture_final_inventory(target, &key.0, key.1, &target_content_types)?;
            entry.insert(inventory);
        }
        let before_inventory = source_inventories
            .get(&(key.0.clone(), key.1))
            .ok_or_else(|| invalid("SVG lifecycle source inventory disappeared"))?;
        let after_inventory = target_inventories
            .get(&(key.0.clone(), key.1))
            .ok_or_else(|| invalid("SVG lifecycle target inventory disappeared"))?;
        let before = capture_final_state_from_inventory(source, before_inventory, key.2)?;
        let after = capture_final_state_from_inventory(target, after_inventory, key.2)?;
        ensure_unchanged_raster(&before, &after)?;
        let expected_attached = expected_by_selector
            .get(&(intent.position, intent.selector))
            .copied()
            .unwrap_or_else(|| before.svg.is_some());
        if after.svg.is_some() != expected_attached {
            return Err(invalid(
                "SVG lifecycle final picture owner does not match the staged operation",
            ));
        }
        guards.push(SvgFinalGuard {
            worksheet: key.0,
            drawing_ordinal: key.1,
            picture_ordinal: key.2,
            before,
            after,
            expected_attached,
        });
    }
    Ok(guards.into_boxed_slice())
}

/// Validate the source state captured when a patch was created before any
/// patch transition is published.  This catches a shared SVG media edit that
/// leaves the selected drawing XML and relationship member bytes unchanged.
pub(super) fn validate_before_guards(workbook: &Workbook, guards: &[SvgFinalGuard]) -> Result<()> {
    validate_guard_states(workbook, guards, true)
}

/// Re-read a patch candidate through the same bounded source-backed worksheet
/// seam used by commit planning.  Only the stored target state is accepted;
/// all XML, relationships, content-type, anchor, raster, and admitted SVG
/// media expectations are exact source-bound values.
pub(super) fn validate_final_guards(workbook: &Workbook, guards: &[SvgFinalGuard]) -> Result<()> {
    validate_guard_states(workbook, guards, false)
}

fn validate_guard_states(
    workbook: &Workbook,
    guards: &[SvgFinalGuard],
    before: bool,
) -> Result<()> {
    let mut inventories = HashMap::new();
    inventories
        .try_reserve(guards.len())
        .map_err(|source| allocation("SVG lifecycle final drawing inventories", source))?;
    let content_types = workbook
        .inner
        .package
        .source_content_types_with_limits(workbook.inner.package.read_limits())?;
    for guard in guards {
        let key = (guard.worksheet.clone(), guard.drawing_ordinal);
        if !inventories.contains_key(&key) {
            let inventory = capture_final_inventory(
                workbook,
                &guard.worksheet,
                guard.drawing_ordinal,
                &content_types,
            )?;
            inventories.insert(key.clone(), inventory);
        }
        let inventory = inventories
            .get(&key)
            .ok_or_else(|| invalid("SVG lifecycle final inventory disappeared"))?;
        let actual = capture_final_state_from_inventory(workbook, inventory, guard.picture_ordinal)
            .map_err(|error| match error {
                Error::PatchConflict { .. } => error,
                _ => Error::PatchConflict {
                    part: guard.worksheet.to_string(),
                },
            })?;
        let expected = if before { &guard.before } else { &guard.after };
        if &actual != expected {
            return Err(Error::PatchConflict {
                part: guard.worksheet.to_string(),
            });
        }
        let expected_attached = if before {
            expected.svg.is_some()
        } else {
            guard.expected_attached
        };
        if actual.svg.is_some() != expected_attached {
            return Err(Error::PatchConflict {
                part: guard.worksheet.to_string(),
            });
        }
    }
    Ok(())
}

fn capture_final_inventory<'a>(
    workbook: &'a Workbook,
    worksheet_uri: &PackURI,
    drawing_ordinal: usize,
    content_types: &OwnedContentTypes,
) -> Result<FinalDrawingInventory<'a>> {
    let worksheet_part = workbook.inner.package.get_part(worksheet_uri)?;
    let (drawing_uri, _) = resolve_drawing_for_uri(workbook, worksheet_uri, drawing_ordinal)?;
    let drawing_part = workbook.inner.package.get_part(&drawing_uri)?;
    let source = Arc::new(SourceDrawing::scan_with_limits(
        drawing_part.blob(),
        drawing_ordinal,
        scan_limits(workbook),
    )?);
    Ok(FinalDrawingInventory {
        drawing: drawing_uri.clone(),
        worksheet_xml: worksheet_part.blob_arc(),
        worksheet_relationships: workbook.inner.package.source_relationships_with_limits(
            worksheet_uri,
            workbook.inner.package.read_limits(),
        )?,
        drawing_xml: drawing_part.blob_arc(),
        drawing_relationships: workbook
            .inner
            .package
            .source_relationships_with_limits(&drawing_uri, workbook.inner.package.read_limits())?,
        content_types: content_types.clone(),
        source,
    })
}

fn capture_final_state_from_inventory(
    workbook: &Workbook,
    inventory: &FinalDrawingInventory<'_>,
    picture_ordinal: usize,
) -> Result<SvgFinalState> {
    let drawing_part = workbook.inner.package.get_part(&inventory.drawing)?;
    let picture = inventory.source.picture(picture_ordinal)?;
    let raster_relationship_id = picture.raster_relationship_id();
    let raster_relationship = drawing_part
        .rels()
        .get(raster_relationship_id)
        .ok_or_else(|| invalid("final raster relationship is missing"))?;
    if raster_relationship.target_mode() != TargetMode::Internal
        || !matches!(raster_relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Err(invalid("final raster relationship is unsupported"));
    }
    let raster_target = raster_relationship.target_partname()?;
    ensure_final_media_uri(&raster_target, "raster")?;
    let raster_part = workbook.inner.package.get_part(&raster_target)?;
    if raster_part.content_type() != ct::PNG || !raster_part.rels().is_empty() {
        return Err(invalid("final raster fallback must be an inert PNG part"));
    }
    let svg = match picture.svg_owner() {
        SvgOwnerState::None | SvgOwnerState::Opaque => None,
        SvgOwnerState::Embedded(owner) => {
            let relationship_id = owner
                .embedded_relationship_id()
                .ok_or_else(|| invalid("embedded SVG owner has no relationship ID"))?;
            let relationship = drawing_part
                .rels()
                .get(relationship_id)
                .ok_or_else(|| invalid("final SVG relationship is missing"))?;
            if relationship.target_mode() != TargetMode::Internal
                || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
            {
                return Err(invalid("final SVG relationship is unsupported"));
            }
            let target = relationship.target_partname()?;
            ensure_final_media_uri(&target, "SVG")?;
            let part = workbook.inner.package.get_part(&target)?;
            if part.content_type() != "image/svg+xml" || !part.rels().is_empty() {
                return Err(invalid("final SVG target must be an inert SVG part"));
            }
            if !manifest_covers_svg(inventory.content_types.bytes(), &target)? {
                return Err(invalid(
                    "final SVG media part has no image/svg+xml content-type declaration",
                ));
            }
            Some(SvgFinalMedia {
                relationship_id: relationship_id.to_owned().into_boxed_str(),
                relationship_type: relationship.reltype().to_owned().into_boxed_str(),
                target_mode: relationship.target_mode(),
                part: part.partname().clone(),
                content_type: part.content_type().to_owned().into_boxed_str(),
                bytes: part.blob_arc(),
            })
        },
        SvgOwnerState::Linked(_) => {
            return Err(invalid("final SVG owner is linked"));
        },
        SvgOwnerState::Ambiguous => {
            return Err(invalid("final picture has duplicate SVG owners"));
        },
        SvgOwnerState::Refused => {
            return Err(invalid("final picture has a refused SVG owner"));
        },
    };
    Ok(SvgFinalState {
        drawing: inventory.drawing.clone(),
        worksheet_xml: Arc::clone(&inventory.worksheet_xml),
        worksheet_relationships: inventory.worksheet_relationships.clone(),
        drawing_xml: Arc::clone(&inventory.drawing_xml),
        drawing_relationships: inventory.drawing_relationships.clone(),
        content_types: inventory.content_types.clone(),
        anchor: *picture.anchor(),
        raster_relationship_id: raster_relationship_id.to_owned().into_boxed_str(),
        raster_relationship_type: raster_relationship.reltype().to_owned().into_boxed_str(),
        raster_target_mode: raster_relationship.target_mode(),
        raster_part: raster_part.partname().clone(),
        raster_content_type: raster_part.content_type().to_owned().into_boxed_str(),
        raster_bytes: raster_part.blob_arc(),
        svg,
    })
}

fn ensure_final_media_uri(uri: &PackURI, kind: &str) -> Result<()> {
    if !is_media_uri(uri) {
        return Err(invalid(format!(
            "final {kind} target is outside canonical /xl/media/"
        )));
    }
    Ok(())
}

fn is_media_uri(uri: &PackURI) -> bool {
    let mut segments = uri.as_str().split('/');
    segments.next() == Some("")
        && segments
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("xl"))
        && segments
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("media"))
        && segments.next().is_some_and(|segment| !segment.is_empty())
}

fn resolve_drawing_for_uri(
    workbook: &Workbook,
    worksheet_uri: &PackURI,
    drawing_ordinal: usize,
) -> Result<(PackURI, Relationship)> {
    let worksheet = workbook.inner.package.get_part(worksheet_uri)?;
    let references = worksheet_drawing_references(workbook, worksheet.blob())?;
    let relationship_id = references
        .get(drawing_ordinal)
        .ok_or_else(|| invalid("drawing ordinal is outside worksheet drawings"))?;
    let relationship = worksheet
        .rels()
        .get(relationship_id)
        .cloned()
        .ok_or_else(|| invalid("worksheet drawing relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING)
    {
        return Err(invalid("worksheet drawing relationship is unsupported"));
    }
    let drawing_uri = relationship.target_partname()?;
    let drawing = workbook.inner.package.get_part(&drawing_uri)?;
    if drawing.content_type() != ct::OFC_DRAWING {
        return Err(invalid(
            "worksheet drawing target has an unsupported content type",
        ));
    }
    Ok((drawing_uri, relationship))
}

fn ensure_unchanged_raster(before: &SvgFinalState, after: &SvgFinalState) -> Result<()> {
    if before.anchor != after.anchor {
        return Err(invalid(
            "SVG lifecycle changed the selected picture anchor geometry",
        ));
    }
    if !before.raster_part.is_equivalent_to(&after.raster_part)
        || before.raster_relationship_id != after.raster_relationship_id
        || before.raster_relationship_type != after.raster_relationship_type
        || before.raster_target_mode != after.raster_target_mode
        || before.raster_content_type != after.raster_content_type
        || before.raster_bytes.as_slice() != after.raster_bytes.as_slice()
    {
        return Err(invalid(
            "SVG lifecycle changed the inert PNG compatibility fallback",
        ));
    }
    Ok(())
}

/// Preflight an operation against the source package plus operations already
/// staged in this transaction.  The source scanner remains authoritative for
/// the initial owner state; this small projection prevents a later detach
/// from silently becoming a no-op after an earlier attach (and permits the
/// corresponding detach-then-attach sequence).  The returned transition is
/// committed by the caller only after all later intent/payload admissions
/// succeed, so a failed stage leaves both semantic maps unchanged.
pub(super) fn preflight_with_pending(
    workbook: &Workbook,
    position: usize,
    selector: PictureSelector,
    attach: bool,
    cache: &mut HashMap<(usize, usize), DrawingPreflight>,
    projected: &mut HashMap<(usize, PictureSelector), ProjectedSvgOwner>,
) -> Result<PreparedSvgTransition> {
    projected
        .try_reserve(1)
        .map_err(|source| allocation("SVG lifecycle projected owner cache", source))?;
    let key = (position, selector_drawing(selector));
    let new_facts = if cache.contains_key(&key) {
        None
    } else {
        cache
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle preflight cache", source))?;
        Some(build_preflight_cache(workbook, position, selector)?)
    };
    let facts = match new_facts.as_ref() {
        Some(facts) => facts,
        None => cache
            .get(&key)
            .ok_or_else(|| invalid("SVG lifecycle preflight cache disappeared"))?,
    };
    let part = workbook.inner.package.get_part(&facts.drawing_uri)?;
    let picture = facts
        .pictures
        .get(selector_picture(selector))
        .ok_or_else(|| invalid("picture ordinal is outside worksheet drawing pictures"))?;
    validate_raster_fallback_id(workbook, part, &picture.raster_relationship_id)?;
    let owner = if let Some(owner) = projected.get(&(position, selector)) {
        *owner
    } else {
        picture.owner
    };
    match (attach, owner) {
        (true, ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque) => Ok(PreparedSvgTransition {
            cache_key: key,
            new_facts,
            projected_key: (position, selector),
            projected_owner: ProjectedSvgOwner::Embedded,
            effective: true,
        }),
        (true, ProjectedSvgOwner::Embedded) => Err(invalid(
            "selected picture already has an embedded SVG owner",
        )),
        (true, ProjectedSvgOwner::Linked) => Err(invalid(
            "linked SVG owners cannot be replaced by attach_svg",
        )),
        (true, ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused) => Err(invalid(
            "selected picture has an ambiguous or refused SVG owner",
        )),
        (false, ProjectedSvgOwner::Embedded) => Ok(PreparedSvgTransition {
            cache_key: key,
            new_facts,
            projected_key: (position, selector),
            projected_owner: ProjectedSvgOwner::None,
            effective: true,
        }),
        (false, ProjectedSvgOwner::Linked) => Err(invalid("linked SVG owners cannot be detached")),
        (false, ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused) => Err(invalid(
            "selected picture has an ambiguous or refused SVG owner",
        )),
        (false, owner @ (ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque)) => {
            Ok(PreparedSvgTransition {
                cache_key: key,
                new_facts,
                projected_key: (position, selector),
                projected_owner: owner,
                effective: false,
            })
        },
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum ProjectedSvgOwner {
    None,
    Opaque,
    Embedded,
    Linked,
    Ambiguous,
    Refused,
}

pub(super) struct PreparedSvgTransition {
    cache_key: (usize, usize),
    new_facts: Option<DrawingPreflight>,
    projected_key: (usize, PictureSelector),
    projected_owner: ProjectedSvgOwner,
    effective: bool,
}

impl PreparedSvgTransition {
    pub(super) fn is_effective(&self) -> bool {
        self.effective
    }

    pub(super) fn commit(
        self,
        cache: &mut HashMap<(usize, usize), DrawingPreflight>,
        projected: &mut HashMap<(usize, PictureSelector), ProjectedSvgOwner>,
    ) {
        if let Some(facts) = self.new_facts {
            cache.insert(self.cache_key, facts);
        }
        projected.insert(self.projected_key, self.projected_owner);
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
    source_owners: Vec<ProjectedSvgOwner>,
    relationship_ids: RelationshipIdAllocator,
}

struct CoalescedGroup {
    indices: Vec<usize>,
    replacements: HashSet<PictureSelector>,
}

fn coalesce_repeated_group(
    state: &DrawingPlanState,
    group: &[usize],
    intents: &[SvgLifecycleIntent],
) -> Result<CoalescedGroup> {
    #[derive(Clone, Copy)]
    struct Entry {
        initial: ProjectedSvgOwner,
        current: ProjectedSvgOwner,
        last_attach: Option<usize>,
        last_detach: Option<usize>,
        saw_attach: bool,
        saw_detach: bool,
    }

    let mut entries = HashMap::<PictureSelector, Entry>::new();
    entries
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle repeated-selector projection", source))?;
    let mut order = Vec::new();
    order
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle repeated-selector order", source))?;
    for index in group {
        let intent = intents
            .get(*index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let picture = selector_picture(intent.selector);
        let initial = *state
            .source_owners
            .get(picture)
            .ok_or_else(|| invalid("picture ordinal is outside worksheet drawing pictures"))?;
        let entry = entries.entry(intent.selector).or_insert_with(|| {
            order.push(intent.selector);
            Entry {
                initial,
                current: initial,
                last_attach: None,
                last_detach: None,
                saw_attach: false,
                saw_detach: false,
            }
        });
        if intent.attach {
            entry.saw_attach = true;
            entry.last_attach = Some(*index);
            match entry.current {
                ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => {
                    entry.current = ProjectedSvgOwner::Embedded;
                },
                ProjectedSvgOwner::Embedded => {
                    return Err(invalid(
                        "selected picture already has an embedded SVG owner",
                    ));
                },
                ProjectedSvgOwner::Linked => {
                    return Err(invalid(
                        "linked SVG owners cannot be replaced by attach_svg",
                    ));
                },
                ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => {
                    return Err(invalid(
                        "selected picture has an ambiguous or refused SVG owner",
                    ));
                },
            }
        } else {
            entry.saw_detach = true;
            entry.last_detach = Some(*index);
            match entry.current {
                ProjectedSvgOwner::Embedded => {
                    entry.current = if entry.initial == ProjectedSvgOwner::Opaque {
                        ProjectedSvgOwner::Opaque
                    } else {
                        ProjectedSvgOwner::None
                    };
                },
                ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => {},
                ProjectedSvgOwner::Linked => {
                    return Err(invalid("linked SVG owners cannot be detached"));
                },
                ProjectedSvgOwner::Ambiguous | ProjectedSvgOwner::Refused => {
                    return Err(invalid(
                        "selected picture has an ambiguous or refused SVG owner",
                    ));
                },
            }
        }
    }

    let mut indices = Vec::new();
    indices
        .try_reserve(order.len())
        .map_err(|source| allocation("SVG lifecycle coalesced selectors", source))?;
    let mut replacements = HashSet::new();
    replacements
        .try_reserve(order.len())
        .map_err(|source| allocation("SVG lifecycle replacement selectors", source))?;
    for selector in order {
        let entry = entries
            .get(&selector)
            .ok_or_else(|| invalid("SVG lifecycle repeated-selector projection disappeared"))?;
        if entry.initial == entry.current {
            if entry.initial == ProjectedSvgOwner::Embedded && entry.saw_detach && entry.saw_attach
            {
                let index = entry
                    .last_attach
                    .ok_or_else(|| invalid("SVG lifecycle replacement payload disappeared"))?;
                indices.push(index);
                replacements.insert(selector);
            }
            continue;
        }
        let index = match entry.current {
            ProjectedSvgOwner::Embedded => entry
                .last_attach
                .ok_or_else(|| invalid("SVG lifecycle attach payload disappeared"))?,
            ProjectedSvgOwner::None | ProjectedSvgOwner::Opaque => entry
                .last_detach
                .ok_or_else(|| invalid("SVG lifecycle detach operation disappeared"))?,
            ProjectedSvgOwner::Linked
            | ProjectedSvgOwner::Ambiguous
            | ProjectedSvgOwner::Refused => {
                return Err(invalid("SVG lifecycle owner projection is unsupported"));
            },
        };
        indices.push(index);
    }
    Ok(CoalescedGroup {
        indices,
        replacements,
    })
}

fn new_drawing_plan_state(
    workbook: &Workbook,
    uri: PackURI,
    drawing_ordinal: usize,
) -> Result<DrawingPlanState> {
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
    let scanned = SourceDrawing::scan_with_limits(
        before_xml.bytes(),
        drawing_ordinal,
        scan_limits(workbook),
    )?;
    let mut source_owners = Vec::new();
    source_owners
        .try_reserve(scanned.pictures().len())
        .map_err(|source| allocation("SVG lifecycle source owner inventory", source))?;
    for picture in scanned.pictures() {
        source_owners.push(match picture.svg_owner() {
            SvgOwnerState::None => ProjectedSvgOwner::None,
            SvgOwnerState::Opaque => ProjectedSvgOwner::Opaque,
            SvgOwnerState::Embedded(_) => ProjectedSvgOwner::Embedded,
            SvgOwnerState::Linked(_) => ProjectedSvgOwner::Linked,
            SvgOwnerState::Ambiguous => ProjectedSvgOwner::Ambiguous,
            SvgOwnerState::Refused => ProjectedSvgOwner::Refused,
        });
    }
    Ok(DrawingPlanState {
        current_xml: before_xml.clone(),
        current_rels: before_rels.clone(),
        before_xml,
        before_rels,
        before_edges: clone_relationships_bounded(part.rels())?,
        current_edges: clone_relationships_bounded(part.rels())?,
        source_owners,
        relationship_ids: RelationshipIdAllocator::from_source(part.rels())?,
        uri,
    })
}

fn index_ordinary_relationship_changes<'a>(
    workbook: &Workbook,
    ordinary_graph: &'a [super::model::GraphChange],
    relationship_changes: &'a [RelationshipChange],
) -> Result<HashMap<String, HashSet<&'a str>>> {
    let capacity = ordinary_graph
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(relationship_changes.len()))
        .ok_or_else(|| invalid("SVG lifecycle ordinary owner count overflows"))?;
    let mut owners: HashMap<String, HashSet<&'a str>> = HashMap::new();
    owners
        .try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle ordinary relationship owners", source))?;
    let mut record = |owner: &PackURI, id: Option<&'a str>| -> Result<()> {
        let key = canonical_owner_key(workbook, owner)?;
        let ids = owners.entry(key).or_default();
        if let Some(id) = id {
            if !ids.contains(id) {
                ids.try_reserve(1).map_err(|source| {
                    allocation("SVG lifecycle ordinary relationship IDs", source)
                })?;
                ids.insert(id);
            }
        }
        Ok(())
    };
    for change in ordinary_graph {
        record(&change.source, Some(change.relationship.r_id()))?;
        record(change.part.partname(), None)?;
    }
    for change in relationship_changes {
        record(
            &change.owner,
            change
                .before
                .as_ref()
                .or(change.after.as_ref())
                .map(|relationship| relationship.r_id()),
        )?;
    }
    Ok(owners)
}

fn plan_composed(
    workbook: &Workbook,
    intents: &[SvgLifecycleIntent],
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
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
    let mut state_by_uri: HashMap<String, usize> = HashMap::new();
    state_by_uri
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle drawing lookup", source))?;
    let mut drawing_by_selector: HashMap<(usize, usize), PackURI> = HashMap::new();
    drawing_by_selector
        .try_reserve(intents.len())
        .map_err(|source| allocation("SVG lifecycle worksheet drawing lookup", source))?;
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
    let mut occupied_part_names = HashSet::new();
    occupied_part_names
        .try_reserve(
            workbook
                .inner
                .package
                .iter_parts()
                .count()
                .saturating_add(ordinary_graph.len())
                .saturating_add(parts.len())
                .saturating_add(intents.len()),
        )
        .map_err(|source| allocation("SVG lifecycle media identity index", source))?;
    let mut occupied_part_prefixes = HashSet::new();
    occupied_part_prefixes
        .try_reserve(
            workbook
                .inner
                .package
                .iter_parts()
                .count()
                .saturating_mul(2)
                .saturating_add(ordinary_graph.len().saturating_mul(2))
                .saturating_add(parts.len().saturating_mul(2))
                .saturating_add(intents.len()),
        )
        .map_err(|source| allocation("SVG lifecycle media prefix index", source))?;
    for part in workbook.inner.package.iter_parts() {
        index_part_name(
            part.partname(),
            &mut occupied_part_names,
            &mut occupied_part_prefixes,
        );
    }
    for change in ordinary_graph {
        if change.action == GraphAction::Add {
            index_part_name(
                change.part.partname(),
                &mut occupied_part_names,
                &mut occupied_part_prefixes,
            );
        }
    }
    for change in parts.iter() {
        index_part_name(
            &change.uri,
            &mut occupied_part_names,
            &mut occupied_part_prefixes,
        );
    }
    let mut next_part_index = 0usize;
    let mut cleanup_targets = CleanupTargets::with_capacity(intents.len())?;
    let manifest_before = workbook
        .inner
        .package
        .source_content_types_with_limits(workbook.inner.package.read_limits())?;
    let mut manifest_after = content_types
        .as_ref()
        .map_or_else(|| manifest_before.clone(), |change| change.after.clone());

    for (intent_index, intent) in intents.iter().enumerate() {
        let selector_key = (intent.position, selector_drawing(intent.selector));
        let drawing_uri = if let Some(uri) = drawing_by_selector.get(&selector_key) {
            uri.clone()
        } else {
            let (uri, _) = resolve_drawing(workbook, intent.position, intent.selector)?;
            drawing_by_selector.insert(selector_key, uri.clone());
            uri
        };
        let drawing_key = drawing_uri.as_str().to_ascii_lowercase();
        let state_index = if let Some(index) = state_by_uri.get(&drawing_key).copied() {
            index
        } else {
            state_by_uri.insert(drawing_key, states.len());
            states.push(new_drawing_plan_state(
                workbook,
                drawing_uri.clone(),
                selector_drawing(intent.selector),
            )?);
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
        let original_group = &state_groups[state_index];
        let mut selectors = HashSet::new();
        selectors
            .try_reserve(original_group.len())
            .map_err(|source| allocation("SVG lifecycle picture selectors", source))?;
        let unique_selectors = original_group
            .iter()
            .all(|index| selectors.insert(intents[*index].selector));
        let coalesced = if unique_selectors {
            CoalescedGroup {
                indices: original_group.clone(),
                replacements: HashSet::new(),
            }
        } else {
            coalesce_repeated_group(&states[state_index], original_group, intents)?
        };
        let group = &coalesced.indices;
        if group.is_empty() {
            continue;
        }
        let all_attach = group.iter().all(|index| intents[*index].attach);
        let all_detach = group.iter().all(|index| !intents[*index].attach);
        if !coalesced.replacements.is_empty() {
            if coalesced.replacements.len() != group.len() || coalesced.replacements.len() != 1 {
                return Err(invalid(
                    "multiple SVG replacement sequences in one drawing are unsupported",
                ));
            }
            let selector = coalesced
                .replacements
                .iter()
                .next()
                .copied()
                .ok_or_else(|| invalid("SVG lifecycle replacement selector disappeared"))?;
            let detach_index = original_group
                .iter()
                .find(|index| intents[**index].selector == selector && !intents[**index].attach)
                .ok_or_else(|| invalid("SVG lifecycle replacement detach disappeared"))?;
            let attach_index = original_group
                .iter()
                .rev()
                .find(|index| intents[**index].selector == selector && intents[**index].attach)
                .ok_or_else(|| invalid("SVG lifecycle replacement attach disappeared"))?;
            plan_composed_replacement(
                workbook,
                &intents[*detach_index],
                &intents[*attach_index],
                &mut states[state_index],
                &mut pending_parts,
                &mut reserved_uris,
                &mut occupied_part_names,
                &mut occupied_part_prefixes,
                &mut next_part_index,
                &mut manifest_after,
                &mut cleanup_targets,
            )?;
            continue;
        }
        if all_attach {
            plan_composed_attach_batch(
                workbook,
                group,
                intents,
                &mut states[state_index],
                &mut pending_parts,
                &mut reserved_uris,
                &mut occupied_part_names,
                &mut occupied_part_prefixes,
                &mut next_part_index,
                &mut manifest_after,
            )?;
        } else if all_detach {
            plan_composed_detach_batch(
                workbook,
                group,
                intents,
                &mut states[state_index],
                &pending_parts,
                &mut manifest_after,
                &mut cleanup_targets,
            )?;
        } else {
            plan_composed_mixed_batch(
                workbook,
                group,
                intents,
                &mut states[state_index],
                &mut pending_parts,
                &mut reserved_uris,
                &mut occupied_part_names,
                &mut occupied_part_prefixes,
                &mut next_part_index,
                &mut manifest_after,
                &mut cleanup_targets,
            )?;
        }
    }

    finalize_svg_target_cleanup(
        workbook,
        &states,
        ordinary_graph,
        relationship_changes,
        &mut pending_parts,
        &cleanup_targets,
    )?;

    // Child drawing planners stage only XML, relationship, and media state.
    // Build the one final content-types token after the package-wide cleanup
    // census, so an old SVG override and its replacement never consume the
    // caller's final mapping/byte cap at the same time.
    let mut pending_svg_uris = Vec::new();
    pending_svg_uris
        .try_reserve(pending_parts.len())
        .map_err(|source| allocation("SVG lifecycle final content-type additions", source))?;
    for change in &pending_parts {
        if change.action == GraphAction::Add && change.part.content_type() == "image/svg+xml" {
            pending_svg_uris.push(change.part.partname().clone());
        }
    }
    let mut final_manifest_removals = Vec::new();
    final_manifest_removals
        .try_reserve(pending_parts.len())
        .map_err(|source| allocation("SVG lifecycle final content-type removals", source))?;
    for change in &pending_parts {
        if change.action == GraphAction::Remove {
            final_manifest_removals.push(change.part.partname().clone());
        }
    }
    manifest_after = materialize_content_types_transition(
        &manifest_after,
        &final_manifest_removals,
        &pending_svg_uris,
        workbook.inner.package.read_limits(),
    )?;

    parts
        .try_reserve(states.len())
        .map_err(|source| allocation("SVG lifecycle final part changes", source))?;
    let mut relationship_capacity = 0usize;
    for state in &states {
        let state_capacity = state
            .before_edges
            .len()
            .checked_add(state.current_edges.len())
            .ok_or_else(|| invalid("SVG lifecycle relationship result count overflows"))?;
        relationship_capacity = relationship_capacity
            .checked_add(state_capacity)
            .ok_or_else(|| invalid("SVG lifecycle relationship result count overflows"))?;
    }
    relationships
        .try_reserve(relationship_capacity)
        .map_err(|source| allocation("SVG lifecycle final relationship changes", source))?;

    let ordinary_changes =
        index_ordinary_relationship_changes(workbook, ordinary_graph, relationship_changes)?;
    for mut state in states {
        let changed_ids = changed_relationship_ids(&state.before_edges, &state.current_edges)?;
        let owner_key = canonical_owner_key(workbook, &state.uri)?;
        let ordinary_owner_changes = ordinary_changes.get(&owner_key);
        if state.current_xml.bytes() != state.before_xml.bytes() {
            parts.push(PartChange {
                uri: state.uri.clone(),
                before: state.before_xml.shared_bytes(),
                after: state.current_xml.shared_bytes(),
                before_source: Some(state.before_xml),
                after_source: Some(state.current_xml),
            });
        }
        if !changed_ids.is_empty() {
            if let Some(ordinary_ids) = ordinary_owner_changes {
                for id in changed_ids {
                    if ordinary_ids.contains(id.as_str()) {
                        return Err(invalid(
                            "SVG lifecycle relationship ID is changed by multiple transaction lanes",
                        ));
                    }
                    relationships.push(RelationshipChange {
                        owner: state.uri.clone(),
                        before: state.before_edges.get(&id).cloned(),
                        after: state.current_edges.get(&id).cloned(),
                        before_source: None,
                        after_source: None,
                    });
                }
            } else {
                state.current_rels = materialize_relationship_transition(
                    workbook,
                    &state.before_rels,
                    &state.before_edges,
                    &state.current_edges,
                )?;
                let id = changed_ids
                    .first()
                    .ok_or_else(|| invalid("SVG lifecycle relationship ID disappeared"))?;
                relationships.push(RelationshipChange {
                    owner: state.uri,
                    before: state.before_edges.get(id).cloned(),
                    after: state.current_edges.get(id).cloned(),
                    before_source: Some(state.before_rels),
                    after_source: Some(state.current_rels),
                });
            }
        }
    }
    svg_parts
        .try_reserve(pending_parts.len())
        .map_err(|source| allocation("SVG lifecycle final SVG parts", source))?;
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

fn materialize_relationship_transition(
    workbook: &Workbook,
    before: &OwnedRelationships,
    before_edges: &Relationships,
    final_edges: &Relationships,
) -> Result<OwnedRelationships> {
    let mut removals = Vec::<String>::new();
    let mut additions = Vec::<Relationship>::new();
    removals
        .try_reserve(before_edges.len())
        .map_err(|source| allocation("SVG lifecycle relationship removals", source))?;
    additions
        .try_reserve(final_edges.len())
        .map_err(|source| allocation("SVG lifecycle relationship additions", source))?;

    for relationship in before_edges.iter() {
        match final_edges.get(relationship.r_id()) {
            None => removals.push(relationship.r_id().to_owned()),
            Some(final_relationship)
                if !same_relationship_values(relationship, final_relationship) =>
            {
                removals.push(relationship.r_id().to_owned());
                additions.push(clone_relationship_bounded(final_relationship)?);
            },
            Some(_) => {},
        }
    }
    for relationship in final_edges.iter() {
        if before_edges.get(relationship.r_id()).is_none() {
            additions.push(clone_relationship_bounded(relationship)?);
        }
    }
    removals.sort_unstable();
    additions.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    if removals.is_empty() && additions.is_empty() {
        return Ok(before.clone());
    }

    let removal_refs = relationship_removal_refs(&removals)?;
    let edits = relationship_edits(&additions)?;
    let limits = workbook.inner.package.read_limits();
    let plan = before.plan_edit(&edits, &removal_refs, limits)?;
    Ok(plan.materialize(limits)?)
}

fn same_relationship_values(left: &Relationship, right: &Relationship) -> bool {
    left.r_id() == right.r_id()
        && left.reltype() == right.reltype()
        && left.target_ref() == right.target_ref()
        && left.target_mode() == right.target_mode()
}

/// Plan the common detach-only batch in one source scan and one drawing XML
/// rewrite. Relationship members remain source-bound tokens; removing their
/// individual edges is cheap compared with rescanning and reallocating the
/// complete drawing once per picture.
fn plan_composed_detach_batch(
    workbook: &Workbook,
    group: &[usize],
    intents: &[SvgLifecycleIntent],
    state: &mut DrawingPlanState,
    pending_parts: &[SvgPartChange],
    _manifest: &mut OwnedContentTypes,
    cleanup_targets: &mut CleanupTargets,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let scanned = SourceDrawing::scan_with_limits(
        state.before_xml.bytes(),
        selector_drawing(
            intents
                .get(
                    *group
                        .first()
                        .ok_or_else(|| invalid("empty SVG lifecycle group"))?,
                )
                .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?
                .selector,
        ),
        limits,
    )?;
    let pending_media_index = index_pending_media_parts(pending_parts)?;
    let mut operations = Vec::new();
    operations
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle detach batch", source))?;
    for index in group {
        let intent = intents
            .get(*index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let picture = scanned.picture(selector_picture(intent.selector))?;
        validate_raster_fallback_edges(workbook, &state.current_edges, picture)?;
        let owner = match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner,
            SvgOwnerState::None | SvgOwnerState::Opaque => {
                return Err(invalid("selected picture has no admitted SVG owner"));
            },
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
        validate_svg_target_candidate(
            workbook,
            &relationship,
            &target,
            pending_parts,
            &pending_media_index,
        )?;
        let (opening, removed) = detach_ranges(state.before_xml.bytes(), picture, owner)?;
        operations.push(DetachOperation {
            relationship_id: relationship_id.to_owned(),
            target,
            opening,
            removed,
            owner_range: owner.extension_range(),
            ext_list: picture
                .ext_list_range()
                .map(|range| (range.range(), range.start_end())),
        });
    }

    let mut xml_removals = Vec::new();
    xml_removals
        .try_reserve(operations.len())
        .map_err(|source| allocation("SVG lifecycle detach XML ranges", source))?;
    for operation in &operations {
        xml_removals.push(XmlRemoval {
            opening: operation.opening.clone(),
            removed: operation.removed,
        });
    }
    let mut ext_lists = Vec::<ExtListRemoval>::new();
    ext_lists
        .try_reserve(operations.len())
        .map_err(|source| allocation("SVG lifecycle detach extension lists", source))?;
    let mut ext_list_index = HashMap::<source::ByteRange, usize>::new();
    ext_list_index
        .try_reserve(operations.len())
        .map_err(|source| allocation("SVG lifecycle detach extension list index", source))?;
    for operation in &operations {
        let Some((range, start_end)) = operation.ext_list else {
            continue;
        };
        if let Some(existing_index) = ext_list_index.get(&range).copied() {
            let existing = ext_lists
                .get_mut(existing_index)
                .ok_or_else(|| invalid("SVG lifecycle extension list index disappeared"))?;
            existing
                .owner_ranges
                .try_reserve(1)
                .map_err(|source| allocation("SVG lifecycle extension list owners", source))?;
            existing.owner_ranges.push(operation.owner_range);
        } else {
            let mut owner_ranges = Vec::new();
            owner_ranges
                .try_reserve(1)
                .map_err(|source| allocation("SVG lifecycle extension list owners", source))?;
            owner_ranges.push(operation.owner_range);
            ext_lists.push(ExtListRemoval {
                range,
                start_end,
                owner_ranges,
            });
            ext_list_index.insert(range, ext_lists.len() - 1);
        }
    }
    coalesce_extension_list_removals(state.before_xml.bytes(), &mut xml_removals, &ext_lists)?;
    xml_removals.sort_by(|left, right| {
        left.opening
            .start
            .cmp(&right.opening.start)
            .then(left.opening.end.cmp(&right.opening.end))
    });
    for pair in xml_removals.windows(2) {
        if pair[0].opening.end > pair[1].opening.start {
            return Err(invalid("SVG detach ranges overlap"));
        }
    }
    let mut updates = Vec::new();
    updates
        .try_reserve(xml_removals.len())
        .map_err(|source| allocation("SVG lifecycle detach XML updates", source))?;
    for removal in &xml_removals {
        updates.push(OwnedElementUpdate {
            start_tag: removal.opening.clone(),
            edit: OwnedElementEdit::Remove,
        });
    }
    state.current_xml = state.before_xml.update_elements(&updates, output_cap)?;

    let references = scanned.relationship_references();
    let mut removed_relationships = Vec::new();
    removed_relationships
        .try_reserve(operations.len())
        .map_err(|source| allocation("SVG lifecycle detach relationships", source))?;
    let mut removed_relationship_ids = HashSet::new();
    removed_relationship_ids
        .try_reserve(operations.len())
        .map_err(|source| allocation("SVG lifecycle detach relationship IDs", source))?;
    let mut removed_ranges = Vec::new();
    removed_ranges
        .try_reserve(xml_removals.len())
        .map_err(|source| allocation("SVG lifecycle removed ranges", source))?;
    removed_ranges.extend(xml_removals.iter().map(|removal| removal.removed));
    removed_ranges.sort_by_key(|range| range.start);
    let mut remaining_references = HashMap::<&str, usize>::new();
    remaining_references
        .try_reserve(references.len())
        .map_err(|source| allocation("SVG lifecycle relationship reference census", source))?;
    let mut removed_range_index = 0usize;
    for reference in references {
        while removed_range_index < removed_ranges.len()
            && removed_ranges[removed_range_index].end <= reference.range().start
        {
            removed_range_index += 1;
        }
        let removed = removed_range_index < removed_ranges.len()
            && ranges_overlap(reference.range(), removed_ranges[removed_range_index]);
        if !removed {
            let entry = remaining_references.entry(reference.id()).or_insert(0);
            *entry = entry
                .checked_add(1)
                .ok_or_else(|| invalid("SVG lifecycle relationship reference count overflow"))?;
        }
    }
    for operation in &operations {
        let remaining = remaining_references
            .get(operation.relationship_id.as_str())
            .copied()
            .unwrap_or(0);
        if remaining != 0 || removed_relationship_ids.contains(&operation.relationship_id) {
            continue;
        }
        removed_relationship_ids.insert(operation.relationship_id.clone());
        removed_relationships.push(operation.relationship_id.clone());
        cleanup_targets.push(operation.target.clone())?;
    }
    for relationship_id in removed_relationships {
        state.current_edges.remove(&relationship_id);
    }
    Ok(())
}

struct DetachOperation {
    relationship_id: String,
    target: PackURI,
    opening: std::ops::Range<usize>,
    removed: source::ByteRange,
    owner_range: source::ByteRange,
    ext_list: Option<(source::ByteRange, usize)>,
}

struct XmlRemoval {
    opening: std::ops::Range<usize>,
    removed: source::ByteRange,
}

struct ExtListRemoval {
    range: source::ByteRange,
    start_end: usize,
    owner_ranges: Vec<source::ByteRange>,
}

fn coalesce_extension_list_removals(
    source: &[u8],
    xml_removals: &mut Vec<XmlRemoval>,
    ext_lists: &[ExtListRemoval],
) -> Result<()> {
    let mut valid_indices = Vec::new();
    valid_indices
        .try_reserve(ext_lists.len())
        .map_err(|source| allocation("SVG lifecycle valid extension lists", source))?;

    let mut range_capacity = xml_removals.len();
    for ext_list in ext_lists {
        range_capacity = range_capacity
            .checked_add(ext_list.owner_ranges.len())
            .ok_or_else(|| invalid("SVG lifecycle extension removal range count overflows"))?;
    }
    let mut removed_ranges = HashSet::<source::ByteRange>::new();
    removed_ranges
        .try_reserve(range_capacity)
        .map_err(|source| allocation("SVG lifecycle extension removal ranges", source))?;

    for (index, ext_list) in ext_lists.iter().enumerate() {
        if !ext_list_contains_only_owner_ranges(source, ext_list.range, &ext_list.owner_ranges)? {
            continue;
        }
        valid_indices.push(index);
        removed_ranges.insert(ext_list.range);
        for owner_range in &ext_list.owner_ranges {
            removed_ranges.insert(*owner_range);
        }
    }
    if valid_indices.is_empty() {
        return Ok(());
    }

    xml_removals.retain(|removal| !removed_ranges.contains(&removal.removed));
    xml_removals
        .try_reserve(valid_indices.len())
        .map_err(|source| allocation("SVG lifecycle extension XML removals", source))?;
    for index in valid_indices {
        let ext_list = ext_lists
            .get(index)
            .ok_or_else(|| invalid("SVG lifecycle extension list disappeared"))?;
        xml_removals.push(XmlRemoval {
            opening: ext_list.range.start..ext_list.start_end,
            removed: ext_list.range,
        });
    }
    Ok(())
}

struct MixedAttachOperation {
    intent_index: usize,
    relationship_id: String,
    part_uri: PackURI,
    parent: std::ops::Range<usize>,
    parent_empty: bool,
    parent_name_len: usize,
    fragment_len: usize,
}

/// Coalesce the only meaningful repeated-selector replacement sequence into
/// one source drawing rewrite.  The public staging projection permits
/// `detach_svg(...).attach_svg(...)` so a caller can replace an existing SVG
/// payload in one ordinary transaction.  Reusing the existing owner range
/// keeps opaque extension siblings byte exact and avoids the intermediate
/// detach/attach XML and relationship allocations.
fn plan_composed_replacement(
    workbook: &Workbook,
    detach: &SvgLifecycleIntent,
    attach: &SvgLifecycleIntent,
    state: &mut DrawingPlanState,
    pending_parts: &mut Vec<SvgPartChange>,
    reserved_uris: &mut HashSet<PackURI>,
    occupied_part_names: &mut HashSet<String>,
    occupied_part_prefixes: &mut HashSet<String>,
    next_part_index: &mut usize,
    _manifest: &mut OwnedContentTypes,
    cleanup_targets: &mut CleanupTargets,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let scanned = SourceDrawing::scan_with_limits(
        state.before_xml.bytes(),
        selector_drawing(detach.selector),
        limits,
    )?;
    let pending_media_index = index_pending_media_parts(pending_parts)?;
    let picture = scanned.picture(selector_picture(detach.selector))?;
    validate_raster_fallback_edges(workbook, &state.current_edges, picture)?;
    let owner = match picture.svg_owner() {
        SvgOwnerState::Embedded(owner) => owner,
        SvgOwnerState::None | SvgOwnerState::Opaque => {
            return Err(invalid("selected picture has no admitted SVG owner"));
        },
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
    let old_target = relationship.target_partname()?;
    validate_svg_target_candidate(
        workbook,
        &relationship,
        &old_target,
        pending_parts,
        &pending_media_index,
    )?;
    let payload = attach
        .payload
        .as_deref()
        .ok_or_else(|| invalid("attach_svg has no staged payload"))?;
    if payload.len() > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle payload exceeds {} bytes",
            output_cap
        )));
    }

    let owner_range = owner.extension_range();
    let opaque_references = scanned
        .relationship_references()
        .iter()
        .filter(|reference| {
            reference.id() == relationship_id && !ranges_overlap(reference.range(), owner_range)
        })
        .count();
    let retain_old_relationship = opaque_references != 0;
    let new_relationship_id = if retain_old_relationship {
        state.relationship_ids.allocate(&state.current_edges)?
    } else {
        relationship_id.to_owned()
    };
    let part_uri = allocate_part_uri_with_reservations(
        reserved_uris,
        occupied_part_names,
        occupied_part_prefixes,
        next_part_index,
    )?;
    reserved_uris.insert(part_uri.clone());
    index_part_name(&part_uri, occupied_part_names, occupied_part_prefixes);
    let relationship_type = if scanned.relationship_dialect() == source::RelationshipDialect::Strict
    {
        rt::STRICT_IMAGE
    } else {
        rt::IMAGE
    };
    let target_ref = part_uri.relative_ref(state.uri.base_uri());
    let next_relationship = Relationship::new_with_mode(
        new_relationship_id.clone(),
        relationship_type.to_owned(),
        target_ref.clone(),
        state.uri.base_uri().to_owned(),
        TargetMode::Internal,
    );
    let next_edges_len = state
        .current_edges
        .len()
        .checked_add(usize::from(retain_old_relationship))
        .ok_or_else(|| invalid("SVG relationship count overflows"))?;
    if next_edges_len
        > workbook
            .inner
            .package
            .read_limits()
            .max_relationships_per_part()
    {
        return Err(invalid("SVG relationship count exceeds caller limit"));
    }
    let blip = picture.blip_range();
    let extension_len =
        generated_extension_size(blip.prefix(), &new_relationship_id, scanned.dialect());
    let owner_opening = opening_tag_range(state.before_xml.bytes(), owner_range)?;
    let output_size = state
        .before_xml
        .bytes()
        .len()
        .checked_sub(owner_range.len()?)
        .and_then(|size| size.checked_add(extension_len))
        .ok_or_else(|| invalid("SVG lifecycle replacement output size overflow"))?;
    if output_size > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle drawing output exceeds {} bytes",
            output_cap
        )));
    }

    let extension = generated_extension(blip.prefix(), &new_relationship_id, scanned.dialect())?;
    let next_xml = state.before_xml.update_elements(
        &[OwnedElementUpdate {
            start_tag: owner_opening,
            edit: OwnedElementEdit::Replace(&extension),
        }],
        output_cap,
    )?;
    if !retain_old_relationship {
        state.current_edges.remove(relationship_id);
        cleanup_targets.push(old_target)?;
    }
    state.current_edges.try_add_relationship(
        next_relationship.reltype().to_owned(),
        next_relationship.target_ref().to_owned(),
        next_relationship.r_id().to_owned(),
        next_relationship.target_mode(),
    )?;
    state.current_xml = next_xml;
    pending_parts.push(SvgPartChange {
        action: GraphAction::Add,
        part: Box::new(BlobPart::new_shared(
            part_uri.clone(),
            "image/svg+xml".to_owned(),
            Arc::clone(
                attach
                    .payload
                    .as_ref()
                    .ok_or_else(|| invalid("attach_svg has no staged payload"))?,
            ),
        )),
    });
    Ok(())
}

/// Plan a drawing group containing one attach/detach operation per picture.
/// Distinct selectors address disjoint source elements, so all structural
/// changes can be published by one drawing XML splice. Repeated selectors
/// are coalesced to their final projected owner state; the one-selector
/// replacement form has a dedicated bounded rewrite and larger replacement
/// batches are refused before any per-operation materialization.
fn plan_composed_mixed_batch(
    workbook: &Workbook,
    group: &[usize],
    intents: &[SvgLifecycleIntent],
    state: &mut DrawingPlanState,
    pending_parts: &mut Vec<SvgPartChange>,
    reserved_uris: &mut HashSet<PackURI>,
    occupied_part_names: &mut HashSet<String>,
    occupied_part_prefixes: &mut HashSet<String>,
    next_part_index: &mut usize,
    _manifest: &mut OwnedContentTypes,
    cleanup_targets: &mut CleanupTargets,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let first = intents
        .get(
            *group
                .first()
                .ok_or_else(|| invalid("empty SVG lifecycle group"))?,
        )
        .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
    let scanned = SourceDrawing::scan_with_limits(
        state.before_xml.bytes(),
        selector_drawing(first.selector),
        limits,
    )?;
    let pending_media_index = index_pending_media_parts(pending_parts)?;
    let relationship_type = if scanned.relationship_dialect() == source::RelationshipDialect::Strict
    {
        rt::STRICT_IMAGE
    } else {
        rt::IMAGE
    };
    let mut planning_edges = clone_relationships_bounded(&state.current_edges)?;
    let mut attachments = Vec::new();
    attachments
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle mixed attachments", source))?;
    let mut additions: Vec<(String, String, String, TargetMode)> = Vec::new();
    additions
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle mixed relationships", source))?;
    let mut detach_operations = Vec::new();
    detach_operations
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle mixed detachments", source))?;

    for index in group {
        let intent = intents
            .get(*index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let picture = scanned.picture(selector_picture(intent.selector))?;
        validate_raster_fallback_edges(workbook, &planning_edges, picture)?;
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
            let relationship_id = state.relationship_ids.allocate(&planning_edges)?;
            let part_uri = allocate_part_uri_with_reservations(
                reserved_uris,
                occupied_part_names,
                occupied_part_prefixes,
                next_part_index,
            )?;
            reserved_uris.insert(part_uri.clone());
            index_part_name(&part_uri, occupied_part_names, occupied_part_prefixes);
            let target_ref = part_uri.relative_ref(state.uri.base_uri());
            if planning_edges.len()
                >= workbook
                    .inner
                    .package
                    .read_limits()
                    .max_relationships_per_part()
            {
                return Err(invalid("SVG relationship count exceeds caller limit"));
            }
            planning_edges.try_add_relationship(
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
            let extension_len =
                generated_extension_size(blip.prefix(), &relationship_id, scanned.dialect());
            let (parent, parent_local, fragment_len) =
                if let Some(ext_list) = picture.ext_list_range() {
                    (
                        ext_list.range().start..ext_list.start_end(),
                        "extLst",
                        extension_len,
                    )
                } else {
                    (
                        blip.range().start..blip.start_end(),
                        "blip",
                        generated_ext_list_size(extension_prefix, extension_len, scanned.dialect()),
                    )
                };
            attachments.push(MixedAttachOperation {
                intent_index: *index,
                relationship_id,
                part_uri: part_uri.clone(),
                parent: parent.clone(),
                parent_empty: if let Some(ext_list) = picture.ext_list_range() {
                    ext_list.is_empty()
                } else {
                    blip.is_empty()
                },
                parent_name_len: element_name_len(
                    if let Some(ext_list) = picture.ext_list_range() {
                        ext_list.prefix()
                    } else {
                        blip.prefix()
                    },
                    parent_local,
                ),
                fragment_len,
            });
        } else {
            let owner = match picture.svg_owner() {
                SvgOwnerState::Embedded(owner) => owner,
                SvgOwnerState::None | SvgOwnerState::Opaque => {
                    return Err(invalid("selected picture has no admitted SVG owner"));
                },
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
            // Detach references the source owner captured by the scanner.  An
            // earlier attach in this same mixed batch may reuse the freed ID,
            // so resolving through the evolving edge map could bind the
            // detach to the new target instead of the source relationship.
            let relationship = state
                .current_edges
                .get(relationship_id)
                .cloned()
                .ok_or_else(|| invalid("embedded SVG relationship is missing"))?;
            let target = relationship.target_partname()?;
            validate_svg_target_candidate(
                workbook,
                &relationship,
                &target,
                pending_parts,
                &pending_media_index,
            )?;
            let (opening, removed) = detach_ranges(state.before_xml.bytes(), picture, owner)?;
            detach_operations.push(DetachOperation {
                relationship_id: relationship_id.to_owned(),
                target,
                opening,
                removed,
                owner_range: owner.extension_range(),
                ext_list: picture
                    .ext_list_range()
                    .map(|range| (range.range(), range.start_end())),
            });
            planning_edges.remove(relationship_id);
        }
    }

    let mut xml_removals = Vec::new();
    xml_removals
        .try_reserve(detach_operations.len())
        .map_err(|source| allocation("SVG lifecycle mixed XML ranges", source))?;
    for operation in &detach_operations {
        xml_removals.push(XmlRemoval {
            opening: operation.opening.clone(),
            removed: operation.removed,
        });
    }
    let mut ext_lists = Vec::<ExtListRemoval>::new();
    ext_lists
        .try_reserve(detach_operations.len())
        .map_err(|source| allocation("SVG lifecycle mixed extension lists", source))?;
    let mut ext_list_index = HashMap::<source::ByteRange, usize>::new();
    ext_list_index
        .try_reserve(detach_operations.len())
        .map_err(|source| allocation("SVG lifecycle mixed extension list index", source))?;
    for operation in &detach_operations {
        let Some((range, start_end)) = operation.ext_list else {
            continue;
        };
        if let Some(existing_index) = ext_list_index.get(&range).copied() {
            let existing = ext_lists
                .get_mut(existing_index)
                .ok_or_else(|| invalid("SVG lifecycle extension list index disappeared"))?;
            existing
                .owner_ranges
                .try_reserve(1)
                .map_err(|source| allocation("SVG lifecycle mixed extension owners", source))?;
            existing.owner_ranges.push(operation.owner_range);
        } else {
            let mut owner_ranges = Vec::new();
            owner_ranges
                .try_reserve(1)
                .map_err(|source| allocation("SVG lifecycle mixed extension owners", source))?;
            owner_ranges.push(operation.owner_range);
            ext_lists.push(ExtListRemoval {
                range,
                start_end,
                owner_ranges,
            });
            ext_list_index.insert(range, ext_lists.len() - 1);
        }
    }
    coalesce_extension_list_removals(state.before_xml.bytes(), &mut xml_removals, &ext_lists)?;

    let references = scanned.relationship_references();
    let mut removed_relationships = Vec::new();
    removed_relationships
        .try_reserve(detach_operations.len())
        .map_err(|source| allocation("SVG lifecycle mixed relationship IDs", source))?;
    let mut removed_relationship_ids = HashSet::new();
    removed_relationship_ids
        .try_reserve(detach_operations.len())
        .map_err(|source| allocation("SVG lifecycle mixed relationship IDs", source))?;
    let mut removed_ranges: Vec<_> = xml_removals.iter().map(|removal| removal.removed).collect();
    removed_ranges.sort_by_key(|range| range.start);
    let mut remaining_references = HashMap::<&str, usize>::new();
    remaining_references
        .try_reserve(references.len())
        .map_err(|source| {
            allocation("SVG lifecycle mixed relationship reference census", source)
        })?;
    let mut removed_range_index = 0usize;
    for reference in references {
        while removed_range_index < removed_ranges.len()
            && removed_ranges[removed_range_index].end <= reference.range().start
        {
            removed_range_index += 1;
        }
        let removed = removed_range_index < removed_ranges.len()
            && ranges_overlap(reference.range(), removed_ranges[removed_range_index]);
        if !removed {
            let entry = remaining_references.entry(reference.id()).or_insert(0);
            *entry = entry
                .checked_add(1)
                .ok_or_else(|| invalid("SVG lifecycle relationship reference count overflow"))?;
        }
    }
    for operation in &detach_operations {
        let remaining = remaining_references
            .get(operation.relationship_id.as_str())
            .copied()
            .unwrap_or(0);
        let final_relationship = planning_edges.get(&operation.relationship_id);
        let source_relationship = state.current_edges.get(&operation.relationship_id);
        let relationship_replaced = final_relationship.is_some()
            && source_relationship.is_some()
            && !same_relationship_option(source_relationship, final_relationship);
        if remaining != 0 && relationship_replaced {
            return Err(invalid(
                "SVG relationship ID is still referenced by opaque drawing XML",
            ));
        }
        if remaining == 0
            && (final_relationship.is_none() || relationship_replaced)
            && !removed_relationship_ids.contains(&operation.relationship_id)
        {
            removed_relationship_ids.insert(operation.relationship_id.clone());
            removed_relationships.push(operation.relationship_id.clone());
            cleanup_targets.push(operation.target.clone())?;
        }
    }

    let mut append_targets = Vec::new();
    append_targets
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle mixed XML updates", source))?;
    for (fragment_index, operation) in attachments.iter().enumerate() {
        append_targets.push(XmlAppendTarget {
            range: operation.parent.clone(),
            parent_empty: operation.parent_empty,
            parent_name_len: operation.parent_name_len,
            fragment_index,
        });
    }
    let mut output_size = state.before_xml.bytes().len();
    for removal in &xml_removals {
        output_size = output_size
            .checked_sub(removal.removed.len()?)
            .ok_or_else(|| invalid("SVG lifecycle mixed XML ranges overlap"))?;
    }
    let mut expanded_parents = HashSet::<(usize, usize)>::new();
    expanded_parents
        .try_reserve(append_targets.len())
        .map_err(|source| allocation("SVG lifecycle mixed XML parents", source))?;
    for target in &append_targets {
        output_size = output_size
            .checked_add(attachments[target.fragment_index].fragment_len)
            .ok_or_else(|| invalid("SVG lifecycle mixed XML output size overflow"))?;
        if target.parent_empty && expanded_parents.insert((target.range.start, target.range.end)) {
            output_size = output_size
                .checked_add(target.parent_name_len.saturating_add(2))
                .ok_or_else(|| invalid("SVG lifecycle mixed XML output size overflow"))?;
        }
    }
    if output_size > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle drawing output exceeds {} bytes",
            output_cap
        )));
    }
    let mut fragments = Vec::new();
    fragments
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle mixed fragments", source))?;
    pending_parts
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle mixed media", source))?;
    for operation in &attachments {
        let intent = intents
            .get(operation.intent_index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let picture = scanned.picture(selector_picture(intent.selector))?;
        let blip = picture.blip_range();
        let extension =
            generated_extension(blip.prefix(), &operation.relationship_id, scanned.dialect())?;
        let fragment = if picture.ext_list_range().is_some() {
            extension
        } else {
            generated_ext_list(
                picture
                    .ext_list_range()
                    .map_or(blip.prefix(), ElementRange::prefix),
                &extension,
                scanned.dialect(),
            )?
        };
        if fragment.len() != operation.fragment_len {
            return Err(invalid("SVG lifecycle extension length preflight mismatch"));
        }
        fragments.push(fragment);
        pending_parts.push(SvgPartChange {
            action: GraphAction::Add,
            part: Box::new(BlobPart::new_shared(
                operation.part_uri.clone(),
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
    let mut updates = Vec::new();
    updates
        .try_reserve(xml_removals.len().saturating_add(append_targets.len()))
        .map_err(|source| allocation("SVG lifecycle mixed XML updates", source))?;
    for removal in &xml_removals {
        updates.push(OwnedElementUpdate {
            start_tag: removal.opening.clone(),
            edit: OwnedElementEdit::Remove,
        });
    }
    for target in &append_targets {
        updates.push(OwnedElementUpdate {
            start_tag: target.range.clone(),
            edit: OwnedElementEdit::AppendChild(&fragments[target.fragment_index]),
        });
    }
    updates.sort_by(|left, right| {
        left.start_tag
            .start
            .cmp(&right.start_tag.start)
            .then(left.start_tag.end.cmp(&right.start_tag.end))
    });
    for pair in updates.windows(2) {
        if pair[0].start_tag.end > pair[1].start_tag.start {
            return Err(invalid("SVG lifecycle mixed XML ranges overlap"));
        }
    }
    state.current_xml = state.before_xml.update_elements(&updates, output_cap)?;
    state.current_edges = planning_edges;
    Ok(())
}

fn plan_composed_attach_batch(
    workbook: &Workbook,
    group: &[usize],
    intents: &[SvgLifecycleIntent],
    state: &mut DrawingPlanState,
    pending_parts: &mut Vec<SvgPartChange>,
    reserved_uris: &mut HashSet<PackURI>,
    occupied_part_names: &mut HashSet<String>,
    occupied_part_prefixes: &mut HashSet<String>,
    next_part_index: &mut usize,
    _manifest: &mut OwnedContentTypes,
) -> Result<()> {
    let limits = scan_limits(workbook);
    let output_cap = output_limit(workbook);
    let first_intent = intents
        .get(
            *group
                .first()
                .ok_or_else(|| invalid("empty SVG lifecycle group"))?,
        )
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
    let mut attachments = Vec::new();
    attachments
        .try_reserve(group.len())
        .map_err(|source| allocation("SVG lifecycle XML update batch", source))?;

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
        let relationship_id = state.relationship_ids.allocate(&state.current_edges)?;
        let part_uri = allocate_part_uri_with_reservations(
            reserved_uris,
            occupied_part_names,
            occupied_part_prefixes,
            next_part_index,
        )?;
        reserved_uris.insert(part_uri.clone());
        index_part_name(&part_uri, occupied_part_names, occupied_part_prefixes);
        let target_ref = part_uri.relative_ref(state.uri.base_uri());
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
        let extension_len =
            generated_extension_size(blip.prefix(), &relationship_id, scanned.dialect());
        let (parent, parent_local, fragment_len) = if let Some(ext_list) = picture.ext_list_range()
        {
            (ext_list, "extLst", extension_len)
        } else {
            (
                blip,
                "blip",
                generated_ext_list_size(extension_prefix, extension_len, scanned.dialect()),
            )
        };
        attachments.push(MixedAttachOperation {
            intent_index: *index,
            relationship_id,
            part_uri: part_uri.clone(),
            parent: parent.range().start..parent.start_end(),
            parent_empty: parent.is_empty(),
            parent_name_len: element_name_len(parent.prefix(), parent_local),
            fragment_len,
        });
    }

    let max_relationships = workbook
        .inner
        .package
        .read_limits()
        .max_relationships_per_part();
    let next_relationship_count = state
        .current_edges
        .len()
        .checked_add(additions.len())
        .ok_or_else(|| invalid("SVG relationship count overflows"))?;
    if next_relationship_count > max_relationships {
        return Err(invalid("SVG relationship count exceeds caller limit"));
    }

    let mut output_size = state.before_xml.bytes().len();
    let mut expanded_parents = HashSet::<(usize, usize)>::new();
    expanded_parents
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle empty XML parents", source))?;
    for target in &attachments {
        output_size = output_size
            .checked_add(target.fragment_len)
            .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
        if target.parent_empty && expanded_parents.insert((target.parent.start, target.parent.end))
        {
            output_size = output_size
                .checked_add(target.parent_name_len.saturating_add(2))
                .ok_or_else(|| invalid("SVG lifecycle drawing output size overflow"))?;
        }
    }
    if output_size > output_cap {
        return Err(invalid(format!(
            "SVG lifecycle drawing output exceeds {} bytes",
            output_cap
        )));
    }

    let mut fragments = Vec::new();
    fragments
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle extension batch", source))?;
    for operation in &attachments {
        let intent = intents
            .get(operation.intent_index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        let picture = scanned.picture(selector_picture(intent.selector))?;
        let blip = picture.blip_range();
        let extension =
            generated_extension(blip.prefix(), &operation.relationship_id, scanned.dialect())?;
        let fragment = if picture.ext_list_range().is_some() {
            extension
        } else {
            generated_ext_list(
                picture
                    .ext_list_range()
                    .map_or(blip.prefix(), ElementRange::prefix),
                &extension,
                scanned.dialect(),
            )?
        };
        if fragment.len() != operation.fragment_len {
            return Err(invalid("SVG lifecycle extension length preflight mismatch"));
        }
        fragments.push(fragment);
    }

    let mut append_targets = Vec::new();
    append_targets
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle XML update batch", source))?;
    for (fragment_index, operation) in attachments.iter().enumerate() {
        append_targets.push(XmlAppendTarget {
            range: operation.parent.clone(),
            parent_empty: operation.parent_empty,
            parent_name_len: operation.parent_name_len,
            fragment_index,
        });
    }
    append_targets.sort_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then(left.range.end.cmp(&right.range.end))
            .then(left.fragment_index.cmp(&right.fragment_index))
    });
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
    let mut next_edges = clone_relationships_bounded(&state.current_edges)?;
    for (reltype, target, id, mode) in &additions {
        next_edges.try_add_relationship(reltype.clone(), target.clone(), id.clone(), *mode)?;
    }
    pending_parts
        .try_reserve(attachments.len())
        .map_err(|source| allocation("SVG lifecycle media batch", source))?;
    for operation in &attachments {
        let intent = intents
            .get(operation.intent_index)
            .ok_or_else(|| invalid("SVG lifecycle intent disappeared"))?;
        pending_parts.push(SvgPartChange {
            action: GraphAction::Add,
            part: Box::new(BlobPart::new_shared(
                operation.part_uri.clone(),
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
    state.current_xml = next_xml;
    state.current_edges = next_edges;
    Ok(())
}

struct XmlAppendTarget {
    range: std::ops::Range<usize>,
    parent_empty: bool,
    parent_name_len: usize,
    fragment_index: usize,
}

fn element_name_len(prefix: &[u8], local: &str) -> usize {
    local.len().saturating_add(if prefix.is_empty() {
        0
    } else {
        prefix.len() + 1
    })
}

fn finalize_svg_target_cleanup(
    workbook: &Workbook,
    states: &[DrawingPlanState],
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
    pending_parts: &mut Vec<SvgPartChange>,
    cleanup_targets: &CleanupTargets,
) -> Result<()> {
    let incoming = incoming_relationship_targets_with_states(
        workbook,
        &cleanup_targets.values,
        states,
        ordinary_graph,
        relationship_changes,
    )?;
    // Index staged additions once.  Cleanup targets can be shared by many
    // pictures; searching/removing from the pending vector for every target
    // made a large detach batch quadratic and also made the final topology
    // depend on vector-removal order.
    let mut pending_add_keys = HashSet::new();
    pending_add_keys
        .try_reserve(pending_parts.len())
        .map_err(|source| allocation("SVG lifecycle pending media index", source))?;
    for change in pending_parts.iter() {
        if change.action == GraphAction::Add {
            pending_add_keys.insert(normalized_uri_key(change.part.partname()));
        }
    }
    let mut cancelled_add_keys = HashSet::new();
    cancelled_add_keys
        .try_reserve(cleanup_targets.len())
        .map_err(|source| allocation("SVG lifecycle cancelled media index", source))?;
    pending_parts
        .try_reserve(cleanup_targets.len())
        .map_err(|source| allocation("SVG lifecycle cleanup media changes", source))?;
    for target in &cleanup_targets.values {
        if incoming.contains(&target.as_str().to_ascii_lowercase()) {
            continue;
        }
        let target_key = normalized_uri_key(target);
        if pending_add_keys.contains(&target_key) {
            cancelled_add_keys.insert(target_key);
        } else {
            let part = workbook.inner.package.get_part(target)?;
            let part_uri = part.partname().clone();
            pending_parts.push(SvgPartChange {
                action: GraphAction::Remove,
                part: Box::new(BlobPart::new_shared(
                    part_uri.clone(),
                    part.content_type().to_owned(),
                    part.blob_arc(),
                )),
            });
            continue;
        }
    }
    if !cancelled_add_keys.is_empty() {
        pending_parts.retain(|change| {
            change.action != GraphAction::Add
                || !cancelled_add_keys.contains(&normalized_uri_key(change.part.partname()))
        });
    }
    Ok(())
}

fn changed_relationship_ids(before: &Relationships, after: &Relationships) -> Result<Vec<String>> {
    let capacity = before
        .len()
        .checked_add(after.len())
        .ok_or_else(|| invalid("relationship ID count overflows"))?;
    let mut ids = Vec::new();
    ids.try_reserve(capacity)
        .map_err(|source| allocation("SVG lifecycle relationship IDs", source))?;
    ids.extend(
        before
            .iter()
            .map(|relationship| relationship.r_id().to_owned()),
    );
    ids.extend(
        after
            .iter()
            .map(|relationship| relationship.r_id().to_owned()),
    );
    ids.sort_unstable();
    ids.dedup();
    Ok(ids
        .into_iter()
        .filter(|id| !same_relationship_option(before.get(id), after.get(id)))
        .collect())
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
    if !is_media_uri(&target) {
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
    pending_media_index: &HashMap<String, usize>,
) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
        || !is_media_uri(target)
    {
        return Err(invalid("SVG relationship target is unsupported"));
    }
    let pending_part = pending_media_index
        .get(&normalized_uri_key(target))
        .and_then(|index| pending_parts.get(*index))
        .filter(|change| change.action == GraphAction::Add)
        .map(|change| change.part.as_ref());
    if let Some(part) = pending_part {
        if part.content_type() != "image/svg+xml" || !part.rels().is_empty() {
            return Err(invalid("SVG target must be an inert image/svg+xml part"));
        }
        return Ok(());
    }
    validate_svg_target(workbook, relationship, target)
}

fn index_pending_media_parts(pending_parts: &[SvgPartChange]) -> Result<HashMap<String, usize>> {
    let mut index = HashMap::new();
    index
        .try_reserve(pending_parts.len())
        .map_err(|source| allocation("SVG lifecycle pending media index", source))?;
    for (position, change) in pending_parts.iter().enumerate() {
        if change.action != GraphAction::Add {
            continue;
        }
        let key = normalized_uri_key(change.part.partname());
        if let Entry::Vacant(entry) = index.entry(key) {
            entry.insert(position);
        }
    }
    Ok(index)
}

fn record_incoming_target(
    workbook: &Workbook,
    target_set: &HashSet<String>,
    incoming: &mut HashSet<String>,
    relationship: &Relationship,
) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal {
        return Ok(());
    }
    let Ok(target) = relationship.target_partname() else {
        return Ok(());
    };
    let canonical = workbook
        .inner
        .package
        .get_part(&target)
        .map(|part| part.partname().clone())
        .unwrap_or(target);
    let key = canonical.as_str().to_ascii_lowercase();
    if target_set.contains(&key) {
        incoming.insert(key);
    }
    Ok(())
}

fn visit_incoming_relationships(
    workbook: &Workbook,
    target_set: &HashSet<String>,
    incoming: &mut HashSet<String>,
    changed_relationships: &HashMap<String, HashSet<String>>,
    relationship_overrides: &HashMap<String, HashMap<String, Option<Relationship>>>,
    owner: &str,
    relationships: &Relationships,
) -> Result<()> {
    let owner_changes = changed_relationships.get(owner);
    let owner_overrides = relationship_overrides.get(owner);
    for relationship in relationships.iter() {
        if owner_changes.is_some_and(|ids| ids.contains(relationship.r_id())) {
            continue;
        }
        if let Some(overrides) =
            owner_overrides.and_then(|overrides| overrides.get(relationship.r_id()))
        {
            if let Some(replacement) = overrides {
                record_incoming_target(workbook, target_set, incoming, replacement)?;
            }
        } else {
            record_incoming_target(workbook, target_set, incoming, relationship)?;
        }
    }
    Ok(())
}

fn incoming_relationship_targets_with_states(
    workbook: &Workbook,
    targets: &[PackURI],
    states: &[DrawingPlanState],
    ordinary_graph: &[super::model::GraphChange],
    relationship_changes: &[RelationshipChange],
) -> Result<HashSet<String>> {
    let mut target_set = HashSet::new();
    target_set
        .try_reserve(targets.len())
        .map_err(|source| allocation("SVG lifecycle cleanup targets", source))?;
    target_set.extend(
        targets
            .iter()
            .map(|target| target.as_str().to_ascii_lowercase()),
    );
    let mut incoming = HashSet::new();
    incoming
        .try_reserve(targets.len())
        .map_err(|source| allocation("SVG lifecycle incoming targets", source))?;
    let mut changed_relationships: HashMap<String, HashSet<String>> = HashMap::new();
    changed_relationships
        .try_reserve(ordinary_graph.len())
        .map_err(|source| allocation("SVG lifecycle changed relationship index", source))?;
    for change in ordinary_graph {
        if change.action != GraphAction::Remove {
            continue;
        }
        let owner = change.source.as_str().to_ascii_lowercase();
        changed_relationships
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle changed relationship index", source))?;
        let ids = changed_relationships.entry(owner).or_default();
        ids.try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle changed relationship IDs", source))?;
        ids.insert(change.relationship.r_id().to_owned());
    }
    let mut removed_part_owners = HashSet::new();
    removed_part_owners
        .try_reserve(ordinary_graph.len())
        .map_err(|source| allocation("SVG lifecycle removed graph owners", source))?;
    for change in ordinary_graph {
        if change.action == GraphAction::Remove {
            removed_part_owners.insert(change.part.partname().as_str().to_ascii_lowercase());
        }
    }
    let mut state_relationships: HashMap<String, &Relationships> = HashMap::new();
    state_relationships
        .try_reserve(states.len())
        .map_err(|source| allocation("SVG lifecycle drawing relationship index", source))?;
    for state in states {
        state_relationships.insert(
            state.uri.as_str().to_ascii_lowercase(),
            &state.current_edges,
        );
    }
    let mut relationship_overrides: HashMap<String, HashMap<String, Option<Relationship>>> =
        HashMap::new();
    relationship_overrides
        .try_reserve(relationship_changes.len())
        .map_err(|source| allocation("SVG lifecycle relationship transition index", source))?;
    for change in relationship_changes {
        let owner = change.owner.as_str().to_ascii_lowercase();
        relationship_overrides
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle relationship transition index", source))?;
        let overrides = relationship_overrides.entry(owner).or_default();
        overrides
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle relationship transition IDs", source))?;
        let id = change
            .before
            .as_ref()
            .or(change.after.as_ref())
            .map(|relationship| relationship.r_id().to_owned())
            .ok_or_else(|| invalid("SVG relationship transition has no relationship ID"))?;
        overrides.insert(id, change.after.clone());
    }
    visit_incoming_relationships(
        workbook,
        &target_set,
        &mut incoming,
        &changed_relationships,
        &relationship_overrides,
        "/",
        workbook.inner.package.rels(),
    )?;
    for part in workbook.inner.package.iter_parts() {
        let owner = part.partname().as_str().to_ascii_lowercase();
        if removed_part_owners.contains(&owner) {
            continue;
        }
        let relationships = state_relationships
            .get(&owner)
            .copied()
            .unwrap_or_else(|| part.rels());
        visit_incoming_relationships(
            workbook,
            &target_set,
            &mut incoming,
            &changed_relationships,
            &relationship_overrides,
            &owner,
            relationships,
        )?;
    }
    // A relationship transition may add a previously absent ID, and an
    // ordinary graph addition may carry opaque child edges of its new part.
    for change in relationship_changes {
        if let Some(relationship) = &change.after {
            record_incoming_target(workbook, &target_set, &mut incoming, relationship)?;
        }
    }
    for change in ordinary_graph {
        if change.action == GraphAction::Add {
            record_incoming_target(workbook, &target_set, &mut incoming, &change.relationship)?;
            for relationship in change.part.rels().iter() {
                record_incoming_target(workbook, &target_set, &mut incoming, relationship)?;
            }
        }
    }
    Ok(incoming)
}

fn allocate_part_uri_with_reservations(
    reserved: &HashSet<PackURI>,
    occupied_part_names: &HashSet<String>,
    occupied_part_prefixes: &HashSet<String>,
    next_index: &mut usize,
) -> Result<PackURI> {
    let start = *next_index;
    for offset in 0..MAX_ID_ATTEMPTS {
        let index = start
            .checked_add(offset)
            .ok_or_else(|| invalid("generated SVG media Part index overflows"))?;
        let name = if index == 0 {
            "/xl/media/vector.svg".to_owned()
        } else {
            format!("/xl/media/vector{index}.svg")
        };
        let uri = PackURI::new(name).map_err(litchi_opc::OpcError::InvalidPackUri)?;
        if !reserved.contains(&uri)
            && !part_name_conflicts(&uri, occupied_part_names, occupied_part_prefixes)
        {
            *next_index = index
                .checked_add(1)
                .ok_or_else(|| invalid("generated SVG media Part index overflows"))?;
            return Ok(uri);
        }
    }
    Err(invalid(format!(
        "generated SVG media Part candidates exceed {}",
        MAX_ID_ATTEMPTS
    )))
}

fn part_name_conflicts(
    candidate: &PackURI,
    occupied_part_names: &HashSet<String>,
    occupied_part_prefixes: &HashSet<String>,
) -> bool {
    let name = candidate.as_str().to_ascii_lowercase();
    if occupied_part_names.contains(&name) || occupied_part_prefixes.contains(&name) {
        return true;
    }
    let mut prefix = String::new();
    let mut segments = name.split('/').skip(1).peekable();
    while let Some(segment) = segments.next() {
        prefix.push('/');
        prefix.push_str(segment);
        if segments.peek().is_some() && occupied_part_names.contains(&prefix) {
            return true;
        }
    }
    false
}

fn index_part_name(
    part: &PackURI,
    occupied_part_names: &mut HashSet<String>,
    occupied_part_prefixes: &mut HashSet<String>,
) {
    let name = part.as_str().to_ascii_lowercase();
    occupied_part_names.insert(name.clone());
    let mut prefix = String::new();
    let mut segments = name.split('/').skip(1).peekable();
    while let Some(segment) = segments.next() {
        prefix.push('/');
        prefix.push_str(segment);
        if segments.peek().is_some() {
            occupied_part_prefixes.insert(prefix.clone());
        }
    }
}

fn materialize_content_types_transition(
    before: &OwnedContentTypes,
    removals: &[PackURI],
    additions: &[PackURI],
    limits: litchi_opc::ReadLimits,
) -> Result<OwnedContentTypes> {
    let (has_default_svg, existing_svg_overrides) = manifest_svg_inventory(before.bytes())?;
    let mut selected_removals = Vec::new();
    selected_removals
        .try_reserve(removals.len())
        .map_err(|source| allocation("SVG lifecycle content-type removals", source))?;
    let mut removal_keys = HashSet::new();
    removal_keys
        .try_reserve(removals.len())
        .map_err(|source| allocation("SVG lifecycle content-type removal index", source))?;
    for uri in removals {
        let key = normalized_uri_key(uri);
        if existing_svg_overrides.contains(&key) && removal_keys.insert(key) {
            selected_removals.push(uri.clone());
        }
    }
    let mut selected_additions = Vec::new();
    if !has_default_svg {
        selected_additions
            .try_reserve(additions.len())
            .map_err(|source| allocation("SVG lifecycle content-type additions", source))?;
        let mut addition_keys = HashSet::new();
        addition_keys
            .try_reserve(additions.len())
            .map_err(|source| allocation("SVG lifecycle content-type addition index", source))?;
        for uri in additions {
            let key = normalized_uri_key(uri);
            if !existing_svg_overrides.contains(&key) && addition_keys.insert(key) {
                selected_additions.push(ContentTypeEdit {
                    part_name: uri,
                    content_type: "image/svg+xml",
                });
            }
        }
    }
    if selected_removals.is_empty() && selected_additions.is_empty() {
        return Ok(before.clone());
    }
    let plan = before.plan_edit(&selected_additions, &selected_removals, limits)?;
    let output_cap = usize::try_from(limits.max_part_bytes())
        .unwrap_or(usize::MAX)
        .min(MAX_OUTPUT_BYTES)
        .min(limits.max_content_types_bytes());
    if plan.final_len() > output_cap {
        return Err(invalid(
            "SVG lifecycle content-types output exceeds caller limit",
        ));
    }
    plan.materialize(limits).map_err(Into::into)
}

/// Scan the retained manifest once to decide which SVG overrides already
/// exist.  The exact source-backed OPC plan performs the final range scan and
/// byte admission; this small inventory only avoids one XML scan per cleanup
/// or media target while keeping Default coverage distinct from Override
/// selectors.
fn manifest_svg_inventory(bytes: &[u8]) -> Result<(bool, HashSet<String>)> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut has_default_svg = false;
    let mut overrides = HashSet::new();
    overrides
        .try_reserve(8)
        .map_err(|source| allocation("SVG lifecycle content-type inventory", source))?;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?
        {
            Event::Empty(element) | Event::Start(element)
                if element.local_name().as_ref() == b"Default"
                    || element.local_name().as_ref() == b"Override" =>
            {
                let mut extension = None;
                let mut part_name = None;
                let mut content_type = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| invalid(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Extension" => extension = Some(value),
                        b"PartName" => part_name = Some(value),
                        b"ContentType" => content_type = Some(value),
                        _ => {},
                    }
                }
                if content_type.as_deref() == Some("image/svg+xml") {
                    if element.local_name().as_ref() == b"Default"
                        && extension
                            .as_deref()
                            .is_some_and(|value| value.eq_ignore_ascii_case("svg"))
                    {
                        has_default_svg = true;
                    } else if element.local_name().as_ref() == b"Override" {
                        let Some(part_name) = part_name else {
                            return Err(invalid(
                                "SVG content-type Override has no PartName attribute",
                            ));
                        };
                        let part_name = PackURI::new(part_name).map_err(invalid)?;
                        overrides.insert(normalized_uri_key(&part_name));
                    }
                }
            },
            Event::Eof => return Ok((has_default_svg, overrides)),
            _ => {},
        }
        buffer.clear();
    }
}

fn relationship_xml_event_count(bytes: &[u8], maximum: usize) -> Result<usize> {
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut count = 0usize;
    loop {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("SVG lifecycle relationship XML event count overflow"))?;
        if count > maximum {
            return Err(invalid(format!(
                "SVG lifecycle relationship XML events exceed {maximum}"
            )));
        }
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        if matches!(event, Event::Eof) {
            return Ok(count);
        }
        buffer.clear();
    }
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
    let limits = crate::drawing::worksheet_source::WorksheetSourceLimits::from_read_limits(
        workbook.inner.package.read_limits(),
    );
    crate::drawing::worksheet_source::scan_worksheet_drawing_references(source, limits)
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
    if !is_media_uri(&target) {
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
        || !is_media_uri(target)
    {
        return Err(invalid("SVG relationship target is unsupported"));
    }
    let part = workbook.inner.package.get_part(target)?;
    if part.content_type() != "image/svg+xml" || !part.rels().is_empty() {
        return Err(invalid("SVG target must be an inert image/svg+xml part"));
    }
    Ok(())
}

fn detach_ranges(
    source: &[u8],
    picture: &source::PictureSource<'_>,
    owner: &source::SvgOwner<'_>,
) -> Result<(std::ops::Range<usize>, source::ByteRange)> {
    if let Some(ext_list) = picture.ext_list_range() {
        if ext_list_contains_only_owner(source, ext_list, owner)? {
            return Ok((
                ext_list.range().start..ext_list.start_end(),
                ext_list.range(),
            ));
        }
    }
    let extension = owner.extension_range();
    let opening = opening_tag_range(source, extension)?;
    Ok((opening, extension))
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
            Event::GeneralRef(_) if depth == 1 => return Ok(false),
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

fn ext_list_contains_only_owner_ranges(
    source: &[u8],
    ext_list: source::ByteRange,
    selected: &[source::ByteRange],
) -> Result<bool> {
    let bytes = ext_list.slice(source)?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut child_start = None;
    let mut children = Vec::new();
    children
        .try_reserve(selected.len())
        .map_err(|source| allocation("SVG lifecycle extension children", source))?;
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
                    if child_start.is_some() {
                        return Ok(false);
                    }
                    child_start = Some(ext_list.start.saturating_add(start));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("SVG extension depth overflows"))?;
            },
            Event::Empty(element) => {
                if depth == 0 && element.attributes().next().is_some() {
                    return Ok(false);
                }
                if depth == 1 {
                    children.push(source::ByteRange::new(
                        ext_list.start.saturating_add(start),
                        ext_list.start.saturating_add(end),
                    ));
                }
            },
            Event::End(_) => {
                if depth == 2 {
                    let child_start = child_start
                        .take()
                        .ok_or_else(|| invalid("SVG extension child has no opening tag"))?;
                    children.push(source::ByteRange::new(
                        child_start,
                        ext_list.start.saturating_add(end),
                    ));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("SVG extension depth underflows"))?;
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) if depth == 1 => {
                return Ok(false);
            },
            Event::GeneralRef(_) if depth == 1 => return Ok(false),
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
    if depth != 0 || children.len() != selected.len() {
        return Ok(false);
    }
    let mut selected_ranges = HashSet::<source::ByteRange>::new();
    selected_ranges
        .try_reserve(selected.len())
        .map_err(|source| allocation("SVG lifecycle selected extension ranges", source))?;
    for owner_range in selected {
        selected_ranges.insert(*owner_range);
    }
    Ok(children.iter().all(|child| selected_ranges.contains(child)))
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

fn generated_extension_size(
    prefix: &[u8],
    relationship_id: &str,
    drawing_dialect: source::DrawingDialect,
) -> usize {
    generated_extension_size_with_lengths(prefix.len(), relationship_id.len(), drawing_dialect)
}

fn generated_extension_size_with_lengths(
    prefix_len: usize,
    relationship_id_len: usize,
    drawing_dialect: source::DrawingDialect,
) -> usize {
    let ext = if prefix_len == 0 {
        "ext".len()
    } else {
        prefix_len.saturating_add(1).saturating_add("ext".len())
    };
    let drawing_binding = if prefix_len == 0 {
        0
    } else {
        " xmlns:"
            .len()
            .saturating_add(prefix_len)
            // Include both the `="` and the closing quote of the namespace
            // attribute.
            .saturating_add(3)
            .saturating_add(drawing_namespace(drawing_dialect).len())
    };
    let authored_namespace = source::RelationshipDialect::authored_svg_attribute_namespace();
    1usize
        .saturating_add(ext)
        .saturating_add(drawing_binding)
        .saturating_add(b" uri=\"".len())
        .saturating_add(SVG_EXTENSION_URI.len())
        .saturating_add(2)
        .saturating_add(b"<asvg:svgBlip xmlns:asvg=\"".len())
        .saturating_add(SVG_NAMESPACE.len())
        .saturating_add(b"\" xmlns:r=\"".len())
        .saturating_add(authored_namespace.len())
        .saturating_add(b"\" r:embed=\"".len())
        .saturating_add(relationship_id_len)
        .saturating_add(b"\"/>".len())
        .saturating_add(2)
        .saturating_add(ext)
        .saturating_add(1)
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

fn generated_ext_list_size(
    prefix: &[u8],
    extension_size: usize,
    drawing_dialect: source::DrawingDialect,
) -> usize {
    let prefix = String::from_utf8_lossy(prefix);
    let name = if prefix.is_empty() {
        "extLst".len()
    } else {
        prefix
            .len()
            .saturating_add(1)
            .saturating_add("extLst".len())
    };
    let drawing_binding = if prefix.is_empty() {
        0
    } else {
        " xmlns:"
            .len()
            .saturating_add(prefix.len())
            .saturating_add(3)
            .saturating_add(drawing_namespace(drawing_dialect).len())
    };
    1usize
        .saturating_add(name)
        .saturating_add(drawing_binding)
        .saturating_add(1)
        .saturating_add(extension_size)
        .saturating_add(2)
        .saturating_add(name)
        .saturating_add(1)
}

fn manifest_covers_svg(bytes: &[u8], uri: &PackURI) -> Result<bool> {
    if manifest_has_override(bytes, uri)? {
        return Ok(true);
    }
    manifest_has_default_svg(bytes)
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
                if content.as_deref() == Some("image/svg+xml") {
                    let Some(part) = part else {
                        continue;
                    };
                    let part = PackURI::new(part).map_err(litchi_opc::OpcError::InvalidPackUri)?;
                    if part.is_equivalent_to(uri) {
                        return Ok(true);
                    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::edit::SvgInput;
    use crate::workbook::edit::model::GraphChange;

    #[test]
    fn chained_graph_additions_measure_final_source_less_relationships() {
        let workbook = Workbook::new().unwrap();
        let limits = workbook.inner.package.read_limits();
        let root = PackURI::new("/").unwrap();
        let owner = PackURI::new("/xl/new-owner.xml").unwrap();
        let child = PackURI::new("/xl/new-child.xml").unwrap();
        let mut part = BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<owner/>".to_vec(),
        );
        part.rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/".to_owned(),
                "existing".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let before_plan = part.rels().plan_canonical(limits).unwrap();
        let before_bytes = before_plan.final_len();
        let before_events = before_plan.event_count();
        let mut final_relationships = part.rels().clone();
        let child_relationship = Relationship::new_with_mode(
            "child".to_owned(),
            "urn:test".to_owned(),
            child.relative_ref(owner.base_uri()),
            owner.base_uri().to_owned(),
            TargetMode::Internal,
        );
        insert_projected_relationship(&mut final_relationships, &child_relationship).unwrap();
        let after_plan = final_relationships.plan_canonical(limits).unwrap();
        let changes = vec![
            GraphChange {
                action: GraphAction::Add,
                source: root.clone(),
                relationship: Relationship::new_with_mode(
                    "new-owner".to_owned(),
                    "urn:test".to_owned(),
                    owner.relative_ref(root.base_uri()),
                    root.base_uri().to_owned(),
                    TargetMode::Internal,
                ),
                part: Box::new(part),
            },
            GraphChange {
                action: GraphAction::Add,
                source: owner,
                relationship: child_relationship,
                part: Box::new(BlobPart::new(
                    child,
                    "application/xml".to_owned(),
                    b"<child/>".to_vec(),
                )),
            },
        ];
        let budget = SvgDrawingBudget {
            added_bytes: 0,
            removed_bytes: 0,
            effective_attach_count: 0,
            effective_payload_bytes: 0,
            payload_lengths: Vec::new(),
            removed_edges: Vec::new(),
            relationship_additions: Vec::new(),
        };
        let measure = |changes: &[GraphChange]| {
            plan_final_relationship_topology(&workbook, changes, &[], &budget, &[], &[], limits)
                .unwrap()
        };
        let first = measure(&changes[..1]);
        let chained = measure(&changes);
        assert_eq!(chained.relationships, first.relationships + 1);
        assert_eq!(
            chained.bytes,
            first.bytes - before_bytes + after_plan.final_len()
        );
        assert_eq!(
            chained.events,
            first.events - before_events + after_plan.event_count()
        );
    }

    fn one_picture_workbook() -> Workbook {
        let baseline = Workbook::new().unwrap();
        let mut package = baseline.inner.package.clone();
        let worksheet_uri = baseline.inner.sheets[0].part_uri.clone();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec(),
            );
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
                ct::OFC_DRAWING.to_owned(),
                br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/media/image1.png").unwrap(),
                ct::PNG.to_owned(),
                b"picture-bytes".to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::DRAWING.to_owned(),
                "../drawings/drawing1.xml".to_owned(),
                "rIdDrawing".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/image1.png".to_owned(),
                "rIdRaster".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        Workbook::from_package(package).unwrap()
    }

    #[test]
    fn failed_attach_admission_leaves_real_preflight_retryable() {
        let mut edit = one_picture_workbook().edit().unwrap();
        let selector = PictureSelector::new(0, 0);
        let payload = SvgInput::borrowed(b"<svg/>");
        assert!(
            edit.stage_svg_attach_after_preflight_failure_for_test(0, selector, payload)
                .is_err()
        );
        assert!(edit.svg_lifecycle.is_empty());
        assert!(edit.svg_preflight.is_empty());
        assert!(edit.svg_projected_owners.is_empty());

        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .attach_svg(selector, payload)
            .unwrap();
        assert_eq!(edit.svg_lifecycle.len(), 1);
        assert_eq!(edit.svg_preflight.len(), 1);
        assert_eq!(
            edit.svg_projected_owners.get(&(0, selector)),
            Some(&ProjectedSvgOwner::Embedded)
        );
    }

    #[test]
    fn no_op_detach_retains_real_preflight_cache() {
        let mut edit = one_picture_workbook().edit().unwrap();
        let selector = PictureSelector::new(0, 0);
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .detach_svg(selector)
            .unwrap();
        assert!(edit.svg_lifecycle.is_empty());
        assert_eq!(edit.svg_preflight.len(), 1);
        assert_eq!(
            edit.svg_projected_owners.get(&(0, selector)),
            Some(&ProjectedSvgOwner::None)
        );
    }
}
