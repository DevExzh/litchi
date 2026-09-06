//! Private raw graph planning for selector-first Keynote media lifecycle edits.
//!
//! This module deliberately owns no public archive identifiers.  It resolves a
//! semantic media selector to an exact, bounded graph snapshot and gives the
//! transaction owner the spans it needs to append/remove references.  Payload
//! edits are performed against borrowed wire views; untouched source payloads
//! are never decoded into generated protobuf values.

use std::mem::size_of;

use litchi_core::Position;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{
    ArchiveObject, ArchiveReferenceOccurrence, ArchiveReferencePolicy, ArchiveReferenceVisitor,
    FieldType, Limits as ArchiveObjectLimits, RawMessage,
};

use super::budget::LifecycleBudget;
use super::comment_graph::{CommentGraphPlan, plan_comment_graph};
use super::{
    BUILD_CHUNK_MESSAGE_TYPE, BUILD_MESSAGE_TYPE, COMMENT_STORAGE_MESSAGE_TYPE, MOVIE_DATA_FIELD,
    MOVIE_MESSAGE_TYPE, POSTER_IMAGE_DATA_FIELD, Package, SLIDE_BUILD_CHUNKS_FIELD,
    SLIDE_BUILDS_FIELD, SLIDE_OWNED_DRAWABLES_FIELD, SlideMediaLifecycleError, SuperUuid,
};
use crate::{MovieKind, MovieSelector, SlideSelector};

const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// One source-order media selection plus its private graph closure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MediaGraphSelection {
    pub(super) slide_position: Position,
    pub(super) movie_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) component_name: Box<str>,
    pub(super) movie_identifier: u64,
    pub(super) kind: MovieKind,
    pub(super) content_identifier: Option<u64>,
    pub(super) poster_identifier: Option<u64>,
    pub(super) private_object_ids: Vec<u64>,
    pub(super) build_ids: Vec<u64>,
    pub(super) chunk_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
    pub(super) comment_graph: Option<CommentGraphPlan>,
}

/// Resolve a checked selector and validate the complete same-component graph.
pub(super) fn select_media(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<MediaGraphSelection, SlideMediaLifecycleError> {
    let slide_position = resolve_slide_position(package, slide_selector, budget)?;
    let slide_record = package
        .slide_record_at(slide_position.get())
        .map_err(|_| SlideMediaLifecycleError::Read)?
        .ok_or(SlideMediaLifecycleError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (component_name, slide) = package
        .object_with_component(slide_record.slide_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_fields(slide.messages.len())?;
    budget.charge_wire_work(slide.messages.len().max(1))?;
    let slide_payload = super::unique_payload(&slide.messages, super::SLIDE_MESSAGE_TYPE)
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let drawable_ids =
        references_in_field(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    let mut movies = Vec::new();
    reserve_u64s(&mut movies, drawable_ids.len(), budget)?;
    for drawable_identifier in drawable_ids {
        let Some((movie_component, movie)) = package.object_with_component(drawable_identifier)
        else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if movie_component != component_name {
            // A slide-owned drawable is a same-component invariant in the
            // native graph.  Skipping a foreign component would shift the
            // source-order selector and make a later edit target the wrong
            // semantic media item.
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        budget.charge_wire_fields(movie.messages.len())?;
        budget.charge_wire_work(movie.messages.len().max(1))?;
        let matches = movie
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .count();
        if matches > 1 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        if matches == 1 {
            push_u64(&mut movies, drawable_identifier, budget)?;
        }
    }
    let movie_position = movie_selector.as_position();
    let movie_identifier = *movies.get(movie_position.get()).ok_or(
        SlideMediaLifecycleError::MoviePositionNotFound {
            position: movie_position,
        },
    )?;
    let movie = package
        .object(movie_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_fields(movie.messages.len())?;
    budget.charge_wire_work(movie.messages.len().max(1))?;
    let movie_payload = movie
        .messages
        .iter()
        .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .next()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let (movie_info, _references) = super::super::decode_movie_info(
        movie_payload,
        limits,
        super::SemanticPath::SlideDrawable {
            slide: slide_position.get(),
            index: movie_position.get(),
        },
    )
    .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    if !matches!(movie_info.kind(), MovieKind::File | MovieKind::Audio) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let comment_graph = match direct_drawable_comment(movie_payload, limits, budget)? {
        Some(root) => {
            let message_index = movie
                .messages
                .iter()
                .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            let info = movie
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            budget.charge_references(info.object_references.len())?;
            if info
                .object_references
                .iter()
                .filter(|identifier| **identifier == root)
                .count()
                != 1
            {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            Some(plan_comment_graph(
                package,
                component_name,
                root,
                limits,
                budget,
            )?)
        },
        None => None,
    };
    let parent = drawable_parent(movie_payload, limits, budget)?;
    if parent != slide_record.slide_identifier {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let content_identifier =
        unique_data_identifier(movie_payload, MOVIE_DATA_FIELD, limits, budget)?;
    let poster_identifier =
        unique_data_identifier(movie_payload, POSTER_IMAGE_DATA_FIELD, limits, budget)?;
    if content_identifier.is_none()
        || (movie_info.kind() == MovieKind::File && poster_identifier.is_none())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let private_object_ids = private_graph(
        package,
        component_name,
        movie_identifier,
        slide_record.slide_identifier,
        comment_graph.as_ref(),
        budget,
    )?;
    if let Some(plan) = comment_graph.as_ref() {
        budget.charge_references(plan.storage_ids.len())?;
        if plan
            .storage_ids
            .iter()
            .any(|id| private_object_ids.binary_search(id).is_err())
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    let build_ids = selected_builds(
        package,
        component_name,
        slide_payload,
        movie_identifier,
        limits,
        budget,
    )?;
    let chunk_ids = selected_chunks(
        package,
        component_name,
        slide_payload,
        &build_ids,
        limits,
        budget,
    )?;
    let data_references = data_edges(
        package,
        component_name,
        &private_object_ids,
        &build_ids,
        &chunk_ids,
        budget,
    )?;
    let component_name = copy_boxed(component_name, budget)?;
    Ok(MediaGraphSelection {
        slide_position,
        movie_position,
        slide_identifier: slide_record.slide_identifier,
        component_name,
        movie_identifier,
        kind: movie_info.kind(),
        content_identifier,
        poster_identifier,
        private_object_ids,
        build_ids,
        chunk_ids,
        data_references,
        comment_graph,
    })
}

pub(super) fn direct_drawable_comment(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Option<u64>, SlideMediaLifecycleError> {
    let root = parse_view(payload, limits, budget, 1)?;
    let mut drawable = None;
    for field in root.fields().filter(|field| field.number() == 1) {
        if drawable.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        drawable = Some(field.payload());
    }
    let drawable = drawable.ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let drawable_view = parse_view(drawable, limits, budget, 2)?;
    let mut comment = None;
    for field in drawable_view.fields().filter(|field| field.number() == 6) {
        if comment.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        comment = Some(strict_reference_identifier(
            field.payload(),
            limits,
            budget,
            3,
        )?);
    }
    Ok(comment)
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut LifecycleBudget,
) -> Result<Position, SlideMediaLifecycleError> {
    budget.charge_wire_work(1)?;
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(|_| SlideMediaLifecycleError::Read)?
            .map(|_| position)
            .ok_or(SlideMediaLifecycleError::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideMediaLifecycleError::EmptySlideName);
            }
            let selected = package
                .show()
                .map_err(|_| SlideMediaLifecycleError::Read)?
                .select_slide(SlideSelector::name(name))
                .map_err(|_| SlideMediaLifecycleError::AmbiguousSelector)?;
            selected
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMediaLifecycleError::SlideNameNotFound)
        },
    }
}

fn references_in_field(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let view = parse_view(payload, limits, budget, 1)?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut references = Vec::new();
    reserve_u64s(&mut references, count, budget)?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        push_u64(
            &mut references,
            reference_identifier(field.payload(), limits, budget, 2)?,
            budget,
        )?;
    }
    Ok(references)
}

fn reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    nesting: usize,
) -> Result<u64, SlideMediaLifecycleError> {
    reference_identifier_with_policy(payload, limits, budget, nesting, false)
}

/// Parse the direct `MovieArchive.super.drawable.comment` reference.
///
/// The legacy reference envelope is intentionally stricter at this one
/// lifecycle admission boundary: deprecated fields and unknown fields cannot
/// silently change comment ownership while the broader compatibility readers
/// continue to preserve their historical behavior.
pub(super) fn strict_reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    nesting: usize,
) -> Result<u64, SlideMediaLifecycleError> {
    reference_identifier_with_policy(payload, limits, budget, nesting, true)
}

fn reference_identifier_with_policy(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    nesting: usize,
    reject_legacy_fields: bool,
) -> Result<u64, SlideMediaLifecycleError> {
    let view = parse_view(payload, limits, budget, nesting)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        if !matches!(field.number(), 1..=3) {
            if reject_legacy_fields {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            continue;
        }
        if field.wire_type() != 0 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        if width != field.payload().len() || litchi_iwa_common::varint::encoded_len(value) != width
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        match field.number() {
            1 => {
                if identifier.replace(value).is_some() || value == 0 {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
            },
            2 => {
                if reject_legacy_fields
                    || deprecated_type.replace(value).is_some()
                    || !canonical_int32(value)
                {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
            },
            3 => {
                if reject_legacy_fields || external.replace(value).is_some() || value > 1 {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
            },
            _ => unreachable!(),
        }
    }
    if external == Some(1) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let _ = deprecated_type;
    identifier.ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn unique_data_identifier(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Option<u64>, SlideMediaLifecycleError> {
    let view = parse_view(payload, limits, budget, 1)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        selected = Some(reference_identifier(field.payload(), limits, budget, 2)?);
    }
    Ok(selected)
}

fn drawable_parent(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<u64, SlideMediaLifecycleError> {
    let view = parse_view(payload, limits, budget, 1)?;
    let mut parent = None;
    for field in view.fields().filter(|field| field.number() == 1) {
        if parent.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        let drawable = parse_view(field.payload(), limits, budget, 2)?;
        for nested in drawable.fields().filter(|nested| nested.number() == 2) {
            if nested.wire_type() != 2 || parent.is_some() {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            nested
                .validate_canonical_framing()
                .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
            parent = Some(reference_identifier(nested.payload(), limits, budget, 3)?);
        }
    }
    parent.ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn private_graph(
    package: &Package,
    component_name: &str,
    root: u64,
    owning_slide: u64,
    comment_graph: Option<&CommentGraphPlan>,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let mut pending = Vec::new();
    push_sorted_unique(&mut pending, root, budget)?;
    let mut selected = Vec::new();
    while let Some(identifier) = pending.pop() {
        let Some((component, object)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        if object
            .messages
            .iter()
            .any(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            && comment_graph.is_none_or(|plan| plan.storage_ids.binary_search(&identifier).is_err())
        {
            return Err(SlideMediaLifecycleError::UnsupportedComment);
        }
        insert_sorted_unique(&mut selected, identifier, budget)?;
        budget.charge_entries(1)?;
        for info in &object.archive_info.message_infos {
            budget.charge_wire_fields(info.field_infos.len())?;
            budget.charge_wire_work(1)?;
            for reference in info.object_references.iter().copied() {
                budget.charge_references(1)?;
                if reference == 0 {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
                if reference == owning_slide {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
                if comment_graph
                    .is_some_and(|plan| plan.author_ids.binary_search(&reference).is_ok())
                {
                    continue;
                }
                match package.object_with_component(reference) {
                    Some((reference_component, _)) if reference_component == component_name => {
                        push_sorted_unique_if_new(&mut pending, &selected, reference, budget)?;
                    },
                    // Native movie graphs may retain shared style/stylesheet
                    // references outside the slide component. They are
                    // dependencies, never clone roots, and remain untouched.
                    Some(_) => {},
                    None => return Err(SlideMediaLifecycleError::InvalidSource),
                }
            }
            for field in &info.field_infos {
                for reference in field.object_references.iter().copied() {
                    budget.charge_references(1)?;
                    if reference == 0 {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    if reference == owning_slide {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    if comment_graph
                        .is_some_and(|plan| plan.author_ids.binary_search(&reference).is_ok())
                    {
                        continue;
                    }
                    match package.object_with_component(reference) {
                        Some((reference_component, _)) if reference_component == component_name => {
                            push_sorted_unique_if_new(&mut pending, &selected, reference, budget)?;
                        },
                        Some(_) => {},
                        None => return Err(SlideMediaLifecycleError::InvalidSource),
                    }
                }
            }
        }
    }
    selected.sort_unstable();
    Ok(selected)
}

fn selected_builds(
    package: &Package,
    component_name: &str,
    slide_payload: &[u8],
    movie_identifier: u64,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let build_ids = references_in_field(slide_payload, SLIDE_BUILDS_FIELD, limits, budget)?;
    let mut selected = Vec::new();
    reserve_u64s(&mut selected, build_ids.len(), budget)?;
    for identifier in build_ids {
        let Some((component, object)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let payload = unique_message_payload(object, BUILD_MESSAGE_TYPE, budget)?;
        let target = build_target_identifier(payload, budget)?;
        if target == movie_identifier {
            push_u64(&mut selected, identifier, budget)?;
        }
    }
    Ok(selected)
}

fn selected_chunks(
    package: &Package,
    component_name: &str,
    slide_payload: &[u8],
    build_ids: &[u64],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let chunk_ids = references_in_field(slide_payload, SLIDE_BUILD_CHUNKS_FIELD, limits, budget)?;
    let mut selected = Vec::new();
    reserve_u64s(&mut selected, chunk_ids.len(), budget)?;
    let build_lookup = sorted_ids(build_ids, budget)?;
    for identifier in chunk_ids {
        let Some((component, object)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let payload = unique_message_payload(object, BUILD_CHUNK_MESSAGE_TYPE, budget)?;
        let target = build_chunk_target_identifier(payload, budget)?;
        if contains_sorted(&build_lookup, target, budget)? {
            push_u64(&mut selected, identifier, budget)?;
        }
        /*
         * The codec validates both UUID edges and all unknown fields.  It is
         * intentionally the source of truth for BuildChunk topology; the
         * graph only retains the selected object identifier.
         */
    }
    Ok(selected)
}

fn data_edges(
    package: &Package,
    component_name: &str,
    private_object_ids: &[u64],
    build_ids: &[u64],
    chunk_ids: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<(u64, u64)>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    let total_ids = private_object_ids
        .len()
        .checked_add(build_ids.len())
        .and_then(|value| value.checked_add(chunk_ids.len()))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    reserve_pairs(&mut output, total_ids, budget)?;
    for identifier in private_object_ids
        .iter()
        .chain(build_ids)
        .chain(chunk_ids)
        .copied()
    {
        let Some((component, object)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        for info in &object.archive_info.message_infos {
            budget.charge_wire_fields(
                1usize
                    .checked_add(info.data_references.len())
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?,
            )?;
            let mut field_references = Vec::new();
            let field_reference_count =
                info.field_infos.iter().try_fold(0usize, |total, field| {
                    total
                        .checked_add(field.data_references.len())
                        .ok_or(SlideMediaLifecycleError::InvalidSource)
                })?;
            reserve_u64s(&mut field_references, field_reference_count, budget)?;
            for field in &info.field_infos {
                if !field.data_references.is_empty()
                    && field.r#type != Some(FieldType::DataReference)
                {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
                for data in field.data_references.iter().copied() {
                    if data == 0 {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    budget.charge_references(1)?;
                    push_u64(&mut field_references, data, budget)?;
                }
            }
            if !field_references.is_empty() {
                let mut aggregate = Vec::new();
                reserve_u64s(&mut aggregate, info.data_references.len(), budget)?;
                for data in info.data_references.iter().copied() {
                    if data == 0 {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    budget.charge_references(1)?;
                    push_u64(&mut aggregate, data, budget)?;
                }
                let mut sorted_fields = field_references;
                sorted_fields.sort_unstable();
                aggregate.sort_unstable();
                if sorted_fields != aggregate {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
            } else {
                budget.charge_references(info.data_references.len())?;
            }
            for data in info.data_references.iter().copied() {
                if data == 0 {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
                push_pair(&mut output, (data, identifier), budget)?;
            }
        }
    }
    Ok(output)
}

/// Prove that removing `removed_ids` leaves no surviving object edge aimed at
/// one of those objects.  Archive metadata is a useful witness, but native
/// build and build-chunk payloads are checked as well because their logical
/// edges may be represented by an opaque or producer-specific field census.
/// The caller supplies the sorted, unique object identifiers that were
/// removed after the candidate slide payload has been rewritten.
pub(super) fn validate_removal_closure(
    package: &Package,
    removed_ids: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_work(removed_ids.len().max(1))?;
    if removed_ids.windows(2).any(|window| window[0] >= window[1]) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let archive_limits = package
        .limits()
        .effective_archive_limits()
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    for component in package.state.source.components().iter() {
        budget.charge_entries(component.archive().objects.len())?;
        for object in &component.archive().objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(SlideMediaLifecycleError::InvalidSource);
            };
            if removed_ids.binary_search(&identifier).is_ok() {
                continue;
            }
            if object.archive_info.message_infos.len() != object.messages.len() {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            let header_limits =
                reserve_core_header_inspection_work(object, archive_limits, budget)?;
            strict_reference_census(object, header_limits, budget)?;
            let metadata_fields =
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |total, info| {
                        total
                            .checked_add(info.field_infos.len())
                            .ok_or(SlideMediaLifecycleError::InvalidSource)
                    })?;
            budget.charge_wire_fields(
                object
                    .archive_info
                    .message_infos
                    .len()
                    .checked_add(metadata_fields)
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?,
            )?;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                for reference in &info.object_references {
                    if removed_ids.binary_search(reference).is_ok() {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                }
                for field in &info.field_infos {
                    for reference in &field.object_references {
                        if removed_ids.binary_search(reference).is_ok() {
                            return Err(SlideMediaLifecycleError::InvalidSource);
                        }
                    }
                }
                if matches!(message.type_, BUILD_MESSAGE_TYPE | BUILD_CHUNK_MESSAGE_TYPE) {
                    budget.charge_wire_work(message.data.len().max(1))?;
                    let reference = if message.type_ == BUILD_MESSAGE_TYPE {
                        build_target_identifier(&message.data, budget)?
                    } else {
                        build_chunk_target_identifier(&message.data, budget)?
                    };
                    if removed_ids.binary_search(&reference).is_ok() {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(())
}

/// Clone one source object and rewrite only known object-reference envelopes.
pub(super) fn clone_object(
    source: &ArchiveObject,
    new_identifier: u64,
    object_remap: &[(u64, u64)],
    limits: ArchiveObjectLimits,
    movie_geometry_offset: bool,
    movie_uuid: Option<SuperUuid>,
    transitive_style_witnesses: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<ArchiveObject, SlideMediaLifecycleError> {
    if source.archive_info.should_merge == Some(true)
        || source.archive_info.message_infos.iter().any(|message| {
            message.base_message_index.is_some()
                || !message.diff_merge_version.is_empty()
                || message.diff_field_path.is_some()
                || !message.fields_to_remove.is_empty()
                || !message.diff_read_version.is_empty()
        })
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let inspection_limits = reserve_core_header_inspection_work(source, limits, budget)?;
    strict_reference_census(source, inspection_limits, budget)?;
    let limits = reserve_core_header_work(source, object_remap.len(), limits, budget)?;
    let mut replacements = Vec::new();
    let replacement_bytes = source
        .messages
        .len()
        .checked_mul(size_of::<RawMessage>())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocations(replacement_bytes)?;
    replacements
        .try_reserve_exact(source.messages.len())
        .map_err(|_| SlideMediaLifecycleError::Allocation {
            amount: replacement_bytes,
        })?;
    for (message_index, message) in source.messages.iter().enumerate() {
        let info = source
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        let wire_limits = payload_limits(message.data.len())?;
        let scratch = message
            .data
            .len()
            .checked_mul(2)
            .and_then(|value| value.checked_add(256))
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        budget.charge_allocations(scratch)?;
        let mut data = if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
            super::comment_clone::rewrite_comment_payload(&message.data, object_remap, budget)?
        } else if message.type_ == BUILD_MESSAGE_TYPE {
            match rewrite_build_payload(&message.data, object_remap, budget)? {
                Some(data) => data,
                None => {
                    if has_mapped_header_reference(info, object_remap, budget)? {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    copy_bytes(&message.data, budget)?
                },
            }
        } else if message.type_ == BUILD_CHUNK_MESSAGE_TYPE {
            match rewrite_build_chunk_payload(&message.data, object_remap, movie_uuid, budget)? {
                Some(data) => data,
                None => {
                    if has_mapped_header_reference(info, object_remap, budget)? {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    copy_bytes(&message.data, budget)?
                },
            }
        } else {
            let mut direct_header_references = Vec::new();
            let source_references =
                if message.type_ == MOVIE_MESSAGE_TYPE && !transitive_style_witnesses.is_empty() {
                    reserve_bytes(
                        &mut direct_header_references,
                        info.object_references.len(),
                        budget,
                    )?;
                    budget.charge_wire_work(info.object_references.len().saturating_mul(2))?;
                    direct_header_references.extend(
                        info.object_references
                            .iter()
                            .copied()
                            .filter(|identifier| !transitive_style_witnesses.contains(identifier)),
                    );
                    direct_header_references.as_slice()
                } else {
                    info.object_references.as_slice()
                };
            let mut budget_error = None;
            let mut charge = |amount| {
                budget.charge_allocations(amount).map_err(|error| {
                    budget_error = Some(error);
                    super::clone_payload::ClonePayloadError::Budget
                })
            };
            match super::clone_payload::remap_clone_payload_with_budget(
                &message.data,
                message.type_,
                object_remap,
                source_references,
                wire_limits,
                &mut charge,
            ) {
                Ok(rewrite) => {
                    let report = rewrite.report();
                    charge_clone_payload_report(budget, report)?;
                    let data = rewrite.into_payload();
                    budget.charge_output(data.len())?;
                    data
                },
                Err(super::clone_payload::ClonePayloadError::UnsupportedMessageType(_)) => {
                    if has_mapped_header_reference(info, object_remap, budget)? {
                        return Err(SlideMediaLifecycleError::InvalidSource);
                    }
                    copy_bytes(&message.data, budget)?
                },
                Err(error) => {
                    return Err(budget_error.unwrap_or_else(|| map_clone_payload_error(error)));
                },
            }
        };
        if message.type_ == MOVIE_MESSAGE_TYPE && movie_geometry_offset {
            data = offset_movie_geometry(&data, payload_limits(data.len())?, budget)?;
        }
        replacements.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    let output_bytes = replacements.iter().try_fold(0usize, |total, message| {
        total
            .checked_add(message.data.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)
    })?;
    budget.charge_allocations(output_bytes)?;
    budget.charge_output(output_bytes)?;
    source
        .clone_with_identity_remap_with_limits(new_identifier, object_remap, &replacements, limits)
        .map_err(map_archive_error)
}

fn rewrite_build_payload(
    payload: &[u8],
    object_remap: &[(u64, u64)],
    budget: &mut LifecycleBudget,
) -> Result<Option<Vec<u8>>, SlideMediaLifecycleError> {
    let options = lifecycle_decode_options(payload);
    let (snapshot, report) =
        litchi_iwa_protos::keynote_media_lifecycle_codec::decode_build_with_report(
            payload, options,
        )
        .map_err(map_lifecycle_decode_error)?;
    charge_decode_report(budget, report)?;
    let source = snapshot
        .drawable()
        .map(|reference| reference.identifier())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_references(1)?;
    let Some(target) = lookup_remap(object_remap, source, budget)? else {
        return Ok(None);
    };
    budget.charge_allocations(payload.len().saturating_mul(2))?;
    let edit = litchi_iwa_protos::keynote_media_lifecycle_codec::BuildLifecycleEdit::drawable(
        litchi_iwa_protos::keynote_media_lifecycle_codec::IdentifierRewrite::new(source, target),
    );
    match litchi_iwa_protos::keynote_media_lifecycle_codec::rewrite_build_with_report(
        payload, edit, options,
    ) {
        Ok((output, report)) => {
            charge_rewrite_report(budget, report)?;
            Ok(Some(output))
        },
        Err(error) => Err(map_lifecycle_decode_error(error)),
    }
}

fn rewrite_build_chunk_payload(
    payload: &[u8],
    object_remap: &[(u64, u64)],
    movie_uuid: Option<SuperUuid>,
    budget: &mut LifecycleBudget,
) -> Result<Option<Vec<u8>>, SlideMediaLifecycleError> {
    let options = lifecycle_decode_options(payload);
    let (snapshot, report) =
        litchi_iwa_protos::keynote_media_lifecycle_codec::decode_build_chunk_with_report(
            payload, options,
        )
        .map_err(map_lifecycle_decode_error)?;
    charge_decode_report(budget, report)?;
    let source_build = snapshot.build().identifier();
    budget.charge_references(1)?;
    let target_build = lookup_remap(object_remap, source_build, budget)?;
    if target_build.is_none() && movie_uuid.is_none() {
        return Ok(None);
    }

    budget.charge_allocations(payload.len().saturating_mul(2))?;
    let mut edit =
        litchi_iwa_protos::keynote_media_lifecycle_codec::BuildChunkLifecycleEdit::empty();
    if let Some(target) = target_build {
        edit = edit.with_build(
            litchi_iwa_protos::keynote_media_lifecycle_codec::IdentifierRewrite::new(
                source_build,
                target,
            ),
        );
    }
    if let Some(target_uuid) = movie_uuid {
        let source_uuid = snapshot
            .chunk_identifier()
            .or_else(|| snapshot.build_id())
            .map(|uuid| uuid.uuid())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        edit = edit.with_uuid(
            litchi_iwa_protos::keynote_media_lifecycle_codec::UuidRewrite::new(
                litchi_iwa_protos::keynote_media_lifecycle_codec::Uuid::new(
                    source_uuid.lower(),
                    source_uuid.upper(),
                ),
                litchi_iwa_protos::keynote_media_lifecycle_codec::Uuid::new(
                    target_uuid.lower,
                    target_uuid.upper,
                ),
            ),
        );
    }
    match litchi_iwa_protos::keynote_media_lifecycle_codec::rewrite_build_chunk_with_report(
        payload, edit, options,
    ) {
        Ok((output, report)) => {
            charge_rewrite_report(budget, report)?;
            Ok(Some(output))
        },
        Err(error) => Err(map_lifecycle_decode_error(error)),
    }
}

fn lookup_remap(
    remap: &[(u64, u64)],
    source: u64,
    budget: &mut LifecycleBudget,
) -> Result<Option<u64>, SlideMediaLifecycleError> {
    budget.charge_wire_work(remap.len().max(1))?;
    Ok(remap
        .binary_search_by_key(&source, |&(old, _)| old)
        .ok()
        .map(|index| remap[index].1))
}

fn offset_movie_geometry(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    // Movie.super -> Drawable.geometry -> Geometry.position -> Point.x/y.
    // The lifecycle owner changes only the position;
    // size, playback, captions, title, and opaque fields remain source bytes.
    let shifted_x = rewrite_nested_f32(payload, &[1, 1, 1, 1], 10.0, limits, budget)?;
    rewrite_nested_f32(&shifted_x, &[1, 1, 1, 2], 10.0, limits, budget)
}

fn rewrite_nested_f32(
    payload: &[u8],
    path: &[u32],
    delta: f32,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    let mut replacements = 0usize;
    let output = rewrite_nested_f32_inner(payload, path, delta, limits, budget, &mut replacements)?;
    if replacements != 1 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(output)
}

fn rewrite_nested_f32_inner(
    payload: &[u8],
    path: &[u32],
    delta: f32,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    replacements: &mut usize,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    let view = parse_view(payload, limits, budget, path.len())?;
    let mut output = Vec::new();
    reserve_bytes(&mut output, payload.len(), budget)?;
    for field in view.fields() {
        if path.len() == 1 && field.number() == path[0] && field.wire_type() == 5 {
            let value = field.payload();
            if value.len() != 4 {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            let bytes: [u8; 4] = value
                .try_into()
                .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
            let current = f32::from_le_bytes(bytes);
            let replacement = (current + delta).to_le_bytes();
            *replacements = replacements
                .checked_add(1)
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            output.extend_from_slice(field.key());
            // Fixed-width replacement preserves the source key and payload
            // span exactly, while retaining any noncanonical outer framing.
            output.extend_from_slice(&replacement);
        } else if field.wire_type() == 2 && path.first() == Some(&field.number()) {
            let nested = rewrite_nested_f32_inner(
                field.payload(),
                &path[1..],
                delta,
                limits,
                budget,
                replacements,
            )?;
            if nested == field.payload() {
                output.extend_from_slice(field.raw());
            } else {
                append_length_delimited(&mut output, field.number(), &nested, limits, budget)?;
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    budget.charge_output(output.len())?;
    Ok(output)
}

fn append_length_delimited(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_work(output.len().max(1))?;
    let payload_len =
        u64::try_from(payload.len()).map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let required = litchi_iwa_common::varint::encoded_len(u64::from(field_number) << 3 | 2)
        .checked_add(litchi_iwa_common::varint::encoded_len(payload_len))
        .and_then(|value| value.checked_add(payload.len()))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if output.capacity().saturating_sub(output.len()) < required {
        budget.charge_allocations(required)?;
    }
    litchi_iwa_common::wire::append_length_delimited_field_with_limits(
        output,
        field_number,
        payload,
        limits,
    )
    .map_err(map_wire_error)
}

fn copy_boxed(
    value: &str,
    budget: &mut LifecycleBudget,
) -> Result<Box<str>, SlideMediaLifecycleError> {
    let mut output = String::new();
    budget.charge_allocations(value.len())?;
    output
        .try_reserve_exact(value.len())
        .map_err(|_| SlideMediaLifecycleError::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

fn payload_limits(length: usize) -> Result<WireLimits, SlideMediaLifecycleError> {
    let output = length
        .checked_add(512)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    WireLimits::default()
        .with_input_bytes(length.max(1))
        .and_then(|limits| limits.with_output_bytes(output))
        .map_err(map_wire_error)
}

fn parse_view<'source>(
    payload: &'source [u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    nesting: usize,
) -> Result<WireView<'source>, SlideMediaLifecycleError> {
    budget.charge_wire_work(payload.len().max(1))?;
    budget.charge_nesting(nesting.max(1))?;
    // WireView is borrowed, but its parser still performs bounded scratch
    // accounting internally.  Precharge the conservative source-width upper
    // bound before entering it so hostile nesting cannot allocate first.
    budget.charge_allocations(payload.len())?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.charge_wire_fields(view.len())?;
    Ok(view)
}

fn reserve_bytes<T>(
    output: &mut Vec<T>,
    additional: usize,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if additional == 0 {
        return Ok(());
    }
    budget.charge_allocations(
        additional
            .checked_mul(size_of::<T>())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?,
    )?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| SlideMediaLifecycleError::Allocation {
            amount: additional.checked_mul(size_of::<T>()).unwrap_or(usize::MAX),
        })
}

fn reserve_u64s(
    output: &mut Vec<u64>,
    additional: usize,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    reserve_bytes(output, additional, budget)
}

fn reserve_pairs(
    output: &mut Vec<(u64, u64)>,
    additional: usize,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    reserve_bytes(output, additional, budget)
}

fn push_u64(
    output: &mut Vec<u64>,
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if output.len() == output.capacity() {
        reserve_u64s(output, 1, budget)?;
    }
    output.push(value);
    Ok(())
}

fn push_pair(
    output: &mut Vec<(u64, u64)>,
    value: (u64, u64),
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if output.len() == output.capacity() {
        reserve_pairs(output, 1, budget)?;
    }
    output.push(value);
    Ok(())
}

fn insert_sorted_unique(
    output: &mut Vec<u64>,
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    match output.binary_search(&value) {
        Ok(_) => Ok(false),
        Err(index) => {
            budget.charge_wire_work(output.len().max(1))?;
            if output.len() == output.capacity() {
                reserve_u64s(output, 1, budget)?;
            }
            output.insert(index, value);
            Ok(true)
        },
    }
}

fn push_sorted_unique(
    output: &mut Vec<u64>,
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let _ = insert_sorted_unique(output, value, budget)?;
    Ok(())
}

fn push_sorted_unique_if_new(
    pending: &mut Vec<u64>,
    selected: &[u64],
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_work(selected.len().max(1))?;
    if selected.binary_search(&value).is_ok() {
        return Ok(());
    }
    let _ = insert_sorted_unique(pending, value, budget)?;
    Ok(())
}

fn reserve_bytes_copy(
    source: &[u8],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    reserve_bytes(&mut output, source.len(), budget)?;
    output.extend_from_slice(source);
    budget.charge_output(output.len())?;
    Ok(output)
}

fn copy_bytes(
    source: &[u8],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    reserve_bytes_copy(source, budget)
}

fn sorted_ids(
    source: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    reserve_u64s(&mut output, source.len(), budget)?;
    output.extend_from_slice(source);
    output.sort_unstable();
    budget.charge_wire_work(source.len().max(1))?;
    Ok(output)
}

fn unique_message_payload<'source>(
    object: &'source ArchiveObject,
    message_type: u32,
    budget: &mut LifecycleBudget,
) -> Result<&'source [u8], SlideMediaLifecycleError> {
    budget.charge_wire_fields(object.messages.len())?;
    budget.charge_wire_work(object.messages.len().max(1))?;
    let mut found = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if found.replace(message.data.as_slice()).is_some() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    found.ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn has_mapped_header_reference(
    info: &litchi_iwa_core::MessageInfo,
    remap: &[(u64, u64)],
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    let field_count = info
        .field_infos
        .len()
        .checked_add(info.object_references.len())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_fields(field_count)?;
    budget.charge_wire_work(field_count.max(1))?;
    let aggregate = info.object_references.iter().any(|reference| {
        remap
            .binary_search_by_key(reference, |&(source, _)| source)
            .is_ok()
    });
    let fields = info.field_infos.iter().any(|field| {
        field.object_references.iter().any(|reference| {
            remap
                .binary_search_by_key(reference, |&(source, _)| source)
                .is_ok()
        })
    });
    Ok(aggregate || fields)
}

fn canonical_int32(value: u64) -> bool {
    value <= i32::MAX as u64 || value >= MIN_SIGN_EXTENDED_INT32
}

fn contains_sorted(
    values: &[u64],
    needle: u64,
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    budget.charge_wire_work(values.len().max(1))?;
    Ok(values.binary_search(&needle).is_ok())
}

fn lifecycle_decode_options(
    payload: &[u8],
) -> litchi_iwa_protos::keynote_media_lifecycle_codec::DecodeOptions {
    litchi_iwa_protos::keynote_media_lifecycle_codec::DecodeOptions::for_source(payload)
}

fn build_target_identifier(
    payload: &[u8],
    budget: &mut LifecycleBudget,
) -> Result<u64, SlideMediaLifecycleError> {
    budget.charge_allocations(payload.len().saturating_mul(2))?;
    match litchi_iwa_protos::keynote_media_lifecycle_codec::decode_build_with_report(
        payload,
        lifecycle_decode_options(payload),
    ) {
        Ok((snapshot, report)) => {
            charge_decode_report(budget, report)?;
            let identifier = snapshot
                .drawable()
                .map(|reference| reference.identifier())
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            budget.charge_references(1)?;
            Ok(identifier)
        },
        Err(error) => Err(map_lifecycle_decode_error(error)),
    }
}

fn build_chunk_target_identifier(
    payload: &[u8],
    budget: &mut LifecycleBudget,
) -> Result<u64, SlideMediaLifecycleError> {
    budget.charge_allocations(payload.len().saturating_mul(2))?;
    match litchi_iwa_protos::keynote_media_lifecycle_codec::decode_build_chunk_with_report(
        payload,
        lifecycle_decode_options(payload),
    ) {
        Ok((snapshot, report)) => {
            charge_decode_report(budget, report)?;
            let identifier = snapshot.build().identifier();
            budget.charge_references(1)?;
            Ok(identifier)
        },
        Err(error) => Err(map_lifecycle_decode_error(error)),
    }
}

fn charge_decode_report(
    budget: &mut LifecycleBudget,
    report: litchi_iwa_protos::keynote_media_lifecycle_codec::DecodeReport,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.scratch_bytes())?;
    for _ in 0..report.allocations() {
        budget.charge_allocations(0)?;
    }
    Ok(())
}

fn charge_rewrite_report(
    budget: &mut LifecycleBudget,
    report: litchi_iwa_protos::keynote_media_lifecycle_codec::RewriteReport,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_output(report.output_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    for _ in 0..report.allocations() {
        budget.charge_allocations(0)?;
    }
    Ok(())
}

fn map_lifecycle_decode_error(
    error: litchi_iwa_protos::keynote_media_lifecycle_codec::DecodeError,
) -> SlideMediaLifecycleError {
    use litchi_iwa_protos::keynote_media_lifecycle_codec::DecodeLimit;
    if let Some(limit) = error.resource_limit() {
        return match limit {
            DecodeLimit::Bytes { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
                kind: super::SlideMediaLifecycleLimitKind::InputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            },
            DecodeLimit::Fields { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
                kind: super::SlideMediaLifecycleLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            },
            DecodeLimit::Work { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
                kind: super::SlideMediaLifecycleLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            },
            DecodeLimit::OutputBytes { observed, maximum } => {
                SlideMediaLifecycleError::LimitExceeded {
                    kind: super::SlideMediaLifecycleLimitKind::OutputBytes,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            DecodeLimit::References { observed, maximum } => {
                SlideMediaLifecycleError::LimitExceeded {
                    kind: super::SlideMediaLifecycleLimitKind::References,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            DecodeLimit::Nesting { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
                kind: super::SlideMediaLifecycleLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            },
            _ => SlideMediaLifecycleError::InvalidSource,
        };
    }
    SlideMediaLifecycleError::InvalidSource
}

fn charge_clone_payload_report(
    budget: &mut LifecycleBudget,
    report: super::clone_payload::ClonePayloadReport,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_references(report.references_seen())?;
    budget.charge_nesting(report.max_depth())?;
    for _ in 0..report.allocations() {
        budget.charge_allocations(0)?;
    }
    Ok(())
}

fn map_clone_payload_error(
    error: super::clone_payload::ClonePayloadError,
) -> SlideMediaLifecycleError {
    match error {
        super::clone_payload::ClonePayloadError::Wire(error) => map_wire_error(error),
        super::clone_payload::ClonePayloadError::UnsupportedMessageType(_)
        | super::clone_payload::ClonePayloadError::InvalidReference
        | super::clone_payload::ClonePayloadError::SourceReferenceWitnessMismatch
        | super::clone_payload::ClonePayloadError::Budget
        | super::clone_payload::ClonePayloadError::InvalidRemap => {
            SlideMediaLifecycleError::InvalidSource
        },
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideMediaLifecycleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => {
            let kind = match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    super::SlideMediaLifecycleLimitKind::InputBytes
                },
                litchi_iwa_common::LimitKind::Fields => {
                    super::SlideMediaLifecycleLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::OutputBytes => {
                    super::SlideMediaLifecycleLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    super::SlideMediaLifecycleLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork => {
                    super::SlideMediaLifecycleLimitKind::WireWork
                },
                _ => return SlideMediaLifecycleError::InvalidSource,
            };
            SlideMediaLifecycleError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideMediaLifecycleError::Allocation { amount }
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_core::Error) -> SlideMediaLifecycleError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::Entries,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMediaLifecycleError::Allocation { amount: requested }
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}

/// Publish a slide payload and its exact aggregate/field reference transition.
pub(super) fn replace_slide_message_with_lifecycle_refs(
    object: &mut ArchiveObject,
    index: usize,
    payload: Vec<u8>,
    limits: ArchiveObjectLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
    use litchi_iwa_protos::keynote_media_lifecycle_codec as codec;
    let source = object
        .messages
        .get(index)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let (before, before_report) = codec::decode_slide_lifecycle_with_report(
        &source.data,
        lifecycle_decode_options(&source.data),
    )
    .map_err(map_lifecycle_decode_error)?;
    charge_decode_report(budget, before_report)?;
    let (after, after_report) =
        codec::decode_slide_lifecycle_with_report(&payload, lifecycle_decode_options(&payload))
            .map_err(map_lifecycle_decode_error)?;
    charge_decode_report(budget, after_report)?;
    let groups = [
        (
            2,
            collect_reference_ids(before.builds(), budget)?,
            collect_reference_ids(after.builds(), budget)?,
        ),
        (
            7,
            collect_reference_ids(before.owned_drawables(), budget)?,
            collect_reference_ids(after.owned_drawables(), budget)?,
        ),
        (
            42,
            collect_reference_ids(before.drawables_z_order(), budget)?,
            collect_reference_ids(after.drawables_z_order(), budget)?,
        ),
        (
            43,
            collect_reference_ids(before.build_chunks(), budget)?,
            collect_reference_ids(after.build_chunks(), budget)?,
        ),
    ];
    let mut before_union = Vec::new();
    let mut after_union = Vec::new();
    for (_, before, after) in &groups {
        reserve_bytes(&mut before_union, before.len(), budget)?;
        before_union.extend_from_slice(before);
        reserve_bytes(&mut after_union, after.len(), budget)?;
        after_union.extend_from_slice(after);
    }
    budget.charge_wire_work(
        before_union
            .len()
            .saturating_add(after_union.len())
            .saturating_mul(usize::BITS as usize),
    )?;
    before_union.sort_unstable();
    before_union.dedup();
    after_union.sort_unstable();
    after_union.dedup();
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    for reference in &before_union {
        budget.charge_wire_work(info.object_references.len())?;
        if !info.object_references.contains(reference) {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    for field in &info.field_infos {
        if let Some((_, before, _)) = groups
            .iter()
            .find(|(number, _, _)| field.path.as_slice() == [*number])
        {
            budget.charge_wire_work(field.object_references.len())?;
            if !field.object_references.is_empty()
                && (field.r#type != Some(FieldType::ObjectReference)
                    || field.object_references != *before)
            {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
        }
    }
    if after_union
        .iter()
        .all(|identifier| before_union.binary_search(identifier).is_ok())
    {
        let mut removed = Vec::new();
        reserve_bytes(&mut removed, before_union.len(), budget)?;
        for identifier in &before_union {
            if after_union.binary_search(identifier).is_err() {
                removed.push(*identifier);
            }
        }
        for identifier in &removed {
            budget.charge_wire_work(info.object_references.len())?;
            if info
                .object_references
                .iter()
                .filter(|value| *value == identifier)
                .count()
                != 1
            {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
        }
        if removed.binary_search(&before.style().identifier()).is_ok() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        for field in &info.field_infos {
            if !groups
                .iter()
                .any(|(number, _, _)| field.path.as_slice() == [*number])
                && field
                    .object_references
                    .iter()
                    .any(|identifier| removed.binary_search(identifier).is_ok())
            {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
        }
        let limits = reserve_core_header_work(object, removed.len(), limits, budget)?;
        return object
            .replace_message_pruning_object_references_preserving_header_with_limits(
                index,
                RawMessage {
                    type_: super::SLIDE_MESSAGE_TYPE,
                    data: payload,
                },
                &removed,
                limits,
            )
            .map(|_| ())
            .map_err(map_archive_error);
    }
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let aggregate_before = copy_reference_ids(&info.object_references, budget)?;
    for reference in &before_union {
        budget.charge_wire_work(aggregate_before.len())?;
        if !aggregate_before.contains(reference) {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    let aggregate_after =
        transition_reference_ids(&aggregate_before, &before_union, &after_union, budget)?;
    let mut owned_fields = Vec::new();
    reserve_bytes(&mut owned_fields, info.field_infos.len(), budget)?;
    for (ordinal, field) in info.field_infos.iter().enumerate() {
        let Some((_, before, after)) = groups
            .iter()
            .find(|(number, _, _)| field.path.as_slice() == [*number])
        else {
            continue;
        };
        if field.object_references.is_empty() {
            continue;
        }
        // A partial or aliased field census needs a separate semantic witness.
        if field.object_references != *before {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let mut path = Vec::new();
        reserve_bytes(&mut path, field.path.as_slice().len(), budget)?;
        path.extend_from_slice(field.path.as_slice());
        let field_before = copy_reference_ids(&field.object_references, budget)?;
        let mut before_set = copy_reference_ids(before, budget)?;
        let mut after_set = copy_reference_ids(after, budget)?;
        budget.charge_wire_work(
            before_set
                .len()
                .saturating_add(after_set.len())
                .saturating_mul(usize::BITS as usize),
        )?;
        before_set.sort_unstable();
        after_set.sort_unstable();
        let field_after = transition_reference_ids(&field_before, &before_set, &after_set, budget)?;
        owned_fields.push((ordinal, path, field_before, field_after));
    }
    let mut fields = Vec::new();
    reserve_bytes(&mut fields, owned_fields.len(), budget)?;
    for (ordinal, path, before, after) in &owned_fields {
        fields.push(FieldObjectReferenceTransition {
            field_info_index: *ordinal,
            expected_path: path,
            before,
            after,
        });
    }
    let growth = fields
        .iter()
        .try_fold(aggregate_after.len(), |total, field| {
            total
                .checked_add(field.after.len())
                .ok_or(SlideMediaLifecycleError::InvalidSource)
        })?;
    let limits = reserve_core_header_work(object, growth, limits, budget)?;
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            index,
            RawMessage {
                type_: super::SLIDE_MESSAGE_TYPE,
                data: payload,
            },
            ObjectReferenceTransition {
                aggregate_before: &aggregate_before,
                aggregate_after: &aggregate_after,
                fields: &fields,
            },
            limits,
        )
        .map(|_| ())
        .map_err(map_archive_error)
}

fn collect_reference_ids<'source>(
    references: impl ExactSizeIterator<
        Item = litchi_iwa_protos::keynote_media_lifecycle_codec::Reference<'source>,
    >,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    reserve_bytes(&mut output, references.len(), budget)?;
    output.extend(references.map(|reference| reference.identifier()));
    Ok(output)
}

fn copy_reference_ids(
    source: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    reserve_bytes(&mut output, source.len(), budget)?;
    output.extend_from_slice(source);
    Ok(output)
}

fn transition_reference_ids(
    source: &[u64],
    before: &[u64],
    after: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, SlideMediaLifecycleError> {
    let mut output = Vec::new();
    reserve_bytes(
        &mut output,
        source
            .len()
            .checked_add(after.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?,
        budget,
    )?;
    budget.charge_wire_work(
        source
            .len()
            .saturating_add(after.len())
            .saturating_mul(usize::BITS as usize),
    )?;
    for &identifier in source {
        if before.binary_search(&identifier).is_err() || after.binary_search(&identifier).is_ok() {
            output.push(identifier);
        }
    }
    for &identifier in after {
        if before.binary_search(&identifier).is_err() {
            // New identifiers must not alias an unrelated retained header edge.
            budget.charge_wire_work(source.len())?;
            if source.contains(&identifier) {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            output.push(identifier);
        }
    }
    Ok(output)
}

struct ReferenceOccurrenceCounter;

impl ArchiveReferenceVisitor for ReferenceOccurrenceCounter {
    fn visit_reference(
        &mut self,
        _occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        Ok(())
    }
}

/// Strictly census one source header before a clone or removal closure scan.
///
/// The core archive reader rejects unknown reference-bearing header metadata;
/// charging its returned occurrence count keeps aggregate and field-level
/// object/data references on the operation-wide lifecycle ledger.
pub(super) fn strict_reference_census(
    source: &ArchiveObject,
    limits: ArchiveObjectLimits,
    budget: &mut LifecycleBudget,
) -> Result<usize, SlideMediaLifecycleError> {
    let mut visitor = ReferenceOccurrenceCounter;
    let occurrences = source
        .inspect_references_with_policy_and_limits(
            &mut visitor,
            ArchiveReferencePolicy::RejectUnknownMetadata,
            limits,
        )
        .map_err(map_archive_error)?;
    budget.charge_references(occurrences)?;
    Ok(occurrences)
}

/// Reserve the source-sized core header arena needed by a read-only reference
/// census.  The decoded limits are identical to the mutation reservation, but
/// the shared lifecycle ledger only admits the canonical source, preflight,
/// and neutral decode arenas that `inspect_references_with_policy_and_limits`
/// can hold concurrently.  A removal scan must not consume four rewrite
/// arenas for every unrelated surviving object.
pub(super) fn reserve_core_header_inspection_work(
    source: &ArchiveObject,
    limits: ArchiveObjectLimits,
    budget: &mut LifecycleBudget,
) -> Result<ArchiveObjectLimits, SlideMediaLifecycleError> {
    reserve_core_header_work_inner(source, 0, limits, budget, false)
}

/// Reserve a source-sized core header arena before entering its bounded codecs.
/// The input is an unchanged physical source header; mutations only add known
/// references or grow identifiers/message lengths to at most ten-byte varints.
pub(super) fn reserve_core_header_work(
    source: &ArchiveObject,
    additional_references: usize,
    limits: ArchiveObjectLimits,
    budget: &mut LifecycleBudget,
) -> Result<ArchiveObjectLimits, SlideMediaLifecycleError> {
    reserve_core_header_work_inner(source, additional_references, limits, budget, true)
}

fn reserve_core_header_work_inner(
    source: &ArchiveObject,
    additional_references: usize,
    limits: ArchiveObjectLimits,
    budget: &mut LifecycleBudget,
    rewrite: bool,
) -> Result<ArchiveObjectLimits, SlideMediaLifecycleError> {
    let source_bytes = usize::try_from(source.header_length)
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    if source_bytes == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let mut variable_fields = additional_references
        .checked_add(2)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let mut containers = source
        .archive_info
        .message_infos
        .len()
        .checked_add(1)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    for message in &source.archive_info.message_infos {
        budget.charge_wire_work(message.field_infos.len().saturating_add(1))?;
        // Merge/diff metadata is outside this lifecycle owner's admission.
        if message.base_message_index.is_some()
            || !message.diff_merge_version.is_empty()
            || message.diff_field_path.is_some()
            || !message.fields_to_remove.is_empty()
            || !message.diff_read_version.is_empty()
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        variable_fields = variable_fields
            .checked_add(message.object_references.len())
            .and_then(|n| n.checked_add(2))
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        containers = containers
            .checked_add(message.field_infos.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        for field in &message.field_infos {
            variable_fields = variable_fields
                .checked_add(field.object_references.len())
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        }
    }
    let growth = variable_fields
        .checked_add(containers)
        .and_then(|n| n.checked_mul(10))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let header_bytes = source_bytes
        .checked_add(growth)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?
        .min(limits.max_header_bytes());
    // Every encoded field/item occupies at least one source byte. A 256-byte
    // cell per byte bounds decoded neutral/generated metadata and raw rewrite
    // cells; the core enforces this arena independently of the shared ledger.
    let memory = header_bytes
        .checked_mul(256)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?
        .min(limits.max_header_memory_bytes());
    let narrowed = limits
        .with_header_bytes(header_bytes)
        .and_then(|value| value.with_header_fields(header_bytes.min(limits.max_header_fields())))
        .and_then(|value| value.with_metadata_items(header_bytes.min(limits.max_metadata_items())))
        .and_then(|value| value.with_header_memory_bytes(memory))
        .map_err(map_archive_error)?;
    // Mutation owns four arenas: canonical-before, raw rewrite, decode, and
    // canonical-after.  A read-only reference census only needs the source
    // canonical bytes plus the preflight/decode projection.  Keep the latter
    // conservative while avoiding a per-object charge that scales as if a
    // rewrite had already begun.
    let (arena_count, header_overhead, event_factor, event_extra) = if rewrite {
        (4, 16, 64, 32)
    } else {
        (2, 8, 2, 4)
    };
    let bytes = memory
        .checked_mul(arena_count)
        .and_then(|n| {
            header_bytes
                .checked_mul(header_overhead)
                .and_then(|h| n.checked_add(h))
        })
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let events = containers
        .checked_mul(event_factor)
        .and_then(|n| n.checked_add(event_extra))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocation_plan(bytes, events)?;
    budget.charge_wire_work(bytes)?;
    Ok(narrowed)
}
