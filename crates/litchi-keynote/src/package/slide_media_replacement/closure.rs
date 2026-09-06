//! Bounded physical ownership checks for one slide-media replacement.
//!
//! The selector identifies a data edge, but the edge is safe to rewrite only
//! when the package metadata and every current owner agree about that edge.
//! This module keeps those checks separate from the ZIP rewrite so callers do
//! not accidentally admit a record by looking at only one representation.

use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireFieldView, WireView},
};
use litchi_iwa_core::{ArchiveObject, FieldType};

use super::{
    MOVIE_MESSAGE_TYPE, MediaBudget, MediaSelection, OwnedMetadataFacts, Package,
    SlideMediaDataError,
};

const DATA_METADATA_MAP_FIELD: u32 = 10;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
const DATA_METADATA_MESSAGE_TYPE: u32 = 11_014;
const DATA_METADATA_MAP_ENTRY_FIELD: u32 = 1;
const DATA_METADATA_MAP_ENTRY_DATA_FIELD: u32 = 1;
const DATA_METADATA_MAP_ENTRY_METADATA_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;

/// Validate the complete current ownership closure for one selected media ID.
///
/// The metadata payload is passed separately from the visitor facts because
/// `DataMetadataMap` is an optional root edge that the media visitor does not
/// project.  The function deliberately borrows all inputs and allocates only
/// a bounded list of map data IDs while checking duplicate entries.
pub(super) fn validate_selected_media_closure(
    package: &Package,
    metadata_payload: &[u8],
    facts: &OwnedMetadataFacts,
    selection: &MediaSelection,
    identifier: u64,
    budget: &mut MediaBudget,
) -> Result<(), SlideMediaDataError> {
    let expected_path = selected_field_path(selection, identifier)?;
    let locator = selection
        .component_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let component = unique_current_component(facts, locator)?;
    let expected_owner_count =
        unique_current_reference_count(facts, component.identifier, identifier)?;

    let mut owner_count = 0usize;
    let mut selected_owner_count = 0usize;
    for (index, owner) in facts.owners.iter().enumerate() {
        if owner.component_identifier != component.identifier
            || owner.data_identifier != identifier
            || owner.versioned
        {
            continue;
        }
        if owner.count == 0 || owner.object_identifier == 0 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        // The duplicate-owner audit is a bounded prefix scan. Charge its
        // work before walking the prefix so hostile owner lists cannot hide
        // quadratic work outside the operation ledger.
        budget.wire_work(index)?;
        if facts.owners[..index].iter().any(|previous| {
            !previous.versioned
                && previous.component_identifier == owner.component_identifier
                && previous.data_identifier == owner.data_identifier
                && previous.object_identifier == owner.object_identifier
        }) {
            return Err(SlideMediaDataError::InvalidSource);
        }
        owner_count = owner_count
            .checked_add(1)
            .ok_or(SlideMediaDataError::InvalidSource)?;
        if owner.object_identifier == selection.movie_identifier {
            selected_owner_count = selected_owner_count
                .checked_add(1)
                .ok_or(SlideMediaDataError::InvalidSource)?;
        }
    }
    if owner_count != expected_owner_count || selected_owner_count != 1 {
        return Err(SlideMediaDataError::InvalidSource);
    }

    for owner in facts.owners.iter().filter(|owner| {
        owner.component_identifier == component.identifier
            && owner.data_identifier == identifier
            && !owner.versioned
    }) {
        let (owner_component, object) = package
            .object_with_component(owner.object_identifier)
            .ok_or(SlideMediaDataError::InvalidSource)?;
        if owner_component != selection.component_name {
            return Err(SlideMediaDataError::InvalidSource);
        }
        let selected_owner = owner.object_identifier == selection.movie_identifier;
        let occurrences = validate_owner_archive_info(
            facts,
            object,
            owner.object_identifier,
            identifier,
            selected_owner.then_some(expected_path),
            budget,
        )?;
        if occurrences
            != usize::try_from(owner.count).map_err(|_| SlideMediaDataError::InvalidSource)?
        {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }

    validate_data_metadata_map(package, metadata_payload, facts, budget)
}

fn selected_field_path(
    selection: &MediaSelection,
    identifier: u64,
) -> Result<&'static [u32], SlideMediaDataError> {
    const CONTENT_PATH: [u32; 1] = [super::MOVIE_DATA_FIELD];
    const POSTER_PATH: [u32; 1] = [super::POSTER_IMAGE_DATA_FIELD];
    let content = selection.content_identifier == Some(identifier);
    let poster = selection.poster_identifier == Some(identifier);
    match (content, poster) {
        (true, false) => Ok(&CONTENT_PATH),
        (false, true) => Ok(&POSTER_PATH),
        _ => Err(SlideMediaDataError::InvalidSource),
    }
}

fn unique_current_component<'a>(
    facts: &'a OwnedMetadataFacts,
    locator: &str,
) -> Result<&'a super::ComponentRecord, SlideMediaDataError> {
    let mut selected = None;
    for component in facts
        .components
        .iter()
        .filter(|component| !component.versioned && component.locator.as_ref() == locator)
    {
        if selected.replace(component).is_some() {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    selected.ok_or(SlideMediaDataError::InvalidSource)
}

fn unique_current_reference_count(
    facts: &OwnedMetadataFacts,
    component_identifier: u64,
    data_identifier: u64,
) -> Result<usize, SlideMediaDataError> {
    let mut count = None;
    for &(component, data, owner_count, versioned) in &facts.references {
        if component != component_identifier || data != data_identifier || versioned {
            continue;
        }
        if count.replace(owner_count).is_some() {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    usize::try_from(count.ok_or(SlideMediaDataError::InvalidSource)?)
        .map_err(|_| SlideMediaDataError::InvalidSource)
}

fn validate_owner_archive_info(
    facts: &OwnedMetadataFacts,
    object: &ArchiveObject,
    object_identifier: u64,
    identifier: u64,
    expected_path: Option<&[u32]>,
    budget: &mut MediaBudget,
) -> Result<usize, SlideMediaDataError> {
    if object.archive_info.identifier != Some(object_identifier)
        || object.archive_info.should_merge == Some(true)
        || object.messages.len() != object.archive_info.message_infos.len()
    {
        return Err(SlideMediaDataError::InvalidSource);
    }

    let mut selected_occurrences = 0usize;
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        budget.wire_work(1)?;
        budget.wire_fields(
            1usize
                .checked_add(info.data_references.len())
                .and_then(|value| value.checked_add(info.field_infos.len()))
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )?;
        budget.references(info.data_references.len())?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
            || info.base_message_index.is_some()
            || !info.diff_merge_version.is_empty()
            || info.diff_field_path.is_some()
            || !info.fields_to_remove.is_empty()
            || !info.diff_read_version.is_empty()
        {
            return Err(SlideMediaDataError::InvalidSource);
        }

        for &data_identifier in &info.data_references {
            if data_identifier == 0 || !has_unique_data_record(facts, data_identifier, budget)? {
                return Err(SlideMediaDataError::InvalidSource);
            }
        }

        budget.wire_work(info.data_references.len())?;
        let aggregate_count = info
            .data_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if expected_path.is_some() && aggregate_count != 0 && message.type_ != MOVIE_MESSAGE_TYPE {
            return Err(SlideMediaDataError::InvalidSource);
        }
        selected_occurrences = selected_occurrences
            .checked_add(aggregate_count)
            .ok_or(SlideMediaDataError::InvalidSource)?;

        let mut field_occurrences = 0usize;
        for field in &info.field_infos {
            let contains_work = info
                .data_references
                .len()
                .checked_mul(field.data_references.len())
                .and_then(|value| value.checked_add(field.data_references.len()))
                .ok_or(SlideMediaDataError::InvalidSource)?;
            budget.wire_work(contains_work)?;
            for &data_identifier in &field.data_references {
                if !info.data_references.contains(&data_identifier)
                    || field.r#type != Some(FieldType::DataReference)
                {
                    return Err(SlideMediaDataError::InvalidSource);
                }
            }
            let selected_in_field = field
                .data_references
                .iter()
                .filter(|candidate| **candidate == identifier)
                .count();
            if selected_in_field != 0 {
                if field.r#type != Some(FieldType::DataReference) {
                    return Err(SlideMediaDataError::InvalidSource);
                }
                if expected_path.is_some_and(|path| field.path.as_slice() != path) {
                    return Err(SlideMediaDataError::InvalidSource);
                }
                field_occurrences = field_occurrences
                    .checked_add(selected_in_field)
                    .ok_or(SlideMediaDataError::InvalidSource)?;
            }
        }
        if expected_path.is_some() && field_occurrences != 0 && field_occurrences != aggregate_count
        {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    if selected_occurrences == 0 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(selected_occurrences)
}

fn has_unique_data_record(
    facts: &OwnedMetadataFacts,
    identifier: u64,
    budget: &mut MediaBudget,
) -> Result<bool, SlideMediaDataError> {
    budget.wire_work(facts.records.len())?;
    Ok(facts
        .records
        .iter()
        .filter(|record| record.identifier == identifier)
        .count()
        == 1)
}

fn validate_data_metadata_map(
    package: &Package,
    metadata_payload: &[u8],
    facts: &OwnedMetadataFacts,
    budget: &mut MediaBudget,
) -> Result<(), SlideMediaDataError> {
    let limits = package
        .semantic_wire_limits()
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.input_bytes(metadata_payload.len())?;
    budget.wire_work(metadata_payload.len())?;
    budget.wire_nesting(1)?;
    let metadata = WireView::parse_with_limits(metadata_payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(metadata.len())?;
    let mut map_identifier = None;
    for field in metadata.fields() {
        if field.number() != DATA_METADATA_MAP_FIELD {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        if map_identifier.is_some() {
            return Err(SlideMediaDataError::InvalidSource);
        }
        budget.input_bytes(field.payload().len())?;
        budget.wire_work(field.payload().len())?;
        budget.wire_nesting(1)?;
        map_identifier = Some(parse_local_reference(field.payload(), limits, budget)?);
    }
    let Some(map_identifier) = map_identifier else {
        return Ok(());
    };

    let (_component, map_object) = package
        .object_with_component(map_identifier)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    validate_archive_object_shape(map_object, map_identifier, budget)?;
    let mut map_payload = None;
    for message in &map_object.messages {
        if message.type_ != DATA_METADATA_MAP_MESSAGE_TYPE {
            continue;
        }
        if map_payload.replace(message.data.as_slice()).is_some() {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    let map_payload = map_payload.ok_or(SlideMediaDataError::InvalidSource)?;
    validate_data_metadata_map_payload(package, facts, map_payload, limits, budget)
}

fn validate_archive_object_shape(
    object: &ArchiveObject,
    expected_identifier: u64,
    budget: &mut MediaBudget,
) -> Result<(), SlideMediaDataError> {
    if object.archive_info.identifier != Some(expected_identifier)
        || object.archive_info.should_merge == Some(true)
        || object.messages.len() != object.archive_info.message_infos.len()
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        budget.wire_work(1)?;
        budget.wire_fields(
            1usize
                .checked_add(info.data_references.len())
                .and_then(|value| value.checked_add(info.field_infos.len()))
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )?;
        budget.references(info.data_references.len())?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
            || info.base_message_index.is_some()
            || !info.diff_merge_version.is_empty()
            || info.diff_field_path.is_some()
            || !info.fields_to_remove.is_empty()
            || !info.diff_read_version.is_empty()
        {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_data_metadata_map_payload(
    package: &Package,
    facts: &OwnedMetadataFacts,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut MediaBudget,
) -> Result<(), SlideMediaDataError> {
    budget.input_bytes(payload.len())?;
    budget.wire_work(payload.len())?;
    budget.wire_nesting(1)?;
    let map = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(map.len())?;
    budget.allocation(
        map.len()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    let mut data_identifiers = Vec::new();
    data_identifiers
        .try_reserve_exact(map.len())
        .map_err(|_| SlideMediaDataError::Allocation { amount: map.len() })?;
    for field in map.fields() {
        if field.number() != DATA_METADATA_MAP_ENTRY_FIELD {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        let entry = WireView::parse_with_limits(field.payload(), limits)
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        budget.input_bytes(field.payload().len())?;
        budget.wire_work(field.payload().len())?;
        budget.wire_nesting(2)?;
        budget.wire_fields(entry.len())?;
        let mut data_identifier = None;
        let mut metadata_identifier = None;
        for entry_field in entry.fields() {
            match entry_field.number() {
                DATA_METADATA_MAP_ENTRY_DATA_FIELD => {
                    if data_identifier.is_some() {
                        return Err(SlideMediaDataError::InvalidSource);
                    }
                    data_identifier = Some(parse_nonzero_varint(
                        entry_field,
                        "DataMetadataMap data identifier",
                        budget,
                    )?);
                },
                DATA_METADATA_MAP_ENTRY_METADATA_FIELD => {
                    if metadata_identifier.is_some() {
                        return Err(SlideMediaDataError::InvalidSource);
                    }
                    entry_field
                        .validate_canonical_framing()
                        .map_err(|_| SlideMediaDataError::InvalidSource)?;
                    budget.input_bytes(entry_field.payload().len())?;
                    budget.wire_work(entry_field.payload().len())?;
                    budget.wire_nesting(2)?;
                    metadata_identifier = Some(parse_local_reference(
                        entry_field.payload(),
                        limits,
                        budget,
                    )?);
                },
                _ => {},
            }
        }
        let data_identifier = data_identifier.ok_or(SlideMediaDataError::InvalidSource)?;
        let metadata_identifier = metadata_identifier.ok_or(SlideMediaDataError::InvalidSource)?;
        budget.wire_work(data_identifiers.len())?;
        if data_identifiers.contains(&data_identifier)
            || !has_unique_data_record(facts, data_identifier, budget)?
        {
            return Err(SlideMediaDataError::InvalidSource);
        }
        data_identifiers.push(data_identifier);

        let (_component, metadata_object) = package
            .object_with_component(metadata_identifier)
            .ok_or(SlideMediaDataError::InvalidSource)?;
        validate_archive_object_shape(metadata_object, metadata_identifier, budget)?;
        if metadata_object
            .messages
            .iter()
            .filter(|message| message.type_ == DATA_METADATA_MESSAGE_TYPE)
            .count()
            != 1
        {
            return Err(SlideMediaDataError::InvalidSource);
        }
    }
    Ok(())
}

fn parse_local_reference(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut MediaBudget,
) -> Result<u64, SlideMediaDataError> {
    budget.input_bytes(payload.len())?;
    budget.wire_work(payload.len())?;
    budget.wire_nesting(1)?;
    let reference = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(reference.len())?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut external = None;
    for field in reference.fields() {
        match field.number() {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(SlideMediaDataError::InvalidSource);
                }
                identifier = Some(parse_nonzero_varint(field, "Reference identifier", budget)?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type_seen {
                    return Err(SlideMediaDataError::InvalidSource);
                }
                deprecated_type_seen = true;
                parse_canonical_int32(field, "Reference deprecated type", budget)?;
            },
            REFERENCE_EXTERNAL_FIELD => {
                if external.is_some() {
                    return Err(SlideMediaDataError::InvalidSource);
                }
                external = Some(parse_canonical_bool(field, "Reference external", budget)?);
            },
            _ => {},
        }
    }
    if external == Some(true) {
        return Err(SlideMediaDataError::InvalidSource);
    }
    identifier.ok_or(SlideMediaDataError::InvalidSource)
}

fn parse_nonzero_varint(
    field: WireFieldView<'_>,
    _context: &'static str,
    budget: &mut MediaBudget,
) -> Result<u64, SlideMediaDataError> {
    let value = parse_canonical_varint(field, budget)?;
    if value == 0 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(value)
}

fn parse_canonical_varint(
    field: WireFieldView<'_>,
    budget: &mut MediaBudget,
) -> Result<u64, SlideMediaDataError> {
    budget.wire_work(field.payload().len())?;
    budget.wire_fields(1)?;
    field
        .validate_canonical_key()
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if field.wire_type() != 0 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let (value, consumed) = decode_varint_from_bytes(field.payload())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if consumed != field.payload().len() || consumed != encoded_len(value) {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(value)
}

fn parse_canonical_int32(
    field: WireFieldView<'_>,
    _context: &'static str,
    budget: &mut MediaBudget,
) -> Result<(), SlideMediaDataError> {
    let value = parse_canonical_varint(field, budget)?;
    if value > i32::MAX as u64 && value < 0xffff_ffff_8000_0000 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(())
}

fn parse_canonical_bool(
    field: WireFieldView<'_>,
    _context: &'static str,
    budget: &mut MediaBudget,
) -> Result<bool, SlideMediaDataError> {
    match parse_canonical_varint(field, budget)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(SlideMediaDataError::InvalidSource),
    }
}
