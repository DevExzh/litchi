//! Transactional native Date & Time smart-field CRUD for TSWP text storage.

use std::collections::HashSet;

use crate::archive::Archive;
use crate::package_metadata::{
    component_identifier_for_entry, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::date_time_object::{
    new_date_time_object, patch_date_time_settings, validate_date_time_object,
};
use super::date_time_types::{
    TextDateTimeDisplayText, TextDateTimeField, TextDateTimeFieldId, TextDateTimeFieldSettings,
};
use super::editor::replace_storage_text;
use super::hyperlink_storage::{
    Boundary, RangedObjectTable, add_range, decoded_boundaries, encode_table,
    ensure_range_available, locate_storage, locate_storage_with_archive, patch_ranged_object_table,
    raw_boundaries, remove_range, validate_range,
};
use super::smart_field_object::{
    ensure_no_metadata_reference, require_exclusive_storage_reference,
};
use super::storage_wire::{StorageLocation, text_utf16_len};
use litchi_iwa_text::position::{TextPosition, TextRange};

const SMART_FIELD_TABLE: RangedObjectTable = RangedObjectTable::SmartField;

pub(crate) fn text_date_time_fields(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextDateTimeField>> {
    let located = locate_storage_with_archive(package, storage_id, SMART_FIELD_TABLE)?;
    let location = &located.location;
    let boundaries = decoded_boundaries(storage_id, location, SMART_FIELD_TABLE)?;
    collect_date_time_fields(storage_id, location, &located.archive, &boundaries)
}

pub(crate) fn add_text_date_time_field(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    settings: &TextDateTimeFieldSettings,
) -> Result<TextDateTimeField> {
    let mut staged = package.clone();
    let id = add_date_time_field_in_place(&mut staged, storage_id, range, settings)?;
    let verified = roundtrip(&staged)?;
    let created = date_time_field_by_id(&verified, storage_id, id)?;
    if created.range != range || created.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork Date & Time field creation failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

pub(crate) fn insert_text_date_time_field(
    package: &mut IWorkPackage,
    storage_id: u64,
    position: TextPosition,
    display_text: &TextDateTimeDisplayText,
    settings: &TextDateTimeFieldSettings,
) -> Result<TextDateTimeField> {
    let start = usize::try_from(position.utf16_index())
        .map_err(|_| Error::ParseError("Date & Time position exceeds usize".to_owned()))?;
    let display_units = display_text.as_str().encode_utf16().count();
    let end = start
        .checked_add(display_units)
        .ok_or_else(|| Error::ParseError("Date & Time range overflow".to_owned()))?;
    let range = TextRange::from_utf16_indexes(start, end)?;
    let mut staged = package.clone();
    replace_storage_text(&mut staged, storage_id, start..start, display_text.as_str())?;
    let id = add_date_time_field_in_place(&mut staged, storage_id, range, settings)?;
    let verified = roundtrip(&staged)?;
    let created = date_time_field_by_id(&verified, storage_id, id)?;
    if created.range != range
        || created.settings != *settings
        || text_in_range(&verified, storage_id, range)? != display_text.as_str()
    {
        return Err(Error::InvalidFormat(
            "iWork Date & Time insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

fn add_date_time_field_in_place(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    settings: &TextDateTimeFieldSettings,
) -> Result<TextDateTimeFieldId> {
    let located = locate_storage_with_archive(package, storage_id, SMART_FIELD_TABLE)?;
    let location = &located.location;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location, SMART_FIELD_TABLE)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        None,
        &location.storage.text,
        SMART_FIELD_TABLE,
    )?;
    let identifier = next_object_identifier(package)?;
    let object = new_date_time_object(identifier, settings)?;
    let archive_name = location.archive_name.clone();
    patch_ranged_object_table(
        package,
        located,
        storage_id,
        SMART_FIELD_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, SMART_FIELD_TABLE)?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                SMART_FIELD_TABLE,
            )?;
            add_range(&mut boundaries, range, identifier)?;
            encode_table(table, boundaries).map(|table| (Some(table), Some(identifier), None))
        },
    )?;
    package.update_archive(&archive_name, |archive| Ok(archive.insert_object(object)?))?;
    set_package_last_object_identifier(package, identifier)?;
    Ok(TextDateTimeFieldId::from_native(identifier))
}

pub(crate) fn update_text_date_time_field(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextDateTimeFieldId,
    range: TextRange,
    settings: &TextDateTimeFieldSettings,
) -> Result<TextDateTimeField> {
    let current = date_time_field_by_id(package, storage_id, id)?;
    if current.range == range && current.settings == *settings {
        return Ok(current);
    }
    let located = locate_storage_with_archive(package, storage_id, SMART_FIELD_TABLE)?;
    let location = &located.location;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location, SMART_FIELD_TABLE)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        Some(id.object_id()),
        &location.storage.text,
        SMART_FIELD_TABLE,
    )?;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "Date & Time")?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        located,
        storage_id,
        SMART_FIELD_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, SMART_FIELD_TABLE)?;
            remove_range(&mut boundaries, id.object_id(), SMART_FIELD_TABLE)?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                SMART_FIELD_TABLE,
            )?;
            add_range(&mut boundaries, range, id.object_id())?;
            encode_table(table, boundaries).map(|table| (Some(table), None, None))
        },
    )?;
    patch_date_time_settings(&mut staged, &archive_name, id.object_id(), settings)?;
    let verified = roundtrip(&staged)?;
    let updated = date_time_field_by_id(&verified, storage_id, id)?;
    if updated.range != range || updated.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork Date & Time field update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(updated)
}

pub(crate) fn remove_text_date_time_field(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextDateTimeFieldId,
) -> Result<TextDateTimeField> {
    let removed = date_time_field_by_id(package, storage_id, id)?;
    let located = locate_storage_with_archive(package, storage_id, SMART_FIELD_TABLE)?;
    let location = &located.location;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "Date & Time")?;
    let registered_component = component_identifier_for_object_uuid(package, id.object_id())?;
    let owning_component = component_identifier_for_entry(package, &location.archive_name)?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        located,
        storage_id,
        SMART_FIELD_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, SMART_FIELD_TABLE)?;
            remove_range(&mut boundaries, id.object_id(), SMART_FIELD_TABLE)?;
            if boundaries
                .iter()
                .any(|boundary| boundary.object_id.is_some())
            {
                encode_table(table, boundaries)
                    .map(|table| (Some(table), None, Some(id.object_id())))
            } else {
                Ok((None, None, Some(id.object_id())))
            }
        },
    )?;
    ensure_no_metadata_reference(&staged, id.object_id(), "Date & Time")?;
    staged.update_archive(&archive_name, |archive| {
        let object = archive.remove_object(id.object_id()).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork Date & Time object {} disappeared during deletion",
                id.object_id()
            ))
        })?;
        validate_date_time_object(id.object_id(), &object).map(|_| ())
    })?;
    if let Some(component) = owning_component {
        remove_component_external_references_to_object(&mut staged, component, id.object_id())?;
    }
    if let Some(component) = registered_component {
        remove_component_object_uuids(&mut staged, component, &[id.object_id()])?;
    }
    release_package_identifier_suffix(&mut staged, &[id.object_id()])?;
    let verified = roundtrip(&staged)?;
    if text_date_time_fields(&verified, storage_id)?
        .iter()
        .any(|field| field.id == id)
    {
        return Err(Error::InvalidFormat(
            "iWork Date & Time field deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(removed)
}

pub(crate) fn remove_unreferenced_date_time_objects(
    package: &mut IWorkPackage,
    archive_name: &str,
    candidates: &HashSet<u64>,
) -> Result<()> {
    if candidates.is_empty() {
        return Ok(());
    }
    let archive = package.archive(archive_name)?;
    let mut identifiers = candidates
        .iter()
        .filter_map(|identifier| {
            archive
                .object(*identifier)
                .map(|object| (*identifier, object))
        })
        .filter_map(
            |(identifier, object)| match validate_date_time_object(identifier, object) {
                Ok(Some(_)) => Some(Ok(identifier)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            },
        )
        .collect::<Result<Vec<_>>>()?;
    if identifiers.is_empty() {
        return Ok(());
    }
    identifiers.sort_unstable_by(|left, right| right.cmp(left));
    for identifier in &identifiers {
        ensure_no_metadata_reference(package, *identifier, "Date & Time")?;
    }
    let owning_component = component_identifier_for_entry(package, archive_name)?;
    let mut staged = package.clone();
    for identifier in &identifiers {
        if let Some(component) = owning_component {
            remove_component_external_references_to_object(&mut staged, component, *identifier)?;
        }
        if let Some(component) = component_identifier_for_object_uuid(&staged, *identifier)? {
            remove_component_object_uuids(&mut staged, component, &[*identifier])?;
        }
        staged.update_archive(archive_name, |archive| {
            let object = archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork Date & Time object {identifier} disappeared during text replacement"
                ))
            })?;
            validate_date_time_object(*identifier, &object).map(|_| ())
        })?;
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn date_time_field_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    id: TextDateTimeFieldId,
) -> Result<TextDateTimeField> {
    let matches = text_date_time_fields(package, storage_id)?
        .into_iter()
        .filter(|field| field.id == id)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [field] => Ok(field.clone()),
        [] => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own Date & Time object {}",
            id.object_id()
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references Date & Time object {} more than once",
            id.object_id()
        ))),
    }
}

fn collect_date_time_fields(
    storage_id: u64,
    location: &StorageLocation,
    archive: &Archive,
    boundaries: &[Boundary],
) -> Result<Vec<TextDateTimeField>> {
    let text_len = text_utf16_len(&location.storage.text)?;
    let mut seen = HashSet::new();
    let mut fields = Vec::new();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(identifier) = boundary.object_id else {
            continue;
        };
        let end = boundaries
            .get(position + 1)
            .map_or(text_len, |next| next.index);
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references missing smart-field object {identifier}"
            ))
        })?;
        let Some(settings) = validate_date_time_object(identifier, object)? else {
            continue;
        };
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references Date & Time object {identifier} more than once"
            )));
        }
        fields.push(TextDateTimeField::new(
            TextDateTimeFieldId::from_native(identifier),
            TextRange::new(
                TextPosition::from_utf16_code_units(boundary.index),
                TextPosition::from_utf16_code_units(end),
            )?,
            settings,
        ));
    }
    Ok(fields)
}

fn text_in_range(package: &IWorkPackage, storage_id: u64, range: TextRange) -> Result<String> {
    let location = locate_storage(package, storage_id, SMART_FIELD_TABLE)?;
    let units = location
        .storage
        .text
        .concat()
        .encode_utf16()
        .collect::<Vec<_>>();
    let start = usize::try_from(range.start().utf16_index())
        .map_err(|_| Error::InvalidFormat("Date & Time range start overflow".to_owned()))?;
    let end = usize::try_from(range.end().utf16_index())
        .map_err(|_| Error::InvalidFormat("Date & Time range end overflow".to_owned()))?;
    let selected = units.get(start..end).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Date & Time range {start}..{end} exceeds the text storage"
        ))
    })?;
    String::from_utf16(selected).map_err(|_| {
        Error::InvalidFormat("Date & Time range splits a UTF-16 surrogate pair".to_owned())
    })
}

fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}

#[cfg(test)]
#[path = "date_time_internal_tests.rs"]
mod tests;
