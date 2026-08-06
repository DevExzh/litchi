//! Transactional native hyperlink CRUD for TSWP text-storage objects.

use std::collections::HashSet;

use crate::archive::Archive;
use crate::package_metadata::{
    component_identifier_for_entry, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::hyperlink_object::{
    new_hyperlink_object, patch_hyperlink_target, validate_hyperlink_object,
};
use super::hyperlink_storage::{
    Boundary, RangedObjectTable, add_range, decoded_boundaries, encode_table,
    ensure_range_available, locate_storage_with_archive, patch_ranged_object_table, raw_boundaries,
    remove_range, validate_range,
};
use super::hyperlink_types::{TextHyperlink, TextHyperlinkId, TextHyperlinkTarget};
use super::smart_field_object::{
    ensure_no_metadata_reference, require_exclusive_storage_reference,
};
use super::storage_wire::{StorageLocation, text_utf16_len};
use litchi_iwa_text::position::{TextPosition, TextRange};

/// Read every native hyperlink in a storage, ordered by text position.
pub(crate) fn text_hyperlinks(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextHyperlink>> {
    let located = locate_storage_with_archive(package, storage_id, RangedObjectTable::SmartField)?;
    let location = &located.location;
    let boundaries = decoded_boundaries(storage_id, location, RangedObjectTable::SmartField)?;
    collect_hyperlinks(storage_id, location, &located.archive, &boundaries)
}

/// Create a hyperlink over a currently unoccupied smart-field range.
pub(crate) fn add_text_hyperlink(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    target: &TextHyperlinkTarget,
) -> Result<TextHyperlink> {
    let located = locate_storage_with_archive(package, storage_id, RangedObjectTable::SmartField)?;
    let location = &located.location;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location, RangedObjectTable::SmartField)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        None,
        &location.storage.text,
        RangedObjectTable::SmartField,
    )?;

    let identifier = next_object_identifier(package)?;
    let id = TextHyperlinkId::from_native(identifier);
    let hyperlink_object = new_hyperlink_object(identifier, target)?;
    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        located,
        storage_id,
        RangedObjectTable::SmartField,
        |table, storage| {
            let mut boundaries =
                raw_boundaries(storage_id, table, storage, RangedObjectTable::SmartField)?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                RangedObjectTable::SmartField,
            )?;
            add_range(&mut boundaries, range, identifier)?;
            encode_table(table, boundaries).map(|table| (Some(table), Some(identifier), None))
        },
    )?;
    staged.update_archive(&archive_name, |archive| {
        Ok(archive.insert_object(hyperlink_object)?)
    })?;
    set_package_last_object_identifier(&mut staged, identifier)?;
    let verified = roundtrip(&staged)?;
    let created = hyperlink_by_id(&verified, storage_id, id)?;
    if created.range != range || created.target != *target {
        return Err(Error::InvalidFormat(
            "iWork hyperlink creation failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

/// Atomically change a hyperlink's range and target without changing its ID.
pub(crate) fn update_text_hyperlink(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHyperlinkId,
    range: TextRange,
    target: &TextHyperlinkTarget,
) -> Result<TextHyperlink> {
    let current = hyperlink_by_id(package, storage_id, id)?;
    if current.range == range && current.target == *target {
        return Ok(current);
    }
    let located = locate_storage_with_archive(package, storage_id, RangedObjectTable::SmartField)?;
    let location = &located.location;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location, RangedObjectTable::SmartField)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        Some(id.object_id()),
        &location.storage.text,
        RangedObjectTable::SmartField,
    )?;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "hyperlink")?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        located,
        storage_id,
        RangedObjectTable::SmartField,
        |table, storage| {
            let mut boundaries =
                raw_boundaries(storage_id, table, storage, RangedObjectTable::SmartField)?;
            remove_range(
                &mut boundaries,
                id.object_id(),
                RangedObjectTable::SmartField,
            )?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                RangedObjectTable::SmartField,
            )?;
            add_range(&mut boundaries, range, id.object_id())?;
            encode_table(table, boundaries).map(|table| (Some(table), None, None))
        },
    )?;
    patch_hyperlink_target(&mut staged, &archive_name, id.object_id(), target)?;
    let verified = roundtrip(&staged)?;
    let updated = hyperlink_by_id(&verified, storage_id, id)?;
    if updated.range != range || updated.target != *target {
        return Err(Error::InvalidFormat(
            "iWork hyperlink update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(updated)
}

/// Delete one hyperlink and reclaim its owned smart-field object.
pub(crate) fn remove_text_hyperlink(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHyperlinkId,
) -> Result<TextHyperlink> {
    let removed = hyperlink_by_id(package, storage_id, id)?;
    let located = locate_storage_with_archive(package, storage_id, RangedObjectTable::SmartField)?;
    let location = &located.location;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "hyperlink")?;
    let registered_component = component_identifier_for_object_uuid(package, id.object_id())?;
    let owning_component = component_identifier_for_entry(package, &location.archive_name)?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        located,
        storage_id,
        RangedObjectTable::SmartField,
        |table, storage| {
            let mut boundaries =
                raw_boundaries(storage_id, table, storage, RangedObjectTable::SmartField)?;
            remove_range(
                &mut boundaries,
                id.object_id(),
                RangedObjectTable::SmartField,
            )?;
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
    ensure_no_metadata_reference(&staged, id.object_id(), "hyperlink")?;
    staged.update_archive(&archive_name, |archive| {
        let object = archive.remove_object(id.object_id()).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork hyperlink object {} disappeared during deletion",
                id.object_id()
            ))
        })?;
        validate_hyperlink_object(id.object_id(), &object).map(|_| ())
    })?;
    if let Some(component) = owning_component {
        remove_component_external_references_to_object(&mut staged, component, id.object_id())?;
    }
    if let Some(component) = registered_component {
        remove_component_object_uuids(&mut staged, component, &[id.object_id()])?;
    }
    release_package_identifier_suffix(&mut staged, &[id.object_id()])?;
    let verified = roundtrip(&staged)?;
    if text_hyperlinks(&verified, storage_id)?
        .iter()
        .any(|hyperlink| hyperlink.id == id)
    {
        return Err(Error::InvalidFormat(
            "iWork hyperlink deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(removed)
}

/// Reclaim hyperlink objects whose table references disappeared during text replacement.
pub(crate) fn remove_unreferenced_hyperlink_objects(
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
            |(identifier, object)| match validate_hyperlink_object(identifier, object) {
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
        ensure_no_metadata_reference(package, *identifier, "hyperlink")?;
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
                    "iWork hyperlink object {identifier} disappeared during text replacement"
                ))
            })?;
            validate_hyperlink_object(*identifier, &object).map(|_| ())
        })?;
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn hyperlink_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    id: TextHyperlinkId,
) -> Result<TextHyperlink> {
    let matches = text_hyperlinks(package, storage_id)?
        .into_iter()
        .filter(|hyperlink| hyperlink.id == id)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [hyperlink] => Ok(hyperlink.clone()),
        [] => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own hyperlink object {}",
            id.object_id()
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references hyperlink object {} more than once",
            id.object_id()
        ))),
    }
}

fn collect_hyperlinks(
    storage_id: u64,
    location: &StorageLocation,
    archive: &Archive,
    boundaries: &[Boundary],
) -> Result<Vec<TextHyperlink>> {
    let text_len = text_utf16_len(&location.storage.text)?;
    let mut seen = HashSet::new();
    let mut hyperlinks = Vec::new();
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
        let Some(target) = validate_hyperlink_object(identifier, object)? else {
            continue;
        };
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references hyperlink object {identifier} more than once"
            )));
        }
        hyperlinks.push(TextHyperlink::new(
            TextHyperlinkId::from_native(identifier),
            TextRange::new(
                TextPosition::from_utf16_code_units(boundary.index),
                TextPosition::from_utf16_code_units(end),
            )?,
            target,
        ));
    }
    Ok(hyperlinks)
}

fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}

#[cfg(test)]
#[path = "hyperlink_internal_tests.rs"]
mod tests;
