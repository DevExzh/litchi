//! Transactional native ranged-bookmark CRUD for TSWP text storage.

use std::collections::HashSet;

use crate::package_metadata::{
    component_identifier_for_entry, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::bookmark_object::{
    new_bookmark_object, patch_bookmark_settings, validate_bookmark_object,
};
use super::bookmark_types::{TextBookmark, TextBookmarkId, TextBookmarkSettings};
use super::hyperlink_storage::{
    Boundary, RangedObjectTable, add_range, decoded_boundaries, encode_table,
    ensure_range_available, locate_storage, patch_ranged_object_table, raw_boundaries,
    remove_range, validate_range,
};
use super::position::{TextPosition, TextRange};
use super::smart_field_object::{
    ensure_no_metadata_reference, require_exclusive_storage_reference,
};
use super::storage_wire::{StorageLocation, text_utf16_len};

const BOOKMARK_TABLE: RangedObjectTable = RangedObjectTable::Bookmark;

pub(crate) fn text_bookmarks(package: &IWorkPackage, storage_id: u64) -> Result<Vec<TextBookmark>> {
    let location = locate_storage(package, storage_id, BOOKMARK_TABLE)?;
    let boundaries = decoded_boundaries(storage_id, &location, BOOKMARK_TABLE)?;
    collect_bookmarks(package, storage_id, &location, &boundaries)
}

pub(crate) fn add_text_bookmark(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    settings: &TextBookmarkSettings,
) -> Result<TextBookmark> {
    let location = locate_storage(package, storage_id, BOOKMARK_TABLE)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, &location, BOOKMARK_TABLE)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        None,
        &location.storage.text,
        BOOKMARK_TABLE,
    )?;

    let identifier = next_object_identifier(package)?;
    let id = TextBookmarkId::from_native(identifier);
    let bookmark_object = new_bookmark_object(identifier, settings)?;
    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        &location,
        storage_id,
        BOOKMARK_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, BOOKMARK_TABLE)?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                BOOKMARK_TABLE,
            )?;
            add_range(&mut boundaries, range, identifier)?;
            encode_table(table, boundaries).map(|table| (Some(table), Some(identifier), None))
        },
    )?;
    staged.update_archive(&location.archive_name, |archive| {
        archive.insert_object(bookmark_object)
    })?;
    set_package_last_object_identifier(&mut staged, identifier)?;
    let verified = roundtrip(&staged)?;
    let created = bookmark_by_id(&verified, storage_id, id)?;
    if created.range != range || created.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork bookmark creation failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

pub(crate) fn update_text_bookmark(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextBookmarkId,
    range: TextRange,
    settings: &TextBookmarkSettings,
) -> Result<TextBookmark> {
    let current = bookmark_by_id(package, storage_id, id)?;
    if current.range == range && current.settings == *settings {
        return Ok(current);
    }
    let location = locate_storage(package, storage_id, BOOKMARK_TABLE)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, &location, BOOKMARK_TABLE)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        Some(id.object_id()),
        &location.storage.text,
        BOOKMARK_TABLE,
    )?;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "bookmark")?;

    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        &location,
        storage_id,
        BOOKMARK_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, BOOKMARK_TABLE)?;
            remove_range(&mut boundaries, id.object_id(), BOOKMARK_TABLE)?;
            ensure_range_available(
                storage_id,
                range,
                &boundaries,
                None,
                &storage.text,
                BOOKMARK_TABLE,
            )?;
            add_range(&mut boundaries, range, id.object_id())?;
            encode_table(table, boundaries).map(|table| (Some(table), None, None))
        },
    )?;
    patch_bookmark_settings(
        &mut staged,
        &location.archive_name,
        id.object_id(),
        settings,
    )?;
    let verified = roundtrip(&staged)?;
    let updated = bookmark_by_id(&verified, storage_id, id)?;
    if updated.range != range || updated.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork bookmark update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(updated)
}

pub(crate) fn remove_text_bookmark(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextBookmarkId,
) -> Result<TextBookmark> {
    let removed = bookmark_by_id(package, storage_id, id)?;
    let location = locate_storage(package, storage_id, BOOKMARK_TABLE)?;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "bookmark")?;
    let registered_component = component_identifier_for_object_uuid(package, id.object_id())?;
    let owning_component = component_identifier_for_entry(package, &location.archive_name)?;

    let mut staged = package.clone();
    patch_ranged_object_table(
        &mut staged,
        &location,
        storage_id,
        BOOKMARK_TABLE,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage, BOOKMARK_TABLE)?;
            remove_range(&mut boundaries, id.object_id(), BOOKMARK_TABLE)?;
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
    ensure_no_metadata_reference(&staged, id.object_id(), "bookmark")?;
    staged.update_archive(&location.archive_name, |archive| {
        let object = archive.remove_object(id.object_id()).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork bookmark object {} disappeared during deletion",
                id.object_id()
            ))
        })?;
        validate_bookmark_object(id.object_id(), &object).map(|_| ())
    })?;
    if let Some(component) = owning_component {
        remove_component_external_references_to_object(&mut staged, component, id.object_id())?;
    }
    if let Some(component) = registered_component {
        remove_component_object_uuids(&mut staged, component, &[id.object_id()])?;
    }
    release_package_identifier_suffix(&mut staged, &[id.object_id()])?;
    let verified = roundtrip(&staged)?;
    if text_bookmarks(&verified, storage_id)?
        .iter()
        .any(|bookmark| bookmark.id == id)
    {
        return Err(Error::InvalidFormat(
            "iWork bookmark deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(removed)
}

pub(crate) fn remove_unreferenced_bookmark_objects(
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
            |(identifier, object)| match validate_bookmark_object(identifier, object) {
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
        ensure_no_metadata_reference(package, *identifier, "bookmark")?;
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
                    "iWork bookmark object {identifier} disappeared during text replacement"
                ))
            })?;
            validate_bookmark_object(*identifier, &object).map(|_| ())
        })?;
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn bookmark_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    id: TextBookmarkId,
) -> Result<TextBookmark> {
    let matches = text_bookmarks(package, storage_id)?
        .into_iter()
        .filter(|bookmark| bookmark.id == id)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [bookmark] => Ok(bookmark.clone()),
        [] => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own bookmark object {}",
            id.object_id()
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references bookmark object {} more than once",
            id.object_id()
        ))),
    }
}

fn collect_bookmarks(
    package: &IWorkPackage,
    storage_id: u64,
    location: &StorageLocation,
    boundaries: &[Boundary],
) -> Result<Vec<TextBookmark>> {
    let text_len = text_utf16_len(&location.storage.text)?;
    let archive = package.archive(&location.archive_name)?;
    let mut seen = HashSet::new();
    let mut bookmarks = Vec::new();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(identifier) = boundary.object_id else {
            continue;
        };
        let end = boundaries
            .get(position + 1)
            .map_or(text_len, |next| next.index);
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references missing bookmark object {identifier}"
            ))
        })?;
        let Some(settings) = validate_bookmark_object(identifier, object)? else {
            continue;
        };
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references bookmark object {identifier} more than once"
            )));
        }
        bookmarks.push(TextBookmark::new(
            TextBookmarkId::from_native(identifier),
            TextRange::new(
                TextPosition::from_native(boundary.index),
                TextPosition::from_native(end),
            )?,
            settings,
        ));
    }
    Ok(bookmarks)
}

fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}

#[cfg(test)]
#[path = "bookmark_internal_tests.rs"]
mod tests;
