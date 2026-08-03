//! Transactional native number-attachment CRUD for TSWP text storage.

use std::collections::HashSet;

use crate::package_metadata::{
    component_identifier_for_entry, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::editor::replace_storage_text;
use super::number_attachment_object::{
    new_number_attachment_object, patch_number_attachment_settings,
    validate_number_attachment_object,
};
use super::number_attachment_storage::{
    decoded_attachment_entries, insert_attachment_reference, locate_attachment_storage,
    locate_attachment_storage_with_archive,
};
use super::number_attachment_types::{
    TextNumberAttachment, TextNumberAttachmentId, TextNumberAttachmentSettings,
};
use super::position::TextPosition;
use super::smart_field_object::{
    ensure_no_metadata_reference, require_exclusive_storage_reference,
};

const OBJECT_REPLACEMENT_CHARACTER: &str = "\u{fffc}";

pub(crate) fn text_number_attachments(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextNumberAttachment>> {
    let located = locate_attachment_storage_with_archive(package, storage_id)?;
    let entries = decoded_attachment_entries(storage_id, &located.location)?;
    let archive = &located.archive;
    let mut seen = HashSet::new();
    let mut attachments = Vec::new();
    for entry in entries {
        let object = archive.object(entry.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references missing attachment object {}",
                entry.object_id
            ))
        })?;
        let Some(settings) = validate_number_attachment_object(entry.object_id, object)? else {
            continue;
        };
        if !seen.insert(entry.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references number-attachment object {} more than once",
                entry.object_id
            )));
        }
        attachments.push(TextNumberAttachment::new(
            TextNumberAttachmentId::from_native(entry.object_id),
            TextPosition::from_native(entry.index),
            settings,
        ));
    }
    Ok(attachments)
}

pub(crate) fn insert_text_number_attachment(
    package: &mut IWorkPackage,
    storage_id: u64,
    position: TextPosition,
    settings: &TextNumberAttachmentSettings,
) -> Result<TextNumberAttachment> {
    let position_usize = usize::try_from(position.utf16_index())
        .map_err(|_| Error::ParseError("number-attachment position exceeds usize".to_owned()))?;
    let location = locate_attachment_storage(package, storage_id)?;
    decoded_attachment_entries(storage_id, &location)?;
    let mut staged = package.clone();
    replace_storage_text(
        &mut staged,
        storage_id,
        position_usize..position_usize,
        OBJECT_REPLACEMENT_CHARACTER,
    )?;
    let located = locate_attachment_storage_with_archive(&staged, storage_id)?;
    decoded_attachment_entries(storage_id, &located.location)?;
    let identifier = next_object_identifier(&staged)?;
    let object = new_number_attachment_object(identifier, settings)?;
    insert_attachment_reference(
        &mut staged,
        located,
        storage_id,
        position.utf16_index(),
        identifier,
        object,
    )?;
    set_package_last_object_identifier(&mut staged, identifier)?;

    let verified = roundtrip(&staged)?;
    let created = number_attachment_by_id(
        &verified,
        storage_id,
        TextNumberAttachmentId::from_native(identifier),
    )?;
    if created.position != position || created.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork number-attachment insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

pub(crate) fn update_text_number_attachment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextNumberAttachmentId,
    settings: &TextNumberAttachmentSettings,
) -> Result<TextNumberAttachment> {
    let current = number_attachment_by_id(package, storage_id, id)?;
    if current.settings == *settings {
        return Ok(current);
    }
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "number attachment")?;
    let located = locate_attachment_storage_with_archive(package, storage_id)?;
    let mut staged = package.clone();
    patch_number_attachment_settings(&mut staged, located, id.object_id(), settings)?;
    let verified = roundtrip(&staged)?;
    let updated = number_attachment_by_id(&verified, storage_id, id)?;
    if updated.position != current.position || updated.settings != *settings {
        return Err(Error::InvalidFormat(
            "iWork number-attachment update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(updated)
}

pub(crate) fn remove_text_number_attachment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextNumberAttachmentId,
) -> Result<TextNumberAttachment> {
    let removed = number_attachment_by_id(package, storage_id, id)?;
    require_exclusive_storage_reference(package, storage_id, id.object_id(), "number attachment")?;
    let start = usize::try_from(removed.position.utf16_index())
        .map_err(|_| Error::ParseError("number-attachment position exceeds usize".to_owned()))?;
    let end = start
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("number-attachment range overflow".to_owned()))?;
    let mut staged = package.clone();
    replace_storage_text(&mut staged, storage_id, start..end, "")?;
    let verified = roundtrip(&staged)?;
    if text_number_attachments(&verified, storage_id)?
        .iter()
        .any(|attachment| attachment.id == id)
    {
        return Err(Error::InvalidFormat(
            "iWork number-attachment deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(removed)
}

pub(crate) fn remove_unreferenced_number_attachment_objects(
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
        .filter_map(|(identifier, object)| {
            match validate_number_attachment_object(identifier, object) {
                Ok(Some(_)) => Some(Ok(identifier)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    if identifiers.is_empty() {
        return Ok(());
    }
    identifiers.sort_unstable_by(|left, right| right.cmp(left));
    for identifier in &identifiers {
        ensure_no_metadata_reference(package, *identifier, "number attachment")?;
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
                    "iWork number-attachment object {identifier} disappeared during text replacement"
                ))
            })?;
            validate_number_attachment_object(*identifier, &object).map(|_| ())
        })?;
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn number_attachment_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    id: TextNumberAttachmentId,
) -> Result<TextNumberAttachment> {
    let mut matches = text_number_attachments(package, storage_id)?
        .into_iter()
        .filter(|attachment| attachment.id == id);
    let Some(attachment) = matches.next() else {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own number-attachment object {}",
            id.object_id()
        )));
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references number-attachment object {} more than once",
            id.object_id()
        )));
    }
    Ok(attachment)
}

fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}

#[cfg(test)]
#[path = "number_attachment_internal_tests.rs"]
mod tests;
