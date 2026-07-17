//! Shared identity and ownership invariants for TSWP smart-field objects.

use crate::{Error, IWorkPackage, Result};

pub(super) fn generated_text_attribute_uuid() -> Result<String> {
    let braced = litchi_core::id::generate_guid_braced();
    braced
        .strip_prefix('{')
        .and_then(|uuid| uuid.strip_suffix('}'))
        .map(str::to_owned)
        .ok_or_else(|| Error::InvalidFormat("generated UUID is not braced".to_owned()))
}

pub(super) fn validate_text_attribute_uuid(identifier: u64, label: &str, uuid: &str) -> Result<()> {
    let valid = uuid.len() == 36
        && uuid.bytes().enumerate().all(|(index, byte)| match index {
            8 | 13 | 18 | 23 => byte == b'-',
            _ => byte.is_ascii_hexdigit(),
        });
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "iWork {label} object {identifier} has an invalid text-attribute UUID"
        )));
    }
    Ok(())
}

pub(super) fn require_exclusive_storage_reference(
    package: &IWorkPackage,
    storage_id: u64,
    identifier: u64,
    label: &str,
) -> Result<()> {
    let mut owners = Vec::new();
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("object in {archive_name} has no identifier"))
            })?;
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&identifier))
            {
                owners.push(object_id);
            }
        }
    }
    if owners != [storage_id] {
        return Err(Error::InvalidFormat(format!(
            "{label} object {identifier} must be referenced only by text storage {storage_id}, found {owners:?}"
        )));
    }
    Ok(())
}

pub(super) fn ensure_no_metadata_reference(
    package: &IWorkPackage,
    identifier: u64,
    label: &str,
) -> Result<()> {
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&identifier)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            }) {
                return Err(Error::InvalidFormat(format!(
                    "{label} object {identifier} retains an indexed reference in {archive_name}"
                )));
            }
        }
    }
    Ok(())
}
