//! Native hyperlink smart-field object encoding, mutation, and ownership checks.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::tswp;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

use super::hyperlink_types::TextHyperlinkTarget;

const HYPERLINK_TARGET_FIELD: u32 = 2;
pub(super) const HYPERLINK_MESSAGE_TYPE: u32 = 2_032;

pub(super) fn validate_hyperlink_object(
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<TextHyperlinkTarget>> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == HYPERLINK_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if payloads.is_empty() {
        return Ok(None);
    }
    let [message] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork hyperlink object {identifier} contains multiple hyperlink payloads"
        )));
    };
    if object.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork hyperlink object {identifier} contains unrelated payloads"
        )));
    }
    let hyperlink = tswp::HyperlinkFieldArchive::decode(message.data.as_slice())?;
    let uuid = hyperlink
        .super_
        .as_ref()
        .and_then(|smart_field| smart_field.text_attribute_uuid_string.as_deref())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork hyperlink object {identifier} is missing its text-attribute UUID"
            ))
        })?;
    validate_uuid(identifier, uuid)?;
    let target = hyperlink.url_ref.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork hyperlink object {identifier} is missing its target"
        ))
    })?;
    TextHyperlinkTarget::new(target.into_boxed_str()).map(Some)
}

fn validate_uuid(identifier: u64, uuid: &str) -> Result<()> {
    let valid = uuid.len() == 36
        && uuid.bytes().enumerate().all(|(index, byte)| match index {
            8 | 13 | 18 | 23 => byte == b'-',
            _ => byte.is_ascii_hexdigit(),
        });
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "iWork hyperlink object {identifier} has an invalid text-attribute UUID"
        )));
    }
    Ok(())
}

pub(super) fn new_hyperlink_object(
    identifier: u64,
    target: &TextHyperlinkTarget,
) -> Result<ArchiveObject> {
    let braced = litchi_core::id::generate_guid_braced();
    let uuid = braced
        .strip_prefix('{')
        .and_then(|uuid| uuid.strip_suffix('}'))
        .ok_or_else(|| Error::InvalidFormat("generated UUID is not braced".to_owned()))?;
    let hyperlink = tswp::HyperlinkFieldArchive {
        super_: Some(tswp::SmartFieldArchive {
            text_attribute_uuid_string: Some(uuid.to_owned()),
        }),
        url_ref: Some(target.as_str().to_owned()),
    };
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: HYPERLINK_MESSAGE_TYPE,
            data: hyperlink.encode_to_vec(),
        }],
    )
}

pub(super) fn patch_hyperlink_target(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
    target: &TextHyperlinkTarget,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork hyperlink object {identifier} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == HYPERLINK_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork hyperlink object {identifier} must contain exactly one hyperlink payload"
            )));
        };
        if object.messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork hyperlink object {identifier} contains unrelated payloads"
            )));
        }
        let original = &object.messages[*index];
        let hyperlink = tswp::HyperlinkFieldArchive::decode(original.data.as_slice())?;
        let data = patch_length_delimited_field(
            &original.data,
            HYPERLINK_TARGET_FIELD,
            hyperlink.url_ref.is_some(),
            Some(target.as_str().as_bytes()),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: HYPERLINK_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn require_exclusive_storage_reference(
    package: &IWorkPackage,
    storage_id: u64,
    identifier: u64,
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
            "hyperlink object {identifier} must be referenced only by text storage {storage_id}, found {owners:?}"
        )));
    }
    Ok(())
}

pub(super) fn ensure_no_metadata_reference(package: &IWorkPackage, identifier: u64) -> Result<()> {
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
                    "hyperlink object {identifier} retains an indexed reference in {archive_name}"
                )));
            }
        }
    }
    Ok(())
}
