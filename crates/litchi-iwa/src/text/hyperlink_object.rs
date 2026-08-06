//! Native hyperlink smart-field object encoding, mutation, and ownership checks.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::tswp;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

use super::smart_field_object::{generated_text_attribute_uuid, validate_text_attribute_uuid};
use litchi_iwa_text::hyperlink::TextHyperlinkTarget;

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
    validate_text_attribute_uuid(identifier, "hyperlink", uuid)?;
    let target = hyperlink.url_ref.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork hyperlink object {identifier} is missing its target"
        ))
    })?;
    Ok(Some(TextHyperlinkTarget::from_boxed(
        target.into_boxed_str(),
    )?))
}

pub(super) fn new_hyperlink_object(
    identifier: u64,
    target: &TextHyperlinkTarget,
) -> Result<ArchiveObject> {
    let uuid = generated_text_attribute_uuid()?;
    let hyperlink = tswp::HyperlinkFieldArchive {
        super_: Some(tswp::SmartFieldArchive {
            text_attribute_uuid_string: Some(uuid),
        }),
        url_ref: Some(target.as_str().to_owned()),
    };
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: HYPERLINK_MESSAGE_TYPE,
            data: hyperlink.encode_to_vec(),
        }],
    )?)
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
