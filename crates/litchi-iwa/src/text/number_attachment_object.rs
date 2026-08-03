//! Native number-attachment encoding and lossless payload mutation.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::tswp;
use crate::wire::{
    patch_length_delimited_field, patch_nested_length_delimited_field, patch_nested_varint_field,
    patch_varint_field,
};
use crate::{Error, IWorkPackage, Result};

use super::number_attachment_types::{
    TextNumberAttachmentFormat, TextNumberAttachmentKind, TextNumberAttachmentSettings,
    TextNumberAttachmentText,
};
use super::storage_wire::{LocatedStorage, update_parsed_archive};

const SUPER_FIELD: u32 = 1;
const STRING_EQUIVALENT_FIELD: u32 = 1;
const KIND_FIELD: u32 = 2;
const NUMBER_FORMAT_FIELD: u32 = 2;
const STRING_VALUE_FIELD: u32 = 3;
const NUMBER_FORMAT_NAME_FIELD: u32 = 4;
pub(super) const NUMBER_ATTACHMENT_MESSAGE_TYPE: u32 = 2_043;

pub(super) fn validate_number_attachment_object(
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<TextNumberAttachmentSettings>> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBER_ATTACHMENT_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if payloads.is_empty() {
        return Ok(None);
    }
    let [message] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork number-attachment object {identifier} contains multiple number payloads"
        )));
    };
    if object.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork number-attachment object {identifier} contains unrelated payloads"
        )));
    }
    let attachment = tswp::NumberAttachmentArchive::decode(message.data.as_slice())?;
    let textual = attachment.super_.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork number-attachment object {identifier} is missing its textual payload"
        ))
    })?;
    let kind = textual.kind.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork number-attachment object {identifier} is missing its kind"
        ))
    })?;
    Ok(Some(TextNumberAttachmentSettings {
        kind: TextNumberAttachmentKind::from_raw(kind),
        string_equivalent: textual
            .string_equivalent
            .map(|value| TextNumberAttachmentText::new(value.into_boxed_str()))
            .transpose()?,
        number_format: attachment
            .number_format
            .map(TextNumberAttachmentFormat::from_native_value),
        string_value: attachment
            .string_value
            .map(|value| TextNumberAttachmentText::new(value.into_boxed_str()))
            .transpose()?,
        number_format_name: attachment
            .number_format_name
            .map(|value| TextNumberAttachmentText::new(value.into_boxed_str()))
            .transpose()?,
    }))
}

pub(super) fn new_number_attachment_object(
    identifier: u64,
    settings: &TextNumberAttachmentSettings,
) -> Result<ArchiveObject> {
    let attachment = tswp::NumberAttachmentArchive {
        super_: Some(tswp::TextualAttachmentArchive {
            string_equivalent: settings
                .string_equivalent
                .as_ref()
                .map(|value| value.as_str().to_owned()),
            kind: Some(settings.kind.as_raw()),
        }),
        number_format: settings.number_format.map(|value| value.native_value()),
        string_value: settings
            .string_value
            .as_ref()
            .map(|value| value.as_str().to_owned()),
        number_format_name: settings
            .number_format_name
            .as_ref()
            .map(|value| value.as_str().to_owned()),
    };
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: NUMBER_ATTACHMENT_MESSAGE_TYPE,
            data: attachment.encode_to_vec(),
        }],
    )
}

pub(super) fn patch_number_attachment_settings(
    package: &mut IWorkPackage,
    located: LocatedStorage,
    identifier: u64,
    settings: &TextNumberAttachmentSettings,
) -> Result<()> {
    let archive_name = located.location.archive_name;
    update_parsed_archive(package, &archive_name, located.archive, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork number-attachment object {identifier} is missing"
            ))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == NUMBER_ATTACHMENT_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork number-attachment object {identifier} must contain exactly one number payload"
            )));
        };
        if object.messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork number-attachment object {identifier} contains unrelated payloads"
            )));
        }
        let original = &object.messages[*index];
        let current = tswp::NumberAttachmentArchive::decode(original.data.as_slice())?;
        let current_textual = current.super_.as_ref().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork number-attachment object {identifier} is missing its textual payload"
            ))
        })?;
        if current_textual.kind.is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork number-attachment object {identifier} is missing its kind"
            )));
        }
        let mut data = patch_nested_length_delimited_field(
            &original.data,
            &[SUPER_FIELD, STRING_EQUIVALENT_FIELD],
            current_textual.string_equivalent.is_some(),
            settings
                .string_equivalent
                .as_ref()
                .map(|value| value.as_str().as_bytes()),
        )?;
        data = patch_nested_varint_field(
            &data,
            &[SUPER_FIELD, KIND_FIELD],
            true,
            Some(i32_varint(settings.kind.as_raw())),
        )?;
        data = patch_varint_field(
            &data,
            NUMBER_FORMAT_FIELD,
            current.number_format.is_some(),
            settings
                .number_format
                .map(|value| u64::from(value.native_value())),
        )?;
        data = patch_optional_text(
            &data,
            STRING_VALUE_FIELD,
            current.string_value.is_some(),
            settings.string_value.as_ref(),
        )?;
        data = patch_optional_text(
            &data,
            NUMBER_FORMAT_NAME_FIELD,
            current.number_format_name.is_some(),
            settings.number_format_name.as_ref(),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: NUMBER_ATTACHMENT_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn patch_optional_text(
    data: &[u8],
    field: u32,
    present: bool,
    replacement: Option<&TextNumberAttachmentText>,
) -> Result<Vec<u8>> {
    patch_length_delimited_field(
        data,
        field,
        present,
        replacement.map(|value| value.as_str().as_bytes()),
    )
}

const fn i32_varint(value: i32) -> u64 {
    value as i64 as u64
}
