//! Native Date & Time smart-field encoding and lossless payload mutation.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tsp, tswp};
use crate::wire::{patch_length_delimited_field, patch_nested_fixed64_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

use super::date_time_types::{
    TextDateTimeFieldSettings, TextDateTimeFormat, TextDateTimeFormatterStyle, TextDateTimeInstant,
    TextDateTimeLocaleIdentifier, TextDateTimeUpdatePlan,
};
use super::smart_field_object::{generated_text_attribute_uuid, validate_text_attribute_uuid};

const FORMAT_FIELD: u32 = 2;
const LOCALE_FIELD: u32 = 3;
const DATE_STYLE_FIELD: u32 = 4;
const TIME_STYLE_FIELD: u32 = 5;
const UPDATE_PLAN_FIELD: u32 = 6;
const NEEDS_UPDATE_FIELD: u32 = 7;
const DATE_FIELD: u32 = 8;
const DATE_SECONDS_FIELD: u32 = 1;
pub(super) const DATE_TIME_MESSAGE_TYPE: u32 = 2_034;

pub(super) fn validate_date_time_object(
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<TextDateTimeFieldSettings>> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == DATE_TIME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if payloads.is_empty() {
        return Ok(None);
    }
    let [message] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork Date & Time object {identifier} contains multiple Date & Time payloads"
        )));
    };
    if object.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork Date & Time object {identifier} contains unrelated payloads"
        )));
    }
    let field = tswp::DateTimeSmartFieldArchive::decode(message.data.as_slice())?;
    let uuid = field
        .super_
        .as_ref()
        .and_then(|smart_field| smart_field.text_attribute_uuid_string.as_deref())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork Date & Time object {identifier} is missing its text-attribute UUID"
            ))
        })?;
    validate_text_attribute_uuid(identifier, "Date & Time", uuid)?;
    Ok(Some(TextDateTimeFieldSettings {
        format: field
            .format
            .map(|value| TextDateTimeFormat::new(value.into_boxed_str()))
            .transpose()?,
        locale_identifier: field
            .locale_identifier
            .map(|value| TextDateTimeLocaleIdentifier::new(value.into_boxed_str()))
            .transpose()?,
        date_style: field.date_style.map(TextDateTimeFormatterStyle::from_raw),
        time_style: field.time_style.map(TextDateTimeFormatterStyle::from_raw),
        update_plan: field.update_plan.map(TextDateTimeUpdatePlan::from_raw),
        needs_update: field.needs_update,
        instant: field
            .date
            .map(|date| TextDateTimeInstant::from_reference_date_seconds(date.seconds))
            .transpose()?,
    }))
}

pub(super) fn new_date_time_object(
    identifier: u64,
    settings: &TextDateTimeFieldSettings,
) -> Result<ArchiveObject> {
    let field = tswp::DateTimeSmartFieldArchive {
        super_: Some(tswp::SmartFieldArchive {
            text_attribute_uuid_string: Some(generated_text_attribute_uuid()?),
        }),
        format: settings
            .format
            .as_ref()
            .map(|value| value.as_str().to_owned()),
        locale_identifier: settings
            .locale_identifier
            .as_ref()
            .map(|value| value.as_str().to_owned()),
        date_style: settings.date_style.map(TextDateTimeFormatterStyle::as_raw),
        time_style: settings.time_style.map(TextDateTimeFormatterStyle::as_raw),
        update_plan: settings.update_plan.map(TextDateTimeUpdatePlan::as_raw),
        needs_update: settings.needs_update,
        date: settings.instant.map(|instant| tsp::Date {
            seconds: instant.reference_date_seconds(),
        }),
    };
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: DATE_TIME_MESSAGE_TYPE,
            data: field.encode_to_vec(),
        }],
    )
}

pub(super) fn patch_date_time_settings(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
    settings: &TextDateTimeFieldSettings,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork Date & Time object {identifier} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == DATE_TIME_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork Date & Time object {identifier} must contain exactly one Date & Time payload"
            )));
        };
        if object.messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork Date & Time object {identifier} contains unrelated payloads"
            )));
        }
        let original = &object.messages[*index];
        let current = tswp::DateTimeSmartFieldArchive::decode(original.data.as_slice())?;
        let mut data = patch_length_delimited_field(
            &original.data,
            FORMAT_FIELD,
            current.format.is_some(),
            settings
                .format
                .as_ref()
                .map(|value| value.as_str().as_bytes()),
        )?;
        data = patch_length_delimited_field(
            &data,
            LOCALE_FIELD,
            current.locale_identifier.is_some(),
            settings
                .locale_identifier
                .as_ref()
                .map(|value| value.as_str().as_bytes()),
        )?;
        for (field_number, present, replacement) in [
            (
                DATE_STYLE_FIELD,
                current.date_style.is_some(),
                settings.date_style.map(TextDateTimeFormatterStyle::as_raw),
            ),
            (
                TIME_STYLE_FIELD,
                current.time_style.is_some(),
                settings.time_style.map(TextDateTimeFormatterStyle::as_raw),
            ),
            (
                UPDATE_PLAN_FIELD,
                current.update_plan.is_some(),
                settings.update_plan.map(TextDateTimeUpdatePlan::as_raw),
            ),
        ] {
            data = patch_varint_field(&data, field_number, present, replacement.map(i32_varint))?;
        }
        data = patch_varint_field(
            &data,
            NEEDS_UPDATE_FIELD,
            current.needs_update.is_some(),
            settings.needs_update.map(u64::from),
        )?;
        data = match (current.date, settings.instant) {
            (Some(_), Some(instant)) => patch_nested_fixed64_field(
                &data,
                &[DATE_FIELD, DATE_SECONDS_FIELD],
                true,
                Some(instant.reference_date_seconds().to_bits()),
            )?,
            (None, Some(instant)) => {
                let date = tsp::Date {
                    seconds: instant.reference_date_seconds(),
                };
                patch_length_delimited_field(&data, DATE_FIELD, false, Some(&date.encode_to_vec()))?
            },
            (Some(_), None) => patch_length_delimited_field(&data, DATE_FIELD, true, None)?,
            (None, None) => data,
        };
        object.replace_message(
            *index,
            RawMessage {
                type_: DATE_TIME_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

const fn i32_varint(value: i32) -> u64 {
    value as i64 as u64
}
