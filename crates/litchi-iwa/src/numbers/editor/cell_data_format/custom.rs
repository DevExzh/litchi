//! Document-level registry for custom table-cell formats.

use std::collections::HashSet;

use prost::Message;

use crate::archive::RawMessage;
use crate::numbers::bnc;
use crate::protobuf::{tsk, tsp, tst};
use litchi_numbers::cell::data_format::custom::{
    Condition, ConditionValue, Custom, DateTime as CustomDateTime, DateTimePattern, Name,
    Number as CustomNumber, NumberPattern, NumberRule, Text as CustomText,
};

const NATIVE_CUSTOM_TEXT_VALUE_TOKEN: char = '\u{e421}';
use crate::{Error, IWorkPackage, Result, wire};

const CUSTOM_FORMAT_LIST_MESSAGE_TYPE: u32 = 222;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
const NATIVE_CUSTOM_NUMBER_FORMAT_TYPE: u32 = 270;
const NATIVE_CUSTOM_TEXT_FORMAT_TYPE: u32 = 271;
const NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE: u32 = 272;
const NATIVE_CUSTOM_FRACTION_SENTINEL: u32 = (-3_i32) as u32;
const UUID_FIELD: u32 = 1;
const CUSTOM_FORMAT_FIELD: u32 = 2;

#[derive(Debug)]
struct RegistryLocation {
    archive_name: String,
    object_id: u64,
    message_index: usize,
    message_type: u32,
    data: Vec<u8>,
    list: tsk::CustomFormatListArchive,
}

pub(super) fn acquire_reference(
    package: &mut IWorkPackage,
    format: &Custom,
) -> Result<tsk::FormatStructArchive> {
    let location = locate_registry(package)?;
    validate_registry(&location)?;
    for (uuid, native) in location
        .list
        .uuids
        .iter()
        .zip(&location.list.custom_formats)
    {
        if custom_format_from_native(native).is_ok_and(|existing| existing == *format) {
            return Ok(reference_archive(format_type(format), *uuid));
        }
    }

    let mut existing = location
        .list
        .uuids
        .iter()
        .map(uuid_key)
        .collect::<HashSet<_>>();
    let uuid = fresh_uuid(&mut existing);
    let native = custom_format_to_native(format)?;
    let mut uuid_payloads = wire::repeated_length_delimited_payloads(&location.data, UUID_FIELD)?
        .into_iter()
        .map(<[u8]>::to_vec)
        .collect::<Vec<_>>();
    let mut format_payloads =
        wire::repeated_length_delimited_payloads(&location.data, CUSTOM_FORMAT_FIELD)?
            .into_iter()
            .map(<[u8]>::to_vec)
            .collect::<Vec<_>>();
    uuid_payloads.push(uuid.encode_to_vec());
    format_payloads.push(native.encode_to_vec());
    let data =
        wire::rewrite_repeated_length_delimited_fields(&location.data, UUID_FIELD, &uuid_payloads)?;
    let data = wire::rewrite_repeated_length_delimited_fields(
        &data,
        CUSTOM_FORMAT_FIELD,
        &format_payloads,
    )?;
    replace_registry(package, &location, data)?;
    Ok(reference_archive(format_type(format), uuid))
}

pub(super) fn resolve_reference(
    package: &IWorkPackage,
    reference: &tsk::FormatStructArchive,
) -> Result<Custom> {
    let uuid = reference_uuid(reference)?.ok_or_else(|| {
        Error::InvalidFormat("custom table-cell format has no registry UUID".to_owned())
    })?;
    let location = locate_registry(package)?;
    validate_registry(&location)?;
    let index = location
        .list
        .uuids
        .iter()
        .position(|candidate| candidate == uuid)
        .ok_or_else(|| {
            Error::InvalidFormat(
                "custom table-cell format references an unknown registry UUID".to_owned(),
            )
        })?;
    let format = custom_format_from_native(&location.list.custom_formats[index])?;
    if format_type(&format) != reference.format_type.expect("validated custom format type") {
        return Err(Error::InvalidFormat(
            "custom table-cell format family disagrees with its registry entry".to_owned(),
        ));
    }
    Ok(format)
}

pub(super) fn release_reference_if_unused(
    package: &mut IWorkPackage,
    reference: &tsk::FormatStructArchive,
) -> Result<()> {
    let Some(uuid) = reference_uuid(reference)? else {
        return Ok(());
    };
    if package_references_uuid(package, uuid)? {
        return Ok(());
    }

    let location = locate_registry(package)?;
    validate_registry(&location)?;
    let index = location
        .list
        .uuids
        .iter()
        .position(|candidate| candidate == uuid)
        .ok_or_else(|| {
            Error::InvalidFormat(
                "released custom table-cell format is missing from the registry".to_owned(),
            )
        })?;
    let mut uuid_payloads = wire::repeated_length_delimited_payloads(&location.data, UUID_FIELD)?
        .into_iter()
        .map(<[u8]>::to_vec)
        .collect::<Vec<_>>();
    let mut format_payloads =
        wire::repeated_length_delimited_payloads(&location.data, CUSTOM_FORMAT_FIELD)?
            .into_iter()
            .map(<[u8]>::to_vec)
            .collect::<Vec<_>>();
    uuid_payloads.remove(index);
    format_payloads.remove(index);
    let data =
        wire::rewrite_repeated_length_delimited_fields(&location.data, UUID_FIELD, &uuid_payloads)?;
    let data = wire::rewrite_repeated_length_delimited_fields(
        &data,
        CUSTOM_FORMAT_FIELD,
        &format_payloads,
    )?;
    replace_registry(package, &location, data)
}

pub(super) fn reference_uuid(native: &tsk::FormatStructArchive) -> Result<Option<&tsp::Uuid>> {
    if !matches!(
        native.format_type,
        Some(
            NATIVE_CUSTOM_NUMBER_FORMAT_TYPE
                | NATIVE_CUSTOM_TEXT_FORMAT_TYPE
                | NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE
        )
    ) {
        return Ok(None);
    }
    let uuid = native.custom_uid.as_ref().ok_or_else(|| {
        Error::InvalidFormat("custom table-cell format has no registry UUID".to_owned())
    })?;
    let expected = reference_archive(
        native
            .format_type
            .expect("matched custom native format type"),
        *uuid,
    );
    if native != &expected {
        return Err(Error::InvalidFormat(
            "custom table-cell format reference contains inconsistent metadata".to_owned(),
        ));
    }
    Ok(Some(uuid))
}

pub(super) const fn scalar_kind(format: &Custom) -> bnc::CellDataFormatKind {
    match format {
        Custom::Number(_) => bnc::CellDataFormatKind::NumberOrPercentage,
        Custom::Text(_) => bnc::CellDataFormatKind::Text,
        Custom::DateTime(_) => bnc::CellDataFormatKind::DateTime,
    }
}

fn custom_format_to_native(format: &Custom) -> Result<tsk::CustomFormatArchive> {
    Ok(match format {
        Custom::Number(format) => tsk::CustomFormatArchive {
            name: format.name().as_str().to_owned(),
            format_type_pre_bnc: NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
            default_format: Box::new(number_pattern_to_native(format.default_pattern())),
            conditions: format.rules().iter().map(number_rule_to_native).collect(),
            format_type: Some(NATIVE_CUSTOM_NUMBER_FORMAT_TYPE),
        },
        Custom::Text(format) => tsk::CustomFormatArchive {
            name: format.name().as_str().to_owned(),
            format_type_pre_bnc: NATIVE_CUSTOM_TEXT_FORMAT_TYPE,
            default_format: Box::new(text_pattern_to_native(format)?),
            conditions: Vec::new(),
            format_type: Some(NATIVE_CUSTOM_TEXT_FORMAT_TYPE),
        },
        Custom::DateTime(format) => tsk::CustomFormatArchive {
            name: format.name().as_str().to_owned(),
            format_type_pre_bnc: NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE,
            default_format: Box::new(date_time_pattern_to_native(format.pattern())),
            conditions: Vec::new(),
            format_type: Some(NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE),
        },
    })
}

fn custom_format_from_native(native: &tsk::CustomFormatArchive) -> Result<Custom> {
    let name = Name::try_new(native.name.clone())?;
    let format_type = native.format_type.ok_or_else(|| {
        Error::InvalidFormat("custom table-cell registry entry has no format family".to_owned())
    })?;
    if native.format_type_pre_bnc != format_type {
        return Err(Error::InvalidFormat(
            "custom table-cell registry entry has inconsistent format families".to_owned(),
        ));
    }
    match format_type {
        NATIVE_CUSTOM_NUMBER_FORMAT_TYPE => {
            let default_pattern = number_pattern_from_native(&native.default_format)?;
            let rules = native
                .conditions
                .iter()
                .map(number_rule_from_native)
                .collect::<Result<Vec<_>>>()?;
            CustomNumber::try_with_rules(name, default_pattern, rules)
                .map(Custom::Number)
                .map_err(Into::into)
        },
        NATIVE_CUSTOM_TEXT_FORMAT_TYPE => {
            if !native.conditions.is_empty() {
                return Err(Error::InvalidFormat(
                    "custom Text format cannot contain numeric conditions".to_owned(),
                ));
            }
            text_pattern_from_native(name, &native.default_format).map(Custom::Text)
        },
        NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE => {
            if !native.conditions.is_empty() {
                return Err(Error::InvalidFormat(
                    "custom Date & Time format cannot contain numeric conditions".to_owned(),
                ));
            }
            let pattern = date_time_pattern_from_native(&native.default_format)?;
            Ok(Custom::DateTime(CustomDateTime::new(name, pattern)))
        },
        _ => Err(Error::InvalidFormat(format!(
            "unsupported custom table-cell format family {format_type}"
        ))),
    }
}

fn number_rule_to_native(rule: &NumberRule) -> tsk::custom_format_archive::Condition {
    tsk::custom_format_archive::Condition {
        condition_type: match rule.condition() {
            Condition::EqualTo(_) => 0,
            Condition::LessThan(_) => 1,
            Condition::LessThanOrEqualTo(_) => 2,
            Condition::GreaterThan(_) => 3,
            Condition::GreaterThanOrEqualTo(_) => 4,
        },
        condition_value: None,
        condition_format: number_pattern_to_native(rule.pattern()),
        condition_value_dbl: Some(rule.condition().threshold().value()),
    }
}

fn number_rule_from_native(native: &tsk::custom_format_archive::Condition) -> Result<NumberRule> {
    let threshold = match (native.condition_value, native.condition_value_dbl) {
        (None, Some(value)) => value,
        (Some(value), None) => f64::from(value),
        _ => {
            return Err(Error::InvalidFormat(
                "custom Number rule must contain exactly one threshold".to_owned(),
            ));
        },
    };
    let threshold = ConditionValue::try_new(threshold)?;
    let condition = match native.condition_type {
        0 => Condition::EqualTo(threshold),
        1 => Condition::LessThan(threshold),
        2 => Condition::LessThanOrEqualTo(threshold),
        3 => Condition::GreaterThan(threshold),
        4 => Condition::GreaterThanOrEqualTo(threshold),
        value => {
            return Err(Error::InvalidFormat(format!(
                "unsupported custom Number condition type {value}"
            )));
        },
    };
    Ok(NumberRule::new(
        condition,
        number_pattern_from_native(&native.condition_format)?,
    ))
}

fn number_pattern_to_native(pattern: &NumberPattern) -> tsk::FormatStructArchive {
    custom_pattern_archive(
        NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        pattern.as_str().to_owned(),
        pattern.as_str().contains(','),
        true,
        pattern_suffix_width(pattern.as_str()),
    )
}

fn number_pattern_from_native(native: &tsk::FormatStructArchive) -> Result<NumberPattern> {
    validate_custom_pattern_archive(native, NATIVE_CUSTOM_NUMBER_FORMAT_TYPE)?;
    NumberPattern::try_new(
        native
            .custom_format_string
            .as_ref()
            .expect("validated custom format string")
            .clone(),
    )
    .map_err(Into::into)
}

fn date_time_pattern_to_native(pattern: &DateTimePattern) -> tsk::FormatStructArchive {
    custom_pattern_archive(
        NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE,
        pattern.as_str().to_owned(),
        false,
        true,
        0,
    )
}

fn date_time_pattern_from_native(native: &tsk::FormatStructArchive) -> Result<DateTimePattern> {
    validate_custom_pattern_archive(native, NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE)?;
    DateTimePattern::try_new(
        native
            .custom_format_string
            .as_ref()
            .expect("validated custom format string")
            .clone(),
    )
    .map_err(Into::into)
}

fn text_pattern_to_native(format: &CustomText) -> Result<tsk::FormatStructArchive> {
    if format.prefix().contains(NATIVE_CUSTOM_TEXT_VALUE_TOKEN)
        || format.suffix().contains(NATIVE_CUSTOM_TEXT_VALUE_TOKEN)
    {
        return Err(Error::InvalidFormat(
            "custom Text affixes cannot contain the native cell-value token".to_owned(),
        ));
    }
    Ok(custom_pattern_archive(
        NATIVE_CUSTOM_TEXT_FORMAT_TYPE,
        native_text_pattern(format),
        false,
        false,
        u32::try_from(format.suffix().encode_utf16().count()).unwrap_or(u32::MAX),
    ))
}

fn native_text_pattern(format: &CustomText) -> String {
    let mut pattern = String::with_capacity(
        format.prefix().len()
            + format.suffix().len()
            + usize::from(format.includes_cell_text()) * NATIVE_CUSTOM_TEXT_VALUE_TOKEN.len_utf8(),
    );
    pattern.push_str(format.prefix());
    if format.includes_cell_text() {
        pattern.push(NATIVE_CUSTOM_TEXT_VALUE_TOKEN);
    }
    pattern.push_str(format.suffix());
    pattern
}

fn text_pattern_from_native(name: Name, native: &tsk::FormatStructArchive) -> Result<CustomText> {
    validate_custom_pattern_archive(native, NATIVE_CUSTOM_TEXT_FORMAT_TYPE)?;
    let pattern = native
        .custom_format_string
        .as_ref()
        .expect("validated custom format string");
    let mut tokens = pattern.match_indices(NATIVE_CUSTOM_TEXT_VALUE_TOKEN);
    let first = tokens.next();
    if tokens.next().is_some() {
        return Err(Error::InvalidFormat(
            "custom Text format contains more than one cell-value token".to_owned(),
        ));
    }
    match first {
        Some((offset, _)) => {
            let suffix_offset = offset + NATIVE_CUSTOM_TEXT_VALUE_TOKEN.len_utf8();
            CustomText::try_new(
                name,
                pattern[..offset].to_owned(),
                pattern[suffix_offset..].to_owned(),
            )
            .map_err(Into::into)
        },
        None => CustomText::try_literal(name, pattern.clone()).map_err(Into::into),
    }
}

fn custom_pattern_archive(
    format_type: u32,
    pattern: String,
    thousands_separator: bool,
    contains_integer_token: bool,
    suffix_width: u32,
) -> tsk::FormatStructArchive {
    tsk::FormatStructArchive {
        format_type: Some(format_type),
        show_thousands_separator: Some(thousands_separator),
        use_accounting_style: Some(false),
        fraction_accuracy: Some(NATIVE_CUSTOM_FRACTION_SENTINEL),
        custom_format_string: Some(pattern),
        scale_factor: Some(1.0),
        requires_fraction_replacement: Some(false),
        decimal_width: Some(0),
        min_integer_width: Some(0),
        num_nonspace_integer_digits: Some(0),
        num_nonspace_decimal_digits: Some(0),
        index_from_right_last_integer: Some(suffix_width),
        num_hash_decimal_digits: Some(0),
        total_num_decimal_digits: Some(0),
        is_complex: Some(false),
        contains_integer_token: Some(contains_integer_token),
        ..Default::default()
    }
}

fn validate_custom_pattern_archive(
    native: &tsk::FormatStructArchive,
    expected_type: u32,
) -> Result<()> {
    let pattern = native.custom_format_string.as_ref().ok_or_else(|| {
        Error::InvalidFormat("custom table-cell format contains no pattern".to_owned())
    })?;
    if native.format_type != Some(expected_type)
        || native.show_thousands_separator.is_none()
        || native.use_accounting_style != Some(false)
        || native.fraction_accuracy != Some(NATIVE_CUSTOM_FRACTION_SENTINEL)
        || native.scale_factor != Some(1.0)
        || native.requires_fraction_replacement != Some(false)
        || native.decimal_width.is_none()
        || native.min_integer_width.is_none()
        || native.num_nonspace_integer_digits.is_none()
        || native.num_nonspace_decimal_digits.is_none()
        || native.index_from_right_last_integer.is_none()
        || native.num_hash_decimal_digits.is_none()
        || native.total_num_decimal_digits.is_none()
        || native.is_complex != Some(false)
        || native.contains_integer_token.is_none()
        || !native.interstitial_strings.is_empty()
        || native.inters_str_insertion_indexes.is_some()
        || native.decimal_places.is_some()
        || native.currency_code.is_some()
        || native.negative_style.is_some()
        || native.duration_style.is_some()
        || native.base.is_some()
        || native.base_places.is_some()
        || native.base_use_minus_sign.is_some()
        || native.suppress_date_format.is_some()
        || native.suppress_time_format.is_some()
        || native.date_time_format.is_some()
        || native.duration_unit_largest.is_some()
        || native.duration_unit_smallest.is_some()
        || native.custom_id.is_some()
        || native.control_minimum.is_some()
        || native.control_maximum.is_some()
        || native.control_increment.is_some()
        || native.control_format_type.is_some()
        || native.slider_orientation.is_some()
        || native.slider_position.is_some()
        || native.multiple_choice_list_initial_value.is_some()
        || native.multiple_choice_list_id.is_some()
        || native.use_automatic_duration_units.is_some()
        || native.custom_uid.is_some()
        || native.custom_format.is_some()
        || native.uses_plus_sign.is_some()
        || native.bool_true_string.is_some()
        || native.bool_false_string.is_some()
    {
        return Err(Error::InvalidFormat(
            "custom table-cell pattern contains unsupported native metadata".to_owned(),
        ));
    }
    if pattern.is_empty() {
        return Err(Error::InvalidFormat(
            "custom table-cell format pattern cannot be empty".to_owned(),
        ));
    }
    Ok(())
}

fn pattern_suffix_width(pattern: &str) -> u32 {
    let suffix = pattern
        .char_indices()
        .rev()
        .find(|(_, character)| matches!(character, '#' | '0'))
        .map_or(pattern, |(offset, character)| {
            &pattern[offset + character.len_utf8()..]
        });
    u32::try_from(suffix.encode_utf16().count()).unwrap_or(u32::MAX)
}

fn format_type(format: &Custom) -> u32 {
    match format {
        Custom::Number(_) => NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        Custom::Text(_) => NATIVE_CUSTOM_TEXT_FORMAT_TYPE,
        Custom::DateTime(_) => NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE,
    }
}

fn reference_archive(format_type: u32, uuid: tsp::Uuid) -> tsk::FormatStructArchive {
    tsk::FormatStructArchive {
        format_type: Some(format_type),
        custom_uid: Some(uuid),
        ..Default::default()
    }
}

fn locate_registry(package: &IWorkPackage) -> Result<RegistryLocation> {
    let mut result = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in archive.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("custom-format registry object has no identifier".to_owned())
            })?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != CUSTOM_FORMAT_LIST_MESSAGE_TYPE {
                    continue;
                }
                let location = RegistryLocation {
                    archive_name: archive_name.to_owned(),
                    object_id,
                    message_index,
                    message_type: message.type_,
                    data: message.data.clone(),
                    list: tsk::CustomFormatListArchive::decode(message.data.as_slice())?,
                };
                if result.replace(location).is_some() {
                    return Err(Error::InvalidFormat(
                        "iWork package contains multiple custom-format registries".to_owned(),
                    ));
                }
            }
        }
    }
    result.ok_or_else(|| {
        Error::InvalidFormat("iWork package has no custom-format registry".to_owned())
    })
}

#[cfg(test)]
pub(super) fn registry_entry_count(package: &IWorkPackage) -> Result<usize> {
    let location = locate_registry(package)?;
    validate_registry(&location)?;
    Ok(location.list.uuids.len())
}

fn validate_registry(location: &RegistryLocation) -> Result<()> {
    if location.list.uuids.len() != location.list.custom_formats.len() {
        return Err(Error::InvalidFormat(
            "custom-format registry UUID and payload counts differ".to_owned(),
        ));
    }
    let raw_uuids = wire::repeated_length_delimited_payloads(&location.data, UUID_FIELD)?;
    let raw_formats =
        wire::repeated_length_delimited_payloads(&location.data, CUSTOM_FORMAT_FIELD)?;
    if raw_uuids.len() != location.list.uuids.len()
        || raw_formats.len() != location.list.custom_formats.len()
    {
        return Err(Error::InvalidFormat(
            "custom-format registry wire and decoded counts differ".to_owned(),
        ));
    }
    let mut seen = HashSet::new();
    for (index, uuid) in location.list.uuids.iter().enumerate() {
        if !seen.insert(uuid_key(uuid)) {
            return Err(Error::InvalidFormat(
                "custom-format registry contains a duplicate UUID".to_owned(),
            ));
        }
        if tsp::Uuid::decode(raw_uuids[index])? != *uuid
            || tsk::CustomFormatArchive::decode(raw_formats[index])?
                != location.list.custom_formats[index]
        {
            return Err(Error::InvalidFormat(
                "custom-format registry wire payload disagrees with its decoded value".to_owned(),
            ));
        }
    }
    Ok(())
}

fn replace_registry(
    package: &mut IWorkPackage,
    location: &RegistryLocation,
    data: Vec<u8>,
) -> Result<()> {
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat("custom-format registry object disappeared".to_owned())
        })?;
        let message = object.messages.get(location.message_index).ok_or_else(|| {
            Error::InvalidFormat("custom-format registry message disappeared".to_owned())
        })?;
        if message.type_ != location.message_type || message.data != location.data {
            return Err(Error::InvalidFormat(
                "custom-format registry changed during mutation".to_owned(),
            ));
        }
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: location.message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn package_references_uuid(package: &IWorkPackage, uuid: &tsp::Uuid) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in object.messages {
                let entries = match message.type_ {
                    TABLE_DATA_LIST_MESSAGE_TYPE => {
                        let list = tst::TableDataList::decode(message.data.as_slice())?;
                        if list.list_type() != tst::table_data_list::ListType::Format {
                            continue;
                        }
                        list.entries
                    },
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE => {
                        let segment = tst::TableDataListSegment::decode(message.data.as_slice())?;
                        if segment.list_type() != tst::table_data_list::ListType::Format {
                            continue;
                        }
                        segment.entries
                    },
                    _ => continue,
                };
                for entry in entries {
                    if entry.refcount == 0 {
                        return Err(Error::InvalidFormat(
                            "format table contains a zero-reference entry".to_owned(),
                        ));
                    }
                    if entry
                        .format
                        .as_ref()
                        .and_then(|format| format.custom_uid.as_ref())
                        == Some(uuid)
                    {
                        return Ok(true);
                    }
                }
            }
        }
    }
    Ok(false)
}

fn fresh_uuid(existing: &mut HashSet<(u64, u64)>) -> tsp::Uuid {
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0; 8];
        lower.copy_from_slice(&bytes[..8]);
        let mut upper = [0; 8];
        upper.copy_from_slice(&bytes[8..]);
        let uuid = tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        };
        if uuid.lower != 0 && uuid.upper != 0 && existing.insert(uuid_key(&uuid)) {
            return uuid;
        }
    }
}

const fn uuid_key(uuid: &tsp::Uuid) -> (u64, u64) {
    (uuid.lower, uuid.upper)
}
