//! Typed payload models shared by the animation behavior parsers.

use super::super::support::{parse_bool1, read_f32, read_i32, read_u32};
use crate::animation::types::{
    TimeAnimateColor, TimeAnimateColorBy, TimeVariantValue, is_valid_time_formula,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) fn parse_generic_time_variant(record: &Record) -> Result<TimeVariantValue> {
    match record.data.first() {
        Some(0) => parse_time_variant_bool(record).map(TimeVariantValue::Boolean),
        Some(1) => parse_time_variant_i32(record).map(TimeVariantValue::Integer),
        Some(2) => parse_time_variant_f32(record).map(TimeVariantValue::Float),
        Some(3) => parse_time_variant_string(record).map(TimeVariantValue::String),
        _ => Err(Error::InvalidFormat(
            "invalid animation keyframe value type".to_string(),
        )),
    }
}

pub(super) fn parse_animate_color_by(data: &[u8]) -> Result<TimeAnimateColorBy> {
    match read_u32(data, 0) {
        0 | 1 => {
            let values = (read_i32(data, 4), read_i32(data, 8), read_i32(data, 12));
            if [values.0, values.1, values.2]
                .iter()
                .any(|value| !(-255..=255).contains(value))
            {
                return Err(Error::InvalidFormat(
                    "color offset component is out of range".to_string(),
                ));
            }
            if read_u32(data, 0) == 0 {
                Ok(TimeAnimateColorBy::Rgb {
                    red: values.0,
                    green: values.1,
                    blue: values.2,
                })
            } else {
                Ok(TimeAnimateColorBy::Hsl {
                    hue: values.0,
                    saturation: values.1,
                    luminance: values.2,
                })
            }
        },
        2 => parse_scheme_color(data).map(TimeAnimateColorBy::Scheme),
        model => Err(Error::InvalidFormat(format!(
            "invalid color-by model {model}"
        ))),
    }
}

pub(super) fn parse_animate_color(data: &[u8]) -> Result<TimeAnimateColor> {
    match read_u32(data, 0) {
        0 => {
            let (red, green, blue) = (read_u32(data, 4), read_u32(data, 8), read_u32(data, 12));
            if red > 255 || green > 255 || blue > 255 {
                return Err(Error::InvalidFormat(
                    "RGB color component is out of range".to_string(),
                ));
            }
            Ok(TimeAnimateColor::Rgb { red, green, blue })
        },
        2 => parse_scheme_color(data).map(TimeAnimateColor::Scheme),
        model => Err(Error::InvalidFormat(format!(
            "invalid absolute color model {model}"
        ))),
    }
}

fn parse_scheme_color(data: &[u8]) -> Result<u32> {
    let index = read_u32(data, 4);
    if index > 7 {
        return Err(Error::InvalidFormat(
            "scheme color index is out of range".to_string(),
        ));
    }
    Ok(index)
}

pub(super) fn parse_time_string_list(record: &Record) -> Result<Vec<String>> {
    super::super::support::require_container(
        record,
        RecordType::TimeVariantList,
        1,
        "TimeStringListContainer",
    )?;
    record
        .children
        .iter()
        .map(|child| {
            if child.record_type != RecordType::TimeVariant || child.version != 0 {
                return Err(Error::InvalidFormat(
                    "invalid TimeStringListContainer child".to_string(),
                ));
            }
            parse_time_variant_string(child)
        })
        .collect()
}

pub(super) fn parse_time_variant_i32(record: &Record) -> Result<i32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 1 {
        return Err(Error::InvalidFormat(
            "invalid integer time variant".to_string(),
        ));
    }
    Ok(read_i32(&record.data, 1))
}

pub(super) fn parse_time_variant_f32(record: &Record) -> Result<f32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 2 {
        return Err(Error::InvalidFormat(
            "invalid floating-point time variant".to_string(),
        ));
    }
    Ok(read_f32(&record.data, 1))
}

pub(super) fn parse_time_variant_bool(record: &Record) -> Result<bool> {
    require_time_variant_payload(record)?;
    if record.data.len() != 2 || record.data[0] != 0 {
        return Err(Error::InvalidFormat(
            "invalid Boolean time variant".to_string(),
        ));
    }
    parse_bool1(record.data[1], "TimeVariant.boolValue")
}

pub(crate) fn parse_time_variant_string(record: &Record) -> Result<String> {
    require_time_variant_payload(record)?;
    if record.data.len() % 2 != 1 || record.data.first() != Some(&3) {
        return Err(Error::InvalidFormat(
            "invalid string time variant".to_string(),
        ));
    }
    String::from_utf16(
        &record.data[1..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect::<Vec<_>>(),
    )
    .map_err(|_err| Error::InvalidFormat("invalid UTF-16 time variant".to_string()))
}

pub(crate) fn require_time_variant_payload(record: &Record) -> Result<()> {
    if record.data_length as usize != record.data.len() {
        return Err(Error::Corrupted(
            "truncated TimeVariant payload".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_time_formula(formula: &str) -> Result<()> {
    if !is_valid_time_formula(formula) {
        return Err(Error::InvalidFormat(
            "invalid animation keyframe formula".to_string(),
        ));
    }
    Ok(())
}
