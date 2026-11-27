//! Shared binary-header and primitive readers for animation parsing.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) fn require_container(
    record: &Record,
    record_type: RecordType,
    instance: u16,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(Error::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0x0F, instance, None, name)?;
    let encoded_children_length = record.children.iter().try_fold(0usize, |length, child| {
        length.checked_add(8 + child.data.len())
    });
    if encoded_children_length != Some(record.data.len()) {
        return Err(Error::Corrupted(format!(
            "{name} child records do not cover its complete payload"
        )));
    }
    Ok(())
}

pub(super) fn require_atom(
    record: &Record,
    record_type: RecordType,
    version: u16,
    length: usize,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(Error::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, version, 0, Some(length), name)
}

pub(super) fn require_header(
    record: &Record,
    version: u16,
    instance: u16,
    length: Option<usize>,
    name: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.data_length as usize != record.data.len()
        || length.is_some_and(|expected_length| record.data.len() != expected_length)
    {
        return Err(Error::Corrupted(format!(
            "invalid {name} header: version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }
    Ok(())
}

pub(super) fn parse_bool1(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(Error::InvalidFormat(format!(
            "{field} has invalid bool1 value {value}"
        ))),
    }
}

pub(super) fn parse_optional_time_value<T>(
    is_set: bool,
    value: u32,
    parse: impl FnOnce(u32) -> Option<T>,
    field: &str,
) -> Result<Option<T>> {
    if is_set {
        parse(value)
            .map(Some)
            .ok_or_else(|| Error::InvalidFormat(format!("{field} has invalid value {value}")))
    } else if value == 0 {
        Ok(None)
    } else {
        Err(Error::InvalidFormat(format!(
            "{field} must be zero when not explicitly set"
        )))
    }
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

pub(super) fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([data[offset], data[offset + 1]])
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

pub(super) fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

pub(super) fn read_f32(data: &[u8], offset: usize) -> f32 {
    f32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}
