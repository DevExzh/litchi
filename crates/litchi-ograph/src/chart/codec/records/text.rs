//! BIFF8 and Excel Unicode string codecs used by chart records.

use std::char;

use super::wire::{byte_at, invalid, invalid_model, u16_at, vec_with_capacity};
use crate::{Error, Result};
use litchi_biff::RecordRef;

pub(super) fn parse_short_text(record: RecordRef<'_>) -> Result<String> {
    let data = record.payload();
    if data.len() < 4 || u16_at(data, 0, record)? != 0 {
        return invalid(
            record,
            "SeriesText is truncated or its reserved field is nonzero",
        );
    }
    let string = data.get(2..).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "SeriesText string is truncated",
    })?;
    parse_string(string, record)
}

pub(super) fn short_text(value: &str) -> Result<Vec<u8>> {
    let string = biff_string(value)?;
    let capacity = 2usize
        .checked_add(string.len())
        .ok_or(Error::SizeOverflow {
            resource: "chart text",
        })?;
    let mut data = vec_with_capacity(capacity, "chart text")?;
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&string);
    Ok(data)
}

pub(super) fn parse_string(data: &[u8], record: RecordRef<'_>) -> Result<String> {
    if data.len() < 2 {
        return invalid(record, "chart string is shorter than two bytes");
    }
    let count = usize::from(byte_at(data, 0, record)?);
    let flags = byte_at(data, 1, record)?;
    parse_string_content(data, 2, count, flags, record)
}

pub(super) fn parse_xl_unicode_string(data: &[u8], record: RecordRef<'_>) -> Result<String> {
    if data.len() < 3 {
        return invalid(record, "Excel chart string is shorter than three bytes");
    }
    let count = usize::from(u16_at(data, 0, record)?);
    let flags = byte_at(data, 2, record)?;
    parse_string_content(data, 3, count, flags, record)
}

fn parse_string_content(
    data: &[u8],
    header: usize,
    count: usize,
    flags: u8,
    record: RecordRef<'_>,
) -> Result<String> {
    if flags & !1 != 0 {
        return invalid(record, "chart string uses unsupported option flags");
    }
    let wide = flags & 1 != 0;
    let width = if wide { 2usize } else { 1usize };
    let content = count.checked_mul(width).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    let expected = header.checked_add(content).ok_or(Error::SizeOverflow {
        resource: "chart string",
    })?;
    if data.len() != expected {
        return invalid(
            record,
            "chart string length does not match its character count",
        );
    }
    let bytes = data.get(header..).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "chart string content is truncated",
    })?;
    let reserve = count
        .checked_mul(if wide { 3 } else { 2 })
        .ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?;
    let mut output = String::new();
    output
        .try_reserve_exact(reserve)
        .ok()
        .ok_or(Error::Allocation {
            resource: "chart string",
        })?;
    if wide {
        let units = bytes.chunks_exact(2).map(|value| match value {
            [low, high] => u16::from_le_bytes([*low, *high]),
            _ => 0,
        });
        for value in char::decode_utf16(units) {
            output.push(value.ok().ok_or(Error::InvalidChart {
                offset: record.offset(),
                reason: "chart string contains invalid UTF-16",
            })?);
        }
    } else {
        for value in bytes {
            output.push(char::from(*value));
        }
    }
    Ok(output)
}

pub(super) fn biff_string(value: &str) -> Result<Vec<u8>> {
    let count = value.encode_utf16().count();
    if count > usize::from(u8::MAX) {
        return invalid_model("text", "chart string exceeds 255 UTF-16 code units");
    }
    let wide = value.encode_utf16().any(|unit| unit > u16::from(u8::MAX));
    let width = if wide { 2usize } else { 1usize };
    let capacity = 2usize
        .checked_add(count.checked_mul(width).ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?)
        .ok_or(Error::SizeOverflow {
            resource: "chart string",
        })?;
    let mut data = vec_with_capacity(capacity, "chart string")?;
    data.push(u8::try_from(count).ok().ok_or(Error::InvalidModel {
        field: "text",
        reason: "chart string exceeds 255 UTF-16 code units",
    })?);
    data.push(u8::from(wide));
    for unit in value.encode_utf16() {
        if wide {
            data.extend_from_slice(&unit.to_le_bytes());
        } else {
            data.push(u8::try_from(unit).ok().ok_or(Error::InvalidModel {
                field: "text",
                reason: "narrow chart string contains a wide code unit",
            })?);
        }
    }
    Ok(data)
}

pub(super) fn xl_unicode_string(value: &str) -> Result<Vec<u8>> {
    let count = value.encode_utf16().count();
    let count = u16::try_from(count).ok().ok_or(Error::InvalidModel {
        field: "cached text",
        reason: "Excel chart string exceeds 65,535 UTF-16 code units",
    })?;
    let wide = value.encode_utf16().any(|unit| unit > u16::from(u8::MAX));
    let width = if wide { 2usize } else { 1usize };
    let capacity = 3usize
        .checked_add(
            usize::from(count)
                .checked_mul(width)
                .ok_or(Error::SizeOverflow {
                    resource: "Excel chart string",
                })?,
        )
        .ok_or(Error::SizeOverflow {
            resource: "Excel chart string",
        })?;
    let mut data = vec_with_capacity(capacity, "Excel chart string")?;
    data.extend_from_slice(&count.to_le_bytes());
    data.push(u8::from(wide));
    for unit in value.encode_utf16() {
        if wide {
            data.extend_from_slice(&unit.to_le_bytes());
        } else {
            data.push(u8::try_from(unit).ok().ok_or(Error::InvalidModel {
                field: "cached text",
                reason: "narrow Excel chart string contains a wide code unit",
            })?);
        }
    }
    Ok(data)
}
