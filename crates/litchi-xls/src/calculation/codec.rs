//! BIFF8 calculation-record field codecs and structural validation.

use crate::error::Result;

use super::invalid;

pub(super) fn parse_bool16(record_type: u16, data: &[u8]) -> Result<bool> {
    require_length(record_type, data, 2)?;
    match read_u16(data, 0) {
        0 => Ok(false),
        1 => Ok(true),
        value => invalid(record_type, format!("Boolean must be 0 or 1, got {value}")),
    }
}

pub(super) fn parse_bool32(record_type: u16, data: &[u8]) -> Result<bool> {
    require_length(record_type, data, 4)?;
    match read_u32(data, 0) {
        0 => Ok(false),
        1 => Ok(true),
        value => invalid(record_type, format!("Boolean must be 0 or 1, got {value}")),
    }
}

pub(super) fn require_future_record_header(
    record_type: u16,
    data: &[u8],
    expected_length: usize,
) -> Result<()> {
    require_length(record_type, data, expected_length)?;
    if read_u16(data, 0) != record_type {
        return invalid(
            record_type,
            "future-record header type does not match containing record",
        );
    }
    if read_u16(data, 2) != 0 || data[4..12].iter().any(|byte| *byte != 0) {
        return invalid(
            record_type,
            "future-record flags and reserved bytes must be zero",
        );
    }
    Ok(())
}

pub(super) fn require_length(record_type: u16, data: &[u8], expected: usize) -> Result<()> {
    if data.len() != expected {
        return invalid(
            record_type,
            format!(
                "payload must be exactly {expected} bytes, got {}",
                data.len()
            ),
        );
    }
    Ok(())
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}
