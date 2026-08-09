//! Structural and semantic validation for the inert OLE-control owner.

use super::model::{Entry, FieldCounts, Metadata, OcxInfo};
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

/// Maximum number of fixed-size records accepted in one metadata table.
pub(crate) const MAX_INFO_COUNT: usize = 1_048_576;
/// Undefined `ODTPersist1` bits retained by the typed model.
pub(crate) const PERSIST1_UNDEFINED: u16 = 0x402D;
/// Bits that MS-DOC requires to be zero in `ODTPersist1`.
const PERSIST1_MUST_BE_ZERO: u16 = 0x0C00;
/// Undefined `ODTPersist2` bits retained by the typed model.
pub(crate) const PERSIST2_UNDEFINED: u16 = 0xFFF0;
/// Bit that MS-DOC requires to be zero in `ODTPersist2`.
const PERSIST2_MUST_BE_ZERO: u16 = 0x0002;
const FIFLD: u16 = 1 << 0;
const ACTIVEX: u16 = 1 << 12;
const STREAM_DATA: u16 = 1 << 13;

/// Validate one `OcxInfo` flags word.
pub(crate) fn flags(raw: u16) -> Result<()> {
    if raw & FIFLD == 0 {
        return Err(corrupted("OcxInfo fifld must be set"));
    }
    Ok(())
}

/// Validate one complete `OcxInfo` record.
pub(crate) fn info(value: &OcxInfo) -> Result<()> {
    flags(value.flags().raw())
}

/// Validate the bounded `RgxOcxInfo` record collection.
pub(crate) fn infos(values: &[OcxInfo]) -> Result<()> {
    if values.len() > MAX_INFO_COUNT {
        return Err(corrupted("RgxOcxInfo count exceeds the metadata limit"));
    }
    let mut cookies = HashSet::with_capacity(values.len());
    for value in values {
        info(value)?;
        if !cookies.insert(value.cookie()) {
            return Err(corrupted("OcxInfo dwCookie values must be unique"));
        }
    }
    Ok(())
}

/// Validate `OcxInfo.ifld` against its story-specific field table.
pub(crate) fn field_indices(values: &[OcxInfo], counts: FieldCounts) -> Result<()> {
    infos(values)?;
    for value in values {
        let count = counts.for_story(value.story());
        if value.field_index() >= count {
            return Err(corrupted(format!(
                "OcxInfo ifld {} is outside the {} field table of length {count}",
                value.field_index(),
                story_name(value.story())
            )));
        }
    }
    Ok(())
}

/// Validate the byte length implied by an `RgxOcxInfo` count.
pub(crate) fn table_size(count: usize, data_len: usize, record_size: usize) -> Result<()> {
    if count > MAX_INFO_COUNT {
        return Err(corrupted("RgxOcxInfo count exceeds the metadata limit"));
    }
    let expected = count
        .checked_mul(record_size)
        .and_then(|size| size.checked_add(4))
        .ok_or_else(|| corrupted("RgxOcxInfo size overflows"))?;
    if expected != data_len {
        return Err(corrupted(format!(
            "RgxOcxInfo requires {expected} bytes for {count} records, got {data_len}"
        )));
    }
    Ok(())
}

/// Validate one raw `ODTPersist1` bitfield.
pub(crate) fn persist1(raw: u16) -> Result<()> {
    if raw & PERSIST1_MUST_BE_ZERO != 0 {
        return Err(corrupted("ODTPersist1 MUST-be-zero bits are set"));
    }
    if raw & STREAM_DATA != 0 && raw & ACTIVEX == 0 {
        return Err(corrupted("ODTPersist1 stream data requires ActiveX"));
    }
    if raw & !((PERSIST1_UNDEFINED) | 0xB3D2) != 0 {
        return Err(corrupted("ODTPersist1 contains unsupported bits"));
    }
    Ok(())
}

/// Validate one raw `ODTPersist2` bitfield.
pub(crate) fn persist2(raw: u16) -> Result<()> {
    if raw & PERSIST2_MUST_BE_ZERO != 0 {
        return Err(corrupted("ODTPersist2 MUST-be-zero bit is set"));
    }
    Ok(())
}

/// Validate one complete `ODT` metadata value.
pub(crate) fn metadata(value: &Metadata) -> Result<()> {
    persist1(value.persist1().raw())?;
    if let Some(value2) = value.persist2() {
        persist2(value2.raw())?;
    }
    Ok(())
}

/// Validate a decimal `ObjectPool` storage name.
pub(crate) fn storage_name(value: &str) -> Result<()> {
    let Some(decimal) = value.strip_prefix('_') else {
        return Err(corrupted("ObjectPool storage name must start with '_'"));
    };
    let digits = decimal.strip_prefix('-').unwrap_or(decimal);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(corrupted(
            "ObjectPool storage name must contain a decimal identifier",
        ));
    }
    Ok(())
}

/// Validate one `ObjectPool` entry and its passive `ActiveX` stream relationship.
pub(crate) fn entry(value: &Entry) -> Result<()> {
    storage_name(value.name().as_str())?;
    if let Some(metadata) = value.metadata() {
        metadata.validate()?;
        if value.control_data_present() && !metadata.is_activex() {
            return Err(corrupted(
                "ObjectPool OCXDATA requires an ActiveX metadata record",
            ));
        }
        if value.control_data_present() != metadata.stores_control_data_in_stream() {
            return Err(corrupted(
                "ObjectPool OCXDATA presence disagrees with ODTPersist1.fStream",
            ));
        }
    } else if value.control_data_present() {
        return Err(corrupted(
            "ObjectPool OCXDATA requires an ObjInfo metadata record",
        ));
    }
    Ok(())
}

/// Validate an ordered `ObjectPool` inventory.
pub(crate) fn pool(values: &[Entry]) -> Result<()> {
    if values.len() > MAX_INFO_COUNT {
        return Err(corrupted(
            "ObjectPool entry count exceeds the metadata limit",
        ));
    }
    let mut names = HashSet::with_capacity(values.len());
    for value in values {
        entry(value)?;
        if !names.insert(value.name().as_str()) {
            return Err(corrupted("ObjectPool storage names must be unique"));
        }
    }
    Ok(())
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn story_name(story: super::model::Story) -> &'static str {
    match story {
        super::model::Story::Main => "main",
        super::model::Story::Header => "header",
        super::model::Story::Footnote => "footnote",
        super::model::Story::Textbox => "textbox",
        super::model::Story::Endnote => "endnote",
        super::model::Story::Comment => "comment",
        super::model::Story::HeaderTextbox => "header-textbox",
    }
}
