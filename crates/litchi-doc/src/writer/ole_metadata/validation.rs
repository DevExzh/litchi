//! Structural validation for the fixed DOC OLE metadata profile.

use super::codec::{
    COMP_OBJ_LEN, COMPOBJ_RESERVED, COMPOBJ_RESERVED_MARKER, COMPOBJ_VERSION, OLE_LEN,
    UNICODE_MARKER,
};
use super::model::{CompObj, Ole};
use std::mem::size_of;

/// Validate a generated `\x01CompObj` stream against its semantic model.
pub(super) fn comp_obj(data: &[u8], expected: CompObj) -> Result<(), &'static str> {
    if data.len() != COMP_OBJ_LEN {
        return Err("CompObj stream has an unexpected length");
    }

    let mut cursor = 0;
    require_bytes(data, &mut cursor, &COMPOBJ_VERSION, "CompObj version")?;
    require_bytes(
        data,
        &mut cursor,
        &COMPOBJ_RESERVED,
        "CompObj reserved field",
    )?;
    require_bytes(
        data,
        &mut cursor,
        &COMPOBJ_RESERVED_MARKER,
        "CompObj reserved marker",
    )?;
    require_bytes(
        data,
        &mut cursor,
        expected.class_id().as_bytes(),
        "CompObj class identifier",
    )?;
    require_string(data, &mut cursor, expected.user_type(), "CompObj user type")?;
    require_string(
        data,
        &mut cursor,
        expected.clipboard_format(),
        "CompObj clipboard format",
    )?;
    require_string(data, &mut cursor, expected.prog_id(), "CompObj ProgID")?;
    require_u32(data, &mut cursor, UNICODE_MARKER, "CompObj Unicode marker")?;
    for label in [
        "CompObj Unicode user type",
        "CompObj Unicode clipboard format",
        "CompObj Unicode ProgID",
    ] {
        require_u32(data, &mut cursor, 0, label)?;
    }
    if cursor == data.len() {
        Ok(())
    } else {
        Err("CompObj stream has trailing bytes")
    }
}

/// Validate a generated `\x01Ole` stream against its semantic model.
pub(super) fn ole(data: &[u8], expected: Ole) -> Result<(), &'static str> {
    if data.len() != OLE_LEN {
        return Err("Ole stream has an unexpected length");
    }
    let version = data
        .get(..size_of::<u32>())
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or("Ole stream is missing its version")?;
    if version != expected.version() {
        return Err("Ole stream has an unexpected version");
    }
    if data[size_of::<u32>()..].iter().any(|&byte| byte != 0) {
        return Err("Ole stream reserved bytes are not zero");
    }
    Ok(())
}

fn require_bytes(
    data: &[u8],
    cursor: &mut usize,
    expected: &[u8],
    label: &'static str,
) -> Result<(), &'static str> {
    let end = cursor
        .checked_add(expected.len())
        .ok_or("OLE metadata cursor overflow")?;
    if data.get(*cursor..end) != Some(expected) {
        return Err(label);
    }
    *cursor = end;
    Ok(())
}

fn require_u32(
    data: &[u8],
    cursor: &mut usize,
    expected: u32,
    label: &'static str,
) -> Result<(), &'static str> {
    let end = cursor
        .checked_add(size_of::<u32>())
        .ok_or("OLE metadata cursor overflow")?;
    let value = data
        .get(*cursor..end)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or(label)?;
    if value != expected {
        return Err(label);
    }
    *cursor = end;
    Ok(())
}

fn require_string(
    data: &[u8],
    cursor: &mut usize,
    expected: &str,
    label: &'static str,
) -> Result<(), &'static str> {
    let length_end = (*cursor).checked_add(size_of::<u32>()).ok_or(label)?;
    let length = usize::try_from(
        data.get(*cursor..length_end)
            .and_then(|bytes| bytes.try_into().ok())
            .map(u32::from_le_bytes)
            .ok_or(label)?,
    )
    .map_err(|_| label)?;
    *cursor = length_end;
    let end = (*cursor).checked_add(length).ok_or(label)?;
    let bytes = data.get(*cursor..end).ok_or(label)?;
    if bytes.len() != expected.len() + 1
        || bytes.get(..expected.len()) != Some(expected.as_bytes())
        || bytes.last() != Some(&0)
    {
        return Err(label);
    }
    *cursor = end;
    Ok(())
}
