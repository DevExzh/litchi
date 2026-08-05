//! BIFF12 `BrtHLink` decoding and encoding.

use crate::raw::{Cursor, Writer};

use super::model::{Error, Hyperlink, PREFIX_LEN, Result};

/// Parse one `BrtHLink` payload.
pub(super) fn parse(data: &[u8]) -> Result<Hyperlink> {
    if data.len() < PREFIX_LEN {
        return Err(Error::InvalidLength {
            expected: PREFIX_LEN,
            found: data.len(),
        });
    }

    let mut cursor = Cursor::new(data, "BrtHLink");
    let row_first = cursor.read_u32()?;
    let row_last = cursor.read_u32()?;
    let col_first = cursor.read_u32()?;
    let col_last = cursor.read_u32()?;
    let r_id = cursor.read_wide_string()?;
    let location = read_optional_string(&mut cursor)?;
    let tooltip = read_optional_string(&mut cursor)?;
    let display = read_optional_string(&mut cursor)?;
    cursor.finish()?;

    Ok(Hyperlink {
        row_first,
        row_last,
        col_first,
        col_last,
        r_id,
        location,
        tooltip,
        display,
        target: None,
    })
}

/// Serialize one `BrtHLink` payload with checked raw-kernel limits.
pub(super) fn serialize(value: &Hyperlink) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_u32(value.row_first)?;
    writer.write_u32(value.row_last)?;
    writer.write_u32(value.col_first)?;
    writer.write_u32(value.col_last)?;
    writer.write_wide_string(&value.r_id)?;
    write_optional_string(&mut writer, value.location.as_deref())?;
    write_optional_string(&mut writer, value.tooltip.as_deref())?;
    write_optional_string(&mut writer, value.display.as_deref())?;
    Ok(writer.finish())
}

/// Serialize using the legacy infallible API.
pub(super) fn serialize_legacy(value: &Hyperlink) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&value.row_first.to_le_bytes());
    data.extend_from_slice(&value.row_last.to_le_bytes());
    data.extend_from_slice(&value.col_first.to_le_bytes());
    data.extend_from_slice(&value.col_last.to_le_bytes());
    write_wide_string_legacy(&mut data, &value.r_id);
    write_wide_string_legacy(&mut data, value.location.as_deref().unwrap_or_default());
    write_wide_string_legacy(&mut data, value.tooltip.as_deref().unwrap_or_default());
    write_wide_string_legacy(&mut data, value.display.as_deref().unwrap_or_default());
    data
}

fn read_optional_string(cursor: &mut Cursor<'_>) -> Result<Option<String>> {
    if cursor.remaining() == 0 {
        return Ok(None);
    }
    let value = cursor.read_wide_string()?;
    Ok((!value.is_empty()).then_some(value))
}

fn write_optional_string(writer: &mut Writer<Vec<u8>>, value: Option<&str>) -> Result<()> {
    writer.write_wide_string(value.unwrap_or_default())?;
    Ok(())
}

fn write_wide_string_legacy(data: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().count();
    data.extend_from_slice(&(units as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}
