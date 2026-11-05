#![allow(
    clippy::cast_possible_truncation,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation to this codec boundary"
)]

//! BIFF12 `BrtHLink` decoding and encoding.

use crate::raw::{Cursor, Encoder};

use super::model::{Error, Hyperlink, PREFIX_LEN, Result};

/// Parse one `BrtHLink` payload.
pub(super) fn parse(data: &[u8], limits: crate::raw::Limits) -> Result<Hyperlink> {
    limits.validate()?;
    if data.len() > limits.payload() {
        return Err(Error::Wire(crate::raw::Error::PayloadLimit {
            length: data.len(),
            limit: limits.payload(),
            offset: 0,
        }));
    }
    if data.len() < PREFIX_LEN {
        return Err(Error::InvalidLength {
            expected: PREFIX_LEN,
            found: data.len(),
        });
    }

    let mut cursor = Cursor::new(data, "BrtHLink", limits);
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
pub(super) fn serialize(value: &Hyperlink, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    let mut writer = Encoder::new(Vec::new(), limits);
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

fn read_optional_string(cursor: &mut Cursor<'_>) -> Result<Option<String>> {
    if cursor.remaining() == 0 {
        return Ok(None);
    }
    let value = cursor.read_wide_string()?;
    Ok((!value.is_empty()).then_some(value))
}

fn write_optional_string(writer: &mut Encoder<Vec<u8>>, value: Option<&str>) -> Result<()> {
    writer.write_wide_string(value.unwrap_or_default())?;
    Ok(())
}
