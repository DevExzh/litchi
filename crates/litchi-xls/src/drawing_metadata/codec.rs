//! `OfficeArt` `ClientAnchor` wire codec for XLS worksheets.

use super::model::{AnchorBehavior, AnchorPoint, SheetAnchor};
use super::validation::{invalid, invalid_input};
use litchi_odraw::{Record, RecordKind};
use std::io;

const CLIENT_ANCHOR: u16 = 0xF010;
const CLIENT_ANCHOR_LEN: usize = 18;
const CLIENT_ANCHOR_LEN_U32: u32 = 18;

pub(crate) fn decode_sheet_anchor(record: &Record<'_>) -> io::Result<SheetAnchor> {
    if record.kind() != RecordKind::ClientAnchor
        || record.raw_kind() != CLIENT_ANCHOR
        || record.version() != 0
        || record.instance() != 0
        || record.len() != CLIENT_ANCHOR_LEN_U32
    {
        return Err(invalid(
            "XLS ClientAnchor does not have the OfficeArtClientAnchorSheet shape",
        ));
    }
    let data = record.data();
    let behavior = AnchorBehavior::from_wire_flags(u16_at(data, 0)?)?;
    let top_left = AnchorPoint::new(
        u16_at(data, 2)?,
        u16_at(data, 6)?,
        i16_at(data, 4)?,
        i16_at(data, 8)?,
    )?;
    let bottom_right = AnchorPoint::new(
        u16_at(data, 10)?,
        u16_at(data, 14)?,
        i16_at(data, 12)?,
        i16_at(data, 16)?,
    )?;
    SheetAnchor::new(top_left, bottom_right, behavior)
}

impl SheetAnchor {
    /// Encodes this metadata as one complete `OfficeArt` `ClientAnchor` atom.
    ///
    /// The method emits inert metadata only; attaching it to a workbook
    /// drawing remains the responsibility of the XLS drawing writer.
    #[must_use]
    pub fn to_record_bytes(self) -> Vec<u8> {
        let mut output = Vec::with_capacity(8 + CLIENT_ANCHOR_LEN);
        output.extend_from_slice(&0u16.to_le_bytes());
        output.extend_from_slice(&CLIENT_ANCHOR.to_le_bytes());
        output.extend_from_slice(&CLIENT_ANCHOR_LEN_U32.to_le_bytes());
        output.extend_from_slice(&self.behavior().wire_flags().to_le_bytes());
        for point in [self.top_left(), self.bottom_right()] {
            let (column, x, row, y) = point.wire_fields();
            output.extend_from_slice(&column.to_le_bytes());
            output.extend_from_slice(&x.to_le_bytes());
            output.extend_from_slice(&row.to_le_bytes());
            output.extend_from_slice(&y.to_le_bytes());
        }
        output
    }
}

fn u16_at(data: &[u8], offset: usize) -> io::Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid_input("worksheet anchor offset overflows"))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| invalid("worksheet ClientAnchor payload is truncated"))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn i16_at(data: &[u8], offset: usize) -> io::Result<i16> {
    Ok(i16::from_le_bytes(u16_at(data, offset)?.to_le_bytes()))
}
