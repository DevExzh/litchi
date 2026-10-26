//! The table extents of a table formula owner across a table-axis edit.
//!
//! The owner's spanning dependencies carry the table's total and body ranges
//! (`total_range_for_table`, `body_range_for_table`); they follow the new
//! geometry. Every other coordinate of the calculation graph moves with the
//! owner's formula-coordinate shift (record 0318), which the caller runs.

use super::wire::{bytes_field, fields, rebuild, single_varint, varint_field};
use super::{Error, Result};

/// The column sentinel of an unset `TSCE.RangeCoordinateArchive`.
const SENTINEL_COLUMN: u64 = 0x7fff;
/// The row sentinel of an unset `TSCE.RangeCoordinateArchive`.
const SENTINEL_ROW: u64 = 0x7fff_ffff;

/// One table's table-sized extents: `[first column, first row, last
/// column, last row]` of the whole table and of its body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Extents {
    pub total: [u64; 4],
    pub body: [u64; 4],
}

impl Extents {
    pub fn of(
        rows: u32,
        columns: u32,
        header_rows: u32,
        header_columns: u32,
        footer_rows: u32,
    ) -> Self {
        let last_column = u64::from(columns.saturating_sub(1));
        let last_row = u64::from(rows.saturating_sub(1));
        Self {
            total: [0, 0, last_column, last_row],
            body: [
                u64::from(header_columns),
                u64::from(header_rows),
                last_column,
                u64::from(rows.saturating_sub(footer_rows).saturating_sub(1)),
            ],
        }
    }
}

fn range(payload: &[u8]) -> Result<[u64; 4]> {
    let mut output = [0; 4];
    for (slot, number) in output.iter_mut().zip(1_u32..) {
        *slot = single_varint(payload, number)?.ok_or(Error::Invalid)?;
    }
    if fields(payload)?.len() != 4 {
        return Err(Error::Unsupported);
    }
    Ok(output)
}

fn encode_range(values: [u64; 4]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for (number, value) in (1_u32..).zip(values) {
        output.extend(varint_field(number, value)?);
    }
    Ok(output)
}

fn is_sentinel(values: [u64; 4]) -> bool {
    values == [SENTINEL_COLUMN, SENTINEL_ROW, SENTINEL_COLUMN, SENTINEL_ROW]
}

/// Rewrite the own owner's table extents from `before` to `after`.
///
/// An unset (sentinel) extent stays. An extent that does not describe the
/// table before the edit (a header or footer change did not update it) moves
/// its far edge on the edited axis by one, as the host editor's
/// `mutate_spanning_ranges` does.
pub fn rewrite_extents(
    owner: &[u8],
    before: Extents,
    after: Extents,
    row_axis: bool,
    insert: bool,
) -> Result<Option<Vec<u8>>> {
    let mut changed = false;
    let output = rebuild(owner, |field| {
        if field.number != 7 && field.number != 8 {
            return Ok(None);
        }
        let spanning = rebuild(field.payload, |inner| {
            let (old, new) = match inner.number {
                2 => (before.total, after.total),
                3 => (before.body, after.body),
                _ => return Ok(None),
            };
            let values = range(inner.payload)?;
            if is_sentinel(values) {
                return Ok(None);
            }
            let new = if values == old {
                new
            } else {
                let mut moved = values;
                let edge = &mut moved[if row_axis { 3 } else { 2 }];
                *edge = if insert {
                    edge.checked_add(1).ok_or(Error::Invalid)?
                } else {
                    edge.checked_sub(1).ok_or(Error::Invalid)?
                };
                moved
            };
            if values == new {
                return Ok(None);
            }
            changed = true;
            bytes_field(inner.number, &encode_range(new)?).map(Some)
        })?;
        if spanning == field.payload {
            return Ok(None);
        }
        bytes_field(field.number, &spanning).map(Some)
    })?;
    Ok(changed.then_some(output))
}
