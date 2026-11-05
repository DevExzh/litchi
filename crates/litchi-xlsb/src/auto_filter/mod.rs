//! The worksheet auto-filter of an XLSB worksheet ([MS-XLSB] sections 2.4.8
//! and 2.4.365).
//!
//! A worksheet stores at most one `BrtBeginAFilter` collection: the filter
//! range followed by its filter columns. [`read`] decodes the range of a
//! worksheet stream; [`edit`] replaces, inserts or removes the collection.

use crate::package::error::{Error, Result};
use crate::part_edit::{SheetLayout, invalid};
use crate::raw::{Records, kind};
pub use crate::sheet::AutoFilter;

pub mod edit;

/// Read the auto-filter range of a worksheet stream.
///
/// # Errors
///
/// Returns an error for a malformed stream or filter range.
pub fn read(worksheet: &[u8], limits: crate::raw::Limits) -> Result<Option<AutoFilter>> {
    read_layout(
        worksheet,
        &SheetLayout::scan_with_limits(worksheet, limits)?,
        limits,
    )
}

pub(crate) fn read_layout(
    worksheet: &[u8],
    layout: &SheetLayout,
    limits: crate::raw::Limits,
) -> Result<Option<AutoFilter>> {
    let mut units = layout.units_of(kind::BEGIN_A_FILTER);
    let Some(unit) = units.next() else {
        return Ok(None);
    };
    if units.next().is_some() {
        return Err(invalid("a worksheet has more than one auto-filter"));
    }
    let header = Records::new(&worksheet[unit.span.clone()], limits)
        .next()
        .transpose()?
        .ok_or_else(|| invalid("an empty auto-filter collection"))?;
    decode(header.payload()).map(Some)
}

fn decode(payload: &[u8]) -> Result<AutoFilter> {
    if payload.len() != 16 {
        return Err(Error::InvalidLength {
            expected: 16,
            found: payload.len(),
        });
    }
    let field = |offset: usize| {
        let mut value = [0u8; 4];
        value.copy_from_slice(&payload[offset..offset + 4]);
        u32::from_le_bytes(value)
    };
    let filter = AutoFilter {
        first_row: field(0),
        last_row: field(4),
        first_column: field(8),
        last_column: field(12),
    };
    check(filter)?;
    Ok(filter)
}

fn check(filter: AutoFilter) -> Result<()> {
    if filter.first_row > filter.last_row
        || filter.last_row > crate::page_breaks::LAST_ROW
        || filter.first_column > filter.last_column
        || filter.last_column > crate::page_breaks::LAST_COLUMN
    {
        return Err(invalid(
            "the auto-filter range is outside the worksheet grid",
        ));
    }
    Ok(())
}

pub(crate) fn encode_range(filter: AutoFilter, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    check(filter)?;
    limits.validate()?;
    if 16 > limits.payload() {
        return Err(Error::Wire(crate::raw::Error::PayloadLimit {
            length: 16,
            limit: limits.payload(),
            offset: 0,
        }));
    }
    let mut payload = Vec::with_capacity(16);
    for field in [
        filter.first_row,
        filter.last_row,
        filter.first_column,
        filter.last_column,
    ] {
        payload.extend_from_slice(&field.to_le_bytes());
    }
    Ok(payload)
}
