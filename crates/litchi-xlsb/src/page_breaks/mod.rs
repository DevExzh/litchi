//! Manual page breaks of an XLSB worksheet ([MS-XLSB] sections 2.4.25,
//! 2.4.197 and 2.4.315).
//!
//! A worksheet stores its horizontal breaks in one `BrtBeginRwBrk`
//! collection and its vertical breaks in one `BrtBeginColBrk` collection,
//! each a list of `BrtBrk` records. [`read`] decodes both from a worksheet
//! stream; [`edit`] replaces them in place.

use crate::package::error::{Error, Result};
use crate::part_edit::{SheetLayout, Unit, invalid};
use crate::raw::{Kind, Records, kind};

pub mod edit;

/// The most breaks one collection holds.
pub const MAX_BREAKS: usize = 1_023;
/// The last row index of a worksheet.
pub const LAST_ROW: u32 = 1_048_575;
/// The last column index of a worksheet.
pub const LAST_COLUMN: u32 = 16_383;

/// One manual page break.
///
/// A row break (`rows`) sits above row `id` and spans the columns
/// `min..=max`; a column break sits left of column `id` and spans the rows
/// `min..=max`. Indexes are zero-based, as the binary records store them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Break {
    /// The zero-based row or column index of the break.
    pub id: u32,
    /// The first column (row breaks) or row (column breaks) it spans.
    pub min: u32,
    /// The last column (row breaks) or row (column breaks) it spans.
    pub max: u32,
    /// Whether the user placed the break (`fMan`).
    pub manual: bool,
    /// Whether a PivotTable created the break (`fPivot`).
    pub pivot: bool,
}

impl Break {
    /// A manual row break above `row` spanning every column.
    #[must_use]
    pub const fn row(row: u32) -> Self {
        Self {
            id: row,
            min: 0,
            max: LAST_COLUMN,
            manual: true,
            pivot: false,
        }
    }

    /// A manual column break left of `column` spanning every row.
    #[must_use]
    pub const fn column(column: u32) -> Self {
        Self {
            id: column,
            min: 0,
            max: LAST_ROW,
            manual: true,
            pivot: false,
        }
    }
}

/// The manual page breaks of a worksheet.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageBreaks {
    /// Horizontal breaks (between rows), in stream order.
    pub rows: Vec<Break>,
    /// Vertical breaks (between columns), in stream order.
    pub columns: Vec<Break>,
}

/// Read the page breaks of a worksheet stream.
///
/// # Errors
///
/// Returns an error for a malformed stream or break collection.
pub fn read(worksheet: &[u8], limits: crate::raw::Limits) -> Result<PageBreaks> {
    let layout = SheetLayout::scan_with_limits(worksheet, limits)?;
    read_layout(worksheet, &layout, limits)
}

pub(crate) fn read_layout(
    worksheet: &[u8],
    layout: &SheetLayout,
    limits: crate::raw::Limits,
) -> Result<PageBreaks> {
    let collection = |begin: Kind, row: bool| -> Result<Vec<Break>> {
        let mut units = layout.units_of(begin);
        let Some(unit) = units.next() else {
            return Ok(Vec::new());
        };
        if units.next().is_some() {
            return Err(invalid(
                "a worksheet has more than one page-break collection",
            ));
        }
        parse(&worksheet[unit.span.clone()], unit, row, limits)
    };
    Ok(PageBreaks {
        rows: collection(kind::BEGIN_RW_BRK, true)?,
        columns: collection(kind::BEGIN_COL_BRK, false)?,
    })
}

fn parse(bytes: &[u8], unit: &Unit, row: bool, limits: crate::raw::Limits) -> Result<Vec<Break>> {
    let end = if row {
        kind::END_RW_BRK
    } else {
        kind::END_COL_BRK
    };
    let mut records = Records::new(bytes, limits);
    let header = records
        .next()
        .transpose()?
        .ok_or_else(|| invalid("an empty page-break collection"))?;
    if header.kind() != unit.kind || header.payload().len() != 8 {
        return Err(Error::InvalidLength {
            expected: 8,
            found: header.payload().len(),
        });
    }
    let declared = u32_at(header.payload(), 0) as usize;
    let mut breaks = Vec::new();
    for record in records {
        let record = record?;
        if record.kind() == end {
            break;
        }
        if record.kind() != kind::BRK || record.payload().len() != 20 {
            return Err(invalid(
                "a page-break collection holds a record that is not BrtBrk",
            ));
        }
        let payload = record.payload();
        let item = Break {
            id: u32_at(payload, 0),
            min: u32_at(payload, 4),
            max: u32_at(payload, 8),
            manual: u32_at(payload, 12) != 0,
            pivot: u32_at(payload, 16) != 0,
        };
        breaks.push(item);
    }
    if breaks.len() != declared {
        return Err(invalid("a page-break collection declares breaks and holds"));
    }
    Ok(breaks)
}

fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    let mut value = [0u8; 4];
    value.copy_from_slice(&bytes[offset..offset + 4]);
    u32::from_le_bytes(value)
}

/// Validate and encode one collection; an empty list encodes nothing.
pub(crate) fn encode(breaks: &[Break], row: bool, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    limits.validate()?;
    if breaks.is_empty() {
        return Ok(Vec::new());
    }
    if breaks.len() > MAX_BREAKS {
        return Err(invalid(
            "a worksheet holds at most page breaks of one direction",
        ));
    }
    let (id_limit, span_limit) = if row {
        (LAST_ROW, LAST_COLUMN)
    } else {
        (LAST_COLUMN, LAST_ROW)
    };
    let mut seen = std::collections::BTreeSet::new();
    for item in breaks {
        if item.id > id_limit || item.max > span_limit || item.min > item.max {
            return Err(invalid("a page break is outside the worksheet grid"));
        }
        if !item.manual {
            return Err(invalid(
                "an XLSB worksheet stores manual page breaks only (fMan must be 1)",
            ));
        }
        if !seen.insert(item.id) {
            return Err(invalid("two page breaks share one index"));
        }
    }
    let (begin, end) = if row {
        (kind::BEGIN_RW_BRK, kind::END_RW_BRK)
    } else {
        (kind::BEGIN_COL_BRK, kind::END_COL_BRK)
    };
    let count = u32::try_from(breaks.len()).map_err(|_overflow| invalid("too many page breaks"))?;
    let mut header = Vec::with_capacity(8);
    header.extend_from_slice(&count.to_le_bytes());
    header.extend_from_slice(&count.to_le_bytes());
    let mut out = crate::part_edit::record_with_limits(begin, &header, limits)?;
    for item in breaks {
        let mut payload = Vec::with_capacity(20);
        for field in [
            item.id,
            item.min,
            item.max,
            u32::from(item.manual),
            u32::from(item.pivot),
        ] {
            payload.extend_from_slice(&field.to_le_bytes());
        }
        out.extend(crate::part_edit::record_with_limits(
            kind::BRK,
            &payload,
            limits,
        )?);
    }
    out.extend(crate::part_edit::record_with_limits(end, &[], limits)?);
    Ok(out)
}
