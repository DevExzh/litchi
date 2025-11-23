//! Worksheet-level BIFF8 record writers.

use crate::ole::xls::{XlsError, XlsResult};
use std::io::Write;

use super::write_record_header;

/// Write WSBOOL record (Additional Workspace Information)
///
/// Record type: 0x0081, Length: 2
/// Writes default flags indicating a normal worksheet (not dialog sheet).
pub fn write_wsbool<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0081, 2)?;
    // Match Apache POI's InternalSheet.createWSBool():
    //   WSBool1 = 0x04, WSBool2 = 0xC1
    // POI serializes as [WSBool2, WSBool1], so the on-disk u16 (little-endian)
    // is 0x04C1.
    writer.write_all(&0x04C1u16.to_le_bytes())?;
    Ok(())
}

/// Write COLINFO record (column formatting and width).
///
/// Record type: 0x007D, Length: 12
///
/// The width is expressed in units of 1/256 of the width of the
/// character "0" in the workbook's default font, matching the
/// semantics of Apache POI's `ColumnInfoRecord.setColumnWidth`.
pub fn write_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
    hidden: bool,
) -> XlsResult<()> {
    write_record_header(writer, 0x007D, 12)?;

    // Column range (inclusive)
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col.to_le_bytes())?;

    // Column width in 1/256 character units
    writer.write_all(&col_width.to_le_bytes())?;

    // XF index: use the default cell XF (15) to match our cell records.
    writer.write_all(&0x000Fu16.to_le_bytes())?;

    // Options bitfield: base value 0x0002 as in POI, plus the hidden flag
    // in the least significant bit when required.
    let mut options: u16 = 0x0002;
    if hidden {
        options |= 0x0001;
    }
    writer.write_all(&options.to_le_bytes())?;

    // Reserved field: POI commonly writes 2 here; Excel tolerates non-zero
    // even though the spec marks this as reserved.
    writer.write_all(&0x0002u16.to_le_bytes())?;

    Ok(())
}

/// Write PANE record (freeze panes / split panes)
///
/// Record type: 0x0041, Length: 10
///
/// For the initial implementation we only support classic freeze panes,
/// matching Apache POI's use of `PaneRecord` for HSSF:
/// - `x` and `y` are the split positions in terms of columns/rows.
/// - `topRow` and `leftColumn` are set to the same values.
/// - `activePane` is derived from which sides are frozen.
pub fn write_pane<W: Write>(writer: &mut W, freeze_rows: u32, freeze_cols: u16) -> XlsResult<()> {
    if freeze_rows == 0 && freeze_cols == 0 {
        return Ok(());
    }

    let y = u16::try_from(freeze_rows).map_err(|_| {
        XlsError::InvalidData(
            "freeze_panes: freeze_rows exceeds BIFF8 limit 65535 for PANE record".to_string(),
        )
    })?;
    let x = freeze_cols;

    let top_row = y;
    let left_col = x;

    // Active pane constants mirror Apache POI's PaneRecord:
    // 0 = lower-right, 1 = upper-right, 2 = lower-left, 3 = upper-left.
    let active_pane: u16 = match (x > 0, y > 0) {
        (true, true) => 0,  // lower-right
        (true, false) => 1, // upper-right
        (false, true) => 2, // lower-left
        (false, false) => 3,
    };

    write_record_header(writer, 0x0041, 10)?;
    writer.write_all(&x.to_le_bytes())?;
    writer.write_all(&y.to_le_bytes())?;
    writer.write_all(&top_row.to_le_bytes())?;
    writer.write_all(&left_col.to_le_bytes())?;
    writer.write_all(&active_pane.to_le_bytes())?;

    Ok(())
}

/// Write WINDOW2 record (Worksheet view settings)
///
/// Record type: 0x023E, Length: 18 (worksheet and macro sheet)
///
/// When `has_freeze_panes` is true, the FREEZE_PANES (0x0008) and
/// FREEZE_PANES_NO_SPLIT (0x0100) bits are set in the options field,
/// mirroring Apache POI's behaviour after `createFreezePane`.
pub fn write_window2<W: Write>(writer: &mut W, has_freeze_panes: bool) -> XlsResult<()> {
    write_record_header(writer, 0x023E, 18)?;

    // Base options value from POI's InternalSheet.createWindowTwo(): 0x06B6
    let mut options: u16 = 0x06B6;

    if has_freeze_panes {
        // Enable freeze panes and indicate that this is a frozen, not split,
        // window. Bits are defined in POI's WindowTwoRecord as:
        //  - 0x0008: freezePanes
        //  - 0x0100: freezePanesNoSplit
        options |= 0x0008 | 0x0100;
    }

    writer.write_all(&options.to_le_bytes())?;

    // rwTop, colLeft
    writer.write_all(&0u16.to_le_bytes())?; // rwTop = 0
    writer.write_all(&0u16.to_le_bytes())?; // colLeft = 0

    // icvHdr (header color). POI uses 0x40; we mirror that here. The header
    // color is stored as a 32-bit value in POI, but we split it across two
    // u16 fields here; little-endian bytes are identical on disk.
    writer.write_all(&0x0040u16.to_le_bytes())?;

    // reserved2
    writer.write_all(&0u16.to_le_bytes())?;

    // wScaleSLV, wScaleNormal, unused, reserved3
    // POI sets both zooms to 0 and reserved to 0; our split-u16 layout yields
    // the same byte pattern on disk (all zeros) for these trailing fields.
    writer.write_all(&0u16.to_le_bytes())?; // wScaleSLV (page break zoom)
    writer.write_all(&0u16.to_le_bytes())?; // wScaleNormal (normal zoom)
    writer.write_all(&0u16.to_le_bytes())?; // unused
    writer.write_all(&0u16.to_le_bytes())?; // reserved3

    Ok(())
}

/// Write DIMENSIONS record (worksheet dimensions)
///
/// Record type: 0x0200
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `first_row` - First used row
/// * `last_row` - Last used row + 1
/// * `first_col` - First used column
/// * `last_col` - Last used column + 1
pub fn write_dimensions<W: Write>(
    writer: &mut W,
    first_row: u32,
    last_row: u32,
    first_col: u16,
    last_col: u16,
) -> XlsResult<()> {
    write_record_header(writer, 0x0200, 14)?;

    writer.write_all(&first_row.to_le_bytes())?;
    writer.write_all(&last_row.to_le_bytes())?;
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col.to_le_bytes())?;

    // Reserved (must be 0)
    writer.write_all(&0u16.to_le_bytes())?;

    Ok(())
}

/// Write ROW record (row metrics including height and hidden flag).
///
/// Record type: 0x0208, Length: 16
///
/// The height is stored in twips (1/20 of a point) as per MS-XLS
/// and Apache POI's `RowRecord` implementation.
pub fn write_row<W: Write>(
    writer: &mut W,
    row_index: u32,
    first_col: u16,
    last_col_plus1: u16,
    height: u16,
    hidden: bool,
) -> XlsResult<()> {
    let row_u16 = u16::try_from(row_index).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for ROW record",
            row_index
        ))
    })?;

    write_record_header(writer, 0x0208, 16)?;

    // Row number
    writer.write_all(&row_u16.to_le_bytes())?;

    // First and last used column indices. A value of 0 for both is
    // accepted by Excel for empty rows (mirrors POI's `setEmpty`).
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col_plus1.to_le_bytes())?;

    // Row height in twips
    writer.write_all(&height.to_le_bytes())?;

    // Optimization hint and reserved fields: keep both at zero, as
    // POI does for generated sheets.
    writer.write_all(&0u16.to_le_bytes())?; // optimize
    writer.write_all(&0u16.to_le_bytes())?; // reserved

    // Option flags: always set bit 8 (0x0100) as in POI's
    // OPTION_BITS_ALWAYS_SET, and toggle the zeroHeight bit (0x0020)
    // when the row is hidden. When a custom height is used
    // (height != 0x00FF), also set the badFontHeight bit (0x0040),
    // mirroring HSSFRow.setHeightInPoints and RowRecord.
    let mut option_flags: u16 = 0x0100;
    if hidden {
        option_flags |= 0x0020;
    }
    if height != 0x00FF {
        option_flags |= 0x0040;
    }
    writer.write_all(&option_flags.to_le_bytes())?;

    // Secondary option flags, including the XF index and border bits.
    // For now we leave this at POI's default of 0x000F.
    writer.write_all(&0x000Fu16.to_le_bytes())?;

    Ok(())
}

pub fn write_mergedcells<W, I>(writer: &mut W, ranges: I) -> XlsResult<()>
where
    W: Write,
    I: IntoIterator<Item = (u32, u32, u16, u16)>,
{
    const MAX_MERGED_REGIONS: usize = 1027;

    let mut chunk: Vec<(u16, u16, u16, u16)> = Vec::new();

    for (first_row_u32, last_row_u32, first_col, last_col) in ranges {
        let first_row = u16::try_from(first_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for MERGEDCELLS record",
                first_row_u32
            ))
        })?;
        let last_row = u16::try_from(last_row_u32).map_err(|_| {
            XlsError::InvalidData(format!(
                "Row index {} exceeds BIFF8 limit 65535 for MERGEDCELLS record",
                last_row_u32
            ))
        })?;

        chunk.push((first_row, last_row, first_col, last_col));

        if chunk.len() == MAX_MERGED_REGIONS {
            write_mergedcells_chunk(writer, &chunk)?;
            chunk.clear();
        }
    }

    if !chunk.is_empty() {
        write_mergedcells_chunk(writer, &chunk)?;
    }

    Ok(())
}

fn write_mergedcells_chunk<W: Write>(
    writer: &mut W,
    ranges: &[(u16, u16, u16, u16)],
) -> XlsResult<()> {
    debug_assert!(!ranges.is_empty());
    debug_assert!(ranges.len() <= 1027);

    let count = u16::try_from(ranges.len()).expect("MERGEDCELLS range count fits in u16");
    let data_len: u16 = 2u16 + count.saturating_mul(8);

    write_record_header(writer, 0x00E5, data_len)?;
    writer.write_all(&count.to_le_bytes())?;

    for &(first_row, last_row, first_col, last_col) in ranges {
        writer.write_all(&first_row.to_le_bytes())?;
        writer.write_all(&last_row.to_le_bytes())?;
        writer.write_all(&first_col.to_le_bytes())?;
        writer.write_all(&last_col.to_le_bytes())?;
    }

    Ok(())
}
