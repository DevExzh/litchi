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

/// Write WINDOW2 record (Worksheet view settings)
///
/// Record type: 0x023E, Length: 18 (worksheet and macro sheet)
/// Writes conservative defaults that are accepted by Excel.
pub fn write_window2<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x023E, 18)?;

    // grbit flags: match Apache POI's InternalSheet.createWindowTwo():
    // options = 0x06B6, which turns on:
    //  - DISPLAY_GRIDLINES
    //  - DISPLAY_ROW_COL_HEADINGS
    //  - DISPLAY_ZEROS
    //  - DEFAULT_HEADER
    //  - DISPLAY_GUTS
    //  - FREEZE_PANES_NO_SPLIT
    //  - ACTIVE
    writer.write_all(&0x06B6u16.to_le_bytes())?;

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
