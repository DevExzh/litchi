//! BIFF record writer for XLS files
//!
//! This module provides functions to generate BIFF8 (Binary Interchange File Format)
//! records for writing XLS files. BIFF8 is the format used by Excel 97-2003.
//!
//! # BIFF Record Structure
//!
//! Each BIFF record consists of:
//! - Record type (2 bytes) - identifies the record
//! - Record length (2 bytes) - length of data in bytes
//! - Record data (variable length)
//!
//! # References
//!
//! Based on Microsoft's "[MS-XLS]: Excel Binary File Format (.xls) Structure" specification
//! and Apache POI's BIFF record generation.

use super::super::{XlsError, XlsResult};
use std::io::Write;

/// Write a BIFF record header
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `record_type` - BIFF record type (e.g., 0x0809 for BOF)
/// * `data_len` - Length of record data in bytes
#[inline]
pub(crate) fn write_record_header<W: Write>(
    writer: &mut W,
    record_type: u16,
    data_len: u16,
) -> XlsResult<()> {
    writer.write_all(&record_type.to_le_bytes())?;
    writer.write_all(&data_len.to_le_bytes())?;
    Ok(())
}

/// Write STYLE record (built-in style)
///
/// Record type: 0x0293
///
/// This helper writes only built-in styles, which use the compact 4-byte
/// payload:
///  - field_1_xf_index (2 bytes): low 12 bits = XF index, bit 15 = isBuiltIn
///  - builtinStyle (1 byte): built-in style identifier (e.g., 0 = Normal)
///  - outlineLevel (1 byte): usually 0xFF for non-outline styles
fn write_style_builtin<W: Write>(
    writer: &mut W,
    xf_index: u16,
    builtin_style_id: u8,
) -> XlsResult<()> {
    // Mask to 12 bits, then set the built-in flag in bit 15.
    let xf_field: u16 = (xf_index & 0x0FFF) | 0x8000;

    write_record_header(writer, 0x0293, 4)?;
    writer.write_all(&xf_field.to_le_bytes())?;
    writer.write_all(&[builtin_style_id])?;
    // Match POI's use of 0xFF ("no outline level").
    writer.write_all(&[0xFF])?;
    Ok(())
}

/// Write the minimal built-in STYLE records used by Excel / Apache POI.
///
/// This mirrors POI's `InternalWorkbook.createStyle(id)` mapping while keeping
/// the implementation compact. The XF indices assume the following XF table:
///
/// - 0..14: style XFs
/// - 15:    default cell XF
/// - 16..20: additional style XFs for comma / currency / percent styles
///
/// Mapping (xf_index, builtin_style_id):
/// - (0x0010, 3)  => Comma
/// - (0x0011, 6)  => Comma [0 decimals]
/// - (0x0012, 4)  => Currency
/// - (0x0013, 7)  => Currency [0 decimals]
/// - (0x0000, 0)  => Normal
/// - (0x0014, 5)  => Percent
pub fn write_builtin_styles<W: Write>(writer: &mut W) -> XlsResult<()> {
    // Order follows POI for easier comparison, but Excel only cares about
    // the XF indices and builtin IDs, not the sequence.
    const MAPPINGS: &[(u16, u8)] = &[
        (0x0010, 3),
        (0x0011, 6),
        (0x0012, 4),
        (0x0013, 7),
        (0x0000, 0),
        (0x0014, 5),
    ];

    for &(xf_index, builtin_id) in MAPPINGS {
        write_style_builtin(writer, xf_index, builtin_id)?;
    }

    Ok(())
}

/// Write UseSelFS (Use Natural Language Formulas) record.
///
/// Record type: 0x0160, Length: 2
/// A value of 0 disables natural language formulas (modern Excel default).
pub fn write_usesel_fs<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0160, 2)?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

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

/// Write BOF (Beginning of File) record
///
/// Record type: 0x0809
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `substream_type` - Type of substream (0x0005 = Workbook, 0x0010 = Worksheet)
pub fn write_bof<W: Write>(writer: &mut W, substream_type: u16) -> XlsResult<()> {
    write_record_header(writer, 0x0809, 16)?;

    // BIFF version (0x0600 = BIFF8)
    writer.write_all(&0x0600u16.to_le_bytes())?;

    // Substream type
    writer.write_all(&substream_type.to_le_bytes())?;

    // Build identifier (arbitrary)
    writer.write_all(&0x0DBBu16.to_le_bytes())?;

    // Build year (e.g., 1996)
    writer.write_all(&0x07CCu16.to_le_bytes())?;

    // File history flags (0x00000000)
    writer.write_all(&0x00000000u32.to_le_bytes())?;

    // Lowest BIFF version (0x06 = BIFF8)
    writer.write_all(&0x00000006u32.to_le_bytes())?;

    Ok(())
}

/// Write EOF (End of File) record
///
/// Record type: 0x000A
pub fn write_eof<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x000A, 0)?;
    Ok(())
}

/// Write CODEPAGE record
///
/// Record type: 0x0042
///
/// # Arguments
///
/// * `writer` - Output writer  
/// * `codepage` - Code page identifier (default: 1252 for Windows Latin 1)
pub fn write_codepage<W: Write>(writer: &mut W, codepage: u16) -> XlsResult<()> {
    write_record_header(writer, 0x0042, 2)?;
    writer.write_all(&codepage.to_le_bytes())?;
    Ok(())
}

/// Write DATE1904 record
///
/// Record type: 0x0022
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `is_1904` - True for 1904 date system (Mac), false for 1900 (Windows)
pub fn write_date1904<W: Write>(writer: &mut W, is_1904: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0022, 2)?;
    let flag = if is_1904 { 1u16 } else { 0u16 };
    writer.write_all(&flag.to_le_bytes())?;
    Ok(())
}

/// Write WINDOW1 record (workbook window properties)
///
/// Record type: 0x003D
pub fn write_window1<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x003D, 18)?;

    // Window position and size (default values)
    writer.write_all(&0u16.to_le_bytes())?; // xWn: horizontal position
    writer.write_all(&0u16.to_le_bytes())?; // yWn: vertical position
    writer.write_all(&0x3000u16.to_le_bytes())?; // dxWn: width
    writer.write_all(&0x1E00u16.to_le_bytes())?; // dyWn: height

    // Options flags
    writer.write_all(&0x0038u16.to_le_bytes())?; // grbit: various flags

    // Active sheet index
    writer.write_all(&0u16.to_le_bytes())?; // itabCur

    // First displayed sheet tab
    writer.write_all(&0u16.to_le_bytes())?; // itabFirst

    // Number of selected sheets
    writer.write_all(&1u16.to_le_bytes())?; // ctabSel

    // Ratio of width of tab to width of horizontal scroll bar
    writer.write_all(&0x0258u16.to_le_bytes())?; // wTabRatio

    Ok(())
}

/// Write BOUNDSHEET8 record (worksheet metadata)
///
/// Record type: 0x0085
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `position` - Absolute stream position of BOF record for this sheet
/// * `name` - Sheet name (max 31 characters)
///
/// The sheet name is encoded as ShortXLUnicodeString per BIFF8: 1-byte character count,
/// 1-byte flags (0x00 = compressed 8-bit, 0x01 = uncompressed UTF-16LE), followed by characters.
pub fn write_boundsheet<W: Write>(writer: &mut W, position: u32, name: &str) -> XlsResult<()> {
    let truncated = if name.len() > 31 { &name[..31] } else { name };

    // Determine encoding: use compressed 8-bit if all ASCII; otherwise UTF-16LE
    let is_ascii = truncated.is_ascii();
    let (cch, flags, name_bytes_vec): (u8, u8, Vec<u8>) = if is_ascii {
        let bytes = truncated.as_bytes();
        (bytes.len() as u8, 0x00, bytes.to_vec())
    } else {
        // UTF-16LE encoding
        let utf16: Vec<u16> = truncated.encode_utf16().collect();
        let mut buf = Vec::with_capacity(utf16.len() * 2);
        for ch in &utf16 {
            buf.extend_from_slice(&ch.to_le_bytes());
        }
        (utf16.len() as u8, 0x01, buf)
    };

    // position(4) + options(2) + cch(1) + flags(1) + name bytes
    let name_bytes_len: u16 = if is_ascii {
        cch as u16
    } else {
        (cch as u16) * 2
    };
    let data_len = 4u16 + 2u16 + 1u16 + 1u16 + name_bytes_len; // 8 + name length
    write_record_header(writer, 0x0085, data_len)?;

    // Absolute stream position
    writer.write_all(&position.to_le_bytes())?;

    // Sheet state and type (0x0000 = visible worksheet, type = worksheet)
    writer.write_all(&0x0000u16.to_le_bytes())?;

    // ShortXLUnicodeString: cch, flags, chars
    writer.write_all(&[cch])?;
    writer.write_all(&[flags])?;
    writer.write_all(&name_bytes_vec[..(cch as usize) * if is_ascii { 1 } else { 2 }])?;

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

/// Write NUMBER record (floating point cell)
///
/// Record type: 0x0203
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Cell value (f64)
pub fn write_number<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: f64,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for NUMBER record",
            row
        ))
    })?;

    write_record_header(writer, 0x0203, 14)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // IEEE 754 floating point value
    writer.write_all(&value.to_le_bytes())?;

    Ok(())
}

/// Write LABELSST record (string cell with reference to SST)
///
/// Record type: 0x00FD
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `sst_index` - Index into shared string table
pub fn write_labelsst<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    sst_index: u32,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for LABELSST record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 4 (sst index) = 10 bytes
    write_record_header(writer, 0x00FD, 10)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // SST index
    writer.write_all(&sst_index.to_le_bytes())?;

    Ok(())
}

/// Write BOOLERR record (boolean or error cell)
///
/// Record type: 0x0205
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `row` - Row index (0-based)
/// * `col` - Column index (0-based)
/// * `value` - Boolean value
pub fn write_boolerr<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: bool,
) -> XlsResult<()> {
    // BIFF8 stores row as a 16-bit index (0..65535)
    let row_u16 = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for BOOLERR record",
            row
        ))
    })?;

    // 2 (row) + 2 (col) + 2 (xf) + 1 (value) + 1 (is-error flag) = 8 bytes
    write_record_header(writer, 0x0205, 8)?;

    writer.write_all(&row_u16.to_le_bytes())?;
    writer.write_all(&col.to_le_bytes())?;

    // XF record index
    writer.write_all(&xf_index.to_le_bytes())?;

    // Boolean value (0 = false, 1 = true) + error flag (0 = boolean)
    writer.write_all(&[if value { 1 } else { 0 }, 0])?;

    Ok(())
}

/// Write CONTINUE record
///
/// Record type: 0x003C
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `data` - Continuation data
fn write_continue<W: Write>(writer: &mut W, data: &[u8]) -> XlsResult<()> {
    let len = data.len().min(8224) as u16; // Max record size
    write_record_header(writer, 0x003C, len)?;
    writer.write_all(&data[..len as usize])?;
    Ok(())
}

/// Write SST (Shared String Table) record with CONTINUE support
///
/// Record type: 0x00FC
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `strings` - Vector of strings to include in SST
///
/// # Implementation Notes
///
/// The SST record has a maximum size of 8224 bytes. If the SST exceeds this size,
/// CONTINUE records (0x003C) are used to store the remaining data.
///
/// This implementation properly handles string splitting across CONTINUE boundaries,
/// based on Apache POI's SSTSerializer.
pub fn write_sst<W: Write>(writer: &mut W, strings: &[String], cst_total: u32) -> XlsResult<()> {
    const MAX_RECORD_DATA: usize = 8224; // max data payload per record

    // We'll build each record's payload into a local buffer, then flush with the header
    let mut first_record = true;
    let mut buffer: Vec<u8> = Vec::with_capacity(MAX_RECORD_DATA);

    // Helper to flush current buffer as either SST or CONTINUE
    let flush = |writer: &mut W, buf: &mut Vec<u8>, first: bool| -> XlsResult<()> {
        if buf.is_empty() {
            return Ok(());
        }
        if first {
            write_record_header(writer, 0x00FC, buf.len() as u16)?;
            writer.write_all(buf)?;
        } else {
            // Delegate CONTINUE records to the helper so we keep the
            // record-writing logic in one place.
            write_continue(writer, buf)?;
        }
        buf.clear();
        Ok(())
    };

    // Initialize first SST record with cstTotal and cstUnique
    buffer.extend_from_slice(&cst_total.to_le_bytes());
    buffer.extend_from_slice(&(strings.len() as u32).to_le_bytes());

    // Available payload for this record
    let mut available = MAX_RECORD_DATA - buffer.len();

    for s in strings {
        let is_ascii = s.is_ascii();
        let cch: usize;
        let mut data8: &[u8] = &[];
        let mut data16: Vec<u8> = Vec::new();
        let high_byte_flag: u8;

        if is_ascii {
            let bytes = s.as_bytes();
            cch = bytes.len().min(0xFFFF);
            data8 = &bytes[..cch];
            high_byte_flag = 0x00;
        } else {
            let utf16: Vec<u16> = s.encode_utf16().collect();
            cch = utf16.len().min(0xFFFF);
            data16.reserve_exact(cch * 2);
            for ch in utf16.iter().take(cch) {
                data16.extend_from_slice(&ch.to_le_bytes());
            }
            high_byte_flag = 0x01;
        }

        // String header is 3 bytes (cch u16 + flags u8). Ensure it fits fully in current record.
        if available < 3 {
            // Flush current record
            flush(writer, &mut buffer, first_record)?;
            first_record = false;
            // Start CONTINUE record; for a new string we do not need leading high-byte flag
            available = MAX_RECORD_DATA;
        }

        // Write header
        buffer.extend_from_slice(&(cch as u16).to_le_bytes());
        buffer.push(high_byte_flag);
        available -= 3;

        // Now write character data, possibly across CONTINUE records
        if high_byte_flag == 0x00 {
            // Compressed 8-bit
            let mut offset = 0;
            while offset < cch {
                let can_write = available.min(cch - offset);
                if can_write == 0 {
                    // Flush and start CONTINUE; for continued strings, first byte is compression flag
                    flush(writer, &mut buffer, first_record)?;
                    first_record = false;
                    buffer.push(high_byte_flag); // continuation header for string
                    available = MAX_RECORD_DATA - 1;
                    continue;
                }
                buffer.extend_from_slice(&data8[offset..offset + can_write]);
                offset += can_write;
                available -= can_write;
            }
        } else {
            // UTF-16LE (2 bytes per char), do not split a character
            let total_bytes = cch * 2;
            let mut written = 0;
            while written < total_bytes {
                // space available in bytes, but keep even number to not split a char
                let mut can_write = available.min(total_bytes - written);
                if can_write == 0 {
                    // Flush and start CONTINUE; for continued strings, first byte is compression flag
                    flush(writer, &mut buffer, first_record)?;
                    first_record = false;
                    buffer.push(high_byte_flag); // continuation header for string
                    available = MAX_RECORD_DATA - 1;
                    continue;
                }
                // ensure even number of bytes
                if !can_write.is_multiple_of(2) {
                    if can_write == 1 {
                        // no space for a full char
                        flush(writer, &mut buffer, first_record)?;
                        first_record = false;
                        buffer.push(high_byte_flag);
                        available = MAX_RECORD_DATA - 1;
                        continue;
                    } else {
                        can_write -= 1;
                    }
                }
                buffer.extend_from_slice(&data16[written..written + can_write]);
                written += can_write;
                available -= can_write;
            }
        }
    }

    // Flush any remaining data
    flush(writer, &mut buffer, first_record)?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_bof() {
        let mut buf = Vec::new();
        write_bof(&mut buf, 0x0005).unwrap();

        // Check record type and length
        assert_eq!(&buf[0..2], &[0x09, 0x08]); // Record type 0x0809
        assert_eq!(&buf[2..4], &[16, 0]); // Length = 16
    }

    #[test]
    fn test_write_eof() {
        let mut buf = Vec::new();
        write_eof(&mut buf).unwrap();

        assert_eq!(&buf[0..2], &[0x0A, 0x00]); // Record type 0x000A
        assert_eq!(&buf[2..4], &[0, 0]); // Length = 0
    }

    #[test]
    fn test_write_number() {
        let mut buf = Vec::new();
        write_number(&mut buf, 0, 0, 0x000F, 42.5).unwrap();

        assert_eq!(&buf[0..2], &[0x03, 0x02]); // Record type 0x0203
        assert_eq!(&buf[2..4], &[14, 0]); // Length = 14
    }
}
