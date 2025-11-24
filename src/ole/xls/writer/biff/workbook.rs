//! Workbook-level BIFF8 record writers.

use crate::ole::xls::XlsResult;
use std::io::Write;

use super::write_record_header;

/// Write FORMAT record (number format string)
///
/// Record type: 0x041E
pub fn write_format_record<W: Write>(
    writer: &mut W,
    index_code: u16,
    format_str: &str,
) -> XlsResult<()> {
    if format_str.is_ascii() {
        let bytes = format_str.as_bytes();
        let cch = bytes.len().min(u16::MAX as usize) as u16;
        let data_len = 2u16 + 2 + 1 + cch; // index_code + cch + flags + chars

        write_record_header(writer, 0x041E, data_len)?;
        writer.write_all(&index_code.to_le_bytes())?;
        writer.write_all(&cch.to_le_bytes())?;
        writer.write_all(&[0x00])?; // compressed 8-bit
        writer.write_all(&bytes[..cch as usize])?;
    } else {
        let utf16: Vec<u16> = format_str.encode_utf16().collect();
        let cch = utf16.len().min(u16::MAX as usize) as u16;
        let data_len = 2u16 + 2 + 1 + cch.saturating_mul(2); // index_code + cch + flags + UTF-16LE

        write_record_header(writer, 0x041E, data_len)?;
        writer.write_all(&index_code.to_le_bytes())?;
        writer.write_all(&cch.to_le_bytes())?;
        writer.write_all(&[0x01])?; // UTF-16LE
        for code_unit in utf16.iter().take(cch as usize) {
            writer.write_all(&code_unit.to_le_bytes())?;
        }
    }

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

pub fn write_workbook_protection<W: Write>(
    writer: &mut W,
    protect_structure: bool,
    protect_windows: bool,
    password_hash: Option<u16>,
) -> XlsResult<()> {
    if !protect_structure && !protect_windows && password_hash.is_none() {
        return Ok(());
    }

    if protect_structure || protect_windows {
        write_record_header(writer, 0x0012, 2)?;
        writer.write_all(&0x0001u16.to_le_bytes())?;
    }

    if protect_windows {
        write_record_header(writer, 0x0019, 2)?;
        writer.write_all(&0x0001u16.to_le_bytes())?;
    }

    if let Some(hash) = password_hash {
        write_record_header(writer, 0x0013, 2)?;
        writer.write_all(&hash.to_le_bytes())?;
    }

    Ok(())
}

/// Write SUPBOOK record for the internal workbook.
///
/// Record type: 0x01AE
///
/// This minimal variant declares that all 3D references refer to the
/// current workbook. The layout for an internal SUPBOOK in BIFF8 is:
///
/// - cTab (2 bytes): number of sheets in the workbook
/// - reserved (2 bytes): MUST be 0x0401
pub fn write_supbook_internal<W: Write>(writer: &mut W, sheet_count: u16) -> XlsResult<()> {
    write_record_header(writer, 0x01AE, 4)?;
    writer.write_all(&sheet_count.to_le_bytes())?;
    writer.write_all(&0x0401u16.to_le_bytes())?;
    Ok(())
}

/// Write EXTERNSHEET record for internal workbook references.
///
/// Record type: 0x0017
///
/// For internal references we generate one XTI entry per worksheet,
/// each mapping to a single sheet in the first (and only) SUPBOOK
/// record. PtgArea3d tokens then use `ixti` to index into this XTI
/// array.
pub fn write_externsheet_internal<W: Write>(writer: &mut W, sheet_count: u16) -> XlsResult<()> {
    if sheet_count == 0 {
        return Ok(());
    }

    let cxti = sheet_count;
    // cXTI (2 bytes) + cXTI * sizeof(XTI)
    let data_len = 2u16.saturating_add(cxti.saturating_mul(6));
    write_record_header(writer, 0x0017, data_len)?;

    // cXTI
    writer.write_all(&cxti.to_le_bytes())?;

    // XTI entries: (ixSupBook, itabFirst, itabLast)
    // ixSupBook is 0 for the first SUPBOOK; sheet indices are 1-based
    // BoundSheet8 indices.
    for i in 0..sheet_count {
        let ix_supbook: u16 = 0;
        let itab: u16 = i.saturating_add(1);
        writer.write_all(&ix_supbook.to_le_bytes())?;
        writer.write_all(&itab.to_le_bytes())?;
        writer.write_all(&itab.to_le_bytes())?;
    }

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
