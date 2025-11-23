//! BIFF8 NAME (Lbl) record writer.
//!
//! This module encodes workbook-level defined names (named ranges) into
//! BIFF8 `NAME` records, following the [MS-XLS] Lbl structure and the
//! option flag layout used by Apache POI's `NameRecord`.

use std::io::Write;

use crate::ole::xls::writer::XlsDefinedName;
use crate::ole::xls::{XlsError, XlsResult};

use super::{has_multibyte_char, write_record_header};

/// Write a single NAME (Lbl) record.
///
/// The `rgce` argument contains the BIFF8 formula bytes for the defined
/// name (for example, a `PtgRef` or `PtgArea` sequence produced by the
/// formula module).
pub(crate) fn write_name<W: Write>(
    writer: &mut W,
    name: &XlsDefinedName,
    rgce: &[u8],
) -> XlsResult<()> {
    if rgce.len() > u16::MAX as usize {
        return Err(XlsError::InvalidData(
            "Named range formula exceeds BIFF8 length limit".to_string(),
        ));
    }

    // Determine character count and encoding for the Name field. For
    // built-in names we follow the POI convention and store a single
    // byte built-in code with fBuiltin set.
    let (cch, is_16bit, built_in_code, name_str): (u8, bool, Option<u8>, Option<&str>) =
        if name.is_built_in {
            let code = name.built_in_code.ok_or_else(|| {
                XlsError::InvalidData("Built-in defined name requires a built_in_code".to_string())
            })?;
            (1, false, Some(code), None)
        } else {
            let char_count = name.name.chars().count();
            if char_count == 0 {
                return Err(XlsError::InvalidData(
                    "Defined name must not be empty".to_string(),
                ));
            }
            if char_count > u8::MAX as usize {
                return Err(XlsError::InvalidData(
                    "Defined name must be at most 255 characters".to_string(),
                ));
            }
            let cch = char_count as u8;
            let is_16bit = has_multibyte_char(&name.name);
            (cch, is_16bit, None, Some(name.name.as_str()))
        };

    let name_bytes_len: u16 = if is_16bit {
        1u16 + (cch as u16) * 2
    } else {
        1u16 + (cch as u16)
    };

    let cce = rgce.len() as u16;
    let data_len: u16 = 14u16.saturating_add(name_bytes_len).saturating_add(cce);

    // Option flags (grbit) map directly to Apache POI's NameRecord.Option
    // constants for easier comparison:
    //  - 0x0001: hidden name (fHidden)
    //  - 0x0002: function name (fFunc)
    //  - 0x0020: built-in name (fBuiltin)
    let mut grbit: u16 = 0;
    if name.hidden {
        grbit |= 0x0001;
    }
    if name.is_function {
        grbit |= 0x0002;
    }
    if name.is_built_in {
        grbit |= 0x0020;
    }

    write_record_header(writer, 0x0018, data_len)?;

    // grbit
    writer.write_all(&grbit.to_le_bytes())?;

    // chKey (no keyboard shortcut when using defined ranges programmatically)
    writer.write_all(&[0u8])?;

    // cch (character count of Name)
    writer.write_all(&[cch])?;

    // cce (length of rgce in bytes)
    writer.write_all(&cce.to_le_bytes())?;

    // reserved3
    writer.write_all(&0u16.to_le_bytes())?;

    // itab: 0 for workbook-scoped names, or 1-based sheet index for
    // local names.
    let itab = name.local_sheet.unwrap_or(0);
    writer.write_all(&itab.to_le_bytes())?;

    // reserved4..reserved7
    writer.write_all(&[0u8; 4])?;

    // Name (XLUnicodeStringNoCch): flags + characters.
    if let Some(code) = built_in_code {
        // Built-in names are always single-byte codes.
        writer.write_all(&[0x00])?; // compressed 8-bit
        writer.write_all(&[code])?;
    } else if let Some(s) = name_str {
        if is_16bit {
            writer.write_all(&[0x01])?; // UTF-16LE
            for code_unit in s.encode_utf16() {
                writer.write_all(&code_unit.to_le_bytes())?;
            }
        } else {
            writer.write_all(&[0x00])?; // compressed 8-bit
            writer.write_all(s.as_bytes())?;
        }
    }

    // rgce (formula tokens)
    writer.write_all(rgce)?;

    Ok(())
}
