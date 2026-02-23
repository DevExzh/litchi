//! Data validation binary serialization for XLSB writer.
//!
//! Serializes data validation rules as `BrtDVal` (0x0040) records inside
//! a `BrtBeginDVals`/`BrtEndDVals` (0x023D/0x023E) wrapper, following the
//! binary layout documented by LibreOffice `worksheetfragment.cxx`.
//!
//! # BrtDVal binary layout
//!
//! ```text
//! flags          u32   — packed bit-field (type, error style, operator, booleans)
//! ranges         BinRangeList
//! errorTitle     XLWideString  (i32 char count + UTF-16LE)
//! errorMessage   XLWideString
//! inputTitle     XLWideString
//! inputMessage   XLWideString
//! formula1       BIFF12 formula (i32 cb_ptg + PTG bytes + i32 cb_adddata + adddata)
//! formula2       BIFF12 formula (i32 cb_ptg + PTG bytes + i32 cb_adddata + adddata)
//! ```
//!
//! ## Flag packing (per LibreOffice)
//!
//! | Bits  | Field           |
//! |-------|-----------------|
//! | 0-3   | Validation type |
//! | 4-6   | Error style     |
//! | 7     | String list     |
//! | 8     | Allow blank     |
//! | 9     | No dropdown     |
//! | 18    | Show input msg  |
//! | 19    | Show error msg  |
//! | 20-23 | Operator        |

use crate::ooxml::xlsb::data_validation::DataValidation;
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::writer::RecordWriter;
use crate::ooxml::xlsb::writer::bin_range::{parse_range_list, write_bin_range_list};
use std::io::Write;

/// Write a complete `BrtBeginDVals` / `BrtDVal`* / `BrtEndDVals` block.
pub fn write_data_validations<W: Write>(
    writer: &mut RecordWriter<W>,
    validations: &[DataValidation],
) -> XlsbResult<()> {
    if validations.is_empty() {
        return Ok(());
    }

    // BrtBeginDVals payload: DVals structure (18 bytes) per [MS-XLSB] 2.5.36.
    //
    // Layout:
    //   fWnClosed(1 bit) + reserved(15 bits)  — u16 (2 bytes)
    //   xLeft                                 — u32 (4 bytes)
    //   yTop                                  — u32 (4 bytes)
    //   unused3                               — u32 (4 bytes, MUST be 0)
    //   idvMac                                — u32 (4 bytes, count of BrtDVal)
    let mut dvals_buf = Vec::with_capacity(18);
    dvals_buf.extend_from_slice(&0u16.to_le_bytes()); // flags (fWnClosed=0)
    dvals_buf.extend_from_slice(&0u32.to_le_bytes()); // xLeft
    dvals_buf.extend_from_slice(&0u32.to_le_bytes()); // yTop
    dvals_buf.extend_from_slice(&0u32.to_le_bytes()); // unused3
    dvals_buf.extend_from_slice(&(validations.len() as u32).to_le_bytes()); // idvMac
    writer.write_record(record_types::BEGIN_D_VALS, &dvals_buf)?;

    for dv in validations {
        let payload = serialize_data_validation(dv)?;
        writer.write_record(record_types::D_VAL, &payload)?;
    }

    writer.write_record(record_types::END_D_VALS, &[])?;
    Ok(())
}

/// Serialize a single [`DataValidation`] into the `BrtDVal` binary payload.
fn serialize_data_validation(dv: &DataValidation) -> XlsbResult<Vec<u8>> {
    let mut buf = Vec::with_capacity(128);

    // --- flags (u32) ---
    let mut flags: u32 = 0;
    // bits 0-3: validation type
    flags |= (dv.validation_type as u32) & 0x0F;
    // bits 4-6: error style
    flags |= ((dv.error_style as u32) & 0x07) << 4;
    // bit 7: string list flag (set for list type with comma-separated values)
    if dv.validation_type == 3 {
        // List type — assume string list unless formula starts with '='
        if dv.formula1.as_ref().is_none_or(|f| !f.starts_with('=')) {
            flags |= 0x0080;
        }
    }
    // bit 8: allow blank
    if dv.allow_blank {
        flags |= 0x0100;
    }
    // bit 9: suppress dropdown (inverted semantics from show_dropdown)
    if !dv.show_dropdown {
        flags |= 0x0200;
    }
    // bit 18: show input message
    if dv.show_input_message {
        flags |= 0x0004_0000;
    }
    // bit 19: show error message
    if dv.show_error_message {
        flags |= 0x0008_0000;
    }
    // bits 20-23: operator
    flags |= ((dv.operator as u32) & 0x0F) << 20;

    buf.extend_from_slice(&flags.to_le_bytes());

    // --- BinRangeList ---
    let ranges = parse_range_list(&dv.cell_ranges)?;
    write_bin_range_list(&ranges, &mut buf)?;

    // --- XLWideStrings: errorTitle, errorMessage, inputTitle, inputMessage ---
    write_xl_wide_string(&mut buf, dv.error_title.as_deref().unwrap_or(""));
    write_xl_wide_string(&mut buf, dv.error_text.as_deref().unwrap_or(""));
    write_xl_wide_string(&mut buf, dv.input_title.as_deref().unwrap_or(""));
    write_xl_wide_string(&mut buf, dv.input_text.as_deref().unwrap_or(""));

    // --- formula1 (BIFF12 formula: cb_ptg + PTG bytes + cb_adddata + adddata) ---
    write_biff12_formula(&mut buf, dv.formula1.as_deref());

    // --- formula2 ---
    write_biff12_formula(&mut buf, dv.formula2.as_deref());

    Ok(buf)
}

/// Write an `XLWideString` (u32 character count + UTF-16LE code units).
fn write_xl_wide_string(buf: &mut Vec<u8>, s: &str) {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    buf.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
    for ch in &utf16 {
        buf.extend_from_slice(&ch.to_le_bytes());
    }
}

/// Write a BIFF12 formula.
///
/// Layout (per `OoxFormulaParserImpl::importBiff12Formula` in LO):
///
/// ```text
/// cb_ptg     i32   — byte count of PTG token stream
/// ptg_bytes  [u8]  — PTG tokens (cb_ptg bytes)
/// cb_adddata i32   — byte count of additional data (0 for simple formulas)
/// adddata    [u8]  — additional data (cb_adddata bytes)
/// ```
///
/// Formula values are encoded as the most appropriate PTG token:
/// - Pure integers 0–65535 → `PtgInt` (0x1E) + `u16` value
/// - Floating-point numbers → `PtgNum` (0x1F) + `f64` value
/// - Quoted strings `"text"` → `PtgStr` (0x17) with outer quotes stripped
/// - Everything else → `PtgStr` (0x17) as fallback
///
/// When no formula is provided, we write `cb_ptg = 0` + `cb_adddata = 0`
/// (8 bytes total) since both formulas are always consumed unconditionally
/// by the reader.
fn write_biff12_formula(buf: &mut Vec<u8>, formula: Option<&str>) {
    match formula {
        Some(text) if !text.is_empty() => {
            // Strip leading '=' if present
            let text = text.strip_prefix('=').unwrap_or(text);

            // Build PTG token bytes into a temporary buffer
            let mut ptg = Vec::new();
            build_ptg_for_value(text, &mut ptg);

            // cb_ptg (i32)
            buf.extend_from_slice(&(ptg.len() as i32).to_le_bytes());
            // PTG bytes
            buf.extend_from_slice(&ptg);
            // cb_adddata (i32) — no additional data for simple formulas
            buf.extend_from_slice(&0i32.to_le_bytes());
        },
        _ => {
            // No formula — write cb_ptg=0 + cb_adddata=0
            buf.extend_from_slice(&0i32.to_le_bytes());
            buf.extend_from_slice(&0i32.to_le_bytes());
        },
    }
}

/// Build the most appropriate PTG token for a formula value string.
///
/// # Token formats (per LO `formulaparser.cxx`)
///
/// - `PtgInt`  (0x1E): `opcode(1) + value(u16)` = 3 bytes
/// - `PtgNum`  (0x1F): `opcode(1) + value(f64)` = 9 bytes
/// - `PtgStr`  (0x17): `opcode(1) + cch(i16) + UTF-16LE` (per
///   `BiffHelper::readString(rStrm, false)` — no flags byte)
fn build_ptg_for_value(text: &str, ptg: &mut Vec<u8>) {
    // Try integer first (PtgInt: 0x1E + u16)
    if let Ok(n) = text.parse::<u64>()
        && n <= u16::MAX as u64
    {
        ptg.push(0x1E);
        ptg.extend_from_slice(&(n as u16).to_le_bytes());
        return;
    }

    // Try float (PtgNum: 0x1F + f64)
    if let Ok(f) = text.parse::<f64>() {
        ptg.push(0x1F);
        ptg.extend_from_slice(&f.to_le_bytes());
        return;
    }

    // Strip outer quotes for string literals (e.g. `"Hello"` → `Hello`)
    let s = text
        .strip_prefix('"')
        .and_then(|t| t.strip_suffix('"'))
        .unwrap_or(text);

    // PtgStr: 0x17 + cch(i16) + UTF-16LE
    let utf16: Vec<u16> = s.encode_utf16().collect();
    ptg.push(0x17);
    ptg.extend_from_slice(&(utf16.len() as i16).to_le_bytes());
    for ch in &utf16 {
        ptg.extend_from_slice(&ch.to_le_bytes());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_serialize_list_validation() {
        let dv = DataValidation {
            validation_type: 3, // list
            operator: 0,
            formula1: Some("Item1,Item2,Item3".to_string()),
            formula2: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges: "A1:A10".to_string(),
        };

        let payload = serialize_data_validation(&dv).unwrap();
        // flags at offset 0 should have type=3, allow_blank, string_list, show_error
        let flags = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        assert_eq!(flags & 0x0F, 3); // type = list
        assert_ne!(flags & 0x0080, 0); // string list
        assert_ne!(flags & 0x0100, 0); // allow blank
        assert_eq!(flags & 0x0200, 0); // show dropdown (not suppressed)
        assert_ne!(flags & 0x0008_0000, 0); // show error
    }
}
