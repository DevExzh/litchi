//! Data validation binary serialization for XLSB writer.
//!
//! Serializes classic `BrtDVal` records and Office 2013 `BrtDVal14` future
//! records, including lossless pre-encoded formulas and validation UI state.
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

use crate::package::data_validation::{
    RecordKind, Settings, Validation, validate_dval_list_formula,
};
use crate::package::error::{Error, Result};
use crate::package::formula::ParsedFormula;
use crate::raw::Writer;
use crate::raw::kind;
use crate::writer::bin_range::{parse_range_list, write_bin_range_list};
use std::io::Write;

/// Write classic and Office 2013 validation collections as required.
pub fn write_data_validations<W: Write>(
    writer: &mut Writer<W>,
    validations: &[Validation],
    classic_settings: Settings,
    extension14_settings: Settings,
) -> Result<()> {
    if validations.is_empty() {
        return Ok(());
    }
    for rule in validations {
        validate_rule(rule)?;
    }

    let classic = validations
        .iter()
        .filter(|rule| rule.record_kind == RecordKind::Classic)
        .collect::<Vec<_>>();
    let extension14 = validations
        .iter()
        .filter(|rule| rule.record_kind == RecordKind::Extension14)
        .collect::<Vec<_>>();
    if classic.len() > 65_534 || extension14.len() > 65_534 {
        return Err(invalid("a validation collection exceeds 65,534 rules"));
    }
    write_classic_collection(writer, &classic, classic_settings)?;
    write_extension14_collection(writer, &extension14, extension14_settings)
}

fn write_classic_collection<W: Write>(
    writer: &mut Writer<W>,
    validations: &[&Validation],
    settings: Settings,
) -> Result<()> {
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
    dvals_buf.extend_from_slice(&u16::from(settings.input_prompts_disabled).to_le_bytes());
    dvals_buf.extend_from_slice(&u32::from(settings.prompt_x).to_le_bytes());
    dvals_buf.extend_from_slice(&u32::from(settings.prompt_y).to_le_bytes());
    dvals_buf.extend_from_slice(&0u32.to_le_bytes()); // unused3
    dvals_buf.extend_from_slice(&(validations.len() as u32).to_le_bytes()); // idvMac
    writer.write_record(kind::BEGIN_D_VALS, &dvals_buf)?;

    for &dv in validations {
        if let Some(list_formula) = &dv.list_formula {
            let mut payload = Vec::new();
            write_xl_wide_string(&mut payload, list_formula);
            writer.write_record(kind::D_VAL_LIST, &payload)?;
        }
        let payload = serialize_data_validation(dv)?;
        writer.write_record(kind::D_VAL, &payload)?;
    }

    writer.write_record(kind::END_D_VALS, &[])?;
    Ok(())
}

fn write_extension14_collection<W: Write>(
    writer: &mut Writer<W>,
    validations: &[&Validation],
    settings: Settings,
) -> Result<()> {
    if validations.is_empty() {
        return Ok(());
    }
    let mut begin = Vec::with_capacity(22);
    begin.extend_from_slice(&0u32.to_le_bytes());
    begin.extend_from_slice(&u16::from(settings.input_prompts_disabled).to_le_bytes());
    begin.extend_from_slice(&u32::from(settings.prompt_x).to_le_bytes());
    begin.extend_from_slice(&u32::from(settings.prompt_y).to_le_bytes());
    begin.extend_from_slice(&0u32.to_le_bytes());
    begin.extend_from_slice(&(validations.len() as u32).to_le_bytes());
    writer.write_record(kind::BEGIN_D_VALS14, &begin)?;
    for &validation in validations {
        writer.write_record(
            kind::D_VAL14,
            &serialize_extension14_validation(validation)?,
        )?;
    }
    writer.write_record(kind::END_D_VALS14, &[])?;
    Ok(())
}

/// Serialize a single [`Validation`] into the `BrtDVal` binary payload.
fn serialize_data_validation(dv: &Validation) -> Result<Vec<u8>> {
    let mut buf = Vec::with_capacity(128);

    // --- flags (u32) ---
    let mut flags: u32 = 0;
    // bits 0-3: validation type
    flags |= (dv.validation_type as u32) & 0x0F;
    // bits 4-6: error style
    flags |= ((dv.error_style as u32) & 0x07) << 4;
    // bit 7: string list flag (set for list type with comma-separated values)
    if dv.string_list {
        flags |= 0x0080;
    }
    // bit 8: allow blank
    if dv.allow_blank {
        flags |= 0x0100;
    }
    // bit 9: suppress dropdown (inverted semantics from show_dropdown)
    if !dv.show_dropdown {
        flags |= 0x0200;
    }
    // bits 10-17: IME mode
    flags |= u32::from(dv.ime_mode) << 10;
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
    write_xl_nullable_wide_string(&mut buf, dv.error_title.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.error_text.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.input_title.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.input_text.as_deref());

    // --- formula1 (BIFF12 formula: cb_ptg + PTG bytes + cb_adddata + adddata) ---
    let formula1_text =
        if dv.formula1.is_none() && dv.formula1_binary.is_none() && dv.list_formula.is_some() {
            Some("\"\"")
        } else {
            dv.formula1.as_deref().or(dv.list_formula.as_deref())
        };
    write_biff12_formula(
        &mut buf,
        dv.formula1_binary.as_ref(),
        formula1_text,
        dv.string_list,
    )?;

    // --- formula2 ---
    write_biff12_formula(
        &mut buf,
        dv.formula2_binary.as_ref(),
        dv.formula2.as_deref(),
        false,
    )?;

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

fn write_xl_nullable_wide_string(buf: &mut Vec<u8>, value: Option<&str>) {
    match value {
        Some(value) => write_xl_wide_string(buf, value),
        None => buf.extend_from_slice(&u32::MAX.to_le_bytes()),
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
/// Formula text is compiled with the XLSB formula compiler. Inline string-list
/// values use a single `PtgStr`, matching Excel and LibreOffice behavior.
///
/// When no formula is provided, we write `cb_ptg = 0` + `cb_adddata = 0`
/// (8 bytes total) since both formulas are always consumed unconditionally
/// by the reader.
fn write_biff12_formula(
    buf: &mut Vec<u8>,
    binary: Option<&ParsedFormula>,
    formula: Option<&str>,
    string_list: bool,
) -> Result<()> {
    let compiled;
    let binary = if let Some(binary) = binary {
        Some(binary)
    } else if let Some(text) = formula.filter(|text| !text.is_empty()) {
        compiled = if string_list {
            compile_string_list(text)?
        } else {
            crate::package::formula::text::Compiler::compile(text)?
        };
        Some(&compiled)
    } else {
        None
    };
    if let Some(binary) = binary {
        buf.extend_from_slice(&(binary.rgce.len() as u32).to_le_bytes());
        buf.extend_from_slice(&binary.rgce);
        buf.extend_from_slice(&(binary.rgcb.len() as u32).to_le_bytes());
        buf.extend_from_slice(&binary.rgcb);
        return Ok(());
    }
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}

fn serialize_extension14_validation(dv: &Validation) -> Result<Vec<u8>> {
    let ranges = parse_range_list(&dv.cell_ranges)?;
    let formula1 = compile_formula(
        dv.formula1_binary.as_ref(),
        dv.formula1.as_deref(),
        dv.string_list,
    )?;
    let formula2 = compile_formula(dv.formula2_binary.as_ref(), dv.formula2.as_deref(), false)?;
    let formula_count = usize::from(formula1.is_some()) + usize::from(formula2.is_some());
    let mut buf = Vec::with_capacity(160);
    let header_flags: u32 = 0x02 | if formula_count == 0 { 0 } else { 0x04 };
    buf.extend_from_slice(&header_flags.to_le_bytes());
    buf.extend_from_slice(&1u32.to_le_bytes());
    buf.extend_from_slice(&2u32.to_le_bytes());
    write_bin_range_list(&ranges, &mut buf)?;
    if formula_count != 0 {
        buf.extend_from_slice(&(formula_count as u32).to_le_bytes());
        for formula in [formula1.as_ref(), formula2.as_ref()].into_iter().flatten() {
            buf.extend_from_slice(&2u32.to_le_bytes());
            buf.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
            buf.extend_from_slice(&(formula.rgcb.len() as u32).to_le_bytes());
            buf.extend_from_slice(&formula.rgce);
            buf.extend_from_slice(&formula.rgcb);
        }
    }

    let mut flags = common_flags(dv);
    if formula1.is_some() {
        flags |= 0x0100_0000;
    }
    if formula2.is_some() {
        flags |= 0x0200_0000;
    }
    buf.extend_from_slice(&flags.to_le_bytes());
    write_xl_nullable_wide_string(&mut buf, dv.error_title.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.error_text.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.input_title.as_deref());
    write_xl_nullable_wide_string(&mut buf, dv.input_text.as_deref());
    Ok(buf)
}

fn compile_formula(
    binary: Option<&ParsedFormula>,
    formula: Option<&str>,
    string_list: bool,
) -> Result<Option<ParsedFormula>> {
    if let Some(binary) = binary {
        return Ok(Some(binary.clone()));
    }
    let Some(text) = formula.filter(|text| !text.is_empty()) else {
        return Ok(None);
    };
    if string_list {
        compile_string_list(text).map(Some)
    } else {
        crate::package::formula::text::Compiler::compile(text).map(Some)
    }
}

fn common_flags(dv: &Validation) -> u32 {
    let mut flags = u32::from(dv.validation_type) | (u32::from(dv.error_style) << 4);
    if dv.string_list {
        flags |= 0x80;
    }
    if dv.allow_blank {
        flags |= 0x100;
    }
    if !dv.show_dropdown {
        flags |= 0x200;
    }
    flags |= u32::from(dv.ime_mode) << 10;
    if dv.show_input_message {
        flags |= 0x0004_0000;
    }
    if dv.show_error_message {
        flags |= 0x0008_0000;
    }
    flags | (u32::from(dv.operator) << 20)
}

fn compile_string_list(text: &str) -> Result<ParsedFormula> {
    let value = text
        .strip_prefix('"')
        .and_then(|text| text.strip_suffix('"'))
        .unwrap_or(text);
    let length = value.encode_utf16().count();
    if length > 255 {
        return Err(Error::InvalidFormula(format!(
            "inline validation string list has {length} UTF-16 units; maximum is 255 (use list_formula for BrtDValList)"
        )));
    }
    let mut rgce = Vec::new();
    build_ptg_for_value(text, &mut rgce);
    Ok(ParsedFormula {
        rgce,
        rgcb: Vec::new(),
    })
}

fn validate_rule(dv: &Validation) -> Result<()> {
    if dv.validation_type > 7 {
        return Err(invalid(format!(
            "validation type {} exceeds 7",
            dv.validation_type
        )));
    }
    if dv.error_style > 2 {
        return Err(invalid(format!("error style {} exceeds 2", dv.error_style)));
    }
    if dv.ime_mode > 10 {
        return Err(invalid(format!("IME mode {} exceeds 10", dv.ime_mode)));
    }
    if !matches!(dv.validation_type, 0 | 3 | 7) && dv.operator > 7 {
        return Err(invalid(format!("operator {} exceeds 7", dv.operator)));
    }
    if dv.record_kind == RecordKind::Extension14 && dv.list_formula.is_some() {
        return Err(invalid(
            "BrtDValList overrides are only valid for classic BrtDVal rules",
        ));
    }
    if let Some(list_formula) = &dv.list_formula {
        validate_dval_list_formula(list_formula)?;
    }
    if dv.string_list
        && dv.formula1_binary.is_none()
        && dv.list_formula.is_none()
        && let Some(formula) = dv.formula1.as_deref()
    {
        let _ = compile_string_list(formula)?;
    }
    for (name, value, maximum) in [
        ("error title", dv.error_title.as_deref(), 32),
        ("error message", dv.error_text.as_deref(), 225),
        ("prompt title", dv.input_title.as_deref(), 32),
        ("prompt message", dv.input_text.as_deref(), 255),
    ] {
        if value.is_some_and(|text| text.encode_utf16().count() > maximum) {
            return Err(invalid(format!(
                "{name} exceeds {maximum} UTF-16 code units"
            )));
        }
    }
    for (name, formula) in [
        ("formula1", dv.formula1_binary.as_ref()),
        ("formula2", dv.formula2_binary.as_ref()),
    ] {
        if let Some(formula) = formula {
            if formula.rgce.is_empty() || formula.rgce.len() > 16_384 {
                return Err(Error::InvalidFormula(format!(
                    "{name} token length {} is outside 1..=16,384",
                    formula.rgce.len()
                )));
            }
            if formula.rgcb.len() > u32::MAX as usize {
                return Err(Error::InvalidFormula(format!(
                    "{name} ancillary stream is too large"
                )));
            }
        }
    }
    let first = dv.formula1_binary.is_some()
        || dv
            .formula1
            .as_deref()
            .is_some_and(|value| !value.is_empty())
        || dv
            .list_formula
            .as_deref()
            .is_some_and(|value| !value.is_empty());
    let second = dv.formula2_binary.is_some()
        || dv
            .formula2
            .as_deref()
            .is_some_and(|value| !value.is_empty());
    let required = match dv.validation_type {
        0 => (false, false),
        3 | 7 => (true, false),
        _ if dv.operator <= 1 => (true, true),
        _ => (true, false),
    };
    if (first, second) != required {
        return Err(Error::InvalidFormula(format!(
            "validation formula presence {first}/{second} does not match required {}/{}",
            required.0, required.1
        )));
    }
    let ranges = parse_range_list(&dv.cell_ranges)?;
    let maximum = if dv.record_kind == RecordKind::Classic {
        8_191
    } else {
        usize::MAX
    };
    if ranges.is_empty() || ranges.len() > maximum {
        return Err(invalid(format!(
            "validation range count {} is outside 1..={maximum}",
            ranges.len()
        )));
    }
    if ranges.iter().any(|range| {
        range.row_first > range.row_last
            || range.col_first > range.col_last
            || range.row_last >= 1_048_576
            || range.col_last >= 16_384
    }) {
        return Err(invalid("invalid validation target range"));
    }
    Ok(())
}

fn invalid(value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: "BrtDVal".to_string(),
        val: value.into(),
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
    use crate::package::formula::Context;
    use crate::raw::Writer;

    #[test]
    fn test_serialize_list_validation() {
        let dv = Validation {
            validation_type: 3, // list
            operator: 0,
            formula1: Some("Item1,Item2,Item3".to_string()),
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: true,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges: "A1:A10".to_string(),
            record_kind: Default::default(),
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

    #[test]
    fn test_serialize_whole_number_validation() {
        let dv = Validation {
            validation_type: 1, // whole number
            operator: 2,        // greater than
            formula1: Some("10".to_string()),
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: false,
            show_dropdown: false,
            show_input_message: true,
            show_error_message: false,
            error_style: 1, // warning
            ime_mode: 0,
            string_list: false,
            list_formula: None,
            input_title: Some("Input Title".to_string()),
            input_text: Some("Enter a number".to_string()),
            error_title: Some("Error".to_string()),
            error_text: Some("Must be > 10".to_string()),
            cell_ranges: "B1:B20".to_string(),
            record_kind: Default::default(),
        };

        let payload = serialize_data_validation(&dv).unwrap();
        let flags = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        assert_eq!(flags & 0x0F, 1); // type = whole number
        assert_eq!((flags >> 4) & 0x07, 1); // error style = warning
        assert_eq!(flags & 0x0100, 0); // allow blank = false
    }

    #[test]
    fn test_serialize_decimal_validation() {
        let dv = Validation {
            validation_type: 2, // decimal
            operator: 4,        // between
            formula1: Some("0.0".to_string()),
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: false,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: false,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges: "C1:C10".to_string(),
            record_kind: Default::default(),
        };

        let payload = serialize_data_validation(&dv).unwrap();
        let flags = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        assert_eq!(flags & 0x0F, 2); // type = decimal
        assert_eq!((flags >> 20) & 0x0F, 4); // operator = between
    }

    #[test]
    fn test_serialize_date_validation() {
        let dv = Validation {
            validation_type: 4, // date
            operator: 3,        // less than
            formula1: Some("2024-01-01".to_string()),
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: false,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: false,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges: "D1:D10".to_string(),
            record_kind: Default::default(),
        };

        let payload = serialize_data_validation(&dv).unwrap();
        let flags = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        assert_eq!(flags & 0x0F, 4); // type = date
    }

    #[test]
    fn test_write_data_validations_empty() {
        let mut buffer = Vec::new();
        let mut writer = Writer::new(&mut buffer);
        let validations: Vec<Validation> = vec![];

        let result = write_data_validations(
            &mut writer,
            &validations,
            Settings::default(),
            Settings::default(),
        );
        assert!(result.is_ok());
        assert!(buffer.is_empty()); // No records written for empty list
    }

    #[test]
    fn test_write_data_validations_single() {
        let mut buffer = Vec::new();
        let mut writer = Writer::new(&mut buffer);
        let validations = vec![Validation {
            validation_type: 3, // list
            operator: 0,
            formula1: Some("Yes,No".to_string()),
            formula2: None,
            formula1_binary: None,
            formula2_binary: None,
            allow_blank: true,
            show_dropdown: true,
            show_input_message: false,
            show_error_message: true,
            error_style: 0,
            ime_mode: 0,
            string_list: true,
            list_formula: None,
            input_title: None,
            input_text: None,
            error_title: None,
            error_text: None,
            cell_ranges: "A1:A10".to_string(),
            record_kind: Default::default(),
        }];

        let result = write_data_validations(
            &mut writer,
            &validations,
            Settings::default(),
            Settings::default(),
        );
        assert!(result.is_ok());
        assert!(!buffer.is_empty());
    }

    #[test]
    fn test_build_ptg_for_value_integer() {
        let mut ptg = Vec::new();
        build_ptg_for_value("42", &mut ptg);
        assert_eq!(ptg.len(), 3); // opcode(1) + u16(2)
        assert_eq!(ptg[0], 0x1E); // PtgInt opcode
        assert_eq!(u16::from_le_bytes([ptg[1], ptg[2]]), 42);
    }

    #[test]
    fn test_build_ptg_for_value_float() {
        let mut ptg = Vec::new();
        build_ptg_for_value("3.14159", &mut ptg);
        assert_eq!(ptg.len(), 9); // opcode(1) + f64(8)
        assert_eq!(ptg[0], 0x1F); // PtgNum opcode
    }

    #[test]
    fn test_build_ptg_for_value_string() {
        let mut ptg = Vec::new();
        build_ptg_for_value("Hello", &mut ptg);
        assert_eq!(ptg[0], 0x17); // PtgStr opcode
    }

    #[test]
    fn test_build_ptg_for_value_quoted_string() {
        let mut ptg = Vec::new();
        build_ptg_for_value("\"Test Value\"", &mut ptg);
        assert_eq!(ptg[0], 0x17); // PtgStr opcode
    }

    #[test]
    fn test_write_biff12_formula_empty() {
        let mut buf = Vec::new();
        write_biff12_formula(&mut buf, None, None, false).unwrap();
        assert_eq!(buf.len(), 8); // cb_ptg(4) + cb_adddata(4) both zeros
        assert_eq!(i32::from_le_bytes(buf[0..4].try_into().unwrap()), 0);
        assert_eq!(i32::from_le_bytes(buf[4..8].try_into().unwrap()), 0);
    }

    #[test]
    fn test_write_biff12_formula_with_equals() {
        let mut buf = Vec::new();
        write_biff12_formula(&mut buf, None, Some("=A1+B1"), false).unwrap();
        // Should strip leading '=' and write as string
        assert!(!buf.is_empty());
    }

    #[test]
    fn test_write_biff12_formula_number() {
        let mut buf = Vec::new();
        write_biff12_formula(&mut buf, None, Some("100"), false).unwrap();
        assert!(!buf.is_empty());
    }

    #[test]
    fn test_write_xl_wide_string() {
        let mut buf = Vec::new();
        write_xl_wide_string(&mut buf, "Test");
        assert!(!buf.is_empty());

        let char_count = u32::from_le_bytes(buf[0..4].try_into().unwrap());
        assert_eq!(char_count, 4); // "Test" has 4 chars
    }

    #[test]
    fn test_write_xl_wide_string_empty() {
        let mut buf = Vec::new();
        write_xl_wide_string(&mut buf, "");
        assert_eq!(buf.len(), 4); // Just the char count (0)
        assert_eq!(u32::from_le_bytes(buf[0..4].try_into().unwrap()), 0);
    }

    #[test]
    fn classic_payload_parses_with_binary_formula_and_nullable_strings() {
        let mut rule = Validation::new(1, "A1:A4 C1".to_string());
        rule.operator = 0;
        rule.formula1 = Some("1".to_string());
        rule.formula2 = Some("4".to_string());
        rule.ime_mode = 5;
        rule.error_title = Some(String::new());
        validate_rule(&rule).unwrap();
        let payload = serialize_data_validation(&rule).unwrap();
        let parsed = Validation::parse_classic(&payload, None, &Context::default()).unwrap();

        assert_eq!(parsed.cell_ranges, "A1:A4 C1");
        assert_eq!(parsed.formula1.as_deref(), Some("1"));
        assert_eq!(parsed.formula2.as_deref(), Some("4"));
        assert!(parsed.formula1_binary.is_some());
        assert_eq!(parsed.error_title.as_deref(), Some(""));
        assert_eq!(parsed.error_text, None);
        assert_eq!(parsed.ime_mode, 5);
    }

    #[test]
    fn extension14_payload_parses_frt_ranges_and_formulas() {
        let mut rule = Validation::new(7, "B2:B12 D2:D12".to_string());
        rule.formula1 = Some("B2>0".to_string());
        rule.record_kind = RecordKind::Extension14;
        validate_rule(&rule).unwrap();
        let payload = serialize_extension14_validation(&rule).unwrap();
        let parsed = Validation::parse_extension14(&payload, &Context::default()).unwrap();

        assert_eq!(parsed.record_kind, RecordKind::Extension14);
        assert_eq!(parsed.cell_ranges, "B2:B12 D2:D12");
        assert_eq!(parsed.formula1.as_deref(), Some("(B2>0)"));
        assert!(parsed.formula1_binary.is_some());
        assert!(parsed.formula2_binary.is_none());

        let mut malformed = payload;
        malformed[4..8].copy_from_slice(&2u32.to_le_bytes());
        assert!(Validation::parse_extension14(&malformed, &Context::default()).is_err());
    }

    #[test]
    fn dval_list_payload_preserves_the_formula_override() {
        let mut payload = Vec::new();
        write_xl_wide_string(&mut payload, "One,\"Two,Three\",Four");
        assert_eq!(
            crate::package::data_validation::parse_dval_list(&payload).unwrap(),
            "One,\"Two,Three\",Four"
        );
        payload.extend_from_slice(&[0, 0]);
        assert!(crate::package::data_validation::parse_dval_list(&payload).is_err());
    }

    #[test]
    fn long_lists_require_and_roundtrip_through_dval_list() {
        let long_list = "x".repeat(256);
        let mut inline = Validation::new(3, "A1".to_string());
        inline.formula1 = Some(long_list.clone());
        assert!(serialize_data_validation(&inline).is_err());

        let mut override_rule = Validation::new(3, "A1".to_string());
        override_rule.list_formula = Some(long_list.clone());
        validate_rule(&override_rule).unwrap();
        let payload = serialize_data_validation(&override_rule).unwrap();
        let parsed =
            Validation::parse_classic(&payload, Some(long_list.clone()), &Context::default())
                .unwrap();
        assert_eq!(parsed.formula1.as_deref(), Some(long_list.as_str()));
        assert_eq!(parsed.list_formula.as_deref(), Some(long_list.as_str()));
    }
}
