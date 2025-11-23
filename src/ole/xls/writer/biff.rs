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

use super::super::XlsResult;
use std::io::Write;

mod cells;
mod conditional_format;
mod sst;
mod validation;
mod workbook;
mod worksheet;

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

/// Write FORMAT record (number format string)
///
/// Record type: 0x041E
pub fn write_format_record<W: Write>(
    writer: &mut W,
    index_code: u16,
    format_str: &str,
) -> XlsResult<()> {
    workbook::write_format_record(writer, index_code, format_str)
}

fn has_multibyte_char(s: &str) -> bool {
    s.chars().any(|c| c as u32 > 0xFF)
}

pub(crate) fn unicode_string_size(value: &str) -> u16 {
    let char_count = value.chars().count() as u16;
    if has_multibyte_char(value) {
        2u16 + 1u16 + char_count.saturating_mul(2)
    } else {
        2u16 + 1u16 + char_count
    }
}

pub(crate) fn write_unicode_string_biff8<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    let char_count: u16 = value.chars().count() as u16;
    writer.write_all(&char_count.to_le_bytes())?;

    let is_16bit = has_multibyte_char(value);
    writer.write_all(&[if is_16bit { 0x01 } else { 0x00 }])?;

    if is_16bit {
        for code_unit in value.encode_utf16() {
            writer.write_all(&code_unit.to_le_bytes())?;
        }
    } else {
        writer.write_all(value.as_bytes())?;
    }

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
    workbook::write_builtin_styles(writer)
}

pub fn write_dval<W: Write>(writer: &mut W, dv_count: u32) -> XlsResult<()> {
    validation::write_dval(writer, dv_count)
}

pub fn write_dv<W: Write>(
    writer: &mut W,
    data_type: u8,
    operator: u8,
    error_style: u8,
    empty_cell_allowed: bool,
    suppress_dropdown_arrow: bool,
    is_explicit_list_formula: bool,
    show_prompt_on_cell_selected: bool,
    prompt_title: Option<&str>,
    prompt_text: Option<&str>,
    show_error_on_invalid_value: bool,
    error_title: Option<&str>,
    error_text: Option<&str>,
    formula1: Option<&[u8]>,
    formula2: Option<&[u8]>,
    ranges: &[(u32, u32, u16, u16)],
) -> XlsResult<()> {
    validation::write_dv(
        writer,
        data_type,
        operator,
        error_style,
        empty_cell_allowed,
        suppress_dropdown_arrow,
        is_explicit_list_formula,
        show_prompt_on_cell_selected,
        prompt_title,
        prompt_text,
        show_error_on_invalid_value,
        error_title,
        error_text,
        formula1,
        formula2,
        ranges,
    )
}

/// Write UseSelFS (Use Natural Language Formulas) record.
///
/// Record type: 0x0160, Length: 2
/// A value of 0 disables natural language formulas (modern Excel default).
pub fn write_usesel_fs<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_usesel_fs(writer)
}

/// Write WSBOOL record (Additional Workspace Information)
///
/// Record type: 0x0081, Length: 2
/// Writes default flags indicating a normal worksheet (not dialog sheet).
pub fn write_wsbool<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_wsbool(writer)
}

/// Write WINDOW2 record (Worksheet view settings)
///
/// Record type: 0x023E, Length: 18 (worksheet and macro sheet)
///
/// The `has_freeze_panes` flag controls whether the FREEZE_PANES and
/// FREEZE_PANES_NO_SPLIT bits are set in the options field.
pub fn write_window2<W: Write>(writer: &mut W, has_freeze_panes: bool) -> XlsResult<()> {
    worksheet::write_window2(writer, has_freeze_panes)
}

/// Write PANE record (freeze panes configuration)
///
/// Record type: 0x0041, Length: 10
pub fn write_pane<W: Write>(writer: &mut W, freeze_rows: u32, freeze_cols: u16) -> XlsResult<()> {
    worksheet::write_pane(writer, freeze_rows, freeze_cols)
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
    workbook::write_bof(writer, substream_type)
}

/// Write EOF (End of File) record
///
/// Record type: 0x000A
pub fn write_eof<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_eof(writer)
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
    workbook::write_codepage(writer, codepage)
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
    workbook::write_date1904(writer, is_1904)
}

/// Write WINDOW1 record (workbook window properties)
///
/// Record type: 0x003D
pub fn write_window1<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_window1(writer)
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
    workbook::write_boundsheet(writer, position, name)
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
    worksheet::write_dimensions(writer, first_row, last_row, first_col, last_col)
}

pub fn write_mergedcells<W, I>(writer: &mut W, ranges: I) -> XlsResult<()>
where
    W: Write,
    I: IntoIterator<Item = (u32, u32, u16, u16)>,
{
    worksheet::write_mergedcells(writer, ranges)
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
    cells::write_number(writer, row, col, xf_index, value)
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
    cells::write_labelsst(writer, row, col, xf_index, sst_index)
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
    cells::write_boolerr(writer, row, col, xf_index, value)
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
    sst::write_sst(writer, strings, cst_total)
}

pub fn write_cfheader<W: Write>(
    writer: &mut W,
    ranges: &[(u32, u32, u16, u16)],
    num_rules: u16,
) -> XlsResult<()> {
    conditional_format::write_cfheader(writer, ranges, num_rules)
}

pub fn write_cfrule<W: Write>(
    writer: &mut W,
    condition_type: u8,
    comparison_op: u8,
    formula1: &[u8],
    formula2: &[u8],
    pattern: Option<(u16, u16, u16)>,
) -> XlsResult<()> {
    conditional_format::write_cfrule(
        writer,
        condition_type,
        comparison_op,
        formula1,
        formula2,
        pattern,
    )
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
