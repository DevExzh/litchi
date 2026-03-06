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
use crate::ole::xls::writer::XlsDefinedName;
use std::io::Write;

mod cells;
mod conditional_format;
mod modern_globals;
mod named_range;
mod pivot;
mod pivot_xfext;
mod sst;
mod validation;
mod workbook;
mod worksheet;

pub(crate) use modern_globals::{
    sxdbex_creation_timestamp_bytes, write_compat12, write_compress_pictures,
    write_pivot_cache_sxaddl_block, write_table_styles,
};
pub(crate) use pivot::{
    PivotCacheFieldInfo, PivotCacheSourceRow, PivotCacheStreamInfo, SxDiConfig, SxExConfig,
    SxVdConfig, SxViConfig, SxViewConfig, generate_pivot_cache_stream, write_dconref,
    write_mso_drawing_group, write_mso_drawing_sheet1, write_pivot_modern_extensions,
    write_pivot_page_mso_drawing, write_pivot_page_obj, write_sx_stream_id, write_sxdi, write_sxex,
    write_sxivd, write_sxli, write_sxpi, write_sxvd, write_sxvdex, write_sxvi, write_sxview,
    write_sxvs,
};
pub(crate) use pivot_xfext::write_pivot_xfext_block;
pub(crate) use validation::DvConfig;
pub use worksheet::AutoFilterConditionWrite;

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

/// Write NAME (Lbl) record for a defined name.
///
/// Record type: 0x0018
pub fn write_name<W: Write>(writer: &mut W, name: &XlsDefinedName, rgce: &[u8]) -> XlsResult<()> {
    named_range::write_name(writer, name, rgce)
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

pub(crate) fn has_multibyte_char(s: &str) -> bool {
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
    cfg: &DvConfig<'_>,
    ranges: &[(u32, u32, u16, u16)],
) -> XlsResult<()> {
    validation::write_dv(writer, cfg, ranges)
}

/// Write UseSelFS (Use Natural Language Formulas) record.
///
/// Record type: 0x0160, Length: 2
/// A value of 0 disables natural language formulas (modern Excel default).
pub fn write_usesel_fs<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_usesel_fs(writer)
}

pub fn write_interface_hdr<W: Write>(writer: &mut W, codepage: u16) -> XlsResult<()> {
    workbook::write_interface_hdr(writer, codepage)
}

pub fn write_mms<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_mms(writer)
}

pub fn write_interface_end<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_interface_end(writer)
}

pub fn write_write_access<W: Write>(writer: &mut W, username: &str) -> XlsResult<()> {
    workbook::write_write_access(writer, username)
}

pub fn write_window_protect<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    workbook::write_window_protect(writer, protect)
}

pub fn write_protect<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    workbook::write_protect(writer, protect)
}

pub use workbook::ExternSheetMode;

pub fn write_password<W: Write>(writer: &mut W, password_hash: u16) -> XlsResult<()> {
    workbook::write_password(writer, password_hash)
}

pub fn write_protection_rev4<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    workbook::write_protection_rev4(writer, protect)
}

pub fn write_password_rev4<W: Write>(writer: &mut W, password_hash: u16) -> XlsResult<()> {
    workbook::write_password_rev4(writer, password_hash)
}

pub fn write_backup<W: Write>(writer: &mut W, backup: bool) -> XlsResult<()> {
    workbook::write_backup(writer, backup)
}

pub fn write_hide_obj<W: Write>(writer: &mut W, mode: u16) -> XlsResult<()> {
    workbook::write_hide_obj(writer, mode)
}

pub fn write_precision<W: Write>(writer: &mut W, full_precision: bool) -> XlsResult<()> {
    workbook::write_precision(writer, full_precision)
}

pub fn write_dsf<W: Write>(writer: &mut W, has_biff5_stream: bool) -> XlsResult<()> {
    workbook::write_dsf(writer, has_biff5_stream)
}

pub fn write_tab_id<W: Write>(writer: &mut W, sheet_count: u16) -> XlsResult<()> {
    workbook::write_tab_id(writer, sheet_count)
}

pub fn write_fn_group_count<W: Write>(writer: &mut W, count: u16) -> XlsResult<()> {
    workbook::write_fn_group_count(writer, count)
}

pub fn write_refresh_all<W: Write>(writer: &mut W, refresh_all: bool) -> XlsResult<()> {
    workbook::write_refresh_all(writer, refresh_all)
}

pub fn write_book_bool<W: Write>(writer: &mut W, save_link_values: bool) -> XlsResult<()> {
    workbook::write_book_bool(writer, save_link_values)
}

pub fn write_country<W: Write>(
    writer: &mut W,
    default_country: u16,
    current_country: u16,
) -> XlsResult<()> {
    workbook::write_country(writer, default_country, current_country)
}

pub fn write_excel9_file<W: Write>(writer: &mut W) -> XlsResult<()> {
    workbook::write_excel9_file(writer)
}

pub fn write_recalc_id<W: Write>(writer: &mut W, engine_id: u32) -> XlsResult<()> {
    workbook::write_recalc_id(writer, engine_id)
}

/// Write WSBOOL record (Additional Workspace Information)
///
/// Record type: 0x0081, Length: 2
/// Writes default flags indicating a normal worksheet (not dialog sheet).
pub fn write_wsbool<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_wsbool(writer)
}

pub fn write_pivot_sheet_preamble<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_pivot_sheet_preamble(writer)
}

pub fn write_pivot_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
) -> XlsResult<()> {
    worksheet::write_pivot_colinfo(writer, first_col, last_col, col_width)
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

pub fn write_pivot_window2<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_pivot_window2(writer)
}

pub fn write_plv<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_plv(writer)
}

pub fn write_selection<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_selection(writer)
}

pub fn write_phonetic_pr<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_phonetic_pr(writer)
}

pub fn write_sheet_ext<W: Write>(writer: &mut W) -> XlsResult<()> {
    worksheet::write_sheet_ext(writer)
}

/// Write PANE record (freeze panes configuration)
///
/// Record type: 0x0041, Length: 10
pub fn write_pane<W: Write>(writer: &mut W, freeze_rows: u32, freeze_cols: u16) -> XlsResult<()> {
    worksheet::write_pane(writer, freeze_rows, freeze_cols)
}

pub fn write_autofilterinfo<W: Write>(writer: &mut W, c_entries: u16) -> XlsResult<()> {
    worksheet::write_autofilterinfo(writer, c_entries)
}

pub fn write_sheet_protection<W: Write>(
    writer: &mut W,
    protect_objects: bool,
    protect_scenarios: bool,
    password_hash: Option<u16>,
) -> XlsResult<()> {
    worksheet::write_sheet_protection(writer, protect_objects, protect_scenarios, password_hash)
}

/// Write HLINK (hyperlink) record for a cell or cell range.
///
/// Record type: 0x01B8
pub fn write_hyperlink<W: Write>(
    writer: &mut W,
    row1: u32,
    row2: u32,
    col1: u16,
    col2: u16,
    url: &str,
) -> XlsResult<()> {
    worksheet::write_hyperlink(writer, row1, row2, col1, col2, url)
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

/// Write internal SUPBOOK record used for 3D references within this
/// workbook.
pub fn write_supbook_internal<W: Write>(writer: &mut W, sheet_count: u16) -> XlsResult<()> {
    workbook::write_supbook_internal(writer, sheet_count)
}

/// Write EXTERNSHEET record for internal workbook references.
pub fn write_externsheet_internal<W: Write>(
    writer: &mut W,
    sheet_count: u16,
    mode: ExternSheetMode,
) -> XlsResult<()> {
    workbook::write_externsheet_internal(writer, sheet_count, mode)
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

/// Write COLINFO record (column formatting and width).
///
/// Record type: 0x007D
pub fn write_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
    hidden: bool,
) -> XlsResult<()> {
    worksheet::write_colinfo(writer, first_col, last_col, col_width, hidden)
}

/// Write DEFCOLWIDTH record.
///
/// Record type: 0x0055
pub fn write_def_col_width<W: Write>(writer: &mut W, width_chars: u16) -> XlsResult<()> {
    worksheet::write_def_col_width(writer, width_chars)
}

/// Write INDEX record.
///
/// Record type: 0x020B
pub fn write_index<W: Write>(
    writer: &mut W,
    first_row: u32,
    last_row_plus1: u32,
    def_col_width_pos: u32,
    dbcell_positions: &[u32],
) -> XlsResult<()> {
    worksheet::write_index(
        writer,
        first_row,
        last_row_plus1,
        def_col_width_pos,
        dbcell_positions,
    )
}

/// Write DBCELL record.
///
/// Record type: 0x00D7
pub fn write_dbcell<W: Write>(
    writer: &mut W,
    row_offset: u32,
    cell_offsets: &[u16],
) -> XlsResult<()> {
    worksheet::write_dbcell(writer, row_offset, cell_offsets)
}

/// Write ROW record (row metrics including height and hidden flag).
///
/// Record type: 0x0208
pub fn write_row<W: Write>(
    writer: &mut W,
    row_index: u32,
    first_col: u16,
    last_col_plus1: u16,
    height: u16,
    hidden: bool,
) -> XlsResult<()> {
    worksheet::write_row(writer, row_index, first_col, last_col_plus1, height, hidden)
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

pub fn write_mulrk<W: Write>(
    writer: &mut W,
    row: u32,
    first_col: u16,
    values: &[(u16, f64)],
) -> XlsResult<()> {
    cells::write_mulrk(writer, row, first_col, values)
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

/// Write an AUTOFILTER record (0x009E) for a single column filter condition.
#[allow(clippy::too_many_arguments)]
pub fn write_autofilter<W: Write>(
    writer: &mut W,
    column_index: u16,
    join_or: bool,
    is_simple: bool,
    is_top10: bool,
    hide_arrow: bool,
    cond1: &AutoFilterConditionWrite,
    cond2: &AutoFilterConditionWrite,
) -> XlsResult<()> {
    worksheet::write_autofilter(
        writer,
        column_index,
        join_or,
        is_simple,
        is_top10,
        hide_arrow,
        cond1,
        cond2,
    )
}

/// Write a SORT record (0x0090).
pub fn write_sort<W: Write>(
    writer: &mut W,
    case_sensitive: bool,
    sort_by_columns: bool,
    keys: &[(u16, bool)],
) -> XlsResult<()> {
    worksheet::write_sort(writer, case_sensitive, sort_by_columns, keys)
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
