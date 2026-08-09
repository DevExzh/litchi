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

use super::super::{Error, Result};
use crate::writer::DefinedName;
use litchi_biff::{Encoder, Kind};
use std::io::Write;

mod cells;
mod comment;
mod conditional_format;
mod drawing;
mod drawing_group;
mod list_object;
mod modern_globals;
mod named_range;
mod pivot;
mod pivot_xfext;
mod scenario;
mod sst;
mod validation;
mod vba;
mod workbook;
mod worksheet;

pub(crate) use cells::write_table;
pub(crate) use cells::{write_formula, write_formula_with_metadata};
pub(crate) use comment::CommentConfig;
pub(crate) use conditional_format::Cf12Config;
pub(crate) use drawing::{PrimitiveShapeConfig, write_worksheet_drawing};
pub(crate) use drawing_group::GroupShapeConfig;
pub(crate) use list_object::write_list_objects;

pub(crate) use modern_globals::{
    sxdbex_creation_timestamp_bytes, write_compat12, write_compress_pictures,
    write_custom_table_styles, write_differential_formats, write_pivot_cache_sxaddl_block,
    write_table_styles,
};
pub(crate) use pivot::{
    PivotCacheFieldInfo, PivotCacheSourceRow, PivotCacheStreamInfo, SxDiConfig, SxExConfig,
    SxVdConfig, SxViConfig, SxViewConfig, generate_pivot_cache_stream, write_dconref,
    write_mso_drawing_group, write_pivot_modern_extensions, write_sx_stream_id, write_sxdi,
    write_sxex, write_sxivd, write_sxli, write_sxpi, write_sxvd, write_sxvdex, write_sxvi,
    write_sxview, write_sxvs,
};
pub(crate) use pivot_xfext::write_pivot_xfext_block;
pub(crate) use validation::{DvConfig, DvalConfig};
pub use worksheet::AutoFilterConditionWrite;

fn map_frame_error(error: litchi_biff::Error) -> Error {
    match error {
        litchi_biff::Error::LimitExceeded {
            resource: litchi_biff::Resource::RecordBytes,
            observed,
            maximum,
        } => Error::InvalidLength {
            expected: usize::try_from(maximum).unwrap_or(usize::MAX),
            found: usize::try_from(observed).unwrap_or(usize::MAX),
        },
        litchi_biff::Error::Allocation { .. } => {
            Error::Allocation("encoding a complete BIFF record")
        },
        error => Error::InvalidData(format!("BIFF frame encoding failed: {error}")),
    }
}

fn record_kind(record_type: u16) -> Result<Kind> {
    Kind::try_from(u64::from(record_type)).map_err(map_frame_error)
}

/// Write one complete BIFF record through the neutral bounded frame encoder.
///
/// Callers should use this only after materializing the semantic payload. The
/// streaming writers below intentionally keep [`write_record_header`] so that
/// large or continuation-sensitive payloads do not acquire a per-record
/// staging allocation.
pub(crate) fn write_record<W: Write>(
    writer: &mut W,
    record_type: u16,
    payload: &[u8],
) -> Result<()> {
    let mut encoder = Encoder::new();
    encoder
        .push(record_kind(record_type)?, payload)
        .map_err(map_frame_error)?;
    writer.write_all(encoder.as_bytes())?;
    Ok(())
}

/// Write a BIFF record header for a payload that will be streamed directly.
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
) -> Result<()> {
    if usize::from(data_len) > litchi_biff::MAX_RECORD_BYTES {
        return Err(Error::InvalidLength {
            expected: litchi_biff::MAX_RECORD_BYTES,
            found: usize::from(data_len),
        });
    }
    let kind = record_kind(record_type)?;
    writer.write_all(&kind.get().to_le_bytes())?;
    writer.write_all(&data_len.to_le_bytes())?;
    Ok(())
}

pub(crate) fn write_force_full_calculation<W: Write>(writer: &mut W, force: bool) -> Result<()> {
    workbook::write_force_full_calculation(writer, force)
}

pub(crate) fn write_mtr_settings<W: Write>(
    writer: &mut W,
    settings: crate::MultithreadedCalculation,
) -> Result<()> {
    workbook::write_mtr_settings(writer, settings)
}

pub(crate) fn write_uncalced<W: Write>(writer: &mut W) -> Result<()> {
    worksheet::write_uncalced(writer)
}

pub(crate) fn write_calculation_settings<W: Write>(
    writer: &mut W,
    settings: &crate::writer::core::CalculationSettings,
) -> Result<()> {
    worksheet::write_calculation_settings(writer, settings)
}

pub(crate) fn write_scenario_manager<W: Write>(
    writer: &mut W,
    manager: &crate::scenario::ScenarioManager,
) -> Result<()> {
    scenario::write_scenario_manager(writer, manager)
}

pub(crate) fn write_ob_proj<W: Write>(writer: &mut W) -> Result<()> {
    vba::write_ob_proj(writer)
}
pub(crate) fn write_ob_no_macros<W: Write>(writer: &mut W) -> Result<()> {
    vba::write_ob_no_macros(writer)
}
pub(crate) fn write_code_name<W: Write>(writer: &mut W, value: &str) -> Result<()> {
    vba::write_code_name(writer, value)
}

/// Write NAME (Lbl) record for a defined name.
///
/// Record type: 0x0018
pub(crate) fn write_name<W: Write>(writer: &mut W, name: &DefinedName, rgce: &[u8]) -> Result<()> {
    named_range::write_name(writer, name, rgce)
}

pub(crate) fn write_name_comment<W: Write>(
    writer: &mut W,
    name: &str,
    comment: &str,
) -> Result<()> {
    named_range::write_name_comment(writer, name, comment)
}

pub(crate) fn write_name_function_group<W: Write>(
    writer: &mut W,
    value: &crate::NameFnGrp12,
) -> Result<()> {
    named_range::write_name_function_group(writer, value)
}
pub(crate) fn write_name_publish<W: Write>(
    writer: &mut W,
    value: &crate::NamePublish,
) -> Result<()> {
    named_range::write_name_publish(writer, value)
}

pub(crate) fn write_defined_name_record<W: Write>(
    writer: &mut W,
    name: &crate::writer::DefinedNameRecordOptions,
) -> Result<()> {
    named_range::write_defined_name_record(writer, name)
}

/// Write FORMAT record (number format string)
///
/// Record type: 0x041E
pub(crate) fn write_format_record<W: Write>(
    writer: &mut W,
    index_code: u16,
    format_str: &str,
) -> Result<()> {
    workbook::write_format_record(writer, index_code, format_str)
}

pub(crate) fn has_multibyte_char(s: &str) -> bool {
    s.chars().any(|c| c as u32 > 0xFF)
}

pub(crate) fn unicode_string_size(value: &str) -> u16 {
    let char_count = crate::utils::truncate_usize_to_u16(value.chars().count());
    if has_multibyte_char(value) {
        2u16 + 1u16 + char_count.saturating_mul(2)
    } else {
        2u16 + 1u16 + char_count
    }
}

pub(crate) fn write_unicode_string_biff8<W: Write>(writer: &mut W, value: &str) -> Result<()> {
    let char_count: u16 = crate::utils::truncate_usize_to_u16(value.chars().count());
    writer.write_all(&char_count.to_le_bytes())?;

    let is_16bit = has_multibyte_char(value);
    writer.write_all(&[u8::from(is_16bit)])?;

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
/// Mapping (`xf_index`, `builtin_style_id)`:
/// - (0x0010, 3)  => Comma
/// - (0x0011, 6)  => Comma [0 decimals]
/// - (0x0012, 4)  => Currency
/// - (0x0013, 7)  => Currency [0 decimals]
/// - (0x0000, 0)  => Normal
/// - (0x0014, 5)  => Percent
pub(crate) fn write_builtin_styles<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_builtin_styles(writer)
}

pub(crate) fn write_dval<W: Write>(writer: &mut W, cfg: DvalConfig) -> Result<()> {
    validation::write_dval(writer, cfg)
}

pub(crate) fn write_dv<W: Write>(
    writer: &mut W,
    cfg: &DvConfig<'_>,
    ranges: &[(u16, u16, u8, u8)],
) -> Result<()> {
    validation::write_dv(writer, cfg, ranges)
}

/// Write `UseSelFS` (Use Natural Language Formulas) record.
///
/// Record type: 0x0160, Length: 2
/// A value of 0 disables natural language formulas (modern Excel default).
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")] // Compatibility wrapper for the former fixed-value writer API.
pub(crate) fn write_usesel_fs<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_usesel_fs(writer)
}

pub(crate) fn write_interface_hdr<W: Write>(writer: &mut W, codepage: u16) -> Result<()> {
    workbook::write_interface_hdr(writer, codepage)
}

pub(crate) fn write_mms<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_mms(writer)
}

pub(crate) fn write_interface_end<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_interface_end(writer)
}

pub(crate) fn write_write_access<W: Write>(writer: &mut W, username: &str) -> Result<()> {
    workbook::write_write_access(writer, username)
}

pub(crate) fn write_window_protect<W: Write>(writer: &mut W, protect: bool) -> Result<()> {
    workbook::write_window_protect(writer, protect)
}

pub(crate) fn write_protect<W: Write>(writer: &mut W, protect: bool) -> Result<()> {
    workbook::write_protect(writer, protect)
}

pub(crate) use workbook::ExternSheetMode;

pub(crate) fn write_external_link_table<W: Write>(
    writer: &mut W,
    internal: Option<(u16, ExternSheetMode)>,
    external: &[crate::writer::core::ExternalWorkbookOptions],
    external_names: &[Vec<crate::writer::core::ExternalDefinedNameOptions>],
    add_in_functions: &[crate::writer::core::AddInFunctionOptions],
    dde_or_ole_links: &[crate::writer::core::DdeOrOleLinkOptions],
) -> Result<()> {
    workbook::write_external_link_table(
        writer,
        internal,
        external,
        external_names,
        add_in_functions,
        dde_or_ole_links,
    )
}

pub(crate) fn write_password<W: Write>(writer: &mut W, password_hash: u16) -> Result<()> {
    workbook::write_password(writer, password_hash)
}

pub(crate) fn write_protection_rev4<W: Write>(writer: &mut W, protect: bool) -> Result<()> {
    workbook::write_protection_rev4(writer, protect)
}

pub(crate) fn write_password_rev4<W: Write>(writer: &mut W, password_hash: u16) -> Result<()> {
    workbook::write_password_rev4(writer, password_hash)
}

pub(crate) fn write_write_protect<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_write_protect(writer)
}

pub(crate) fn write_file_sharing<W: Write>(
    writer: &mut W,
    read_only_recommended: bool,
    password_hash: Option<u16>,
    user_name: &str,
) -> Result<()> {
    workbook::write_file_sharing(writer, read_only_recommended, password_hash, user_name)
}

pub(crate) fn write_backup<W: Write>(writer: &mut W, backup: bool) -> Result<()> {
    workbook::write_backup(writer, backup)
}

pub(crate) fn write_hide_obj<W: Write>(writer: &mut W, mode: u16) -> Result<()> {
    workbook::write_hide_obj(writer, mode)
}

pub(crate) fn write_precision<W: Write>(writer: &mut W, full_precision: bool) -> Result<()> {
    workbook::write_precision(writer, full_precision)
}

pub(crate) fn write_dsf<W: Write>(writer: &mut W, has_biff5_stream: bool) -> Result<()> {
    workbook::write_dsf(writer, has_biff5_stream)
}

pub(crate) fn write_tab_id<W: Write>(writer: &mut W, sheet_count: u16) -> Result<()> {
    workbook::write_tab_id(writer, sheet_count)
}

pub(crate) fn write_function_groups<W: Write>(
    writer: &mut W,
    options: &crate::writer::core::FunctionGroupOptions,
) -> Result<()> {
    workbook::write_function_groups(writer, options)
}

pub(crate) fn write_refresh_all<W: Write>(writer: &mut W, refresh_all: bool) -> Result<()> {
    workbook::write_refresh_all(writer, refresh_all)
}

#[allow(dead_code, reason = "retained as a BIFF compatibility building block")] // Compatibility wrapper retained for existing internal callers.
pub(crate) fn write_book_bool<W: Write>(writer: &mut W, save_link_values: bool) -> Result<()> {
    workbook::write_book_bool(writer, save_link_values)
}
pub(crate) fn write_book_bool_raw<W: Write>(writer: &mut W, bits: u16) -> Result<()> {
    workbook::write_book_bool_raw(writer, bits)
}
pub(crate) fn write_usesel_fs_value<W: Write>(writer: &mut W, enabled: bool) -> Result<()> {
    workbook::write_usesel_fs_value(writer, enabled)
}
pub(crate) fn write_template<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_template(writer)
}

pub(crate) fn write_country<W: Write>(
    writer: &mut W,
    default_country: u16,
    current_country: u16,
) -> Result<()> {
    workbook::write_country(writer, default_country, current_country)
}

pub(crate) fn write_book_ext<W: Write>(writer: &mut W, value: &crate::BookExt) -> Result<()> {
    workbook::write_book_ext(writer, value)
}

pub(crate) fn write_real_time_data<W: Write>(
    writer: &mut W,
    value: &crate::real_time_data::Record,
) -> Result<()> {
    workbook::write_real_time_data(writer, value)
}

pub(crate) fn write_web_pub<W: Write>(writer: &mut W, value: &crate::WebPub) -> Result<()> {
    workbook::write_web_pub(writer, value)
}

pub(crate) fn write_style_ext<W: Write>(writer: &mut W, value: &crate::StyleExt) -> Result<()> {
    workbook::write_style_ext(writer, value)
}

pub(crate) fn write_theme<W: Write>(writer: &mut W, value: &crate::Theme) -> Result<()> {
    workbook::write_theme(writer, value)
}

pub(crate) fn write_mdx_metadata<W: Write>(
    writer: &mut W,
    value: &crate::MdxMetadata,
) -> Result<()> {
    workbook::write_mdx_metadata(writer, value)
}

pub(crate) fn write_xfcrc<W: Write>(writer: &mut W, xf_count: u16) -> Result<()> {
    workbook::write_xfcrc(writer, xf_count)
}

pub(crate) fn write_xf_ext<W: Write>(writer: &mut W, value: &crate::XfExt) -> Result<()> {
    workbook::write_xf_ext(writer, value)
}

pub(crate) fn write_excel9_file<W: Write>(writer: &mut W) -> Result<()> {
    workbook::write_excel9_file(writer)
}

pub(crate) fn write_recalc_id<W: Write>(writer: &mut W, engine_id: u32) -> Result<()> {
    workbook::write_recalc_id(writer, engine_id)
}

pub(crate) fn write_worksheet_layout<W: Write>(
    writer: &mut W,
    options: &crate::writer::core::WorksheetLayoutOptions,
) -> Result<()> {
    worksheet::write_worksheet_layout(writer, options)
}

pub(crate) fn write_pivot_sheet_preamble<W: Write>(
    writer: &mut W,
    options: &crate::writer::core::WorksheetLayoutOptions,
) -> Result<()> {
    worksheet::write_pivot_sheet_preamble(writer, options)
}

pub(crate) fn write_pivot_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
) -> Result<()> {
    worksheet::write_pivot_colinfo(writer, first_col, last_col, col_width)
}

/// Write WINDOW2 record (Worksheet view settings)
///
/// Record type: 0x023E, Length: 18 (worksheet and macro sheet)
///
/// The `has_freeze_panes` flag controls whether the `FREEZE_PANES` and
/// `FREEZE_PANES_NO_SPLIT` bits are set in the options field.
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub(crate) fn write_window2<W: Write>(writer: &mut W, has_freeze_panes: bool) -> Result<()> {
    worksheet::write_window2(writer, has_freeze_panes)
}

pub(crate) fn write_scl<W: Write>(writer: &mut W, numerator: u16, denominator: u16) -> Result<()> {
    worksheet::write_scl(writer, numerator, denominator)
}

pub(crate) fn write_window2_options<W: Write>(
    writer: &mut W,
    options: &crate::writer::view::View,
) -> Result<()> {
    worksheet::write_window2_options(writer, options)
}

pub(crate) fn write_pane_options<W: Write>(
    writer: &mut W,
    pane: &crate::writer::view::Pane,
) -> Result<()> {
    worksheet::write_pane_options(writer, pane)
}

pub(crate) fn write_selection_options<W: Write>(
    writer: &mut W,
    selection: &crate::writer::view::Selection,
) -> Result<()> {
    worksheet::write_selection_options(writer, selection)
}

#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub(crate) fn write_default_selection<W: Write>(
    writer: &mut W,
    freeze_rows: u16,
    freeze_cols: u16,
) -> Result<()> {
    worksheet::write_default_selection(writer, freeze_rows, freeze_cols)
}

pub(crate) fn write_pivot_window2<W: Write>(writer: &mut W, selected: bool) -> Result<()> {
    worksheet::write_pivot_window2(writer, selected)
}

pub(crate) fn write_plv<W: Write>(writer: &mut W) -> Result<()> {
    worksheet::write_plv(writer)
}

pub(crate) fn write_selection<W: Write>(writer: &mut W) -> Result<()> {
    worksheet::write_selection(writer)
}

pub(crate) fn write_phonetic_pr<W: Write>(writer: &mut W) -> Result<()> {
    worksheet::write_phonetic_pr(writer)
}

pub(crate) fn write_sheet_ext<W: Write>(writer: &mut W) -> Result<()> {
    worksheet::write_sheet_ext(writer)
}

pub(crate) fn write_sheet_ext_tab_color<W: Write>(writer: &mut W, tab_color: u8) -> Result<()> {
    worksheet::write_sheet_ext_tab_color(writer, tab_color)
}

pub(crate) fn write_phonetic_info<W: Write>(
    writer: &mut W,
    value: &crate::PhoneticInfo,
) -> Result<()> {
    worksheet::write_phonetic_info(writer, value)
}

/// Write PANE record (freeze panes configuration)
///
/// Record type: 0x0041, Length: 10
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub(crate) fn write_pane<W: Write>(
    writer: &mut W,
    freeze_rows: u32,
    freeze_cols: u16,
) -> Result<()> {
    worksheet::write_pane(writer, freeze_rows, freeze_cols)
}

pub(crate) fn write_autofilterinfo<W: Write>(writer: &mut W, c_entries: u16) -> Result<()> {
    worksheet::write_autofilterinfo(writer, c_entries)
}

pub(crate) fn write_sheet_protection<W: Write>(
    writer: &mut W,
    protect_objects: bool,
    protect_scenarios: bool,
    password_hash: Option<u16>,
) -> Result<()> {
    worksheet::write_sheet_protection(writer, protect_objects, protect_scenarios, password_hash)
}

pub(crate) fn write_page_settings<W: Write>(
    writer: &mut W,
    options: &crate::writer::core::PageSetupOptions,
    horizontal_breaks: &[(u16, u16, u16)],
    vertical_breaks: &[(u16, u16, u16)],
) -> Result<()> {
    worksheet::write_page_settings(writer, options, horizontal_breaks, vertical_breaks)
}

pub(crate) fn write_consolidation<W: Write>(
    writer: &mut W,
    consolidation: &crate::Consolidation,
) -> Result<()> {
    worksheet::write_consolidation(writer, consolidation)
}

/// Write HLINK (hyperlink) record for a cell or cell range.
///
/// Record type: 0x01B8
pub(crate) fn write_hyperlink<W: Write>(
    writer: &mut W,
    row1: u32,
    row2: u32,
    col1: u16,
    col2: u16,
    url: &str,
) -> Result<()> {
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
pub(crate) fn write_bof<W: Write>(writer: &mut W, substream_type: u16) -> Result<()> {
    workbook::write_bof(writer, substream_type)
}

/// Write EOF (End of File) record
///
/// Record type: 0x000A
pub(crate) fn write_eof<W: Write>(writer: &mut W) -> Result<()> {
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
pub(crate) fn write_codepage<W: Write>(writer: &mut W, codepage: u16) -> Result<()> {
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
pub(crate) fn write_date1904<W: Write>(writer: &mut W, is_1904: bool) -> Result<()> {
    workbook::write_date1904(writer, is_1904)
}

/// Write WINDOW1 record (workbook window properties)
///
/// Record type: 0x003D
pub(crate) fn write_window1<W: Write>(
    writer: &mut W,
    options: &crate::writer::core::WorkbookWindowOptions,
    sheet_count: usize,
) -> Result<()> {
    workbook::write_window1(writer, options, sheet_count)
}

/// Write internal SUPBOOK record used for 3D references within this
/// workbook.
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
/// The sheet name is encoded as `ShortXLUnicodeString` per BIFF8: 1-byte character count,
/// 1-byte flags (0x00 = compressed 8-bit, 0x01 = uncompressed UTF-16LE), followed by characters.
pub(crate) fn write_boundsheet<W: Write>(writer: &mut W, position: u32, name: &str) -> Result<()> {
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
pub(crate) fn write_dimensions<W: Write>(
    writer: &mut W,
    first_row: u32,
    last_row: u32,
    first_col: u16,
    last_col: u16,
) -> Result<()> {
    worksheet::write_dimensions(writer, first_row, last_row, first_col, last_col)
}

pub(crate) fn write_mergedcells<W, I>(writer: &mut W, ranges: I) -> Result<()>
where
    W: Write,
    I: IntoIterator<Item = (u16, u16, u8, u8)>,
{
    worksheet::write_mergedcells(writer, ranges)
}

/// Write COLINFO record (column formatting and width).
///
/// Record type: 0x007D
pub(crate) fn write_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
    hidden: bool,
) -> Result<()> {
    worksheet::write_colinfo(writer, first_col, last_col, col_width, hidden)
}

/// Write DEFCOLWIDTH record.
///
/// Record type: 0x0055
pub(crate) fn write_def_col_width<W: Write>(writer: &mut W, width_chars: u16) -> Result<()> {
    worksheet::write_def_col_width(writer, width_chars)
}

/// Write INDEX record.
///
/// Record type: 0x020B
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub(crate) fn write_index<W: Write>(
    writer: &mut W,
    first_row: u32,
    last_row_plus1: u32,
    def_col_width_pos: u32,
    dbcell_positions: &[u32],
) -> Result<()> {
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
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub(crate) fn write_dbcell<W: Write>(
    writer: &mut W,
    row_offset: u32,
    cell_offsets: &[u16],
) -> Result<()> {
    worksheet::write_dbcell(writer, row_offset, cell_offsets)
}

/// Write ROW record (row metrics including height and hidden flag).
///
/// Record type: 0x0208
pub(crate) fn write_row<W: Write>(
    writer: &mut W,
    row_index: u32,
    first_col: u16,
    last_col_plus1: u16,
    height: u16,
    hidden: bool,
) -> Result<()> {
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
pub(crate) fn write_number<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: f64,
) -> Result<()> {
    cells::write_number(writer, row, col, xf_index, value)
}

pub(crate) fn write_mulrk<W: Write>(
    writer: &mut W,
    row: u32,
    first_col: u16,
    values: &[(u16, f64)],
) -> Result<()> {
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
pub(crate) fn write_labelsst<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    sst_index: u32,
) -> Result<()> {
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
pub(crate) fn write_boolerr<W: Write>(
    writer: &mut W,
    row: u32,
    col: u16,
    xf_index: u16,
    value: bool,
) -> Result<()> {
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
/// based on Apache POI's `SSTSerializer`.
pub(crate) fn write_sst<W: Write>(
    writer: &mut W,
    strings: &[String],
    cst_total: u32,
) -> Result<()> {
    sst::write_sst(writer, strings, cst_total)
}

/// Write an AUTOFILTER record (0x009E) for a single column filter condition.
#[allow(
    clippy::too_many_arguments,
    reason = "arguments map positionally to BIFF fields"
)]
pub(crate) fn write_autofilter<W: Write>(
    writer: &mut W,
    column_index: u16,
    join_or: bool,
    is_simple: bool,
    is_top10: bool,
    hide_arrow: bool,
    cond1: &AutoFilterConditionWrite,
    cond2: &AutoFilterConditionWrite,
) -> Result<()> {
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
pub(crate) fn write_sort<W: Write>(
    writer: &mut W,
    case_sensitive: bool,
    sort_by_columns: bool,
    keys: &[(u16, bool)],
) -> Result<()> {
    worksheet::write_sort(writer, case_sensitive, sort_by_columns, keys)
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn write_cfheader<W: Write>(
    writer: &mut W,
    ranges: &[(u32, u32, u16, u16)],
    num_rules: u16,
) -> Result<()> {
    conditional_format::write_cfheader(writer, ranges, num_rules)
}

pub(crate) fn write_cfrule<W: Write>(
    writer: &mut W,
    condition_type: u8,
    comparison_op: u8,
    formula1: &[u8],
    formula2: &[u8],
    pattern: Option<(u16, u16, u16)>,
) -> Result<()> {
    conditional_format::write_cfrule(
        writer,
        condition_type,
        comparison_op,
        formula1,
        formula2,
        pattern,
    )
}

pub(crate) fn write_cfheader_with_identifier<W: Write>(
    writer: &mut W,
    ranges: &[(u32, u32, u16, u16)],
    num_rules: u16,
    identifier: u16,
) -> Result<()> {
    conditional_format::write_cfheader_with_identifier(writer, ranges, num_rules, identifier)
}

pub(crate) fn write_condfmt12<W: Write>(
    writer: &mut W,
    ranges: &[(u32, u32, u16, u16)],
    num_rules: u16,
    identifier: u16,
) -> Result<()> {
    conditional_format::write_condfmt12(writer, ranges, num_rules, identifier)
}

pub(crate) fn write_cf12<W: Write>(writer: &mut W, config: &Cf12Config<'_>) -> Result<()> {
    conditional_format::write_cf12(writer, config)
}

/// Write an exact `CFEx` marker which must immediately precede its associated `CF12` record.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn write_cfex12_marker<W: Write>(
    writer: &mut W,
    identifier: u16,
    enclosing: (u16, u16, u16, u16),
) -> Result<()> {
    conditional_format::write_cfex12_marker(writer, identifier, enclosing)
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
    fn test_write_bof_worksheet() {
        let mut buf = Vec::new();
        write_bof(&mut buf, 0x0010).unwrap(); // Worksheet substream

        assert_eq!(&buf[0..2], &[0x09, 0x08]); // Record type 0x0809
        // Record header is 4 bytes, then BIFF8 BOF structure follows
        // Substream type is at offset 8 within the record data
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

    #[test]
    fn test_write_number_different_values() {
        let mut buf = Vec::new();
        write_number(&mut buf, 5, 3, 0x0010, -123.456).unwrap();

        assert_eq!(&buf[0..2], &[0x03, 0x02]); // Record type 0x0203
        assert_eq!(buf.len(), 18); // Header (4) + data (14)

        // Verify row index
        let row = u16::from_le_bytes([buf[4], buf[5]]);
        assert_eq!(row, 5);

        // Verify column index
        let col = u16::from_le_bytes([buf[6], buf[7]]);
        assert_eq!(col, 3);
    }

    #[test]
    fn test_write_labelsst() {
        let mut buf = Vec::new();
        write_labelsst(&mut buf, 1, 2, 0x000F, 42).unwrap();

        assert_eq!(&buf[0..2], &[0xFD, 0x00]); // Record type 0x00FD
        // LABELSST record structure: row(2) + col(2) + xf_idx(2) + sst_idx(4) = 10 bytes
    }

    #[test]
    fn test_write_boolerr() {
        let mut buf = Vec::new();
        write_boolerr(&mut buf, 2, 3, 0x000F, true).unwrap();

        assert_eq!(&buf[0..2], &[0x05, 0x02]); // Record type 0x0205

        // Check row and col
        let row = u16::from_le_bytes([buf[4], buf[5]]);
        let col = u16::from_le_bytes([buf[6], buf[7]]);
        assert_eq!(row, 2);
        assert_eq!(col, 3);
    }

    #[test]
    fn test_write_dimensions() {
        let mut buf = Vec::new();
        write_dimensions(&mut buf, 0, 100, 0, 50).unwrap();

        assert_eq!(&buf[0..2], &[0x00, 0x02]); // Record type 0x0200
        assert_eq!(&buf[2..4], &[14, 0]); // Length = 14

        // Verify dimensions
        let first_row = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
        let last_row = u32::from_le_bytes([buf[8], buf[9], buf[10], buf[11]]);
        assert_eq!(first_row, 0);
        assert_eq!(last_row, 100);
    }

    #[test]
    fn test_write_codepage() {
        let mut buf = Vec::new();
        write_codepage(&mut buf, 1252).unwrap();

        assert_eq!(&buf[0..2], &[0x42, 0x00]); // Record type 0x0042
        assert_eq!(&buf[2..4], &[2, 0]); // Length = 2
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 1252);
    }

    #[test]
    fn test_write_date1904() {
        let mut buf = Vec::new();
        write_date1904(&mut buf, false).unwrap();

        assert_eq!(&buf[0..2], &[0x22, 0x00]); // Record type 0x0022
        assert_eq!(&buf[2..4], &[2, 0]); // Length = 2
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 0);

        let mut buf2 = Vec::new();
        write_date1904(&mut buf2, true).unwrap();
        assert_eq!(u16::from_le_bytes([buf2[4], buf2[5]]), 1);
    }

    #[test]
    fn test_write_boundsheet() {
        let mut buf = Vec::new();
        write_boundsheet(&mut buf, 4096, "Sheet1").unwrap();

        assert_eq!(&buf[0..2], &[0x85, 0x00]); // Record type 0x0085

        // Check position
        let pos = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
        assert_eq!(pos, 4096);
    }

    #[test]
    fn test_write_row() {
        let mut buf = Vec::new();
        write_row(&mut buf, 10, 0, 5, 255, false).unwrap();

        assert_eq!(&buf[0..2], &[0x08, 0x02]); // Record type 0x0208

        // Verify row index
        let row_idx = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
        assert_eq!(row_idx, 10);
    }

    #[test]
    fn test_write_colinfo() {
        let mut buf = Vec::new();
        write_colinfo(&mut buf, 0, 5, 2560, false).unwrap();

        assert_eq!(&buf[0..2], &[0x7D, 0x00]); // Record type 0x007D
        assert_eq!(&buf[2..4], &[12, 0]); // Length = 12

        let first_col = u16::from_le_bytes([buf[4], buf[5]]);
        let last_col = u16::from_le_bytes([buf[6], buf[7]]);
        assert_eq!(first_col, 0);
        assert_eq!(last_col, 5);
    }

    #[test]
    fn test_write_window1() {
        let mut buf = Vec::new();
        write_window1(
            &mut buf,
            &crate::writer::core::WorkbookWindowOptions::default(),
            1,
        )
        .unwrap();

        assert_eq!(&buf[0..2], &[0x3D, 0x00]); // Record type 0x003D
        assert_eq!(&buf[2..4], &[18, 0]); // Length = 18
    }

    #[test]
    fn test_write_def_col_width() {
        let mut buf = Vec::new();
        write_def_col_width(&mut buf, 8).unwrap();

        assert_eq!(&buf[0..2], &[0x55, 0x00]); // Record type 0x0055
        assert_eq!(&buf[2..4], &[2, 0]); // Length = 2
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 8);
    }

    #[test]
    fn test_write_record_header() {
        let mut buf = Vec::new();
        write_record_header(&mut buf, 0x0203, 14).unwrap();

        assert_eq!(u16::from_le_bytes([buf[0], buf[1]]), 0x0203);
        assert_eq!(u16::from_le_bytes([buf[2], buf[3]]), 14);
    }

    #[test]
    fn complete_record_encoder_preserves_exact_frame_bytes_at_limit() {
        let payload = vec![0xA5; litchi_biff::MAX_RECORD_BYTES];
        let mut encoded = Vec::new();

        write_record(&mut encoded, 0x7FFF, &payload).unwrap();

        assert_eq!(encoded.len(), 4 + litchi_biff::MAX_RECORD_BYTES);
        assert_eq!(&encoded[..4], &[0xFF, 0x7F, 0x20, 0x20]);
        assert_eq!(&encoded[4..], payload.as_slice());
    }

    #[test]
    fn complete_record_encoder_rejects_payload_above_biff_limit_without_output() {
        let maximum = litchi_biff::MAX_RECORD_BYTES;
        let payload = vec![0x5A; maximum + 1];
        let mut encoded = Vec::new();

        let error = write_record(&mut encoded, 0x0001, &payload).unwrap_err();

        assert!(matches!(
            error,
            Error::InvalidLength {
                expected,
                found,
            }
            if expected == maximum && found == maximum + 1
        ));
        assert!(encoded.is_empty());
    }

    #[test]
    fn test_has_multibyte_char() {
        assert!(!has_multibyte_char("Hello"));
        assert!(!has_multibyte_char("ABC123"));
        assert!(has_multibyte_char("日本語"));
        assert!(has_multibyte_char("Hello 世界"));
    }

    #[test]
    fn test_unicode_string_size() {
        let size_ascii = unicode_string_size("Hello");
        assert_eq!(size_ascii, 2 + 1 + 5); // char_count (2) + options (1) + chars (5)

        let size_unicode = unicode_string_size("日本語");
        assert_eq!(size_unicode, 2 + 1 + 6); // char_count (2) + options (1) + chars*2 (6)
    }

    #[test]
    fn test_write_unicode_string_biff8_ascii() {
        let mut buf = Vec::new();
        write_unicode_string_biff8(&mut buf, "Test").unwrap();

        assert_eq!(u16::from_le_bytes([buf[0], buf[1]]), 4); // char count
        assert_eq!(buf[2], 0x00); // ASCII flag
        assert_eq!(&buf[3..7], b"Test");
    }

    #[test]
    fn test_write_format_record() {
        let mut buf = Vec::new();
        write_format_record(&mut buf, 164, "0.00\"mm\"").unwrap();

        assert_eq!(&buf[0..2], &[0x1E, 0x04]); // Record type 0x041E
        // Format index should be at offset 4
        let fmt_idx = u16::from_le_bytes([buf[4], buf[5]]);
        assert_eq!(fmt_idx, 164);
    }

    #[test]
    fn test_write_tab_id() {
        let mut buf = Vec::new();
        write_tab_id(&mut buf, 3).unwrap();

        assert_eq!(&buf[0..2], &[0x3D, 0x01]); // Record type 0x013D
        assert_eq!(&buf[2..4], &[6, 0]); // Length = 6 (3 * 2 bytes)
    }

    #[test]
    fn test_write_protect() {
        let mut buf = Vec::new();
        write_protect(&mut buf, true).unwrap();

        assert_eq!(&buf[0..2], &[0x12, 0x00]); // Record type 0x0012
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 1);

        let mut buf2 = Vec::new();
        write_protect(&mut buf2, false).unwrap();
        assert_eq!(u16::from_le_bytes([buf2[4], buf2[5]]), 0);
    }

    #[test]
    fn test_write_window_protect() {
        let mut buf = Vec::new();
        write_window_protect(&mut buf, true).unwrap();

        assert_eq!(&buf[0..2], &[0x19, 0x00]); // Record type 0x0019
    }

    #[test]
    fn test_write_backup() {
        let mut buf = Vec::new();
        write_backup(&mut buf, true).unwrap();

        assert_eq!(&buf[0..2], &[0x40, 0x00]); // Record type 0x0040
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 1);
    }

    #[test]
    fn test_write_refresh_all() {
        let mut buf = Vec::new();
        write_refresh_all(&mut buf, false).unwrap();

        assert_eq!(&buf[0..2], &[0xB7, 0x01]); // Record type 0x01B7
    }

    #[test]
    fn test_write_book_bool() {
        let mut buf = Vec::new();
        write_book_bool(&mut buf, true).unwrap();

        assert_eq!(&buf[0..2], &[0xDA, 0x00]); // Record type 0x00DA
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 1);
    }

    #[test]
    fn test_write_country() {
        let mut buf = Vec::new();
        write_country(&mut buf, 1, 1).unwrap();

        assert_eq!(&buf[0..2], &[0x8C, 0x00]); // Record type 0x008C
        assert_eq!(&buf[2..4], &[4, 0]); // Length = 4
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 1);
        assert_eq!(u16::from_le_bytes([buf[6], buf[7]]), 1);
    }

    #[test]
    fn test_write_precision() {
        let mut buf = Vec::new();
        write_precision(&mut buf, true).unwrap();

        assert_eq!(&buf[0..2], &[0x0E, 0x00]); // Record type 0x000E
    }

    #[test]
    fn test_write_hide_obj() {
        let mut buf = Vec::new();
        write_hide_obj(&mut buf, 2).unwrap();

        assert_eq!(&buf[0..2], &[0x8D, 0x00]); // Record type 0x008D
        assert_eq!(u16::from_le_bytes([buf[4], buf[5]]), 2);
    }

    #[test]
    fn test_write_function_groups() {
        let mut buf = Vec::new();
        write_function_groups(
            &mut buf,
            &crate::writer::core::FunctionGroupOptions::default(),
        )
        .unwrap();

        assert_eq!(&buf[0..2], &[0x9C, 0x00]); // Record type 0x009C
        // Record data follows the 4-byte header
    }
}
