//! Workbook-level BIFF8 record writers.

use crate::xls::{XlsError, XlsResult};
use std::io::Write;

use super::write_record_header;

const WRITE_ACCESS_DATA_LEN: u16 = 112;

/// Workbook-global EXTERNSHEET layout mode.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExternSheetMode {
    /// Emit one XTI per worksheet so 3D tokens can use the sheet index as `ixti`.
    PerSheet,
    /// Emit a single XTI spanning the whole workbook, matching Excel's pivot-only output.
    WorkbookWide,
}

/// Write INTERFACEHDR record.
///
/// Record type: 0x00E1
pub fn write_interface_hdr<W: Write>(writer: &mut W, codepage: u16) -> XlsResult<()> {
    write_record_header(writer, 0x00E1, 2)?;
    writer.write_all(&codepage.to_le_bytes())?;
    Ok(())
}

/// Write MMS record.
///
/// Record type: 0x00C1
pub fn write_mms<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x00C1, 2)?;
    writer.write_all(&[0u8, 0u8])?;
    Ok(())
}

/// Write INTERFACEEND record.
///
/// Record type: 0x00E2
pub fn write_interface_end<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x00E2, 0)?;
    Ok(())
}

/// Write WRITEACCESS record.
///
/// Record type: 0x005C
pub fn write_write_access<W: Write>(writer: &mut W, username: &str) -> XlsResult<()> {
    let access = crate::xls::access::XlsWriteAccess::try_new(username)?;
    write_record_header(writer, 0x005C, WRITE_ACCESS_DATA_LEN)?;
    writer.write_all(&access.to_payload()?)?;
    Ok(())
}

/// Write WINDOWPROTECT record.
///
/// Record type: 0x0019
pub fn write_window_protect<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0019, 2)?;
    writer.write_all(&(u16::from(protect)).to_le_bytes())?;
    Ok(())
}

/// Write PROTECT record.
///
/// Record type: 0x0012
pub fn write_protect<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0012, 2)?;
    writer.write_all(&(u16::from(protect)).to_le_bytes())?;
    Ok(())
}

/// Write PASSWORD record.
///
/// Record type: 0x0013
pub fn write_password<W: Write>(writer: &mut W, password_hash: u16) -> XlsResult<()> {
    write_record_header(writer, 0x0013, 2)?;
    writer.write_all(&password_hash.to_le_bytes())?;
    Ok(())
}

/// Write PROTECTIONREV4 record.
///
/// Record type: 0x01AF
pub fn write_protection_rev4<W: Write>(writer: &mut W, protect: bool) -> XlsResult<()> {
    write_record_header(writer, 0x01AF, 2)?;
    writer.write_all(&(u16::from(protect)).to_le_bytes())?;
    Ok(())
}

/// Write PASSWORDREV4 record.
///
/// Record type: 0x01BC
pub fn write_password_rev4<W: Write>(writer: &mut W, password_hash: u16) -> XlsResult<()> {
    write_record_header(writer, 0x01BC, 2)?;
    writer.write_all(&password_hash.to_le_bytes())?;
    Ok(())
}

/// Write the empty WRITEPROTECT marker.
pub fn write_write_protect<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0086, 0)?;
    Ok(())
}

/// Write FILESHARING write-reservation metadata.
pub fn write_file_sharing<W: Write>(
    writer: &mut W,
    read_only_recommended: bool,
    password_hash: Option<u16>,
    user_name: &str,
) -> XlsResult<()> {
    let password_hash = password_hash.unwrap_or(0);
    if password_hash == 0 {
        write_record_header(writer, 0x005B, 6)?;
        writer.write_all(&u16::from(read_only_recommended).to_le_bytes())?;
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&0u16.to_le_bytes())?;
        return Ok(());
    }

    let utf16 = user_name.encode_utf16().collect::<Vec<_>>();
    if utf16.len() > 54 {
        return Err(XlsError::InvalidData(
            "FILESHARING username exceeds 54 UTF-16 code units".to_string(),
        ));
    }
    let compressed = utf16.iter().all(|unit| *unit <= 0x00FF);
    let char_bytes = utf16.len() * if compressed { 1 } else { 2 };
    let data_len = u16::try_from(7 + char_bytes).map_err(|_| {
        XlsError::InvalidData("FILESHARING payload exceeds BIFF8 record size".to_string())
    })?;
    write_record_header(writer, 0x005B, data_len)?;
    writer.write_all(&u16::from(read_only_recommended).to_le_bytes())?;
    writer.write_all(&password_hash.to_le_bytes())?;
    writer.write_all(&(utf16.len() as u16).to_le_bytes())?;
    writer.write_all(&[u8::from(!compressed)])?;
    for unit in utf16 {
        if compressed {
            writer.write_all(&[unit as u8])?;
        } else {
            writer.write_all(&unit.to_le_bytes())?;
        }
    }
    Ok(())
}

/// Write BACKUP record.
///
/// Record type: 0x0040
pub fn write_backup<W: Write>(writer: &mut W, backup: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0040, 2)?;
    writer.write_all(&(u16::from(backup)).to_le_bytes())?;
    Ok(())
}

/// Write HIDEOBJ record.
///
/// Record type: 0x008D
pub fn write_hide_obj<W: Write>(writer: &mut W, mode: u16) -> XlsResult<()> {
    write_record_header(writer, 0x008D, 2)?;
    writer.write_all(&mode.to_le_bytes())?;
    Ok(())
}

/// Write PRECISION record.
///
/// Record type: 0x000E
pub fn write_precision<W: Write>(writer: &mut W, full_precision: bool) -> XlsResult<()> {
    write_record_header(writer, 0x000E, 2)?;
    writer.write_all(&(u16::from(full_precision)).to_le_bytes())?;
    Ok(())
}

/// Write DSF record.
///
/// Record type: 0x0161
pub fn write_dsf<W: Write>(writer: &mut W, has_biff5_stream: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0161, 2)?;
    writer.write_all(&(u16::from(has_biff5_stream)).to_le_bytes())?;
    Ok(())
}

/// Write TABID record.
///
/// Record type: 0x013D
pub fn write_tab_id<W: Write>(writer: &mut W, sheet_count: u16) -> XlsResult<()> {
    if sheet_count == 0 || sheet_count > 4112 {
        return Err(XlsError::InvalidData("RRTabId requires 1..=4112 sheets".to_string()));
    }
    write_record_header(writer, 0x013D, sheet_count * 2)?;
    for sheet_idx in 0..sheet_count {
        writer.write_all(&sheet_idx.saturating_add(1).to_le_bytes())?;
    }
    Ok(())
}

/// Write FNGROUPCOUNT record.
///
/// Record type: 0x009C
pub fn write_fn_group_count<W: Write>(writer: &mut W, count: u16) -> XlsResult<()> {
    write_record_header(writer, 0x009C, 2)?;
    writer.write_all(&count.to_le_bytes())?;
    Ok(())
}

fn write_function_group_name<W: Write>(
    writer: &mut W,
    record_type: u16,
    name: &str,
) -> XlsResult<()> {
    let units = name.encode_utf16().collect::<Vec<_>>();
    if units.len() > 32 {
        return Err(XlsError::InvalidData("function category name exceeds 32 UTF-16 code units".to_string()));
    }
    let compressed = units.iter().all(|unit| *unit <= 0x00ff);
    let string_size = 3 + units.len() * if compressed { 1 } else { 2 };
    let header_size = if record_type == 0x0898 { 12 } else { 0 };
    write_record_header(writer, record_type, (header_size + string_size) as u16)?;
    if record_type == 0x0898 {
        writer.write_all(&0x0898u16.to_le_bytes())?;
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&[0; 8])?;
    }
    writer.write_all(&(units.len() as u16).to_le_bytes())?;
    writer.write_all(&[u8::from(!compressed)])?;
    for unit in units {
        if compressed {
            writer.write_all(&[unit as u8])?;
        } else {
            writer.write_all(&unit.to_le_bytes())?;
        }
    }
    Ok(())
}

pub fn write_function_groups<W: Write>(
    writer: &mut W,
    options: &crate::xls::writer::core::XlsFunctionGroupOptions,
) -> XlsResult<()> {
    options.validate()?;
    let built_in = options.built_in.count();
    write_fn_group_count(writer, built_in)?;
    let classic_count = options.custom_categories.len().min(32 - usize::from(built_in));
    for name in &options.custom_categories[..classic_count] {
        write_function_group_name(writer, 0x009a, name)?;
    }
    for name in &options.custom_categories[classic_count..] {
        write_function_group_name(writer, 0x0898, name)?;
    }
    Ok(())
}

/// Write REFRESHALL record.
///
/// Record type: 0x01B7
pub fn write_refresh_all<W: Write>(writer: &mut W, refresh_all: bool) -> XlsResult<()> {
    write_record_header(writer, 0x01B7, 2)?;
    writer.write_all(&(u16::from(refresh_all)).to_le_bytes())?;
    Ok(())
}

/// Write BOOKBOOL record.
///
/// Record type: 0x00DA
#[allow(dead_code)] // Compatibility implementation for the former fixed-bit API.
pub fn write_book_bool<W: Write>(writer: &mut W, save_link_values: bool) -> XlsResult<()> {
    write_record_header(writer, 0x00DA, 2)?;
    writer.write_all(&(u16::from(save_link_values)).to_le_bytes())?;
    Ok(())
}

pub fn write_book_bool_raw<W: Write>(writer: &mut W, bits: u16) -> XlsResult<()> {
    write_record_header(writer, 0x00DA, 2)?;
    writer.write_all(&bits.to_le_bytes())?;
    Ok(())
}

/// Write COUNTRY record.
///
/// Record type: 0x008C
pub fn write_country<W: Write>(
    writer: &mut W,
    default_country: u16,
    current_country: u16,
) -> XlsResult<()> {
    write_record_header(writer, 0x008C, 4)?;
    writer.write_all(&default_country.to_le_bytes())?;
    writer.write_all(&current_country.to_le_bytes())?;
    Ok(())
}

/// Write EXCEL9FILE record.
///
/// Record type: 0x01C0
pub fn write_excel9_file<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x01C0, 0)?;
    Ok(())
}

/// Write RECALCID record.
///
/// Record type: 0x01C1
pub fn write_recalc_id<W: Write>(writer: &mut W, engine_id: u32) -> XlsResult<()> {
    write_record_header(writer, 0x01C1, 8)?;
    writer.write_all(&0x01C1u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&engine_id.to_le_bytes())?;
    Ok(())
}

pub fn write_force_full_calculation<W: Write>(writer: &mut W, force: bool) -> XlsResult<()> {
    write_record_header(writer, 0x08A3, 16)?;
    writer.write_all(&0x08A3u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u64.to_le_bytes())?;
    writer.write_all(&u32::from(force).to_le_bytes())?;
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
#[allow(dead_code)] // Compatibility implementation for the former fixed-value API.
pub fn write_usesel_fs<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0160, 2)?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

pub fn write_usesel_fs_value<W: Write>(writer: &mut W, enabled: bool) -> XlsResult<()> {
    write_record_header(writer, 0x0160, 2)?;
    writer.write_all(&u16::from(enabled).to_le_bytes())?;
    Ok(())
}

pub fn write_template<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0060, 0)
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

    writer.write_all(&0x0600u16.to_le_bytes())?;
    writer.write_all(&substream_type.to_le_bytes())?;
    writer.write_all(&0x4F5Au16.to_le_bytes())?;
    writer.write_all(&0x07CDu16.to_le_bytes())?;
    writer.write_all(&0x0002_00C1u32.to_le_bytes())?;
    writer.write_all(&0x0000_0806u32.to_le_bytes())?;

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
pub fn write_window1<W: Write>(
    writer: &mut W,
    options: &crate::xls::writer::core::XlsWorkbookWindowOptions,
    sheet_count: usize,
) -> XlsResult<()> {
    options.validate_for_sheet_count(sheet_count)?;
    write_record_header(writer, 0x003D, 18)?;
    writer.write_all(&options.horizontal_position_twips.to_le_bytes())?;
    writer.write_all(&options.vertical_position_twips.to_le_bytes())?;
    writer.write_all(&options.width_twips.to_le_bytes())?;
    writer.write_all(&options.height_twips.to_le_bytes())?;
    let flags = u16::from(options.hidden)
        | (u16::from(options.minimized) << 1)
        | (u16::from(options.very_hidden) << 2)
        | (u16::from(options.show_horizontal_scrollbar) << 3)
        | (u16::from(options.show_vertical_scrollbar) << 4)
        | (u16::from(options.show_sheet_tabs) << 5)
        | (u16::from(!options.group_dates_in_autofilter) << 6);
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&options.active_sheet_index.to_le_bytes())?;
    writer.write_all(&options.first_visible_sheet_index.to_le_bytes())?;
    writer.write_all(&options.selected_sheet_count.to_le_bytes())?;
    writer.write_all(&options.sheet_tab_ratio_per_mille.to_le_bytes())?;

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
/// For internal references we either generate one XTI per worksheet,
/// or a single workbook-wide XTI spanning all sheets. The workbook-wide
/// form matches Excel's pivot-only output, while the per-sheet form keeps
/// `PtgArea3d.ixti == target_sheet` valid for defined names.
fn write_unicode_string<W: Write>(writer: &mut W, value: &str, include_count: bool) -> XlsResult<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let compressed = units.iter().all(|unit| *unit <= 0x00ff);
    if include_count {
        writer.write_all(&(units.len() as u16).to_le_bytes())?;
    }
    writer.write_all(&[u8::from(!compressed)])?;
    for unit in units {
        if compressed { writer.write_all(&[unit as u8])?; }
        else { writer.write_all(&unit.to_le_bytes())?; }
    }
    Ok(())
}

fn write_external_supbook<W: Write>(
    writer: &mut W,
    book: &crate::xls::writer::core::XlsExternalWorkbookOptions,
    names: &[crate::xls::writer::core::XlsExternalDefinedNameOptions],
) -> XlsResult<()> {
    book.validate()?;
    let path_units = book.encoded_virtual_path.encode_utf16().count();
    let mut data = Vec::new();
    data.extend_from_slice(&(book.sheets.len() as u16).to_le_bytes());
    data.extend_from_slice(&(path_units as u16).to_le_bytes());
    write_unicode_string(&mut data, &book.encoded_virtual_path, false)?;
    for sheet in &book.sheets { write_unicode_string(&mut data, &sheet.name, true)?; }
    if data.len() > 8224 {
        return Err(XlsError::InvalidData("external SupBook exceeds BIFF8 record size".to_string()));
    }
    write_record_header(writer, 0x01ae, data.len() as u16)?;
    writer.write_all(&data)?;
    for name in names { write_external_defined_name(writer, name)?; }
    for (sheet_index, sheet) in book.sheets.iter().enumerate() {
        if sheet.cache_rows.is_empty() { continue; }
        write_record_header(writer, 0x0059, 4)?;
        writer.write_all(&(sheet.cache_rows.len() as i16).to_le_bytes())?;
        writer.write_all(&(sheet_index as u16).to_le_bytes())?;
        for row in &sheet.cache_rows { write_crn(writer, row)?; }
    }
    Ok(())
}

fn write_short_unicode_string<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    writer.write_all(&[units.len() as u8])?;
    write_unicode_string(writer, value, false)
}

fn write_external_defined_name<W: Write>(
    writer: &mut W,
    name: &crate::xls::writer::core::XlsExternalDefinedNameOptions,
) -> XlsResult<()> {
    let mut data = Vec::new();
    data.extend_from_slice(&u16::from(name.built_in).to_le_bytes());
    data.extend_from_slice(&name.sheet_index.map_or(0, |index| index + 1).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    write_short_unicode_string(&mut data, &name.name)?;
    data.extend_from_slice(&(name.formula_bytes.len() as u16).to_le_bytes());
    data.extend_from_slice(&name.formula_bytes);
    if data.len() > 8224 {
        return Err(XlsError::InvalidData("external defined name exceeds BIFF8 record size".to_string()));
    }
    write_record_header(writer, 0x0023, data.len() as u16)?;
    writer.write_all(&data)?;
    Ok(())
}

fn write_add_in_supbook<W: Write>(
    writer: &mut W,
    functions: &[crate::xls::writer::core::XlsAddInFunctionOptions],
) -> XlsResult<()> {
    write_record_header(writer, 0x01ae, 4)?;
    writer.write_all(&1u16.to_le_bytes())?;
    writer.write_all(&0x3a01u16.to_le_bytes())?;
    for function in functions {
        let mut data = vec![0; 6];
        write_short_unicode_string(&mut data, &function.name)?;
        data.extend_from_slice(&(function.unused_data.len() as u16).to_le_bytes());
        data.extend_from_slice(&function.unused_data);
        if data.len() > 8224 {
            return Err(XlsError::InvalidData("add-in ExternName exceeds BIFF8 record size".to_string()));
        }
        write_record_header(writer, 0x0023, data.len() as u16)?;
        writer.write_all(&data)?;
    }
    Ok(())
}

fn encode_clipboard_format(value: i16) -> u16 { (value as u16) & 0x03ff }

fn write_dde_or_ole_supbook<W: Write>(
    writer: &mut W,
    link: &crate::xls::writer::core::XlsDdeOrOleLinkOptions,
) -> XlsResult<()> {
    let mut supbook = Vec::new();
    supbook.extend_from_slice(&0u16.to_le_bytes());
    supbook.extend_from_slice(&(link.encoded_virtual_path.encode_utf16().count() as u16).to_le_bytes());
    write_unicode_string(&mut supbook, &link.encoded_virtual_path, false)?;
    write_record_header(writer, 0x01ae, supbook.len() as u16)?;
    writer.write_all(&supbook)?;
    for item in &link.items {
        let mut flags = (encode_clipboard_format(item.clipboard_format) << 5)
            | u16::from(item.automatic) << 1
            | u16::from(item.picture) << 2
            | u16::from(item.standard_document_name) << 3
            | u16::from(item.ole_link) << 4
            | u16::from(item.displayed_as_icon) << 15;
        flags &= 0xffff;
        let mut data = Vec::new();
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&item.storage_id.to_le_bytes());
        write_short_unicode_string(&mut data, &item.name)?;
        data.extend_from_slice(&item.opaque_data);
        if data.len() > 8224 {
            return Err(XlsError::InvalidData("DDE/OLE ExternName exceeds BIFF8 record size".to_string()));
        }
        write_record_header(writer, 0x0023, data.len() as u16)?;
        writer.write_all(&data)?;
        for chunk in &item.continuation_chunks {
            write_record_header(writer, 0x003c, chunk.len() as u16)?;
            writer.write_all(chunk)?;
        }
    }
    Ok(())
}

fn write_crn<W: Write>(
    writer: &mut W,
    row: &crate::xls::writer::core::XlsExternalCacheRowOptions,
) -> XlsResult<()> {
    let mut data = Vec::new();
    data.push(row.first_column + row.values.len() as u8 - 1);
    data.push(row.first_column);
    data.extend_from_slice(&row.row.to_le_bytes());
    for value in &row.values {
        match value {
            crate::xls::XlsExternalCachedValue::Blank => data.extend_from_slice(&[0; 9]),
            crate::xls::XlsExternalCachedValue::Number(value) => {
                data.push(0x01); data.extend_from_slice(&value.to_le_bytes());
            },
            crate::xls::XlsExternalCachedValue::Text(value) => {
                data.push(0x02); write_unicode_string(&mut data, value, true)?;
            },
            crate::xls::XlsExternalCachedValue::Boolean(value) => {
                data.extend_from_slice(&[0x04, u8::from(*value), 0, 0, 0, 0, 0, 0, 0]);
            },
            crate::xls::XlsExternalCachedValue::Error(error) => {
                data.extend_from_slice(&[0x10, error.code(), 0, 0, 0, 0, 0, 0, 0]);
            },
        }
    }
    write_record_header(writer, 0x005a, data.len() as u16)?;
    writer.write_all(&data)?;
    Ok(())
}

pub fn write_external_link_table<W: Write>(
    writer: &mut W,
    internal: Option<(u16, ExternSheetMode)>,
    external: &[crate::xls::writer::core::XlsExternalWorkbookOptions],
    external_names: &[Vec<crate::xls::writer::core::XlsExternalDefinedNameOptions>],
    add_in_functions: &[crate::xls::writer::core::XlsAddInFunctionOptions],
    dde_or_ole_links: &[crate::xls::writer::core::XlsDdeOrOleLinkOptions],
) -> XlsResult<()> {
    if let Some((sheet_count, _)) = internal { write_supbook_internal(writer, sheet_count)?; }
    if external.len() != external_names.len() {
        return Err(XlsError::InvalidData("external workbook/name table cardinality mismatch".to_string()));
    }
    for (book, names) in external.iter().zip(external_names) {
        write_external_supbook(writer, book, names)?;
    }
    if !add_in_functions.is_empty() { write_add_in_supbook(writer, add_in_functions)?; }
    for link in dde_or_ole_links { write_dde_or_ole_supbook(writer, link)?; }

    let internal_count = internal.map_or(0usize, |(sheet_count, mode)| match mode {
        ExternSheetMode::PerSheet => usize::from(sheet_count),
        ExternSheetMode::WorkbookWide => 1,
    });
    let external_count = external.iter().map(|book| book.sheets.len()).sum::<usize>();
    let add_in_count = usize::from(!add_in_functions.is_empty());
    let count = internal_count + external_count + add_in_count + dde_or_ole_links.len();
    if count > 1370 {
        return Err(XlsError::InvalidData("ExternSheet reference count exceeds BIFF8 record bound".to_string()));
    }
    write_record_header(writer, 0x0017, (2 + count * 6) as u16)?;
    writer.write_all(&(count as u16).to_le_bytes())?;
    if let Some((sheet_count, mode)) = internal {
        match mode {
            ExternSheetMode::PerSheet => for sheet in 0..sheet_count {
                writer.write_all(&0u16.to_le_bytes())?;
                writer.write_all(&(sheet as i16).to_le_bytes())?;
                writer.write_all(&(sheet as i16).to_le_bytes())?;
            },
            ExternSheetMode::WorkbookWide => {
                writer.write_all(&0u16.to_le_bytes())?;
                writer.write_all(&0i16.to_le_bytes())?;
                writer.write_all(&(sheet_count as i16 - 1).to_le_bytes())?;
            },
        }
    }
    let first_external_book = usize::from(internal.is_some());
    for (book_offset, book) in external.iter().enumerate() {
        let book_index = u16::try_from(first_external_book + book_offset)
            .map_err(|_| XlsError::InvalidData("supporting-book index exceeds u16".to_string()))?;
        for sheet in 0..book.sheets.len() {
            writer.write_all(&book_index.to_le_bytes())?;
            writer.write_all(&(sheet as i16).to_le_bytes())?;
            writer.write_all(&(sheet as i16).to_le_bytes())?;
        }
    }
    let add_in_book = first_external_book + external.len();
    if !add_in_functions.is_empty() {
        writer.write_all(&(add_in_book as u16).to_le_bytes())?;
        writer.write_all(&(-2i16).to_le_bytes())?;
        writer.write_all(&(-2i16).to_le_bytes())?;
    }
    let first_dde_book = add_in_book + add_in_count;
    for offset in 0..dde_or_ole_links.len() {
        writer.write_all(&u16::try_from(first_dde_book + offset)
            .map_err(|_| XlsError::InvalidData("supporting-book index exceeds u16".to_string()))?
            .to_le_bytes())?;
        writer.write_all(&(-2i16).to_le_bytes())?;
        writer.write_all(&(-2i16).to_le_bytes())?;
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
