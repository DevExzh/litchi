//! Worksheet-level BIFF8 record writers.

use crate::page_setup::{XlsPrintComments, XlsPrintErrors, XlsPrintOrder, XlsPrintOrientation};
use crate::writer::core::{XlsPageSetupOptions, XlsWorksheetLayoutOptions};
use crate::{XlsError, XlsResult};
use std::io::Write;

use super::write_record_header;

fn write_unicode_string<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let compressed = units.iter().all(|unit| *unit <= 0xff);
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

fn write_dcon_file<W: Write>(writer: &mut W, file: &crate::XlsConsolidationFile) -> XlsResult<()> {
    let units = file.encoded_path().encode_utf16().collect::<Vec<_>>();
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    writer.write_all(&(units.len() as u16).to_le_bytes())?;
    writer.write_all(&[u8::from(!compressed)])?;
    for unit in units {
        if compressed {
            writer.write_all(&[unit as u8])?;
        } else {
            writer.write_all(&unit.to_le_bytes())?;
        }
    }
    if file.is_self_reference() {
        writer.write_all(if compressed { &[0][..] } else { &[0, 0][..] })?;
    }
    Ok(())
}

/// Write one complete contiguous `DCON` worksheet directory.
pub fn write_consolidation<W: Write>(
    writer: &mut W,
    consolidation: &crate::XlsConsolidation,
) -> XlsResult<()> {
    consolidation.validate_for_write()?;
    write_record_header(writer, crate::consolidation::DCON_RECORD_TYPE, 8)?;
    writer.write_all(&consolidation.function().code().to_le_bytes())?;
    writer.write_all(&u16::from(consolidation.uses_left_labels()).to_le_bytes())?;
    writer.write_all(&u16::from(consolidation.uses_top_labels()).to_le_bytes())?;
    writer.write_all(&u16::from(consolidation.creates_links()).to_le_bytes())?;
    for source in consolidation.sources() {
        let mut payload = Vec::new();
        let record_type = match source {
            crate::XlsConsolidationSource::CellRange { range, file } => {
                payload.extend_from_slice(&range.first_row().to_le_bytes());
                payload.extend_from_slice(&range.last_row().to_le_bytes());
                payload.push(range.first_column());
                payload.push(range.last_column());
                write_dcon_file(&mut payload, file)?;
                crate::consolidation::DCON_REF_RECORD_TYPE
            },
            crate::XlsConsolidationSource::DefinedName { name, file } => {
                write_unicode_string(&mut payload, name)?;
                if let Some(file) = file {
                    write_dcon_file(&mut payload, file)?;
                } else {
                    payload.extend_from_slice(&0u16.to_le_bytes());
                }
                crate::consolidation::DCON_NAME_RECORD_TYPE
            },
            crate::XlsConsolidationSource::BuiltInName { name, file } => {
                payload.push(name.code());
                payload.extend_from_slice(&0u16.to_le_bytes());
                payload.push(0);
                if let Some(file) = file {
                    write_dcon_file(&mut payload, file)?;
                } else {
                    payload.extend_from_slice(&0u16.to_le_bytes());
                }
                crate::consolidation::DCON_BIN_RECORD_TYPE
            },
        };
        write_record_header(writer, record_type, payload.len() as u16)?;
        writer.write_all(&payload)?;
    }
    Ok(())
}

fn write_bool_record<W: Write>(writer: &mut W, record_type: u16, value: bool) -> XlsResult<()> {
    write_record_header(writer, record_type, 2)?;
    writer.write_all(&u16::from(value).to_le_bytes())?;
    Ok(())
}

fn write_header_footer<W: Write>(writer: &mut W, record_type: u16, text: &str) -> XlsResult<()> {
    let units = text.encode_utf16().collect::<Vec<_>>();
    if units.len() > 255 {
        return Err(XlsError::InvalidData(
            "header/footer exceeds 255 UTF-16 code units".to_string(),
        ));
    }
    let compressed = units.iter().all(|unit| *unit <= 0x00ff);
    let data_len = 3 + units.len() * if compressed { 1 } else { 2 };
    write_record_header(writer, record_type, data_len as u16)?;
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

fn write_page_breaks<W: Write>(
    writer: &mut W,
    record_type: u16,
    page_breaks: &[(u16, u16, u16)],
) -> XlsResult<()> {
    if page_breaks.is_empty() {
        return Ok(());
    }
    let maximum = if record_type == 0x001b { 1026 } else { 255 };
    if page_breaks.len() > maximum {
        return Err(XlsError::InvalidData(
            "page-break count exceeds BIFF8 limit".to_string(),
        ));
    }
    let mut ordered = page_breaks.to_vec();
    ordered.sort_unstable();
    for (index, &(position, range_start, range_end)) in ordered.iter().enumerate() {
        if range_end <= range_start
            || (record_type == 0x001b && range_end > 16383)
            || (record_type == 0x001a && position > 255)
        {
            return Err(XlsError::InvalidData(
                "page-break range is invalid".to_string(),
            ));
        }
        if index > 0 {
            let previous = ordered[index - 1];
            if position == previous.0 && range_start <= previous.2 {
                return Err(XlsError::InvalidData(
                    "page-break ranges overlap".to_string(),
                ));
            }
        }
    }
    write_record_header(writer, record_type, 2 + ordered.len() as u16 * 6)?;
    writer.write_all(&(ordered.len() as u16).to_le_bytes())?;
    for (position, range_start, range_end) in ordered {
        writer.write_all(&position.to_le_bytes())?;
        writer.write_all(&range_start.to_le_bytes())?;
        writer.write_all(&range_end.to_le_bytes())?;
    }
    Ok(())
}

fn write_pls<W: Write>(writer: &mut W, driver_data: &[u8]) -> XlsResult<()> {
    const MAX_PAYLOAD: usize = 8224;
    let first_len = driver_data.len().min(MAX_PAYLOAD - 2);
    write_record_header(writer, 0x004d, (first_len + 2) as u16)?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&driver_data[..first_len])?;
    for chunk in driver_data[first_len..].chunks(MAX_PAYLOAD) {
        write_record_header(writer, 0x003c, chunk.len() as u16)?;
        writer.write_all(chunk)?;
    }
    Ok(())
}

/// Write the primary BIFF8 worksheet page-settings records in canonical order.
pub fn write_page_settings<W: Write>(
    writer: &mut W,
    options: &XlsPageSetupOptions,
    horizontal_breaks: &[(u16, u16, u16)],
    vertical_breaks: &[(u16, u16, u16)],
) -> XlsResult<()> {
    write_bool_record(writer, 0x002a, options.print_headers)?;
    write_bool_record(writer, 0x002b, options.print_gridlines)?;
    write_page_breaks(writer, 0x001b, horizontal_breaks)?;
    write_page_breaks(writer, 0x001a, vertical_breaks)?;
    write_header_footer(writer, 0x0014, &options.header)?;
    write_header_footer(writer, 0x0015, &options.footer)?;
    write_bool_record(writer, 0x0083, options.horizontally_centered)?;
    write_bool_record(writer, 0x0084, options.vertically_centered)?;
    for (record_type, value) in [
        (0x0026, options.left_margin_inches),
        (0x0027, options.right_margin_inches),
        (0x0028, options.top_margin_inches),
        (0x0029, options.bottom_margin_inches),
    ] {
        write_record_header(writer, record_type, 8)?;
        writer.write_all(&value.to_le_bytes())?;
    }
    if let Some(driver_data) = &options.printer_driver_data {
        write_pls(writer, driver_data)?;
    }

    let mut flags = 0u16;
    if options.print_order == XlsPrintOrder::OverThenDown {
        flags |= 0x0001;
    }
    if options.printer_driver_data.is_none() {
        flags |= 0x0004;
    }
    match options.orientation {
        Some(XlsPrintOrientation::Portrait) => flags |= 0x0002,
        Some(XlsPrintOrientation::Landscape) => {},
        None => flags |= 0x0040,
    }
    if options.black_and_white {
        flags |= 0x0008;
    }
    if options.draft_quality {
        flags |= 0x0010;
    }
    match options.comments {
        XlsPrintComments::None => {},
        XlsPrintComments::AsDisplayed => flags |= 0x0020,
        XlsPrintComments::AtEnd => flags |= 0x0220,
    }
    if options.starting_page_number.is_some() {
        flags |= 0x0080;
    }
    flags |= match options.errors {
        XlsPrintErrors::Displayed => 0,
        XlsPrintErrors::Blank => 1 << 10,
        XlsPrintErrors::Dashes => 2 << 10,
        XlsPrintErrors::NotAvailable => 3 << 10,
    };
    write_record_header(writer, 0x00a1, 34)?;
    writer.write_all(&options.paper_size.to_le_bytes())?;
    writer.write_all(&options.scale_percent.to_le_bytes())?;
    writer.write_all(&(options.starting_page_number.unwrap_or(1) as u16).to_le_bytes())?;
    writer.write_all(&options.fit_width_pages.to_le_bytes())?;
    writer.write_all(&options.fit_height_pages.to_le_bytes())?;
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&options.horizontal_resolution_dpi.to_le_bytes())?;
    writer.write_all(&options.vertical_resolution_dpi.to_le_bytes())?;
    writer.write_all(&options.header_margin_inches.to_le_bytes())?;
    writer.write_all(&options.footer_margin_inches.to_le_bytes())?;
    writer.write_all(&options.copies.to_le_bytes())?;
    // HeaderFooter (0x089C) follows the PAGESETUP group in the worksheet
    // substream grammar.
    if let Some(header_footer) = &options.header_footer {
        let payload = header_footer.to_payload()?;
        write_record_header(writer, 0x089C, payload.len() as u16)?;
        writer.write_all(&payload)?;
    }
    Ok(())
}

/// Write DEFCOLWIDTH record.
///
/// Record type: 0x0055, Length: 2
pub fn write_def_col_width<W: Write>(writer: &mut W, width_chars: u16) -> XlsResult<()> {
    if width_chars > 255 {
        return Err(XlsError::InvalidData(
            "default column width exceeds 255 characters".to_string(),
        ));
    }
    write_record_header(writer, 0x0055, 2)?;
    writer.write_all(&width_chars.to_le_bytes())?;
    Ok(())
}

/// Write INDEX record.
///
/// Record type: 0x020B, Length: 16 + 4 * cDbCell
#[allow(dead_code)]
pub fn write_index<W: Write>(
    writer: &mut W,
    first_row: u32,
    last_row_plus1: u32,
    def_col_width_pos: u32,
    dbcell_positions: &[u32],
) -> XlsResult<()> {
    let data_len = 16u16
        + u16::try_from(dbcell_positions.len() * 4).map_err(|_| {
            XlsError::InvalidData(
                "INDEX record DBCell pointer list exceeds BIFF8 size limit".to_string(),
            )
        })?;
    write_record_header(writer, 0x020B, data_len)?;
    writer.write_all(&0u32.to_le_bytes())?;
    writer.write_all(&first_row.to_le_bytes())?;
    writer.write_all(&last_row_plus1.to_le_bytes())?;
    writer.write_all(&def_col_width_pos.to_le_bytes())?;
    for position in dbcell_positions {
        writer.write_all(&position.to_le_bytes())?;
    }
    Ok(())
}

/// Write DBCELL record.
///
/// Record type: 0x00D7, Length: 4 + 2 * cOffsets
#[allow(dead_code)]
pub fn write_dbcell<W: Write>(
    writer: &mut W,
    row_offset: u32,
    cell_offsets: &[u16],
) -> XlsResult<()> {
    let data_len = 4u16
        + u16::try_from(cell_offsets.len() * 2).map_err(|_| {
            XlsError::InvalidData("DBCELL row offset list exceeds BIFF8 size limit".to_string())
        })?;
    write_record_header(writer, 0x00D7, data_len)?;
    writer.write_all(&row_offset.to_le_bytes())?;
    for offset in cell_offsets {
        writer.write_all(&offset.to_le_bytes())?;
    }
    Ok(())
}

/// Write GUTS, DEFAULTROWHEIGHT, and WSBOOL from typed worksheet settings.
pub fn write_worksheet_layout<W: Write>(
    writer: &mut W,
    options: &XlsWorksheetLayoutOptions,
) -> XlsResult<()> {
    options.validate()?;
    write_record_header(writer, 0x0080, 8)?;
    writer.write_all(&options.row_gutter_width.to_le_bytes())?;
    writer.write_all(&options.column_gutter_height.to_le_bytes())?;
    let row_level = if options.max_row_outline_level == 0 {
        0
    } else {
        u16::from(options.max_row_outline_level) + 1
    };
    let column_level = if options.max_column_outline_level == 0 {
        0
    } else {
        u16::from(options.max_column_outline_level) + 1
    };
    writer.write_all(&row_level.to_le_bytes())?;
    writer.write_all(&column_level.to_le_bytes())?;

    let row_flags = u16::from(options.default_row_height_unsynced)
        | (u16::from(options.empty_rows_hidden) << 1)
        | (u16::from(options.thick_top_border) << 2)
        | (u16::from(options.thick_bottom_border) << 3);
    write_record_header(writer, 0x0225, 4)?;
    writer.write_all(&row_flags.to_le_bytes())?;
    writer.write_all(&options.default_row_height_twips.to_le_bytes())?;

    let wsbool = u16::from(options.show_automatic_page_breaks)
        | (u16::from(options.apply_styles_to_outlines) << 5)
        | (u16::from(options.summary_rows_below) << 6)
        | (u16::from(options.summary_columns_right) << 7)
        | (u16::from(options.fit_to_page) << 8)
        | (u16::from(options.synchronize_horizontal_scrolling) << 12)
        | (u16::from(options.synchronize_vertical_scrolling) << 13)
        | (u16::from(options.alternate_expression_evaluation) << 14)
        | (u16::from(options.alternate_formula_entry) << 15);
    write_record_header(writer, 0x0081, 2)?;
    writer.write_all(&wsbool.to_le_bytes())?;
    Ok(())
}

pub fn write_uncalced<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x005E, 2)?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

pub fn write_calculation_settings<W: Write>(
    writer: &mut W,
    settings: &crate::writer::core::XlsCalculationSettings,
) -> XlsResult<()> {
    if !(1..=32_767).contains(&settings.maximum_iterations)
        || !settings.iteration_delta.is_finite()
        || settings.iteration_delta < 0.0
    {
        return Err(XlsError::InvalidData(
            "invalid BIFF8 calculation settings".to_string(),
        ));
    }
    let mode = match settings.mode {
        crate::XlsCalculationMode::Manual => 0u16,
        crate::XlsCalculationMode::Automatic => 1u16,
        crate::XlsCalculationMode::AutomaticExceptTables => 2u16,
    };
    let reference_mode = match settings.reference_mode {
        crate::XlsReferenceMode::R1C1 => 0u16,
        crate::XlsReferenceMode::A1 => 1u16,
    };
    write_record_header(writer, 0x000D, 2)?;
    writer.write_all(&mode.to_le_bytes())?;
    write_record_header(writer, 0x000C, 2)?;
    writer.write_all(&settings.maximum_iterations.to_le_bytes())?;
    write_record_header(writer, 0x000F, 2)?;
    writer.write_all(&reference_mode.to_le_bytes())?;
    write_record_header(writer, 0x0011, 2)?;
    writer.write_all(&u16::from(settings.iteration_enabled).to_le_bytes())?;
    write_record_header(writer, 0x0010, 8)?;
    writer.write_all(&settings.iteration_delta.to_le_bytes())?;
    write_record_header(writer, 0x005F, 2)?;
    writer.write_all(&u16::from(settings.recalculate_before_save).to_le_bytes())?;
    Ok(())
}

pub fn write_pivot_sheet_preamble<W: Write>(
    writer: &mut W,
    options: &XlsWorksheetLayoutOptions,
) -> XlsResult<()> {
    const MARGIN_BYTES: [u8; 8] = [0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0xE6, 0x3F];
    const TOP_BOTTOM_MARGIN_BYTES: [u8; 8] = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE8, 0x3F];
    const PRINT_SETUP_BYTES: [u8; 34] = [
        0x00, 0x00, 0x16, 0x01, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00, 0x04, 0x00, 0x02, 0x00, 0x01,
        0xFF, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xD3, 0x3F, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33,
        0xD3, 0x3F, 0x0F, 0x00,
    ];
    const HEADER_FOOTER_BYTES: [u8; 38] = [
        0x9C, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3C, 0x33,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ];

    write_record_header(writer, 0x002A, 2)?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    write_record_header(writer, 0x002B, 2)?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    write_record_header(writer, 0x0082, 2)?;
    writer.write_all(&0x0001u16.to_le_bytes())?;
    write_worksheet_layout(writer, options)?;
    write_record_header(writer, 0x0014, 0)?;
    write_record_header(writer, 0x0015, 0)?;
    write_record_header(writer, 0x0083, 2)?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    write_record_header(writer, 0x0084, 2)?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    write_record_header(writer, 0x0026, 8)?;
    writer.write_all(&MARGIN_BYTES)?;
    write_record_header(writer, 0x0027, 8)?;
    writer.write_all(&MARGIN_BYTES)?;
    write_record_header(writer, 0x0028, 8)?;
    writer.write_all(&TOP_BOTTOM_MARGIN_BYTES)?;
    write_record_header(writer, 0x0029, 8)?;
    writer.write_all(&TOP_BOTTOM_MARGIN_BYTES)?;
    write_record_header(writer, 0x00A1, 34)?;
    writer.write_all(&PRINT_SETUP_BYTES)?;
    write_record_header(writer, 0x089C, 38)?;
    writer.write_all(&HEADER_FOOTER_BYTES)?;
    Ok(())
}

pub fn write_pivot_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
) -> XlsResult<()> {
    write_record_header(writer, 0x007D, 12)?;
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col.to_le_bytes())?;
    writer.write_all(&col_width.to_le_bytes())?;
    writer.write_all(&0x000Fu16.to_le_bytes())?;
    writer.write_all(&0x0006u16.to_le_bytes())?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    Ok(())
}

/// Write COLINFO record (column formatting and width).
///
/// Record type: 0x007D, Length: 12
///
/// The width is expressed in units of 1/256 of the width of the
/// character "0" in the workbook's default font, matching the
/// semantics of Apache POI's `ColumnInfoRecord.setColumnWidth`.
pub fn write_colinfo<W: Write>(
    writer: &mut W,
    first_col: u16,
    last_col: u16,
    col_width: u16,
    hidden: bool,
) -> XlsResult<()> {
    write_record_header(writer, 0x007D, 12)?;

    // Column range (inclusive)
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col.to_le_bytes())?;

    // Column width in 1/256 character units
    writer.write_all(&col_width.to_le_bytes())?;

    // XF index: use the default cell XF (15) to match our cell records.
    writer.write_all(&0x000Fu16.to_le_bytes())?;

    // Options bitfield: base value 0x0002 as in POI, plus the hidden flag
    // in the least significant bit when required.
    let mut options: u16 = 0x0002;
    if hidden {
        options |= 0x0001;
    }
    writer.write_all(&options.to_le_bytes())?;

    // Reserved field: POI commonly writes 2 here; Excel tolerates non-zero
    // even though the spec marks this as reserved.
    writer.write_all(&0x0002u16.to_le_bytes())?;

    Ok(())
}

/// Write PANE record (freeze panes / split panes)
///
/// Record type: 0x0041, Length: 10
///
/// For the initial implementation we only support classic freeze panes,
/// matching Apache POI's use of `PaneRecord` for HSSF:
/// - `x` and `y` are the split positions in terms of columns/rows.
/// - `topRow` and `leftColumn` are set to the same values.
/// - `activePane` is derived from which sides are frozen.
#[allow(dead_code)]
pub fn write_pane<W: Write>(writer: &mut W, freeze_rows: u32, freeze_cols: u16) -> XlsResult<()> {
    if freeze_rows == 0 && freeze_cols == 0 {
        return Ok(());
    }

    let y = u16::try_from(freeze_rows).map_err(|_| {
        XlsError::InvalidData(
            "freeze_panes: freeze_rows exceeds BIFF8 limit 65535 for PANE record".to_string(),
        )
    })?;
    if freeze_cols > 255 {
        return Err(XlsError::InvalidData(
            "freeze_panes: freeze_cols exceeds BIFF8 frozen PANE limit 255".to_string(),
        ));
    }
    let x = freeze_cols;

    let top_row = y;
    let left_col = x;

    // Active pane constants mirror Apache POI's PaneRecord:
    // 0 = lower-right, 1 = upper-right, 2 = lower-left, 3 = upper-left.
    let active_pane: u16 = match (x > 0, y > 0) {
        (true, true) => 0,  // lower-right
        (true, false) => 1, // upper-right
        (false, true) => 2, // lower-left
        (false, false) => 3,
    };

    write_record_header(writer, 0x0041, 10)?;
    writer.write_all(&x.to_le_bytes())?;
    writer.write_all(&y.to_le_bytes())?;
    writer.write_all(&top_row.to_le_bytes())?;
    writer.write_all(&left_col.to_le_bytes())?;
    writer.write_all(&active_pane.to_le_bytes())?;

    Ok(())
}

pub fn write_autofilterinfo<W: Write>(writer: &mut W, c_entries: u16) -> XlsResult<()> {
    if c_entries == 0 {
        return Ok(());
    }

    write_record_header(writer, 0x009D, 2)?;
    writer.write_all(&c_entries.to_le_bytes())?;
    Ok(())
}

/// Write an AUTOFILTER record (0x009E) for a single column filter condition.
///
/// # Record layout (MS-XLS 2.4.6)
///
/// ```text
/// Offset  Size  Field
/// 0       2     iEntry     — column index within the AutoFilter range (0-based)
/// 2       2     grbit      — option flags
/// 4       10    doper1     — first DOPER condition
/// 14      10    doper2     — second DOPER condition
/// 24      var   rgch1      — string for condition 1 (if applicable)
///         var   rgch2      — string for condition 2 (if applicable)
/// ```
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
    let mut grbit: u16 = 0;
    if join_or {
        grbit |= 0x0001;
    }
    if is_simple {
        grbit |= 0x0002;
    }
    if is_top10 {
        grbit |= 0x0010;
    }
    if hide_arrow {
        grbit |= 0x0020;
    }

    let (doper1, str1) = cond1.to_doper();
    let (doper2, str2) = cond2.to_doper();

    let str1_bytes = encode_autofilter_string(str1);
    let str2_bytes = encode_autofilter_string(str2);

    let data_len = 24u16 + str1_bytes.len() as u16 + str2_bytes.len() as u16;
    write_record_header(writer, 0x009E, data_len)?;

    writer.write_all(&column_index.to_le_bytes())?;
    writer.write_all(&grbit.to_le_bytes())?;
    writer.write_all(&doper1)?;
    writer.write_all(&doper2)?;
    writer.write_all(&str1_bytes)?;
    writer.write_all(&str2_bytes)?;

    Ok(())
}

/// A single filter condition for writing an AUTOFILTER record.
#[derive(Debug, Clone)]
pub enum AutoFilterConditionWrite {
    /// No condition / unused.
    None,
    /// Numeric condition (IEEE 754 double).
    Number { operator: u8, value: f64 },
    /// String condition.
    String { operator: u8, value: String },
    /// Boolean condition.
    Bool { operator: u8, value: bool },
    /// Match all / blanks.
    MatchAll { operator: u8 },
}

impl AutoFilterConditionWrite {
    /// Serialize to a 10-byte DOPER structure + optional string.
    ///
    /// Returns `(doper: [u8; 10], optional_string: Option<&str>)`.
    fn to_doper(&self) -> ([u8; 10], Option<&str>) {
        let mut doper = [0u8; 10];
        match self {
            Self::None => (doper, None),
            Self::Number { operator, value } => {
                doper[0] = 0x04; // vt = IEEE double
                doper[1] = *operator;
                doper[2..10].copy_from_slice(&value.to_le_bytes());
                (doper, None)
            },
            Self::String { operator, value } => {
                doper[0] = 0x06; // vt = string
                doper[1] = *operator;
                // doper[2] = unused, doper[3] = byte length of string
                let byte_len = value.len().min(255) as u8;
                doper[3] = byte_len;
                (doper, Some(value.as_str()))
            },
            Self::Bool { operator, value } => {
                doper[0] = 0x08; // vt = boolean/error
                doper[1] = *operator;
                doper[2] = if *value { 1 } else { 0 };
                doper[3] = 0; // is_error = false
                (doper, None)
            },
            Self::MatchAll { operator } => {
                doper[0] = 0x0C; // vt = match all
                doper[1] = *operator;
                (doper, None)
            },
        }
    }
}

/// Encode a DOPER string operand as XLUnicodeStringNoCch for the trailing bytes.
fn encode_autofilter_string(s: Option<&str>) -> Vec<u8> {
    match s {
        None => Vec::new(),
        Some("") => Vec::new(),
        Some(s) => {
            let is_ascii = s.is_ascii();
            if is_ascii {
                let mut buf = Vec::with_capacity(1 + s.len());
                buf.push(0x00); // flags: compressed
                buf.extend_from_slice(s.as_bytes());
                buf
            } else {
                let utf16: Vec<u16> = s.encode_utf16().collect();
                let mut buf = Vec::with_capacity(1 + utf16.len() * 2);
                buf.push(0x01); // flags: UTF-16LE
                for ch in &utf16 {
                    buf.extend_from_slice(&ch.to_le_bytes());
                }
                buf
            }
        },
    }
}

/// Write a SORT record (0x0090).
///
/// # Record layout
///
/// ```text
/// Offset  Size  Field
/// 0       2     flags   — bit 0: case-sensitive, bit 2: sort by columns (not rows),
///                          bits 4-6: key descending flags
/// 2       2     col1    — first sort key column index
/// 4       2     col2    — second sort key column index (0 if unused)
/// 6       2     col3    — third sort key column index (0 if unused)
/// 8       2     reserved
/// ```
pub fn write_sort<W: Write>(
    writer: &mut W,
    case_sensitive: bool,
    sort_by_columns: bool,
    keys: &[(u16, bool)], // (column_index, descending)
) -> XlsResult<()> {
    if keys.is_empty() {
        return Ok(());
    }

    let mut flags: u16 = 0;
    if case_sensitive {
        flags |= 0x0001;
    }
    if sort_by_columns {
        flags |= 0x0004;
    }
    // Number of keys encoded in bit 1 (0 = 1 key, 1 = 2+ keys)
    if keys.len() >= 2 {
        flags |= 0x0002;
    }
    // Descending flags for each key
    if keys.first().is_some_and(|(_, desc)| *desc) {
        flags |= 0x0010;
    }
    if keys.get(1).is_some_and(|(_, desc)| *desc) {
        flags |= 0x0020;
    }
    if keys.get(2).is_some_and(|(_, desc)| *desc) {
        flags |= 0x0040;
    }

    let col1 = keys.first().map_or(0, |(c, _)| *c);
    let col2 = keys.get(1).map_or(0, |(c, _)| *c);
    let col3 = keys.get(2).map_or(0, |(c, _)| *c);

    write_record_header(writer, 0x0090, 10)?;
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&col1.to_le_bytes())?;
    writer.write_all(&col2.to_le_bytes())?;
    writer.write_all(&col3.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?; // reserved

    Ok(())
}

pub fn write_sheet_protection<W: Write>(
    writer: &mut W,
    protect_objects: bool,
    protect_scenarios: bool,
    password_hash: Option<u16>,
) -> XlsResult<()> {
    write_record_header(writer, 0x0012, 2)?;
    writer.write_all(&0x0001u16.to_le_bytes())?;

    if protect_objects {
        write_record_header(writer, 0x0063, 2)?;
        writer.write_all(&0x0001u16.to_le_bytes())?;
    }

    if protect_scenarios {
        write_record_header(writer, 0x00DD, 2)?;
        writer.write_all(&0x0001u16.to_le_bytes())?;
    }

    if let Some(hash) = password_hash {
        write_record_header(writer, 0x0013, 2)?;
        writer.write_all(&hash.to_le_bytes())?;
    }

    Ok(())
}

fn encode_web_url_bytes(url: &str) -> Vec<u8> {
    // For URL hyperlinks we follow Apache POI's HyperlinkRecord layout:
    // the address is stored as a UTF-16LE string with a single trailing
    // NUL character and the length field contains the size in bytes
    // (2 bytes per character).
    let mut terminated = String::with_capacity(url.len().saturating_add(1));
    terminated.push_str(url);
    terminated.push('\0');

    let mut out = Vec::with_capacity(terminated.len().saturating_mul(2));
    for unit in terminated.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn write_hyperlink_web<W: Write>(
    writer: &mut W,
    row1: u16,
    row2: u16,
    col1: u16,
    col2: u16,
    url: &str,
) -> XlsResult<()> {
    if url.is_empty() {
        return Ok(());
    }

    // Constants taken from PhpSpreadsheet's writeUrlWeb implementation.
    const UNKNOWN1: [u8; 20] = [
        0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B, 0x02, 0x00, 0x00, 0x00,
    ];
    const UNKNOWN2: [u8; 16] = [
        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B,
    ];

    let url_bytes = encode_web_url_bytes(url);
    let url_len = u32::try_from(url_bytes.len()).map_err(|_| {
        XlsError::InvalidData("Hyperlink URL exceeds BIFF8 length limit".to_string())
    })?;

    // Base size (0x34) matches POI's HyperlinkRecord.getDataSize():
    //  - 8 bytes Ref8U (rwFirst, rwLast, colFirst, colLast)
    //  - 16 bytes GUID
    //  - 4 bytes streamVersion
    //  - 4 bytes linkOpts
    //  - 16 bytes URL moniker CLSID
    //  - 4 bytes address length (byte count)
    let data_len = 0x34u32.saturating_add(url_len);
    if data_len > u16::MAX as u32 {
        return Err(XlsError::InvalidData(
            "Hyperlink record exceeds BIFF8 length limit".to_string(),
        ));
    }

    write_record_header(writer, 0x01B8, data_len as u16)?;

    writer.write_all(&row1.to_le_bytes())?;
    writer.write_all(&row2.to_le_bytes())?;
    writer.write_all(&col1.to_le_bytes())?;
    writer.write_all(&col2.to_le_bytes())?;

    writer.write_all(&UNKNOWN1)?;

    // Option flags: 0x00000003 for standard URL/UNC hyperlink.
    writer.write_all(&0x0000_0003u32.to_le_bytes())?;

    writer.write_all(&UNKNOWN2)?;
    writer.write_all(&url_len.to_le_bytes())?;
    writer.write_all(&url_bytes)?;

    Ok(())
}

fn write_hyperlink_internal<W: Write>(
    writer: &mut W,
    row1: u16,
    row2: u16,
    col1: u16,
    col2: u16,
    url: &str,
) -> XlsResult<()> {
    if url.is_empty() {
        return Ok(());
    }

    const UNKNOWN1: [u8; 20] = [
        0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B, 0x02, 0x00, 0x00, 0x00,
    ];

    // Strip explicit internal: prefix if present.
    let target = url.strip_prefix("internal:").unwrap_or(url);

    // Append a single NUL terminator, then encode as UTF-16LE.
    let mut terminated = String::with_capacity(target.len().saturating_add(1));
    terminated.push_str(target);
    terminated.push('\0');

    let char_count = terminated.chars().count();
    let mut wide = Vec::with_capacity(char_count.saturating_mul(2));
    for unit in terminated.encode_utf16() {
        wide.extend_from_slice(&unit.to_le_bytes());
    }

    let url_len = u32::try_from(char_count)
        .map_err(|_| XlsError::InvalidData("Internal hyperlink target is too long".to_string()))?;

    let data_len = 0x24u32.saturating_add(u32::from(wide.len() as u16));
    if data_len > u16::MAX as u32 {
        return Err(XlsError::InvalidData(
            "Internal hyperlink record exceeds BIFF8 length limit".to_string(),
        ));
    }

    write_record_header(writer, 0x01B8, data_len as u16)?;

    writer.write_all(&row1.to_le_bytes())?;
    writer.write_all(&row2.to_le_bytes())?;
    writer.write_all(&col1.to_le_bytes())?;
    writer.write_all(&col2.to_le_bytes())?;

    writer.write_all(&UNKNOWN1)?;

    // Option flags: 0x00000008 for internal document reference.
    writer.write_all(&0x0000_0008u32.to_le_bytes())?;

    writer.write_all(&url_len.to_le_bytes())?;
    writer.write_all(&wide)?;

    Ok(())
}

/// Write HLINK (hyperlink) record for a single cell or cell range.
///
/// For now we support standard web/mail/ftp URLs and internal workbook
/// references. External file hyperlinks can be added later using the
/// more complex BIFF8 layout if required.
pub fn write_hyperlink<W: Write>(
    writer: &mut W,
    row1: u32,
    row2: u32,
    col1: u16,
    col2: u16,
    url: &str,
) -> XlsResult<()> {
    if row1 > u16::MAX as u32 || row2 > u16::MAX as u32 {
        return Err(XlsError::InvalidData(
            "Hyperlink row index must be <= 65535 for BIFF8".to_string(),
        ));
    }

    let r1 = row1 as u16;
    let r2 = row2 as u16;

    let trimmed = url.trim();
    if trimmed.is_empty() {
        return Ok(());
    }

    let is_web_like = trimmed.starts_with("http://")
        || trimmed.starts_with("https://")
        || trimmed.starts_with("ftp://")
        || trimmed.starts_with("mailto:");

    let is_internal = trimmed.starts_with("internal:")
        || (!is_web_like && trimmed.contains('!') && !trimmed.contains("://"));

    if is_internal {
        write_hyperlink_internal(writer, r1, r2, col1, col2, trimmed)
    } else {
        write_hyperlink_web(writer, r1, r2, col1, col2, trimmed)
    }
}

/// Write WINDOW2 record (Worksheet view settings)
///
/// Record type: 0x023E, Length: 18 (worksheet and macro sheet)
///
/// When `has_freeze_panes` is true, the FREEZE_PANES (0x0008) and
/// FREEZE_PANES_NO_SPLIT (0x0100) bits are set in the options field,
/// mirroring Apache POI's behaviour after `createFreezePane`.
#[allow(dead_code)]
pub fn write_window2<W: Write>(writer: &mut W, has_freeze_panes: bool) -> XlsResult<()> {
    write_record_header(writer, 0x023E, 18)?;

    // Base options value from POI's InternalSheet.createWindowTwo(): 0x06B6
    let mut options: u16 = 0x06B6;

    if has_freeze_panes {
        // Enable freeze panes and indicate that this is a frozen, not split,
        // window. Bits are defined in POI's WindowTwoRecord as:
        //  - 0x0008: freezePanes
        //  - 0x0100: freezePanesNoSplit
        options |= 0x0008 | 0x0100;
    }

    writer.write_all(&options.to_le_bytes())?;

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

/// Write an SCL record for a non-default worksheet zoom fraction.
pub fn write_scl<W: Write>(writer: &mut W, numerator: u16, denominator: u16) -> XlsResult<()> {
    if numerator == 0
        || denominator == 0
        || numerator > i16::MAX as u16
        || denominator > i16::MAX as u16
        || u32::from(numerator) * 10 < u32::from(denominator)
        || u32::from(numerator) > u32::from(denominator) * 4
    {
        return Err(XlsError::InvalidData(
            "SCL zoom fraction must be between 1/10 and 4 with positive terms".to_string(),
        ));
    }
    write_record_header(writer, 0x00A0, 4)?;
    writer.write_all(&numerator.to_le_bytes())?;
    writer.write_all(&denominator.to_le_bytes())?;
    Ok(())
}

pub fn write_window2_options<W: Write>(
    writer: &mut W,
    options: &crate::writer::view::XlsWorksheetViewOptions,
) -> XlsResult<()> {
    options.validate()?;
    let mut flags = 0u16;
    flags |= u16::from(options.show_formulas);
    flags |= u16::from(options.show_gridlines) << 1;
    flags |= u16::from(options.show_row_column_headers) << 2;
    if options
        .pane
        .is_some_and(|pane| pane.mode == crate::writer::view::XlsPaneMode::Frozen)
    {
        flags |= 0x0108;
    }
    flags |= u16::from(options.show_zero_values) << 4;
    flags |= u16::from(options.gridline_color_index.is_none()) << 5;
    flags |= u16::from(options.right_to_left) << 6;
    flags |= u16::from(options.show_outline_symbols) << 7;
    flags |= u16::from(options.selected) << 9;
    flags |= u16::from(options.displayed) << 10;
    flags |= u16::from(options.page_break_preview) << 11;
    write_record_header(writer, 0x023e, 18)?;
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&options.first_visible_row.to_le_bytes())?;
    writer.write_all(&u16::from(options.first_visible_column).to_le_bytes())?;
    writer.write_all(&options.gridline_color_index.unwrap_or(64).to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&options.page_break_zoom_percent.unwrap_or(0).to_le_bytes())?;
    writer.write_all(&options.normal_zoom_percent.unwrap_or(0).to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

pub fn write_pane_options<W: Write>(
    writer: &mut W,
    pane: &crate::writer::view::XlsWorksheetPaneOptions,
) -> XlsResult<()> {
    pane.validate()?;
    write_record_header(writer, 0x0041, 10)?;
    writer.write_all(&pane.horizontal_split.to_le_bytes())?;
    writer.write_all(&pane.vertical_split.to_le_bytes())?;
    writer.write_all(&pane.bottom_pane_top_row.to_le_bytes())?;
    writer.write_all(&u16::from(pane.right_pane_left_column).to_le_bytes())?;
    writer.write_all(&[pane.active_pane.code(), 0])?;
    Ok(())
}

pub fn write_selection_options<W: Write>(
    writer: &mut W,
    selection: &crate::writer::view::XlsWorksheetSelectionOptions,
) -> XlsResult<()> {
    if selection.ranges.is_empty()
        || selection.ranges.len() > crate::writer::view::MAX_SELECTION_RANGES
    {
        return Err(XlsError::InvalidData(
            "SELECTION range count must be 1..=1369".to_string(),
        ));
    }
    let payload_len = 9usize
        .checked_add(selection.ranges.len().checked_mul(6).ok_or_else(|| {
            XlsError::InvalidData("SELECTION payload length overflow".to_string())
        })?)
        .ok_or_else(|| XlsError::InvalidData("SELECTION payload length overflow".to_string()))?;
    write_record_header(
        writer,
        0x001d,
        u16::try_from(payload_len).map_err(|_| {
            XlsError::InvalidData("SELECTION payload exceeds BIFF8 limit".to_string())
        })?,
    )?;
    writer.write_all(&[selection.pane.code()])?;
    writer.write_all(&selection.active_row.to_le_bytes())?;
    writer.write_all(&u16::from(selection.active_column).to_le_bytes())?;
    writer.write_all(&selection.active_range_index.to_le_bytes())?;
    writer.write_all(&(selection.ranges.len() as u16).to_le_bytes())?;
    for range in &selection.ranges {
        writer.write_all(&range.first_row().to_le_bytes())?;
        writer.write_all(&range.last_row().to_le_bytes())?;
        writer.write_all(&[range.first_column(), range.last_column()])?;
    }
    Ok(())
}

/// Write the primary selection for a regular worksheet window.
#[allow(dead_code)]
pub fn write_default_selection<W: Write>(
    writer: &mut W,
    freeze_rows: u16,
    freeze_cols: u16,
) -> XlsResult<()> {
    if freeze_cols > 255 {
        return Err(XlsError::InvalidData(
            "SELECTION active column exceeds BIFF8 limit 255".to_string(),
        ));
    }
    let pane: u8 = match (freeze_cols > 0, freeze_rows > 0) {
        (true, true) => 0,
        (true, false) => 1,
        (false, true) => 2,
        (false, false) => 3,
    };
    write_record_header(writer, 0x001D, 15)?;
    writer.write_all(&[pane])?;
    writer.write_all(&freeze_rows.to_le_bytes())?;
    writer.write_all(&freeze_cols.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&1u16.to_le_bytes())?;
    writer.write_all(&freeze_rows.to_le_bytes())?;
    writer.write_all(&freeze_rows.to_le_bytes())?;
    writer.write_all(&[freeze_cols as u8, freeze_cols as u8])?;
    Ok(())
}

pub fn write_pivot_window2<W: Write>(writer: &mut W, selected: bool) -> XlsResult<()> {
    const FLAGS: u16 = 0x00B6;
    static DATA: &[u8] = &[
        0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x11, 0x00, 0x00,
        0x00,
    ];
    write_record_header(writer, 0x023E, 18)?;
    let flags = FLAGS | if selected { 0x0200 } else { 0 };
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(DATA)?;
    Ok(())
}

pub fn write_plv<W: Write>(writer: &mut W) -> XlsResult<()> {
    static DATA: &[u8] = &[
        0x8B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x12,
        0x00,
    ];
    write_record_header(writer, 0x088B, DATA.len() as u16)?;
    writer.write_all(DATA)?;
    Ok(())
}

pub fn write_selection<W: Write>(writer: &mut W) -> XlsResult<()> {
    static DATA: &[u8] = &[
        0x03, 0x0F, 0x00, 0x03, 0x00, 0x00, 0x00, 0x01, 0x00, 0x0F, 0x00, 0x0F, 0x00, 0x03, 0x03,
    ];
    write_record_header(writer, 0x001D, DATA.len() as u16)?;
    writer.write_all(DATA)?;
    Ok(())
}

pub fn write_phonetic_pr<W: Write>(writer: &mut W) -> XlsResult<()> {
    static DATA: &[u8] = &[0x17, 0x00, 0x37, 0x00, 0x00, 0x00];
    write_record_header(writer, 0x00EF, DATA.len() as u16)?;
    writer.write_all(DATA)?;
    Ok(())
}

pub fn write_sheet_ext<W: Write>(writer: &mut W) -> XlsResult<()> {
    static DATA: &[u8] = &[
        0x67, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
        0xFF, 0xFF, 0xFF, 0xFF, 0x03, 0x44, 0x00, 0x00,
    ];
    write_record_header(writer, 0x0867, DATA.len() as u16)?;
    writer.write_all(DATA)?;
    Ok(())
}

/// Write a PHONETICINFO record (MS-XLS 2.4.192) with phonetic defaults,
/// chunking long range lists into Continue records.
///
/// Record type: 0x00EF
pub fn write_phonetic_info<W: Write>(
    writer: &mut W,
    value: &crate::XlsPhoneticInfo,
) -> XlsResult<()> {
    const MAX_RECORD_PAYLOAD: usize = 8_224;
    const CONTINUE_RECORD_TYPE: u16 = 0x003C;
    let payload = value.to_payload();
    let first_chunk = payload.len().min(MAX_RECORD_PAYLOAD);
    write_record_header(
        writer,
        crate::phonetic_info::PHONETIC_INFO_RECORD_TYPE,
        first_chunk as u16,
    )?;
    writer.write_all(&payload[..first_chunk])?;
    for chunk in payload[first_chunk..].chunks(MAX_RECORD_PAYLOAD) {
        write_record_header(writer, CONTINUE_RECORD_TYPE, chunk.len() as u16)?;
        writer.write_all(chunk)?;
    }
    Ok(())
}

/// Write a SHEETEXT record (MS-XLS 2.4.259) carrying a sheet tab color.
///
/// Record type: 0x0862
pub fn write_sheet_ext_tab_color<W: Write>(writer: &mut W, tab_color: u8) -> XlsResult<()> {
    let payload = crate::sheet_ext::XlsSheetExt::from_tab_color(Some(tab_color)).to_payload();
    write_record_header(
        writer,
        crate::sheet_ext::SHEET_EXT_RECORD_TYPE,
        payload.len() as u16,
    )?;
    writer.write_all(&payload)?;
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

/// Write ROW record (row metrics including height and hidden flag).
///
/// Record type: 0x0208, Length: 16
///
/// The height is stored in twips (1/20 of a point) as per MS-XLS
/// and Apache POI's `RowRecord` implementation.
pub fn write_row<W: Write>(
    writer: &mut W,
    row_index: u32,
    first_col: u16,
    last_col_plus1: u16,
    height: u16,
    hidden: bool,
) -> XlsResult<()> {
    let row_u16 = u16::try_from(row_index).map_err(|_| {
        XlsError::InvalidData(format!(
            "Row index {} exceeds BIFF8 limit 65535 for ROW record",
            row_index
        ))
    })?;

    write_record_header(writer, 0x0208, 16)?;

    // Row number
    writer.write_all(&row_u16.to_le_bytes())?;

    // First and last used column indices. A value of 0 for both is
    // accepted by Excel for empty rows (mirrors POI's `setEmpty`).
    writer.write_all(&first_col.to_le_bytes())?;
    writer.write_all(&last_col_plus1.to_le_bytes())?;

    // Row height in twips
    writer.write_all(&height.to_le_bytes())?;

    // Optimization hint and reserved fields: keep both at zero, as
    // POI does for generated sheets.
    writer.write_all(&0u16.to_le_bytes())?; // optimize
    writer.write_all(&0u16.to_le_bytes())?; // reserved

    // Option flags: always set bit 8 (0x0100) as in POI's
    // OPTION_BITS_ALWAYS_SET, and toggle the zeroHeight bit (0x0020)
    // when the row is hidden. When a custom height is used
    // (height != 0x00FF), also set the badFontHeight bit (0x0040),
    // mirroring HSSFRow.setHeightInPoints and RowRecord.
    let mut option_flags: u16 = 0x0100;
    if hidden {
        option_flags |= 0x0020;
    }
    if height != 0x00FF && height != 0x0116 {
        option_flags |= 0x0040;
    }
    writer.write_all(&option_flags.to_le_bytes())?;

    // Secondary option flags, including the XF index and border bits.
    // For now we leave this at POI's default of 0x000F.
    writer.write_all(&0x000Fu16.to_le_bytes())?;

    Ok(())
}

pub fn write_mergedcells<W, I>(writer: &mut W, ranges: I) -> XlsResult<()>
where
    W: Write,
    I: IntoIterator<Item = (u16, u16, u8, u8)>,
{
    const MAX_MERGED_REGIONS: usize = 1027;

    let mut chunk: Vec<(u16, u16, u8, u8)> = Vec::new();

    for (first_row, last_row, first_col, last_col) in ranges {
        if first_row > last_row || first_col > last_col {
            return Err(XlsError::InvalidCellReference(
                "MERGEDCELLS contains a reversed range".to_string(),
            ));
        }
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
    ranges: &[(u16, u16, u8, u8)],
) -> XlsResult<()> {
    debug_assert!(!ranges.is_empty());
    debug_assert!(ranges.len() <= 1027);

    let count = u16::try_from(ranges.len())
        .map_err(|_| XlsError::InvalidData("MERGEDCELLS range count exceeds u16".to_string()))?;
    let data_len: u16 = 2u16 + count.saturating_mul(8);

    write_record_header(writer, 0x00E5, data_len)?;
    writer.write_all(&count.to_le_bytes())?;

    for &(first_row, last_row, first_col, last_col) in ranges {
        writer.write_all(&first_row.to_le_bytes())?;
        writer.write_all(&last_row.to_le_bytes())?;
        writer.write_all(&u16::from(first_col).to_le_bytes())?;
        writer.write_all(&u16::from(last_col).to_le_bytes())?;
    }

    Ok(())
}

#[cfg(test)]
mod view_round_trip_tests {
    use super::*;
    use crate::view::{
        PANE_RECORD_TYPE, SCL_RECORD_TYPE, SELECTION_RECORD_TYPE, ViewCollector,
        WINDOW2_RECORD_TYPE,
    };

    fn payload(record: &[u8]) -> &[u8] {
        let length = usize::from(u16::from_le_bytes([record[2], record[3]]));
        &record[4..4 + length]
    }

    #[test]
    fn writes_view_records_that_round_trip_through_reader() {
        let mut window = Vec::new();
        let mut scl = Vec::new();
        let mut pane = Vec::new();
        let mut selection = Vec::new();
        write_window2(&mut window, true).unwrap();
        write_scl(&mut scl, 3, 4).unwrap();
        write_pane(&mut pane, 7, 5).unwrap();
        write_default_selection(&mut selection, 7, 5).unwrap();

        let mut collector = ViewCollector::new();
        collector
            .feed_record(WINDOW2_RECORD_TYPE, payload(&window))
            .unwrap();
        collector
            .feed_record(SCL_RECORD_TYPE, payload(&scl))
            .unwrap();
        collector
            .feed_record(PANE_RECORD_TYPE, payload(&pane))
            .unwrap();
        collector
            .feed_record(SELECTION_RECORD_TYPE, payload(&selection))
            .unwrap();
        let views = collector.finish().unwrap();
        let view = &views[0];
        assert_eq!(view.zoom_fraction(), Some((3, 4)));
        assert_eq!(
            view.pane().unwrap().active_pane(),
            crate::view::XlsPaneType::LowerRight
        );
        assert_eq!(view.selections()[0].active_row(), 7);
        assert_eq!(view.selections()[0].active_column(), 5);
    }

    #[test]
    fn rejects_invalid_writer_view_bounds() {
        assert!(write_scl(&mut Vec::new(), 0, 1).is_err());
        assert!(write_scl(&mut Vec::new(), 5, 1).is_err());
        assert!(write_pane(&mut Vec::new(), 1, 256).is_err());
        assert!(write_default_selection(&mut Vec::new(), 0, 256).is_err());
    }
}
