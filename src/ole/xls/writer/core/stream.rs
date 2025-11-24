use std::collections::HashMap;

use crate::ole::xls::writer::biff;
use crate::ole::xls::writer::formatting::FormattingManager;
use crate::ole::xls::{XlsError, XlsResult};

use super::named_range::XlsDefinedName as InternalDefinedName;
use super::worksheet::WritableWorksheet;
use super::{XlsCellValue, XlsWorkbookProtection};

#[allow(clippy::too_many_arguments)] // TODO: Refactor this function to accept a struct
pub(crate) fn generate_workbook_stream(
    use_1904_dates: bool,
    fmt: &FormattingManager,
    defined_names: &[InternalDefinedName],
    shared_strings: &[String],
    sst_total: u32,
    workbook_protection: Option<XlsWorkbookProtection>,
    worksheets: &[WritableWorksheet],
    string_map: &HashMap<String, u32>,
) -> XlsResult<Vec<u8>> {
    let mut stream = Vec::new();

    // === Workbook Globals ===

    // BOF record (workbook)
    biff::write_bof(&mut stream, 0x0005)?;

    // CodePage record - BIFF8 requires Unicode codepage 1200 (0x04B0)
    biff::write_codepage(&mut stream, 0x04B0)?;

    // Date1904 record
    biff::write_date1904(&mut stream, use_1904_dates)?;

    // Window1 record (workbook window properties)
    biff::write_window1(&mut stream)?;

    if let Some(protection) = workbook_protection {
        biff::write_workbook_protection(
            &mut stream,
            protection.protect_structure,
            protection.protect_windows,
            protection.password_hash,
        )?;
    }

    // Write minimal formatting tables so XF index 0 is valid.
    // Order mirrors Apache POI's workbook creation:
    //  - FONT records
    //  - FORMAT records (built-in 0..7 + custom)
    //  - XF records (style and cell formats)
    fmt.write_fonts(&mut stream)?;
    fmt.write_number_formats(&mut stream)?;
    fmt.write_formats(&mut stream)?;

    // Built-in STYLE records and UseSelFS flag to align with Excel / POI
    // defaults. This makes standard cell styles (Normal, Currency, Percent,
    // etc.) visible to Excel even though we currently only use the default
    // cell XF (index 15) for all cells.
    biff::write_builtin_styles(&mut stream)?;
    biff::write_usesel_fs(&mut stream)?;

    // BoundSheet8 records (one per worksheet)
    // We need to calculate positions, so we'll write them after we know the sizes
    let mut boundsheet_positions = Vec::new();
    for worksheet in worksheets {
        // Placeholder - we'll update positions later
        boundsheet_positions.push(stream.len());
        biff::write_boundsheet(&mut stream, 0, &worksheet.name)?;
    }

    // Internal SUPBOOK / EXTERNSHEET records are required for 3D
    // references used by defined names (NameParsedFormula). We keep
    // the model minimal by generating a single internal SUPBOOK and
    // one XTI entry per worksheet.
    if !defined_names.is_empty() && !worksheets.is_empty() {
        let sheet_count = u16::try_from(worksheets.len()).unwrap_or(u16::MAX);
        biff::write_supbook_internal(&mut stream, sheet_count)?;
        biff::write_externsheet_internal(&mut stream, sheet_count)?;
    }

    // NAME (Lbl) records for workbook- and sheet-scoped defined names.
    // These are stored in the globals substream and reference cell
    // areas using BIFF8 formula tokens.
    for defined_name in defined_names {
        let rgce = defined_name.to_biff_formula()?;
        biff::write_name(&mut stream, defined_name, &rgce)?;
    }

    // SST record (shared string table)
    if !shared_strings.is_empty() {
        biff::write_sst(&mut stream, shared_strings, sst_total)?;
    }

    // EOF record (end of workbook globals)
    biff::write_eof(&mut stream)?;

    // === Worksheets ===

    // Track actual worksheet positions
    let mut actual_positions = Vec::new();

    for worksheet in worksheets {
        // Record the position of this worksheet's BOF
        let worksheet_pos = stream.len() as u32;
        actual_positions.push(worksheet_pos);

        // BOF record (worksheet)
        biff::write_bof(&mut stream, 0x0010)?;

        // DIMENSIONS record
        biff::write_dimensions(
            &mut stream,
            worksheet.first_row,
            worksheet.last_row,
            worksheet.first_col,
            worksheet.last_col,
        )?;

        // Required sheet records for worksheet substream per MS-XLS.
        //
        // Apache POI writes WINDOW2 first and then (optionally) PANE
        // immediately afterwards when freeze panes are configured. We
        // mirror that ordering here to avoid Excel interpreting the
        // pane as a generic split window.
        biff::write_wsbool(&mut stream)?;
        let has_freeze_panes = worksheet.freeze_panes.is_some();
        biff::write_window2(&mut stream, has_freeze_panes)?;

        if let Some(panes) = worksheet.freeze_panes {
            biff::write_pane(&mut stream, panes.freeze_rows, panes.freeze_cols)?;
        }

        if let Some(protection) = worksheet.sheet_protection {
            biff::write_sheet_protection(
                &mut stream,
                protection.protect_objects,
                protection.protect_scenarios,
                protection.password_hash,
            )?;
        }

        if let Some(af) = worksheet.auto_filter {
            let _row_span = af.last_row.saturating_sub(af.first_row).saturating_add(1);
            let width = u32::from(af.last_col)
                .saturating_sub(u32::from(af.first_col))
                .saturating_add(1);
            let c_entries = u16::try_from(width).map_err(|_| {
                XlsError::InvalidData(
                    "set_auto_filter: auto-filter column span exceeds BIFF8 limit".to_string(),
                )
            })?;
            biff::write_autofilterinfo(&mut stream, c_entries)?;
        }

        // Column width / hidden state via COLINFO records.
        if !worksheet.column_widths.is_empty() || !worksheet.hidden_columns.is_empty() {
            use std::collections::BTreeSet;

            let mut cols = BTreeSet::<u16>::new();
            cols.extend(worksheet.column_widths.keys().copied());
            cols.extend(worksheet.hidden_columns.iter().copied());

            for col in cols {
                let width_units = worksheet
                    .column_widths
                    .get(&col)
                    .copied()
                    // Default matches POI's ColumnInfoRecord constructor.
                    .unwrap_or(2275u16);
                let hidden = worksheet.hidden_columns.contains(&col);
                biff::write_colinfo(&mut stream, col, col, width_units, hidden)?;
            }
        }

        // Pre-compute row spans (first/last used column per row) for ROW records.
        use std::collections::HashMap as StdHashMap;
        let mut row_spans: StdHashMap<u32, (u16, u16)> = StdHashMap::new();

        for &(row, col) in worksheet.cells.keys() {
            let entry = row_spans.entry(row).or_insert((col, col.saturating_add(1)));
            if col < entry.0 {
                entry.0 = col;
            }
            if col.saturating_add(1) > entry.1 {
                entry.1 = col.saturating_add(1);
            }
        }

        // ROW records for rows with custom height or hidden state.
        if !worksheet.row_heights.is_empty() || !worksheet.hidden_rows.is_empty() {
            use std::collections::BTreeSet;

            let mut rows = BTreeSet::<u32>::new();
            rows.extend(worksheet.row_heights.keys().copied());
            rows.extend(worksheet.hidden_rows.iter().copied());

            for row in rows {
                let (first_col, last_col_plus1) = row_spans.get(&row).copied().unwrap_or((0, 0));
                let height = worksheet
                    .row_heights
                    .get(&row)
                    // Default height matches POI's RowRecord constructor (0x00FF).
                    .copied()
                    .unwrap_or(0x00FFu16);
                let hidden = worksheet.hidden_rows.contains(&row);
                biff::write_row(&mut stream, row, first_col, last_col_plus1, height, hidden)?;
            }
        }

        // Cell records (sorted by row, then column)
        let mut sorted_cells: Vec<_> = worksheet.cells.iter().collect();
        sorted_cells.sort_by_key(|(k, _)| *k);

        for ((row, col), cell) in sorted_cells {
            let xf_index = fmt.cell_xf_index_for(cell.format_idx);
            match &cell.value {
                XlsCellValue::Number(value) => {
                    biff::write_number(&mut stream, *row, *col, xf_index, *value)?;
                },
                XlsCellValue::String(s) => {
                    let sst_index = *string_map.get(s).unwrap();
                    biff::write_labelsst(&mut stream, *row, *col, xf_index, sst_index)?;
                },
                XlsCellValue::Boolean(value) => {
                    biff::write_boolerr(&mut stream, *row, *col, xf_index, *value)?;
                },
                XlsCellValue::Formula(_formula) => {
                    // Formula tokenization not yet implemented
                    // Write as blank cell for now
                    // Future enhancement: Parse formula to RPN tokens and write FORMULA record
                },
                XlsCellValue::Blank => {
                    // Skip blank cells
                },
            }
        }

        // Hyperlink records for cells or ranges.
        for hyperlink in &worksheet.hyperlinks {
            biff::write_hyperlink(
                &mut stream,
                hyperlink.first_row,
                hyperlink.last_row,
                hyperlink.first_col,
                hyperlink.last_col,
                &hyperlink.url,
            )?;
        }

        if !worksheet.merged_ranges.is_empty() {
            biff::write_mergedcells(
                &mut stream,
                worksheet
                    .merged_ranges
                    .iter()
                    .map(|r| (r.first_row, r.last_row, r.first_col, r.last_col)),
            )?;
        }

        if !worksheet.data_validations.is_empty() {
            let dv_count = worksheet.data_validations.len() as u32;
            biff::write_dval(&mut stream, dv_count)?;

            for dv in &worksheet.data_validations {
                let payload = dv.validation_type.to_biff_payload()?;

                let ranges = [(dv.first_row, dv.last_row, dv.first_col, dv.last_col)];

                let formula1 = payload.formula1.as_deref();
                let formula2 = payload.formula2.as_deref();

                let dv_config = biff::DvConfig {
                    data_type: payload.data_type,
                    operator: payload.operator,
                    error_style: 0, // errorStyle: STOP
                    empty_cell_allowed: true,
                    suppress_dropdown_arrow: false,
                    is_explicit_list_formula: payload.is_explicit_list,
                    show_prompt_on_cell_selected: dv.show_input_message,
                    prompt_title: dv.input_title.as_deref(),
                    prompt_text: dv.input_message.as_deref(),
                    show_error_on_invalid_value: dv.show_error_alert,
                    error_title: dv.error_title.as_deref(),
                    error_text: dv.error_message.as_deref(),
                    formula1,
                    formula2,
                };

                biff::write_dv(&mut stream, &dv_config, &ranges)?;
            }
        }

        if !worksheet.conditional_formats.is_empty() {
            for cf in &worksheet.conditional_formats {
                let ranges = [(cf.first_row, cf.last_row, cf.first_col, cf.last_col)];

                // One CFHEADER per rule with a single region keeps the
                // implementation simple and matches Excel's expectations.
                biff::write_cfheader(&mut stream, &ranges, 1)?;

                let (condition_type, comparison_op, formula1, formula2) =
                    cf.format_type.to_biff_payload()?;

                biff::write_cfrule(
                    &mut stream,
                    condition_type,
                    comparison_op,
                    &formula1,
                    &formula2,
                    cf.to_biff_pattern(),
                )?;
            }
        }

        // EOF record (end of worksheet)
        biff::write_eof(&mut stream)?;
    }

    // Go back and update BoundSheet positions
    for (i, &pos) in actual_positions.iter().enumerate() {
        let boundsheet_pos = boundsheet_positions[i];
        // Position field starts at offset 4 in the record (after header)
        let pos_offset = boundsheet_pos + 4;
        stream[pos_offset..pos_offset + 4].copy_from_slice(&pos.to_le_bytes());
    }

    Ok(stream)
}
