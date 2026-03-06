use std::collections::HashMap;

use crate::ole::xls::writer::biff;
use crate::ole::xls::writer::formatting::FormattingManager;
use crate::ole::xls::{XlsError, XlsResult};

use super::named_range::XlsDefinedName as InternalDefinedName;
use super::worksheet::WritableWorksheet;
use super::{XlsCellValue, XlsWorkbookProtection};

const DEFAULT_WRITE_ACCESS_USER: &str = "litchi";
const DEFAULT_RECALC_ENGINE_ID: u32 = 0x000E_EA35;
const DEFAULT_FUNCTION_GROUP_COUNT: u16 = 17;
const DEFAULT_COUNTRY_CODE: u16 = 1;

/// Result of generating the workbook: the Workbook stream plus any pivot
/// cache storage streams that must be placed in `_SX_DB_CUR/nnnn`.
pub(crate) struct WorkbookStreams {
    /// The main Workbook BIFF stream.
    pub workbook: Vec<u8>,
    /// Pivot cache streams: `(stream_id, data)`.  Each goes into
    /// `_SX_DB_CUR/{stream_id:04X}`.
    pub pivot_caches: Vec<(u16, Vec<u8>)>,
}

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
) -> XlsResult<WorkbookStreams> {
    let mut stream = Vec::new();
    let has_pivot_tables = worksheets.iter().any(|ws| !ws.pivot_tables.is_empty());
    let sheet_count = u16::try_from(worksheets.len()).unwrap_or(u16::MAX);
    let (protect_structure, protect_windows, password_hash) = workbook_protection
        .map(|protection| {
            (
                protection.protect_structure,
                protection.protect_windows,
                protection.password_hash.unwrap_or(0),
            )
        })
        .unwrap_or((false, false, 0));

    // === Workbook Globals ===

    // BOF record (workbook)
    biff::write_bof(&mut stream, 0x0005)?;

    biff::write_interface_hdr(&mut stream, 0x04B0)?;
    biff::write_mms(&mut stream)?;
    biff::write_interface_end(&mut stream)?;
    biff::write_write_access(&mut stream, DEFAULT_WRITE_ACCESS_USER)?;

    // CodePage record - BIFF8 requires Unicode codepage 1200 (0x04B0)
    biff::write_codepage(&mut stream, 0x04B0)?;

    biff::write_dsf(&mut stream, false)?;
    if has_pivot_tables {
        biff::write_excel9_file(&mut stream)?;
    }
    biff::write_tab_id(&mut stream, sheet_count)?;
    biff::write_fn_group_count(&mut stream, DEFAULT_FUNCTION_GROUP_COUNT)?;
    biff::write_window_protect(&mut stream, protect_windows)?;
    biff::write_protect(&mut stream, protect_structure || protect_windows)?;
    biff::write_password(&mut stream, password_hash)?;
    biff::write_protection_rev4(&mut stream, false)?;
    biff::write_password_rev4(&mut stream, 0)?;

    // Window1 record (workbook window properties)
    biff::write_window1(&mut stream)?;

    biff::write_backup(&mut stream, false)?;
    biff::write_hide_obj(&mut stream, 0)?;
    biff::write_date1904(&mut stream, use_1904_dates)?;
    biff::write_precision(&mut stream, true)?;
    biff::write_refresh_all(&mut stream, false)?;
    biff::write_book_bool(&mut stream, false)?;

    // Write minimal formatting tables so XF index 0 is valid.
    // Order mirrors Apache POI's workbook creation:
    //  - FONT records
    //  - FORMAT records (built-in 0..7 + custom)
    //  - XF records (style and cell formats)
    fmt.write_fonts(&mut stream)?;
    fmt.write_number_formats(&mut stream)?;
    fmt.write_formats(&mut stream)?;
    if has_pivot_tables {
        biff::write_pivot_xfext_block(&mut stream)?;
    }

    // Built-in STYLE records to align with Excel / POI
    // defaults. This makes standard cell styles (Normal, Currency, Percent,
    // etc.) visible to Excel even though we currently only use the default
    // cell XF (index 15) for all cells.
    biff::write_builtin_styles(&mut stream)?;
    if has_pivot_tables {
        biff::write_table_styles(&mut stream)?;
    }

    // BoundSheet8 records are emitted later so pivot cache definitions can stay
    // adjacent to the workbook formatting block, matching Excel's globals order.
    let mut boundsheet_positions = Vec::new();

    // Internal SUPBOOK / EXTERNSHEET records are required for 3D
    // references used by defined names (NameParsedFormula) and pivot caches.
    if (!defined_names.is_empty() || has_pivot_tables) && !worksheets.is_empty() {
        let externsheet_mode = if defined_names.is_empty() {
            biff::ExternSheetMode::WorkbookWide
        } else {
            biff::ExternSheetMode::PerSheet
        };
        biff::write_supbook_internal(&mut stream, sheet_count)?;
        biff::write_externsheet_internal(&mut stream, sheet_count, externsheet_mode)?;
    }

    // NAME (Lbl) records for workbook- and sheet-scoped defined names.
    // These are stored in the globals substream and reference cell
    // areas using BIFF8 formula tokens.
    for defined_name in defined_names {
        let rgce = defined_name.to_biff_formula()?;
        biff::write_name(&mut stream, defined_name, &rgce)?;
    }

    // Pivot table globals: PIVOTCACHEDEFINITION records.
    //
    // Per MS-XLS §2.1.7.20.3 globals ABNF:
    //   PIVOTCACHEDEFINITION = SXStreamID [SXVS] [SXSRC]
    //   SXSRC = DREF (= DConRef for worksheet sources)
    //
    // The actual cache data (SXDB + SXFDB) goes in a *separate* OLE
    // storage `_SX_DB_CUR/nnnn`, NOT in the Workbook stream.
    let mut pivot_caches: Vec<(u16, Vec<u8>)> = Vec::new();
    let has_any_page_fields: bool;
    {
        // Collect all pivot tables across worksheets.
        let all_pts: Vec<&super::worksheet::WritablePivotTable> = worksheets
            .iter()
            .flat_map(|ws| ws.pivot_tables.iter())
            .collect();
        has_any_page_fields = all_pts.iter().any(|pt| !pt.page_entries.is_empty());

        for (idx, pt) in all_pts.iter().enumerate() {
            // LO uses 1-based IDs: maPCInfo.mnStrmId = nListIdx + 1
            let id = (idx + 1) as u16;

            // PIVOTCACHEDEFINITION in globals: SxStreamID + SXVS + DCONREF
            biff::write_sx_stream_id(&mut stream, id)?;
            biff::write_sxvs(&mut stream, pt.source_type)?;
            biff::write_dconref(
                &mut stream,
                pt.source_first_row,
                pt.source_last_row,
                pt.source_first_col as u8,
                pt.source_last_col as u8,
                &pt.source_sheet_name,
            )?;
            biff::write_pivot_cache_sxaddl_block(&mut stream)?;

            // Build per-field cache info from the dedicated cache_items lists.
            // cache_name is the source column header (SXFDB name).
            // cache_items are the unique source data values (SXSTRING records).
            let field_item_refs: Vec<Vec<&str>> = pt
                .fields
                .iter()
                .map(|f| f.cache_items.iter().map(String::as_str).collect())
                .collect();

            // Count unique numeric values per numeric field from source_data.
            let unique_numeric_counts: Vec<u16> = pt
                .fields
                .iter()
                .enumerate()
                .map(|(fi, f)| {
                    if !f.is_numeric {
                        return 0;
                    }
                    let mut vals: Vec<u64> = pt
                        .source_data
                        .iter()
                        .filter_map(|row| {
                            row.get(fi).and_then(|v| match v {
                                super::PivotCacheValue::Number(n) => Some(n.to_bits()),
                                _ => None,
                            })
                        })
                        .collect();
                    vals.sort_unstable();
                    vals.dedup();
                    vals.len() as u16
                })
                .collect();

            let cache_fields: Vec<biff::PivotCacheFieldInfo<'_>> = pt
                .fields
                .iter()
                .zip(field_item_refs.iter())
                .zip(unique_numeric_counts.iter())
                .map(|((f, items), &uniq_count)| biff::PivotCacheFieldInfo {
                    name: &f.cache_name,
                    items: items.as_slice(),
                    is_numeric: f.is_numeric,
                    unique_numeric_count: uniq_count,
                })
                .collect();

            // Build source rows: split each PivotCacheValue row into
            // string_indices (for SXDBB) and numeric_values (for SXNUM).
            let num_string_fields = pt.fields.iter().filter(|f| !f.is_numeric).count();
            let num_numeric_fields = pt.fields.iter().filter(|f| f.is_numeric).count();
            let mut row_string_indices: Vec<Vec<u8>> = Vec::with_capacity(pt.source_data.len());
            let mut row_numeric_values: Vec<Vec<f64>> = Vec::with_capacity(pt.source_data.len());
            for row in &pt.source_data {
                let mut si = Vec::with_capacity(num_string_fields);
                let mut nv = Vec::with_capacity(num_numeric_fields);
                for (fi, val) in row.iter().enumerate() {
                    let is_num = pt.fields.get(fi).is_some_and(|f| f.is_numeric);
                    match val {
                        super::PivotCacheValue::StringIndex(idx) if !is_num => si.push(*idx),
                        super::PivotCacheValue::Number(v) if is_num => nv.push(*v),
                        _ => {}, // type mismatch — skip
                    }
                }
                row_string_indices.push(si);
                row_numeric_values.push(nv);
            }
            let source_rows: Vec<biff::PivotCacheSourceRow<'_>> = row_string_indices
                .iter()
                .zip(row_numeric_values.iter())
                .map(|(si, nv)| biff::PivotCacheSourceRow {
                    string_indices: si.as_slice(),
                    numeric_values: nv.as_slice(),
                })
                .collect();

            let record_count = pt.source_last_row.saturating_sub(pt.source_first_row) as u32;
            let cache_stream = biff::generate_pivot_cache_stream(&biff::PivotCacheStreamInfo {
                stream_id: id,
                record_count,
                fields: &cache_fields,
                source_rows: &source_rows,
            })?;
            pivot_caches.push((id, cache_stream));
        }
    }

    biff::write_usesel_fs(&mut stream)?;

    for worksheet in worksheets {
        boundsheet_positions.push(stream.len());
        biff::write_boundsheet(&mut stream, 0, &worksheet.name)?;
    }

    if has_pivot_tables {
        biff::write_compress_pictures(&mut stream)?;
        biff::write_compat12(&mut stream)?;
    }

    // MsoDrawingGroup (0x00EB) — if we have page fields, Excel expects a drawing group
    if has_any_page_fields {
        biff::write_mso_drawing_group(&mut stream)?;
    }

    biff::write_country(&mut stream, DEFAULT_COUNTRY_CODE, DEFAULT_COUNTRY_CODE)?;

    // SST record (shared string table)
    if !shared_strings.is_empty() {
        biff::write_sst(&mut stream, shared_strings, sst_total)?;
    }

    if !has_pivot_tables {
        biff::write_excel9_file(&mut stream)?;
    }
    biff::write_recalc_id(&mut stream, DEFAULT_RECALC_ENGINE_ID)?;

    // EOF record (end of workbook globals)
    biff::write_eof(&mut stream)?;

    // === Worksheets ===

    // Track actual worksheet positions
    let mut actual_positions = Vec::new();

    for worksheet in worksheets {
        // Record the position of this worksheet's BOF
        let worksheet_pos = stream.len() as u32;
        actual_positions.push(worksheet_pos);

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

        let pivot_first_col = if !worksheet.pivot_tables.is_empty() {
            worksheet
                .pivot_tables
                .iter()
                .map(|pt| pt.first_col)
                .min()
                .unwrap_or(0)
        } else {
            0
        };
        let pivot_last_col_plus1 = if !worksheet.pivot_tables.is_empty() {
            worksheet
                .pivot_tables
                .iter()
                .map(|pt| pt.last_col.saturating_add(1))
                .max()
                .unwrap_or(0)
        } else {
            0
        };

        let emitted_rows: Vec<u32> = {
            use std::collections::BTreeSet;

            let mut rows = BTreeSet::<u32>::new();
            rows.extend(worksheet.row_heights.keys().copied());
            rows.extend(worksheet.hidden_rows.iter().copied());
            if !worksheet.pivot_tables.is_empty() {
                rows.extend(row_spans.keys().copied());
            }
            rows.into_iter().collect()
        };
        let pivot_first_used_row = row_spans.keys().min().copied().unwrap_or(0);
        let pivot_last_used_row_plus1 = row_spans
            .keys()
            .max()
            .copied()
            .map(|row| row.saturating_add(1))
            .unwrap_or(0);
        let pivot_row_block_count = emitted_rows.len().div_ceil(32);

        // BOF record (worksheet)
        biff::write_bof(&mut stream, 0x0010)?;

        let pivot_index_record_pos =
            if !worksheet.pivot_tables.is_empty() && pivot_row_block_count > 0 {
                let index_record_pos = stream.len();
                let zero_dbcells = vec![0u32; pivot_row_block_count];
                biff::write_index(
                    &mut stream,
                    pivot_first_used_row,
                    pivot_last_used_row_plus1,
                    0,
                    &zero_dbcells,
                )?;
                Some(index_record_pos)
            } else {
                None
            };

        if worksheet.pivot_tables.is_empty() {
            biff::write_dimensions(
                &mut stream,
                worksheet.first_row,
                worksheet.last_row,
                worksheet.first_col,
                worksheet.last_col,
            )?;
        } else {
            biff::write_pivot_sheet_preamble(&mut stream)?;
        }

        // Required sheet records for worksheet substream per MS-XLS.
        //
        // Apache POI writes WINDOW2 first and then (optionally) PANE
        // immediately afterwards when freeze panes are configured. We
        // mirror that ordering here to avoid Excel interpreting the
        // pane as a generic split window.
        if worksheet.pivot_tables.is_empty() {
            biff::write_wsbool(&mut stream)?;
        }
        let has_freeze_panes = worksheet.freeze_panes.is_some();
        if worksheet.pivot_tables.is_empty() {
            biff::write_window2(&mut stream, has_freeze_panes)?;
        }

        // MsoDrawing (0x00EC) for non-pivot sheets
        if has_any_page_fields && worksheet.pivot_tables.is_empty() {
            biff::write_mso_drawing_sheet1(&mut stream)?;
        }

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

            // Write per-column AUTOFILTER records with filter conditions.
            for col_def in &worksheet.auto_filter_columns {
                biff::write_autofilter(
                    &mut stream,
                    col_def.column_index,
                    col_def.join_or,
                    false, // is_simple
                    false, // is_top10
                    false, // hide_arrow
                    &col_def.condition1,
                    &col_def.condition2,
                )?;
            }
        }

        // SORT record (if configured).
        if let Some(ref sort) = worksheet.sort_config {
            biff::write_sort(
                &mut stream,
                sort.case_sensitive,
                sort.sort_by_columns,
                &sort.keys,
            )?;
        }

        // Column width / hidden state via COLINFO records.
        let pivot_def_col_width_pos = if !worksheet.pivot_tables.is_empty() {
            let pos = stream.len() as u32;
            biff::write_def_col_width(&mut stream, 8)?;
            Some(pos)
        } else {
            None
        };

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
                if worksheet.pivot_tables.is_empty() {
                    biff::write_colinfo(&mut stream, col, col, width_units, hidden)?;
                } else {
                    biff::write_pivot_colinfo(&mut stream, col, col, width_units)?;
                }
            }
        }

        if !worksheet.pivot_tables.is_empty() {
            biff::write_dimensions(
                &mut stream,
                worksheet.first_row,
                worksheet.last_row,
                worksheet.first_col,
                worksheet.last_col,
            )?;
        }

        // ROW records for rows with custom height or hidden state.
        // Pivot worksheets also emit ROW records for used rows even when the
        // height is default, which appears to be part of Excel's expected
        // substream scaffolding for page-field dropdowns.
        let mut row_record_positions = StdHashMap::<u32, u32>::new();
        if !emitted_rows.is_empty() {
            for row in &emitted_rows {
                row_record_positions.insert(*row, stream.len() as u32);
                let (mut first_col, mut last_col_plus1) =
                    row_spans.get(row).copied().unwrap_or((0, 0));
                if !worksheet.pivot_tables.is_empty() {
                    first_col = first_col.min(pivot_first_col);
                    last_col_plus1 = last_col_plus1.max(pivot_last_col_plus1);
                }
                let height = worksheet
                    .row_heights
                    .get(row)
                    // Excel-authored pivot sheets use the default row height
                    // stored in DEFAULTROWHEIGHT (0x0116) for emitted ROW records.
                    .copied()
                    .unwrap_or(if !worksheet.pivot_tables.is_empty() {
                        0x0116u16
                    } else {
                        0x00FFu16
                    });
                let hidden = worksheet.hidden_rows.contains(row);
                biff::write_row(&mut stream, *row, first_col, last_col_plus1, height, hidden)?;
            }
        }

        // Cell records (sorted by row, then column)
        let mut sorted_cells: Vec<_> = worksheet.cells.iter().collect();
        sorted_cells.sort_by_key(|(k, _)| *k);

        let mut row_first_cell_positions = StdHashMap::<u32, u32>::new();
        let pivot_xf_indices = fmt.pivot_xf_indices();

        let mut cell_index = 0usize;
        while cell_index < sorted_cells.len() {
            let ((row, col), cell) = sorted_cells[cell_index];
            let xf_index = match cell.pivot_xf_role {
                Some(super::worksheet::PivotCellXfRole::HeaderAccent) => {
                    pivot_xf_indices.header_accent
                },
                Some(super::worksheet::PivotCellXfRole::HeaderPlain) => {
                    fmt.cell_xf_index_for(cell.format_idx)
                },
                Some(super::worksheet::PivotCellXfRole::RowLabel) => pivot_xf_indices.row_label,
                Some(super::worksheet::PivotCellXfRole::Value) => pivot_xf_indices.value,
                None => fmt.cell_xf_index_for(cell.format_idx),
            };

            if !worksheet.pivot_tables.is_empty()
                && matches!(
                    cell.pivot_xf_role,
                    Some(super::worksheet::PivotCellXfRole::Value)
                )
                && matches!(cell.value, XlsCellValue::Number(_))
            {
                let mut mulrk_values = Vec::new();
                let mut next_index = cell_index;
                let mut expected_col = *col;

                while next_index < sorted_cells.len() {
                    let ((next_row, next_col), next_cell) = sorted_cells[next_index];
                    if next_row != row || *next_col != expected_col {
                        break;
                    }
                    if !matches!(
                        next_cell.pivot_xf_role,
                        Some(super::worksheet::PivotCellXfRole::Value)
                    ) {
                        break;
                    }
                    let next_xf_index = pivot_xf_indices.value;
                    let XlsCellValue::Number(next_value) = &next_cell.value else {
                        break;
                    };
                    mulrk_values.push((next_xf_index, *next_value));
                    expected_col = expected_col.saturating_add(1);
                    next_index += 1;
                }

                if mulrk_values.len() >= 2 {
                    row_first_cell_positions
                        .entry(*row)
                        .or_insert(stream.len() as u32);
                    biff::write_mulrk(&mut stream, *row, *col, &mulrk_values)?;
                    cell_index = next_index;
                    continue;
                }
            }

            match &cell.value {
                XlsCellValue::Number(value) => {
                    row_first_cell_positions
                        .entry(*row)
                        .or_insert(stream.len() as u32);
                    biff::write_number(&mut stream, *row, *col, xf_index, *value)?;
                },
                XlsCellValue::String(s) => {
                    let sst_index = *string_map.get(s).unwrap();
                    row_first_cell_positions
                        .entry(*row)
                        .or_insert(stream.len() as u32);
                    biff::write_labelsst(&mut stream, *row, *col, xf_index, sst_index)?;
                },
                XlsCellValue::Boolean(value) => {
                    row_first_cell_positions
                        .entry(*row)
                        .or_insert(stream.len() as u32);
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
            cell_index += 1;
        }

        let mut pivot_dbcell_positions = Vec::new();
        if !worksheet.pivot_tables.is_empty() && !emitted_rows.is_empty() {
            for row_block in emitted_rows.chunks(32) {
                let dbcell_pos = stream.len() as u32;
                let first_row_pos = row_record_positions
                    .get(&row_block[0])
                    .copied()
                    .ok_or_else(|| {
                        XlsError::InvalidData(
                            "pivot worksheet row block missing ROW record offset for DBCELL"
                                .to_string(),
                        )
                    })?;
                let mut cell_offsets = Vec::new();
                let mut previous_row_first_cell_pos = None;
                for row in row_block {
                    if let Some(first_cell_pos) = row_first_cell_positions.get(row).copied() {
                        let offset = if let Some(previous_pos) = previous_row_first_cell_pos {
                            first_cell_pos.saturating_sub(previous_pos)
                        } else {
                            first_cell_pos.saturating_sub(first_row_pos.saturating_add(20))
                        };
                        cell_offsets.push(u16::try_from(offset).map_err(|_| {
                            XlsError::InvalidData(
                                "pivot worksheet DBCELL cell offset exceeds BIFF8 limit"
                                    .to_string(),
                            )
                        })?);
                        previous_row_first_cell_pos = Some(first_cell_pos);
                    }
                }

                biff::write_dbcell(
                    &mut stream,
                    dbcell_pos.saturating_sub(first_row_pos),
                    &cell_offsets,
                )?;
                pivot_dbcell_positions.push(dbcell_pos);
            }
        }

        if let Some(index_record_pos) = pivot_index_record_pos {
            let mut index_record = Vec::new();
            biff::write_index(
                &mut index_record,
                pivot_first_used_row,
                pivot_last_used_row_plus1,
                pivot_def_col_width_pos.unwrap_or(0),
                &pivot_dbcell_positions,
            )?;
            stream[index_record_pos..index_record_pos + index_record.len()]
                .copy_from_slice(&index_record);
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

        if worksheet.pivot_tables.iter().any(|pt| {
            pt.page_entries
                .iter()
                .any(|&(_, _, object_id)| object_id != 0xFFFF)
        }) {
            biff::write_pivot_page_mso_drawing(&mut stream)?;
            biff::write_pivot_page_obj(&mut stream)?;
        }

        // Pivot table records (SX* family).
        //
        // Record order per LibreOffice xepivot.cxx XclExpPivotTable::Save():
        //   SxView
        //   *(SXVD *(SXVI) SXVDEx)    — per-field group
        //   SXIVD (row fields)
        //   SXIVD (col fields)
        //   [SXPI]
        //   *SXDI
        //   SxEx
        for (pt_local_idx, pt) in worksheet.pivot_tables.iter().enumerate() {
            let field_count = pt.fields.len() as u16;
            let data_field_count = pt.data_items.len() as u16;

            // Collect field indices per axis.
            let mut row_field_indices: Vec<u16> = Vec::new();
            let mut col_field_indices: Vec<u16> = Vec::new();
            let mut page_field_count: u16 = 0;
            for (i, f) in pt.fields.iter().enumerate() {
                match f.axis {
                    0x0001 => row_field_indices.push(i as u16),
                    0x0002 => col_field_indices.push(i as u16),
                    0x0004 => page_field_count += 1,
                    _ => {},
                }
            }

            let effective_data_axis = pt.data_axis;
            let mut effective_data_position = pt.data_position;

            // LibreOffice keeps single-data-field pivots row-oriented without the
            // EXC_SXIVD_DATA pseudo-field. The data layout pseudo-field is only
            // relevant when there are multiple data fields.
            if data_field_count <= 1 {
                effective_data_position = 0xFFFF;
            } else {
                let target_axis = match effective_data_axis {
                    0x0002 => Some(&mut col_field_indices),
                    0x0001 => Some(&mut row_field_indices),
                    _ => None,
                };

                if let Some(axis_fields) = target_axis {
                    if let Some(existing_pos) = axis_fields.iter().position(|&idx| idx == 0xFFFE) {
                        if axis_fields.last().copied() != Some(0xFFFE) {
                            effective_data_position = existing_pos as u16;
                        } else {
                            effective_data_position = 0xFFFF;
                        }
                    } else {
                        axis_fields.push(0xFFFE);
                        effective_data_position = 0xFFFF;
                    }
                }
            }

            let row_fields = row_field_indices.len() as u16;
            let col_fields = col_field_indices.len() as u16;

            // cRw / cCol = visible data body dimensions.
            // Per LO Finalize():
            //   rnDataXclCol = rnXclCol1 + mnRowFields
            //   rnDataXclRow = rnXclRow1 + mnColFields + 1
            //   mnDataCols = rnXclCol2 - rnDataXclCol + 1
            //   mnDataRows = rnXclRow2 - rnDataXclRow + 1
            let data_row_count = pt.last_row.saturating_sub(pt.first_data_row) + 1;
            let data_col_count = pt.last_col.saturating_sub(pt.first_data_col) + 1;

            let cache_index = pt_local_idx as u16;

            // 1) SXVIEW — view definition
            biff::write_sxview(
                &mut stream,
                &biff::SxViewConfig {
                    first_row: pt.first_row,
                    last_row: pt.last_row,
                    first_col: pt.first_col,
                    last_col: pt.last_col,
                    first_header_row: pt.first_header_row,
                    first_data_row: pt.first_data_row,
                    first_data_col: pt.first_data_col,
                    cache_index,
                    data_axis: effective_data_axis,
                    data_position: effective_data_position,
                    field_count,
                    row_field_count: row_fields,
                    col_field_count: col_fields,
                    page_field_count,
                    data_field_count,
                    data_row_count,
                    data_col_count,
                    // fRwGrand(0x01) | fColGrand(0x02) | fAutoFormat(0x08) | fAtrProc(0x200)
                    flags: 0x020B,
                    auto_format_index: 1,
                    name: &pt.name,
                    data_field_name: &pt.data_field_name,
                },
            )?;

            // 2) Per-field: SXVD + SXVI items + SXVDEx
            for field in &pt.fields {
                biff::write_sxvd(
                    &mut stream,
                    &biff::SxVdConfig {
                        axis: field.axis,
                        subtotal_count: field.subtotal_count,
                        subtotal_flags: field.subtotal_flags,
                        item_count: field.items.len() as u16,
                        name: field.name.as_deref(),
                    },
                )?;

                for item in &field.items {
                    biff::write_sxvi(
                        &mut stream,
                        &biff::SxViConfig {
                            item_type: item.item_type,
                            flags: item.flags,
                            cache_index: item.cache_index,
                            name: item.name.as_deref(),
                        },
                    )?;
                }

                // SXVDEx — mandatory per LibreOffice
                biff::write_sxvdex(&mut stream)?;
            }

            // 3) SXIVD — row field index list
            biff::write_sxivd(&mut stream, &row_field_indices)?;

            // 4) SXIVD — column field index list
            biff::write_sxivd(&mut stream, &col_field_indices)?;

            // 5) SXPI — page field entries
            if !pt.page_entries.is_empty() {
                biff::write_sxpi(&mut stream, &pt.page_entries)?;
            }

            // 6) SXDI — data items
            for di in &pt.data_items {
                biff::write_sxdi(
                    &mut stream,
                    &biff::SxDiConfig {
                        source_field_index: di.source_field_index,
                        function: di.function,
                        display_format: di.display_format,
                        base_field_index: di.base_field_index,
                        base_item_index: di.base_item_index,
                        num_format_index: di.num_format_index,
                        name: &di.name,
                    },
                )?;
            }

            // 7) SXLI — row line items, then column line items
            //    Per LO: WriteSxli(mnDataRows, mnRowFields) then
            //             WriteSxli(mnDataCols, mnColFields)
            biff::write_sxli(&mut stream, data_row_count, row_fields)?;
            biff::write_sxli(&mut stream, data_col_count, col_fields)?;

            // 8) SxEx — extended view properties
            // Per LO Finalize(): mnPagePerRow = mnPageFields,
            //                    mnPagePerCol = (mnPageFields > 0) ? 1 : 0
            biff::write_sxex(
                &mut stream,
                &biff::SxExConfig {
                    page_rows: page_field_count,
                    page_cols: if page_field_count > 0 { 1 } else { 0 },
                    ..biff::SxExConfig::default()
                },
            )?;

            let pivot_field_names: Vec<&str> = pt
                .fields
                .iter()
                .map(|field| field.cache_name.as_str())
                .collect();

            biff::write_pivot_modern_extensions(&mut stream, &pt.name, &pivot_field_names)?;
            biff::write_pivot_window2(&mut stream)?;
            biff::write_plv(&mut stream)?;
            biff::write_selection(&mut stream)?;
            biff::write_phonetic_pr(&mut stream)?;
            biff::write_sheet_ext(&mut stream)?;
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

    Ok(WorkbookStreams {
        workbook: stream,
        pivot_caches,
    })
}
