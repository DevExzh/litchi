use std::collections::{HashMap, HashSet};

use crate::writer::biff;
use crate::writer::formatting::FormattingManager;
use crate::writer::formula::{FormulaTokenizer, encode_ptg_tokens};
use crate::{XlsError, XlsResult};

use super::named_range::XlsDefinedName as InternalDefinedName;
use super::worksheet::WritableWorksheet;
use super::{
    XlsCalculationSettings, XlsCellValue, XlsCustomTableStyles, XlsExternalWorkbookOptions,
    XlsFileSharing, XlsFunctionGroupOptions, XlsVbaWriteMetadata, XlsWorkbookEnvironmentOptions,
    XlsWorkbookProtection, XlsWorkbookWindowOptions,
};

const DEFAULT_WRITE_ACCESS_USER: &str = "litchi";

/// Result of generating the workbook: the Workbook stream plus any pivot
/// cache storage streams that must be placed in `_SX_DB_CUR/nnnn`.
pub(crate) struct WorkbookStreams {
    /// The main Workbook BIFF stream.
    pub workbook: Vec<u8>,
    /// Pivot cache streams: `(stream_id, data)`.  Each goes into
    /// `_SX_DB_CUR/{stream_id:04X}`.
    pub pivot_caches: Vec<(u16, Vec<u8>)>,
}

#[derive(Clone, Copy)]
struct PivotCacheIdentity {
    /// Zero-based index used by SXVIEW.iCache.
    cache_index: u16,
    /// One-based identifier used by SXStreamID and `_SX_DB_CUR/nnnn`.
    stream_id: u16,
}

fn stage_pivot_cache_identities(
    worksheets: &[WritableWorksheet],
) -> XlsResult<Vec<Vec<PivotCacheIdentity>>> {
    let pivot_count = worksheets.iter().try_fold(0usize, |count, worksheet| {
        count
            .checked_add(worksheet.pivot_tables.len())
            .ok_or_else(|| XlsError::InvalidData("PivotTable cache count overflow".to_string()))
    })?;
    if pivot_count > usize::from(u16::MAX) {
        return Err(XlsError::InvalidData(format!(
            "PivotTable cache count {pivot_count} exceeds the BIFF8 limit of {}",
            u16::MAX
        )));
    }

    let mut next_index = 0usize;
    worksheets
        .iter()
        .map(|worksheet| {
            (0..worksheet.pivot_tables.len())
                .map(|_| {
                    let cache_index = u16::try_from(next_index).map_err(|_| {
                        XlsError::InvalidData("PivotTable cache index overflow".to_string())
                    })?;
                    let stream_id = cache_index.checked_add(1).ok_or_else(|| {
                        XlsError::InvalidData("PivotTable cache stream ID overflow".to_string())
                    })?;
                    next_index = next_index.checked_add(1).ok_or_else(|| {
                        XlsError::InvalidData("PivotTable cache index overflow".to_string())
                    })?;
                    Ok(PivotCacheIdentity {
                        cache_index,
                        stream_id,
                    })
                })
                .collect()
        })
        .collect()
}

#[allow(clippy::too_many_arguments)] // TODO: Refactor this function to accept a struct
pub(crate) fn generate_workbook_stream(
    use_1904_dates: bool,
    calculation_settings: XlsCalculationSettings,
    vba_metadata: Option<&XlsVbaWriteMetadata>,
    environment: XlsWorkbookEnvironmentOptions,
    workbook_window: XlsWorkbookWindowOptions,
    function_groups: &XlsFunctionGroupOptions,
    external_workbooks: &[XlsExternalWorkbookOptions],
    external_names: &[Vec<super::XlsExternalDefinedNameOptions>],
    add_in_functions: &[super::XlsAddInFunctionOptions],
    dde_or_ole_links: &[super::XlsDdeOrOleLinkOptions],
    fmt: &FormattingManager,
    custom_table_styles: Option<&XlsCustomTableStyles>,
    defined_names: &[InternalDefinedName],
    defined_name_records: &[(
        super::XlsDefinedNameRecordOptions,
        crate::XlsDefinedNameFutureRecords,
    )],
    shared_strings: &[String],
    sst_total: u32,
    workbook_protection: Option<XlsWorkbookProtection>,
    file_sharing: Option<&XlsFileSharing>,
    book_ext: Option<&crate::XlsBookExt>,
    theme: Option<&crate::XlsTheme>,
    mdx_metadata: Option<&crate::XlsMdxMetadata>,
    real_time_data: &[crate::XlsRealTimeData],
    web_publications: &[crate::XlsWebPub],
    xf_extensions: &[crate::XlsXfExt],
    style_extensions: &[crate::XlsStyleExt],
    worksheets: &[WritableWorksheet],
    string_map: &HashMap<String, u32>,
) -> XlsResult<WorkbookStreams> {
    if let Some(styles) = custom_table_styles {
        styles.validate(fmt)?;
    }
    super::validate_list_object_relationships(
        worksheets,
        custom_table_styles,
        defined_names,
        defined_name_records,
    )?;
    workbook_window.validate_for_sheet_count(worksheets.len())?;
    let active_sheet = usize::from(workbook_window.active_sheet_index);
    if !worksheets[active_sheet].view.selected {
        return Err(XlsError::InvalidData(format!(
            "active worksheet {active_sheet} must be selected in Window2"
        )));
    }
    let selected_sheet_count = worksheets
        .iter()
        .filter(|sheet| sheet.view.selected)
        .count();
    if selected_sheet_count != usize::from(workbook_window.selected_sheet_count) {
        return Err(XlsError::InvalidData(format!(
            "Window1 selected sheet count {} disagrees with Window2 selected state ({selected_sheet_count})",
            workbook_window.selected_sheet_count
        )));
    }
    // Stage the complete workbook-global cache identity map before emitting any
    // BIFF bytes. SXVIEW uses the zero-based index while SXStreamID and the OLE
    // cache stream name use the corresponding one-based stream ID.
    let pivot_cache_identities = stage_pivot_cache_identities(worksheets)?;
    let mut stream = Vec::new();
    let has_pivot_tables = worksheets.iter().any(|ws| !ws.pivot_tables.is_empty());
    let sheet_count = u16::try_from(worksheets.len()).unwrap_or(u16::MAX);
    let (protect_structure, protect_windows, password_hash, protect_revisions, revision_hash) =
        workbook_protection
            .map(|protection| {
                (
                    protection.protect_structure,
                    protection.protect_windows,
                    protection.password_hash.unwrap_or(0),
                    protection.protect_revisions,
                    protection.revision_password_hash.unwrap_or(0),
                )
            })
            .unwrap_or((false, false, 0, false, 0));

    // === Workbook Globals ===

    // BOF record (workbook)
    biff::write_bof(&mut stream, 0x0005)?;

    if environment.template {
        biff::write_template(&mut stream)?;
    }

    biff::write_interface_hdr(&mut stream, 0x04B0)?;
    biff::write_mms(&mut stream)?;
    biff::write_interface_end(&mut stream)?;
    biff::write_write_access(&mut stream, DEFAULT_WRITE_ACCESS_USER)?;
    if let Some(sharing) = file_sharing {
        if sharing.password_hash.is_some() {
            biff::write_write_protect(&mut stream)?;
        }
        biff::write_file_sharing(
            &mut stream,
            sharing.read_only_recommended,
            sharing.password_hash,
            &sharing.user_name,
        )?;
    }

    // CodePage record - BIFF8 requires Unicode codepage 1200 (0x04B0)
    biff::write_codepage(&mut stream, 0x04B0)?;

    biff::write_dsf(&mut stream, environment.has_biff5_stream)?;
    biff::write_excel9_file(&mut stream)?;
    biff::write_tab_id(&mut stream, sheet_count)?;
    if let Some(metadata) = vba_metadata {
        biff::write_ob_proj(&mut stream)?;
        if metadata.project.is_module_free() {
            biff::write_ob_no_macros(&mut stream)?;
        }
        biff::write_code_name(&mut stream, &metadata.workbook_code_name)?;
    }
    biff::write_function_groups(&mut stream, function_groups)?;
    biff::write_window_protect(&mut stream, protect_windows)?;
    biff::write_protect(&mut stream, protect_structure)?;
    biff::write_password(&mut stream, password_hash)?;
    biff::write_protection_rev4(&mut stream, protect_revisions)?;
    biff::write_password_rev4(&mut stream, revision_hash)?;

    // Window1 record (workbook window properties)
    biff::write_window1(&mut stream, &workbook_window, worksheets.len())?;

    biff::write_backup(&mut stream, environment.create_backup_copy)?;
    biff::write_hide_obj(
        &mut stream,
        match environment.object_display_mode {
            crate::XlsObjectDisplayMode::ShowAll => 0,
            crate::XlsObjectDisplayMode::ShowPlaceholders => 1,
            crate::XlsObjectDisplayMode::HideAll => 2,
        },
    )?;
    biff::write_date1904(&mut stream, use_1904_dates)?;
    biff::write_precision(&mut stream, calculation_settings.full_precision)?;
    biff::write_refresh_all(&mut stream, environment.refresh_external_data_on_load)?;
    biff::write_book_bool_raw(&mut stream, environment.book_bool_bits())?;

    // Write minimal formatting tables so XF index 0 is valid.
    // Order mirrors Apache POI's workbook creation:
    //  - FONT records
    //  - FORMAT records (built-in 0..7 + custom)
    //  - XF records (style and cell formats)
    fmt.write_fonts(&mut stream)?;
    fmt.write_number_formats(&mut stream)?;
    fmt.write_formats(&mut stream)?;
    if !xf_extensions.is_empty() {
        let xf_count = fmt.xf_record_count();
        for extension in xf_extensions {
            if extension.xf_index() >= xf_count {
                return Err(XlsError::InvalidData(format!(
                    "XFExt references XF index {} but only {xf_count} XF records are written",
                    extension.xf_index()
                )));
            }
        }
        // XFS = 16*XF [XFCRC 16*4050XFExt]: the extensions follow the XF
        // table directly; the pivot path writes its own XFCRC block later.
        if !has_pivot_tables {
            biff::write_xfcrc(&mut stream, xf_count)?;
        }
        for extension in xf_extensions {
            biff::write_xf_ext(&mut stream, extension)?;
        }
    }
    if let Some(styles) = custom_table_styles {
        biff::write_differential_formats(&mut stream, styles.differential_formats())?;
    }
    if has_pivot_tables {
        biff::write_pivot_xfext_block(&mut stream)?;
    }

    // Built-in STYLE records to align with Excel / POI
    // defaults. This makes standard cell styles (Normal, Currency, Percent,
    // etc.) visible to Excel even though we currently only use the default
    // cell XF (index 15) for all cells.
    biff::write_builtin_styles(&mut stream)?;
    for style_ext in style_extensions {
        biff::write_style_ext(&mut stream, style_ext)?;
    }
    if let Some(styles) = custom_table_styles {
        biff::write_custom_table_styles(
            &mut stream,
            styles.catalog(),
            styles.differential_formats().len(),
        )?;
    } else if has_pivot_tables {
        biff::write_table_styles(&mut stream)?;
    }

    // BoundSheet8 records are emitted later so pivot cache definitions can stay
    // adjacent to the workbook formatting block, matching Excel's globals order.
    let mut boundsheet_positions = Vec::new();

    // Internal SUPBOOK / EXTERNSHEET records are required for 3D
    // references used by defined names (NameParsedFormula) and pivot caches.
    let internal_links =
        (!defined_names.is_empty() || !defined_name_records.is_empty() || has_pivot_tables)
            && !worksheets.is_empty();
    if internal_links
        || !external_workbooks.is_empty()
        || !add_in_functions.is_empty()
        || !dde_or_ole_links.is_empty()
    {
        let externsheet_mode = if defined_names.is_empty() {
            biff::ExternSheetMode::WorkbookWide
        } else {
            biff::ExternSheetMode::PerSheet
        };
        biff::write_external_link_table(
            &mut stream,
            internal_links.then_some((sheet_count, externsheet_mode)),
            external_workbooks,
            external_names,
            add_in_functions,
            dde_or_ole_links,
        )?;
    }

    // NAME (Lbl) records for workbook- and sheet-scoped defined names.
    // These are stored in the globals substream and reference cell
    // areas using BIFF8 formula tokens.
    for defined_name in defined_names {
        let rgce = defined_name.to_biff_formula()?;
        biff::write_name(&mut stream, defined_name, &rgce)?;
        if let Some(comment) = &defined_name.comment {
            biff::write_name_comment(&mut stream, &defined_name.name, comment)?;
        }
    }
    let classic_limit = 32usize - usize::from(function_groups.built_in.count());
    let extended_count = function_groups
        .custom_categories
        .len()
        .saturating_sub(classic_limit);
    for (defined_name, future) in defined_name_records {
        if future
            .function_group
            .as_ref()
            .is_some_and(|value| value.category_index() >= extended_count)
        {
            return Err(XlsError::InvalidData(
                "NameFnGrp12 category does not reference an emitted FnGrp12 record".to_string(),
            ));
        }
        biff::write_defined_name_record(&mut stream, defined_name)?;
        if let Some(value) = &future.function_group {
            biff::write_name_function_group(&mut stream, value)?;
        }
        if let Some(value) = &future.publication {
            biff::write_name_publish(&mut stream, value)?;
        }
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
    {
        // Collect all pivot tables across worksheets.
        let all_pts: Vec<&super::worksheet::WritablePivotTable> = worksheets
            .iter()
            .flat_map(|ws| ws.pivot_tables.iter())
            .collect();
        let all_identities = pivot_cache_identities.iter().flatten();
        for (pt, identity) in all_pts.iter().zip(all_identities) {
            // LO uses 1-based IDs: maPCInfo.mnStrmId = nListIdx + 1.
            let id = identity.stream_id;

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
                .enumerate()
                .zip(unique_numeric_counts.iter())
                .map(|((field_index, f), &uniq_count)| biff::PivotCacheFieldInfo {
                    name: &f.cache_name,
                    items: &f.cache_items,
                    is_numeric: f.is_numeric,
                    unique_numeric_count: uniq_count,
                    grouping: f.grouping.as_ref(),
                    group_child: pt.fields.iter().position(|candidate| matches!(&candidate.grouping, Some(crate::PivotCacheGrouping::Discrete(value)) if usize::from(value.base_field_index) == field_index)).map(|index| index as u16),
                    is_source_field: !matches!(f.grouping, Some(crate::PivotCacheGrouping::Discrete(_))),
                })
                .collect();

            // Build source rows: split each PivotCacheValue row into
            // string_indices (for SXDBB) and numeric_values (for SXNUM).
            let num_string_fields = pt
                .fields
                .iter()
                .filter(|f| {
                    !f.cache_items.is_empty()
                        && !matches!(f.grouping, Some(crate::PivotCacheGrouping::Discrete(_)))
                })
                .count();
            let num_numeric_fields = pt
                .fields
                .iter()
                .filter(|f| f.is_numeric && f.cache_items.is_empty() && f.grouping.is_none())
                .count();
            let mut row_item_indices: Vec<Vec<u16>> = Vec::with_capacity(pt.source_data.len());
            let mut row_numeric_values: Vec<Vec<f64>> = Vec::with_capacity(pt.source_data.len());
            for row in &pt.source_data {
                let mut si = Vec::with_capacity(num_string_fields);
                let mut nv = Vec::with_capacity(num_numeric_fields);
                let mut values = row.iter();
                for field in &pt.fields {
                    if matches!(field.grouping, Some(crate::PivotCacheGrouping::Discrete(_))) {
                        continue;
                    }
                    let val = values.next().unwrap();
                    let is_num = field.is_numeric
                        && field.cache_items.is_empty()
                        && field.grouping.is_none();
                    match val {
                        super::PivotCacheValue::StringIndex(idx) if !is_num => {
                            si.push(u16::from(*idx))
                        },
                        super::PivotCacheValue::SharedItemIndex(idx) if !is_num => si.push(*idx),
                        super::PivotCacheValue::Number(v) if is_num => nv.push(*v),
                        value if !is_num => {
                            let item = match value {
                                super::PivotCacheValue::Number(number) => {
                                    Some(crate::PivotCacheItem::Number(*number))
                                },
                                _ => value.shared_item(),
                            };
                            if let Some(item) = item
                                && let Some(index) = field
                                    .cache_items
                                    .iter()
                                    .position(|candidate| candidate == &item)
                            {
                                si.push(index as u16);
                            }
                        },
                        _ => {}, // type mismatch — skip
                    }
                }
                row_item_indices.push(si);
                row_numeric_values.push(nv);
            }
            let source_rows: Vec<biff::PivotCacheSourceRow<'_>> = row_item_indices
                .iter()
                .zip(row_numeric_values.iter())
                .map(|(si, nv)| biff::PivotCacheSourceRow {
                    item_indices: si.as_slice(),
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

    let mut drawing_clusters = Vec::<(u32, u32)>::new();
    let mut worksheet_drawing_ids = Vec::with_capacity(worksheets.len());
    let mut next_drawing_id = 1u32;
    for worksheet in worksheets {
        let pivot_ids = worksheet
            .pivot_tables
            .iter()
            .flat_map(|table| table.page_entries.iter().map(|entry| entry.2))
            .filter(|id| *id != 0 && *id != u16::MAX)
            .collect::<HashSet<_>>();
        let group_object_count = worksheet
            .shape_groups
            .iter()
            .map(|group| 1 + group.children.len())
            .sum::<usize>();
        let object_count = pivot_ids.len()
            + worksheet.shapes.len()
            + group_object_count
            + worksheet.comments.len();
        if object_count > 1022 {
            return Err(XlsError::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        if object_count == 0 {
            worksheet_drawing_ids.push(None);
        } else {
            let drawing_id = next_drawing_id;
            next_drawing_id = next_drawing_id.checked_add(1).ok_or_else(|| {
                XlsError::InvalidData("workbook drawing IDs are exhausted".to_string())
            })?;
            drawing_clusters.push((drawing_id, object_count as u32 + 1));
            worksheet_drawing_ids.push(Some(drawing_id));
        }
    }

    biff::write_usesel_fs_value(&mut stream, environment.supports_natural_language_formulas)?;

    for worksheet in worksheets {
        boundsheet_positions.push(stream.len());
        biff::write_boundsheet(&mut stream, 0, &worksheet.name)?;
    }

    if let Some(settings) = calculation_settings.multithreaded_calculation {
        biff::write_mtr_settings(&mut stream, settings)?;
    }
    if calculation_settings.force_full_calculation {
        biff::write_force_full_calculation(&mut stream, true)?;
    }

    biff::write_country(
        &mut stream,
        environment.default_country_code,
        environment.current_country_code,
    )?;

    // RTD precedes RecalcId in the workbook globals grammar.
    for topic in real_time_data {
        biff::write_real_time_data(&mut stream, topic)?;
    }

    biff::write_recalc_id(&mut stream, calculation_settings.recalculation_engine_id)?;

    if has_pivot_tables {
        biff::write_compress_pictures(&mut stream)?;
        biff::write_compat12(&mut stream)?;
    }

    // MsoDrawingGroup (0x00EB) — if we have page fields, Excel expects a drawing group
    if !drawing_clusters.is_empty() {
        biff::write_mso_drawing_group(&mut stream, &drawing_clusters)?;
    }

    // SST record (shared string table)
    if !shared_strings.is_empty() {
        biff::write_sst(&mut stream, shared_strings, sst_total)?;
    }

    // WEBPUB follows ExtSST in the workbook globals grammar.
    for publication in web_publications {
        biff::write_web_pub(&mut stream, publication)?;
    }

    // BOOKEXT follows the shared-string/web-publishing region in the
    // workbook globals grammar.
    if let Some(book_ext) = book_ext {
        biff::write_book_ext(&mut stream, book_ext)?;
    }

    // THEME follows the BookExt/data-connection region in the workbook
    // globals grammar.
    if let Some(theme) = theme {
        biff::write_theme(&mut stream, theme)?;
    }

    // METADATA follows the Theme/custom-view region in the workbook globals
    // grammar.
    if let Some(metadata) = mdx_metadata
        && !metadata.is_empty()
    {
        biff::write_mdx_metadata(&mut stream, metadata)?;
    }

    // EOF record (end of workbook globals)
    biff::write_eof(&mut stream)?;

    // === Worksheets ===

    // Track actual worksheet positions
    let mut actual_positions = Vec::new();

    for (worksheet_index, worksheet) in worksheets.iter().enumerate() {
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
            rows.extend(row_spans.keys().copied());
            rows.into_iter().collect()
        };
        // BOF record (worksheet)
        biff::write_bof(&mut stream, 0x0010)?;

        if worksheet.formulas_pending_recalculation {
            biff::write_uncalced(&mut stream)?;
        }

        let index_record_pos = stream.len();

        biff::write_calculation_settings(&mut stream, &calculation_settings)?;

        if worksheet.pivot_tables.is_empty() {
            biff::write_worksheet_layout(&mut stream, &worksheet.sheet_layout)?;
        }

        if let Some(page_setup) = &worksheet.page_setup {
            biff::write_page_settings(
                &mut stream,
                page_setup,
                &worksheet.horizontal_page_breaks,
                &worksheet.vertical_page_breaks,
            )?;
        }

        if let Some(manager) = &worksheet.scenario_manager {
            biff::write_scenario_manager(&mut stream, manager)?;
        }

        let mut def_col_width_pos = None;
        if worksheet.pivot_tables.is_empty() {
            def_col_width_pos = Some(stream.len());
            biff::write_def_col_width(
                &mut stream,
                worksheet.sheet_layout.default_column_width_chars,
            )?;
            biff::write_dimensions(
                &mut stream,
                worksheet.first_row,
                worksheet.last_row,
                worksheet.first_col,
                worksheet.last_col,
            )?;
        } else {
            biff::write_pivot_sheet_preamble(&mut stream, &worksheet.sheet_layout)?;
        }

        if let Some(consolidation) = &worksheet.consolidation {
            biff::write_consolidation(&mut stream, consolidation)?;
        }

        // Required sheet records for worksheet substream per MS-XLS.
        //
        // Apache POI writes WINDOW2 first and then (optionally) PANE
        // immediately afterwards when freeze panes are configured. We
        // mirror that ordering here to avoid Excel interpreting the
        // pane as a generic split window.
        if worksheet.pivot_tables.is_empty() {
            biff::write_window2_options(&mut stream, &worksheet.view)?;
            if let Some(scale) = worksheet.view.scale
                && scale.numerator != scale.denominator
            {
                biff::write_scl(&mut stream, scale.numerator, scale.denominator)?;
            }
            if let Some(pane) = worksheet.view.pane.as_ref() {
                biff::write_pane_options(&mut stream, pane)?;
            }
            for selection in &worksheet.view.selections {
                biff::write_selection_options(&mut stream, selection)?;
            }
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
        if let Some(ref sort_data) = worksheet.sort_data
            && !matches!(sort_data.parent(), crate::XlsSortParent::Table { .. })
        {
            sort_data.write_biff_records(&mut stream)?;
        }

        // Column width / hidden state via COLINFO records.
        if !worksheet.pivot_tables.is_empty() {
            def_col_width_pos = Some(stream.len());
            biff::write_def_col_width(
                &mut stream,
                worksheet.sheet_layout.default_column_width_chars,
            )?;
        }

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
        let row_table_start = stream.len();
        if !emitted_rows.is_empty() {
            for row in &emitted_rows {
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
                    biff::write_mulrk(&mut stream, *row, *col, &mulrk_values)?;
                    cell_index = next_index;
                    continue;
                }
            }

            // A data-table anchor cell emits its PtgTbl formula immediately
            // followed by the Table record (MS-XLS 2.4.319).
            if let Some((_, _, table)) = worksheet
                .data_tables
                .iter()
                .find(|(anchor_row, anchor_col, _)| (*anchor_row, *anchor_col) == (*row, *col))
            {
                let tokens = table.ptg_tbl_tokens();
                biff::write_formula(&mut stream, *row, *col, xf_index, &tokens)?;
                biff::write_table(&mut stream, table)?;
                cell_index += 1;
                continue;
            }

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
                XlsCellValue::Formula(formula) => {
                    let expression = formula.strip_prefix('=').unwrap_or(formula);
                    let tokens = FormulaTokenizer::new().tokenize(expression)?;
                    let encoded = encode_ptg_tokens(&tokens);
                    biff::write_formula(&mut stream, *row, *col, xf_index, &encoded)?;
                },
                XlsCellValue::Blank => {
                    // Skip blank cells
                },
            }
            cell_index += 1;
        }

        let staged_row_table = stream.split_off(row_table_start);
        let def_col_width_pos = def_col_width_pos.ok_or_else(|| {
            XlsError::InvalidData("worksheet is missing DEFCOLWIDTH for INDEX".to_string())
        })?;
        let plan =
            crate::writer::row_blocks::XlsRowBlockLayoutPlan::generate_from_staged(
                u64::try_from(index_record_pos).map_err(|_| {
                    XlsError::InvalidData("worksheet INDEX position overflow".to_string())
                })?,
                u64::try_from(
                    row_table_start
                        .checked_sub(index_record_pos)
                        .ok_or_else(|| {
                            XlsError::InvalidData("worksheet row table precedes INDEX".to_string())
                        })?,
                )
                .map_err(|_| {
                    XlsError::InvalidData("worksheet row-table position overflow".to_string())
                })?,
                u64::try_from(def_col_width_pos.checked_sub(index_record_pos).ok_or_else(
                    || XlsError::InvalidData("worksheet DEFCOLWIDTH precedes INDEX".to_string()),
                )?)
                .map_err(|_| {
                    XlsError::InvalidData("worksheet DEFCOLWIDTH position overflow".to_string())
                })?,
                &staged_row_table,
            )?;
        let (index_record, row_table) = plan.into_records();
        stream.splice(index_record_pos..index_record_pos, index_record);
        stream.extend_from_slice(&row_table);

        if let Some(drawing_id) = worksheet_drawing_ids[worksheet_index] {
            let pivot_object_ids = worksheet
                .pivot_tables
                .iter()
                .flat_map(|table| table.page_entries.iter().map(|entry| entry.2))
                .filter(|id| *id != 0 && *id != u16::MAX)
                .collect::<Vec<_>>();
            let mut reserved = HashSet::new();
            for &object_id in &pivot_object_ids {
                if !reserved.insert(object_id) {
                    return Err(XlsError::InvalidData(
                        "pivot page object ID is duplicated on the worksheet".to_string(),
                    ));
                }
            }
            let mut primitive_configs = Vec::with_capacity(worksheet.shapes.len());
            for shape in &worksheet.shapes {
                let object_id = shape.object_id.ok_or_else(|| {
                    XlsError::InvalidData("writable shape has no assigned object ID".to_string())
                })?;
                if object_id == 0 || object_id == u16::MAX || !reserved.insert(object_id) {
                    return Err(XlsError::InvalidData(
                        "shape object ID is reserved or duplicated on the worksheet".to_string(),
                    ));
                }
                primitive_configs.push(biff::PrimitiveShapeConfig { shape, object_id });
            }
            let mut group_configs = Vec::with_capacity(worksheet.shape_groups.len());
            for group in &worksheet.shape_groups {
                let object_id = group.object_id.ok_or_else(|| {
                    XlsError::InvalidData(
                        "writable shape group has no assigned object ID".to_string(),
                    )
                })?;
                if object_id == 0 || object_id == u16::MAX || !reserved.insert(object_id) {
                    return Err(XlsError::InvalidData(
                        "shape object ID is reserved or duplicated on the worksheet".to_string(),
                    ));
                }
                let mut child_object_ids = Vec::with_capacity(group.children.len());
                for child in &group.children {
                    let child_id = child.object_id.ok_or_else(|| {
                        XlsError::InvalidData("grouped shape has no assigned object ID".to_string())
                    })?;
                    if child_id == 0 || child_id == u16::MAX || !reserved.insert(child_id) {
                        return Err(XlsError::InvalidData(
                            "shape object ID is reserved or duplicated on the worksheet"
                                .to_string(),
                        ));
                    }
                    child_object_ids.push(child_id);
                }
                group_configs.push(biff::GroupShapeConfig {
                    group,
                    object_id,
                    child_object_ids,
                });
            }
            let mut next_object_id = 1u16;
            let mut configs = Vec::with_capacity(worksheet.comments.len());
            let mut guids = HashSet::new();
            for comment in &worksheet.comments {
                while reserved.contains(&next_object_id) {
                    next_object_id = next_object_id.checked_add(1).ok_or_else(|| {
                        XlsError::InvalidData(
                            "worksheet comment object IDs are exhausted".to_string(),
                        )
                    })?;
                }
                let object_id = next_object_id;
                reserved.insert(object_id);
                next_object_id = next_object_id.saturating_add(1);
                let guid = comment.options.guid.unwrap_or_else(|| {
                    super::comment::deterministic_comment_guid(
                        worksheet_index,
                        comment.row,
                        comment.column,
                        object_id,
                    )
                });
                if !guids.insert(guid) {
                    return Err(XlsError::InvalidData(
                        "comment GUID is duplicated after planning".to_string(),
                    ));
                }
                configs.push(biff::CommentConfig {
                    row: comment.row,
                    column: comment.column,
                    author: &comment.author,
                    text: &comment.text,
                    visible: comment.options.visible,
                    shared: comment.options.shared,
                    anchor: comment.anchor(),
                    text_runs: &comment.options.text_runs,
                    font_when_empty: comment.options.font_when_empty,
                    guid,
                    object_id,
                });
            }
            biff::write_worksheet_drawing(
                &mut stream,
                drawing_id,
                &pivot_object_ids,
                &primitive_configs,
                &group_configs,
                &configs,
            )?;
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

        if !worksheet.data_validations.is_empty()
            || worksheet.data_validation_table_options.is_some()
        {
            let dv_count = worksheet.data_validations.len() as u32;
            let table = worksheet.data_validation_table_options.unwrap_or_default();
            biff::write_dval(
                &mut stream,
                biff::DvalConfig {
                    window_closed: table.window_closed,
                    x_left: table.x_left,
                    y_top: table.y_top,
                    dropdown_object_id: table.dropdown_object_id,
                    dv_count,
                },
            )?;

            for writable in &worksheet.data_validations {
                let dv = &writable.validation;
                let payload = dv.validation_type.to_biff_payload()?;

                let ranges = writable
                    .ranges
                    .iter()
                    .map(|range| {
                        (
                            range.first_row,
                            range.last_row,
                            range.first_col,
                            range.last_col,
                        )
                    })
                    .collect::<Vec<_>>();

                let formula1 = payload.formula1.as_deref();
                let formula2 = payload.formula2.as_deref();

                let dv_config = biff::DvConfig {
                    data_type: payload.data_type,
                    operator: payload.operator,
                    error_style: writable.options.error_style.to_biff_code(),
                    empty_cell_allowed: writable.options.allow_blank,
                    suppress_dropdown_arrow: writable.options.suppress_dropdown,
                    is_explicit_list_formula: payload.is_explicit_list,
                    ime_mode: writable.options.ime_mode.to_biff_code(),
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

        // PHONETICINFO precedes CONDFMTS in the worksheet substream grammar.
        if let Some(phonetic_info) = &worksheet.phonetic_info {
            biff::write_phonetic_info(&mut stream, phonetic_info)?;
        }

        if !worksheet.conditional_formats.is_empty() {
            for (identifier, group) in worksheet.conditional_formats.iter().enumerate() {
                let ranges = group
                    .ranges
                    .iter()
                    .map(|range| {
                        (
                            range.first_row,
                            range.last_row,
                            range.first_col,
                            range.last_col,
                        )
                    })
                    .collect::<Vec<_>>();
                biff::write_cfheader_with_identifier(
                    &mut stream,
                    &ranges,
                    group.rules.len() as u16,
                    identifier as u16,
                )?;
                for rule in &group.rules {
                    let (condition_type, comparison_op, formula1, formula2) =
                        rule.format_type.to_biff_payload()?;
                    let pattern = rule.pattern.as_ref().map(|pat| {
                        (
                            pat.pattern as u16,
                            pat.foreground_color & 0x7f,
                            pat.background_color & 0x7f,
                        )
                    });
                    biff::write_cfrule(
                        &mut stream,
                        condition_type,
                        comparison_op,
                        &formula1,
                        &formula2,
                        pattern,
                    )?;
                }
            }
        }

        for (offset, group) in worksheet.conditional_formats12.iter().enumerate() {
            let identifier = worksheet.conditional_formats.len() + offset;
            let ranges = group
                .ranges
                .iter()
                .map(|range| {
                    (
                        range.first_row,
                        range.last_row,
                        range.first_col,
                        range.last_col,
                    )
                })
                .collect::<Vec<_>>();
            biff::write_condfmt12(
                &mut stream,
                &ranges,
                group.rules.len() as u16,
                identifier as u16,
            )?;
            for rule in &group.rules {
                let (condition_type, comparison, formula1, formula2, active_formula, payload) =
                    rule.format_type.biff_parts();
                let config = biff::Cf12Config {
                    condition_type,
                    comparison,
                    differential_format: &rule.differential_format,
                    formula1,
                    formula2,
                    active_formula,
                    stop_if_true: rule.stop_if_true,
                    priority: rule.priority,
                    template: rule.template,
                    template_parameters: rule.template_parameters,
                    rule_payload: payload,
                };
                biff::write_cf12(&mut stream, &config)?;
            }
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
        for (pt, identity) in worksheet
            .pivot_tables
            .iter()
            .zip(&pivot_cache_identities[worksheet_index])
        {
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

            let cache_index = identity.cache_index;

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
            biff::write_pivot_window2(&mut stream, worksheet.view.selected)?;
            biff::write_plv(&mut stream)?;
            biff::write_selection(&mut stream)?;
            biff::write_sheet_ext(&mut stream)?;
        }

        // PHONETICINFO is a per-sheet record; emit the pivot-era stub at
        // most once even when several pivot tables share the sheet.
        if worksheet.phonetic_info.is_none() && !worksheet.pivot_tables.is_empty() {
            biff::write_phonetic_pr(&mut stream)?;
        }

        if let Some(code_name) = &worksheet.vba_code_name {
            biff::write_code_name(&mut stream, code_name)?;
        }

        // WEBPUB follows CodeName and precedes SheetExt in the worksheet
        // substream grammar.
        for publication in &worksheet.web_publications {
            biff::write_web_pub(&mut stream, publication)?;
        }

        // SHEETEXT sits after CodeName and before the FEAT records
        // (list objects) in the worksheet substream grammar.
        if let Some(tab_color) = worksheet.tab_color {
            biff::write_sheet_ext_tab_color(&mut stream, tab_color)?;
        }

        biff::write_list_objects(
            &mut stream,
            &worksheet.list_objects,
            worksheet.sort_data.as_ref(),
        )?;

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
