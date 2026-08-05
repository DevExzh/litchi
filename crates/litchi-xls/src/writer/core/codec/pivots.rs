use super::super::model::validate_pivot_table_config;
use super::super::*;
use crate::error::{Error, Result};

impl Writer {
    /// Add a pivot table definition to a worksheet.
    ///
    /// This writes the SX* record family (SXVS, SXVIEW, SXVD, SXVI, SXDI,
    /// SXPI) to the worksheet stream. The pivot table must be fully
    /// configured before calling this method.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `config` — pivot table configuration (see [`PivotTableConfig`])
    pub fn add_pivot_table(&mut self, sheet: usize, config: PivotTableConfig) -> Result<()> {
        validate_pivot_table_config(&config)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Generate pivot output cells BEFORE consuming config.fields / config.data_items.
        // Excel validates that DIMENSIONS and cell content are consistent with the
        // pivot table definition; missing cells cause a "corrupt file" repair dialog.
        Self::generate_pivot_output_cells(worksheet, &config)?;
        self.fmt.enable_pivot_xfs();

        let fields: Vec<WritablePivotField> = config
            .fields
            .into_iter()
            .map(|f| {
                let mut items: Vec<WritablePivotItem> = f
                    .items
                    .into_iter()
                    .map(|i| WritablePivotItem {
                        item_type: i.item_type,
                        flags: i.flags,
                        cache_index: i.cache_index,
                        name: i.name,
                    })
                    .collect();

                // Sort data items (item_type=0x0000) alphabetically by their
                // cache label to match Excel's default SXVI ordering.  Non-data
                // items (subtotals etc.) stay at the end.
                let data_end = items
                    .iter()
                    .position(|i| i.item_type != 0x0000)
                    .unwrap_or(items.len());
                items[..data_end].sort_unstable_by(|a, b| {
                    let al = f
                        .cache_items
                        .get(a.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    let bl = f
                        .cache_items
                        .get(b.cache_index as usize)
                        .map(crate::PivotCacheItem::display_text)
                        .unwrap_or_default();
                    al.cmp(&bl)
                });

                WritablePivotField {
                    axis: f.axis,
                    subtotal_count: f.subtotal_count,
                    subtotal_flags: f.subtotal_flags,
                    items,
                    name: f.name,
                    cache_name: f.cache_name,
                    cache_items: f.cache_items,
                    is_numeric: f.is_numeric,
                    grouping: f.grouping,
                }
            })
            .collect();

        let data_items: Vec<WritablePivotDataItem> = config
            .data_items
            .into_iter()
            .map(|d| WritablePivotDataItem {
                source_field_index: d.source_field_index,
                function: d.function,
                display_format: d.display_format,
                base_field_index: d.base_field_index,
                base_item_index: d.base_item_index,
                num_format_index: d.num_format_index,
                name: d.name,
            })
            .collect();

        worksheet.add_pivot_table(WritablePivotTable {
            name: config.name,
            source_type: config.source_type,
            source_sheet_name: config.source_sheet_name,
            source_first_row: config.source_first_row,
            source_last_row: config.source_last_row,
            source_first_col: config.source_first_col,
            source_last_col: config.source_last_col,
            first_row: config.first_row,
            last_row: config.last_row,
            first_col: config.first_col,
            last_col: config.last_col,
            first_header_row: config.first_header_row,
            first_data_row: config.first_data_row,
            first_data_col: config.first_data_col,
            data_field_name: config.data_field_name,
            data_axis: config.data_axis,
            data_position: config.data_position,
            fields,
            data_items,
            page_entries: config.page_entries,
            source_data: config.source_data,
        });

        Ok(())
    }

    /// Generate the cell data that Excel expects in the SXVIEW output area.
    ///
    /// The layout (for a single row-field, single col-field, single page-field,
    /// single data-field configuration) is:
    ///
    /// ```text
    /// (first_row-2, 0)       : page field name    (first_row-2, 1)       : "(All)"
    /// (first_row,   0)       : data item name      (first_row, first_data_col): "Column Labels"
    /// (first_header_row, 0)  : "Row Labels"        (fhr, fdc+j)           : col item names …
    /// (first_data_row+i, 0)  : row item name       (fdr+i, fdc+j)         : aggregated value
    /// (last_row, 0)          : "Grand Total"        (lr, fdc+j)            : column totals
    /// ```
    fn generate_pivot_output_cells(
        ws: &mut WritableWorksheet,
        cfg: &PivotTableConfig,
    ) -> Result<()> {
        // Identify fields per axis.
        let row_field = cfg.fields.iter().find(|f| f.axis == 0x0001);
        let col_field = cfg.fields.iter().find(|f| f.axis == 0x0002);
        let page_field = cfg.fields.iter().find(|f| f.axis == 0x0004);

        let data_item = cfg.data_items.first();

        // Helper: find the field index for a given field by cache_name.
        let field_idx_of =
            |name: &str| -> Option<usize> { cfg.fields.iter().position(|f| f.cache_name == name) };

        // Collect row/col item labels from cache_items, sorted alphabetically
        // to match Excel's default SXVI ordering.  Also build a mapping from
        // cache_index → sorted position so the aggregation grid uses the same
        // order as the output rows/columns.
        let (row_items, row_cache_to_sorted) = Self::sorted_cache_items(row_field);
        let (col_items, col_cache_to_sorted) = Self::sorted_cache_items(col_field);

        let fr = cfg.first_row;
        let fhr = cfg.first_header_row;
        let fdr = cfg.first_data_row;
        let fdc = cfg.first_data_col;
        let lr = cfg.last_row;
        let lc = cfg.last_col;
        let fc = cfg.first_col;

        let offset = |base: u16, amount: usize| -> Result<u16> {
            let amount = u16::try_from(amount).map_err(|_| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })?;
            base.checked_add(amount).ok_or_else(|| {
                Error::InvalidCellReference("PivotTable output exceeds the BIFF8 grid".to_string())
            })
        };
        let mut staged = Vec::new();
        let mut add = |row: u16,
                       col: u16,
                       value: CellValue,
                       pivot_xf_role: Option<PivotCellXfRole>|
         -> Result<()> {
            staged.push(WritableCell::new(
                CellPos::try_new(u32::from(row), col)?,
                value,
                0,
                pivot_xf_role,
            ));
            Ok(())
        };

        // --- Page field area (above SXVIEW range) ---
        if let Some(pf) = page_field {
            let page_row = fr.saturating_sub(2);
            add(
                page_row,
                0,
                CellValue::String(pf.cache_name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
            add(
                page_row,
                1,
                CellValue::String("(All)".to_string()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }

        // --- Row at first_row: data item name + "Column Labels" ---
        if let Some(di) = data_item {
            add(
                fr,
                fc,
                CellValue::String(di.name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }
        if col_field.is_some() {
            add(
                fr,
                fdc,
                CellValue::String("Column Labels".to_string()),
                Some(PivotCellXfRole::HeaderAccent),
            )?;
        }

        // --- Row at first_header_row: "Row Labels" + column item names + "Grand Total" ---
        add(
            fhr,
            fc,
            CellValue::String("Row Labels".to_string()),
            Some(PivotCellXfRole::HeaderAccent),
        )?;
        for (j, ci) in col_items.iter().enumerate() {
            add(
                fhr,
                offset(fdc, j)?,
                CellValue::String(ci.clone()),
                Some(PivotCellXfRole::HeaderPlain),
            )?;
        }
        add(
            fhr,
            lc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::HeaderPlain),
        )?;

        // --- Compute aggregated values from source_data ---
        let row_fi = row_field.and_then(|f| field_idx_of(&f.cache_name));
        let col_fi = col_field.and_then(|f| field_idx_of(&f.cache_name));
        let data_fi = data_item.map(|di| di.source_field_index as usize);

        let nr = row_items.len();
        let nc = col_items.len();
        let mut grid = vec![vec![0.0f64; nc]; nr];
        let mut row_totals = vec![0.0f64; nr];
        let mut col_totals = vec![0.0f64; nc];
        let mut grand_total = 0.0f64;

        for row_data in &cfg.source_data {
            // Map cache indices through the sorted permutation so that
            // grid positions match the alphabetically-sorted output.
            let ri = row_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    row_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let ci = col_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    col_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let val = data_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::Number(v)) => Some(*v),
                _ => None,
            });

            if let (Some(ri), Some(ci), Some(val)) = (ri, ci, val)
                && ri < nr
                && ci < nc
            {
                grid[ri][ci] += val;
                row_totals[ri] += val;
                col_totals[ci] += val;
                grand_total += val;
            }
        }

        // --- Data rows ---
        for (i, (ri_name, row_total)) in row_items.iter().zip(row_totals.iter()).enumerate() {
            let r = offset(fdr, i)?;
            add(
                r,
                fc,
                CellValue::String(ri_name.clone()),
                Some(PivotCellXfRole::RowLabel),
            )?;
            for (j, cell_val) in grid[i].iter().enumerate() {
                add(
                    r,
                    offset(fdc, j)?,
                    CellValue::Number(*cell_val),
                    Some(PivotCellXfRole::Value),
                )?;
            }
            add(
                r,
                lc,
                CellValue::Number(*row_total),
                Some(PivotCellXfRole::Value),
            )?;
        }

        // --- Grand total row ---
        add(
            lr,
            fc,
            CellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::RowLabel),
        )?;
        for (j, col_total) in col_totals.iter().enumerate() {
            add(
                lr,
                offset(fdc, j)?,
                CellValue::Number(*col_total),
                Some(PivotCellXfRole::Value),
            )?;
        }
        add(
            lr,
            lc,
            CellValue::Number(grand_total),
            Some(PivotCellXfRole::Value),
        )?;
        for cell in staged {
            ws.add_cell(cell);
        }
        Ok(())
    }

    /// Sort a field's cache items alphabetically and return the sorted labels
    /// plus a mapping from original cache index to sorted position.
    ///
    /// Returns `(sorted_labels, cache_to_sorted)` where `cache_to_sorted[i]`
    /// gives the position of original cache item `i` in the sorted output.
    fn sorted_cache_items(field: Option<&PivotFieldConfig>) -> (Vec<String>, Vec<usize>) {
        let Some(f) = field else {
            return (Vec::new(), Vec::new());
        };

        // Build (original_index, label) pairs and sort by label.
        let mut indexed: Vec<(usize, String)> = f
            .cache_items
            .iter()
            .enumerate()
            .map(|(i, item)| (i, item.display_text()))
            .collect();
        indexed.sort_unstable_by(|a, b| a.1.cmp(&b.1));

        let sorted_labels: Vec<String> = indexed.iter().map(|(_, value)| value.clone()).collect();

        // cache_to_sorted[original_cache_idx] = position in sorted output
        let mut cache_to_sorted = vec![0usize; f.cache_items.len()];
        for (sorted_pos, (orig_idx, _)) in indexed.iter().enumerate() {
            cache_to_sorted[*orig_idx] = sorted_pos;
        }

        (sorted_labels, cache_to_sorted)
    }
}
