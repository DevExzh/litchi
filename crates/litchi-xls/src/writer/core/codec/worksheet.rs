use super::super::super::biff::AutoFilterConditionWrite;
use super::super::model::{a1_cell, prepare_data_validation};
use super::super::*;
use crate::error::{Error, Result};
use std::collections::HashSet;

impl Writer {
    pub fn set_auto_filter(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        MergedRange::try_new(first_row, last_row, first_col, last_col)?;
        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.list_objects.iter().any(|table| {
            first_row <= u32::from(table.range().last_row())
                && u32::from(table.range().first_row()) <= last_row
                && first_col <= table.range().last_column()
                && table.range().first_column() <= last_col
        }) {
            return Err(Error::InvalidData(
                "set_auto_filter: range overlaps a worksheet table".to_string(),
            ));
        }
        let itab = sheet
            .checked_add(1)
            .and_then(|index| u16::try_from(index).ok())
            .ok_or_else(|| {
                Error::InvalidData(
                    "set_auto_filter: sheet index exceeds BIFF8 itab limit".to_string(),
                )
            })?;
        let target_sheet = u16::try_from(sheet).map_err(|_| {
            Error::InvalidData("set_auto_filter: sheet index exceeds BIFF8 limit".to_string())
        })?;

        let start_ref = a1_cell(first_row, first_col);
        let end_ref = a1_cell(last_row, last_col);
        let reference = format!("{start_ref}:{end_ref}");
        let defined_name = DefinedName {
            name: "_FilterDatabase".to_string(),
            reference,
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(target_sheet),
            hidden: true,
            is_function: false,
            is_built_in: true,
            built_in_code: Some(0x0D),
        };

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.auto_filter = Some(AutoFilterRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });
        self.defined_names.retain(|n| {
            !(n.is_built_in && n.built_in_code == Some(0x0D) && n.local_sheet == Some(itab))
        });
        self.defined_names.push(defined_name);

        Ok(())
    }

    /// Add a filter condition to a specific column within the AutoFilter range.
    ///
    /// The AutoFilter range must first be set via [`Self::set_auto_filter`]. The
    /// `column_index` is 0-based relative to the filter range start column.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `column_index` — column within the filter range (0-based relative)
    /// * `join_or` — `true` to join conditions with OR, `false` for AND
    /// * `cond1` — first filter condition
    /// * `cond2` — second filter condition (use `AutoFilterConditionWrite::None` if unused)
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi_xls::writer::biff::AutoFilterConditionWrite;
    ///
    /// // Filter column 2: value > 100
    /// writer.add_filter_condition(
    ///     sheet_idx, 2, false,
    ///     AutoFilterConditionWrite::Number { operator: 0x04, value: 100.0 },
    ///     AutoFilterConditionWrite::None,
    /// )?;
    /// ```
    pub fn add_filter_condition(
        &mut self,
        sheet: usize,
        column_index: u16,
        join_or: bool,
        cond1: AutoFilterConditionWrite,
        cond2: AutoFilterConditionWrite,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let Some(filter) = worksheet.auto_filter else {
            return Err(Error::InvalidData(
                "add_filter_condition: call set_auto_filter first".to_string(),
            ));
        };
        let width = filter.last_col - filter.first_col + 1;
        if column_index >= width {
            return Err(Error::InvalidCellReference(format!(
                "AutoFilter relative column {column_index} exceeds its {width}-column range"
            )));
        }
        if worksheet
            .auto_filter_columns
            .iter()
            .any(|entry| entry.column_index == column_index)
        {
            return Err(Error::InvalidData(
                "AutoFilter column already has a condition".to_string(),
            ));
        }

        worksheet.add_auto_filter_column(AutoFilterColumnDef {
            column_index,
            join_or,
            condition1: cond1,
            condition2: cond2,
        });

        Ok(())
    }

    /// Set the sort configuration for a worksheet.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `case_sensitive` — whether sorting is case-sensitive
    /// * `sort_by_columns` — `true` for left-to-right sort, `false` for top-to-bottom
    /// * `keys` — up to 3 sort keys as `(column_index, descending)` tuples
    pub fn set_sort(
        &mut self,
        sheet: usize,
        case_sensitive: bool,
        sort_by_columns: bool,
        keys: &[(u16, bool)],
    ) -> Result<()> {
        if keys.is_empty() || keys.len() > 3 {
            return Err(Error::InvalidData(
                "set_sort: must provide 1..3 sort keys".to_string(),
            ));
        }
        if keys.iter().any(|(col, _)| *col > u16::from(u8::MAX)) {
            return Err(Error::InvalidCellReference(
                "sort key column is outside the BIFF8 grid".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.set_sort_config(SortConfig {
            case_sensitive,
            sort_by_columns,
            keys: keys.to_vec(),
        });

        Ok(())
    }

    /// Replace the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Unlike [`set_sort`](Self::set_sort), this preserves the complete
    /// `SortData` model, including an explicit range, more than three keys,
    /// custom lists, differential-format colors, and icon sets. The previous
    /// owned configuration is returned.
    pub fn put_sort(
        &mut self,
        sheet: usize,
        sort: crate::writer::sort::Config,
    ) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if let crate::writer::sort::Parent::Table { id } = sort.parent()
            && !worksheet
                .list_objects
                .iter()
                .any(|table| table.id().value() == id)
        {
            return Err(Error::InvalidData(
                "table SortData references an unknown ListObject identifier".to_string(),
            ));
        }
        Ok(worksheet.put_sort(sort))
    }

    /// Remove and return the extended BIFF8 sort metadata for a worksheet.
    ///
    /// Removing an absent configuration succeeds and returns `None`.
    pub fn remove_sort(&mut self, sheet: usize) -> Result<Option<crate::writer::sort::Config>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        Ok(worksheet.remove_sort())
    }

    /// Set the width of a column in character units.
    ///
    /// The column index is 0-based (0 = column A), matching the rest of the
    /// XLS writer API. The width is specified in the same units as Excel's
    /// UI, i.e. the number of characters of the "0" glyph in the default
    /// font. Internally this is converted to BIFF8 units of 1/256 characters
    /// for the COLINFO record.
    pub fn set_column_width(
        &mut self,
        sheet: usize,
        column: crate::writer::Column,
        width_chars: f64,
    ) -> Result<()> {
        if !(width_chars.is_finite()) || width_chars <= 0.0 {
            return Err(Error::InvalidData(
                "set_column_width: width must be a positive finite value".to_string(),
            ));
        }

        let max_units = 255u32 * 256u32; // Excel maximum column width
        let width_units_f = (width_chars * 256.0).round();
        if width_units_f <= 0.0 || width_units_f > max_units as f64 {
            return Err(Error::InvalidData(
                "set_column_width: width exceeds Excel's maximum (255 characters)".to_string(),
            ));
        }

        let width_units = width_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_column_width(column, width_units);
        Ok(())
    }

    /// Hide a column.
    pub fn hide_column(&mut self, sheet: usize, column: crate::writer::Column) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_column(column);
        Ok(())
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, sheet: usize, column: crate::writer::Column) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_column(column);
        Ok(())
    }

    pub fn merge_cells(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        let range = MergedRange::try_new(first_row, last_row, first_col, last_col)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if worksheet
            .merged_ranges
            .iter()
            .any(|existing| range.overlaps(*existing))
        {
            return Err(Error::InvalidData("merged-cell ranges overlap".to_string()));
        }

        worksheet.add_merged_range(range);

        Ok(())
    }

    /// Configure freeze panes for the specified worksheet.
    ///
    /// The checked counts represent the number of rows/columns at the top/left
    /// that remain frozen.
    pub fn freeze_panes(&mut self, sheet: usize, panes: crate::writer::FrozenPanes) -> Result<()> {
        if panes.is_empty() {
            let worksheet = self
                .worksheets
                .get_mut(sheet)
                .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
            worksheet.clear_freeze_panes();
            return Ok(());
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.set_freeze_panes(panes)
    }

    /// Remove any freeze panes from the specified worksheet.
    pub fn unfreeze_panes(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.clear_freeze_panes();
        Ok(())
    }

    /// Replace the worksheet's checked BIFF8 zoom scale.
    pub fn put_scale(
        &mut self,
        sheet: usize,
        scale: Option<crate::writer::view::Scale>,
    ) -> Result<Option<crate::writer::view::Scale>> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(worksheet.put_scale(scale))
    }

    /// Replace a worksheet view after validating the complete prospective state.
    pub fn put_view(
        &mut self,
        sheet: usize,
        view: crate::writer::view::View,
    ) -> Result<crate::writer::view::View> {
        view.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        Ok(std::mem::replace(&mut worksheet.view, view))
    }

    /// Replace a worksheet pane and its selections as one failure-atomic edit.
    pub fn put_pane(
        &mut self,
        sheet: usize,
        pane: crate::writer::view::Pane,
        selections: Vec<crate::writer::view::Selection>,
    ) -> Result<(
        Option<crate::writer::view::Pane>,
        Vec<crate::writer::view::Selection>,
    )> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.view.put_pane(pane, selections)
    }

    /// Set the height of a row in points.
    ///
    /// The checked row is zero-based (0 = first row), and the height is specified
    /// in typographic points. Internally this is converted to twips
    /// (1/20th of a point) for the BIFF8 ROW record.
    pub fn set_row_height(
        &mut self,
        sheet: usize,
        row: crate::writer::Row,
        height_points: f64,
    ) -> Result<()> {
        if !(height_points.is_finite()) || height_points <= 0.0 {
            return Err(Error::InvalidData(
                "set_row_height: height must be a positive finite value".to_string(),
            ));
        }

        let height_units_f = (height_points * 20.0).round();
        if height_units_f <= 0.0 || height_units_f > u16::MAX as f64 {
            return Err(Error::InvalidData(
                "set_row_height: height exceeds BIFF8 row height limit".to_string(),
            ));
        }

        let height_units = height_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_row_height(row, height_units);
        Ok(())
    }

    /// Hide a row.
    pub fn hide_row(&mut self, sheet: usize, row: crate::writer::Row) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_row(row);
        Ok(())
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, sheet: usize, row: crate::writer::Row) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_row(row);
        Ok(())
    }

    /// Add a data validation rule to the specified worksheet.
    pub fn add_data_validation(&mut self, sheet: usize, validation: DataValidation) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let range = validation.range;
        worksheet.add_data_validation(
            validation,
            payload,
            vec![range],
            DataValidationOptions::default(),
        );

        Ok(())
    }

    /// Add a validation with typed flags and additional target ranges.
    pub fn add_data_validation_with_options(
        &mut self,
        sheet: usize,
        validation: DataValidation,
        additional_ranges: &[DataValidationRange],
        options: DataValidationOptions,
    ) -> Result<()> {
        let payload = prepare_data_validation(&validation)?;
        let range_count = 1usize
            .checked_add(additional_ranges.len())
            .ok_or_else(|| Error::InvalidData("DV range count overflows".to_string()))?;
        if range_count > 432 {
            return Err(Error::InvalidData("DV range count exceeds 432".to_string()));
        }
        let mut ranges = Vec::with_capacity(range_count);
        ranges.push(validation.range);
        ranges.extend_from_slice(additional_ranges);
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.add_data_validation(validation, payload, ranges, options);
        Ok(())
    }

    /// Configure worksheet-level DVAL window/dropdown metadata.
    pub fn set_data_validation_table_options(
        &mut self,
        sheet: usize,
        options: DataValidationTableOptions,
    ) -> Result<()> {
        if options.x_left > 65_535
            || options.y_top > 65_535
            || matches!(options.dropdown_object_id, Some(0))
            || options.dropdown_object_id.is_some_and(|id| id > 32_767)
        {
            return Err(Error::InvalidData(
                "DVAL metadata is out of range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.data_validation_table_options = Some(options);
        Ok(())
    }

    pub fn add_conditional_format(&mut self, sheet: usize, cf: ConditionalFormat) -> Result<()> {
        if cf.first_row > cf.last_row
            || cf.first_col > cf.last_col
            || cf.last_row > 65_535
            || cf.last_col > 255
        {
            return Err(Error::InvalidData(
                "add_conditional_format: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_conditional_format(cf);

        Ok(())
    }

    /// Add one legacy `CONDFMT` collection with ordered ranges and one to three ordered rules.
    pub fn add_conditional_format_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormatGroup,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > 3 {
            return Err(Error::InvalidData(
                "legacy conditional-format rule count must be 1..=3".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        for rule in &group.rules {
            rule.format_type.to_biff_payload()?;
        }
        self.worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?
            .add_conditional_format_group(group);
        Ok(())
    }

    /// Add one future `CondFmt12` collection. Formula tokens and visual
    /// payloads are serialized exactly and are never evaluated.
    pub fn add_conditional_format12_group(
        &mut self,
        sheet: usize,
        group: ConditionalFormat12Group,
    ) -> Result<()> {
        if group.ranges.is_empty() || group.ranges.len() > 1026 {
            return Err(Error::InvalidData(
                "future conditional-format range count must be 1..=1026".to_string(),
            ));
        }
        if group.rules.is_empty() || group.rules.len() > usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "future conditional-format rule count must be 1..=65535".to_string(),
            ));
        }
        for range in &group.ranges {
            if range.first_row > range.last_row
                || range.first_col > range.last_col
                || range.last_row > 65_535
                || range.last_col > 255
            {
                return Err(Error::InvalidData(
                    "future conditional-format range is outside BIFF8 bounds".to_string(),
                ));
            }
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.conditional_formats.len() + worksheet.conditional_formats12.len() >= 32_768 {
            return Err(Error::InvalidData(
                "conditional-format group count exceeds the 15-bit BIFF identifier space"
                    .to_string(),
            ));
        }
        let mut priorities = worksheet
            .conditional_formats12
            .iter()
            .flat_map(|existing| existing.rules.iter().map(|rule| rule.priority))
            .collect::<HashSet<_>>();
        for rule in &group.rules {
            if rule.priority == 0 || !priorities.insert(rule.priority) {
                return Err(Error::InvalidData(
                    "future conditional-format priorities must be nonzero and unique per sheet"
                        .to_string(),
                ));
            }
            if !matches!(rule.template, 0..=5 | 7..=12 | 15..=27 | 29 | 30) {
                return Err(Error::InvalidData(
                    "future conditional-format template is invalid".to_string(),
                ));
            }
            let between = matches!(
                rule.format_type,
                ConditionalFormat12Type::CellValue {
                    operator: ConditionalFormatOperator::Between
                        | ConditionalFormatOperator::NotBetween,
                    ..
                }
            );
            if let ConditionalFormat12Type::CellValue { formula2, .. } = &rule.format_type
                && between != formula2.is_some()
            {
                return Err(Error::InvalidData(
                        "between/not-between CF12 rules require two formulas; other comparisons require one".to_string(),
                    ));
            }
            let visual = matches!(
                rule.format_type,
                ConditionalFormat12Type::ColorScale { .. }
                    | ConditionalFormat12Type::DataBar { .. }
                    | ConditionalFormat12Type::IconSet { .. }
            );
            if visual && (rule.stop_if_true || rule.differential_format != [0, 0, 0, 0, 0, 0]) {
                return Err(Error::InvalidData(
                    "visual CF12 rules require an empty DXFN12 and cannot stop-if-true".to_string(),
                ));
            }
            let (condition_type, comparison, formula1, formula2, active_formula, payload) =
                rule.format_type.biff_parts();
            let config = crate::writer::biff::Cf12Config {
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
            crate::writer::biff::write_cf12(&mut Vec::new(), &config)?;
        }
        worksheet.add_conditional_format12_group(group);
        Ok(())
    }

    /// Set a worksheet's tab color as a BIFF8 palette index (SHEETEXT
    /// `icvPlain`, MS-XLS 2.4.259). Valid indices are 0x08 through 0x3F;
    /// `None` clears an explicitly set color.
    pub fn set_worksheet_tab_color(&mut self, sheet: usize, tab_color: Option<u8>) -> Result<()> {
        if let Some(index) = tab_color
            && !(0x08..=0x3F).contains(&index)
        {
            return Err(Error::InvalidData(format!(
                "sheet tab color index {index:#04X} is outside the Icv palette"
            )));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.tab_color = tab_color;
        Ok(())
    }

    /// Author a what-if data table (MS-XLS 2.4.319) anchored at a formula
    /// cell. The anchor cell is written as a `PtgTbl` formula immediately
    /// followed by the `Table` record; it must lie outside the table range
    /// and must not already carry a value.
    pub fn add_data_table(
        &mut self,
        sheet: usize,
        anchor_row: u32,
        anchor_col: u16,
        table: crate::DataTable,
    ) -> Result<()> {
        let anchor_pos = CellPos::try_new(anchor_row, anchor_col)?;
        let range = table.range();
        let inside = (u32::from(range.first_row())..=u32::from(range.last_row()))
            .contains(&anchor_row)
            && (range.first_col()..=range.last_col()).contains(&anchor_pos.col());
        if inside {
            return Err(Error::InvalidData(
                "data-table anchor formula cell must lie outside the table range".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet
            .data_tables
            .iter()
            .any(|(row, col, _)| (*row, *col) == (anchor_row, anchor_col))
        {
            return Err(Error::InvalidData(
                "duplicate data-table anchor cell".to_string(),
            ));
        }
        if worksheet.cells.iter().any(|(&(row, col), cell)| {
            (u32::from(range.first_row())..=u32::from(range.last_row())).contains(&row)
                && (u16::from(range.first_col())..=u16::from(range.last_col())).contains(&col)
                && cell
                    .formula_metadata
                    .as_ref()
                    .is_some_and(|metadata| metadata.array_owner().is_some())
        }) {
            return Err(Error::InvalidData(
                "data-table range overlaps an array-formula group".to_string(),
            ));
        }
        let anchor_missing = if let Some(cell) = worksheet.cells.get(&(anchor_row, anchor_col)) {
            if !matches!(cell.value, CellValue::Blank) {
                return Err(Error::InvalidData(
                    "data-table anchor cell already carries a value".to_string(),
                ));
            }
            false
        } else {
            true
        };
        worksheet
            .data_tables
            .try_reserve(1)
            .map_err(|_| Error::Allocation("reserving worksheet data-table storage"))?;
        if anchor_missing {
            worksheet
                .cells
                .try_reserve(1)
                .map_err(|_| Error::Allocation("reserving data-table anchor cell"))?;
            worksheet.add_cell(WritableCell::new(anchor_pos, CellValue::Blank, 0, None));
        }
        worksheet.data_tables.push((anchor_row, anchor_col, table));
        Ok(())
    }

    pub fn set_worksheet_vba_code_name(
        &mut self,
        sheet: usize,
        code_name: Option<&str>,
    ) -> Result<()> {
        if self.vba_metadata.is_none() && code_name.is_some() {
            return Err(Error::InvalidData(
                "worksheet VBA code names require an enabled VBA project".to_string(),
            ));
        }
        if let Some(value) = code_name {
            crate::vba::validate_code_name(value)?;
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.vba_code_name = code_name.map(str::to_string);
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_worksheet_layout(
        &mut self,
        sheet: usize,
        options: WorksheetLayoutOptions,
    ) -> Result<()> {
        options.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_layout = options;
        Ok(())
    }

    /// Configure the complete primary worksheet print/page settings block.
    pub fn set_page_setup(&mut self, sheet: usize, options: PageSetupOptions) -> Result<()> {
        let valid_margin = |value: f64| value.is_finite() && (0.0..49.0).contains(&value);
        if !valid_margin(options.left_margin_inches)
            || !valid_margin(options.right_margin_inches)
            || !valid_margin(options.top_margin_inches)
            || !valid_margin(options.bottom_margin_inches)
            || !valid_margin(options.header_margin_inches)
            || !valid_margin(options.footer_margin_inches)
        {
            return Err(Error::InvalidData(
                "page margins must be finite and between 0 and 49 inches".to_string(),
            ));
        }
        if options.header.encode_utf16().count() > 255
            || options.footer.encode_utf16().count() > 255
        {
            return Err(Error::InvalidData(
                "header and footer must not exceed 255 UTF-16 code units".to_string(),
            ));
        }
        if (118..=255).contains(&options.paper_size)
            || !(10..=400).contains(&options.scale_percent)
            || options.fit_width_pages > 32767
            || options.fit_height_pages > 32767
            || options.horizontal_resolution_dpi == 0
            || options.vertical_resolution_dpi == 0
            || options.copies == 0
            || options.copies > 32767
        {
            return Err(Error::InvalidData(
                "page setup contains an out-of-range dimension".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = Some(options);
        Ok(())
    }

    pub fn clear_page_setup(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.page_setup = None;
        worksheet.horizontal_page_breaks.clear();
        worksheet.vertical_page_breaks.clear();
        Ok(())
    }

    /// Add a horizontal break at the first row below the break.
    pub fn add_horizontal_page_break(
        &mut self,
        sheet: usize,
        row: u32,
        col_start: u16,
        col_end: u16,
    ) -> Result<()> {
        let page_break = HorizontalPageBreak::try_new(row, col_start, col_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.horizontal_page_breaks.len() >= 1026 {
            return Err(Error::InvalidData(
                "horizontal page-break count exceeds 1026".to_string(),
            ));
        }
        if worksheet
            .horizontal_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "horizontal page-break ranges overlap".to_string(),
            ));
        }
        worksheet.horizontal_page_breaks.push(page_break);
        Ok(())
    }

    /// Add a vertical break at the first column right of the break.
    pub fn add_vertical_page_break(
        &mut self,
        sheet: usize,
        column: u16,
        row_start: u32,
        row_end: u32,
    ) -> Result<()> {
        let page_break = VerticalPageBreak::try_new(column, row_start, row_end)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        if worksheet.vertical_page_breaks.len() >= 255 {
            return Err(Error::InvalidData(
                "vertical page-break count exceeds 255".to_string(),
            ));
        }
        if worksheet
            .vertical_page_breaks
            .iter()
            .any(|existing| page_break.overlaps(*existing))
        {
            return Err(Error::InvalidData(
                "vertical page-break ranges overlap".to_string(),
            ));
        }
        worksheet.vertical_page_breaks.push(page_break);
        Ok(())
    }
}
