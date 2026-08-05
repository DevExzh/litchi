//! Table and cell editing semantics.

#![allow(unused_imports)]

use super::*;

impl NumbersEditor {
    /// List absolute pivot categories backed by valid calculation-engine
    /// aggregate coordinates.
    pub fn pivot_categories(&self) -> Result<Vec<NumbersPivotCategoryInfo>> {
        let mut categories = formula_pivot_categories(&self.package)?
            .into_iter()
            .map(|(key, value)| NumbersPivotCategoryInfo {
                reference: FormulaPivotCategoryReference::new(
                    key.group_by_uid,
                    key.column_uid,
                    key.group_uid,
                    value.aggregate_type,
                    value.group_level,
                ),
                label: value.label,
            })
            .collect::<Vec<_>>();
        categories.sort_by(|left, right| {
            left.reference
                .group_by_uid
                .cmp(&right.reference.group_by_uid)
                .then_with(|| left.reference.column_uid.cmp(&right.reference.column_uid))
                .then_with(|| left.reference.group_level.cmp(&right.reference.group_level))
                .then_with(|| left.label.cmp(&right.label))
                .then_with(|| left.reference.group_uid.cmp(&right.reference.group_uid))
        });
        Ok(categories)
    }

    /// Set or clear a cell in a table identified by its IWA object ID.
    ///
    /// Cached results of dependent numeric/Boolean formulas are refreshed in
    /// dependency order. If an impacted formula is outside the strict local
    /// evaluator subset, the entire edit is rejected without changing the
    /// package rather than persisting a stale displayed result.
    pub fn set_cell(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        value: CellValue,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_cell_in_package(&mut staged, table_id, row, column, value)?;
        formula_cache::refresh_formula_caches_after_cell_write(&mut staged, table_id, row, column)?;
        // Exercise every serialization boundary before committing the edit.
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Set several cells in one table as one transaction.
    ///
    /// The batch must contain unique coordinates. It clones and serializes the
    /// package once, reuses one table/object lookup context for every cell, and
    /// refreshes all impacted formula caches from the final batch state in one
    /// dependency pass. The returned count equals the number of applied cells.
    pub fn set_cells(
        &mut self,
        table_id: u64,
        updates: impl IntoIterator<Item = TableCellUpdate>,
    ) -> Result<usize> {
        let batch = table_cells::TableCellBatch::collect(updates)?;
        if batch.is_empty() {
            attached_table_descriptor(&self.package, table_id)?;
            return Ok(0);
        }
        let expected = batch.len();
        let mut staged = self.package.clone();
        let applied = batch.apply_numbers(&mut staged, table_id)?;
        if applied != expected {
            return Err(Error::InvalidFormat(format!(
                "Table cell batch applied {applied} updates, expected {expected}"
            )));
        }
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(applied)
    }

    pub fn clear_cell(&mut self, table_id: u64, row: usize, column: usize) -> Result<()> {
        self.set_cell(table_id, row, column, CellValue::Empty)
    }

    /// Read the explicit data format for one zero-based table cell.
    pub fn table_cell_data_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<DataFormat> {
        cell_data_format::cell_data_format(&self.package, table_id, row, column)
    }

    /// Create, replace, or reset one cell's typed data format transactionally.
    pub fn set_table_cell_data_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: DataFormat,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_data_format::set_cell_data_format(&mut staged, table_id, row, column, &format)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_data_format(table_id, row, column)? != format {
            return Err(Error::InvalidFormat(
                "Numbers table-cell data format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read an explicit decimal-number format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_number_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Number>> {
        cell_data_format::cell_number_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit decimal-number format transactionally.
    pub fn set_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Number,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore iWork's automatic data format for one table cell.
    pub fn reset_table_cell_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_number_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified
                .table_cell_number_format(table_id, row, column)?
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "Numbers table-cell number-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Text format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_text_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Text>> {
        cell_data_format::cell_text_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Text format transactionally.
    pub fn set_table_cell_text_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_table_cell_data_format(
            table_id,
            row,
            column,
            Text.into(),
        )
    }

    /// Restore Automatic from an explicit Text cell.
    pub fn reset_table_cell_text_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_data_format::reset_cell_text_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Text-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read a named custom Number, Date & Time, or Text format.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_custom_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Custom>> {
        cell_data_format::cell_custom_format(&self.package, table_id, row, column)
    }

    /// Create or replace a named custom format transactionally.
    pub fn set_table_cell_custom_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Custom,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from a named custom format.
    pub fn reset_table_cell_custom_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_custom_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Custom-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit currency format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_currency_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Currency>> {
        cell_data_format::cell_currency_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit currency format transactionally.
    pub fn set_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Currency,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Currency cell.
    pub fn reset_table_cell_currency_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_currency_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers currency-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit percentage format for one zero-based table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_percentage_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Percentage>> {
        cell_data_format::cell_percentage_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit percentage format transactionally.
    pub fn set_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Percentage,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore iWork's automatic format from an explicit Percentage cell.
    pub fn reset_table_cell_percentage_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_percentage_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers percentage-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit scientific-notation format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_scientific_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Scientific>> {
        cell_data_format::cell_scientific_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit scientific-notation format transactionally.
    pub fn set_table_cell_scientific_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Scientific,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Scientific cell.
    pub fn reset_table_cell_scientific_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_scientific_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers scientific-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit mixed-fraction format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_fraction_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Fraction>> {
        cell_data_format::cell_fraction_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit mixed-fraction format transactionally.
    pub fn set_table_cell_fraction_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Fraction,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Fraction cell.
    pub fn reset_table_cell_fraction_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_fraction_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers fraction-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit positional numeral-system format for one table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_numeral_system_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<NumeralSystem>> {
        cell_data_format::cell_numeral_system_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit positional numeral-system format transactionally.
    pub fn set_table_cell_numeral_system_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: NumeralSystem,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Numeral System cell.
    pub fn reset_table_cell_numeral_system_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_numeral_system_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers numeral-system reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Date & Time format for one table cell.
    ///
    /// `None` means the Date value uses iWork's automatic data format.
    pub fn table_cell_date_time_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<DateTime>> {
        cell_data_format::cell_date_time_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Date & Time format transactionally.
    pub fn set_table_cell_date_time_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: DateTime,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Date & Time cell.
    pub fn reset_table_cell_date_time_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_date_time_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Date & Time reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Duration format for one table cell.
    ///
    /// `None` means the Duration value uses iWork's automatic data format.
    pub fn table_cell_duration_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Duration>> {
        cell_data_format::cell_duration_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit Duration format transactionally.
    pub fn set_table_cell_duration_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Duration,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Duration cell.
    pub fn reset_table_cell_duration_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_duration_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Duration reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Checkbox format for one table cell.
    pub fn table_cell_checkbox_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Checkbox>> {
        cell_data_format::cell_checkbox_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit native Checkbox format transactionally.
    pub fn set_table_cell_checkbox_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Checkbox,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Checkbox cell.
    pub fn reset_table_cell_checkbox_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_checkbox_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Checkbox reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Star Rating format for one table cell.
    pub fn table_cell_star_rating_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<StarRating>> {
        cell_data_format::cell_star_rating_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit native five-star rating transactionally.
    pub fn set_table_cell_star_rating_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: StarRating,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Star Rating cell.
    pub fn reset_table_cell_star_rating_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_star_rating_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Star Rating reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Slider format for one table cell.
    pub fn table_cell_slider_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Slider>> {
        cell_data_format::cell_slider_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit native Slider format transactionally.
    pub fn set_table_cell_slider_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Slider,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Slider cell.
    pub fn reset_table_cell_slider_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_slider_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Slider reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Stepper format for one table cell.
    pub fn table_cell_stepper_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Stepper>> {
        cell_data_format::cell_stepper_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit native Stepper format transactionally.
    pub fn set_table_cell_stepper_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: Stepper,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Stepper cell.
    pub fn reset_table_cell_stepper_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_stepper_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Stepper reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Pop-Up Menu format for one table cell.
    pub fn table_cell_pop_up_menu_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<PopUpMenu>> {
        cell_data_format::cell_pop_up_menu_format(&self.package, table_id, row, column)
    }

    /// Create or replace an explicit native Pop-Up Menu format transactionally.
    pub fn set_table_cell_pop_up_menu_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        format: PopUpMenu,
    ) -> Result<()> {
        self.set_table_cell_data_format(table_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Pop-Up Menu cell.
    pub fn reset_table_cell_pop_up_menu_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_data_format::reset_cell_pop_up_menu_format(&mut staged, table_id, row, column)?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            if verified.table_cell_data_format(table_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Numbers Pop-Up Menu reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective text layout for one zero-based table cell.
    pub fn table_cell_layout(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::table_cell_layout::TableCellLayout> {
        cell_layout::cell_layout(&self.package, table_id, row, column)
    }

    /// Create or replace local text-layout overrides for one table cell.
    pub fn set_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        layout: crate::table_cell_layout::TableCellLayout,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_layout::set_cell_layout(&mut staged, table_id, row, column, layout)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_layout(table_id, row, column)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers table-cell layout failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local text-layout overrides and restore inherited cell values.
    pub fn reset_table_cell_layout(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_layout::reset_cell_layout(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective horizontal text alignment for one zero-based table cell.
    pub fn table_cell_text_alignment(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextAlignment> {
        cell_paragraph_style::alignment(&self.package, table_id, row, column)
    }

    /// Create or replace a local horizontal text-alignment override.
    pub fn set_table_cell_text_alignment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        alignment: NumbersTableCellTextAlignment,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_alignment(&mut staged, table_id, row, column, alignment)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_alignment(table_id, row, column)? != alignment {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text alignment failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local horizontal alignment and restore the inherited table style.
    pub fn reset_table_cell_text_alignment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_alignment(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph line spacing for one table cell.
    pub fn table_cell_paragraph_line_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphLineSpacing> {
        cell_paragraph_style::line_spacing(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell paragraph line-spacing override.
    pub fn set_table_cell_paragraph_line_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: NumbersTableCellParagraphLineSpacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_line_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_line_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph line spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local line spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_line_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_line_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing for one table cell.
    pub fn table_cell_paragraph_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphSpacing> {
        cell_paragraph_style::spacing(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell before/after paragraph spacing.
    pub fn set_table_cell_paragraph_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: NumbersTableCellParagraphSpacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local before/after spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the canonical list preset applied uniformly to a table cell.
    pub fn table_cell_paragraph_list(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphList> {
        cell_paragraph_list::paragraph_list(&self.package, table_id, row, column)
    }

    /// Promote a plain text cell when necessary and apply one native list preset.
    pub fn set_table_cell_paragraph_list(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        list: NumbersTableCellParagraphList,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list(&mut staged, table_id, row, column, list)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list(table_id, row, column)? != list {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the canonical None list preset for a table cell.
    pub fn reset_table_cell_paragraph_list(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_list::reset_paragraph_list(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read all paragraph-scoped list preset boundaries in a table cell.
    pub fn table_cell_paragraph_lists(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<NumbersTableCellParagraphListPlacement>> {
        cell_paragraph_list::paragraph_lists(&self.package, table_id, row, column)
    }

    /// Promote a plain cell when necessary and replace all list preset boundaries.
    pub fn set_table_cell_paragraph_lists(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        placements: &[NumbersTableCellParagraphListPlacement],
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_lists(&mut staged, table_id, row, column, placements)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let expected = cell_paragraph_list::paragraph_lists(&staged, table_id, row, column)?;
        if verified.table_cell_paragraph_lists(table_id, row, column)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph-list placements failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read every effective list-level boundary in a table cell.
    pub fn table_cell_paragraph_list_levels(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<NumbersTableCellParagraphListLevelPlacement>> {
        cell_paragraph_list::paragraph_list_levels(&self.package, table_id, row, column)
    }

    /// Set one validated paragraph's list level without changing later paragraphs.
    pub fn set_table_cell_paragraph_list_level(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        level: NumbersTableCellParagraphListLevel,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_level(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            level,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if cell_paragraph_list::paragraph_list_level(
            &verified.package,
            table_id,
            row,
            column,
            paragraph,
        )? != level
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list level failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_table_cell_paragraph_list_level(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_level(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read whether one table-cell paragraph continues or restarts list numbering.
    pub fn table_cell_paragraph_list_numbering(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListNumbering> {
        cell_paragraph_list::paragraph_list_numbering(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Continue or restart numbered-list sequencing at one table-cell paragraph.
    pub fn set_table_cell_paragraph_list_numbering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        numbering: NumbersTableCellParagraphListNumbering,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_numbering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            numbering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if cell_paragraph_list::paragraph_list_numbering(
            &verified.package,
            table_id,
            row,
            column,
            paragraph,
        )? != numbering
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list numbering failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one numbered table-cell paragraph's effective label format.
    pub fn table_cell_paragraph_list_number_format(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListNumberFormat> {
        cell_paragraph_list::paragraph_list_number_format(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered table-cell paragraph's locale-aware label format.
    pub fn set_table_cell_paragraph_list_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        format: NumbersTableCellParagraphListNumberFormat,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_format(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_format(table_id, row, column, paragraph)?
            != format
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number format failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_table_cell_paragraph_list_number_format(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_format(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read whether one numbered table-cell paragraph displays hierarchical numbering.
    pub fn table_cell_paragraph_list_number_tiering(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListNumberTiering> {
        cell_paragraph_list::paragraph_list_number_tiering(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Choose flat or hierarchical numbering for one table-cell list level.
    pub fn set_table_cell_paragraph_list_number_tiering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        tiering: NumbersTableCellParagraphListNumberTiering,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_tiering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            tiering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_tiering(table_id, row, column, paragraph)?
            != tiering
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number tiering failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore flat numbering for one table-cell list level.
    pub fn reset_table_cell_paragraph_list_number_tiering(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_tiering(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one numbered table-cell paragraph's number-label size.
    pub fn table_cell_paragraph_list_number_scale(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListNumberScale> {
        cell_paragraph_list::paragraph_list_number_scale(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered table-cell paragraph's number-label size.
    pub fn set_table_cell_paragraph_list_number_scale(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        scale: NumbersTableCellParagraphListNumberScale,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_number_scale(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            scale,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_number_scale(table_id, row, column, paragraph)?
            != scale
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-number scale failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_table_cell_paragraph_list_number_scale(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_number_scale(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell paragraph's effective text-bullet marker.
    pub fn table_cell_paragraph_list_bullet(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListBullet> {
        cell_paragraph_list::paragraph_list_bullet(&self.package, table_id, row, column, paragraph)
    }

    /// Set one table-cell paragraph's text-bullet marker.
    pub fn set_table_cell_paragraph_list_bullet(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        bullet: &NumbersTableCellParagraphListBullet,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_bullet(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            bullet,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_bullet(table_id, row, column, paragraph)? != *bullet {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph text bullet failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one table-cell paragraph.
    pub fn reset_table_cell_paragraph_list_bullet(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_bullet(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell paragraph's effective bullet size and baseline.
    pub fn table_cell_paragraph_list_bullet_geometry(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListBulletGeometry> {
        cell_paragraph_list::paragraph_list_bullet_geometry(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell paragraph's bullet size and baseline.
    pub fn set_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        geometry: NumbersTableCellParagraphListBulletGeometry,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_bullet_geometry(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            geometry,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_bullet_geometry(table_id, row, column, paragraph)?
            != geometry
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph bullet geometry failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_bullet_geometry(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell list paragraph's label and text-gap indentation.
    pub fn table_cell_paragraph_list_indentation(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListIndentation> {
        cell_paragraph_list::paragraph_list_indentation(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell list paragraph's label and text-gap indentation.
    pub fn set_table_cell_paragraph_list_indentation(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        indentation: NumbersTableCellParagraphListIndentation,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_indentation(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            indentation,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_indentation(table_id, row, column, paragraph)?
            != indentation
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list indentation failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_table_cell_paragraph_list_indentation(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_indentation(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read one table-cell list paragraph's effective label color.
    pub fn table_cell_paragraph_list_label_color(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<NumbersTableCellParagraphListLabelColor> {
        cell_paragraph_list::paragraph_list_label_color(
            &self.package,
            table_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one table-cell list paragraph's bullet or number color.
    pub fn set_table_cell_paragraph_list_label_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
        color: NumbersTableCellParagraphListLabelColor,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_list::set_paragraph_list_label_color(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
            color,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_list_label_color(table_id, row, column, paragraph)?
            != color
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph list-label color failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_table_cell_paragraph_list_label_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        paragraph: ParagraphStart,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_list::reset_paragraph_list_label_color(
            &mut staged,
            table_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right paragraph indents.
    pub fn table_cell_paragraph_indents(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphIndents> {
        cell_paragraph_style::indents(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell paragraph indents.
    pub fn set_table_cell_paragraph_indents(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        indents: NumbersTableCellParagraphIndents,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_indents(&mut staged, table_id, row, column, indents)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_indents(table_id, row, column)? != indents {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph indents failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local paragraph indents and restore the inherited table style.
    pub fn reset_table_cell_paragraph_indents(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_indents(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the ordered explicit ruler tab stops for one table cell.
    pub fn table_cell_paragraph_tab_stops(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellParagraphTabStops> {
        cell_paragraph_style::tab_stops(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell ruler tab stops.
    pub fn set_table_cell_paragraph_tab_stops(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        stops: NumbersTableCellParagraphTabStops,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_tab_stops(&mut staged, table_id, row, column, stops.clone())?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_paragraph_tab_stops(table_id, row, column)? != stops {
            return Err(Error::InvalidFormat(
                "Numbers table-cell paragraph tab stops failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local ruler tab stops and restore the inherited table style.
    pub fn reset_table_cell_paragraph_tab_stops(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_tab_stops(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective background painted behind one table cell's text.
    pub fn table_cell_text_background(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextBackground> {
        cell_paragraph_style::background(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text-background override.
    pub fn set_table_cell_text_background(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        background: NumbersTableCellTextBackground,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_background(&mut staged, table_id, row, column, background)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_background(table_id, row, column)? != background {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text background failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text background and restore the inherited value.
    pub fn reset_table_cell_text_background(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_background(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of one table cell.
    pub fn table_cell_text_baseline_shift(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextBaselineShift> {
        cell_paragraph_style::baseline_shift(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell custom baseline displacement.
    pub fn set_table_cell_text_baseline_shift(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        shift: NumbersTableCellTextBaselineShift,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_baseline_shift(&mut staged, table_id, row, column, shift)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_baseline_shift(table_id, row, column)? != shift {
            return Err(Error::InvalidFormat(
                "Numbers table-cell baseline shift failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local baseline displacement and restore the inherited value.
    pub fn reset_table_cell_text_baseline_shift(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_baseline_shift(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective capitalization of one table cell.
    pub fn table_cell_text_capitalization(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextCapitalization> {
        cell_paragraph_style::capitalization(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell capitalization.
    pub fn set_table_cell_text_capitalization(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        capitalization: NumbersTableCellTextCapitalization,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_capitalization(
            &mut staged,
            table_id,
            row,
            column,
            capitalization,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_capitalization(table_id, row, column)? != capitalization {
            return Err(Error::InvalidFormat(
                "Numbers table-cell capitalization failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local capitalization and restore the inherited value.
    pub fn reset_table_cell_text_capitalization(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_capitalization(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of one table cell.
    pub fn table_cell_text_character_spacing(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextCharacterSpacing> {
        cell_paragraph_style::character_spacing(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell character spacing.
    pub fn set_table_cell_text_character_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        spacing: NumbersTableCellTextCharacterSpacing,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_character_spacing(&mut staged, table_id, row, column, spacing)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_character_spacing(table_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers table-cell character spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local character spacing and restore the inherited value.
    pub fn reset_table_cell_text_character_spacing(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed =
            cell_paragraph_style::reset_character_spacing(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective foreground text color of one table cell.
    pub fn table_cell_text_color(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextColor> {
        cell_paragraph_style::text_color(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell foreground text-color override.
    pub fn set_table_cell_text_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        color: NumbersTableCellTextColor,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_text_color(&mut staged, table_id, row, column, color)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_color(table_id, row, column)? != color {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text color failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text-color override and restore the inherited color.
    pub fn reset_table_cell_text_color(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_text_color(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective whole-cell underline and strikethrough formatting.
    pub fn table_cell_text_decorations(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextDecorations> {
        cell_paragraph_style::decorations(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell underline and strikethrough formatting.
    pub fn set_table_cell_text_decorations(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        decorations: NumbersTableCellTextDecorations,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_decorations(&mut staged, table_id, row, column, decorations)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_decorations(table_id, row, column)? != decorations {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text decorations failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local decorations and restore the inherited cell formatting.
    pub fn reset_table_cell_text_decorations(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_decorations(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of one table cell.
    pub fn table_cell_text_font(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextFont> {
        cell_paragraph_style::font(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell PostScript font override.
    pub fn set_table_cell_text_font(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        font: NumbersTableCellTextFont,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_font(&mut staged, table_id, row, column, font.clone())?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_font(table_id, row, column)? != font {
            return Err(Error::InvalidFormat(
                "Numbers table-cell font failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local font override and restore the inherited table font.
    pub fn reset_table_cell_text_font(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_font(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of one table cell.
    pub fn table_cell_text_ligatures(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextLigatures> {
        cell_paragraph_style::ligatures(&self.package, table_id, row, column)
    }

    /// Create or replace the whole-cell ligature policy.
    pub fn set_table_cell_text_ligatures(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        ligatures: NumbersTableCellTextLigatures,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_ligatures(&mut staged, table_id, row, column, ligatures)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_ligatures(table_id, row, column)? != ligatures {
            return Err(Error::InvalidFormat(
                "Numbers table-cell ligatures failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local ligature policy and restore the inherited value.
    pub fn reset_table_cell_text_ligatures(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_ligatures(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective outline of one table cell's text.
    pub fn table_cell_text_outline(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextOutline> {
        cell_paragraph_style::outline(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text outline.
    pub fn set_table_cell_text_outline(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        outline: NumbersTableCellTextOutline,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_outline(&mut staged, table_id, row, column, outline)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_outline(table_id, row, column)? != outline {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text outline failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text outline and restore the inherited value.
    pub fn reset_table_cell_text_outline(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_outline(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective normal, superscript, or subscript formatting.
    pub fn table_cell_text_script(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextScript> {
        cell_paragraph_style::script(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell baseline script formatting.
    pub fn set_table_cell_text_script(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        script: NumbersTableCellTextScript,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_script(&mut staged, table_id, row, column, script)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_script(table_id, row, column)? != script {
            return Err(Error::InvalidFormat(
                "Numbers table-cell baseline script failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local baseline script formatting and restore the inherited value.
    pub fn reset_table_cell_text_script(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_script(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective drop shadow of one table cell's text.
    pub fn table_cell_text_shadow(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextShadow> {
        cell_paragraph_style::shadow(&self.package, table_id, row, column)
    }

    /// Create or replace a whole-cell text drop shadow.
    pub fn set_table_cell_text_shadow(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        shadow: NumbersTableCellTextShadow,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_shadow(&mut staged, table_id, row, column, shadow)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_shadow(table_id, row, column)? != shadow {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text shadow failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text shadow and restore the inherited value.
    pub fn reset_table_cell_text_shadow(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_shadow(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read effective whole-cell point size, bold, and italic formatting.
    pub fn table_cell_text_style(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<NumbersTableCellTextStyle> {
        cell_paragraph_style::text_style(&self.package, table_id, row, column)
    }

    /// Create or replace whole-cell point size, bold, and italic formatting.
    pub fn set_table_cell_text_style(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        style: NumbersTableCellTextStyle,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_paragraph_style::set_text_style(&mut staged, table_id, row, column, style)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_text_style(table_id, row, column)? != style {
            return Err(Error::InvalidFormat(
                "Numbers table-cell text style failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local point size, bold, and italic formatting.
    pub fn reset_table_cell_text_style(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_paragraph_style::reset_text_style(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective fill for one zero-based table cell.
    pub fn table_cell_fill(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::shapes::ShapeFill> {
        cell_fill::cell_fill(&self.package, table_id, row, column)
    }

    /// Create or replace a local table-cell fill transactionally.
    pub fn set_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        fill: &crate::shapes::ShapeFill,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        cell_fill::set_cell_fill(&mut staged, table_id, row, column, fill)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if &verified.table_cell_fill(table_id, row, column)? != fill {
            return Err(Error::InvalidFormat(
                "Numbers table-cell fill failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct fill override and restore the inherited table style.
    pub fn reset_table_cell_fill(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_fill::reset_cell_fill(&mut staged, table_id, row, column)?;
        if changed {
            *self = Self::from_bytes(&staged.to_bytes()?)?;
        }
        Ok(changed)
    }

    /// Read the effective explicit borders for one zero-based table cell.
    pub fn table_cell_borders(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::table_cell_border::TableCellBorders> {
        stroke_layers::cell_borders(&self.package, table_id, row, column)
    }

    /// Create or replace one explicit table-cell border transactionally.
    pub fn set_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
        stroke: crate::shapes::ShapeStroke,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, Some(stroke))
    }

    /// Explicitly clear one table-cell border transactionally.
    pub fn clear_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
    ) -> Result<()> {
        self.update_table_cell_border(table_id, row, column, side, None)
    }

    fn update_table_cell_border(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        side: crate::table_cell_border::TableCellBorderSide,
        stroke: Option<crate::shapes::ShapeStroke>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        stroke_layers::set_cell_border(&mut staged, table_id, row, column, side, stroke)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .table_cell_borders(table_id, row, column)?
            .get(side)
            != stroke
        {
            return Err(Error::InvalidFormat(
                "Numbers table-cell border failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// List every native merged-cell rectangle in one attached table.
    pub fn table_cell_merges(&self, table_id: u64) -> Result<Vec<IWorkTableCellRegion>> {
        cell_merge::regions_in_package(&self.package, table_id)
    }

    /// Merge one non-overlapping rectangular cell region transactionally.
    pub fn merge_cells(&mut self, table_id: u64, region: IWorkTableCellRegion) -> Result<()> {
        let mut staged = self.package.clone();
        cell_merge::merge_in_package(&mut staged, table_id, region)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if !verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell merge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove one exact merged-cell rectangle, returning whether it existed.
    pub fn unmerge_cells(&mut self, table_id: u64, region: IWorkTableCellRegion) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = cell_merge::unmerge_in_package(&mut staged, table_id, region)?;
        if !changed {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_cell_merges(table_id)?.contains(&region) {
            return Err(Error::InvalidFormat(
                "Numbers table-cell unmerge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Read the comment attached to a writable BNC cell.
    pub fn cell_comment(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<TableCellComment>> {
        cell_comment_in_package(&self.package, table_id, row, column)
    }

    /// Create or replace a cell comment without changing the cell value or style.
    pub fn set_cell_comment(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_cell_comment_in_package(&mut staged, table_id, row, column, text.into())?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Delete a cell comment without changing the cell value or style.
    pub fn clear_cell_comment(&mut self, table_id: u64, row: usize, column: usize) -> Result<()> {
        let mut staged = self.package.clone();
        clear_cell_comment_in_package(&mut staged, table_id, row, column)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    /// Inspect the conditional-highlight style set attached to a writable BNC cell.
    pub fn cell_conditional_highlighting(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<TableCellConditionalHighlightInfo>> {
        conditional_highlight::info_in_package(&self.package, table_id, row, column)
    }

    /// Read the supported ordered conditional-highlight rules attached to a cell.
    pub fn cell_conditional_highlight_rules(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Vec<Rule>>> {
        conditional_highlight::rules_in_package(&self.package, table_id, row, column)
    }

    /// Delete conditional highlighting from one cell without changing its value or base style.
    pub fn clear_cell_conditional_highlighting(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        conditional_highlight::clear_in_package(&mut staged, table_id, row, column)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .cell_conditional_highlighting(table_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a cell's conditional highlighting and return its storage identity.
    pub fn set_cell_conditional_highlighting(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        rules: &[Rule],
    ) -> Result<TableCellConditionalHighlightInfo> {
        let mut staged = self.package.clone();
        conditional_highlight::set_in_package(&mut staged, table_id, row, column, rules)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let actual = verified
            .cell_conditional_highlighting(table_id, row, column)?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-highlight creation failed validation".to_owned(),
                )
            })?;
        if actual.rule_count as usize != rules.len() {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight rule count failed validation".to_owned(),
            ));
        }
        if verified
            .cell_conditional_highlight_rules(table_id, row, column)?
            .as_deref()
            != Some(rules)
        {
            return Err(Error::InvalidFormat(
                "Numbers conditional-highlight rules failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(actual)
    }

    /// Read the direct replies attached to a cell comment in stored order.
    pub fn cell_comment_replies(
        &self,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<TableCellReply>> {
        cell_comment_replies_in_package(&self.package, table_id, row, column)
    }

    /// Append a direct reply to an existing cell comment.
    pub fn add_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<u64> {
        let mut staged = self.package.clone();
        let reply_id =
            add_cell_comment_reply_in_package(&mut staged, table_id, row, column, text.into())?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Replace one direct reply and return its new copy-on-write object ID.
    pub fn set_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        let mut staged = self.package.clone();
        let reply_id = set_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
            text.into(),
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Remove one direct reply from an existing cell comment.
    pub fn remove_cell_comment_reply(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        remove_cell_comment_reply_in_package(
            &mut staged,
            table_id,
            row,
            column,
            reply_storage_object_id,
        )?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }
}
