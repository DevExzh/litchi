//! Semantic table models and transactional Pages table editing.

use super::*;
use crate::numbers::editor::table::cell::Borders;
use litchi_iwa_common::comment::Comment;
use litchi_iwa_common::table::lock::State as TableLockState;
use litchi_iwa_common::table::cell::BorderSide;
use litchi_numbers::cell::data_format::{
    Checkbox, Currency, Custom, DataFormat, DateTime, Duration, Fraction, Number, NumeralSystem,
    Percentage, PopUpMenu, Scientific, Slider, StarRating, Stepper, Text,
};
use litchi_numbers::table::merge::Region;
use crate::text::{Alignment, Indents, LineSpacing, Spacing};

/// Stable identity and dimensions of one native table attached to the Pages body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesTableInfo {
    /// Object identifier of the body-owned table drawable.
    pub drawable_object_id: u64,
    /// Object identifier accepted by cell-editing APIs.
    pub model_object_id: u64,
    /// UTF-16 body position of the object-replacement character.
    pub anchor_character_index: usize,
    /// Table name stored in the native table model.
    pub name: String,
    /// Number of addressable rows.
    pub rows: usize,
    /// Number of addressable columns.
    pub columns: usize,
    /// Effective alternating-row and automatic-sizing settings.
    pub appearance: TableAppearance,
    /// Interactive editing lock shown in the Arrange inspector.
    pub lock_state: TableLockState,
}

/// Materialized values from one Pages table.
#[derive(Debug, Clone)]
pub struct PagesTable {
    /// Stable identity and dimensions of this table.
    pub info: PagesTableInfo,
    semantic_table: litchi_numbers::Table,
    comments: Box<[((usize, usize), Comment)]>,
    merges: Vec<Region>,
}

impl PagesTable {
    /// Borrow a materialized cell value, or return `None` for an empty cell.
    pub fn get_cell(&self, row: usize, column: usize) -> Option<&PagesCellValue> {
        let position = litchi_numbers::Position::try_from_usize(row, column).ok()?;
        self.semantic_table.get(position)
    }

    /// Iterate over materialized cells without exposing the backing map.
    pub fn iter_cells(&self) -> impl Iterator<Item = ((usize, usize), &PagesCellValue)> + '_ {
        self.semantic_table.iter_cells().map(|cell| {
            (
                (
                    cell.position().row() as usize,
                    cell.position().column() as usize,
                ),
                cell.value(),
            )
        })
    }

    /// Return the number of materialized cells, including explicit empty cells.
    pub fn cell_count(&self) -> usize {
        self.semantic_table.cell_count()
    }

    /// Borrow the comment attached to a materialized cell, if any.
    pub fn get_comment(&self, row: usize, column: usize) -> Option<&Comment> {
        self.comments
            .binary_search_by_key(&(row, column), |(position, _comment)| *position)
            .ok()
            .map(|index| &self.comments[index].1)
    }

    /// Iterate over cell comments without exposing the backing map.
    pub fn iter_comments(
        &self,
    ) -> impl Iterator<Item = ((usize, usize), &Comment)> + '_ {
        self.comments
            .iter()
            .map(|(position, comment)| (*position, comment))
    }

    /// Return the number of materialized cell comments.
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Borrow native merged-cell rectangles in formula-store order.
    pub fn merges(&self) -> &[Region] {
        &self.merges
    }
}

impl PagesEditor {
    /// List native tables anchored in the main body in document order.
    pub fn tables(&self) -> Result<Vec<PagesTableInfo>> {
        Ok(body_table_graphs(self)?
            .into_iter()
            .map(|graph| graph.info)
            .collect())
    }

    /// Read all materialized cell values from one reachable body table.
    pub fn table(&self, model_object_id: u64) -> Result<PagesTable> {
        let info = self
            .tables()?
            .into_iter()
            .find(|table| table.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table model {model_object_id} is not attached to the body"
                ))
            })?;
        let bytes = self.package().to_bytes()?;
        let bundle = Bundle::from_bytes(&bytes)?;
        let index = ObjectIndex::from_bundle(&bundle)?;
        let object = index
            .resolve_ref_id(&bundle, model_object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Pages table model {model_object_id} is missing"))
            })?;
        let table = TableDataExtractor::new(&bundle, &index)
            .extract_table_from_object(&object)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages object {model_object_id} has no native table model"
                ))
            })?;
        let (semantic_table, comments) = table.into_semantic_parts()?;
        Ok(PagesTable {
            info,
            semantic_table,
            comments,
            merges: crate::numbers::editor::table_cell_merges_in_package(
                self.package(),
                model_object_id,
            )?,
        })
    }

    /// Set or clear one cell in a reachable body table transactionally.
    ///
    /// Supported dependent formula caches are refreshed before commit;
    /// unsupported impacted formulas reject the edit instead of remaining
    /// visibly stale.
    pub fn set_table_cell(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        value: PagesCellValue,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            value,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Set several body-table cells with one package clone and dependency pass.
    ///
    /// Coordinates must be unique. Any invalid value, coordinate, or impacted
    /// formula rejects the complete batch without changing the editor.
    pub fn set_table_cells(
        &mut self,
        model_object_id: u64,
        updates: impl IntoIterator<Item = PagesTableCellUpdate>,
    ) -> Result<usize> {
        self.require_body_table(model_object_id)?;
        let batch = crate::numbers::editor::TableCellBatch::collect(updates)?;
        if batch.is_empty() {
            return Ok(0);
        }
        let expected = batch.len();
        let mut staged = self.package().clone();
        let applied = batch.apply_attached(&mut staged, model_object_id)?;
        if applied != expected {
            return Err(Error::InvalidFormat(format!(
                "Pages table-cell batch applied {applied} updates, expected {expected}"
            )));
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        *self = verified;
        Ok(applied)
    }

    /// Clear one cell in a reachable body table.
    pub fn clear_table_cell(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_table_cell(model_object_id, row, column, PagesCellValue::Empty)
    }

    /// Read the explicit typed data format for one body-table cell.
    pub fn table_cell_data_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<DataFormat> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_data_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create, replace, or reset one body-table cell's data format.
    pub fn set_table_cell_data_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: DataFormat,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_data_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            &format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_data_format(model_object_id, row, column)? != format {
            return Err(Error::InvalidFormat(
                "Pages table-cell data format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read an explicit decimal-number format for one body-table cell.
    ///
    /// `None` means the cell uses iWork's automatic data format.
    pub fn table_cell_number_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Number>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::common_table_cell_number_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit decimal-number format transactionally.
    pub fn set_table_cell_number_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Number,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_common_table_cell_number_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_number_format(model_object_id, row, column)? != Some(format) {
            return Err(Error::InvalidFormat(
                "Pages table-cell number format failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore iWork's automatic data format for one body-table cell.
    pub fn reset_table_cell_number_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_number_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified
                .table_cell_number_format(model_object_id, row, column)?
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "Pages table-cell number-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Text format for one body-table cell.
    pub fn table_cell_text_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Text>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Text format transactionally.
    pub fn set_table_cell_text_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.set_table_cell_data_format(
            model_object_id,
            row,
            column,
            Text.into(),
        )
    }

    /// Restore Automatic from an explicit Text body-table cell.
    pub fn reset_table_cell_text_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Text-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read a named custom format for one body-table cell.
    pub fn table_cell_custom_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Custom>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_custom_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a named custom format transactionally.
    pub fn set_table_cell_custom_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Custom,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from a named custom body-table format.
    pub fn reset_table_cell_custom_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_custom_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Custom-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit currency format for one body-table cell.
    pub fn table_cell_currency_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Currency>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_currency_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit currency format transactionally.
    pub fn set_table_cell_currency_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Currency,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Currency body-table cell.
    pub fn reset_table_cell_currency_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_currency_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages currency-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit percentage format for one body-table cell.
    pub fn table_cell_percentage_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Percentage>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_percentage_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit percentage format transactionally.
    pub fn set_table_cell_percentage_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Percentage,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Percentage body-table cell.
    pub fn reset_table_cell_percentage_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_percentage_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages percentage-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit scientific-notation format for one body-table cell.
    pub fn table_cell_scientific_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Scientific>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_scientific_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit scientific-notation format transactionally.
    pub fn set_table_cell_scientific_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Scientific,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Scientific body-table cell.
    pub fn reset_table_cell_scientific_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_scientific_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages scientific-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit mixed-fraction format for one body-table cell.
    pub fn table_cell_fraction_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Fraction>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_fraction_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit mixed-fraction format transactionally.
    pub fn set_table_cell_fraction_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Fraction,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Fraction body-table cell.
    pub fn reset_table_cell_fraction_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_fraction_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages fraction-format reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit positional numeral-system format for one body-table cell.
    pub fn table_cell_numeral_system_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<NumeralSystem>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_numeral_system_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit positional numeral-system format transactionally.
    pub fn set_table_cell_numeral_system_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: NumeralSystem,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Numeral System body-table cell.
    pub fn reset_table_cell_numeral_system_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_numeral_system_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages numeral-system reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Date & Time format for one body-table cell.
    pub fn table_cell_date_time_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<DateTime>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_date_time_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Date & Time format transactionally.
    pub fn set_table_cell_date_time_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: DateTime,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Date & Time body-table cell.
    pub fn reset_table_cell_date_time_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_date_time_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Date & Time reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Duration format for one body-table cell.
    pub fn table_cell_duration_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Duration>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_duration_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit Duration format transactionally.
    pub fn set_table_cell_duration_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Duration,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Duration body-table cell.
    pub fn reset_table_cell_duration_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_duration_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Duration reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Checkbox format for one body-table cell.
    pub fn table_cell_checkbox_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Checkbox>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_checkbox_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Checkbox format transactionally.
    pub fn set_table_cell_checkbox_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Checkbox,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Checkbox body-table cell.
    pub fn reset_table_cell_checkbox_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_checkbox_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Checkbox reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Star Rating format for one body-table cell.
    pub fn table_cell_star_rating_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<StarRating>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_star_rating_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native five-star rating transactionally.
    pub fn set_table_cell_star_rating_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: StarRating,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Star Rating body-table cell.
    pub fn reset_table_cell_star_rating_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_star_rating_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Star Rating reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Slider format for one body-table cell.
    pub fn table_cell_slider_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Slider>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_slider_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Slider format transactionally.
    pub fn set_table_cell_slider_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Slider,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Slider body-table cell.
    pub fn reset_table_cell_slider_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_slider_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Slider reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Stepper format for one body-table cell.
    pub fn table_cell_stepper_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<Stepper>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_stepper_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Stepper format transactionally.
    pub fn set_table_cell_stepper_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: Stepper,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Stepper body-table cell.
    pub fn reset_table_cell_stepper_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_stepper_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Stepper reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read an explicit Pop-Up Menu format for one body-table cell.
    pub fn table_cell_pop_up_menu_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<PopUpMenu>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_pop_up_menu_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace an explicit native Pop-Up Menu format transactionally.
    pub fn set_table_cell_pop_up_menu_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        format: PopUpMenu,
    ) -> Result<()> {
        self.set_table_cell_data_format(model_object_id, row, column, format.into())
    }

    /// Restore Automatic from an explicit Pop-Up Menu body-table cell.
    pub fn reset_table_cell_pop_up_menu_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_pop_up_menu_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            if verified.table_cell_data_format(model_object_id, row, column)?
                != DataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Pop-Up Menu reset failed package validation".to_owned(),
                ));
            }
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective text layout for one body-table cell.
    pub fn table_cell_layout(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellLayout> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_layout_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace local text-layout overrides for one body-table cell.
    pub fn set_table_cell_layout(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        layout: PagesTableCellLayout,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_layout_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            layout,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_layout(model_object_id, row, column)? != layout {
            return Err(Error::InvalidFormat(
                "Pages table-cell layout failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local text-layout overrides and restore inherited cell values.
    pub fn reset_table_cell_layout(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_layout_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective horizontal text alignment for one body-table cell.
    pub fn table_cell_text_alignment(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Alignment> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_alignment_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a local horizontal text-alignment override.
    pub fn set_table_cell_text_alignment(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        alignment: Alignment,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_alignment_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            alignment,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_alignment(model_object_id, row, column)? != alignment {
            return Err(Error::InvalidFormat(
                "Pages table-cell text alignment failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local horizontal alignment and restore the inherited table style.
    pub fn reset_table_cell_text_alignment(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_alignment_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective paragraph line spacing for one body-table cell.
    pub fn table_cell_paragraph_line_spacing(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<LineSpacing> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_line_spacing_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell paragraph line spacing.
    pub fn set_table_cell_paragraph_line_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        spacing: LineSpacing,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_line_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            spacing,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_line_spacing(model_object_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph line spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local line spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_line_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_line_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing for one body-table cell.
    pub fn table_cell_paragraph_spacing(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Spacing> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_spacing_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell before/after paragraph spacing.
    pub fn set_table_cell_paragraph_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        spacing: Spacing,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            spacing,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_spacing(model_object_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local before/after spacing and restore the inherited table style.
    pub fn reset_table_cell_paragraph_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the canonical list preset applied uniformly to one body-table cell.
    pub fn table_cell_paragraph_list(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellParagraphList> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Promote a plain text cell when necessary and apply one native list preset.
    pub fn set_table_cell_paragraph_list(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        list: PagesTableCellParagraphList,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            list,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list(model_object_id, row, column)? != list {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the canonical None list preset for one body-table cell.
    pub fn reset_table_cell_paragraph_list(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_list_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read all paragraph-scoped list preset boundaries in one body-table cell.
    pub fn table_cell_paragraph_lists(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<PagesTableCellParagraphListPlacement>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_lists_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Promote a plain cell when necessary and replace all list preset boundaries.
    pub fn set_table_cell_paragraph_lists(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        placements: &[PagesTableCellParagraphListPlacement],
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_lists_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            placements,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        let expected = crate::numbers::editor::table_cell_paragraph_lists_in_package(
            &staged,
            model_object_id,
            row,
            column,
        )?;
        if verified.table_cell_paragraph_lists(model_object_id, row, column)? != expected {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph-list placements failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read every effective list-level boundary in one body-table cell.
    pub fn table_cell_paragraph_list_levels(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<PagesTableCellParagraphListLevelPlacement>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_levels_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Set one paragraph's list level without changing later paragraphs.
    pub fn set_table_cell_paragraph_list_level(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        level: PagesTableCellParagraphListLevel,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_level_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            level,
        )?;
        let expected = crate::numbers::editor::table_cell_paragraph_list_levels_in_package(
            &staged,
            model_object_id,
            row,
            column,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_levels(model_object_id, row, column)? != expected {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list levels failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_table_cell_paragraph_list_level(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_list_level_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read whether one body-table paragraph continues or restarts list numbering.
    pub fn table_cell_paragraph_list_numbering(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListNumbering> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_numbering_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Continue or restart numbered-list sequencing at one body-table paragraph.
    pub fn set_table_cell_paragraph_list_numbering(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        numbering: PagesTableCellParagraphListNumbering,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_numbering_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            numbering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_numbering(model_object_id, row, column, paragraph)?
            != numbering
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list numbering failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one numbered body-table paragraph's effective label format.
    pub fn table_cell_paragraph_list_number_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListNumberFormat> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_number_format_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered body-table paragraph's locale-aware label format.
    pub fn set_table_cell_paragraph_list_number_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        format: PagesTableCellParagraphListNumberFormat,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_number_format_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            format,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_number_format(
            model_object_id,
            row,
            column,
            paragraph,
        )? != format
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list-number format failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_table_cell_paragraph_list_number_format(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_number_format_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read whether one numbered body-table paragraph displays hierarchical numbering.
    pub fn table_cell_paragraph_list_number_tiering(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListNumberTiering> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_number_tiering_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Choose flat or hierarchical numbering for one body-table list level.
    pub fn set_table_cell_paragraph_list_number_tiering(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        tiering: PagesTableCellParagraphListNumberTiering,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_number_tiering_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            tiering,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_number_tiering(
            model_object_id,
            row,
            column,
            paragraph,
        )? != tiering
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list-number tiering failed package validation"
                    .to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore flat numbering for one body-table list level.
    pub fn reset_table_cell_paragraph_list_number_tiering(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_number_tiering_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read one numbered body-table paragraph's number-label size.
    pub fn table_cell_paragraph_list_number_scale(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListNumberScale> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_number_scale_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one numbered body-table paragraph's number-label size.
    pub fn set_table_cell_paragraph_list_number_scale(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        scale: PagesTableCellParagraphListNumberScale,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_number_scale_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            scale,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_number_scale(
            model_object_id,
            row,
            column,
            paragraph,
        )? != scale
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list-number scale failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_table_cell_paragraph_list_number_scale(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_number_scale_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read one body-table paragraph's effective text-bullet marker.
    pub fn table_cell_paragraph_list_bullet(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListBullet> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_bullet_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one body-table paragraph's text-bullet marker.
    pub fn set_table_cell_paragraph_list_bullet(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        bullet: &PagesTableCellParagraphListBullet,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_bullet_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            bullet,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_bullet(model_object_id, row, column, paragraph)?
            != *bullet
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph text bullet failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one body-table paragraph.
    pub fn reset_table_cell_paragraph_list_bullet(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_list_bullet_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read one body-table paragraph's effective bullet size and baseline.
    pub fn table_cell_paragraph_list_bullet_geometry(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListBulletGeometry> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_bullet_geometry_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one body-table paragraph's bullet size and baseline.
    pub fn set_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        geometry: PagesTableCellParagraphListBulletGeometry,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_bullet_geometry_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            geometry,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_bullet_geometry(
            model_object_id,
            row,
            column,
            paragraph,
        )? != geometry
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph bullet geometry failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_table_cell_paragraph_list_bullet_geometry(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_bullet_geometry_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read one body-table list paragraph's label and text-gap indentation.
    pub fn table_cell_paragraph_list_indentation(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListIndentation> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_indentation_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one body-table list paragraph's label and text-gap indentation.
    pub fn set_table_cell_paragraph_list_indentation(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        indentation: PagesTableCellParagraphListIndentation,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_indentation_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            indentation,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_indentation(
            model_object_id,
            row,
            column,
            paragraph,
        )? != indentation
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list indentation failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_table_cell_paragraph_list_indentation(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_indentation_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read one body-table list paragraph's effective label color.
    pub fn table_cell_paragraph_list_label_color(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<PagesTableCellParagraphListLabelColor> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_list_label_color_in_package(
            self.package(),
            model_object_id,
            row,
            column,
            paragraph,
        )
    }

    /// Set one body-table list paragraph's bullet or number color.
    pub fn set_table_cell_paragraph_list_label_color(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
        color: PagesTableCellParagraphListLabelColor,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_list_label_color_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            paragraph,
            color,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_list_label_color(
            model_object_id,
            row,
            column,
            paragraph,
        )? != color
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph list-label color failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_table_cell_paragraph_list_label_color(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        paragraph: TextPosition,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed =
            crate::numbers::editor::reset_table_cell_paragraph_list_label_color_in_package(
                &mut staged,
                model_object_id,
                row,
                column,
                paragraph,
            )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right paragraph indents.
    pub fn table_cell_paragraph_indents(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Indents> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_indents_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell paragraph indents.
    pub fn set_table_cell_paragraph_indents(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        indents: Indents,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_indents_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            indents,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_indents(model_object_id, row, column)? != indents {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph indents failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local paragraph indents and restore the inherited table style.
    pub fn reset_table_cell_paragraph_indents(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_indents_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the ordered explicit ruler tab stops for one body-table cell.
    pub fn table_cell_paragraph_tab_stops(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellParagraphTabStops> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_paragraph_tab_stops_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell ruler tab stops.
    pub fn set_table_cell_paragraph_tab_stops(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        stops: PagesTableCellParagraphTabStops,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_paragraph_tab_stops_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            stops.clone(),
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_paragraph_tab_stops(model_object_id, row, column)? != stops {
            return Err(Error::InvalidFormat(
                "Pages table-cell paragraph tab stops failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local ruler tab stops and restore the inherited table style.
    pub fn reset_table_cell_paragraph_tab_stops(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_paragraph_tab_stops_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective background painted behind one body-table cell's text.
    pub fn table_cell_text_background(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextBackground> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_background_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell text-background override.
    pub fn set_table_cell_text_background(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        background: PagesTableCellTextBackground,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_background_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            background,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_background(model_object_id, row, column)? != background {
            return Err(Error::InvalidFormat(
                "Pages table-cell text background failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text background and restore the inherited value.
    pub fn reset_table_cell_text_background(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_background_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of one body-table cell.
    pub fn table_cell_text_baseline_shift(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextBaselineShift> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_baseline_shift_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell custom baseline displacement.
    pub fn set_table_cell_text_baseline_shift(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        shift: PagesTableCellTextBaselineShift,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_baseline_shift_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            shift,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_baseline_shift(model_object_id, row, column)? != shift {
            return Err(Error::InvalidFormat(
                "Pages table-cell baseline shift failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local baseline displacement and restore the inherited value.
    pub fn reset_table_cell_text_baseline_shift(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_baseline_shift_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective capitalization of one body-table cell.
    pub fn table_cell_text_capitalization(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextCapitalization> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_capitalization_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell capitalization.
    pub fn set_table_cell_text_capitalization(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        capitalization: PagesTableCellTextCapitalization,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_capitalization_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            capitalization,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_capitalization(model_object_id, row, column)? != capitalization
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell capitalization failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local capitalization and restore the inherited value.
    pub fn reset_table_cell_text_capitalization(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_capitalization_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of one body-table cell.
    pub fn table_cell_text_character_spacing(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextCharacterSpacing> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_character_spacing_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell character spacing.
    pub fn set_table_cell_text_character_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        spacing: PagesTableCellTextCharacterSpacing,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_character_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            spacing,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_character_spacing(model_object_id, row, column)? != spacing {
            return Err(Error::InvalidFormat(
                "Pages table-cell character spacing failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local character spacing and restore the inherited value.
    pub fn reset_table_cell_text_character_spacing(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_character_spacing_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective foreground text color of one body-table cell.
    pub fn table_cell_text_color(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextColor> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_color_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell foreground text-color override.
    pub fn set_table_cell_text_color(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        color: PagesTableCellTextColor,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_color_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            color,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_color(model_object_id, row, column)? != color {
            return Err(Error::InvalidFormat(
                "Pages table-cell text color failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text-color override and restore the inherited color.
    pub fn reset_table_cell_text_color(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_color_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective whole-cell underline and strikethrough formatting.
    pub fn table_cell_text_decorations(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextDecorations> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_decorations_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell underline and strikethrough formatting.
    pub fn set_table_cell_text_decorations(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        decorations: PagesTableCellTextDecorations,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_decorations_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            decorations,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_decorations(model_object_id, row, column)? != decorations {
            return Err(Error::InvalidFormat(
                "Pages table-cell text decorations failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local decorations and restore the inherited cell formatting.
    pub fn reset_table_cell_text_decorations(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_decorations_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of one body-table cell.
    pub fn table_cell_text_font(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextFont> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_font_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell PostScript font override.
    pub fn set_table_cell_text_font(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        font: PagesTableCellTextFont,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_font_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            font.clone(),
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_font(model_object_id, row, column)? != font {
            return Err(Error::InvalidFormat(
                "Pages table-cell font failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local font override and restore the inherited table font.
    pub fn reset_table_cell_text_font(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_font_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of one body-table cell.
    pub fn table_cell_text_ligatures(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextLigatures> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_ligatures_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace the whole-cell ligature policy.
    pub fn set_table_cell_text_ligatures(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        ligatures: PagesTableCellTextLigatures,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_ligatures_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            ligatures,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_ligatures(model_object_id, row, column)? != ligatures {
            return Err(Error::InvalidFormat(
                "Pages table-cell ligatures failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local ligature policy and restore the inherited value.
    pub fn reset_table_cell_text_ligatures(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_ligatures_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective outline of one body-table cell's text.
    pub fn table_cell_text_outline(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextOutline> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_outline_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell text outline.
    pub fn set_table_cell_text_outline(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        outline: PagesTableCellTextOutline,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_outline_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            outline,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_outline(model_object_id, row, column)? != outline {
            return Err(Error::InvalidFormat(
                "Pages table-cell text outline failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text outline and restore the inherited value.
    pub fn reset_table_cell_text_outline(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_outline_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective normal, superscript, or subscript formatting.
    pub fn table_cell_text_script(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextScript> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_script_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell baseline script formatting.
    pub fn set_table_cell_text_script(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        script: PagesTableCellTextScript,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_script_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            script,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_script(model_object_id, row, column)? != script {
            return Err(Error::InvalidFormat(
                "Pages table-cell baseline script failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local baseline script formatting and restore the inherited value.
    pub fn reset_table_cell_text_script(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_script_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective drop shadow of one body-table cell's text.
    pub fn table_cell_text_shadow(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextShadow> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_shadow_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a whole-cell text drop shadow.
    pub fn set_table_cell_text_shadow(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        shadow: PagesTableCellTextShadow,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_shadow_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            shadow,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_shadow(model_object_id, row, column)? != shadow {
            return Err(Error::InvalidFormat(
                "Pages table-cell text shadow failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a local text shadow and restore the inherited value.
    pub fn reset_table_cell_text_shadow(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_shadow_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read effective whole-cell point size, bold, and italic formatting.
    pub fn table_cell_text_style(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<PagesTableCellTextStyle> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_text_style_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace whole-cell point size, bold, and italic formatting.
    pub fn set_table_cell_text_style(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        style: PagesTableCellTextStyle,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_text_style_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            style,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_cell_text_style(model_object_id, row, column)? != style {
            return Err(Error::InvalidFormat(
                "Pages table-cell text style failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove local whole-cell point size, bold, and italic formatting.
    pub fn reset_table_cell_text_style(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_text_style_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective fill for one body-table cell.
    pub fn table_cell_fill(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<crate::shapes::ShapeFill> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_fill_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace one local body-table cell fill.
    pub fn set_table_cell_fill(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        fill: &crate::shapes::ShapeFill,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_fill_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            fill,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if &verified.table_cell_fill(model_object_id, row, column)? != fill {
            return Err(Error::InvalidFormat(
                "Pages table-cell fill failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct fill override and restore the inherited table style.
    pub fn reset_table_cell_fill(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::reset_table_cell_fill_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        if changed {
            let verified = Self::from_bytes(&staged.to_bytes()?)?;
            verified.require_body_table(model_object_id)?;
            *self = verified;
        }
        Ok(changed)
    }

    /// Read the effective explicit borders for one body-table cell.
    pub fn table_cell_borders(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Borders> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_borders_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace one explicit body-table cell border.
    pub fn set_table_cell_border(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
        stroke: crate::shapes::ShapeStroke,
    ) -> Result<()> {
        self.update_table_cell_border(model_object_id, row, column, side, Some(stroke))
    }

    /// Explicitly clear one body-table cell border.
    pub fn clear_table_cell_border(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
    ) -> Result<()> {
        self.update_table_cell_border(model_object_id, row, column, side, None)
    }

    fn update_table_cell_border(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: BorderSide,
        stroke: Option<crate::shapes::ShapeStroke>,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_border_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            side,
            stroke,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified
            .table_cell_borders(model_object_id, row, column)?
            .get(side)
            != stroke
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell border failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// List native merged-cell rectangles in one reachable body table.
    pub fn table_cell_merges(&self, model_object_id: u64) -> Result<Vec<Region>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_merges_in_package(self.package(), model_object_id)
    }

    /// Merge one non-overlapping body-table rectangle transactionally.
    pub fn merge_table_cells(
        &mut self,
        model_object_id: u64,
        region: Region,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::merge_table_cells_in_package(&mut staged, model_object_id, region)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if !verified
            .table_cell_merges(model_object_id)?
            .contains(&region)
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell merge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove one exact body-table merge, returning whether it existed.
    pub fn unmerge_table_cells(
        &mut self,
        model_object_id: u64,
        region: Region,
    ) -> Result<bool> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let changed = crate::numbers::editor::unmerge_table_cells_in_package(
            &mut staged,
            model_object_id,
            region,
        )?;
        if !changed {
            return Ok(false);
        }
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified
            .table_cell_merges(model_object_id)?
            .contains(&region)
        {
            return Err(Error::InvalidFormat(
                "Pages table-cell unmerge failed package validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(true)
    }

    /// Rename a reachable body table transactionally.
    pub fn rename_table(&mut self, model_object_id: u64, name: &str) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::rename_table_in_package(&mut staged, model_object_id, name)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let renamed = verified.require_body_table(model_object_id)?;
        if renamed.info.name != name {
            return Err(Error::InvalidFormat(
                "Pages table rename failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Resize a reachable body table while preserving existing cells and UIDs.
    ///
    /// Growth creates blank trailing rows or columns. Shrinkage is rejected if
    /// any removed row or column contains stored cell data.
    pub fn resize_table(
        &mut self,
        model_object_id: u64,
        rows: usize,
        columns: usize,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::resize_table_in_package(
            &mut staged,
            model_object_id,
            rows,
            columns,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let resized = verified.require_body_table(model_object_id)?;
        if (resized.info.rows, resized.info.columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Pages table resize failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Insert an independent empty native table at a UTF-16 body position.
    ///
    /// An existing body table supplies native style and storage templates. A
    /// table-less document created by [`PagesEditor::create`] bootstraps that
    /// native scaffold automatically. No cells or formula state are shared,
    /// and the insertion transaction shifts later text attributes and
    /// attachments.
    pub fn add_table(
        &mut self,
        anchor_character_index: usize,
        name: &str,
        rows: usize,
        columns: usize,
    ) -> Result<PagesTableInfo> {
        let template = body_table_graphs(self)?.into_iter().next();
        let body_length = self.body_text()?.encode_utf16().count();
        if anchor_character_index > body_length {
            return Err(Error::ParseError(format!(
                "Pages table anchor {anchor_character_index} exceeds body length {body_length}"
            )));
        }

        let source = self.package();
        let mut staged = source.clone();
        let (new_info_id, new_model_id, new_attachment_id) = if let Some(template) = template {
            let (new_info_id, new_model_id) =
                crate::numbers::editor::create_empty_table_graph_in_package(
                    &mut staged,
                    template.info.drawable_object_id,
                    template.info.model_object_id,
                    self.body_storage_id.get(),
                    name,
                    rows,
                    columns,
                )?;

            let new_attachment_id = clone_body_table_attachment(
                source,
                &mut staged,
                template.attachment_object_id,
                template.info.drawable_object_id,
                new_info_id,
            )?;
            (new_info_id, new_model_id, new_attachment_id)
        } else {
            let graph = crate::pages::creation::bootstrap_first_table_graph(
                &mut staged,
                self.body_storage_id.get(),
                name,
                rows,
                columns,
            )?;
            (
                graph.info_object_id,
                graph.model_object_id,
                graph.attachment_object_id,
            )
        };

        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id.get(),
            anchor_character_index,
            new_attachment_id,
        )?;
        let last_identifier = next_object_identifier(&staged)?
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("Pages package has no object IDs".to_owned()))?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified.require_body_table(new_model_id)?;
        if created.info.drawable_object_id != new_info_id
            || created.info.anchor_character_index != anchor_character_index
            || created.info.name != name
            || (created.info.rows, created.info.columns) != (rows, columns)
        {
            return Err(Error::InvalidFormat(
                "Pages table insertion produced unexpected properties".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Duplicate a populated body table at a UTF-16 body position.
    ///
    /// The clone has independent table storage, object identifiers, table UUID,
    /// and formula-owner state. Its attachment is inserted at
    /// `anchor_character_index`, shifting later body content exactly as a
    /// normal inline table insertion would.
    pub fn duplicate_table(
        &mut self,
        model_object_id: u64,
        anchor_character_index: usize,
    ) -> Result<PagesTableInfo> {
        let body_length = self.body_text()?.encode_utf16().count();
        if anchor_character_index > body_length {
            return Err(Error::ParseError(format!(
                "Pages table anchor {anchor_character_index} exceeds body length {body_length}"
            )));
        }
        let tables = body_table_graphs(self)?;
        let source = tables
            .iter()
            .find(|graph| graph.info.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table model {model_object_id} is not attached to the body"
                ))
            })?;
        let existing_names = tables
            .iter()
            .map(|graph| graph.info.name.as_str())
            .collect::<HashSet<_>>();
        let name =
            crate::numbers::editor::duplicate_table_name(&source.info.name, &existing_names)?;
        let source_info_id = source.info.drawable_object_id;
        let source_model_id = source.info.model_object_id;
        let source_attachment_id = source.attachment_object_id;
        let source_rows = source.info.rows;
        let source_columns = source.info.columns;

        let package = self.package();
        let mut staged = package.clone();
        let cloned = crate::numbers::editor::duplicate_attached_table_graph_in_package(
            package,
            &mut staged,
            source_info_id,
            source_model_id,
            &name,
            INLINE_TABLE_DUPLICATE_OFFSET,
        )?;
        let attachment_id = clone_body_table_attachment(
            package,
            &mut staged,
            source_attachment_id,
            source_info_id,
            cloned.info_object_id,
        )?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id.get(),
            anchor_character_index,
            attachment_id,
        )?;
        let last_identifier = next_object_identifier(&staged)?
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("Pages package has no object IDs".to_owned()))?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified.require_body_table(cloned.model_object_id)?;
        if created.info.drawable_object_id != cloned.info_object_id
            || created.info.anchor_character_index != anchor_character_index
            || created.info.name != name
            || (created.info.rows, created.info.columns) != (source_rows, source_columns)
        {
            return Err(Error::InvalidFormat(
                "Pages table duplication produced unexpected properties".to_owned(),
            ));
        }
        *self = verified;
        Ok(created.info)
    }

    /// Remove a reachable body table and its private native storage graph.
    ///
    /// The body attachment marker, drawable, model, private cell stores,
    /// formula owner family, component references, and UUID registrations are
    /// removed transactionally. Storage shared with another table is retained.
    pub fn remove_table(&mut self, model_object_id: u64) -> Result<PagesTableInfo> {
        let graph = self.require_body_table(model_object_id)?;
        let tables = body_table_graphs(self)?;
        let owned = crate::numbers::editor::table_owned_object_ids_in_package(
            self.package(),
            model_object_id,
        )?;
        let mut shared_owned = HashSet::new();
        for table in tables
            .iter()
            .filter(|table| table.info.model_object_id != model_object_id)
        {
            shared_owned.extend(crate::numbers::editor::table_owned_object_ids_in_package(
                self.package(),
                table.info.model_object_id,
            )?);
        }
        let private_owned = owned
            .into_iter()
            .filter(|identifier| !shared_owned.contains(identifier));
        let mut removed_identifiers = vec![
            graph.attachment_object_id,
            graph.info.drawable_object_id,
            graph.info.model_object_id,
        ];
        removed_identifiers.extend(private_owned);
        let unique = removed_identifiers.iter().copied().collect::<HashSet<_>>();
        if unique.len() != removed_identifiers.len() {
            return Err(Error::InvalidFormat(format!(
                "Pages table model {model_object_id} reuses private graph identifiers"
            )));
        }

        let mut object_components = Vec::with_capacity(removed_identifiers.len());
        for &identifier in &removed_identifiers {
            let archive_name = find_object_archive(self.package(), identifier)?;
            let component = component_identifier_for_entry(self.package(), &archive_name)?;
            object_components.push((identifier, archive_name, component));
        }

        let mut text_editor = IWorkTextEditor::from_package(self.package().clone());
        let anchor_end = graph
            .info
            .anchor_character_index
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Pages table anchor overflow".to_owned()))?;
        text_editor.replace_text(
            self.body_storage_id,
            graph.info.anchor_character_index..anchor_end,
            "",
        )?;
        let mut staged = text_editor.into_package();
        let mut formula_context_ids = graph.formula_context_ids.clone();
        for &identifier in &removed_identifiers {
            if !formula_context_ids.contains(&identifier) {
                formula_context_ids.push(identifier);
            }
        }
        let formula_identifiers = crate::numbers::editor::remove_table_formula_graph_in_package(
            &mut staged,
            &formula_context_ids,
        )?;
        let mut removed_components = HashMap::new();
        for (identifier, archive_name, component) in object_components {
            if let Some(component) = component {
                remove_component_external_references_to_object(&mut staged, component, identifier)?;
                if component_uuid_identifiers(&staged, component)?
                    .is_some_and(|identifiers| identifiers.contains(&identifier))
                {
                    remove_component_object_uuids(&mut staged, component, &[identifier])?;
                }
            }
            if remove_table_object(&mut staged, &archive_name, identifier)?
                && let Some(component) = component
            {
                removed_components.insert(archive_name, component);
            }
        }
        for (archive_name, component) in removed_components {
            if !staged.contains_entry(&archive_name) {
                remove_component_registration(&mut staged, component)?;
            }
        }
        removed_identifiers.extend(formula_identifiers);
        let mut pending = graph.formula_context_ids.clone();
        let mut examined = HashSet::new();
        while let Some(identifier) = pending.pop() {
            if !examined.insert(identifier)
                || removed_identifiers.contains(&identifier)
                || package_references_object(&staged, identifier)?
            {
                continue;
            }
            let Ok(archive_name) = find_object_archive(&staged, identifier) else {
                continue;
            };
            let archive = staged.archive(&archive_name)?;
            let object = archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages table context object {identifier} is missing"
                ))
            })?;
            pending.extend(
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .flat_map(|message| {
                        message.object_references.iter().copied().chain(
                            message
                                .field_infos
                                .iter()
                                .flat_map(|field| field.object_references.iter().copied()),
                        )
                    }),
            );
            let component = component_identifier_for_entry(&staged, &archive_name)?;
            if let Some(component) = component {
                remove_component_external_references_to_object(&mut staged, component, identifier)?;
                if component_uuid_identifiers(&staged, component)?
                    .is_some_and(|identifiers| identifiers.contains(&identifier))
                {
                    remove_component_object_uuids(&mut staged, component, &[identifier])?;
                }
            }
            if remove_table_object(&mut staged, &archive_name, identifier)?
                && let Some(component) = component
            {
                remove_component_registration(&mut staged, component)?;
            }
            removed_identifiers.push(identifier);
        }
        release_package_identifier_suffix(&mut staged, &removed_identifiers)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .tables()?
            .iter()
            .any(|table| table.model_object_id == model_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages table deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(graph.info)
    }
}
