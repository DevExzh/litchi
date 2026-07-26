//! Native table discovery and cell editing for Pages body attachments.

mod appearance;
mod comments;
mod formula;
mod hidden_axes;
mod layout;
mod lock;
mod sort;
mod title;
mod topology;

pub use comments::{
    PagesTableCellComment, PagesTableCellCommentInfo, PagesTableCellCommentReplyInfo,
};
pub use formula::{
    PagesTableFormulaAxisReference, PagesTableFormulaBinaryOperator, PagesTableFormulaCachedValue,
    PagesTableFormulaCellReference, PagesTableFormulaExpression,
};
pub use hidden_axes::{PagesTableAxisIndex, PagesTableHiddenAxes};
pub use layout::{
    PagesTableDimension, PagesTableDimensionSize, PagesTableHeaderCount, PagesTableHeaderSettings,
    PagesTablePoints,
};
pub use sort::{
    PagesTableSortColumnIndex, PagesTableSortDirection, PagesTableSortOrder,
    PagesTableSortRowRange, PagesTableSortRule, PagesTableSortScope,
};
pub use title::PagesTableTitleSettings;

use std::collections::{HashMap, HashSet};

use prost::Message;

use super::*;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::protobuf::tst::{TableInfoArchive, TableModelArchive};
use crate::table_appearance::TableAppearance;
use crate::table_lock::{TableLockState, table_lock_state_from_message};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const OBJECT_REPLACEMENT_CHARACTER: u16 = 0xfffc;
const INLINE_TABLE_DUPLICATE_OFFSET: f32 = 0.0;

/// Strongly typed cell value shared by Pages and Numbers table storage.
pub type PagesCellValue = crate::numbers::CellValue;
/// One mutation in a transactional Pages table-cell batch.
pub type PagesTableCellUpdate = crate::numbers::TableCellUpdate;
/// Section-relative row deletion shared by native iWork tables.
pub type PagesTableRowDeletion = crate::numbers::TableRowDeletion;
/// Section-relative column deletion shared by native iWork tables.
pub type PagesTableColumnDeletion = crate::numbers::TableColumnDeletion;
/// Section-relative row insertion shared by native iWork tables.
pub type PagesTableRowInsertion = crate::numbers::TableRowInsertion;
/// Section-relative column insertion shared by native iWork tables.
pub type PagesTableColumnInsertion = crate::numbers::TableColumnInsertion;
/// A validated native merged-cell rectangle.
pub type PagesTableCellRegion = crate::numbers::editor::IWorkTableCellRegion;
pub use crate::table_cell_border::{
    TableCellBorderSide as PagesTableCellBorderSide, TableCellBorders as PagesTableCellBorders,
};
pub use crate::table_cell_data_format::{
    TableCellCurrencyFormat as PagesTableCellCurrencyFormat,
    TableCellDataFormat as PagesTableCellDataFormat,
    TableCellDateTimeFormat as PagesTableCellDateTimeFormat,
    TableCellFractionFormat as PagesTableCellFractionFormat,
    TableCellNumeralSystemFormat as PagesTableCellNumeralSystemFormat,
    TableCellPercentageFormat as PagesTableCellPercentageFormat,
    TableCellScientificFormat as PagesTableCellScientificFormat,
};
pub use crate::table_cell_layout::{
    TableCellInset as PagesTableCellInset, TableCellInsets as PagesTableCellInsets,
    TableCellLayout as PagesTableCellLayout, TableCellTextWrap as PagesTableCellTextWrap,
    TableCellVerticalAlignment as PagesTableCellVerticalAlignment,
};
pub use crate::table_cell_number_format::{
    TableCellDecimalPlaces as PagesTableCellDecimalPlaces,
    TableCellFixedDecimalPlaces as PagesTableCellFixedDecimalPlaces,
    TableCellNegativeNumberStyle as PagesTableCellNegativeNumberStyle,
    TableCellNumberFormat as PagesTableCellNumberFormat,
    TableCellThousandsSeparator as PagesTableCellThousandsSeparator,
};

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
    /// Materialized non-empty cells indexed by `(row, column)`.
    pub cells: HashMap<(usize, usize), PagesCellValue>,
    /// Comments indexed independently from cell values by `(row, column)`.
    pub comments: HashMap<(usize, usize), PagesTableCellComment>,
    /// Native merged-cell rectangles in formula-store order.
    pub merges: Vec<PagesTableCellRegion>,
}

impl PagesTable {
    /// Borrow a materialized cell value, or return `None` for an empty cell.
    pub fn get_cell(&self, row: usize, column: usize) -> Option<&PagesCellValue> {
        self.cells.get(&(row, column))
    }

    /// Borrow the comment attached to a materialized cell, if any.
    pub fn get_comment(&self, row: usize, column: usize) -> Option<&PagesTableCellComment> {
        self.comments.get(&(row, column))
    }
}

#[derive(Debug, Clone)]
struct PagesTableGraph {
    info: PagesTableInfo,
    drawable_archive: String,
    attachment_object_id: u64,
    formula_context_ids: Vec<u64>,
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
            .resolve_object(&bundle, model_object_id)?
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
        Ok(PagesTable {
            info,
            cells: table.cells,
            comments: table.comments,
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
    ) -> Result<PagesTableCellDataFormat> {
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
        format: PagesTableCellDataFormat,
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
    ) -> Result<Option<PagesTableCellNumberFormat>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_number_format_in_package(
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
        format: PagesTableCellNumberFormat,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_number_format_in_package(
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

    /// Read an explicit currency format for one body-table cell.
    pub fn table_cell_currency_format(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<PagesTableCellCurrencyFormat>> {
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
        format: PagesTableCellCurrencyFormat,
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
                != PagesTableCellDataFormat::Automatic
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
    ) -> Result<Option<PagesTableCellPercentageFormat>> {
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
        format: PagesTableCellPercentageFormat,
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
                != PagesTableCellDataFormat::Automatic
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
    ) -> Result<Option<PagesTableCellScientificFormat>> {
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
        format: PagesTableCellScientificFormat,
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
                != PagesTableCellDataFormat::Automatic
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
    ) -> Result<Option<PagesTableCellFractionFormat>> {
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
        format: PagesTableCellFractionFormat,
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
                != PagesTableCellDataFormat::Automatic
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
    ) -> Result<Option<PagesTableCellNumeralSystemFormat>> {
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
        format: PagesTableCellNumeralSystemFormat,
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
                != PagesTableCellDataFormat::Automatic
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
    ) -> Result<Option<PagesTableCellDateTimeFormat>> {
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
        format: PagesTableCellDateTimeFormat,
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
                != PagesTableCellDataFormat::Automatic
            {
                return Err(Error::InvalidFormat(
                    "Pages Date & Time reset failed package validation".to_owned(),
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
    ) -> Result<PagesTableCellBorders> {
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
        side: PagesTableCellBorderSide,
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
        side: PagesTableCellBorderSide,
    ) -> Result<()> {
        self.update_table_cell_border(model_object_id, row, column, side, None)
    }

    fn update_table_cell_border(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        side: PagesTableCellBorderSide,
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
    pub fn table_cell_merges(&self, model_object_id: u64) -> Result<Vec<PagesTableCellRegion>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_merges_in_package(self.package(), model_object_id)
    }

    /// Merge one non-overlapping body-table rectangle transactionally.
    pub fn merge_table_cells(
        &mut self,
        model_object_id: u64,
        region: PagesTableCellRegion,
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
        region: PagesTableCellRegion,
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
                    self.body_storage_id,
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
                self.body_storage_id,
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
            self.body_storage_id,
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
            self.body_storage_id,
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
            if remove_table_object(&mut staged, &archive_name, identifier)? {
                if let Some(component) = component {
                    removed_components.insert(archive_name, component);
                }
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
            if remove_table_object(&mut staged, &archive_name, identifier)? {
                if let Some(component) = component {
                    remove_component_registration(&mut staged, component)?;
                }
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

    fn require_body_table(&self, model_object_id: u64) -> Result<PagesTableGraph> {
        body_table_graphs(self)?
            .into_iter()
            .find(|graph| graph.info.model_object_id == model_object_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages table model {model_object_id} is not attached to the body"
                ))
            })
    }
}

fn body_table_graphs(editor: &PagesEditor) -> Result<Vec<PagesTableGraph>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    let mut seen_drawables = HashSet::new();
    let mut seen_models = HashSet::new();
    let mut result = Vec::new();

    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        let Some(attachment) = decode_optional_typed_package_object::<DrawableAttachmentArchive>(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )?
        else {
            continue;
        };
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        let archive_name = find_object_archive(editor.package(), drawable.identifier)?;
        let archive = editor.package().archive(&archive_name)?;
        let object = archive.object(drawable.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table drawable {} is missing",
                drawable.identifier
            ))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} repeats its table-info payload",
                drawable.identifier
            )));
        };
        let table_info = TableInfoArchive::decode(message.data.as_slice())?;
        if table_info.super_.parent.map(|parent| parent.identifier) != Some(editor.body_storage_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} is not owned by the body",
                drawable.identifier
            )));
        }
        if body_units.get(entry.character_index as usize) != Some(&OBJECT_REPLACEMENT_CHARACTER) {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} has no object-replacement character",
                drawable.identifier
            )));
        }
        let model_id = table_info.table_model.identifier;
        let model_archive_name = find_object_archive(editor.package(), model_id)?;
        let model_archive = editor.package().archive(&model_archive_name)?;
        let model_object = model_archive.object(model_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages table model {model_id} is missing"))
        })?;
        let models = model_object
            .messages
            .iter()
            .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
            .filter_map(|message| TableModelArchive::decode(message.data.as_slice()).ok())
            .collect::<Vec<_>>();
        let [model] = models.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages table model {model_id} must contain exactly one table-model payload"
            )));
        };
        if !seen_drawables.insert(drawable.identifier) || !seen_models.insert(model_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} or model {model_id} is attached more than once",
                drawable.identifier
            )));
        }
        let mut formula_context_ids = vec![drawable.identifier, model_id];
        for reference in object
            .archive_info
            .message_infos
            .iter()
            .chain(&model_object.archive_info.message_infos)
            .flat_map(|message| {
                message.object_references.iter().copied().chain(
                    message
                        .field_infos
                        .iter()
                        .flat_map(|field| field.object_references.iter().copied()),
                )
            })
        {
            if !formula_context_ids.contains(&reference) {
                formula_context_ids.push(reference);
            }
        }
        result.push(PagesTableGraph {
            drawable_archive: archive_name,
            attachment_object_id: attachment_reference.identifier,
            formula_context_ids,
            info: PagesTableInfo {
                drawable_object_id: drawable.identifier,
                model_object_id: model_id,
                anchor_character_index: entry.character_index as usize,
                name: model.table_name.clone(),
                rows: model.number_of_rows as usize,
                columns: model.number_of_columns as usize,
                appearance: crate::table_appearance::table_appearance(editor.package(), model_id)?,
                lock_state: table_lock_state_from_message(&message.data)?,
            },
        });
    }
    let table_roots = result
        .iter()
        .flat_map(|graph| {
            [
                graph.attachment_object_id,
                graph.info.drawable_object_id,
                graph.info.model_object_id,
            ]
        })
        .collect::<HashSet<_>>();
    for graph in &mut result {
        let mut excluded = table_roots.clone();
        excluded.remove(&graph.attachment_object_id);
        excluded.remove(&graph.info.drawable_object_id);
        excluded.remove(&graph.info.model_object_id);
        excluded.insert(editor.body_storage_id);
        expand_formula_contexts(editor.package(), &mut graph.formula_context_ids, &excluded)?;
    }
    result.sort_by_key(|graph| graph.info.anchor_character_index);
    Ok(result)
}

fn clone_body_table_attachment(
    source: &IWorkPackage,
    staged: &mut IWorkPackage,
    source_attachment_id: u64,
    source_drawable_id: u64,
    new_drawable_id: u64,
) -> Result<u64> {
    let new_attachment_id = next_object_identifier(staged)?;
    let attachment_archive_name = find_object_archive(source, source_attachment_id)?;
    let attachment_archive = source.archive(&attachment_archive_name)?;
    let attachment_object = attachment_archive
        .object(source_attachment_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table attachment {source_attachment_id} is missing"
            ))
        })?;
    let remap = HashMap::from([
        (source_attachment_id, new_attachment_id),
        (source_drawable_id, new_drawable_id),
    ]);
    let cloned_attachment = clone_pages_drawable_graph_object(attachment_object, &remap)?;
    staged.update_archive(&attachment_archive_name, |archive| {
        archive.insert_object(cloned_attachment)
    })?;

    if let Some(component) = component_identifier_for_entry(source, &attachment_archive_name)? {
        if component_uuid_identifiers(source, component)?
            .is_some_and(|identifiers| identifiers.contains(&source_attachment_id))
        {
            add_component_object_uuids(staged, component, &[new_attachment_id])?;
        }
        let new_drawable_archive = find_object_archive(staged, new_drawable_id)?;
        if let Some(target_component) =
            component_identifier_for_entry(staged, &new_drawable_archive)?
            && target_component != component
        {
            add_component_external_reference(staged, component, target_component, new_drawable_id)?;
        }
    }
    Ok(new_attachment_id)
}

fn expand_formula_contexts(
    package: &IWorkPackage,
    contexts: &mut Vec<u64>,
    excluded: &HashSet<u64>,
) -> Result<()> {
    const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
    const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
    const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;

    contexts.retain(|identifier| !excluded.contains(identifier));
    let mut seen = contexts.iter().copied().collect::<HashSet<_>>();
    let mut cursor = 0usize;
    while cursor < contexts.len() {
        let identifier = contexts[cursor];
        cursor += 1;
        let Ok(archive_name) = find_object_archive(package, identifier) else {
            continue;
        };
        let archive = package.archive(&archive_name)?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table context object {identifier} is missing"
            ))
        })?;
        if object.messages.iter().any(|message| {
            matches!(
                message.type_,
                CALCULATION_ENGINE_MESSAGE_TYPE
                    | FORMULA_OWNER_MESSAGE_TYPE
                    | CELL_RECORD_TILE_MESSAGE_TYPE
            )
        }) {
            continue;
        }
        for reference in object
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
            })
        {
            if !excluded.contains(&reference) && seen.insert(reference) {
                contexts.push(reference);
            }
        }
    }
    Ok(())
}

fn remove_table_object(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
) -> Result<bool> {
    let mut archive = package.archive(archive_name)?;
    archive.remove_object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages table object {identifier} is missing"))
    })?;
    if archive.objects.is_empty() {
        package.remove_entry(archive_name).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages table component {archive_name} is missing"))
        })?;
        Ok(true)
    } else {
        package.replace_archive(archive_name, &archive)?;
        Ok(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;
    use crate::table_cell_data_format::{
        TableCellCurrencyCode, TableCellCurrencyStyle, TableCellFractionAccuracy,
        TableCellNumeralSystemBase, TableCellNumeralSystemFixedPlaces,
        TableCellNumeralSystemNegativeStyle, TableCellNumeralSystemPlaces,
    };

    const SOURCE_BUILT_TABLE_INFO_OBJECT_ID: u64 = 9;

    #[test]
    fn source_built_table_roundtrips_cell_border_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Borders", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let stroke = crate::shapes::ShapeStroke::new(
            crate::shapes::RgbaColor::new(0.1, 0.3, 0.9, 1.0, crate::shapes::RgbColorSpace::Srgb)
                .unwrap(),
            crate::shapes::StrokeWidth::new(2.0).unwrap(),
            crate::shapes::StrokePattern::RoundedDash,
        );
        editor
            .set_table_cell_border(model_id, 1, 1, PagesTableCellBorderSide::Right, stroke)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_borders(model_id, 1, 1).unwrap().right,
            Some(stroke)
        );
        reopened
            .clear_table_cell_border(model_id, 1, 1, PagesTableCellBorderSide::Right)
            .unwrap();
        assert_eq!(
            reopened.table_cell_borders(model_id, 1, 1).unwrap().right,
            None
        );
    }

    #[test]
    fn source_built_table_roundtrips_cell_fill_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Fills", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let inherited = editor.table_cell_fill(model_id, 1, 1).unwrap();
        let fill = crate::shapes::ShapeFill::Solid(
            crate::shapes::RgbaColor::new(0.2, 0.75, 0.35, 1.0, crate::shapes::RgbColorSpace::Srgb)
                .unwrap(),
        );
        editor.set_table_cell_fill(model_id, 1, 1, &fill).unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_cell_fill(model_id, 1, 1).unwrap(), fill);
        assert!(reopened.reset_table_cell_fill(model_id, 1, 1).unwrap());
        assert_eq!(reopened.table_cell_fill(model_id, 1, 1).unwrap(), inherited);
    }

    #[test]
    fn source_built_table_roundtrips_cell_layout_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Layout", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let inherited = editor.table_cell_layout(model_id, 1, 1).unwrap();
        let layout = PagesTableCellLayout::default()
            .with_text_wrap(PagesTableCellTextWrap::Wrapped)
            .with_vertical_alignment(PagesTableCellVerticalAlignment::Bottom)
            .with_insets(PagesTableCellInsets::uniform(
                PagesTableCellInset::from_points(5.0).unwrap(),
            ));
        editor
            .set_table_cell_layout(model_id, 1, 1, layout)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_cell_layout(model_id, 1, 1).unwrap(), layout);
        assert!(reopened.reset_table_cell_layout(model_id, 1, 1).unwrap());
        assert_eq!(
            reopened.table_cell_layout(model_id, 1, 1).unwrap(),
            inherited
        );
    }

    #[test]
    fn source_built_table_roundtrips_number_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Formats", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellNumberFormat::new(
            PagesTableCellDecimalPlaces::fixed(2).unwrap(),
            PagesTableCellNegativeNumberStyle::Parentheses,
            PagesTableCellThousandsSeparator::Shown,
        );
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(1_234.5))
            .unwrap();
        editor
            .set_table_cell_number_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_number_format(model_id, 1, 1).unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_number_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_number_format(model_id, 1, 1).unwrap(),
            None
        );
    }

    #[test]
    fn source_built_table_roundtrips_percentage_data_format() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Percentages", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellPercentageFormat::new(
            PagesTableCellDecimalPlaces::fixed(2).unwrap(),
            PagesTableCellNegativeNumberStyle::Parentheses,
            PagesTableCellThousandsSeparator::Shown,
        );
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(-12.345))
            .unwrap();
        editor
            .set_table_cell_percentage_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Percentage(format)
        );
        assert_eq!(
            reopened
                .table_cell_percentage_format(model_id, 1, 1)
                .unwrap(),
            Some(format)
        );
        reopened
            .set_table_cell_data_format(model_id, 1, 1, PagesTableCellDataFormat::Automatic)
            .unwrap();
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_currency_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Currencies", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellCurrencyFormat::new(
            TableCellCurrencyCode::EUR,
            PagesTableCellDecimalPlaces::fixed(2).unwrap(),
            PagesTableCellNegativeNumberStyle::Parentheses,
            PagesTableCellThousandsSeparator::Shown,
            TableCellCurrencyStyle::Accounting,
        );
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_currency_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_currency_format(model_id, 1, 1).unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_currency_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_scientific_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Scientific", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format =
            PagesTableCellScientificFormat::new(PagesTableCellFixedDecimalPlaces::new(5).unwrap());
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_scientific_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_scientific_format(model_id, 1, 1)
                .unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_scientific_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_fraction_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Fractions", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellFractionFormat::new(TableCellFractionAccuracy::Eighths);
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(-12.375))
            .unwrap();
        editor
            .set_table_cell_fraction_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_fraction_format(model_id, 1, 1).unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_fraction_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_numeral_system_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Numeral Systems", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellNumeralSystemFormat::new(
            TableCellNumeralSystemBase::HEXADECIMAL,
            TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
            TableCellNumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_numeral_system_format(model_id, 1, 1, format)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_numeral_system_format(model_id, 1, 1)
                .unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_numeral_system_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_date_time_format_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Dates", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let format = PagesTableCellDateTimeFormat::iso_date_time_24_hour_with_seconds();
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Date(789_332_889.0))
            .unwrap();
        editor
            .set_table_cell_date_time_format(model_id, 1, 1, format.clone())
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_date_time_format(model_id, 1, 1)
                .unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_date_time_format(model_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
            PagesTableCellDataFormat::Automatic
        );
    }

    #[test]
    fn source_built_table_roundtrips_cell_updates() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Report 🙂\n")
            .body_table("Results", 3, 2)
            .build()
            .unwrap();
        let tables = editor.tables().unwrap();
        assert_eq!(tables.len(), 1);
        let info = &tables[0];
        assert_eq!(info.name, "Results");
        assert_eq!((info.rows, info.columns), (3, 2));
        assert_eq!(
            info.anchor_character_index,
            "Report 🙂\n".encode_utf16().count()
        );
        let model_id = info.model_object_id;
        assert!(editor.table(model_id).unwrap().cells.is_empty());

        assert_eq!(
            editor
                .set_table_cells(
                    model_id,
                    [
                        PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Header".to_owned()),),
                        PagesTableCellUpdate::new(1, 1, PagesCellValue::Number(42.5)),
                    ],
                )
                .unwrap(),
            2
        );
        let before_invalid_batch = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cells(
                    model_id,
                    [
                        PagesTableCellUpdate::new(2, 0, PagesCellValue::Boolean(true)),
                        PagesTableCellUpdate::clear(2, 0),
                    ],
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid_batch);
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let table = reopened.table(model_id).unwrap();
        assert_eq!(
            table.get_cell(0, 0),
            Some(&PagesCellValue::Text("Header".to_owned()))
        );
        assert_eq!(table.get_cell(1, 1), Some(&PagesCellValue::Number(42.5)));
        reopened.clear_table_cell(model_id, 0, 0).unwrap();
        assert!(reopened.table(model_id).unwrap().get_cell(0, 0).is_none());
    }

    #[test]
    fn source_built_table_duplication_clones_formula_storage_and_attachment() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Budget\n")
            .body_table("Budget", 3, 2)
            .build()
            .unwrap();
        let source = editor.tables().unwrap().remove(0);
        editor
            .set_table_cells(
                source.model_object_id,
                [
                    PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Category".to_owned())),
                    PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("Travel".to_owned())),
                    PagesTableCellUpdate::new(1, 1, PagesCellValue::Number(125.0)),
                ],
            )
            .unwrap();
        editor
            .set_table_formula(
                source.model_object_id,
                2,
                1,
                PagesTableFormulaExpression::function(
                    "SUM",
                    [
                        PagesTableFormulaExpression::Number(100.0),
                        PagesTableFormulaExpression::Number(25.0),
                    ],
                ),
                PagesTableFormulaCachedValue::Number(125.0),
            )
            .unwrap();

        let insertion_anchor = editor.body_text().unwrap().encode_utf16().count();
        let copied = editor
            .duplicate_table(source.model_object_id, insertion_anchor)
            .unwrap();
        assert_ne!(copied.drawable_object_id, source.drawable_object_id);
        assert_ne!(copied.model_object_id, source.model_object_id);
        assert_eq!(copied.name, "Budget copy");
        assert_eq!(copied.anchor_character_index, insertion_anchor);
        assert_eq!((copied.rows, copied.columns), (source.rows, source.columns));
        assert_eq!(
            editor.table(copied.model_object_id).unwrap().get_cell(1, 0),
            Some(&PagesCellValue::Text("Travel".to_owned()))
        );
        assert_eq!(
            editor
                .table_formula(copied.model_object_id, 2, 1)
                .unwrap()
                .as_deref(),
            Some("=SUM(100,25)")
        );

        editor
            .set_table_cell(
                copied.model_object_id,
                1,
                0,
                PagesCellValue::Text("Lodging".to_owned()),
            )
            .unwrap();
        assert_eq!(
            editor.table(source.model_object_id).unwrap().get_cell(1, 0),
            Some(&PagesCellValue::Text("Travel".to_owned()))
        );

        let second_anchor = editor.body_text().unwrap().encode_utf16().count();
        assert_eq!(
            editor
                .duplicate_table(source.model_object_id, second_anchor)
                .unwrap()
                .name,
            "Budget copy 2"
        );
        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .duplicate_table(source.model_object_id, usize::MAX)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.tables().unwrap().len(), 3);
        let mut reopened = reopened;
        reopened.remove_table(copied.model_object_id).unwrap();
        assert_eq!(reopened.tables().unwrap().len(), 2);
        assert_eq!(
            reopened
                .table(source.model_object_id)
                .unwrap()
                .get_cell(1, 0),
            Some(&PagesCellValue::Text("Travel".to_owned()))
        );
    }

    #[test]
    fn source_built_table_roundtrips_full_table_sort_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Cities", 5, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_header_settings(
                model_id,
                PagesTableHeaderSettings {
                    header_rows: Some(PagesTableHeaderCount::ONE),
                    ..Default::default()
                },
            )
            .unwrap();
        editor
            .set_table_cells(
                model_id,
                [
                    PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Name".to_owned())),
                    PagesTableCellUpdate::new(0, 1, PagesCellValue::Text("Marker".to_owned())),
                    PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("zebra".to_owned())),
                    PagesTableCellUpdate::new(1, 1, PagesCellValue::Text("last".to_owned())),
                    PagesTableCellUpdate::new(2, 0, PagesCellValue::Text("apple".to_owned())),
                    PagesTableCellUpdate::new(2, 1, PagesCellValue::Text("first apple".to_owned())),
                    PagesTableCellUpdate::new(3, 0, PagesCellValue::Text("banana".to_owned())),
                    PagesTableCellUpdate::new(3, 1, PagesCellValue::Text("middle".to_owned())),
                    PagesTableCellUpdate::new(4, 0, PagesCellValue::Text("apple".to_owned())),
                    PagesTableCellUpdate::new(
                        4,
                        1,
                        PagesCellValue::Text("second apple".to_owned()),
                    ),
                ],
            )
            .unwrap();
        editor
            .set_table_cell_comment(model_id, 1, 1, "Zebra comment follows row")
            .unwrap();
        let reply_id = editor
            .add_table_cell_comment_reply(model_id, 1, 1, "Zebra reply follows row")
            .unwrap();
        let comment_id = editor
            .table_cell_comment(model_id, 1, 1)
            .unwrap()
            .unwrap()
            .storage_object_id;
        let hidden = PagesTableHiddenAxes::new([PagesTableAxisIndex::row(2)]).unwrap();
        editor.set_table_hidden_axes(model_id, &hidden).unwrap();
        let order = PagesTableSortOrder::new([PagesTableSortRule::new(
            PagesTableSortColumnIndex::new(0).unwrap(),
            PagesTableSortDirection::Ascending,
        )])
        .unwrap();

        assert_eq!(editor.table_sort_order(model_id).unwrap(), None);
        editor
            .set_table_sort_order(model_id, order.clone())
            .unwrap();
        assert_eq!(
            editor.table_sort_order(model_id).unwrap(),
            Some(order.clone())
        );
        assert!(editor.apply_table_sort_order(model_id).unwrap());
        assert!(!editor.apply_table_sort_order(model_id).unwrap());
        let table = editor.table(model_id).unwrap();
        assert_eq!(
            table.get_cell(0, 0),
            Some(&PagesCellValue::Text("Name".to_owned()))
        );
        assert_eq!(
            table.get_cell(1, 0),
            Some(&PagesCellValue::Text("apple".to_owned()))
        );
        assert_eq!(
            table.get_cell(1, 1),
            Some(&PagesCellValue::Text("first apple".to_owned()))
        );
        assert_eq!(
            table.get_cell(2, 1),
            Some(&PagesCellValue::Text("second apple".to_owned()))
        );
        assert_eq!(
            table.get_cell(3, 0),
            Some(&PagesCellValue::Text("banana".to_owned()))
        );
        assert_eq!(
            table.get_cell(4, 0),
            Some(&PagesCellValue::Text("zebra".to_owned()))
        );
        assert!(editor.table_cell_comment(model_id, 1, 1).unwrap().is_none());
        assert_eq!(
            editor
                .table_cell_comment(model_id, 4, 1)
                .unwrap()
                .unwrap()
                .storage_object_id,
            comment_id
        );
        assert_eq!(
            editor.table_cell_comment_replies(model_id, 4, 1).unwrap()[0].storage_object_id,
            reply_id
        );
        assert_eq!(editor.table_hidden_axes(model_id).unwrap(), hidden);

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_sort_order(model_id).unwrap(),
            Some(order.clone())
        );
        assert_eq!(
            reopened
                .table_cell_comment(model_id, 4, 1)
                .unwrap()
                .unwrap()
                .storage_object_id,
            comment_id
        );
        assert_eq!(
            reopened.table_cell_comment_replies(model_id, 4, 1).unwrap()[0].storage_object_id,
            reply_id
        );
        assert_eq!(reopened.table_hidden_axes(model_id).unwrap(), hidden);
        let unchanged = reopened.to_bytes().unwrap();
        reopened.set_table_sort_order(model_id, order).unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), unchanged);

        let invalid = PagesTableSortOrder::new([PagesTableSortRule::new(
            PagesTableSortColumnIndex::new(2).unwrap(),
            PagesTableSortDirection::Ascending,
        )])
        .unwrap();
        let before_invalid = reopened.to_bytes().unwrap();
        assert!(reopened.set_table_sort_order(model_id, invalid).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), before_invalid);

        let selected_order = PagesTableSortOrder::selected_rows([PagesTableSortRule::new(
            PagesTableSortColumnIndex::new(0).unwrap(),
            PagesTableSortDirection::Descending,
        )])
        .unwrap();
        reopened
            .set_table_sort_order(model_id, selected_order.clone())
            .unwrap();
        assert_eq!(
            reopened.table_sort_order(model_id).unwrap(),
            Some(selected_order)
        );
        let before_wrong_executor = reopened.to_bytes().unwrap();
        assert!(reopened.apply_table_sort_order(model_id).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), before_wrong_executor);
        assert!(
            reopened
                .apply_table_sort_order_to_rows(
                    model_id,
                    PagesTableSortRowRange::new(1, 4).unwrap(),
                )
                .unwrap()
        );
        let table = reopened.table(model_id).unwrap();
        assert_eq!(
            table.get_cell(1, 1),
            Some(&PagesCellValue::Text("first apple".to_owned()))
        );
        assert_eq!(
            table.get_cell(2, 0),
            Some(&PagesCellValue::Text("zebra".to_owned()))
        );
        assert_eq!(
            table.get_cell(3, 0),
            Some(&PagesCellValue::Text("banana".to_owned()))
        );
        assert_eq!(
            table.get_cell(4, 1),
            Some(&PagesCellValue::Text("second apple".to_owned()))
        );
        assert!(
            reopened
                .table_cell_comment(model_id, 4, 1)
                .unwrap()
                .is_none()
        );
        assert_eq!(
            reopened
                .table_cell_comment(model_id, 2, 1)
                .unwrap()
                .unwrap()
                .storage_object_id,
            comment_id
        );
        assert_eq!(
            reopened.table_cell_comment_replies(model_id, 2, 1).unwrap()[0].storage_object_id,
            reply_id
        );
        assert_eq!(reopened.table_hidden_axes(model_id).unwrap(), hidden);

        reopened.clear_table_sort_order(model_id).unwrap();
        assert_eq!(reopened.table_sort_order(model_id).unwrap(), None);
        let unchanged = reopened.to_bytes().unwrap();
        reopened.clear_table_sort_order(model_id).unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), unchanged);
        assert!(reopened.apply_table_sort_order(model_id).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), unchanged);
    }

    #[test]
    fn source_built_table_roundtrips_formula_crud_transactionally() {
        let editor = PagesDocumentBuilder::new()
            .body_table("Formula", 3, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let mut package = editor.into_package();
        let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
        package
            .insert_entry("Index/CalculationEngine-82.iwa", engine)
            .unwrap();
        let mut editor = PagesEditor::from_package(package).unwrap();
        editor
            .set_table_formula(
                model_id,
                1,
                1,
                PagesTableFormulaExpression::function(
                    "SUM",
                    [
                        PagesTableFormulaExpression::Number(1.0),
                        PagesTableFormulaExpression::Number(2.0),
                    ],
                ),
                PagesTableFormulaCachedValue::Number(3.0),
            )
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_formula(model_id, 1, 1).unwrap().as_deref(),
            Some("=SUM(1,2)")
        );
        reopened
            .set_table_formula(
                model_id,
                1,
                1,
                PagesTableFormulaExpression::function(
                    "SUM",
                    [
                        PagesTableFormulaExpression::Number(3.0),
                        PagesTableFormulaExpression::Number(4.0),
                    ],
                ),
                PagesTableFormulaCachedValue::Number(7.0),
            )
            .unwrap();
        assert_eq!(
            reopened.table_formula(model_id, 1, 1).unwrap().as_deref(),
            Some("=SUM(3,4)")
        );

        let before = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_table_formula(
                    model_id,
                    usize::MAX,
                    1,
                    PagesTableFormulaExpression::Number(1.0),
                    PagesTableFormulaCachedValue::Number(1.0),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before);
        assert_eq!(
            reopened.clear_table_formula(model_id, 1, 1).unwrap(),
            "=SUM(3,4)"
        );
        assert_eq!(reopened.table_formula(model_id, 1, 1).unwrap(), None);
        let cleared = reopened.to_bytes().unwrap();
        assert!(reopened.clear_table_formula(model_id, 1, 1).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), cleared);
    }

    #[test]
    fn source_built_table_roundtrips_section_relative_axis_crud_transactionally() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Topology", 4, 4)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let row_size = PagesTableDimensionSize::points(33.0).unwrap();
        let column_size = PagesTableDimensionSize::points(77.0).unwrap();
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Text("shift me".to_owned()))
            .unwrap();
        editor
            .set_table_formula(
                model_id,
                2,
                2,
                PagesTableFormulaExpression::cell(PagesTableFormulaCellReference::relative(1, 1)),
                PagesTableFormulaCachedValue::Number(7.0),
            )
            .unwrap();
        editor.set_table_row_height(model_id, 1, row_size).unwrap();
        editor
            .set_table_column_width(model_id, 1, column_size)
            .unwrap();
        let baseline = editor.to_bytes().unwrap();

        editor
            .insert_table_row(model_id, PagesTableRowInsertion::body(1))
            .unwrap();
        editor
            .insert_table_column(model_id, PagesTableColumnInsertion::body(1))
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let table = reopened.table(model_id).unwrap();
        assert_eq!((table.info.rows, table.info.columns), (5, 5));
        assert_eq!(
            table.get_cell(1, 1),
            Some(&PagesCellValue::Text("shift me".to_owned()))
        );
        assert_eq!(
            table.get_cell(3, 3),
            Some(&PagesCellValue::Formula("=B2".to_owned()))
        );
        assert_eq!(reopened.table_row_height(model_id, 1).unwrap(), row_size);
        assert_eq!(
            reopened.table_column_width(model_id, 1).unwrap(),
            column_size
        );

        editor
            .remove_table_column(model_id, PagesTableColumnDeletion::body(1))
            .unwrap();
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::body(1))
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let before_error = editor.to_bytes().unwrap();
        assert!(
            editor
                .insert_table_row(model_id, PagesTableRowInsertion::body(usize::MAX))
                .is_err()
        );
        assert!(
            editor
                .remove_table_column(model_id, PagesTableColumnDeletion::body(usize::MAX))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_error);
    }

    #[test]
    fn source_built_footer_formula_expands_and_contracts_with_body_rows() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Footer aggregate", 4, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_header_settings(
                model_id,
                PagesTableHeaderSettings {
                    footer_rows: Some(PagesTableHeaderCount::ONE),
                    ..Default::default()
                },
            )
            .unwrap();
        editor
            .set_table_formula(
                model_id,
                3,
                1,
                PagesTableFormulaExpression::function(
                    "SUM",
                    [PagesTableFormulaExpression::range(
                        PagesTableFormulaCellReference::relative(1, 1),
                        PagesTableFormulaCellReference::relative(2, 1),
                    )],
                ),
                PagesTableFormulaCachedValue::Number(3.0),
            )
            .unwrap();
        let mut package = editor.into_package();
        let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
        package
            .insert_entry("Index/CalculationEngine-42-2.iwa", engine)
            .unwrap();
        let mut editor = PagesEditor::from_package(package).unwrap();
        let baseline = editor.to_bytes().unwrap();

        editor
            .insert_table_row(model_id, PagesTableRowInsertion::body(3))
            .unwrap();
        assert_eq!(
            editor.table_formula(model_id, 4, 1).unwrap().as_deref(),
            Some("=SUM(B2:B4)")
        );
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::body(3))
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn source_built_fixed_table_sections_roundtrip_full_axis_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Fixed sections", 4, 4)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_header_settings(
                model_id,
                PagesTableHeaderSettings {
                    header_rows: Some(PagesTableHeaderCount::ONE),
                    header_columns: Some(PagesTableHeaderCount::ONE),
                    footer_rows: Some(PagesTableHeaderCount::ONE),
                    ..Default::default()
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();

        editor
            .insert_table_row(model_id, PagesTableRowInsertion::header(1))
            .unwrap();
        editor
            .insert_table_row(model_id, PagesTableRowInsertion::footer(0))
            .unwrap();
        editor
            .insert_table_column(model_id, PagesTableColumnInsertion::header(1))
            .unwrap();
        let settings = editor.table_header_settings(model_id).unwrap();
        assert_eq!(settings.header_row_count(), 2);
        assert_eq!(settings.footer_row_count(), 2);
        assert_eq!(settings.header_column_count(), 2);
        let table = editor.table(model_id).unwrap();
        assert_eq!((table.info.rows, table.info.columns), (6, 5));

        editor
            .remove_table_column(model_id, PagesTableColumnDeletion::header(1))
            .unwrap();
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::footer(0))
            .unwrap();
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::header(1))
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .set_table_cell(model_id, 0, 0, PagesCellValue::Text("Header".to_owned()))
            .unwrap();
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Text("Body".to_owned()))
            .unwrap();
        editor
            .set_table_cell(model_id, 3, 2, PagesCellValue::Text("Footer".to_owned()))
            .unwrap();
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::header(0))
            .unwrap();
        editor
            .remove_table_row(model_id, PagesTableRowDeletion::footer(0))
            .unwrap();
        editor
            .remove_table_column(model_id, PagesTableColumnDeletion::header(0))
            .unwrap();
        let settings = editor.table_header_settings(model_id).unwrap();
        assert_eq!(settings.header_row_count(), 0);
        assert_eq!(settings.footer_row_count(), 0);
        assert_eq!(settings.header_column_count(), 0);
        let table = editor.table(model_id).unwrap();
        assert_eq!((table.info.rows, table.info.columns), (2, 3));
        assert_eq!(
            table.get_cell(0, 0),
            Some(&PagesCellValue::Text("Body".to_owned()))
        );
        assert!(!table.cells.values().any(|value| matches!(
            value,
            PagesCellValue::Text(text) if text == "Header" || text == "Footer"
        )));
    }

    #[test]
    fn source_built_table_roundtrips_title_settings_transactionally() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Revenue", 2, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let visible = PagesTableTitleSettings {
            visible: Some(true),
            outlined: Some(true),
        };
        let initially_hidden = PagesTableTitleSettings {
            visible: Some(false),
            outlined: None,
        };
        assert_eq!(
            editor.table_title_settings(model_id).unwrap(),
            initially_hidden
        );
        editor.set_table_title_settings(model_id, visible).unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_title_settings(model_id).unwrap(), visible);
        let unchanged = reopened.to_bytes().unwrap();
        reopened
            .set_table_title_settings(model_id, visible)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), unchanged);

        let explicit_hidden = PagesTableTitleSettings {
            visible: Some(false),
            outlined: Some(false),
        };
        reopened
            .set_table_title_settings(model_id, explicit_hidden)
            .unwrap();
        assert_eq!(
            reopened.table_title_settings(model_id).unwrap(),
            explicit_hidden
        );
        reopened
            .set_table_title_settings(model_id, PagesTableTitleSettings::default())
            .unwrap();
        assert_eq!(
            reopened.table_title_settings(model_id).unwrap(),
            PagesTableTitleSettings::default()
        );

        let before_error = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_table_title_settings(u64::MAX, visible)
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_error);
    }

    #[test]
    fn inserts_independent_table_and_shifts_existing_anchor() {
        let body = "Alpha 🙂\nBeta\n";
        let mut editor = PagesDocumentBuilder::new()
            .body_text(body)
            .body_table("Source", 3, 2)
            .build()
            .unwrap();
        let source = editor.tables().unwrap().remove(0);
        editor
            .set_table_cell(
                source.model_object_id,
                1,
                1,
                PagesCellValue::Text("source only".to_owned()),
            )
            .unwrap();
        let anchor = "Alpha 🙂\n".encode_utf16().count();
        let mut package = editor.into_package();
        let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
        package
            .insert_entry("Index/CalculationEngine-83.iwa", engine)
            .unwrap();
        let mut editor = PagesEditor::from_package(package).unwrap();

        let inserted = editor.add_table(anchor, "Inserted", 4, 3).unwrap();
        assert_eq!(inserted.anchor_character_index, anchor);
        assert_eq!((inserted.rows, inserted.columns), (4, 3));
        assert!(
            editor
                .table(inserted.model_object_id)
                .unwrap()
                .cells
                .is_empty()
        );
        let tables = editor.tables().unwrap();
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0], inserted);
        assert_eq!(
            tables[1].anchor_character_index,
            body.encode_utf16().count() + 1
        );
        assert_eq!(tables[1].model_object_id, source.model_object_id);

        editor
            .set_table_cell(
                inserted.model_object_id,
                0,
                0,
                PagesCellValue::Text("inserted only".to_owned()),
            )
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table(source.model_object_id)
                .unwrap()
                .get_cell(1, 1),
            Some(&PagesCellValue::Text("source only".to_owned()))
        );
        assert_eq!(
            reopened
                .table(inserted.model_object_id)
                .unwrap()
                .get_cell(0, 0),
            Some(&PagesCellValue::Text("inserted only".to_owned()))
        );
        reopened.remove_table(inserted.model_object_id).unwrap();
        let retained = reopened.tables().unwrap();
        assert_eq!(retained.len(), 1);
        assert_eq!(retained[0].model_object_id, source.model_object_id);
        assert_eq!(
            retained[0].anchor_character_index,
            body.encode_utf16().count()
        );
        assert_eq!(
            reopened
                .table(source.model_object_id)
                .unwrap()
                .get_cell(1, 1),
            Some(&PagesCellValue::Text("source only".to_owned()))
        );
    }

    #[test]
    fn inserts_and_removes_first_table_without_a_template() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Before 🙂 after")
            .build()
            .unwrap();
        let anchor = "Before 🙂".encode_utf16().count();
        let created = editor.add_table(anchor, "First runtime", 2, 3).unwrap();
        assert_eq!(created.anchor_character_index, anchor);
        assert_eq!((created.rows, created.columns), (2, 3));
        editor
            .set_table_cell(
                created.model_object_id,
                1,
                2,
                PagesCellValue::Text("bootstrapped".to_owned()),
            )
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table(created.model_object_id)
                .unwrap()
                .get_cell(1, 2),
            Some(&PagesCellValue::Text("bootstrapped".to_owned()))
        );
        reopened.remove_table(created.model_object_id).unwrap();
        assert!(reopened.tables().unwrap().is_empty());
        assert_eq!(reopened.body_text().unwrap(), "Before 🙂 after");
    }

    #[test]
    fn first_table_bootstrap_rejects_reserved_id_collision_transactionally() {
        let mut package = PagesDocumentBuilder::new()
            .body_text("Collision")
            .build_package()
            .unwrap();
        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.insert_object(ArchiveObject::new(
                    SOURCE_BUILT_TABLE_INFO_OBJECT_ID,
                    vec![RawMessage {
                        type_: u32::MAX,
                        data: Vec::new(),
                    }],
                )?)
            })
            .unwrap();
        let mut editor = PagesEditor::from_package(package).unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(editor.add_table(0, "Collision", 2, 2).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn invalid_insertion_anchor_is_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Short")
            .body_table("Template", 2, 2)
            .build()
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(editor.add_table(usize::MAX, "Invalid", 2, 2).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn out_of_bounds_cell_update_is_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Bounded", 2, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell(model_id, 2, 0, PagesCellValue::Boolean(true))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn source_built_table_roundtrips_rename_and_resize() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Original", 3, 2)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_cell(model_id, 1, 1, PagesCellValue::Text("kept".to_owned()))
            .unwrap();

        editor.rename_table(model_id, "Renamed").unwrap();
        editor.resize_table(model_id, 5, 4).unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let info = reopened.tables().unwrap().remove(0);
        assert_eq!(info.name, "Renamed");
        assert_eq!((info.rows, info.columns), (5, 4));
        assert_eq!(
            reopened.table(model_id).unwrap().get_cell(1, 1),
            Some(&PagesCellValue::Text("kept".to_owned()))
        );

        reopened.resize_table(model_id, 2, 2).unwrap();
        let info = reopened.tables().unwrap().remove(0);
        assert_eq!((info.rows, info.columns), (2, 2));
    }

    #[test]
    fn source_built_table_roundtrips_layout_crud_transactionally() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Layout", 4, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        let settings = PagesTableHeaderSettings {
            header_rows: Some(PagesTableHeaderCount::TWO),
            header_columns: Some(PagesTableHeaderCount::ONE),
            footer_rows: Some(PagesTableHeaderCount::ONE),
            ..Default::default()
        };

        editor
            .set_table_header_settings(model_id, settings)
            .unwrap();
        editor
            .set_table_column_width(model_id, 0, PagesTableDimensionSize::points(150.0).unwrap())
            .unwrap();
        editor
            .set_table_row_height(model_id, 2, PagesTableDimensionSize::points(42.0).unwrap())
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_header_settings(model_id).unwrap(), settings);
        assert_eq!(
            reopened.table_column_width(model_id, 0).unwrap(),
            PagesTableDimensionSize::points(150.0).unwrap()
        );
        assert_eq!(
            reopened.table_row_height(model_id, 2).unwrap(),
            PagesTableDimensionSize::points(42.0).unwrap()
        );
        assert_eq!(
            reopened.table_column_width(model_id, 1).unwrap(),
            PagesTableDimensionSize::Default
        );

        reopened
            .set_table_column_width(model_id, 0, PagesTableDimensionSize::Default)
            .unwrap();
        reopened
            .set_table_row_height(model_id, 2, PagesTableDimensionSize::Default)
            .unwrap();
        assert_eq!(
            reopened.table_column_width(model_id, 0).unwrap(),
            PagesTableDimensionSize::Default
        );
        assert_eq!(
            reopened.table_row_height(model_id, 2).unwrap(),
            PagesTableDimensionSize::Default
        );

        let before = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_table_row_height(
                    model_id,
                    usize::MAX,
                    PagesTableDimensionSize::points(20.0).unwrap(),
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before);
        assert!(
            reopened
                .set_table_header_settings(
                    model_id,
                    PagesTableHeaderSettings {
                        header_rows: Some(PagesTableHeaderCount::FOUR),
                        footer_rows: Some(PagesTableHeaderCount::ONE),
                        ..Default::default()
                    },
                )
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before);
    }

    #[test]
    fn table_rename_and_occupied_shrink_are_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Protected", 3, 3)
            .build()
            .unwrap();
        let model_id = editor.tables().unwrap()[0].model_object_id;
        editor
            .set_table_cell(model_id, 2, 2, PagesCellValue::Number(7.0))
            .unwrap();

        let before = editor.to_bytes().unwrap();
        assert!(editor.rename_table(model_id, "").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(editor.resize_table(model_id, 2, 2).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn source_built_table_deletion_removes_private_graph_and_anchor() {
        let body = "Before 🙂 after";
        let mut editor = PagesDocumentBuilder::new()
            .body_text(body)
            .body_table("Disposable", 3, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap().remove(0);
        let owned = crate::numbers::editor::table_owned_object_ids_in_package(
            editor.package(),
            table.model_object_id,
        )
        .unwrap();
        editor
            .set_table_cell(
                table.model_object_id,
                1,
                1,
                PagesCellValue::Text("removed".to_owned()),
            )
            .unwrap();

        let removed = editor.remove_table(table.model_object_id).unwrap();
        assert_eq!(removed, table);
        assert!(editor.tables().unwrap().is_empty());
        assert_eq!(editor.body_text().unwrap(), body);
        let mut removed_ids = owned;
        removed_ids.extend([table.drawable_object_id, table.model_object_id]);
        for identifier in removed_ids {
            assert!(find_object_archive(editor.package(), identifier).is_err());
        }
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(reopened.tables().unwrap().is_empty());
        assert_eq!(reopened.body_text().unwrap(), body);
    }

    #[test]
    fn missing_table_deletion_is_transactional() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Retained", 2, 2)
            .build()
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(editor.remove_table(u64::MAX).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
