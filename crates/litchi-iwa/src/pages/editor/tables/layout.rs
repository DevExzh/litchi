//! Typed header, footer, row-height, and column-width editing for Pages tables.

use super::*;

use litchi_numbers::table::headers::Settings as HeaderSettings;
/// One row or column addressed by zero-based index.
pub type PagesTableDimension = crate::numbers::NumbersTableDimension;
/// A validated positive point measurement for a table axis.
pub type PagesTablePoints = crate::numbers::NumbersTablePoints;
/// Either a table style's default axis size or an explicit point override.
pub type PagesTableDimensionSize = crate::numbers::NumbersTableDimensionSize;

impl PagesEditor {
    /// Read a body table's lossless header and footer configuration.
    pub fn table_header_settings(&self, model_object_id: u64) -> Result<HeaderSettings> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_header_settings_in_package(self.package(), model_object_id)
    }

    /// Replace a body table's header and footer configuration transactionally.
    pub fn set_table_header_settings(
        &mut self,
        model_object_id: u64,
        settings: HeaderSettings,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_header_settings(model_object_id)? == settings {
            return Ok(());
        }

        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_header_settings_in_package(
            &mut staged,
            model_object_id,
            settings,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_header_settings(model_object_id)? != settings {
            return Err(Error::InvalidFormat(
                "Pages table header settings failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one row-height or column-width override.
    pub fn table_dimension_size(
        &self,
        model_object_id: u64,
        dimension: PagesTableDimension,
    ) -> Result<PagesTableDimensionSize> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_dimension_size_in_package(
            self.package(),
            model_object_id,
            dimension,
        )
    }

    /// Set or clear one row-height or column-width override transactionally.
    pub fn set_table_dimension_size(
        &mut self,
        model_object_id: u64,
        dimension: PagesTableDimension,
        size: PagesTableDimensionSize,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_dimension_size(model_object_id, dimension)? == size {
            return Ok(());
        }

        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_dimension_size_in_package(
            &mut staged,
            model_object_id,
            dimension,
            size,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_dimension_size(model_object_id, dimension)? != size {
            return Err(Error::InvalidFormat(
                "Pages table dimension update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read one row-height override.
    pub fn table_row_height(
        &self,
        model_object_id: u64,
        row: usize,
    ) -> Result<PagesTableDimensionSize> {
        self.table_dimension_size(model_object_id, PagesTableDimension::Row(row))
    }

    /// Set or clear one row-height override.
    pub fn set_table_row_height(
        &mut self,
        model_object_id: u64,
        row: usize,
        size: PagesTableDimensionSize,
    ) -> Result<()> {
        self.set_table_dimension_size(model_object_id, PagesTableDimension::Row(row), size)
    }

    /// Read one column-width override.
    pub fn table_column_width(
        &self,
        model_object_id: u64,
        column: usize,
    ) -> Result<PagesTableDimensionSize> {
        self.table_dimension_size(model_object_id, PagesTableDimension::Column(column))
    }

    /// Set or clear one column-width override.
    pub fn set_table_column_width(
        &mut self,
        model_object_id: u64,
        column: usize,
        size: PagesTableDimensionSize,
    ) -> Result<()> {
        self.set_table_dimension_size(model_object_id, PagesTableDimension::Column(column), size)
    }
}
