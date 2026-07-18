//! Transactional physical row and column CRUD for Keynote slide tables.

use super::*;
use crate::numbers::editor::TableTopologyMutation;

impl KeynoteEditor {
    /// Insert one blank row before `row`; `row == table.rows` appends.
    ///
    /// Cell storage, formulas, dimension overrides, stable row identities, and
    /// the slide drawable bounds move together transactionally.
    pub fn insert_slide_table_row(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
    ) -> Result<()> {
        self.edit_slide_table_topology(
            slide_index,
            model_object_id,
            TableTopologyMutation::InsertRow(row),
        )
    }

    /// Insert one blank column before `column`; `column == table.columns` appends.
    pub fn insert_slide_table_column(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        column: usize,
    ) -> Result<()> {
        self.edit_slide_table_topology(
            slide_index,
            model_object_id,
            TableTopologyMutation::InsertColumn(column),
        )
    }

    /// Delete one physical row and compact all following rows.
    ///
    /// The operation fails unchanged when a surviving formula still references
    /// the deleted row or native topology cannot be rewritten safely.
    pub fn remove_slide_table_row(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
    ) -> Result<()> {
        self.edit_slide_table_topology(
            slide_index,
            model_object_id,
            TableTopologyMutation::RemoveRow(row),
        )
    }

    /// Delete one physical column and compact all following columns.
    pub fn remove_slide_table_column(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        column: usize,
    ) -> Result<()> {
        self.edit_slide_table_topology(
            slide_index,
            model_object_id,
            TableTopologyMutation::RemoveColumn(column),
        )
    }

    fn edit_slide_table_topology(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        edit: TableTopologyMutation,
    ) -> Result<()> {
        let source = require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        let expected = edit.apply(&mut staged, model_object_id)?;
        let (width, height) =
            crate::numbers::editor::table_size_points_in_package(&staged, model_object_id)?;
        let geometry = DrawableGeometry {
            size: Some(DrawableSize { width, height }),
            ..source.geometry
        };
        set_table_geometry_in_package(&mut staged, source.drawable_object_id, geometry)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let table = require_table_model(&verified, slide_index, model_object_id)?;
        if (table.rows, table.columns) != expected || table.geometry != geometry {
            return Err(Error::InvalidFormat(
                "Keynote table topology update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}
