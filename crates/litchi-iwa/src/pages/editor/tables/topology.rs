//! Transactional section-relative row and column CRUD for Pages body tables.

use super::*;
use crate::numbers::editor::TableTopologyMutation;

impl PagesEditor {
    /// Insert one blank row at a section-relative position.
    ///
    /// Cell storage, formulas, dimension overrides, and stable row identities
    /// move together. Unsupported filtered, grouped, pivot, spill, or hidden
    /// topology is rejected without changing the editor.
    pub fn insert_table_row(
        &mut self,
        model_object_id: u64,
        insertion: RowInsertion,
    ) -> Result<()> {
        self.edit_table_topology(model_object_id, TableTopologyMutation::InsertRow(insertion))
    }

    /// Insert one blank column at a section-relative position.
    pub fn insert_table_column(
        &mut self,
        model_object_id: u64,
        insertion: ColumnInsertion,
    ) -> Result<()> {
        self.edit_table_topology(
            model_object_id,
            TableTopologyMutation::InsertColumn(insertion),
        )
    }

    /// Delete one row from a semantic table section and compact following rows.
    ///
    /// The operation fails unchanged when a surviving formula still references
    /// the deleted row or native topology cannot be rewritten safely.
    pub fn remove_table_row(
        &mut self,
        model_object_id: u64,
        deletion: RowDeletion,
    ) -> Result<()> {
        self.edit_table_topology(model_object_id, TableTopologyMutation::RemoveRow(deletion))
    }

    /// Delete one column from a semantic table section and compact following columns.
    pub fn remove_table_column(
        &mut self,
        model_object_id: u64,
        deletion: ColumnDeletion,
    ) -> Result<()> {
        self.edit_table_topology(
            model_object_id,
            TableTopologyMutation::RemoveColumn(deletion),
        )
    }

    fn edit_table_topology(
        &mut self,
        model_object_id: u64,
        edit: TableTopologyMutation,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        let expected = edit.apply(&mut staged, model_object_id)?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let table = verified.require_body_table(model_object_id)?;
        if (table.info.rows, table.info.columns) != expected {
            return Err(Error::InvalidFormat(
                "Pages table topology update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}
