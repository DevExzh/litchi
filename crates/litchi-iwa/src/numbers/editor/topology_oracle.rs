//! Unverified row and column edits for differential tests.
//!
//! The verified editor methods reopen the staged package and read every
//! table's appearance before they publish. Several native corpus files fail
//! that appearance read for reasons unrelated to the axis edit, which would
//! hide the formula-coordinate shift from the record 0318 differential test
//! against the focused `litchi-numbers` transaction. These entry points stage
//! the same edit without that reopen; they are not a supported editing API.

use super::package::TableTopologyMutation;
use super::*;

/// One row or column edit, for [`Draft::__edit_table_topology_unverified`].
#[doc(hidden)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnverifiedTopologyEdit {
    /// Insert a row.
    InsertRow(RowInsertion),
    /// Insert a column.
    InsertColumn(ColumnInsertion),
    /// Remove a row.
    RemoveRow(RowDeletion),
    /// Remove a column.
    RemoveColumn(ColumnDeletion),
}

impl Draft {
    /// Stage one row or column edit without the post-edit appearance
    /// verification of the public methods. Differential tests only.
    ///
    /// # Errors
    ///
    /// The edit's own refusals; nothing is changed on error.
    #[doc(hidden)]
    pub fn __edit_table_topology_unverified(
        &mut self,
        selector: litchi_numbers::TableSelector<'_>,
        edit: UnverifiedTopologyEdit,
    ) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        let mutation = match edit {
            UnverifiedTopologyEdit::InsertRow(insertion) => {
                TableTopologyMutation::InsertRow(insertion)
            },
            UnverifiedTopologyEdit::InsertColumn(insertion) => {
                TableTopologyMutation::InsertColumn(insertion)
            },
            UnverifiedTopologyEdit::RemoveRow(deletion) => {
                TableTopologyMutation::RemoveRow(deletion)
            },
            UnverifiedTopologyEdit::RemoveColumn(deletion) => {
                TableTopologyMutation::RemoveColumn(deletion)
            },
        };
        mutation.apply(&mut staged, table_id)?;
        self.package = staged;
        Ok(())
    }
}
