//! Immutable worksheet query-table edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_query_tables`] starts an [`Edit`] of one worksheet; every operation
//! runs on a private candidate graph, so an error stages nothing.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_query_tables`] replays a patch onto an equal package. The
//! source package never changes. A changed edit of a signed workbook is
//! refused with [`Error::Signed`]; an exact no-op keeps every byte.

use super::model::{Conformance, Part, Table};
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("query-table");

pub use crate::opc::LimitKind;

/// A staged worksheet query-table edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    worksheet: litchi_opc::PackURI,
}

impl Package {
    /// Start a worksheet query-table edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, a `NotFound` refusal for a
    /// selector that matches no sheet, and [`Error::NotWorksheet`] for a
    /// chart or dialog sheet.
    pub fn edit_query_tables<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Edit> {
        let staged = Staged::new(self, "edit_query_tables")?;
        let worksheet = self.worksheet_part(sheet, "edit_query_tables")?;
        Ok(Edit { staged, worksheet })
    }

    /// Apply a worksheet query-table patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_query_tables(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_query_tables")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The worksheet's direct and table-owned query tables, ordered by owner
    /// part name and then relationship ID.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the worksheet's query-table graph.
    pub fn query_tables(&self) -> Result<Vec<Part>> {
        super::package::load_worksheet_query_tables(self.staged.opc(), &self.worksheet)
    }

    /// The staged query table with an unambiguous relationship ID.
    /// IDs are local to owner parts; duplicate IDs across owners are refused.
    ///
    /// # Errors
    ///
    /// As [`Self::query_tables`].
    pub fn query_table(&self, relationship_id: &str) -> Result<Option<Part>> {
        super::package::find_worksheet_query_table(
            self.staged.opc(),
            &self.worksheet,
            relationship_id,
        )
    }

    /// Add an inert query-table part; the referenced connection must
    /// already exist.
    ///
    /// # Errors
    ///
    /// Returns a missing-connection or validation error, and
    /// [`Error::Signed`] for a signed package.
    pub fn add(&mut self, query_table: Table, conformance: Conformance) -> Result<Part> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::add_worksheet_query_table(
                candidate,
                &worksheet,
                query_table,
                conformance,
            )
        })
    }

    /// Replace the definition of one existing query-table part.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent relationship, a missing connection or
    /// an invalid definition, and [`Error::Signed`] for a signed package.
    pub fn replace(
        &mut self,
        relationship_id: &str,
        query_table: Table,
        conformance: Conformance,
    ) -> Result<()> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::replace_worksheet_query_table(
                candidate,
                &worksheet,
                relationship_id,
                query_table,
                conformance,
            )
        })
    }

    /// Replace a query-table definition selected by its full package part name.
    /// This disambiguates equal relationship IDs belonging to different tables.
    /// The selected part must belong to this worksheet, directly or through one
    /// of its table definitions; table and worksheet XML remain unchanged.
    pub fn replace_part(
        &mut self,
        part_name: &str,
        query_table: Table,
        conformance: Conformance,
    ) -> Result<()> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::replace_by_part_name(
                candidate,
                &worksheet,
                part_name,
                query_table,
                conformance,
            )
        })
    }

    /// Remove one query-table relationship and its unreferenced part.
    /// Returns whether one was removed.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the graph and [`Error::Signed`] for
    /// a change to a signed package.
    pub fn remove(&mut self, relationship_id: &str) -> Result<bool> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::remove_worksheet_query_table(candidate, &worksheet, relationship_id)
        })
    }

    /// Reassign deterministic query-table relationship IDs in the given
    /// order, which must be a permutation of the current IDs.
    ///
    /// # Errors
    ///
    /// Returns an error for an order that is not a permutation and
    /// [`Error::Signed`] for a change to a signed package.
    pub fn reorder(&mut self, ordered_relationship_ids: &[String]) -> Result<Vec<Part>> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::reorder_worksheet_query_tables(
                candidate,
                &worksheet,
                ordered_relationship_ids,
            )
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed signed package and the
    /// workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
