//! Existing-cell edits of a [`SourceBackedSpreadsheet`] (record 0358a):
//! scalar values, formulas and direct hyperlinks of cells that already
//! exist, published by replacing only the changed rows of `content.xml`.
//!
//! [`SourceBackedSpreadsheet::edit_cells`] starts an [`Edit`];
//! [`Edit::commit`] returns a [`Commit`] holding the new spreadsheet, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`SourceBackedSpreadsheet::apply_cells`] replays a patch. The source
//! spreadsheet never changes, and nothing is written until the committed
//! spreadsheet's [`write_to`](SourceBackedSpreadsheet::write_to) or
//! [`save`](SourceBackedSpreadsheet::save).
//!
//! ```no_run
//! use litchi_ods::{Cell, CellValue, SourceBackedSpreadsheet};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let spreadsheet = SourceBackedSpreadsheet::open("book.ods")?;
//! let mut edit = spreadsheet.edit_cells()?;
//! edit.set_cell("Data", 0, 0, Cell::new(CellValue::Text("one".into()), "one"))?;
//! let commit = edit.commit()?;
//! assert_eq!(commit.diagnostics().changed_cells(), 1);
//! commit.package().save("book-updated.ods")?;
//! let restored = commit.package().apply_cells(&commit.patch().inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```

use crate::SourceBackedSpreadsheet;
use crate::facade::SourceCellEdit;
use crate::worksheet::{Cell, CellChange, Selector};

pub use crate::facade::{FormulaChange, HyperlinkChange, HyperlinkOperation};
pub use litchi_odf_common::{Error, LimitKind};

super::content_patch!("an existing-cell edit");

/// The result of an existing-cell edit's `commit`: the new source-backed
/// spreadsheet, the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<SourceBackedSpreadsheet, Patch, Diagnostics>;

/// What an existing-cell commit changed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    changed_cells: usize,
}

impl Diagnostics {
    /// Whether the commit changed `content.xml`.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Number of effective unique cells the commit changed.
    #[must_use]
    pub const fn changed_cells(&self) -> usize {
        self.changed_cells
    }
}

/// A staged existing-cell edit of one [`SourceBackedSpreadsheet`]. Dropping
/// it discards it.
pub struct Edit<'a> {
    source: &'a SourceBackedSpreadsheet,
    inner: SourceCellEdit<'a>,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("changed_cells", &self.inner.changed_cells())
            .finish_non_exhaustive()
    }
}

impl SourceBackedSpreadsheet {
    /// Start an existing-cell edit.
    ///
    /// # Errors
    ///
    /// Returns a source change, or the parse failure of a committed
    /// spreadsheet's `content.xml`.
    pub fn edit_cells(&self) -> Result<Edit<'_>, Error> {
        let base = self.effective()?;
        Ok(Edit {
            source: self,
            inner: base.cell_snapshot()?.edit()?,
        })
    }

    /// Apply an existing-cell patch to a new spreadsheet; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this spreadsheet is over another
    /// source or does not hold the patch's before `content.xml`.
    pub fn apply_cells(&self, patch: &Patch) -> Result<Commit, Error> {
        let package = patch
            .change
            .publish(self, "ods.source_backed.cells.apply")?;
        let changed = !patch.is_noop();
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics {
                changed,
                changed_cells: 0,
            },
        ))
    }
}

impl Edit<'_> {
    /// Number of effective unique cell changes staged so far.
    #[must_use]
    pub fn changed_cells(&self) -> usize {
        self.inner.changed_cells()
    }

    /// Whether the edit currently stages no effective change.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.is_no_op()
    }

    /// Replace one existing ordinary scalar cell. `Ok(None)` means the
    /// worksheet is missing; `Ok(Some(false))` that the cell already holds
    /// the value.
    ///
    /// # Errors
    ///
    /// Refuses a missing or unsupported cell; the edit is unchanged.
    pub fn set_cell<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<Option<bool>, Error> {
        Ok(self.inner.set_cell(selector, row, column, cell)?)
    }

    /// Replace one existing formula cell while retaining its cached value,
    /// displayed text, style, and other cell metadata.
    ///
    /// # Errors
    ///
    /// As [`Self::set_cell`].
    pub fn set_formula<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<Option<bool>, Error> {
        Ok(self.inner.set_formula(selector, row, column, formula)?)
    }

    /// Atomically apply ordered direct-hyperlink operations to existing
    /// cells, returning the number of effective cell changes.
    ///
    /// Missing cells, repeated physical rows, merged covered cells, duplicate
    /// coordinates, and unsupported row shapes are refused. New or changed
    /// links are authoring-validated; unchanged parsed links are only
    /// required to pass storage validation.
    ///
    /// # Errors
    ///
    /// As [`Self::set_cell`].
    pub fn set_hyperlinks<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        changes: Vec<HyperlinkChange>,
    ) -> Result<Option<usize>, Error> {
        Ok(self.inner.set_hyperlinks(selector, changes)?)
    }

    /// Atomically replace a bounded batch of existing formula cells.
    ///
    /// Missing cells, scalar cells, repeated physical rows, merged cells,
    /// unknown values, and duplicate coordinates are refused. Formula
    /// changes never evaluate or replace cached values.
    ///
    /// # Errors
    ///
    /// As [`Self::set_cell`].
    pub fn set_formulas<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        changes: Vec<FormulaChange>,
    ) -> Result<Option<usize>, Error> {
        Ok(self.inner.set_formulas(selector, changes)?)
    }

    /// Atomically replace one bounded batch of existing ordinary scalar
    /// cells.
    ///
    /// Missing cells, repeated physical rows, formulas, merged cells, unknown
    /// values, and style retargeting are refused. Repeated cell runs may be
    /// split because the containing physical row remains the same owner.
    ///
    /// # Errors
    ///
    /// As [`Self::set_cell`].
    pub fn set_cells<'s>(
        &mut self,
        selector: impl Into<Selector<'s>>,
        changes: Vec<CellChange>,
    ) -> Result<Option<usize>, Error> {
        Ok(self.inner.set_cells(selector, changes)?)
    }

    /// Discard every staged change.
    ///
    /// # Errors
    ///
    /// Returns a source change.
    pub fn reset(&mut self) -> Result<(), Error> {
        self.inner = self.source.effective()?.cell_snapshot()?.edit()?;
        Ok(())
    }

    /// Validate the row-local replacement and commit it to a new
    /// spreadsheet. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns the readback or eligibility failure of the candidate, or a
    /// source change.
    pub fn commit(self) -> Result<Commit, Error> {
        let base = self.source.effective()?;
        let commit = self.inner.commit()?;
        let changed_cells = commit.changed_cells();
        let change = super::Change::new(base, std::sync::Arc::clone(commit.content_xml()));
        let changed = !change.is_noop();
        let package = if changed {
            base.with_content(std::sync::Arc::clone(commit.content_xml()))
        } else {
            self.source.share()
        };
        Ok(Commit::new(
            package,
            Patch { change },
            Diagnostics {
                changed,
                changed_cells,
            },
        ))
    }
}
