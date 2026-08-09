//! Failure-atomic source-bound worksheet page-break edits.

use litchi_opc::OpcPackage;

use super::{Collection, Commit, PageBreaks, Patch, Snapshot};
use crate::Selector;
use crate::error::{Error, Result, invalid};

/// One isolated page-break transaction over a selected worksheet.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    staged: PageBreaks,
}

impl<'a> Transaction<'a> {
    /// Resolve a worksheet and start an isolated transaction.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or selector is invalid, the selected
    /// sheet is not a worksheet, or its page-break XML is invalid.
    pub fn new<'selector>(
        target: &'a mut OpcPackage,
        selector: impl Into<Selector<'selector>>,
    ) -> Result<Self> {
        let before = Snapshot::load(target, selector)?;
        let staged = before.page_breaks().clone();
        Ok(Self {
            target,
            before,
            staged,
        })
    }

    /// Exact source state captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-break state.
    #[must_use]
    pub const fn page_breaks(&self) -> &PageBreaks {
        &self.staged
    }

    /// Replace the complete staged page-break state.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` violates page-break invariants.
    pub fn set(&mut self, value: PageBreaks) -> Result<bool> {
        value.validate()?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Clone-edit the complete staged value; a failed closure changes nothing.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or a page-break validation error.
    pub fn edit(&mut self, edit: impl FnOnce(&mut PageBreaks) -> Result<()>) -> Result<bool> {
        let mut draft = self.staged.clone();
        edit(&mut draft)?;
        self.set(draft)
    }

    /// Replace horizontal row breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not horizontal or is invalid.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_horizontal(collection))
    }

    /// Replace vertical column breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not vertical or is invalid.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_vertical(collection))
    }

    /// Remove horizontal row breaks.
    pub fn remove_horizontal(&mut self) -> bool {
        self.staged.remove_horizontal()
    }

    /// Remove vertical column breaks.
    pub fn remove_vertical(&mut self) -> bool {
        self.staged.remove_vertical()
    }

    /// Remove both authored collections.
    pub fn clear(&mut self) -> bool {
        self.staged.clear()
    }

    /// Whether the exact authored semantic state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.page_breaks() != &self.staged
    }

    /// Validate, rewrite one worksheet, read it back, and atomically publish.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes, signed packages, invalid
    /// staged state, rewrite failures, or semantic readback failures.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        if !self.before.matches_current_source(self.target) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet_part_name().to_string(),
            });
        }
        let output = super::replace(self.before.source_xml(), &self.staged)?;
        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(self.before.worksheet_part_name())?
            .set_blob(output);
        let snapshot = Snapshot::load(&candidate, self.before.sheet_position())?;
        if snapshot.page_breaks() != &self.staged {
            return Err(invalid(
                "page-break publication changed the staged semantic state",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}
