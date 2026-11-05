//! Immutable Scenario Manager edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_scenarios`] starts an [`Edit`]; [`Edit::set_scenarios`]
//! adds or replaces a worksheet's Scenario Manager and
//! [`Edit::remove_scenarios`] removes it. Every other worksheet record keeps
//! its order and bytes, and the scenario owner validates the complete
//! candidate worksheet before it is staged. Scenario values are inert: they
//! are never applied to cells or recalculated. [`Edit::commit`] returns a
//! [`Commit`]; [`Package::apply_scenarios`] replays its [`Patch`]. An exact
//! no-op keeps every byte, and a changed edit of a signed workbook is
//! refused with [`Error::Signed`].

use super::Manager;
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("Scenario Manager");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Scenarios of one worksheet ([`super::MAX_SCENARIOS`]).
    Scenarios,
}

/// A staged Scenario Manager edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a Scenario Manager edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_scenarios(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
        })
    }

    /// Apply a Scenario Manager patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_scenarios(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged Scenario Manager of the worksheet at complete-catalog
    /// position `sheet`, if any.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab or a malformed
    /// scenario collection.
    pub fn scenarios(&self, sheet: usize) -> Result<Option<Manager>> {
        let worksheet = self.staged.worksheet(sheet)?;
        super::parse_worksheet(self.staged.blob(&worksheet)?, self.staged.raw_limits())
    }

    /// Add or replace a worksheet's Scenario Manager. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, an invalid
    /// manager, a malformed or ambiguous scenario collection, and a changed
    /// edit of a signed workbook.
    pub fn set_scenarios(&mut self, sheet: usize, scenarios: &Manager) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let updated = super::replace_worksheet(
            self.staged.blob(&worksheet)?,
            Some(scenarios),
            self.staged.raw_limits(),
        )?;
        super::parse_worksheet(&updated, self.staged.raw_limits())?;
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Remove a worksheet's Scenario Manager. Returns whether one was
    /// removed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a malformed
    /// scenario collection, and a changed edit of a signed workbook.
    pub fn remove_scenarios(&mut self, sheet: usize) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let updated = super::replace_worksheet(
            self.staged.blob(&worksheet)?,
            None,
            self.staged.raw_limits(),
        )?;
        super::parse_worksheet(&updated, self.staged.raw_limits())?;
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Validate the staged worksheets and publish them to a new package. An
    /// unchanged staging returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed staging of a signed workbook
    /// and the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
