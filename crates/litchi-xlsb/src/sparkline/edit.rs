//! Immutable sparkline edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_sparklines`] starts an [`Edit`]; [`Edit::set`] inserts or
//! replaces a worksheet's sparkline collection and [`Edit::remove`] removes
//! it. Each change is fully encoded, read back and its formulas proven
//! against the workbook's defined names and supporting links (never
//! evaluated); every record outside the sparkline block keeps its bytes.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_sparklines`] replays a patch onto an equal package. An
//! exact no-op keeps every byte, and a changed edit of a signed workbook is
//! refused with [`Error::Signed`].

use super::workbook::validate_commit_context;
use super::{Groups, Limits, worksheet};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("sparkline");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of one worksheet stream.
    WorksheetBytes,
    /// Sparkline groups of one worksheet.
    Groups,
}

/// The configuration of a sparkline edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Select the sparkline policy of the edit; the default is the package's
    /// [`crate::SemanticLimits::sparklines_value`].
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// The configured sparkline policy, if any.
    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// A staged sparkline edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    limits: Limits,
}

impl Package {
    /// Start a sparkline edit under the package's limits. Starting one
    /// copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_sparklines(&self) -> Result<Edit> {
        self.edit_sparklines_with_options(EditOptions::default())
    }

    /// Start a sparkline edit with explicit options.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_sparklines_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            limits: options
                .limits
                .unwrap_or_else(|| self.semantic_limits().sparklines_value()),
        })
    }

    /// Apply a sparkline patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_sparklines(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    fn snapshot(&self, sheet: usize) -> Result<(litchi_opc::PackURI, worksheet::Snapshot)> {
        let uri = self.staged.worksheet(sheet)?;
        let snapshot = worksheet::read(self.staged.blob(&uri)?, self.limits).map_err(|_error| {
            crate::part_edit::invalid("the worksheet sparkline records are not valid")
        })?;
        Ok((uri, snapshot))
    }

    /// The staged sparkline groups of the worksheet at complete-catalog
    /// position `sheet`, if it owns a sparkline block.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab or a malformed
    /// sparkline block.
    pub fn groups(&self, sheet: usize) -> Result<Option<Groups>> {
        Ok(self.snapshot(sheet)?.1.groups().cloned())
    }

    fn stage(&mut self, sheet: usize, groups: Option<Groups>) -> Result<bool> {
        let (uri, snapshot) = self.snapshot(sheet)?;
        let mut draft = snapshot.edit();
        match groups {
            Some(groups) => {
                draft.set(groups);
            },
            None => {
                draft.remove();
            },
        }
        let commit = draft.commit().map_err(|_error| {
            crate::part_edit::invalid("the worksheet sparkline records are not valid")
        })?;
        let workbook = self.staged.workbook()?;
        validate_commit_context(&commit, &workbook.formula_context)?;
        if commit.is_noop() {
            return Ok(false);
        }
        let (_, _, patch) = commit.into_publication();
        let (_, updated) = patch
            .apply_owned(self.staged.blob(&uri)?)
            .map_err(|_error| {
                crate::part_edit::invalid("the worksheet sparkline records are not valid")
            })?;
        self.staged.replace_worksheet(&uri, updated)
    }

    /// Insert or replace the complete sparkline collection of a worksheet.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a collection the
    /// policy refuses, a formula naming a missing name or link, and a
    /// changed edit of a signed workbook.
    pub fn set(&mut self, sheet: usize, groups: Groups) -> Result<bool> {
        self.stage(sheet, Some(groups))
    }

    /// Remove the sparkline collection of a worksheet. Returns whether one
    /// was removed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set`].
    pub fn remove(&mut self, sheet: usize) -> Result<bool> {
        self.stage(sheet, None)
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
