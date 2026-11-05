//! The source-backed worksheet-visibility family: same-length
//! `BoundSheet8` visibility splices (the former
//! `sheet_visibility::Edit::commit_source_backed`).

use crate::Result;
use crate::sheet_visibility::Selector;
use crate::{SheetVisibility, SourceBackedWorkbook};

source_family!("worksheet visibility", "edit sheet visibility");

/// Content-free source-backed visibility publication diagnostics.
pub use crate::sheet_visibility::source_diagnostics::Diagnostics;

/// A staged source-backed worksheet-visibility edit. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    source: SourceBackedWorkbook,
    inner: Option<crate::sheet_visibility::Edit>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("materialized", &self.inner.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn inner_mut(&mut self) -> Result<&mut crate::sheet_visibility::Edit> {
        if self.inner.is_none() {
            let package = self.source.staging_package().map_err(|error| {
                crate::source_backed::protected_edit_error(error, "edit sheet visibility")
            })?;
            let inner = package.edit_sheet_visibility().map_err(|error| {
                crate::source_backed::protected_edit_error(error, "edit sheet visibility")
            })?;
            self.inner = Some(inner);
        }
        self.inner.as_mut().ok_or(Error::Internal(
            "source-backed visibility edit engine was not retained",
        ))
    }

    /// Stage one exact visible, hidden or very-hidden outcome.
    ///
    /// # Errors
    ///
    /// Returns a missing, ambiguous or non-worksheet selector refusal.
    pub fn set_visibility(
        &mut self,
        selector: Selector<'_>,
        visibility: SheetVisibility,
    ) -> Result<&mut Self> {
        self.inner_mut()?.set_visibility(selector, visibility)?;
        Ok(self)
    }

    /// Stage a hidden worksheet.
    ///
    /// # Errors
    ///
    /// As [`Self::set_visibility`].
    pub fn hide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::Hidden)
    }

    /// Stage a very-hidden worksheet.
    ///
    /// # Errors
    ///
    /// As [`Self::set_visibility`].
    pub fn very_hide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::VeryHidden)
    }

    /// Stage a visible worksheet.
    ///
    /// # Errors
    ///
    /// As [`Self::set_visibility`].
    pub fn unhide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::Visible)
    }

    /// Atomically stage several visibility outcomes.
    ///
    /// # Errors
    ///
    /// Returns the first refusal; nothing is staged.
    pub fn set_visibility_batch<'a, I>(&mut self, changes: I) -> Result<&mut Self>
    where
        I: IntoIterator<Item = (Selector<'a>, SheetVisibility)>,
    {
        self.inner_mut()?.set_visibility_batch(changes)?;
        Ok(self)
    }

    /// Validate the staged edits and commit them as a workbook that writes
    /// the checked overlay plan. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns a refusal for a protected, signed, encrypted, stale or
    /// last-visible-sheet edit, or a failed readback.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source.clone();
        source.ensure_original_source_current()?;
        let inner = match self.inner {
            Some(inner) => inner,
            None => {
                let package = source.staging_package().map_err(|error| {
                    crate::source_backed::protected_edit_error(error, "edit sheet visibility")
                })?;
                package.edit_sheet_visibility().map_err(|error| {
                    crate::source_backed::protected_edit_error(error, "edit sheet visibility")
                })?
            },
        };
        let (plan, diagnostics) = inner.commit_source_backed()?.into_plan();
        source.ensure_original_source_current()?;
        let commit = published(&source, plan, diagnostics)?;
        source.ensure_original_source_current()?;
        Ok(commit)
    }
}

impl SourceBackedWorkbook {
    /// Start a source-backed worksheet-visibility edit of this workbook.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook's sheet table cannot be read.
    pub fn edit_sheet_visibility(&self) -> Result<Edit> {
        self.ensure_original_source_current()?;
        Ok(Edit {
            source: self.clone(),
            inner: None,
        })
    }

    /// Replay a source-backed worksheet-visibility patch onto matching
    /// workbook content, including an independently opened equal-content
    /// workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for another workbook.
    pub fn apply_sheet_visibility(&self, patch: &Patch) -> Result<Commit> {
        replay(self, patch)
    }
}
