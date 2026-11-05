//! The source-backed comment family: same-length `TXO`/`CONTINUE` comment
//! text and `NOTE` visibility splices (the former
//! `comments::Edit::commit_source_backed`).

use crate::Result;
use crate::SourceBackedWorkbook;
use crate::cell_values::{Reference, Selector};
use crate::comments::{Update, Value, Visibility};

source_family!("comment", "edit comments");

/// Content-free source-backed comment publication diagnostics.
pub use crate::comments::source_diagnostics::Diagnostics;

/// A staged source-backed comment edit. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    source: SourceBackedWorkbook,
    inner: Option<crate::comments::Edit>,
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
    fn inner_mut(&mut self) -> Result<&mut crate::comments::Edit> {
        if self.inner.is_none() {
            let package = self.source.staging_package().map_err(|error| {
                crate::source_backed::protected_edit_error(error, "edit comments")
            })?;
            let inner = package.edit_comments().map_err(|error| {
                crate::source_backed::protected_edit_error(error, "edit comments")
            })?;
            self.inner = Some(inner);
        }
        self.inner.as_mut().ok_or(Error::Internal(
            "source-backed comment edit engine was not retained",
        ))
    }

    /// Replace the complete text and author of one existing comment.
    ///
    /// # Errors
    ///
    /// Returns a selector, protection, ambiguity, encoding or bound refusal.
    pub fn replace(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        self.inner_mut()?.replace(selector, reference, value)
    }

    /// Atomically stage up to 256 distinct replacements on one worksheet.
    ///
    /// # Errors
    ///
    /// Returns the first refusal; nothing is staged.
    pub fn replace_many(
        &mut self,
        selector: Selector<'_>,
        updates: impl IntoIterator<Item = Update>,
    ) -> Result<()> {
        self.inner_mut()?.replace_many(selector, updates)
    }

    /// Stage an existing comment's visibility.
    ///
    /// # Errors
    ///
    /// Returns a selector, missing-comment or ownership refusal.
    pub fn set_visibility(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        visibility: Visibility,
    ) -> Result<()> {
        self.inner_mut()?
            .set_visibility(selector, reference, visibility)
    }

    /// Validate the staged edits and commit them as a workbook that writes
    /// the checked overlay plan. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns a refusal for a protected, signed, encrypted, stale or
    /// length-changing edit, or a failed readback.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source.clone();
        source.ensure_original_source_current()?;
        let inner = match self.inner {
            Some(inner) => inner,
            None => {
                let package = source.staging_package().map_err(|error| {
                    crate::source_backed::protected_edit_error(error, "edit comments")
                })?;
                package.edit_comments().map_err(|error| {
                    crate::source_backed::protected_edit_error(error, "edit comments")
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
    /// Start a source-backed comment edit of this workbook.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook's comments cannot be read.
    pub fn edit_comments(&self) -> Result<Edit> {
        self.ensure_original_source_current()?;
        Ok(Edit {
            source: self.clone(),
            inner: None,
        })
    }

    /// Replay a source-backed comment patch onto matching workbook content,
    /// including an independently opened equal-content workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for another workbook.
    pub fn apply_comments(&self, patch: &Patch) -> Result<Commit> {
        replay(self, patch)
    }
}
