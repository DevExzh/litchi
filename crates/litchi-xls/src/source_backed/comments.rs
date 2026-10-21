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
    inner: crate::comments::Edit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
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
        self.inner.replace(selector, reference, value)
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
        self.inner.replace_many(selector, updates)
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
        self.inner.set_visibility(selector, reference, visibility)
    }

    /// Validate the staged edits and commit them as a workbook that writes
    /// the checked overlay plan. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns a refusal for a protected, signed, encrypted, stale or
    /// length-changing edit, or a failed readback.
    pub fn commit(self) -> Result<Commit> {
        let (plan, diagnostics) = self.inner.commit_source_backed()?.into_plan();
        published(&self.source, plan, diagnostics)
    }
}

impl SourceBackedWorkbook {
    /// Start a source-backed comment edit of this workbook.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook's comments cannot be read.
    pub fn edit_comments(&self) -> Result<Edit> {
        Ok(Edit {
            source: self.clone(),
            inner: self.staging_package()?.edit_comments()?,
        })
    }

    /// Replay a source-backed comment patch onto the exact workbook it was
    /// made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for another workbook.
    pub fn apply_comments(&self, patch: &Patch) -> Result<Commit> {
        replay(self, patch)
    }
}
