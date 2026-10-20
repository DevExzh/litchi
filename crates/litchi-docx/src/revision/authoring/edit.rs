//! The tracked-revision family (record 0355): accepting and rejecting the
//! main story's tracked insertions, deletions, moves and property changes.
//!
//! [`Package::tracked_revisions`](crate::Package::tracked_revisions) reads
//! the changes; [`Package::edit_tracked_revisions`](crate::Package::edit_tracked_revisions)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_tracked_revisions`](crate::Package::apply_tracked_revisions)
//! replays its [`Patch`].

use super::{Disposition, Draft, Snapshot};
use crate::package::family::{Staged, graph_family};

graph_family!("tracked-revision", own_limits);

/// The budgets of the tracked-revision family, reported by
/// [`crate::Error::RevisionLimit`].
pub use crate::revision::LimitKind;

/// A staged tracked-revision disposition over one package snapshot. The
/// source is never changed: [`Self::commit`] publishes a new package, and
/// dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    draft: Option<Draft>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn draft(&mut self) -> Result<&mut Draft> {
        if self.draft.is_none() {
            self.draft = Some(self.staged.working().tracked_revisions()?.edit());
        }
        Ok(self
            .draft
            .as_mut()
            .expect("tracked-revision draft was created above"))
    }

    /// Accept the change at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for an index past the changes.
    pub fn accept(&mut self, index: usize) -> Result<&mut Self> {
        self.draft()?.accept(index)?;
        Ok(self)
    }

    /// Reject the change at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for an index past the changes.
    pub fn reject(&mut self, index: usize) -> Result<&mut Self> {
        self.draft()?.reject(index)?;
        Ok(self)
    }

    /// Stage `disposition` for the change at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::OutOfBounds`] for an index past the changes.
    pub fn set(&mut self, index: usize, disposition: Disposition) -> Result<&mut Self> {
        self.draft()?.set(index, disposition)?;
        Ok(self)
    }

    /// Accept every change.
    ///
    /// # Errors
    ///
    /// Returns an error when the main story cannot be read.
    pub fn accept_all(&mut self) -> Result<&mut Self> {
        self.draft()?.accept_all()?;
        Ok(self)
    }

    /// Reject every change.
    ///
    /// # Errors
    ///
    /// Returns an error when the main story cannot be read.
    pub fn reject_all(&mut self) -> Result<&mut Self> {
        self.draft()?.reject_all()?;
        Ok(self)
    }

    /// The changes the edit selects from.
    ///
    /// # Errors
    ///
    /// Returns an error when the main story cannot be read.
    pub fn snapshot_changes(&mut self) -> Result<Snapshot> {
        self.staged.working().tracked_revisions()
    }

    /// Discard every staged disposition, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Publish the staged dispositions as a new package with their exact
    /// patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a disposition the story rewrite refuses.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some(draft) = self.draft.take() {
            let planned = draft.commit()?;
            let mut working = self.staged.working().clone();
            working.publish_tracked_revision_draft(&planned)?;
            *self.staged.working_mut() = working;
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage tracked-revision dispositions over this package. O(1).
    #[must_use]
    pub fn edit_tracked_revisions(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            draft: None,
        }
    }

    /// Replay a tracked-revision patch onto this package, which must hold
    /// the patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_tracked_revisions(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_tracked_revisions")
    }
}
