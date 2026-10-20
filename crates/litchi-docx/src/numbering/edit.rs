//! The numbering family (record 0355): the numbering definitions part.
//!
//! [`Package::numbering_snapshot`](crate::Package::numbering_snapshot) reads
//! the source-preserving [`Snapshot`](super::Snapshot);
//! [`Package::edit_numbering`](crate::Package::edit_numbering) stages an
//! [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_numbering`](crate::Package::apply_numbering) replays its
//! [`Patch`]. Only the changed attributes are rewritten.

use super::transaction::Transaction as Draft;
use crate::package::family::{Staged, graph_family};

graph_family!("numbering", own_limits);

pub use super::limit::LimitKind;

/// A staged numbering change over one package snapshot. The source is never
/// changed: [`Self::commit`] publishes a new package, and dropping the edit
/// discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    draft: Option<(super::Snapshot, Draft)>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn draft(&mut self) -> Result<&mut Draft> {
        if self.draft.is_none() {
            let snapshot = self
                .staged
                .working()
                .numbering_snapshot()?
                .ok_or_else(|| crate::Error::PartNotFound("numbering part".into()))?;
            let draft = snapshot.edit();
            self.draft = Some((snapshot, draft));
        }
        Ok(&mut self
            .draft
            .as_mut()
            .expect("numbering draft was created above")
            .1)
    }

    /// The staged `w15:restartNumberingAfterBreak` of an abstract
    /// definition.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no numbering part or no such
    /// definition.
    pub fn restart_numbering_after_break(&mut self, abstract_num_id: u32) -> Result<Option<bool>> {
        self.draft()?.restart_numbering_after_break(abstract_num_id)
    }

    /// Set or remove `w15:restartNumberingAfterBreak` of an abstract
    /// definition.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no numbering part or no such
    /// definition.
    pub fn set_restart_numbering_after_break(
        &mut self,
        abstract_num_id: u32,
        value: Option<bool>,
    ) -> Result<&mut Self> {
        self.draft()?
            .set_restart_numbering_after_break(abstract_num_id, value)?;
        Ok(self)
    }

    /// Remove `w15:restartNumberingAfterBreak` of an abstract definition.
    ///
    /// # Errors
    ///
    /// As [`Self::set_restart_numbering_after_break`].
    pub fn clear_restart_numbering_after_break(
        &mut self,
        abstract_num_id: u32,
    ) -> Result<&mut Self> {
        self.set_restart_numbering_after_break(abstract_num_id, None)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Publish the staged changes as a new package with their exact patch.
    ///
    /// # Errors
    ///
    /// Returns an error when the rewrite cannot preserve the numbering part.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some((source, draft)) = self.draft.take() {
            let patch = draft.commit()?.into_patch();
            let mut working = self.staged.working().clone();
            working.apply_numbering_patch(&source, &patch)?;
            *self.staged.working_mut() = working;
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a numbering change over this package. O(1).
    #[must_use]
    pub fn edit_numbering(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            draft: None,
        }
    }

    /// Replay a numbering patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_numbering(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_numbering")
    }
}
