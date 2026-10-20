//! The conflict-markup family (record 0355): Word conflict annotations and
//! paired conflict ranges of one story.
//!
//! [`Package::conflicts`](crate::Package::conflicts) reads the main story and
//! [`Package::conflict_stories`](crate::Package::conflict_stories) every
//! story; [`Package::edit_conflicts`](crate::Package::edit_conflicts) stages
//! an [`Edit`] of the main story ([`Package::edit_story_conflicts`](crate::Package::edit_story_conflicts)
//! of another one) whose commit publishes a new package, and
//! [`Package::apply_conflicts`](crate::Package::apply_conflicts) replays its
//! [`Patch`].

use litchi_opc::PackURI;

use super::transaction::{Draft, Operation};
use crate::package::family::{Staged, graph_family};

graph_family!("conflict-markup", own_limits);

pub use super::limit::LimitKind;

/// Options of a conflict-markup edit (record 0355a).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<super::Limits>,
}

impl EditOptions {
    /// Bound the edit's story snapshot and its rewrite (including
    /// `max_output_bytes`) by `limits` instead of the package's semantic
    /// conflict limits.
    #[must_use]
    pub const fn limits(mut self, limits: super::Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// The configured limits; `None` uses the package's semantic limits.
    #[must_use]
    pub const fn limits_value(&self) -> Option<super::Limits> {
        self.limits
    }
}

/// A staged conflict-markup change of one story over one package snapshot.
/// The source is never changed: [`Self::commit`] publishes a new package,
/// and dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    story: Option<PackURI>,
    limits: super::Limits,
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
            let package = self.staged.working();
            let snapshot = match &self.story {
                None => package.conflicts_bounded(self.limits)?,
                Some(part) => package
                    .conflict_stories_bounded(self.limits)?
                    .into_iter()
                    .find(|story| story.part() == part)
                    .ok_or_else(|| crate::Error::PartNotFound(format!("conflict story {part}")))?
                    .snapshot()
                    .clone(),
            };
            self.draft = Some(snapshot.edit());
        }
        Ok(self
            .draft
            .as_mut()
            .expect("conflict draft was created above"))
    }

    /// Stage one operation.
    ///
    /// # Errors
    ///
    /// Returns an error for an operation the story's markup refuses.
    pub fn apply(&mut self, operation: Operation) -> Result<&mut Self> {
        self.draft()?.apply(operation)?;
        Ok(self)
    }

    /// Replace the author of the conflict element at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the conflicts or an invalid author.
    pub fn set_conflict_author(
        &mut self,
        index: usize,
        author: impl Into<String>,
    ) -> Result<&mut Self> {
        self.draft()?.set_conflict_author(index, author)?;
        Ok(self)
    }

    /// Replace or remove the date of the conflict element at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the conflicts or an invalid date.
    pub fn set_conflict_date(&mut self, index: usize, date: Option<String>) -> Result<&mut Self> {
        self.draft()?.set_conflict_date(index, date)?;
        Ok(self)
    }

    /// Remove the conflict annotation at `index`, keeping its inline payload.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the conflicts.
    pub fn remove_conflict(&mut self, index: usize) -> Result<&mut Self> {
        self.draft()?.remove_conflict(index)?;
        Ok(self)
    }

    /// Replace the author on the start marker of the range at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the ranges or an invalid author.
    pub fn set_range_author(
        &mut self,
        index: usize,
        author: impl Into<String>,
    ) -> Result<&mut Self> {
        self.draft()?.set_range_author(index, author)?;
        Ok(self)
    }

    /// Replace or remove the date on the start marker of the range at
    /// `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the ranges or an invalid date.
    pub fn set_range_date(&mut self, index: usize, date: Option<String>) -> Result<&mut Self> {
        self.draft()?.set_range_date(index, date)?;
        Ok(self)
    }

    /// Remove both markers of the range at `index`, keeping its payload.
    ///
    /// # Errors
    ///
    /// Returns an error for an index past the ranges.
    pub fn remove_range(&mut self, index: usize) -> Result<&mut Self> {
        self.draft()?.remove_range(index)?;
        Ok(self)
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
    /// Returns an error when the story rewrite cannot preserve the markup.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some(draft) = self.draft.take() {
            let planned = draft.commit()?;
            let mut working = self.staged.working().clone();
            working.publish_conflict_draft(&planned)?;
            *self.staged.working_mut() = working;
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a conflict-markup change of the main story. O(1).
    #[must_use]
    pub fn edit_conflicts(&self) -> Edit {
        self.edit_conflicts_with_options(EditOptions::default())
    }

    /// Stage a conflict-markup change of the main story under `options`.
    /// O(1).
    #[must_use]
    pub fn edit_conflicts_with_options(&self, options: EditOptions) -> Edit {
        self.conflict_edit(None, options)
    }

    fn conflict_edit(&self, story: Option<PackURI>, options: EditOptions) -> Edit {
        Edit {
            staged: Staged::new(self),
            story,
            limits: options
                .limits
                .unwrap_or_else(|| self.semantic_limits().conflicts_value()),
            draft: None,
        }
    }

    /// Stage a conflict-markup change of the story part `part` (one of
    /// [`Self::conflict_stories`]). O(1).
    #[must_use]
    pub fn edit_story_conflicts(&self, part: &PackURI) -> Edit {
        self.edit_story_conflicts_with_options(part, EditOptions::default())
    }

    /// Stage a conflict-markup change of the story part `part` under
    /// `options`. O(1).
    #[must_use]
    pub fn edit_story_conflicts_with_options(&self, part: &PackURI, options: EditOptions) -> Edit {
        self.conflict_edit(Some(part.clone()), options)
    }

    /// Replay a conflict-markup patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_conflicts(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_conflicts")
    }
}
