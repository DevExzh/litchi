//! Immutable comment-author presence edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_author_presence`] starts an [`Edit`] over a staged copy of the
//! package; [`Edit::set`] creates or replaces one comment author's `p15:presenceInfo` extension and [`Edit::remove`]
//! removes it. [`Edit::commit`] returns a [`Commit`] with the new package,
//! the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_author_presence`] replays a patch onto an equal package. The
//! source package never changes, an exact no-op keeps every byte, and a
//! changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use super::PresenceInfo;
use super::package;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("comment-author presence");

/// A staged comment-author presence edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a comment-author presence edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_author_presence(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_author_presence")?,
        })
    }

    /// Apply a comment-author presence patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_author_presence(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_author_presence",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .comments()
                    .constrained_by_source(target.read_limits());
                patch.validate_changed_parts(target, |_, part| {
                    if part.content_type() == crate::comments::AUTHORS_CONTENT_TYPE {
                        super::validate_presence_replay(part.blob(), limits)
                    } else {
                        crate::comments::validate_replay_part(part, limits)
                    }
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    fn limits(&self) -> super::super::Limits {
        self.staged
            .semantic_limits()
            .comments()
            .constrained_by_source(self.staged.opc().read_limits())
    }

    /// The staged presence extension of the comment author `author_id`.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed comment-author markup.
    pub fn presence(&self, author_id: u32) -> Result<Option<PresenceInfo>> {
        Ok(package::load_presence_snapshot_with_limits(
            self.staged.opc(),
            author_id,
            self.limits(),
        )?
        .and_then(|snapshot| snapshot.value().cloned()))
    }

    /// Create or replace the presence extension of the comment author
    /// `author_id`. Returns the previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing author, malformed markup, or a changed
    /// edit of a signed package.
    pub fn set(&mut self, author_id: u32, value: PresenceInfo) -> Result<Option<PresenceInfo>> {
        let limits = self.limits();
        self.staged
            .update(move |opc| package::put_presence_with_limits(opc, author_id, value, limits))
    }

    /// Remove the presence extension of the comment author `author_id`.
    /// Returns the removed value; removing an absent value stages nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed markup or a changed edit of a signed
    /// package.
    pub fn remove(&mut self, author_id: u32) -> Result<Option<PresenceInfo>> {
        let limits = self.limits();
        self.staged
            .update(move |opc| package::remove_presence_with_limits(opc, author_id, limits))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        Ok(commit_of(self.staged.commit()?))
    }
}
