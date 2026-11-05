//! Immutable comment threading edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_comment_threading`] starts an [`Edit`] over a staged copy of the
//! package; [`Edit::set`] creates or replaces one legacy comment's `p15:threadingInfo` extension and [`Edit::remove`]
//! removes it. [`Edit::commit`] returns a [`Commit`] with the new package,
//! the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_comment_threading`] replays a patch onto an equal package. The
//! source package never changes, an exact no-op keeps every byte, and a
//! changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use super::ThreadingInfo;
use super::package;
use crate::part_edit::{Staged, family};
use crate::slide::Selector;
use crate::{Package, Result};

family!("comment threading");

/// A staged comment threading edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a comment threading edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_comment_threading(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_comment_threading")?,
        })
    }

    /// Apply a comment threading patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_comment_threading(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_comment_threading",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .comments()
                    .constrained_by_source(target.read_limits());
                patch.validate_changed_parts(target, |_, part| {
                    if part.content_type() == crate::comments::COMMENTS_CONTENT_TYPE {
                        super::validate_threading_replay(part.blob(), limits)
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

    /// The staged threading extension of the legacy comment `index` (its
    /// `p:cm/@idx`) by the author `author_id` on the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed comment markup.
    pub fn threading<'s>(
        &self,
        slide: impl Into<Selector<'s>>,
        author_id: u32,
        index: u32,
    ) -> Result<Option<ThreadingInfo>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        Ok(package::load_threading_snapshot_with_limits(
            self.staged.opc(),
            slide_part_name.as_str(),
            author_id,
            index,
            self.limits(),
        )?
        .and_then(|snapshot| snapshot.value().cloned()))
    }

    /// Create or replace the threading extension of one legacy comment.
    /// Returns the previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing comment, malformed markup, or a
    /// changed edit of a signed package.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        author_id: u32,
        index: u32,
        value: ThreadingInfo,
    ) -> Result<Option<ThreadingInfo>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            package::put_threading_with_limits(
                opc,
                slide_part_name.as_str(),
                author_id,
                index,
                value,
                limits,
            )
        })
    }

    /// Remove the threading extension of one legacy comment. Returns the
    /// removed value; removing an absent value stages nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed markup or a changed edit of a signed
    /// package.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        author_id: u32,
        index: u32,
    ) -> Result<Option<ThreadingInfo>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            package::remove_threading_with_limits(
                opc,
                slide_part_name.as_str(),
                author_id,
                index,
                limits,
            )
        })
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
