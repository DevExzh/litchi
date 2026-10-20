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
        Ok(commit_of((
            patch.0.apply(self, "apply_comment_threading")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// The staged threading extension of the legacy comment `index` by the
    /// author `author_id` on the slide part `slide_part_name`.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed comment markup.
    pub fn threading(
        &self,
        slide_part_name: &str,
        author_id: u32,
        index: u32,
    ) -> Result<Option<ThreadingInfo>> {
        package::load_threading(self.staged.opc(), slide_part_name, author_id, index)
    }

    /// Create or replace the threading extension of one legacy comment.
    /// Returns the previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing comment, malformed markup, or a
    /// changed edit of a signed package.
    pub fn set(
        &mut self,
        slide_part_name: &str,
        author_id: u32,
        index: u32,
        value: ThreadingInfo,
    ) -> Result<Option<ThreadingInfo>> {
        self.staged.update(move |opc| {
            package::put_threading(opc, slide_part_name, author_id, index, value)
        })
    }

    /// Remove the threading extension of one legacy comment. Returns the
    /// removed value; removing an absent value stages nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed markup or a changed edit of a signed
    /// package.
    pub fn remove(
        &mut self,
        slide_part_name: &str,
        author_id: u32,
        index: u32,
    ) -> Result<Option<ThreadingInfo>> {
        self.staged
            .update(move |opc| package::remove_threading(opc, slide_part_name, author_id, index))
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
