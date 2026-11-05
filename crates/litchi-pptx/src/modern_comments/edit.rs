//! Immutable modern comment edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_modern_comments`](crate::Package::edit_modern_comments) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_modern_comments`](crate::Package::apply_modern_comments) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::semantic::extensions::List as Extensions;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("modern comment");

/// A staged modern comment edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a modern comment edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_modern_comments(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_modern_comments")?,
        })
    }

    /// Apply a modern comment patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_modern_comments(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_modern_comments")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn comments(&self) -> Result<Vec<Part>> {
        load_modern_comments(self.staged.opc())
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn graph(&self) -> Result<Graph> {
        load_modern_comment_graph(self.staged.opc())
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn authors(&self) -> Result<Option<AuthorPart>> {
        load_modern_comment_authors(self.staged.opc())
    }

    /// Add a new modern Comment part after validating the complete existing graph.
    /// Existing parts are deliberately not overwritten.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Part) -> Result<()> {
        self.staged
            .update(move |opc| store_modern_comment(opc, value))
    }

    /// Add a new modern Comment Author part after validating the complete graph.
    /// Existing Author parts are deliberately not overwritten.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set_authors(&mut self, value: &AuthorPart) -> Result<()> {
        self.staged
            .update(move |opc| store_modern_comment_authors(opc, value))
    }

    /// The staged modern comment `comment_id` on the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comment<'s>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
    ) -> Result<Option<Comment>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        find_modern_comment(self.staged.opc(), &slide_part_name, comment_id)
    }

    /// Add a modern comment to a slide, creating a collision-safe part when necessary.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: Comment,
    ) -> Result<Part> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| add_modern_comment(opc, &slide_part_name, comment))
    }

    /// Update a modern comment without permitting its stable GUID to change.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Comment),
    {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| update_modern_comment(opc, &slide_part_name, comment_id, update))
    }

    /// Replace a modern comment without permitting its stable GUID to change.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        replacement: Comment,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            replace_modern_comment(opc, &slide_part_name, comment_id, replacement)
        })
    }

    /// Remove a modern comment and remove an empty per-slide part unless another owner shares it.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| remove_modern_comment(opc, &slide_part_name, comment_id))
    }

    /// Reorder every modern comment in one slide part by a complete GUID list.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        ordered_comment_ids: &[String],
    ) -> Result<Vec<Comment>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| reorder_modern_comments(opc, &slide_part_name, ordered_comment_ids))
    }

    /// Find a reply by its stable GUID within a modern comment thread.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reply<'s>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<Option<Reply>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        find_modern_comment_reply(self.staged.opc(), &slide_part_name, comment_id, reply_id)
    }

    /// Add a reply to a modern comment thread.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply: Reply,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| add_modern_comment_reply(opc, &slide_part_name, comment_id, reply))
    }

    /// Update a reply without permitting its stable GUID to change.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_reply<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Reply),
    {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            update_modern_comment_reply(opc, &slide_part_name, comment_id, reply_id, update)
        })
    }

    /// Replace a reply without permitting its stable GUID to change.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
        replacement: Reply,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            replace_modern_comment_reply(opc, &slide_part_name, comment_id, reply_id, replacement)
        })
    }

    /// Remove a reply from a modern comment thread.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            remove_modern_comment_reply(opc, &slide_part_name, comment_id, reply_id)
        })
    }

    /// Read the typed extension envelope for one comment without exposing
    /// relationship or collaboration behavior.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn extensions<'s>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
    ) -> Result<Option<Extensions>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        load_modern_comment_extensions(self.staged.opc(), &slide_part_name, comment_id)
    }

    /// Transactionally update one comment's typed task/reaction extension
    /// envelope while retaining all opaque extension entries.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Extensions),
    {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            update_modern_comment_extensions(opc, &slide_part_name, comment_id, update)
        })
    }

    /// Read the typed extension envelope for one reply.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn reply_extensions<'s>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<Option<Extensions>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        load_modern_comment_reply_extensions(
            self.staged.opc(),
            &slide_part_name,
            comment_id,
            reply_id,
        )
    }

    /// Transactionally update one reply's typed reaction extension envelope.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_reply_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut Extensions),
    {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            update_modern_comment_reply_extensions(
                opc,
                &slide_part_name,
                comment_id,
                reply_id,
                update,
            )
        })
    }

    /// The staged modern comment author `author_id`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn author(&self, author_id: &str) -> Result<Option<Author>> {
        find_modern_comment_author(self.staged.opc(), author_id)
    }

    /// Add a modern comment author, allocating a collision-safe part and relationship if needed.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_author(&mut self, author: Author) -> Result<AuthorPart> {
        self.staged
            .update(move |opc| add_modern_comment_author(opc, author))
    }

    /// Update a modern comment author while keeping its stable GUID unchanged.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_author<F>(&mut self, author_id: &str, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Author),
    {
        self.staged
            .update(move |opc| update_modern_comment_author(opc, author_id, update))
    }

    /// Replace a modern comment author while keeping its stable GUID unchanged.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_author(&mut self, author_id: &str, replacement: Author) -> Result<bool> {
        self.staged
            .update(move |opc| replace_modern_comment_author(opc, author_id, replacement))
    }

    /// Remove an unreferenced modern comment author.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_author(&mut self, author_id: &str) -> Result<bool> {
        self.staged
            .update(move |opc| remove_modern_comment_author(opc, author_id))
    }

    /// Reorder modern comment authors by a complete, duplicate-free GUID list.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_authors(&mut self, ordered_author_ids: &[String]) -> Result<Vec<Author>> {
        self.staged
            .update(move |opc| reorder_modern_comment_authors(opc, ordered_author_ids))
    }

    /// Load every typed modern-comment command while leaving all other change XML
    /// untouched and inert.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn changes(&self) -> Result<Vec<ChangeCommand>> {
        load_modern_comment_changes(self.staged.opc(), self.staged.changes_limits())
    }

    /// Update typed modern-comment commands transactionally.
    ///
    /// The callback may edit typed values but cannot add/remove/reorder command
    /// fragments. Such structural edits would risk changing unrelated future
    /// command data and are rejected. No command is interpreted or executed.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_changes<F>(&mut self, update: F) -> Result<bool>
    where
        F: FnOnce(&mut Vec<ChangeCommand>),
    {
        let limits = self.staged.changes_limits();
        self.staged
            .update(move |opc| update_modern_comment_changes(opc, limits, update))
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
