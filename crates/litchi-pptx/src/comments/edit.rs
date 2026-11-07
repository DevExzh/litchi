//! Immutable legacy comment edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_comments`](crate::Package::edit_comments) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_comments`](crate::Package::apply_comments) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("legacy comment");

/// A staged legacy comment edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a legacy comment edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_comments(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_comments")?,
        })
    }

    /// Apply a legacy comment patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_comments(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_comments",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .comments()
                    .constrained_by_source(target.read_limits());
                patch.validate_changed_parts(target, |_, part| validate_replay_part(part, limits))
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    fn limits(&self) -> Limits {
        self.staged
            .semantic_limits()
            .comments()
            .constrained_by_source(self.staged.opc().read_limits())
    }

    /// Load the presentation's legacy comment graph.
    ///
    /// The read applies the `PowerPoint`-compatible comment repairs of audit
    /// finding 6.12 ([`CommentDefect`](crate::leniency::CommentDefect)): a comment with `idx="0"` or one
    /// repeating an earlier `(authorId, idx)` key is kept instead of failing the
    /// whole graph. [`load_presentation_comments_with_tolerance`] reports what
    /// was repaired; the verbs of this edit load through
    /// [`load_presentation_comments_strict`] and refuse a graph that needed a
    /// repair, so an edit never silently rewrites (or drops) such comments.
    ///
    /// [`Comments`] stays a plain value with public fields (callers construct
    /// it to store a new graph), so the report travels beside it rather than
    /// inside it.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn comments(&self) -> Result<Option<Comments>> {
        load_presentation_comments_with_tolerance_and_limits(
            self.staged.opc(),
            &mut crate::leniency::ToleranceReport::default(),
            self.limits(),
        )
    }

    /// The staged comment author `author_id`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn author(&self, author_id: u32) -> Result<Option<Author>> {
        Ok(load_presentation_comments_with_tolerance_and_limits(
            self.staged.opc(),
            &mut crate::leniency::ToleranceReport::default(),
            self.limits(),
        )?
        .and_then(|graph| {
            graph
                .authors
                .into_iter()
                .find(|author| author.id == author_id)
        }))
    }

    /// The staged legacy comment `index` (its `p:cm/@idx`) by the author `author_id` on the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comment<'s>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author_id: u32,
        index: u32,
    ) -> Result<Option<Comment>> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        Ok(load_presentation_comments_with_tolerance_and_limits(
            self.staged.opc(),
            &mut crate::leniency::ToleranceReport::default(),
            self.limits(),
        )?
        .and_then(|graph| {
            graph
                .slides
                .into_iter()
                .find(|slide| slide.slide_part_name == slide_part_name.as_str())
                .and_then(|slide| {
                    slide
                        .comments
                        .into_iter()
                        .find(|comment| comment.author_id == author_id && comment.index == index)
                })
        }))
    }

    /// Create the complete legacy comment graph (authors and slide comments).
    ///
    /// The package must not already contain a legacy graph. Use the author and
    /// comment update/remove/reorder verbs to edit an existing graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Comments, conformance: Conformance) -> Result<()> {
        let limits = self.limits();
        self.staged.update(move |opc| {
            store_presentation_comments_with_limits(opc, value, conformance, limits)
        })
    }

    /// Add one comment author; its ID must be new.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_author(&mut self, author: Author, conformance: Conformance) -> Result<()> {
        let limits = self.limits();
        self.staged.update(move |opc| {
            add_presentation_comment_author_with_limits(opc, author, conformance, limits)
        })
    }

    /// Replace the comment author `author_id`.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update_author(
        &mut self,
        author_id: u32,
        author: Author,
        conformance: Conformance,
    ) -> Result<()> {
        let limits = self.limits();
        self.staged.update(move |opc| {
            update_presentation_comment_author_with_limits(
                opc,
                author_id,
                author,
                conformance,
                limits,
            )
        })
    }

    /// Remove the comment author `author_id`. Returns whether it existed.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_author(&mut self, author_id: u32, conformance: Conformance) -> Result<bool> {
        let limits = self.limits();
        self.staged.update(move |opc| {
            remove_presentation_comment_author_with_limits(opc, author_id, conformance, limits)
        })
    }

    /// Reorder the comment authors to `ordered_ids`, which must name each author once.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_authors(&mut self, ordered_ids: &[u32], conformance: Conformance) -> Result<()> {
        let limits = self.limits();
        self.staged.update(move |opc| {
            reorder_presentation_comment_authors_with_limits(opc, ordered_ids, conformance, limits)
        })
    }

    /// Add one legacy comment to the selected slide.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: Comment,
        conformance: Conformance,
    ) -> Result<()> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            add_presentation_comment_with_limits(
                opc,
                slide_part_name.as_str(),
                comment,
                conformance,
                limits,
            )
        })
    }

    /// Replace the legacy comment `index` by the author `author_id` on the selected slide.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author_id: u32,
        index: u32,
        replacement: Comment,
        conformance: Conformance,
    ) -> Result<()> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            update_presentation_comment_with_limits(
                opc,
                slide_part_name.as_str(),
                author_id,
                index,
                replacement,
                conformance,
                limits,
            )
        })
    }

    /// Remove the legacy comment `index` by the author `author_id` from the selected slide. Returns whether it existed.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author_id: u32,
        index: u32,
        conformance: Conformance,
    ) -> Result<bool> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            remove_presentation_comment_with_limits(
                opc,
                slide_part_name.as_str(),
                author_id,
                index,
                conformance,
                limits,
            )
        })
    }

    /// Reorder the selected slide's legacy comments to `ordered_keys` (`(author_id, index)` pairs).
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        ordered_keys: &[(u32, u32)],
        conformance: Conformance,
    ) -> Result<()> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            reorder_presentation_comments_with_limits(
                opc,
                slide_part_name.as_str(),
                ordered_keys,
                conformance,
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
