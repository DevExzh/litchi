//! Immutable slide-tag edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_tags`] starts an [`Edit`]; [`Edit::set`] creates or
//! replaces one slide's direct programmable-tag list and [`Edit::remove`]
//! removes it. [`Edit::commit`] returns a [`Commit`] with the new package,
//! the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_tags`] replays a patch onto an equal package. The source
//! package never changes, an exact no-op keeps every byte, and a changed edit
//! of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).
//!
//! ```
//! use litchi_pptx::tag::{List, Tag};
//!
//! # fn main() -> litchi_pptx::Result<()> {
//! let mut builder = litchi_pptx::builder::Builder::new()?;
//! builder.presentation()?.add_slide()?;
//! let package = builder.build()?;
//! let mut tags = List::new();
//! tags.add(Tag::new("Owner", "Alice")?)?;
//! let mut edit = package.edit_tags()?;
//! assert_eq!(edit.set(0_usize, tags)?, None);
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let after = commit.package();
//! assert_eq!(after.tags(0_usize)?.unwrap().get("owner")?.value(), "Alice");
//! assert!(package.tags(0_usize)?.is_none());
//! # Ok(())
//! # }
//! ```

use super::List;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide-tag");

/// A staged slide-tag edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a slide-tag edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_tags(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_tags")?,
        })
    }

    /// Apply a slide-tag patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_tags(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_tags",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .tag()
                    .constrained_by_source(target.read_limits());
                patch.validate_changed_parts(target, |name, part| {
                    super::validate_replay_part(target, name, part, limits)
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// The staged direct tag list of the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, or a
    /// malformed tag graph.
    pub fn tags<'s>(&self, slide: impl Into<crate::slide::Selector<'s>>) -> Result<Option<List>> {
        let opc = self.staged.opc();
        let slide = crate::package::resolve_slide(opc, slide.into())?;
        let limits = self
            .staged
            .semantic_limits()
            .tag()
            .constrained_by_source(opc.read_limits());
        Ok(super::load_with_limits(opc, &slide, limits)?.map(super::Source::into_list))
    }

    /// Create or replace one slide's direct programmable-tag list. Returns
    /// the previous list.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, a
    /// malformed tag graph, or a changed edit of a signed package.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        list: List,
    ) -> Result<Option<List>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self
            .staged
            .semantic_limits()
            .tag()
            .constrained_by_source(self.staged.opc().read_limits());
        self.staged
            .update(move |opc| super::put_with_limits(opc, &slide, list, limits))
    }

    /// Remove one slide's direct programmable-tag list. Removal is
    /// idempotent and collects an orphaned tag part only after the
    /// package-wide inbound-edge check succeeds. Returns the removed list.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, a
    /// malformed tag graph, or a changed edit of a signed package.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
    ) -> Result<Option<List>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self
            .staged
            .semantic_limits()
            .tag()
            .constrained_by_source(self.staged.opc().read_limits());
        self.staged
            .update(move |opc| super::remove_with_limits(opc, &slide, limits))
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
