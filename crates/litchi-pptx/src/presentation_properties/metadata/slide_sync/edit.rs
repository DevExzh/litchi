//! Immutable slide-synchronization edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_slide_sync`](crate::Package::edit_slide_sync) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_slide_sync`](crate::Package::apply_slide_sync) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide-synchronization");

/// Per-edit bounds for Slide Synchronization operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Select an optional tighter Slide Synchronization profile.
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// Return the explicitly selected profile, if any.
    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// A staged slide-synchronization edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    limits: Limits,
}

impl Package {
    /// Start a slide-synchronization edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_slide_sync(&self) -> Result<Edit<'_>> {
        self.edit_slide_sync_with_options(EditOptions::default())
    }

    /// Start a Slide Synchronization edit under explicit operation options.
    pub fn edit_slide_sync_with_options(&self, options: EditOptions) -> Result<Edit<'_>> {
        let staged = Staged::new(self, "edit_slide_sync")?;
        let limits = options
            .limits_value()
            .unwrap_or_else(|| staged.slide_sync_limits())
            .constrained_by_source(self.opc.read_limits());
        Ok(Edit { staged, limits })
    }

    /// Apply a slide-synchronization patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_slide_sync(&self, patch: &Patch) -> Result<Commit> {
        let semantic_limits = self.semantic_limits();
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_slide_sync",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part_policy(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::SlideSync,
                        name,
                        part,
                        semantic_limits,
                    )
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Load every synchronization part and validate its slide relationship graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn slide_sync(&self) -> Result<Vec<Part>> {
        load(self.staged.opc(), self.limits)
    }

    /// Attach one synchronization part to its source slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Part) -> Result<()> {
        let limits = self.limits;
        self.staged
            .update(move |opc| raw::store(opc, value, limits))
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
