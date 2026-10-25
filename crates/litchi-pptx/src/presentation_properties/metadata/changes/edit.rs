//! Immutable Changes Information edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_changes_info`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of the presentation's Changes Information part. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_changes_info`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use super::Draft;
use super::Limits;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("Changes Information");

/// Per-edit bounds for Changes Information operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Select an optional tighter Changes Information profile.
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

/// A staged Changes Information edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    limits: Limits,
}

impl Package {
    /// Start a Changes Information edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_changes_info(&self) -> Result<Edit<'_>> {
        self.edit_changes_info_with_options(EditOptions::default())
    }

    /// Start a Changes Information edit under explicit operation options.
    pub fn edit_changes_info_with_options(&self, options: EditOptions) -> Result<Edit<'_>> {
        let staged = Staged::new(self, "edit_changes_info")?;
        let limits = options
            .limits_value()
            .unwrap_or_else(|| staged.changes_limits())
            .constrained_by_source(self.opc.read_limits());
        Ok(Edit { staged, limits })
    }

    /// Apply a Changes Information patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_changes_info(&self, patch: &Patch) -> Result<Commit> {
        let semantic_limits = self.semantic_limits();
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_changes_info",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part_policy(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::Changes,
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
    /// Stage one checked [`Draft`] of the presentation's Changes Information part: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut Draft) -> Result<T>) -> Result<T> {
        let limits = self.limits;
        self.staged.update(move |opc| {
            let mut draft = load_snapshot(opc, limits)?
                .ok_or_else(|| {
                    Error::refused(
                        litchi_core::edit::RefusalKind::NotFound,
                        "update",
                        "the package has no Changes Information part",
                    )
                })?
                .edit();
            let value = operation(&mut draft)?;
            package::apply_commit(opc, draft.commit()?)?;
            Ok(value)
        })
    }

    /// Read the unique Changes Information part after validating its OPC graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn changes_info(&self) -> Result<Option<Part>> {
        load(self.staged.opc(), self.limits)
    }

    /// Add a new Changes Information owner atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Part) -> Result<()> {
        let limits = self.limits;
        self.staged
            .update(move |opc| package::store(opc, value, limits))
    }

    /// Remove the Changes Information owner and its Presentation relationship.
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
    pub fn remove(&mut self) -> Result<Option<Part>> {
        let limits = self.limits;
        self.staged.update(move |opc| package::remove(opc, limits))
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
