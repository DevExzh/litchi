//! Immutable slide-show event edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_show_events`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`ListDraft`] of one slide's slide-show event list. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_show_events`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use crate::slide::Selector;

use super::ListDraft;
use super::codec::Limits;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide-show event");

/// Per-edit bounds for inert slide-show event operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// A staged slide-show event edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    limits: Limits,
}

impl Package {
    /// Start a slide-show event edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_show_events(&self) -> Result<Edit<'_>> {
        self.edit_show_events_with_options(EditOptions::default())
    }

    /// Start a slide-show event edit under explicit operation options.
    pub fn edit_show_events_with_options(&self, options: EditOptions) -> Result<Edit<'_>> {
        let staged = Staged::new(self, "edit_show_events")?;
        let limits = options
            .limits_value()
            .unwrap_or_else(|| staged.events_limits())
            .constrained_by_source(self.opc.read_limits());
        Ok(Edit { staged, limits })
    }

    /// Apply a slide-show event patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_show_events(&self, patch: &Patch) -> Result<Commit> {
        let semantic_limits = self.semantic_limits();
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_show_events",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part_policy(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::Events,
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
    /// Stage one checked [`ListDraft`] of one slide's slide-show event list: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<'s, T>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        operation: impl FnOnce(&mut ListDraft) -> Result<T>,
    ) -> Result<T> {
        let slide_part_name = &crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits;
        self.staged.update(move |opc| {
            let mut draft = load_snapshot(opc, slide_part_name, limits)?
                .ok_or_else(|| {
                    Error::refused(
                        litchi_core::edit::RefusalKind::NotFound,
                        "update",
                        "the slide has no slide-show event list",
                    )
                })?
                .edit();
            let value = operation(&mut draft)?;
            package::apply_commit(opc, draft.commit()?)?;
            Ok(value)
        })
    }

    /// Discover the typed event records on one slide without executing them.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn show_events<'s>(&self, slide: impl Into<Selector<'s>>) -> Result<Option<Vec<Event>>> {
        let slide_part_name = &crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        load(self.staged.opc(), slide_part_name, self.limits)
    }

    /// Add a new slide-show event extension atomically.
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
    pub fn set<'s>(&mut self, slide: impl Into<Selector<'s>>, events: &[Draft]) -> Result<()> {
        let slide_part_name = &crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits;
        self.staged
            .update(move |opc| package::store(opc, slide_part_name, events, limits))
    }

    /// Remove only the MS-PPTX show-event extension, retaining neighboring opaque
    /// extension blocks and all other slide XML.
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
    pub fn remove<'s>(&mut self, slide: impl Into<Selector<'s>>) -> Result<Option<Snapshot>> {
        let slide_part_name = &crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits;
        self.staged
            .update(move |opc| package::remove(opc, slide_part_name, limits))
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
