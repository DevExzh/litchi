//! Immutable presentation-structure edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_structure`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of the presentation's custom-show and section graph. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_structure`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use super::super::custom_show::Show;
use super::super::sections::Section;
use super::Draft;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("presentation-structure");

/// A staged presentation-structure edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a presentation-structure edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_structure(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_structure")?,
        })
    }

    /// Apply a presentation-structure patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_structure(&self, patch: &Patch) -> Result<Commit> {
        let package =
            patch
                .0
                .apply_with_replay_validation(self, "apply_structure", |_, target, patch| {
                    patch.validate_changed_parts(target, |name, part| {
                        crate::presentation_properties::metadata::validate_replay_part(
                            target,
                            crate::presentation_properties::metadata::ReplayFamily::Structure,
                            name,
                            part,
                        )
                    })
                })?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Stage one checked [`Draft`] of the presentation's custom-show and section graph: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut Draft) -> Result<T>) -> Result<T> {
        self.staged.update(move |opc| {
            let limits = Limits::from_source_limits(opc.read_limits());
            let mut draft = load_snapshot(opc, limits)?.edit();
            let value = operation(&mut draft)?;
            package::apply_commit(opc, draft.commit()?)?;
            Ok(value)
        })
    }

    /// Load and validate the presentation structure graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn structure(&self) -> Result<Graph> {
        load(
            self.staged.opc(),
            Limits::from_source_limits(self.staged.opc().read_limits()),
        )
    }

    /// Atomically replace custom shows and sections while preserving unrelated XML.
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
    pub fn set(&mut self, value: &Graph) -> Result<()> {
        self.staged.update(move |opc| package::store(opc, value))
    }

    /// The staged custom show `id`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_custom_show(&self, id: u32) -> Result<Option<Show>> {
        find_custom_show(self.staged.opc(), id)
    }

    /// Add one custom show; its ID and name must be new.
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
    pub fn add_custom_show(&mut self, show: Show) -> Result<()> {
        self.staged
            .update(move |opc| package::add_custom_show(opc, show))
    }

    /// Replace the custom show `id` with `replacement`.
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
    pub fn update_custom_show(&mut self, id: u32, replacement: Show) -> Result<()> {
        self.staged
            .update(move |opc| package::update_custom_show(opc, id, replacement))
    }

    /// Remove the custom show `id`. Returns whether it existed.
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
    pub fn remove_custom_show(&mut self, id: u32) -> Result<bool> {
        self.staged
            .update(move |opc| package::remove_custom_show(opc, id))
    }

    /// Reorder the custom shows to `ordered_ids`, which must name each show once.
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
    pub fn reorder_custom_shows(&mut self, ordered_ids: &[u32]) -> Result<()> {
        self.staged
            .update(move |opc| package::reorder_custom_shows(opc, ordered_ids))
    }

    /// Append the slide `slide_id` to the custom show `show_id`.
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
    pub fn add_custom_show_slide(&mut self, show_id: u32, slide_id: u32) -> Result<()> {
        self.staged
            .update(move |opc| package::add_custom_show_slide(opc, show_id, slide_id))
    }

    /// Remove the slide `slide_id` from the custom show `show_id`. Returns whether it was a member.
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
    pub fn remove_custom_show_slide(&mut self, show_id: u32, slide_id: u32) -> Result<bool> {
        self.staged
            .update(move |opc| package::remove_custom_show_slide(opc, show_id, slide_id))
    }

    /// Reorder the slides of the custom show `show_id` to `ordered_slide_ids`.
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
    pub fn reorder_custom_show_slides(
        &mut self,
        show_id: u32,
        ordered_slide_ids: &[u32],
    ) -> Result<()> {
        self.staged
            .update(move |opc| package::reorder_custom_show_slides(opc, show_id, ordered_slide_ids))
    }

    /// The staged section `id`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_section(&self, id: &str) -> Result<Option<Section>> {
        find_section(self.staged.opc(), id)
    }

    /// Add one section, generating its GUID when it has none. Returns the section ID.
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
    pub fn add_section(&mut self, section: Section) -> Result<String> {
        self.staged
            .update(move |opc| package::add_section(opc, section))
    }

    /// Replace the section `id` with `replacement`, keeping slide membership in slide order.
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
    pub fn update_section(&mut self, id: &str, replacement: Section) -> Result<()> {
        self.staged
            .update(move |opc| package::update_section(opc, id, replacement))
    }

    /// Remove the section `id`. Returns whether it existed.
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
    pub fn remove_section(&mut self, id: &str) -> Result<bool> {
        self.staged
            .update(move |opc| package::remove_section(opc, id))
    }

    /// Reorder the sections to `ordered_ids`, which must name each section once.
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
    pub fn reorder_sections(&mut self, ordered_ids: &[String]) -> Result<()> {
        self.staged
            .update(move |opc| package::reorder_sections(opc, ordered_ids))
    }

    /// Add the slide `slide_id` to the section `section_id`.
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
    pub fn add_section_slide(&mut self, section_id: &str, slide_id: u32) -> Result<()> {
        self.staged
            .update(move |opc| package::add_section_slide(opc, section_id, slide_id))
    }

    /// Remove the slide `slide_id` from the section `section_id`. Returns whether it was a member.
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
    pub fn remove_section_slide(&mut self, section_id: &str, slide_id: u32) -> Result<bool> {
        self.staged
            .update(move |opc| package::remove_section_slide(opc, section_id, slide_id))
    }

    /// Reorder the slides of the section `section_id` to `ordered_slide_ids`.
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
    pub fn reorder_section_slides(
        &mut self,
        section_id: &str,
        ordered_slide_ids: &[u32],
    ) -> Result<()> {
        self.staged
            .update(move |opc| package::reorder_section_slides(opc, section_id, ordered_slide_ids))
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
