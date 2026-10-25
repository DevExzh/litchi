//! Immutable color-map edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_color_maps`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of the color map or color-map override of one slide master, slide layout or slide part. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_color_maps`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use super::{Draft, Prepared, Snapshot};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::part_edit::{Staged, family};
use crate::slide::Owner;
use crate::{Package, Result};

family!("color-map");

/// A staged color-map edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a color-map edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_color_maps(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_color_maps")?,
        })
    }

    /// Apply a color-map patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_color_maps(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_color_maps",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::ColorMaps,
                        name,
                        part,
                    )
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Stage one checked [`Draft`] of the color map or color-map override of one slide master, slide layout or slide part: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<'o, T>(
        &mut self,
        owner: impl Into<Owner<'o>>,
        operation: impl FnOnce(&mut Draft) -> Result<T>,
    ) -> Result<T> {
        let part_name = &crate::master_layout::resolve_owner(self.staged.opc(), owner.into())?;
        self.staged.update(move |opc| {
            let mut draft = source(opc, part_name)?.edit();
            let value = operation(&mut draft)?;
            publish(opc, part_name, &draft.commit()?)?;
            Ok(value)
        })
    }

    /// The staged color map or color-map override of the slide master, slide
    /// layout or slide `owner`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing part, a part of another type, or
    /// malformed color-map markup.
    pub fn color_map<'o>(&self, owner: impl Into<Owner<'o>>) -> Result<Snapshot> {
        let part_name = &crate::master_layout::resolve_owner(self.staged.opc(), owner.into())?;
        source(self.staged.opc(), part_name)
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

/// Capture the color-map source of one slide master, slide layout or slide.
fn source(opc: &OpcPackage, part_name: &PackURI) -> Result<Snapshot> {
    let part = opc.get_part(part_name)?;
    let xml = part.blob().to_vec();
    let limits = Limits::from_source_limits(opc.read_limits());
    match part.content_type() {
        ct::PML_SLIDE_MASTER => Snapshot::from_master_owned(xml, limits),
        ct::PML_SLIDE_LAYOUT => {
            Snapshot::from_override_owned(xml, b"sldLayout", "slide layout", limits)
        },
        ct::PML_SLIDE => Snapshot::from_override_owned(xml, b"sld", "slide", limits),
        _ => Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "update",
            "a color map lives in a slide master, slide layout or slide part",
        )),
    }
}

/// Replace the owner part's XML with the prepared source; an unchanged
/// draft stages nothing.
fn publish(opc: &mut OpcPackage, part_name: &PackURI, prepared: &Prepared) -> Result<()> {
    if prepared.patch().is_noop() {
        return Ok(());
    }
    refuse_signed_change(opc)?;
    let mut xml = opc.get_part(part_name)?.blob().to_vec();
    prepared
        .patch()
        .apply(&mut xml, Limits::from_source_limits(opc.read_limits()))?;
    opc.get_part_mut(part_name)?.set_blob(xml);
    settle_unsigned_change(opc)
}
