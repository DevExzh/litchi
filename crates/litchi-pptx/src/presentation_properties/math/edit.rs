//! Immutable document math-defaults edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_math_properties`] starts an [`Edit`] over the package; [`Edit::set`] replaces the document-level math defaults in the
//! presentation properties and [`Edit::remove`] removes them.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_math_properties`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::Properties;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("document math-defaults");

/// A staged document math-defaults edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a document math-defaults edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_math_properties(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_math_properties")?,
        })
    }

    /// Apply a document math-defaults patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_math_properties(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_math_properties",
            |_, target, patch| {
                let source_limits = target.read_limits();
                let limits =
                    crate::presentation_properties::Limits::from_source_limits(source_limits);
                patch.validate_changed_parts(target, |_, part| {
                    crate::presentation_properties::validate_replay_part(
                        part,
                        limits,
                        source_limits,
                    )
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Replace the document-level math defaults. Returns the previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed presentation-properties part, or a changed edit of a signed package.
    pub fn set(&mut self, value: Properties) -> Result<Option<Properties>> {
        let limits = crate::presentation_properties::Limits::from_source_limits(
            self.staged.opc().read_limits(),
        );
        self.staged.update(move |opc| {
            crate::presentation_properties::put_math_to_package(opc, value, limits)
        })
    }

    /// Remove the document-level math defaults. Returns the removed value.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed presentation-properties part, or a changed edit of a signed package.
    pub fn remove(&mut self) -> Result<Option<Properties>> {
        let limits = crate::presentation_properties::Limits::from_source_limits(
            self.staged.opc().read_limits(),
        );
        self.staged.update(move |opc| {
            crate::presentation_properties::remove_math_from_package(opc, limits)
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
