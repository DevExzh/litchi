//! Immutable read-only-recommendation edits of a [`Package`] (record 0351
//! §2.3).
//!
//! [`Package::edit_read_only_recommended`] starts an [`Edit`]: [`Edit::set`]
//! writes the inert `p1710:readonlyRecommended` hint and [`Edit::remove`]
//! removes it while retaining its owner part. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_read_only_recommended`] replays a patch
//! onto an equal package. The recommendation is document data only: it never
//! makes a package read-only. An exact no-op keeps every byte, and a changed
//! edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).
//!
//! ```
//! # fn main() -> litchi_pptx::Result<()> {
//! let package = litchi_pptx::Package::new()?;
//! let mut edit = package.edit_read_only_recommended()?;
//! assert_eq!(edit.set(true)?, None);
//! let after = edit.commit()?.into_package();
//! assert_eq!(after.read_only_recommended()?, Some(true));
//! assert_eq!(package.read_only_recommended()?, None);
//! # Ok(())
//! # }
//! ```

use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("read-only-recommendation");

/// A staged read-only-recommendation edit of one [`Package`]. Dropping it
/// discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a read-only-recommendation edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_read_only_recommended(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_read_only_recommended")?,
        })
    }

    /// Apply a read-only-recommendation patch to a new package; this one is
    /// unchanged. An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_read_only_recommended(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_read_only_recommended",
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
    /// The staged recommendation; `None` when it is absent.
    ///
    /// # Errors
    ///
    /// Returns an error if the presentation-properties owner is malformed or
    /// ambiguous.
    pub fn value(&self) -> Result<Option<bool>> {
        super::load(self.staged.opc())
    }

    /// Set the recommendation. Returns the previous value.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner is malformed or ambiguous, or for a
    /// changed edit of a signed package.
    pub fn set(&mut self, value: bool) -> Result<Option<bool>> {
        self.staged.update(move |opc| super::put(opc, value))
    }

    /// Remove the recommendation while retaining its owner part. Returns the
    /// removed value.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn remove(&mut self) -> Result<Option<bool>> {
        self.staged.update(super::remove)
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
