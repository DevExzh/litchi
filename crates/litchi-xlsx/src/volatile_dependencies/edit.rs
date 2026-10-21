//! Immutable volatile-dependencies edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_volatile_dependencies`] starts an [`Edit`]; every operation
//! runs on a private candidate graph, so an error stages nothing.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_volatile_dependencies`] replays a patch onto an equal package. The
//! source package never changes. A changed edit of a signed workbook is
//! refused with [`Error::Signed`]; an exact no-op keeps every byte.

use super::model::{VolatileDependencies, VolatileDependenciesConformance};
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("volatile-dependencies");

pub use crate::opc::LimitKind;

/// A staged volatile-dependencies edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a volatile-dependencies edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_volatile_dependencies(&self) -> Result<Edit> {
        let staged = Staged::new(self, "edit_volatile_dependencies")?;
        Ok(Edit { staged })
    }

    /// Apply a volatile-dependencies patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_volatile_dependencies(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_volatile_dependencies")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged volatile-dependency records with their namespace family.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the graph.
    pub fn volatile_dependencies(
        &self,
    ) -> Result<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
        super::package::load_from_package_with_conformance(self.staged.opc())
    }

    /// Store or replace the inert records in `conformance`.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a signed
    /// package.
    pub fn set(
        &mut self,
        value: &VolatileDependencies,
        conformance: VolatileDependenciesConformance,
    ) -> Result<()> {
        self.staged
            .update(|candidate| super::package::store_in_package(candidate, value, conformance))
    }

    /// Remove the part and its workbook relationship. Returns whether it
    /// was present.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a change to a
    /// signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.staged.update(super::package::remove_from_package)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed signed package and the
    /// workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
