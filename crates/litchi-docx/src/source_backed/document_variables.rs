//! The document-variable family of a source-backed package (record
//! 0355a): the variables of the existing settings part.
//!
//! [`Package::edit_document_variables`](crate::source_backed::Package::edit_document_variables)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the settings part and copies every other ZIP
//! member, and an exact no-op reproduces the source artifact.
//! [`Package::apply_document_variables`](crate::source_backed::Package::apply_document_variables)
//! replays a [`Patch`].

use crate::Result;
use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::variables::{self as record, Variables};

/// The detached settings-part patch the family publishes.
#[derive(Clone)]
pub(crate) struct VariablesChange(record::Patch);

impl Change for VariablesChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        self.0.is_empty()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        let target = package.document_variables_target(&self.0, "commit_document_variables")?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        let (partname, xml) = target;
        package.deferred_part_overlay(partname, xml)
    }
}

source_family!("document-variable", VariablesChange);

impl From<record::Patch> for Patch {
    /// Replay a detached settings-part record patch onto a package: it
    /// applies to a package whose settings part holds its before state.
    fn from(patch: record::Patch) -> Self {
        Self {
            change: VariablesChange(patch),
            restore: None,
        }
    }
}

/// The budgets of this family: the settings part it rewrites.
pub use crate::settings::LimitKind;

/// A staged document-variable change over one source-backed package. The
/// source is never changed: [`Self::commit`] returns a new package, and
/// dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    base: record::Snapshot,
    inner: record::Edit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The staged variables.
    #[must_use]
    pub const fn variables(&self) -> &Variables {
        self.inner.variables()
    }

    /// Insert or replace one variable.
    ///
    /// # Errors
    ///
    /// Returns an error for a name or value Word refuses.
    pub fn set(&mut self, name: impl Into<String>, value: impl Into<String>) -> Result<&mut Self> {
        self.inner.set(name, value)?;
        Ok(self)
    }

    /// Remove one variable and return its previous value.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.inner.remove(name)
    }

    /// Remove every variable.
    pub fn clear(&mut self) -> &mut Self {
        self.inner.clear();
        self
    }

    /// Replace the whole collection.
    ///
    /// # Errors
    ///
    /// Returns an error for a collection Word refuses.
    pub fn replace(&mut self, value: Variables) -> Result<&mut Self> {
        self.inner.replace(value)?;
        Ok(self)
    }

    /// Apply a staged collection mutation atomically.
    ///
    /// # Errors
    ///
    /// Returns the mutation's error, or an error for a result Word refuses.
    pub fn edit_variables(
        &mut self,
        edit: impl FnOnce(&mut Variables) -> Result<()>,
    ) -> Result<&mut Self> {
        self.inner.edit_variables(edit)?;
        Ok(self)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.inner = self.base.edit();
    }

    /// Commit the staged variables as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source-conflict, relationship, protection,
    /// preservation, signature or limit error before any output.
    pub fn commit(self) -> Result<Commit> {
        let patch = self.inner.commit()?.into_patch();
        commit_change(&self.package, VariablesChange(patch))
    }
}

impl Package {
    /// Stage a document-variable change of the existing settings part.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::document_variables_snapshot`].
    pub fn edit_document_variables(&self) -> Result<Edit> {
        let base = self.document_variables_snapshot()?;
        let inner = base.try_edit()?;
        Ok(Edit {
            package: self.clone(),
            base,
            inner,
        })
    }

    /// Replay a document-variable patch onto this package, which must hold
    /// the patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for another settings state,
    /// and the errors of [`Edit::commit`].
    pub fn apply_document_variables(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
