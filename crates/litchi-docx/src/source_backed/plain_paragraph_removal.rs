//! The plain-paragraph removal family of a source-backed package (record
//! 0355a): remove one plain main-story paragraph.
//!
//! [`Package::edit_plain_paragraph_removal`](crate::source_backed::Package::edit_plain_paragraph_removal)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the main-document part.

use litchi_core::Position;

use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::paragraph_copy::{self as record, RemovalEffectReport as EffectReport};

/// The record patch the family publishes.
#[derive(Clone)]
pub(crate) struct RemovalChange(record::RemovalPatch);

impl Change for RemovalChange {
    type Error = record::Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package, record::Error> {
        let (main, xml) = package.plain_paragraph_removal_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        let base = package.clone();
        Ok(package.deferred(move |writer| base.write_part_overlay(writer, &main, &xml))?)
    }
}

source_family!("plain-paragraph removal", RemovalChange, record::Error);

impl From<record::RemovalPatch> for Patch {
    /// Replay a record patch onto a package whose main document holds its
    /// before bytes.
    fn from(patch: record::RemovalPatch) -> Self {
        Self {
            change: RemovalChange(patch),
            restore: None,
        }
    }
}

/// The budgets of this family, reported by [`record::Error::Limit`].
pub use crate::source_backed::paragraph_copy::LimitKind;

/// A staged removal of one plain paragraph. The source is never changed:
/// [`Self::commit`] returns a new package.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    inner: record::RemovalEdit,
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

    /// The projected main story.
    #[must_use]
    pub const fn projected(&self) -> &record::Snapshot {
        self.inner.projected()
    }

    /// Remove the direct paragraph at `position`, in source order.
    ///
    /// # Errors
    ///
    /// Returns a bounds, refusal or limit error, staging nothing.
    pub fn remove_plain_paragraph(&mut self, position: Position) -> Result<&mut Self, Error> {
        self.inner.remove_plain_paragraph(position)?;
        Ok(self)
    }

    /// The effect the commit will have.
    #[must_use]
    pub fn effect_report(&self) -> EffectReport {
        self.inner.clone().commit().effect_report()
    }

    /// Commit the removal as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source, signature or limit error before any output.
    pub fn commit(self) -> Result<Commit, Error> {
        let patch = self.inner.commit().patch().clone();
        commit_change(&self.package, RemovalChange(patch))
    }
}

impl Package {
    /// Stage a plain-paragraph removal from the main story.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::plain_paragraph_removal_snapshot`].
    pub fn edit_plain_paragraph_removal(&self) -> Result<Edit, Error> {
        Ok(Edit {
            package: self.clone(),
            inner: self.plain_paragraph_removal_snapshot()?.removal_edit(),
        })
    }

    /// Replay a plain-paragraph removal patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns [`record::Error::PatchConflict`] for another main story.
    pub fn apply_plain_paragraph_removal(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_change(self, patch)
    }
}
