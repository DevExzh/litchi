//! Reversible, source-checked directory catalog patches.

use super::catalog::Catalog;
use super::snapshot::Snapshot;
use super::transaction::Revision;
use litchi_cfb::{DirectoryEntry, OleError};

/// The typed before/after catalog values represented by a patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    before: Catalog,
    after: Catalog,
}

impl Change {
    pub(crate) fn new(before: &Catalog, after: &Catalog) -> Self {
        Self {
            before: before.clone(),
            after: after.clone(),
        }
    }

    /// Borrows the typed catalog required before the edit.
    #[must_use]
    pub const fn before(&self) -> &Catalog {
        &self.before
    }

    /// Borrows the typed catalog produced by the edit.
    #[must_use]
    pub const fn after(&self) -> &Catalog {
        &self.after
    }
}

/// A reversible, source-checked replacement of a complete directory catalog.
///
/// Applying a patch checks both the compact revision and every retained raw
/// directory field.  The catalog layer never serializes or mutates a CFB
/// container; the patch only replaces an inert metadata snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Catalog,
    after: Catalog,
    change: Option<Change>,
}

impl Patch {
    pub(crate) fn new(before: &Snapshot, after: &Snapshot) -> Self {
        let before_catalog = before.catalog_clone();
        let after_catalog = after.catalog_clone();
        let change = (!before_catalog.raw_equal(&after_catalog))
            .then(|| Change::new(&before_catalog, &after_catalog));
        Self {
            base: before.revision(),
            target: after.revision(),
            before: before_catalog,
            after: after_catalog,
            change,
        }
    }

    /// Returns the expected source revision.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Returns the revision produced by this patch.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Returns the expected source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    /// Returns the resulting fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    /// Borrows the exact raw directory values required by this patch.
    #[must_use]
    pub fn before_entries(&self) -> &[DirectoryEntry] {
        self.before.raw_entries()
    }

    /// Borrows the exact raw directory values produced by this patch.
    #[must_use]
    pub fn after_entries(&self) -> &[DirectoryEntry] {
        self.after.raw_entries()
    }

    /// Borrows the typed source catalog represented by this patch.
    #[must_use]
    pub const fn before_catalog(&self) -> &Catalog {
        &self.before
    }

    /// Borrows the typed target catalog represented by this patch.
    #[must_use]
    pub const fn after_catalog(&self) -> &Catalog {
        &self.after
    }

    /// Alias for [`Self::before_entries`].
    #[must_use]
    pub fn before(&self) -> &[DirectoryEntry] {
        self.before_entries()
    }

    /// Alias for [`Self::after_entries`].
    #[must_use]
    pub fn after(&self) -> &[DirectoryEntry] {
        self.after_entries()
    }

    /// Returns the typed change, or `None` for an exact raw no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&Change> {
        self.change.as_ref()
    }

    /// Whether this patch preserves every raw source field.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.raw_equal(&self.after)
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Applies the patch only to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` does not match the patch's exact base
    /// catalog and revision.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, OleError> {
        if source.revision() != self.base || !source.catalog().raw_equal(&self.before) {
            return Err(OleError::InvalidFormat(
                "CFB directory patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Ok(Snapshot::from_catalog(self.after.clone()))
    }

    /// Applies the inverse to the exact committed target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `target` does not match the patch's exact result
    /// catalog and revision.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot, OleError> {
        self.inverse().apply(target)
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            change: self
                .change
                .as_ref()
                .map(|change| Change::new(&change.after, &change.before)),
        }
    }
}
