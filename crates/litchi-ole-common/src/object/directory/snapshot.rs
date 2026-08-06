//! Immutable, source-preserving CFB directory snapshots.

use super::catalog::{Catalog, Entry};
use super::model::{Limits, Links, Metadata};
use super::transaction::{Revision, Transaction};
use litchi_cfb::{DirectoryEntry, OleError};
use std::ops::Deref;
use std::sync::Arc;

/// An immutable, cheaply clonable directory catalog snapshot.
///
/// The exact source `DirectoryEntry` allocation is retained.  A no-op edit
/// therefore returns the original allocation and leaves all raw or unknown
/// fields untouched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    catalog: Catalog,
    revision: Revision,
}

impl Snapshot {
    /// Parses a bounded directory catalog from borrowed CFB entries.
    pub fn parse(entries: &[DirectoryEntry], limits: Limits) -> Result<Self, OleError> {
        let catalog = Catalog::parse(entries, limits)?;
        Ok(Self::from_catalog(catalog))
    }

    /// Parses a directory catalog with the default bounds.
    pub fn parse_default(entries: &[DirectoryEntry]) -> Result<Self, OleError> {
        Self::parse(entries, Limits::default())
    }

    /// Captures owned directory entries without retaining a CFB container.
    pub fn from_entries(entries: Vec<DirectoryEntry>, limits: Limits) -> Result<Self, OleError> {
        Ok(Self::from_catalog(Catalog::from_entries(entries, limits)?))
    }

    /// Captures an already shared directory-entry allocation without copying.
    pub fn from_entries_shared(
        entries: Arc<[DirectoryEntry]>,
        limits: Limits,
    ) -> Result<Self, OleError> {
        Ok(Self::from_catalog(Catalog::from_entries_shared(
            entries, limits,
        )?))
    }

    /// Publishes a validated catalog as a source-preserving snapshot.
    pub fn from_catalog(catalog: Catalog) -> Self {
        let revision = Revision::of(catalog.raw_entries());
        Self { catalog, revision }
    }

    /// Borrows the typed and raw directory catalog.
    #[must_use]
    pub const fn catalog(&self) -> &Catalog {
        &self.catalog
    }

    /// Returns entries in source order.
    pub fn entries(&self) -> impl Iterator<Item = Entry<'_>> + '_ {
        self.catalog.entries()
    }

    /// Returns one checked metadata value by SID.
    #[must_use]
    pub fn metadata(&self, sid: super::Sid) -> Option<&Metadata> {
        self.catalog.metadata(sid)
    }

    /// Returns one checked containment-link value by SID.
    #[must_use]
    pub fn links(&self, sid: super::Sid) -> Option<Links> {
        self.catalog.links(sid)
    }

    /// The exact raw directory values retained by this snapshot.
    #[must_use]
    pub fn raw_entries(&self) -> &[DirectoryEntry] {
        self.catalog.raw_entries()
    }

    /// Shared ownership of the exact raw directory values.
    #[must_use]
    pub fn raw_entries_shared(&self) -> Arc<[DirectoryEntry]> {
        self.catalog.raw_entries_shared()
    }

    /// The deterministic identity of the exact source directory catalog.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// The compact source fingerprint used by patches.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Resource bounds retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.catalog.limits()
    }

    /// Starts an isolated typed directory edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`] for transactional call sites.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    pub(crate) fn catalog_clone(&self) -> Catalog {
        self.catalog.clone()
    }
}

impl Deref for Snapshot {
    type Target = Catalog;

    fn deref(&self) -> &Self::Target {
        self.catalog()
    }
}
