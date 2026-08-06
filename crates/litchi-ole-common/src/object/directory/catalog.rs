//! Bounded, lossless projections of CFB directory entries.

use super::codec;
use super::model::{EntryKind, Limits, Links, Metadata, Sid};
use super::validation;
use litchi_cfb::{DirectoryEntry, OleError};
use std::sync::Arc;

/// The typed projection retained alongside one raw CFB directory entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Projection {
    pub(crate) metadata: Option<Metadata>,
    pub(crate) links: Links,
}

/// A borrowed catalog entry view.
///
/// Known CFB object kinds expose [`Metadata`].  An unsupported producer entry
/// remains available through [`Self::raw`] and has no typed metadata, so a
/// future entry kind is never silently rewritten by a no-op transaction.
#[derive(Debug, Clone, Copy)]
pub struct Entry<'a> {
    raw: &'a DirectoryEntry,
    metadata: Option<&'a Metadata>,
    links: Links,
}

impl<'a> Entry<'a> {
    /// The checked SID carried by this validated directory entry.
    #[must_use]
    pub const fn sid(self) -> Sid {
        // Catalog construction validates every SID before exposing an entry.
        Sid::from_checked(self.raw.sid)
    }

    /// The original unsigned SID, including an opaque invalid value.
    #[must_use]
    pub const fn raw_sid(self) -> u32 {
        self.raw.sid
    }

    /// The producer-visible directory name.
    #[must_use]
    pub fn name(self) -> &'a str {
        &self.raw.name
    }

    /// The raw CFB entry type byte.
    #[must_use]
    pub const fn raw_kind(self) -> u8 {
        self.raw.entry_type
    }

    /// The known typed object kind, or `None` for an opaque future kind.
    #[must_use]
    pub const fn kind(self) -> Option<EntryKind> {
        match self.raw.entry_type {
            0x01 => Some(EntryKind::Storage),
            0x02 => Some(EntryKind::Stream),
            0x05 => Some(EntryKind::Root),
            _ => None,
        }
    }

    /// Whether the raw entry kind is not understood by this crate.
    #[must_use]
    pub const fn is_unknown(self) -> bool {
        self.kind().is_none()
    }

    /// The typed metadata projection for a known CFB object kind.
    #[must_use]
    pub const fn metadata(self) -> Option<&'a Metadata> {
        self.metadata
    }

    /// The checked containment links, including for an opaque entry kind.
    #[must_use]
    pub const fn links(self) -> Links {
        self.links
    }

    /// The complete raw directory value retained by the catalog.
    #[must_use]
    pub const fn raw(self) -> &'a DirectoryEntry {
        self.raw
    }
}

/// An immutable bounded catalog of raw CFB directory entries and typed views.
///
/// The raw entries are shared between clones.  The catalog does not own CFB
/// sectors, stream payloads, or a writer, and therefore cannot activate or
/// mutate an OLE container.
#[derive(Debug, Clone)]
pub struct Catalog {
    raw: Arc<[DirectoryEntry]>,
    projections: Arc<[Projection]>,
    limits: Limits,
}

impl Catalog {
    /// Parses a directory catalog under explicit resource bounds.
    pub fn parse(entries: &[DirectoryEntry], limits: Limits) -> Result<Self, OleError> {
        let projections = validation::validate_catalog(entries, limits)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(entries.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB directory catalog entries",
                source,
            })?;
        owned.extend(entries.iter().cloned());
        let raw: Arc<[DirectoryEntry]> = owned.into();
        Ok(Self::from_parts(raw, projections, limits))
    }

    /// Parses a directory catalog with the default resource bounds.
    pub fn parse_default(entries: &[DirectoryEntry]) -> Result<Self, OleError> {
        Self::parse(entries, Limits::default())
    }

    /// Captures owned directory entries under explicit resource bounds.
    pub fn from_entries(entries: Vec<DirectoryEntry>, limits: Limits) -> Result<Self, OleError> {
        let projections = validation::validate_catalog(&entries, limits)?;
        let raw: Arc<[DirectoryEntry]> = entries.into();
        Ok(Self::from_parts(raw, projections, limits))
    }

    /// Captures an already shared directory-entry allocation without copying.
    pub fn from_entries_shared(
        entries: Arc<[DirectoryEntry]>,
        limits: Limits,
    ) -> Result<Self, OleError> {
        Self::from_shared(entries, limits)
    }

    pub(crate) fn from_shared(
        raw: Arc<[DirectoryEntry]>,
        limits: Limits,
    ) -> Result<Self, OleError> {
        let projections = validation::validate_catalog(&raw, limits)?;
        Ok(Self::from_parts(raw, projections, limits))
    }

    fn from_parts(
        raw: Arc<[DirectoryEntry]>,
        projections: Vec<Projection>,
        limits: Limits,
    ) -> Self {
        Self {
            raw,
            projections: projections.into(),
            limits,
        }
    }

    /// Returns every entry in source order as a borrowed typed view.
    pub fn entries(&self) -> impl Iterator<Item = Entry<'_>> + '_ {
        self.raw
            .iter()
            .zip(self.projections.iter())
            .map(|(raw, projection)| Entry {
                raw,
                metadata: projection.metadata.as_ref(),
                links: projection.links,
            })
    }

    /// Returns one entry by checked SID.
    #[must_use]
    pub fn get(&self, sid: Sid) -> Option<Entry<'_>> {
        self.index_of(sid).and_then(|index| self.entry_at(index))
    }

    /// Returns one entry by source position.
    #[must_use]
    pub fn at(&self, index: usize) -> Option<Entry<'_>> {
        self.entry_at(index)
    }

    /// Returns the typed metadata for one checked SID.
    #[must_use]
    pub fn metadata(&self, sid: Sid) -> Option<&Metadata> {
        self.get(sid).and_then(Entry::metadata)
    }

    /// Returns the checked containment links for one checked SID.
    #[must_use]
    pub fn links(&self, sid: Sid) -> Option<Links> {
        self.get(sid).map(Entry::links)
    }

    /// The exact raw directory values retained by this catalog.
    #[must_use]
    pub fn raw_entries(&self) -> &[DirectoryEntry] {
        &self.raw
    }

    /// Shared ownership of the exact raw directory values.
    #[must_use]
    pub fn raw_entries_shared(&self) -> Arc<[DirectoryEntry]> {
        Arc::clone(&self.raw)
    }

    /// The number of top-level directory entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.raw.len()
    }

    /// Whether the catalog contains no top-level entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.raw.is_empty()
    }

    /// Resource bounds retained for future transactions.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn index_of(&self, sid: Sid) -> Option<usize> {
        self.raw.iter().position(|entry| entry.sid == sid.raw())
    }

    pub(crate) fn entry_at(&self, index: usize) -> Option<Entry<'_>> {
        let projection = self.projections.get(index)?;
        Some(Entry {
            raw: self.raw.get(index)?,
            metadata: projection.metadata.as_ref(),
            links: projection.links,
        })
    }

    pub(crate) fn projection_at(&self, index: usize) -> Option<Projection> {
        self.projections.get(index).copied()
    }

    pub(crate) fn raw_equal(&self, other: &Self) -> bool {
        codec::raw_catalog_equal(&self.raw, &other.raw)
    }
}

impl PartialEq for Catalog {
    fn eq(&self, other: &Self) -> bool {
        self.raw_equal(other)
    }
}

impl Eq for Catalog {}
