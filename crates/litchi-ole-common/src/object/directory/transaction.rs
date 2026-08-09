//! Failure-atomic typed edits for a CFB directory catalog.

use super::catalog::Catalog;
use super::codec;
use super::model::{EntryKind, Links, Metadata, Sid};
use super::patch::Patch;
use super::snapshot::Snapshot;
use super::validation;
use crate::property_set::Guid;
use litchi_cfb::{DirectoryEntry, OleError};

/// A deterministic identity for one exact raw directory catalog.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(crate) fn of(entries: &[DirectoryEntry]) -> Self {
        Self(codec::fingerprint(entries))
    }

    /// Returns the raw source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

/// An isolated, failure-atomic edit over one directory snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Catalog,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.catalog_clone(),
            source,
        }
    }

    /// Borrows the immutable source snapshot used by this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the current typed and raw catalog draft.
    #[must_use]
    pub const fn catalog(&self) -> &Catalog {
        &self.candidate
    }

    /// Returns entries in the current draft's source order.
    pub fn entries(&self) -> impl Iterator<Item = super::Entry<'_>> + '_ {
        self.candidate.entries()
    }

    /// Whether the current raw directory catalog differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.candidate.raw_equal(self.source.catalog())
    }

    /// Replaces a producer-visible directory name while retaining all other
    /// raw fields and unknown entry data.
    ///
    /// # Errors
    ///
    /// Returns an error when the SID is absent, the name is invalid, or the
    /// resulting catalog violates its retained resource limits.
    pub fn set_name(&mut self, sid: Sid, name: impl Into<String>) -> Result<&mut Self, OleError> {
        let replacement_name = name.into();
        validation::validate_name_for_edit(&replacement_name, self.source.limits())?;
        let index = self.index(sid)?;
        self.update_raw(move |entries| {
            entries[index].name = replacement_name;
            Ok(())
        })
    }

    /// Replaces containment links without requiring an entry kind to be known.
    ///
    /// # Errors
    ///
    /// Returns an error when the SID is absent, a link is self-referential, or
    /// the resulting catalog violates CFB invariants or resource limits.
    pub fn set_links(&mut self, sid: Sid, links: Links) -> Result<&mut Self, OleError> {
        validate_links(sid, links)?;
        let index = self.index(sid)?;
        self.update_raw(move |entries| {
            entries[index].sid_left = links.left().map_or(super::NOSTREAM, Sid::raw);
            entries[index].sid_right = links.right().map_or(super::NOSTREAM, Sid::raw);
            entries[index].sid_child = links.child().map_or(super::NOSTREAM, Sid::raw);
            Ok(())
        })
    }

    /// Replaces one complete known typed metadata value.
    ///
    /// # Errors
    ///
    /// Returns an error when the SID is absent or opaque, the replacement has
    /// a different SID, or the resulting catalog violates CFB invariants or
    /// resource limits.
    pub fn set_metadata(&mut self, sid: Sid, metadata: Metadata) -> Result<&mut Self, OleError> {
        let index = self.index(sid)?;
        let before = self
            .candidate
            .projection_at(index)
            .and_then(|projection| projection.metadata)
            .ok_or_else(|| opaque_metadata(sid))?;
        if metadata.sid() != sid {
            return Err(OleError::InvalidFormat(
                "CFB directory metadata replacement must retain its SID".into(),
            ));
        }
        validation::validate(metadata)?;
        self.update_raw(move |entries| {
            codec::apply_metadata(&mut entries[index], before, metadata);
            Ok(())
        })
    }

    /// Alias for [`Self::set_metadata`].
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_metadata`].
    pub fn replace_metadata(
        &mut self,
        sid: Sid,
        metadata: Metadata,
    ) -> Result<&mut Self, OleError> {
        self.set_metadata(sid, metadata)
    }

    /// Applies a checked closure to one known metadata projection.
    ///
    /// The candidate is cloned first; a failed closure or invariant check
    /// leaves this transaction unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when the SID is absent or opaque, `edit` fails, or the
    /// resulting metadata violates CFB invariants or resource limits.
    pub fn update_metadata<F>(&mut self, sid: Sid, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut Metadata) -> Result<(), OleError>,
    {
        let index = self.index(sid)?;
        let before = self
            .candidate
            .projection_at(index)
            .and_then(|projection| projection.metadata)
            .ok_or_else(|| opaque_metadata(sid))?;
        let mut after = before;
        edit(&mut after)?;
        self.set_metadata(sid, after)
    }

    /// Changes a storage or root CLSID while preserving the original raw
    /// spelling when the typed value is unchanged.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::update_metadata`].
    pub fn set_class_id(
        &mut self,
        sid: Sid,
        class_id: Option<Guid>,
    ) -> Result<&mut Self, OleError> {
        self.update_metadata(sid, |metadata| {
            metadata.set_class_id(class_id);
            Ok(())
        })
    }

    /// Changes the typed object kind, subject to the normal CFB invariants.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::update_metadata`].
    pub fn set_kind(&mut self, sid: Sid, kind: EntryKind) -> Result<&mut Self, OleError> {
        self.update_metadata(sid, |metadata| {
            metadata.set_kind(kind);
            Ok(())
        })
    }

    /// Changes the starting FAT or `MiniFAT` sector.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::update_metadata`].
    pub fn set_start_sector(&mut self, sid: Sid, value: u32) -> Result<&mut Self, OleError> {
        self.update_metadata(sid, |metadata| {
            metadata.set_start_sector(value);
            Ok(())
        })
    }

    /// Changes the parsed CFB stream size.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::update_metadata`].
    pub fn set_stream_size(&mut self, sid: Sid, value: u64) -> Result<&mut Self, OleError> {
        self.update_metadata(sid, |metadata| {
            metadata.set_stream_size(value);
            Ok(())
        })
    }

    /// Changes `MiniFAT` placement.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::update_metadata`].
    pub fn set_uses_mini_stream(&mut self, sid: Sid, value: bool) -> Result<&mut Self, OleError> {
        self.update_metadata(sid, |metadata| {
            metadata.set_uses_mini_stream(value);
            Ok(())
        })
    }

    /// Applies an inert raw edit and republishes only after catalog validation.
    ///
    /// This is the escape hatch for producer-defined fields exposed by
    /// `litchi_cfb::DirectoryEntry`; it still cannot mutate a CFB container or
    /// bypass SID, name, containment, kind, and resource checks.
    ///
    /// # Errors
    ///
    /// Returns an error when `edit` fails or its changes violate CFB
    /// invariants or the transaction's resource limits.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut [DirectoryEntry]) -> Result<(), OleError>,
    {
        self.update_raw(edit)
    }

    /// Captures the current candidate as a validated snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when a changed candidate no longer validates under the
    /// transaction's retained CFB resource limits.
    pub fn snapshot(&self) -> Result<Snapshot, OleError> {
        self.materialize()
    }

    /// Discards the candidate and returns its source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate without mutating its source.
    ///
    /// # Errors
    ///
    /// Returns an error when a changed candidate no longer validates under the
    /// transaction's retained CFB resource limits.
    pub fn commit(self) -> Result<Commit, OleError> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(&self.source, &snapshot);
        Ok(Commit { snapshot, patch })
    }

    fn update_raw<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut [DirectoryEntry]) -> Result<(), OleError>,
    {
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(self.candidate.raw_entries().len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB directory transaction entries",
                source,
            })?;
        entries.extend(self.candidate.raw_entries().iter().cloned());
        edit(&mut entries)?;
        let candidate = Catalog::from_entries(entries, self.source.limits())?;
        self.candidate = candidate;
        Ok(self)
    }

    fn index(&self, sid: Sid) -> Result<usize, OleError> {
        self.candidate.index_of(sid).ok_or_else(|| {
            OleError::InvalidFormat(format!("CFB directory SID {} is not present", sid.raw()))
        })
    }

    fn materialize(&self) -> Result<Snapshot, OleError> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let catalog = Catalog::from_entries_shared(
            self.candidate.raw_entries_shared(),
            self.source.limits(),
        )?;
        Ok(Snapshot::from_catalog(catalog))
    }
}

/// A successful directory publication containing its reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether the raw directory source changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the committed snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the commit into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Runs one source-checked directory edit atomically.
///
/// # Errors
///
/// Returns an error when `edit` fails or its changes violate CFB invariants or
/// the snapshot's retained resource limits.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, OleError>
where
    F: FnOnce(&mut Transaction) -> Result<(), OleError>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}

fn validate_links(sid: Sid, links: Links) -> Result<(), OleError> {
    if [links.left(), links.right(), links.child()]
        .into_iter()
        .flatten()
        .any(|link| link == sid)
    {
        return Err(OleError::InvalidFormat(
            "CFB directory metadata contains a self-referential link".into(),
        ));
    }
    Ok(())
}

fn opaque_metadata(sid: Sid) -> OleError {
    OleError::InvalidFormat(format!(
        "CFB directory SID {} has no typed metadata projection",
        sid.raw()
    ))
}
