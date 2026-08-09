//! Failure-atomic typed link transactions.

use super::{Link, Patch, Snapshot, Times, validation};
use crate::property_set::Guid;
use litchi_cfb::OleError;

/// A deterministic identity for one exact serialized OLEDS link stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(crate) fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
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

/// An isolated typed edit over one source link snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Link,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.link.clone(),
            source,
        }
    }

    /// Borrows the source snapshot used for this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the projected typed link value.
    #[must_use]
    pub const fn link(&self) -> &Link {
        &self.candidate
    }

    /// Whether any typed field differs from the source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.link
    }

    /// Updates the cache hint while retaining every other flags bit.
    pub fn set_cache_hint(&mut self, enabled: bool) -> &mut Self {
        self.candidate.set_cache_hint(enabled);
        self
    }

    /// Replaces the raw flags while preserving the embedded/linked layout.
    ///
    /// # Errors
    ///
    /// Returns an error when `flags` changes the embedded/linked layout.
    pub fn set_flags(&mut self, flags: u32) -> Result<&mut Self, OleError> {
        self.candidate.set_flags(flags)?;
        Ok(self)
    }

    /// Replaces the implementation-specific OLEDS update option.
    pub fn set_link_update_option(&mut self, value: u32) -> &mut Self {
        self.candidate.set_link_update_option(value);
        self
    }

    /// Replaces the linked object's class identifier.
    ///
    /// # Errors
    ///
    /// Returns an error when the link has no class identifier field in its
    /// wire layout.
    pub fn set_class_id(&mut self, value: Guid) -> Result<&mut Self, OleError> {
        self.candidate.set_class_id(value)?;
        Ok(self)
    }

    /// Replaces the linked object's FILETIME group.
    ///
    /// # Errors
    ///
    /// Returns an error when the link has no timestamp group in its wire
    /// layout.
    pub fn set_times(&mut self, value: Times) -> Result<&mut Self, OleError> {
        self.candidate.set_times(value)?;
        Ok(self)
    }

    /// Applies a custom typed edit using the existing inert [`Link`] setters.
    ///
    /// # Errors
    ///
    /// Returns an error when `edit` fails or leaves the candidate with an
    /// invalid OLEDS layout.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut Link) -> Result<(), OleError>,
    {
        let mut candidate = self.candidate.clone();
        edit(&mut candidate)?;
        validation::validate(&candidate)?;
        self.candidate = candidate;
        Ok(self)
    }

    /// Projects and validates the current candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate has an invalid OLEDS layout.
    pub fn snapshot(&self) -> Result<Snapshot, OleError> {
        self.materialize()
    }

    /// Restores the source candidate and returns the immutable source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate without mutating its source.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate has an invalid OLEDS layout.
    pub fn commit(self) -> Result<Commit, OleError> {
        let snapshot = self.materialize()?;
        let patch = self.source.patch_to(&snapshot);
        Ok(Commit { snapshot, patch })
    }

    fn materialize(&self) -> Result<Snapshot, OleError> {
        validation::validate(&self.candidate)?;
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let bytes = self.candidate.to_bytes().into();
        Snapshot::parse_shared(bytes)
    }
}

/// A successful link publication containing a new snapshot and reversible
/// source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether the publication changed any source bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Moves the published snapshot out of this commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Moves the reversible patch out of this commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits this publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Runs one typed link edit and publishes it atomically.
///
/// # Errors
///
/// Returns an error when `edit` fails or leaves the link with an invalid OLEDS
/// layout.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, OleError>
where
    F: FnOnce(&mut Transaction) -> Result<(), OleError>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}
