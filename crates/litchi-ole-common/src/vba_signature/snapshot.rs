//! Immutable, source-preserving VBA signature edit snapshots.

use std::sync::Arc;

use super::model::{Blob, Error, Info, Kind, Limits};
use super::transaction::{Revision, Transaction};

/// An immutable, cheaply clonable owner snapshot for one complete
/// `[MS-OSHARED]` `DigSigBlob` or `WordSigBlob`.
///
/// The parsed [`Blob`] remains the single typed projection of the wire
/// layout. This layer adds the source identity and retained limits needed for
/// failure-atomic, source-checked payload edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    blob: Blob,
    revision: Revision,
}

impl Snapshot {
    /// Parses a complete blob using conservative default limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the source is malformed or exceeds the default
    /// resource limits.
    pub fn parse(source: &[u8], kind: Kind) -> Result<Self, Error> {
        Ok(Self::from_blob(Blob::parse(source, kind)?))
    }

    /// Parses a complete blob using caller-provided resource limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the source violates the wire invariants or the
    /// supplied resource limits.
    pub fn parse_with(source: &[u8], kind: Kind, limits: Limits) -> Result<Self, Error> {
        Ok(Self::from_blob(Blob::parse_with(source, kind, limits)?))
    }

    /// Alias for [`Self::parse_with`] using explicit-limit terminology.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::parse_with`].
    pub fn parse_with_limits(source: &[u8], kind: Kind, limits: Limits) -> Result<Self, Error> {
        Self::parse_with(source, kind, limits)
    }

    /// Parses an already-shared source allocation without copying it.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the source violates the wire invariants or the
    /// supplied resource limits.
    pub fn parse_shared(source: Arc<[u8]>, kind: Kind, limits: Limits) -> Result<Self, Error> {
        Ok(Self::from_blob(Blob::parse_shared(source, kind, limits)?))
    }

    /// Parses a complete `DigSigBlob` using default limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the source is malformed or exceeds the default
    /// resource limits.
    pub fn parse_property(source: &[u8]) -> Result<Self, Error> {
        Self::parse(source, Kind::Property)
    }

    /// Parses a complete `WordSigBlob` using default limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the source is malformed or exceeds the default
    /// resource limits.
    pub fn parse_word(source: &[u8]) -> Result<Self, Error> {
        Self::parse(source, Kind::Word)
    }

    /// Captures an already validated immutable [`Blob`] without reparsing it.
    #[must_use]
    pub fn from_blob(blob: Blob) -> Self {
        let revision = Revision::of(blob.bytes());
        Self { blob, revision }
    }

    /// Returns the parsed immutable blob owner.
    #[must_use]
    pub const fn blob(&self) -> &Blob {
        &self.blob
    }

    /// Returns the source container kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.blob.kind()
    }

    /// Borrows the typed nested signature information.
    #[must_use]
    pub const fn info(&self) -> Info<'_> {
        self.blob.info()
    }

    /// Returns the exact source bytes, including unknown gaps and padding.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.blob.bytes()
    }

    /// Clones ownership of the exact source allocation without copying bytes.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        self.blob.bytes_shared()
    }

    /// Moves the exact source allocation out of this snapshot.
    #[must_use]
    pub fn into_bytes(self) -> Arc<[u8]> {
        self.blob.into_bytes()
    }

    /// Returns the compact source identity used for stale-source checks.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Returns the compact source fingerprint.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Returns the resource limits retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.blob.limits()
    }

    /// Starts an isolated source-checked payload transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`] for callers using transactional terminology.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    /// Consumes this snapshot back into its immutable [`Blob`] owner.
    #[must_use]
    pub fn into_blob(self) -> Blob {
        self.blob
    }

    pub(super) fn layout(&self) -> &super::codec::Layout {
        self.blob.layout()
    }
}
