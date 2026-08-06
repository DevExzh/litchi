//! Immutable, source-preserving DOC tracked-revision snapshots.

use super::super::{Limits, Revision, RevisionEditor};
use super::{Result, Transaction};
use crate::package::Result as PackageResult;
use std::sync::Arc;

/// An immutable validated snapshot of a complete legacy DOC artifact.
///
/// The snapshot retains the exact source CFB bytes and reparses only the
/// bounded tracked-revision projection when a caller requests it. No-op
/// publication therefore never renders a replacement package, and unknown
/// streams, table blocks, and SPRM bytes remain owned by the source artifact.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
}

impl Snapshot {
    /// Opens and validates an owned DOC artifact with explicit resource limits.
    pub fn open(input: impl Into<Vec<u8>>, limits: Limits) -> PackageResult<Self> {
        let bytes = input.into();
        RevisionEditor::open(bytes.clone(), limits)?;
        Ok(Self {
            source: Arc::from(bytes.into_boxed_slice()),
            limits,
        })
    }

    /// Parses a borrowed DOC artifact using the default resource limits.
    pub fn parse(input: &[u8]) -> PackageResult<Self> {
        Self::open(input.to_vec(), Limits::default())
    }

    /// Parses an owned DOC artifact using the default resource limits.
    pub fn from_bytes(input: Vec<u8>) -> PackageResult<Self> {
        Self::open(input, Limits::default())
    }

    /// Returns the exact source CFB bytes captured by this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Returns shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Returns a stable first-stage fingerprint for stale-source checks.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        fingerprint(&self.source)
    }

    /// Lists the typed tracked revisions in source order.
    pub fn revisions(&self) -> PackageResult<Vec<Revision>> {
        self.editor()?.revisions()
    }

    /// Returns the revision-author table in its stored order.
    pub fn authors(&self) -> PackageResult<Vec<String>> {
        Ok(self.editor()?.authors().to_vec())
    }

    /// Starts an isolated clone-first tracked-revision transaction.
    pub fn edit(&self) -> Result<Transaction> {
        Transaction::new(self.clone())
    }

    /// Alias for [`Self::edit`].
    pub fn transaction(&self) -> Result<Transaction> {
        self.edit()
    }

    /// Returns the exact source bytes. A snapshot itself is always a no-op.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    pub(super) const fn limits(&self) -> Limits {
        self.limits
    }

    fn editor(&self) -> PackageResult<RevisionEditor> {
        RevisionEditor::open(self.source.as_ref().to_vec(), self.limits)
    }
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.len())
            .field("fingerprint", &self.fingerprint())
            .field("limits", &self.limits)
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

pub(super) fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
