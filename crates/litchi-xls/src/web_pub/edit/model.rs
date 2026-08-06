//! Immutable snapshots, commits, and reversible patches for `WebPub`.

use std::fmt;
use std::sync::Arc;

use crate::{Error, Result};

use super::super::model::WebPub;

/// An immutable, source-preserving BIFF8 `WebPub` payload.
///
/// The semantic publication is kept beside the exact source bytes. Opening a
/// snapshot never resolves or fetches its URL/path metadata, and publishing a
/// no-op returns the original bytes byte-for-byte.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    publication: WebPub,
    fingerprint: u64,
}

impl Snapshot {
    /// Parses and retains one complete `WebPub` record payload.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let bytes = Arc::<[u8]>::from(bytes.as_ref().to_vec().into_boxed_slice());
        Self::parse_shared(bytes)
    }

    /// Parses an already shared payload without another byte copy.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self> {
        let publication = WebPub::parse(&bytes)?;
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            publication,
        })
    }

    /// Returns the typed, inert web-publication view.
    #[must_use]
    pub const fn publication(&self) -> &WebPub {
        &self.publication
    }

    /// Alias for callers using the BIFF record name.
    #[must_use]
    pub const fn web_pub(&self) -> &WebPub {
        self.publication()
    }

    /// Alias for callers using the semantic feature name.
    #[must_use]
    pub const fn web_publication(&self) -> &WebPub {
        self.publication()
    }

    /// Returns the exact source payload, including opaque reserved bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// Returns the compact source identity used by stale-source checks.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.fingerprint
    }

    /// Starts a detached, failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Publishes the exact validated source payload.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.bytes.len())
            .field("source", &self.publication.source)
            .field("page_type", &self.publication.page_type)
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
    }
}

impl Eq for Snapshot {}

/// A source-checked, reversible replacement of one complete `WebPub` payload.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    pub(super) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            source_fingerprint: before.fingerprint(),
            target_fingerprint: after.fingerprint(),
            before: before.bytes_shared(),
            after: after.bytes_shared(),
        }
    }

    /// Returns the source fingerprint required by this patch.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Returns the target fingerprint produced by this patch.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Returns the exact source payload required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact payload produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether this patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.source_fingerprint || source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "WebPub patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_noop() {
            Ok(source.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.after))
        }
    }

    /// Applies the exact inverse replacement to its committed target.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.fingerprint() != self.target_fingerprint || target.bytes() != self.after() {
            return Err(Error::UnsafeEdit(
                "WebPub patch target does not match its committed snapshot".into(),
            ));
        }
        if self.is_noop() {
            Ok(target.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.before))
        }
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// A successful publication containing the typed result and reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(super) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Whether publication changed any payload byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Returns the post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the publication into its post-edit snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the publication into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consumes the publication into exact BIFF8 payload bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.finish()
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
