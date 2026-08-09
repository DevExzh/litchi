//! Immutable snapshots, commits, and reversible byte-preserving patches.

use std::fmt;
use std::sync::Arc;

use crate::{Error, Result};

use super::super::model::{Record, UnknownRecord};
use super::super::package::Package;

/// An immutable, validated workbook-global BIFF stream containing inert RTD
/// metadata and its surrounding opaque records.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    package: Package,
    fingerprint: u64,
}

impl Snapshot {
    /// Parse and retain one complete BIFF8 record stream.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let bytes = Arc::<[u8]>::from(bytes.as_ref().to_vec().into_boxed_slice());
        Self::parse_shared(bytes)
    }

    /// Parse an already shared source allocation without another byte copy.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self> {
        let package = Package::parse(&bytes)?;
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            package,
        })
    }

    /// Alias for [`Self::parse`] in stream-oriented callers.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn read(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse(bytes)
    }

    /// Borrow the typed RTD records in source order.
    #[must_use]
    pub fn records(&self) -> &[Record] {
        self.package.real_time_data()
    }

    /// Contextual alias for [`Self::records`].
    #[must_use]
    pub fn real_time_data(&self) -> &[Record] {
        self.records()
    }

    /// Contextual alias for callers that use the topic terminology.
    #[must_use]
    pub fn topics(&self) -> &[Record] {
        self.records()
    }

    /// Borrow unknown BIFF records in their original source order.
    #[must_use]
    pub fn unknown_records(&self) -> impl Iterator<Item = UnknownRecord<'_>> + '_ {
        self.package.unknown_records(&self.bytes)
    }

    /// Return the exact source stream, including opaque records and producer
    /// padding.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Contextual alias for [`Self::bytes`].
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.bytes()
    }

    /// Share the exact source allocation without copying it.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// The compact identity required by stale-source checks.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.fingerprint
    }

    /// Start a detached, failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Publish the exact validated source stream.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }

    /// This snapshot is always bound to its exact source stream.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        true
    }

    pub(crate) fn package(&self) -> &Package {
        &self.package
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.bytes.len())
            .field("real_time_data", &self.records().len())
            .field("unknown_records", &self.package.unknown_record_count())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
    }
}

impl Eq for Snapshot {}

/// A source-checked, reversible replacement of one complete BIFF stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            source_fingerprint: before.fingerprint(),
            target_fingerprint: after.fingerprint(),
            before: before.bytes_shared(),
            after: after.bytes_shared(),
        }
    }

    /// The source identity required by this patch.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// The target identity produced by this patch.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// The exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// The exact bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether this patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Apply this patch only to its exact source snapshot.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.source_fingerprint || source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "RealTimeData patch source does not match its base snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(source.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.after))
        }
    }

    /// Apply the exact inverse replacement to its committed target.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.fingerprint() != self.target_fingerprint || target.bytes() != self.after() {
            return Err(Error::UnsafeEdit(
                "RealTimeData patch target does not match its committed snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(target.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.before))
        }
    }

    /// Return the exact inverse patch.
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

/// A successful transaction publication containing its resulting snapshot and
/// reversible source-checked patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Whether publication changed any stream byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrow the post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the publication into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the publication into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consume the publication into exact BIFF stream bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.finish()
    }

    /// Consume the publication into both resulting artifacts.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
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
