//! Immutable external-link snapshots and reversible stream patches.

use std::fmt;
use std::sync::Arc;

use crate::{Error, Result};

use super::super::package::Package;

/// An immutable, validated BIFF8 workbook-global stream containing external
/// link metadata.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    package: Package,
    fingerprint: u64,
}

impl Snapshot {
    /// Parses a BIFF8 workbook-global stream without resolving any target.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let bytes = Arc::<[u8]>::from(bytes.as_ref().to_vec().into_boxed_slice());
        Self::parse_shared(bytes)
    }

    /// Parses an already shared source allocation without another byte copy.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self> {
        let package = Package::parse(&bytes)?;
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            package,
        })
    }

    /// Returns the typed, contextual external-link view.
    #[must_use]
    pub fn links(&self) -> &crate::external_link::Links {
        self.package.links()
    }

    /// Alias for [`Self::links`] for callers using the workbook-global name.
    #[must_use]
    pub fn external_links(&self) -> &crate::external_link::Links {
        self.links()
    }

    /// Returns the exact source stream, including unknown records and
    /// producer-specific continuation payloads.
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

    /// Starts a detached failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Publishes the exact validated source stream.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }

    pub(super) fn package(&self) -> &Package {
        &self.package
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.bytes.len())
            .field("supporting_books", &self.links().supporting_books().len())
            .field("external_names", &self.links().external_names().len())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
    }
}

impl Eq for Snapshot {}

/// A source-checked replacement of one complete BIFF8 workbook-global
/// stream.
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

    /// Returns the exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact bytes produced by this patch.
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
                "external-link patch source does not match its base snapshot".into(),
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
                "external-link patch target does not match its committed snapshot".into(),
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
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Whether publication changed any stream byte.
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

    /// Consumes the publication into exact BIFF8 stream bytes.
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
