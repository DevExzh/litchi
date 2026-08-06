//! Source-checked reversible DOC embedded-object patches.

use super::TransactionError;
use super::snapshot::{Snapshot, fingerprint};
use std::sync::Arc;

/// A complete, source-checked replacement of one DOC artifact.
///
/// The patch boundary is the whole compound file because DOC field CPs,
/// FIB/table pointers, CFB allocation, `ObjectPool` ownership, unknown streams,
/// and opaque embedded payloads form one dependency closure. Applying the
/// patch never executes or resolves any embedded content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    before_fingerprint: u64,
    after_fingerprint: u64,
}

impl Patch {
    pub(super) fn new(before: Vec<u8>, after: Vec<u8>) -> Self {
        Self {
            before_fingerprint: fingerprint(&before),
            after_fingerprint: fingerprint(&after),
            before: Arc::from(before),
            after: Arc::from(after),
        }
    }

    /// Returns the exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact result bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Returns the source fingerprint used as a fast conflict precheck.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.before_fingerprint
    }

    /// Returns the result fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.after_fingerprint
    }

    /// Whether the complete DOC artifact is byte-for-byte unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies the patch only to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns a conflict when the source snapshot bytes do not exactly match
    /// the patch base, or a validation error when the replacement is invalid.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, TransactionError> {
        if source.fingerprint() != self.before_fingerprint || source.bytes() != self.before.as_ref()
        {
            return Err(TransactionError::Conflict);
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::open(self.after.as_ref().to_vec(), source.limits())
            .map_err(TransactionError::Invalid)
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
        }
    }
}

/// A validated post-edit snapshot paired with its reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(super) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Whether publication changed any DOC artifact byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the validated post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its post-edit snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the commit into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits the commit into its result snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
