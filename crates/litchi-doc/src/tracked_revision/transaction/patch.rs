//! Source-checked reversible DOC tracked-revision patches.

use super::snapshot::{Snapshot, fingerprint};
use super::{Result, TransactionError};

/// A reversible replacement of a complete validated DOC artifact.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    before_fingerprint: u64,
    after_fingerprint: u64,
}

impl Patch {
    pub(super) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            before_fingerprint: fingerprint(before.bytes()),
            after_fingerprint: fingerprint(after.bytes()),
            before,
            after,
        }
    }

    /// Returns the exact source snapshot required by this patch.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Returns the exact replacement snapshot produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Returns the source fingerprint used as a fast conflict precheck.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.before_fingerprint
    }

    /// Returns the replacement fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.after_fingerprint
    }

    /// Whether the complete DOC artifact is byte-for-byte unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.before_fingerprint || source.bytes() != self.before.bytes()
        {
            return Err(TransactionError::Conflict);
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
        }
    }
}
