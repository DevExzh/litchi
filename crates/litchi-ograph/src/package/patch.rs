//! Source-checked whole-package patches for Graph chart metadata edits.

use std::sync::Arc;

use crate::{Error, Result, chart};

use super::Snapshot;

/// Published result of a successful Graph package transaction.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(super) fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether the complete compound-file artifact differs from its source.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the validated post-edit package snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked reversible package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its exact output artifact bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.into_bytes()
    }

    /// Consume the commit into its output snapshot and reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, source-bound whole-package patch.
#[derive(Debug, Clone)]
pub struct Patch {
    source: Arc<[u8]>,
    replacement: Arc<[u8]>,
    chart: chart::Patch,
}

impl Patch {
    pub(super) fn new(source: Arc<[u8]>, replacement: Arc<[u8]>, chart: chart::Patch) -> Self {
        Self {
            source,
            replacement,
            chart,
        }
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }

    /// Exact replacement bytes produced by this patch.
    #[must_use]
    pub fn replacement(&self) -> &[u8] {
        &self.replacement
    }

    /// The typed chart patch nested inside this package publication.
    #[must_use]
    pub const fn chart(&self) -> &chart::Patch {
        &self.chart
    }

    /// Whether the typed package edit was a semantic no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source.as_ref() == self.replacement.as_ref() && self.chart.is_empty()
    }

    /// Returns the source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.replacement),
            replacement: Arc::clone(&self.source),
            chart: self.chart.inverse(),
        }
    }

    /// Applies this patch only to the exact source snapshot it captured.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes.as_ref() != self.source.as_ref() {
            return Err(Error::UnsupportedMutation {
                operation: "package-patch",
                reason: "patch source package does not match the target snapshot",
            });
        }
        Snapshot::from_bytes_with_limits(self.replacement.to_vec(), source.limits())
    }

    /// Reverts this patch only from its exact replacement snapshot.
    pub fn revert(&self, replacement: &Snapshot) -> Result<Snapshot> {
        if replacement.bytes.as_ref() != self.replacement.as_ref() {
            return Err(Error::UnsupportedMutation {
                operation: "package-patch",
                reason: "patch replacement package does not match the target snapshot",
            });
        }
        Snapshot::from_bytes_with_limits(self.source.to_vec(), replacement.limits())
    }
}
