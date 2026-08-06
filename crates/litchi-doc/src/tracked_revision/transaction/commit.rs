//! Validated tracked-revision transaction results.

use super::{Patch, Snapshot};

/// A validated post-edit snapshot paired with its reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    pub(super) snapshot: Snapshot,
    pub(super) patch: Patch,
}

impl Commit {
    /// Whether publication changed any DOC artifact byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Alias for [`Self::changed`] emphasizing exact no-op publication.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed()
    }

    /// Whether publication retained the exact source artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        !self.changed()
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

    /// Splits the commit into its post-edit snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
