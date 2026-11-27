//! Exact reversible worksheet page-setup patches.

use super::Snapshot;

/// Exact source-checked page-setup worksheet replacement.
#[derive(Clone, Debug)]
pub(crate) struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required source state.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch preserves its source byte-for-byte.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}
