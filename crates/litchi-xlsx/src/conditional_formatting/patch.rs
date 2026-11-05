//! Exact reversible worksheet conditional-formatting patches.

use super::Snapshot;

/// Exact-source-checked worksheet owner replacement.
#[derive(Clone, Debug)]
pub(crate) struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    /// Exact source state required for application.
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    /// Exact state produced by successful application.
    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    #[must_use]
    /// Whether application preserves the selected source byte-for-byte.
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    /// Construct a patch that restores this patch's exact source state.
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}
