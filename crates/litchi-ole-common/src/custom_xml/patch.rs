//! Reversible, source-checked Custom XML store patches.

use super::model::Result;
use super::snapshot::{Revision, Snapshot};

/// The typed before/after snapshots represented by one patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    pub(super) fn new(before: &Snapshot, after: &Snapshot) -> Self {
        Self {
            before: before.clone(),
            after: after.clone(),
        }
    }

    /// Borrow the complete source state required before the edit.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the complete replacement state produced by the edit.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }
}

/// A reversible replacement of one complete Custom XML store projection.
///
/// Applying a patch requires both the source fingerprint and the exact typed
/// projection, so a same-fingerprint or same-size unrelated source cannot be
/// mistaken for the patch base.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Snapshot,
    after: Snapshot,
    change: Option<Change>,
}

impl Patch {
    pub(super) fn new(before: Snapshot, after: Snapshot) -> Self {
        let change = (!before.same_source(&after)).then(|| Change::new(&before, &after));
        Self {
            base: before.revision(),
            target: after.revision(),
            before,
            after,
            change,
        }
    }

    /// Return the expected source revision.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Return the produced target revision.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Return the expected source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    /// Return the resulting target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    /// Borrow the complete source snapshot retained by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the complete replacement snapshot retained by this patch.
    #[must_use]
    pub const fn replacement(&self) -> &Snapshot {
        &self.after
    }

    /// Return the typed before/after change, or `None` for an exact no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&Change> {
        self.change.as_ref()
    }

    /// Whether this patch retains the source exactly.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Apply the patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(super::model::invalid(
                "custom XML patch source does not match its base snapshot",
            ));
        }
        Ok(self.after.clone())
    }

    /// Revert the patch only from its exact replacement snapshot.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(target)
    }

    /// Build the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            change: self.change.as_ref().map(|change| Change {
                before: change.after.clone(),
                after: change.before.clone(),
            }),
        }
    }
}
