//! Source-checked, reversible smart-tag property-bag patches.

use super::model::Error;
use super::snapshot::Snapshot;
use super::transaction::Revision;

/// The typed before/after snapshots represented by a patch.
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

    /// Borrows the complete typed state required before the edit.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrows the complete typed state produced by the edit.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }
}

/// A reversible replacement of one complete smart-tag property-bag payload.
///
/// Applying a patch requires both the expected revision and exact source
/// bytes. This prevents a same-length edit from another producer from being
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
        let change = (before.store() != after.store() || before.bags() != after.bags())
            .then(|| Change::new(&before, &after));
        Self {
            base: before.revision(),
            target: after.revision(),
            before,
            after,
            change,
        }
    }

    /// Returns the expected source revision.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Returns the revision produced by this patch.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Returns the expected source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    /// Returns the resulting fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    /// Borrows the exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Alias for [`Self::before_bytes`].
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before_bytes()
    }

    /// Borrows the exact bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Alias for [`Self::after_bytes`].
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after_bytes()
    }

    /// Returns the complete source snapshot retained by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Returns the complete replacement snapshot retained by this patch.
    #[must_use]
    pub const fn replacement(&self) -> &Snapshot {
        &self.after
    }

    /// Returns the typed change, or `None` for a semantic no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&Change> {
        self.change.as_ref()
    }

    /// Whether this patch preserves the source byte-for-byte.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Applies the patch only to the exact source snapshot used to create it.
    ///
    /// # Errors
    ///
    /// Returns an error if `source` does not match this patch's base snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.revision() != self.base || source.bytes() != self.before.bytes() {
            return Err(Error::new(
                "smart-tag patch source does not match its base snapshot",
            ));
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse replacement.
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
