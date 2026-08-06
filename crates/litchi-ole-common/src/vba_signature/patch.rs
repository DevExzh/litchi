//! Source-checked, reversible VBA signature payload patches.

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

    /// Borrows the complete typed state required before the change.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrows the complete typed state produced by the change.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }
}

/// A reversible replacement of one complete VBA signature blob.
///
/// Applying a patch requires both the expected source fingerprint and exact
/// source bytes. This prevents a same-length edit from another producer from
/// being mistaken for the patch base.
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
        let change = (before.info().signature() != after.info().signature()
            || before.info().certificate_store() != after.info().certificate_store())
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

    /// Returns the typed payload change, or `None` for an exact no-op.
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
    /// Returns [`Error`] when the supplied snapshot is stale or belongs to a
    /// different blob kind.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.kind() != self.before.kind()
            || source.revision() != self.base
            || source.bytes() != self.before.bytes()
        {
            return Err(Error::invalid(
                "VBA signature patch source does not match its base snapshot",
            ));
        }
        if self.is_noop() {
            Ok(source.clone())
        } else {
            Ok(self.after.clone())
        }
    }

    /// Applies the exact inverse to the committed target snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the supplied target is stale or does not match
    /// the committed target bytes.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot, Error> {
        self.inverse().apply(target)
    }

    /// Alias for [`Self::revert`].
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the supplied target is stale or does not match
    /// the committed target bytes.
    pub fn undo(&self, target: &Snapshot) -> Result<Snapshot, Error> {
        self.revert(target)
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
