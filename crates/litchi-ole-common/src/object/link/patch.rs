//! Source-checked, reversible link patches.

use super::{Link, Revision, Snapshot};
use litchi_cfb::OleError;
use std::sync::Arc;

/// The typed before/after values represented by a link patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    before: Link,
    after: Link,
}

impl Change {
    pub(crate) fn new(before: &Link, after: &Link) -> Self {
        Self {
            before: before.clone(),
            after: after.clone(),
        }
    }

    /// Borrows the typed value required before the change.
    #[must_use]
    pub const fn before(&self) -> &Link {
        &self.before
    }

    /// Borrows the typed value produced by the change.
    #[must_use]
    pub const fn after(&self) -> &Link {
        &self.after
    }
}

/// A reversible, source-checked replacement of one complete OLEDS link
/// stream.
///
/// The patch retains both exact byte snapshots and typed before/after values.
/// Applying it requires the expected revision and exact source bytes, so a
/// same-length edit from another producer cannot be mistaken for its base.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    change: Option<Change>,
}

impl Patch {
    pub(crate) fn new(before: &Snapshot, after: &Snapshot) -> Self {
        let change = (before.link != after.link).then(|| Change::new(&before.link, &after.link));
        Self {
            base: before.revision,
            target: after.revision,
            before: before.link.bytes_shared(),
            after: after.link.bytes_shared(),
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
        &self.before
    }

    /// Alias for [`Self::before_bytes`].
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before_bytes()
    }

    /// Borrows the exact bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    /// Alias for [`Self::after_bytes`].
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after_bytes()
    }

    /// Returns the typed change, or `None` for an exact no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&Change> {
        self.change.as_ref()
    }

    /// Whether this patch preserves the source byte-for-byte.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
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
    /// Returns an error when `source` does not match the patch's exact base
    /// bytes and revision, or when its target bytes no longer parse as OLEDS
    /// link metadata.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, OleError> {
        if source.revision != self.base || source.bytes() != self.before.as_ref() {
            return Err(OleError::InvalidFormat(
                "OLE link patch source does not match its base snapshot".into(),
            ));
        }
        Snapshot::parse_shared(Arc::clone(&self.after))
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            change: self.change.as_ref().map(|change| Change {
                before: change.after.clone(),
                after: change.before.clone(),
            }),
        }
    }
}
