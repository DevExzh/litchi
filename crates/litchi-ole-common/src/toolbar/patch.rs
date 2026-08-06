//! Source-checked, reversible toolbar-control patches.

use std::sync::Arc;

use super::Error;
use super::control::Control;
use super::snapshot::{Revision, Snapshot};

/// The typed before/after control states represented by a patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    before: Control<'static>,
    after: Control<'static>,
}

impl Change {
    pub(crate) fn new(before: &Control<'static>, after: &Control<'static>) -> Self {
        Self {
            before: before.clone(),
            after: after.clone(),
        }
    }

    /// Borrow the typed source control state.
    pub const fn before(&self) -> &Control<'static> {
        &self.before
    }

    /// Borrow the typed target control state.
    pub const fn after(&self) -> &Control<'static> {
        &self.after
    }
}

/// A reversible, source-checked replacement of one complete `TBC` control.
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
        let change = (before.control() != after.control())
            .then(|| Change::new(before.control(), after.control()));
        Self {
            base: before.revision(),
            target: after.revision(),
            before: before.bytes_shared(),
            after: after.bytes_shared(),
            change,
        }
    }

    /// Return the required source revision.
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Return the produced target revision.
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Return the required source fingerprint.
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    /// Return the resulting target fingerprint.
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    /// Borrow the exact source bytes required by this patch.
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }

    /// Alias for [`Self::before_bytes`].
    pub fn before(&self) -> &[u8] {
        self.before_bytes()
    }

    /// Borrow the exact bytes produced by this patch.
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    /// Alias for [`Self::after_bytes`].
    pub fn after(&self) -> &[u8] {
        self.after_bytes()
    }

    /// Return the typed change, or `None` for an exact no-op.
    pub const fn change(&self) -> Option<&Change> {
        self.change.as_ref()
    }

    /// Whether this patch preserves the source byte-for-byte.
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Alias for [`Self::is_noop`].
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Apply this patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.revision() != self.base || source.bytes() != self.before.as_ref() {
            return Err(Error::invalid(
                "toolbar patch source does not match its base snapshot",
            ));
        }
        let Some(change) = &self.change else {
            return Ok(source.clone());
        };
        Snapshot::from_parts(Arc::clone(&self.after), change.after.clone())
    }

    /// Revert this patch only from its exact target snapshot.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot, Error> {
        self.inverse().apply(target)
    }

    /// Build the exact inverse replacement.
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
