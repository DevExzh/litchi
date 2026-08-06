//! Typed snapshots and reversible changes for one `Chart` record.

use litchi_ograph::chart::Rect;

use crate::{Error, Result};

use super::validation;

/// A source snapshot of the fixed-size `[MS-XLS]`/`[MS-OGRAPH]` `Chart`
/// geometry.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Snapshot {
    rect: Rect,
}

impl Snapshot {
    /// Creates a semantically valid snapshot for a new edit.
    pub fn try_new(rect: Rect) -> Result<Self> {
        validation::ensure(rect)?;
        Ok(Self { rect })
    }

    /// Returns the decoded fixed-point rectangle.
    pub const fn rect(self) -> Rect {
        self.rect
    }

    /// Opens a source-checked transaction over this snapshot.
    pub fn edit(self) -> Result<super::Transaction> {
        super::Transaction::new(self)
    }

    pub(crate) const fn from_wire(rect: Rect) -> Self {
        Self { rect }
    }
}

impl Default for Snapshot {
    fn default() -> Self {
        Self::from_wire(Rect::default())
    }
}

/// One source-checked replacement of the fixed chart-area payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    pub(crate) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Rectangle required in the source snapshot.
    pub const fn before(self) -> Rect {
        self.before.rect
    }

    /// Rectangle produced by the change.
    pub const fn after(self) -> Rect {
        self.after.rect
    }

    pub(crate) const fn before_snapshot(self) -> Snapshot {
        self.before
    }

    pub(crate) const fn after_snapshot(self) -> Snapshot {
        self.after
    }

    pub(crate) const fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }
}

/// Reversible, source-checked chart-area patch.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Patch {
    change: Option<Change>,
}

impl Patch {
    pub(crate) const fn new(change: Option<Change>) -> Self {
        Self { change }
    }

    /// Returns the effective change, or `None` for a semantic no-op.
    pub const fn change(self) -> Option<Change> {
        self.change
    }

    /// Whether this patch changes no chart-area bytes.
    pub const fn is_empty(self) -> bool {
        self.change.is_none()
    }

    /// Returns the exact inverse change.
    pub const fn inverse(self) -> Self {
        let change = match self.change {
            Some(change) => Some(change.inverse()),
            None => None,
        };
        Self { change }
    }

    /// Applies the patch only to its matching source snapshot.
    pub fn apply(self, source: Snapshot) -> Result<Commit> {
        let Some(change) = self.change else {
            return Ok(Commit::new(source, self));
        };
        if source != change.before_snapshot() {
            return Err(Error::UnsafeEdit(
                "chart-area patch source does not match the snapshot".into(),
            ));
        }
        Ok(Commit::new(change.after_snapshot(), self))
    }
}

/// Published result of a chart-area transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) const fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Returns the post-edit snapshot.
    pub const fn snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Returns the reversible semantic patch.
    pub const fn patch(self) -> Patch {
        self.patch
    }
}
