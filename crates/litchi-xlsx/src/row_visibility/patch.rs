//! Exact reversible row-visibility worksheet patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::cell_values;
use crate::error::Result;

/// Exact source-bound replacement of direct existing-row visibility.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    inner: cell_values::Patch,
}

impl Patch {
    pub(super) fn new(before: Snapshot, after: Snapshot) -> Self {
        let inner = cell_values::Patch::new(before.inner().clone(), after.inner().clone());
        Self {
            before,
            after,
            inner,
        }
    }

    /// Required exact source state.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact target state.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether application preserves the selected source bytes exactly.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Exact source-bound inverse, including the producer's original lexical
    /// representation of the `hidden` attribute.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            inner: self.inner.inverse(),
        }
    }

    /// Apply atomically after checking the complete retained package closure.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.inner.apply(package)
    }
}

/// Successful row-visibility publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    pub(super) const fn new(snapshot: Snapshot, patch: Patch, changed_rows: usize) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: Diagnostics { changed_rows },
        }
    }

    /// Whether at least one direct row owner changed lexically.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.diagnostics.changed_rows != 0
    }

    /// Resulting immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free transaction diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Content-free row-visibility publication diagnostics.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_rows: usize,
}

impl Diagnostics {
    /// Number of existing row owners whose `hidden` attribute changed.
    #[must_use]
    pub const fn changed_rows(self) -> usize {
        self.changed_rows
    }

    /// Number of worksheet Parts replaced (zero or one).
    #[must_use]
    pub const fn touched_worksheets(self) -> usize {
        if self.changed_rows == 0 { 0 } else { 1 }
    }
}
