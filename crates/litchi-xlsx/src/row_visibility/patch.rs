//! Exact reversible row-visibility worksheet records.

use super::Snapshot;
use crate::cell_values::patch::SheetRecord;

/// Exact source-bound replacement of direct existing-row visibility: the
/// replay payload of [`super::source::Patch`].
#[derive(Clone, Debug)]
pub(crate) struct Record {
    before: Snapshot,
    after: Snapshot,
    inner: SheetRecord,
    changed_rows: usize,
}

impl Record {
    pub(super) fn new(before: Snapshot, after: Snapshot, changed_rows: usize) -> Self {
        let inner = SheetRecord::new(before.inner().clone(), after.inner().clone(), changed_rows);
        Self {
            before,
            after,
            inner,
            changed_rows,
        }
    }

    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    pub(crate) const fn inner(&self) -> &SheetRecord {
        &self.inner
    }

    /// The record with its target worksheet payload copied out of a managed
    /// reservation under `maximum_bytes`.
    pub(crate) fn materialized(&self, maximum_bytes: usize) -> crate::Result<Self> {
        Ok(Self {
            before: self.before.clone(),
            after: self.after.clone(),
            inner: self.inner.materialized(maximum_bytes)?,
            changed_rows: self.changed_rows,
        })
    }

    pub(crate) const fn changed_rows(&self) -> usize {
        self.changed_rows
    }

    /// Exact source-bound inverse, including the producer's original lexical
    /// representation of the `hidden` attribute.
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            inner: self.inner.inverse(),
            changed_rows: self.changed_rows,
        }
    }
}
