//! Typed semantic values and transaction state for `Series` metadata.

use super::super::super::super::{Chart, Count, DataKind, Series};
use super::validation;
use crate::{Error, Result};

/// The editable scalar metadata carried by one `[MS-OGRAPH]` `Series` record.
///
/// `sdtY` and `sdtBSize` are deliberately not exposed: the specification
/// fixes both fields to numeric (`0x0001`), and changing them would alter the
/// record's grammar rather than its bounded metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Metadata {
    category_kind: DataKind,
    category_count: Count,
    value_count: Count,
    bubble_count: Count,
}

impl Metadata {
    /// Creates typed series metadata.  Wire-range validation is performed by
    /// [`Transaction::set`] before the value can be published.
    #[must_use]
    pub const fn new(
        category_kind: DataKind,
        category_count: Count,
        value_count: Count,
        bubble_count: Count,
    ) -> Self {
        Self {
            category_kind,
            category_count,
            value_count,
            bubble_count,
        }
    }

    /// Data kind of category or horizontal values (`sdtX`).
    #[must_use]
    pub const fn category_kind(self) -> DataKind {
        self.category_kind
    }

    /// Number of category or horizontal values (`cValx`).
    #[must_use]
    pub const fn category_count(self) -> Count {
        self.category_count
    }

    /// Number of vertical values (`cValy`).
    #[must_use]
    pub const fn value_count(self) -> Count {
        self.value_count
    }

    /// Number of bubble-size values (`cValBSize`).
    #[must_use]
    pub const fn bubble_count(self) -> Count {
        self.bubble_count
    }

    pub(super) const fn from_series(series: &Series) -> Self {
        Self::new(
            series.category_kind,
            series.category_count,
            series.value_count,
            series.bubble_count,
        )
    }

    pub(super) fn apply(self, series: &mut Series) {
        series.category_kind = self.category_kind;
        series.category_count = self.category_count;
        series.value_count = self.value_count;
        series.bubble_count = self.bubble_count;
    }
}

/// One source-checked replacement for an existing `Series` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    index: usize,
    offset: usize,
    before: Metadata,
    after: Metadata,
}

impl Change {
    pub(super) const fn new(
        index: usize,
        offset: usize,
        before: Metadata,
        after: Metadata,
    ) -> Self {
        Self {
            index,
            offset,
            before,
            after,
        }
    }

    /// Zero-based index of the existing semantic series.
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Source offset of the fixed-width `Series` record.
    #[must_use]
    pub const fn offset(self) -> usize {
        self.offset
    }

    /// Metadata required before this change can be applied.
    #[must_use]
    pub const fn before(self) -> Metadata {
        self.before
    }

    /// Metadata produced by this change.
    #[must_use]
    pub const fn after(self) -> Metadata {
        self.after
    }

    pub(super) const fn inverse(self) -> Self {
        Self::new(self.index, self.offset, self.after, self.before)
    }
}

/// Reversible source-checked patch for a series-metadata transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    source_before: u64,
    source_after: u64,
    changes: Box<[Change]>,
}

impl Patch {
    pub(super) fn new(source_before: u64, source_after: u64, changes: Vec<Change>) -> Self {
        Self {
            source_before,
            source_after,
            changes: changes.into_boxed_slice(),
        }
    }

    /// Ordered series changes in transaction staging order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Number of effective metadata changes.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.changes.len()
    }

    /// Whether this patch is a semantic no-op.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// FNV-1a source fingerprint required before application.
    #[must_use]
    pub const fn source_before(&self) -> u64 {
        self.source_before
    }

    /// FNV-1a source fingerprint produced after application.
    #[must_use]
    pub const fn source_after(&self) -> u64 {
        self.source_after
    }

    /// Returns the source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_before: self.source_after,
            source_after: self.source_before,
            changes: self
                .changes
                .iter()
                .rev()
                .copied()
                .map(Change::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
        }
    }

    /// Applies the patch to an exact matching parsed chart source.
    pub fn apply(&self, chart: Chart) -> Result<Commit> {
        let snapshot = Snapshot::from_chart(chart)?;
        if snapshot.source != self.source_before {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch source fingerprint does not match the target snapshot",
            });
        }
        let mut transaction = snapshot.edit();
        for change in &self.changes {
            transaction.set_expected(*change)?;
        }
        transaction.commit()
    }
}

/// Parsed chart snapshot with a validated `Series` record inventory.
#[derive(Debug)]
pub struct Snapshot {
    pub(super) chart: Chart,
    pub(super) entries: Box<[Entry]>,
    pub(super) source: u64,
}

impl Snapshot {
    /// Captures a pristine parsed chart without copying its retained source.
    pub fn from_chart(chart: Chart) -> Result<Self> {
        if !chart.is_pristine() {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "only a pristine parsed chart has a replayable source stream",
            });
        }
        let scan = super::codec::scan(&chart)?;
        Ok(Self {
            chart,
            entries: scan.entries.into_boxed_slice(),
            source: scan.source,
        })
    }

    /// Borrow the unchanged chart snapshot.
    #[must_use]
    pub const fn chart(&self) -> &Chart {
        &self.chart
    }

    /// Number of series records in the snapshot.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the chart contains no series records.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Returns one typed series metadata value by zero-based index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<Metadata> {
        self.entries.get(index).map(|entry| entry.metadata)
    }

    /// Iterates metadata in the source `Series` order without allocation.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = Metadata> + '_ {
        self.entries.iter().map(|entry| entry.metadata)
    }

    /// Starts a transaction consuming this snapshot.
    #[must_use]
    pub fn edit(self) -> Transaction {
        Transaction {
            chart: self.chart,
            entries: self.entries,
            source: self.source,
            requests: Vec::new(),
        }
    }

    /// Consumes the snapshot without editing it.
    #[must_use]
    pub fn into_chart(self) -> Chart {
        self.chart
    }
}

/// Published result of a successful series-metadata transaction.
#[derive(Debug)]
pub struct Commit {
    chart: Chart,
    patch: Patch,
}

impl Commit {
    pub(super) const fn new(chart: Chart, patch: Patch) -> Self {
        Self { chart, patch }
    }

    /// Borrow the post-edit chart snapshot.
    #[must_use]
    pub const fn chart(&self) -> &Chart {
        &self.chart
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit and return the post-edit chart.
    #[must_use]
    pub fn into_chart(self) -> Chart {
        self.chart
    }

    /// Split the commit into chart and patch.
    #[must_use]
    pub fn into_parts(self) -> (Chart, Patch) {
        (self.chart, self.patch)
    }
}

/// A bounded transaction over one parsed chart snapshot.
#[derive(Debug)]
pub struct Transaction {
    pub(super) chart: Chart,
    pub(super) entries: Box<[Entry]>,
    pub(super) source: u64,
    pub(super) requests: Vec<Request>,
}

impl Transaction {
    /// Number of distinct staged metadata replacements.
    #[must_use]
    pub fn len(&self) -> usize {
        self.requests.len()
    }

    /// Whether no metadata replacement is staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.requests.is_empty()
    }

    /// Stages a typed replacement for one existing `Series` record.
    pub fn set(&mut self, index: usize, metadata: Metadata) -> Result<&mut Self> {
        let entry = self.entries.get(index).ok_or(Error::InvalidModel {
            field: "series metadata",
            reason: "series index is outside the parsed chart",
        })?;
        validation::ensure(metadata)?;
        if let Some(request) = self
            .requests
            .iter_mut()
            .find(|request| request.index == index)
        {
            request.metadata = metadata;
            return Ok(self);
        }
        if self.requests.len() >= self.chart.limits.max_series {
            return Err(Error::LimitExceeded {
                resource: "series metadata edit count",
                observed: u64::try_from(self.requests.len().saturating_add(1)).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.chart.limits.max_series).unwrap_or(u64::MAX),
            });
        }
        self.requests.try_reserve(1).ok().ok_or(Error::Allocation {
            resource: "series metadata edits",
        })?;
        let _ = entry;
        self.requests.push(Request { index, metadata });
        Ok(self)
    }

    /// Validates and publishes the staged metadata edits atomically.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            mut chart,
            entries,
            source,
            requests,
        } = self;
        let current = super::codec::scan(&chart)?;
        if current.source != source {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "transaction source changed before publication",
            });
        }

        let mut changes = Vec::new();
        changes
            .try_reserve_exact(requests.len())
            .ok()
            .ok_or(Error::Allocation {
                resource: "series metadata patch changes",
            })?;
        for request in requests {
            let entry = entries.get(request.index).ok_or(Error::InvalidModel {
                field: "series metadata",
                reason: "series index disappeared during transaction",
            })?;
            if entry.metadata == request.metadata {
                continue;
            }
            changes.push(Change::new(
                request.index,
                entry.offset,
                entry.metadata,
                request.metadata,
            ));
        }

        let source_after = if changes.is_empty() {
            source
        } else {
            super::codec::patch(&mut chart, source, &changes)?
        };
        let patch = Patch::new(source, source_after, changes);
        Ok(Commit::new(chart, patch))
    }

    fn set_expected(&mut self, change: Change) -> Result<&mut Self> {
        let entry = self.entries.get(change.index).ok_or(Error::InvalidModel {
            field: "series metadata",
            reason: "patch series index is outside the target chart",
        })?;
        if entry.offset != change.offset {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch Series record offset does not match the target snapshot",
            });
        }
        if entry.metadata != change.before {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch source metadata does not match the target snapshot",
            });
        }
        self.set(change.index, change.after)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Request {
    pub(super) index: usize,
    pub(super) metadata: Metadata,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Entry {
    pub(super) index: usize,
    pub(super) offset: usize,
    pub(super) metadata: Metadata,
}
