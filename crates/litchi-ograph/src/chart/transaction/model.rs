//! Typed transaction values and reversible change metadata.

use super::super::{Cache, Chart, RowCol, Value, XlValue, cache};
use super::chart_area;
use crate::{Error, Result};

/// Producer-specific value accepted by the bounded cache patcher.
#[derive(Debug, Clone, PartialEq)]
pub enum CacheValue {
    /// A standalone Graph datasheet value.
    Graph(Value),
    /// An Excel chart cache value, including `BoolErr` values.
    Excel(XlValue),
}

impl From<Value> for CacheValue {
    fn from(value: Value) -> Self {
        Self::Graph(value)
    }
}

impl From<XlValue> for CacheValue {
    fn from(value: XlValue) -> Self {
        Self::Excel(value)
    }
}

impl CacheValue {
    pub(crate) fn from_cache(cache: &Cache) -> Self {
        match cache {
            Cache::Graph { value, .. } => Self::Graph(value.clone()),
            Cache::Excel { value, .. } => Self::Excel(value.clone()),
        }
    }

    pub(crate) fn replace_cache(cache: &mut Cache, value: Self) -> Result<()> {
        match (cache, value) {
            (Cache::Graph { value: current, .. }, Self::Graph(value)) => *current = value,
            (Cache::Excel { value: current, .. }, Self::Excel(value)) => *current = value,
            _ => {
                return Err(Error::InvalidModel {
                    field: "cache",
                    reason: "replacement producer does not match the chart cache",
                });
            },
        }
        Ok(())
    }
}

/// Stable semantic identity of an existing cache cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Identity {
    /// Graph row, column, and IFmt identity.
    Graph {
        row: RowCol,
        col: RowCol,
        ifmt: cache::Ifmt,
    },
    /// Excel section, row, column, and XF identity.
    Excel {
        section: cache::Index,
        row: u16,
        col: u8,
        xf: cache::Xf,
    },
}

impl Identity {
    pub(crate) fn from_cache(cache: &Cache) -> Self {
        match cache {
            Cache::Graph { row, col, ifmt, .. } => Self::Graph {
                row: *row,
                col: *col,
                ifmt: *ifmt,
            },
            Cache::Excel {
                section,
                row,
                col,
                xf,
                ..
            } => Self::Excel {
                section: *section,
                row: *row,
                col: *col,
                xf: *xf,
            },
        }
    }
}

/// One source-checked semantic cache-value change.
#[derive(Debug, Clone, PartialEq)]
pub struct Change {
    index: usize,
    identity: Identity,
    before: CacheValue,
    after: CacheValue,
}

impl Change {
    pub(crate) const fn new(
        index: usize,
        identity: Identity,
        before: CacheValue,
        after: CacheValue,
    ) -> Self {
        Self {
            index,
            identity,
            before,
            after,
        }
    }

    /// Zero-based position in the chart's existing cache inventory.
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Stable source identity retained by this change.
    pub const fn identity(&self) -> &Identity {
        &self.identity
    }

    /// Value required before this change can be applied.
    pub const fn before(&self) -> &CacheValue {
        &self.before
    }

    /// Value produced by this change.
    pub const fn after(&self) -> &CacheValue {
        &self.after
    }
}

/// Reversible, deterministic semantic patch for chart snapshot edits.
#[derive(Debug, Clone, PartialEq)]
pub struct Patch {
    pub(super) changes: Box<[Change]>,
    pub(super) chart_area: Option<chart_area::Change>,
}

impl Patch {
    pub(crate) fn new(changes: Vec<Change>, chart_area: Option<chart_area::Change>) -> Self {
        Self {
            changes: changes.into_boxed_slice(),
            chart_area,
        }
    }

    /// Ordered cache changes in their transaction staging order.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Number of effective semantic changes, including the optional chart area.
    pub fn len(&self) -> usize {
        self.changes
            .len()
            .saturating_add(usize::from(self.chart_area.is_some()))
    }

    /// Whether the transaction was a semantic no-op.
    pub fn is_empty(&self) -> bool {
        self.changes.is_empty() && self.chart_area.is_none()
    }

    /// One reversible change to the fixed-size `[MS-OGRAPH]` `Chart` record.
    pub const fn chart_area(&self) -> Option<&chart_area::Change> {
        self.chart_area.as_ref()
    }

    /// Returns the source-checked inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            changes: self
                .changes
                .iter()
                .rev()
                .map(|change| {
                    Change::new(
                        change.index,
                        change.identity,
                        change.after.clone(),
                        change.before.clone(),
                    )
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            chart_area: self.chart_area.map(chart_area::Change::inverse),
        }
    }
}

/// Published result of a successful chart transaction.
#[derive(Debug)]
pub struct Commit {
    chart: Chart,
    patch: Patch,
}

impl Commit {
    pub(crate) const fn new(chart: Chart, patch: Patch) -> Self {
        Self { chart, patch }
    }

    /// Borrow the post-edit chart snapshot.
    pub const fn chart(&self) -> &Chart {
        &self.chart
    }

    /// Borrow the reversible semantic patch.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit and return the post-edit chart snapshot.
    pub fn into_chart(self) -> Chart {
        self.chart
    }

    /// Split the commit into the chart snapshot and reversible patch.
    pub fn into_parts(self) -> (Chart, Patch) {
        (self.chart, self.patch)
    }
}

/// One staged cache replacement, kept private to the transaction layer.
#[derive(Debug)]
pub(crate) struct Request {
    pub(crate) index: usize,
    pub(crate) value: CacheValue,
}
