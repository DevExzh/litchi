//! Failure-atomic typed edits at the standalone Graph package boundary.

use crate::{Result, chart};

use super::{Snapshot, codec, patch};

/// Source-checked transaction over the one standalone Graph Workbook chart.
///
/// Only existing fixed-width chart payloads exposed by the chart editor are
/// reachable here. The transaction never recalculates formulas, follows links,
/// renders a chart, or activates an OLE object.
#[derive(Debug)]
pub struct Transaction {
    source: Snapshot,
    editor: chart::Editor,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Result<Self> {
        let chart = codec::read_chart(&source)?;
        Ok(Self {
            source,
            editor: chart.edit()?,
        })
    }

    /// Immutable package source used for conflict checks and rollback.
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Number of typed chart operations staged in this transaction.
    pub fn len(&self) -> usize {
        self.editor.len()
    }

    /// Whether no chart payload has been staged.
    pub fn is_empty(&self) -> bool {
        self.editor.is_empty()
    }

    /// Whether the package will differ after a successful commit.
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Replace one existing Graph cache value without evaluating its source.
    pub fn set_cache_value<V>(&mut self, index: usize, value: V) -> Result<&mut Self>
    where
        V: Into<chart::CacheValue>,
    {
        self.editor.set_cache_value(index, value)?;
        Ok(self)
    }

    /// Replace the fixed-size Graph chart-area rectangle.
    pub fn set_rect(&mut self, value: chart::Rect) -> Result<&mut Self> {
        self.editor.set_rect(value)?;
        Ok(self)
    }

    /// Replace the fixed-size Graph `ShtProps` metadata payload.
    pub fn set_props(&mut self, value: chart::Props) -> Result<&mut Self> {
        self.editor.set_props(value)?;
        Ok(self)
    }

    /// Validate and publish the package edit atomically.
    pub fn commit(self) -> Result<patch::Commit> {
        let Self { source, editor } = self;
        let chart_commit = editor.commit()?;
        let chart_patch = chart_commit.patch().clone();
        if chart_patch.is_empty() {
            let patch = patch::Patch::new(source.source_arc(), source.source_arc(), chart_patch);
            return Ok(patch::Commit::new(source.clone(), patch, false));
        }

        let chart = chart_commit.into_chart();
        let stream = chart.encode()?;
        let bytes = codec::replace_chart(&source, stream.as_bytes())?;
        let snapshot = Snapshot::from_bytes_with_limits(bytes, source.limits())?;
        let patch = patch::Patch::new(source.source_arc(), snapshot.source_arc(), chart_patch);
        Ok(patch::Commit::new(snapshot, patch, true))
    }

    /// Discard staged operations and recover the exact source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}
