//! Guarded source-backed existing-row visibility transactions.

use std::collections::BTreeMap;
use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::ReadLimits;
use litchi_sheet::Row;

use super::{Commit, Patch, Snapshot, rewrite};
use crate::Selector;
use crate::cell_values;
use crate::error::{Error, Result, invalid};

/// Maximum unique existing row owners in one atomic visibility transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One direct visibility edit for an existing row owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct RowVisibilityEdit {
    row: Row,
    hidden: bool,
}

impl RowVisibilityEdit {
    /// Construct a checked hide operation.
    #[must_use]
    pub const fn hide(row: Row) -> Self {
        Self { row, hidden: true }
    }

    /// Construct a checked unhide operation.
    #[must_use]
    pub const fn unhide(row: Row) -> Self {
        Self { row, hidden: false }
    }

    /// Target zero-based row coordinate.
    #[must_use]
    pub const fn row(self) -> Row {
        self.row
    }

    /// Requested effective visibility state.
    #[must_use]
    pub const fn hidden(self) -> bool {
        self.hidden
    }
}

/// Owning source-backed editor for one exact XLSX artifact.
pub struct SourceBackedEditor {
    inner: cell_values::SourceBackedEditor,
}

/// Clone-staged atomic row-visibility changes over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Vec<RowVisibilityEdit>,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with explicit OPC ingress limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Ok(Self {
            inner: cell_values::SourceBackedEditor::from_read_at_with_limits(source, limits)?,
        })
    }

    /// Capture the exact safe visibility closure for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        Snapshot::from_inner(self.inner.snapshot(selector)?)
    }

    /// Begin one atomic existing-row visibility edit.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "row-visibility batch staging",
                source,
            })?;
        Ok(SourceEdit {
            before: self.snapshot(selector)?,
            staged,
        })
    }

    /// Content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.inner.cache_diagnostics()
    }

    /// Publish one exact-source-checked worksheet overlay to a sequential sink.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        let current = self.snapshot(commit.patch().before().sheet_position())?;
        if !current.same_source(commit.patch().before()) {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current
        } else {
            commit.patch().after().clone()
        };
        self.inner
            .write_snapshot_overlay_to_stream(writer, target.inner())?;
        Ok(target)
    }
}

impl SourceEdit {
    /// Exact source snapshot captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Number of unique staged row selectors.
    #[must_use]
    pub fn len(&self) -> usize {
        self.staged.len()
    }

    /// Whether no row selectors have been staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.staged.is_empty()
    }

    /// Stage canonical hiding of one existing row owner.
    pub fn hide(&mut self, row: Row) -> Result<()> {
        self.apply_batch([RowVisibilityEdit::hide(row)])
    }

    /// Stage canonical unhiding of one existing row owner.
    pub fn unhide(&mut self, row: Row) -> Result<()> {
        self.apply_batch([RowVisibilityEdit::unhide(row)])
    }

    /// Stage a bounded visibility batch atomically.
    pub fn apply_batch(
        &mut self,
        edits: impl IntoIterator<Item = RowVisibilityEdit>,
    ) -> Result<()> {
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "row-visibility pending batch",
                source,
            })?;
        for edit in edits {
            if self.staged.len() + pending.len() == MAX_BATCH_EDITS {
                return Err(invalid(format!(
                    "row-visibility batch exceeds {MAX_BATCH_EDITS} unique rows"
                )));
            }
            if self
                .staged
                .iter()
                .chain(&pending)
                .any(|stored| stored.row == edit.row)
            {
                return Err(invalid(format!(
                    "duplicate row-visibility selector '{}'",
                    edit.row
                )));
            }
            if !self.before.contains_row(edit.row) {
                return Err(invalid(format!(
                    "row-visibility selector '{}' is not an existing row owner",
                    edit.row
                )));
            }
            pending.push(edit);
        }
        self.staged.extend(pending);
        Ok(())
    }

    /// Validate, rewrite once, and freeze an exact reversible commit.
    pub fn commit(self) -> Result<Commit> {
        let actions = self
            .staged
            .iter()
            .map(|edit| (edit.row, edit.hidden))
            .collect::<BTreeMap<_, _>>();
        let (output, changed_rows) = rewrite::rewrite(self.before.source_xml(), &actions)?;
        if changed_rows == 0 {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, 0));
        }
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        for edit in &self.staged {
            if snapshot.is_hidden(edit.row) != Some(edit.hidden) {
                return Err(invalid(
                    "row-visibility publication readback differs from staged state",
                ));
            }
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, changed_rows))
    }
}
