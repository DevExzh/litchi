//! Guarded source-backed existing-row visibility edits of a
//! [`SourceBackedWorkbook`] (record 0360, the 0355a §1 shape).
//!
//! [`SourceBackedWorkbook::edit_row_visibility`] stages hide/unhide changes of
//! existing row owners of one worksheet; [`Edit::commit`] returns a
//! [`Commit`] whose output is the one-worksheet topology publication of the
//! value-only family; [`SourceBackedWorkbook::apply_row_visibility`] replays
//! a [`Patch`].

use std::collections::BTreeMap;
use std::io::Write;

use litchi_opc::SourceBackedPackage;
use litchi_sheet::Row;

use super::patch::Record;
use super::{Snapshot, rewrite};
use crate::error::{Result, invalid};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!(
    "row-visibility",
    Record,
    own_limits,
    own_diagnostics
);

crate::record::limit_kinds!("source-backed row-visibility", RowVisibility {
    /// Unique rows of one batch ([`MAX_BATCH_EDITS`]).
    BatchEdits => "rows of one row-visibility batch",
});

/// What a committed row-visibility change changed.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_rows: usize,
}

impl Diagnostics {
    /// Whether at least one row owner changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed_rows != 0
    }

    /// Number of existing row owners whose `hidden` attribute changed.
    #[must_use]
    pub const fn changed_rows(&self) -> usize {
        self.changed_rows
    }

    /// Number of worksheet parts replaced (zero or one).
    #[must_use]
    pub const fn touched_worksheets(&self) -> usize {
        if self.changed_rows == 0 { 0 } else { 1 }
    }
}

impl crate::workbook::source::deferred::Report for Record {
    type Diagnostics = Diagnostics;

    fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            changed_rows: self.changed_rows(),
        }
    }
}

impl Patch {
    /// The exact worksheet state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact worksheet state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }

    /// The patch with its target worksheet payload copied out of a managed
    /// source's reservation under an explicit `maximum_bytes` bound, so a
    /// patch made over a managed source replays onto a workbook outside that
    /// budget.
    ///
    /// # Errors
    ///
    /// Returns an error when the payload exceeds the bound.
    pub fn materialized(&self, maximum_bytes: usize) -> Result<Self> {
        Ok(Self(
            std::sync::Arc::new(self.0.materialized(maximum_bytes)?),
            None,
        ))
    }
}

impl crate::workbook::source::deferred::Change for Record {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        Record::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        crate::workbook::source::deferred::Change::check(self.inner(), package)
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        crate::workbook::source::deferred::Change::write(self.inner(), package, writer)
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        crate::workbook::source::deferred::Change::write_authenticated(
            self.inner(),
            package,
            writer,
        )
    }
}

impl SourceBackedWorkbook {
    /// Begin one atomic existing-row visibility edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector does not name a worksheet, its row
    /// closure is unsafe to edit, or the source cannot be read.
    pub fn edit_row_visibility<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        let before = self.edit_sheet_values(selector)?.before()?.clone();
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "row-visibility batch staging",
                source,
            })?;
        Ok(Edit {
            workbook: self.clone(),
            before: Snapshot::from_inner(before)?,
            staged,
        })
    }

    /// Replay a row-visibility patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the worksheet does not hold the
    /// patch's before state, and the signed-source refusal for a changed patch
    /// on a signed source.
    pub fn apply_row_visibility(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// Maximum unique existing row owners in one atomic visibility transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One direct visibility edit for an existing row owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct Change {
    row: Row,
    hidden: bool,
}

impl Change {
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

/// Clone-staged atomic row-visibility changes over one exact source worksheet
/// of a [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    staged: Vec<Change>,
}

impl Edit {
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
        self.apply_batch([Change::hide(row)])
    }

    /// Stage canonical unhiding of one existing row owner.
    pub fn unhide(&mut self, row: Row) -> Result<()> {
        self.apply_batch([Change::unhide(row)])
    }

    /// Stage a bounded visibility batch atomically.
    pub fn apply_batch(&mut self, edits: impl IntoIterator<Item = Change>) -> Result<()> {
        self.before.check_execution()?;
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "row-visibility pending batch",
                source,
            })?;
        for edit in edits {
            self.before.check_execution()?;
            if self.staged.len() + pending.len() == MAX_BATCH_EDITS {
                return Err(Error::limit(
                    LimitKind::BatchEdits,
                    MAX_BATCH_EDITS + 1,
                    MAX_BATCH_EDITS,
                ));
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
        self.before.check_execution()?;
        let original_len = self.staged.len();
        self.staged.extend(pending);
        if let Err(error) = self.before.check_execution() {
            self.staged.truncate(original_len);
            return Err(error);
        }
        Ok(())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged.clear();
    }

    /// Validate and rewrite the staged changes and return the committed
    /// workbook, the exact patch and the diagnostics.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        let actions = self
            .staged
            .iter()
            .map(|edit| (edit.row, edit.hidden))
            .collect::<BTreeMap<_, _>>();
        let (output, changed_rows) = rewrite::rewrite(self.before.source_xml(), &actions)?;
        self.before.check_execution()?;
        if changed_rows == 0 {
            let patch = Record::new(self.before.clone(), self.before.clone(), 0);
            self.before.check_execution()?;
            return commit_change(&self.workbook, patch);
        }
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        snapshot.check_execution()?;
        for edit in &self.staged {
            if snapshot.is_hidden(edit.row) != Some(edit.hidden) {
                return Err(invalid(
                    "row-visibility publication readback differs from staged state",
                ));
            }
        }
        self.before.check_execution()?;
        commit_change(
            &self.workbook,
            Record::new(self.before, snapshot, changed_rows),
        )
    }
}
