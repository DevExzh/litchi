//! Guarded source-backed value-only worksheet transactions.

use std::collections::BTreeMap;
use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{ReadLimits, SourceBackedPackage};
use litchi_sheet::Cell as Address;

use super::{Commit, Patch, Snapshot};
use crate::Selector;
use crate::cell::{Content, Value};
use crate::error::{Error, Result, invalid};
use crate::raw::worksheet::edit::{Action, rewrite};

/// Maximum unique existing cells in one atomic value transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One typed value-only edit for an existing scalar cell.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellValueEdit {
    /// Replace the stored scalar value without changing local style.
    Set { address: Address, value: Value },
    /// Remove the stored scalar payload while retaining the cell record.
    Clear { address: Address },
    /// Remove the complete stored scalar cell record.
    ///
    /// Unlike [`Self::Clear`], this also removes cell-local type and style
    /// attributes. The containing row and the producer's declared worksheet
    /// dimension are retained conservatively.
    Remove { address: Address },
}

impl CellValueEdit {
    /// Construct a checked value replacement.
    pub fn set(address: Address, value: impl Into<Value>) -> Self {
        Self::Set {
            address,
            value: value.into(),
        }
    }

    /// Construct a scalar-value clear.
    #[must_use]
    pub const fn clear(address: Address) -> Self {
        Self::Clear { address }
    }

    /// Construct a complete scalar-cell owner removal.
    #[must_use]
    pub const fn remove(address: Address) -> Self {
        Self::Remove { address }
    }

    /// Target coordinate.
    #[must_use]
    pub const fn address(&self) -> Address {
        match self {
            Self::Set { address, .. } | Self::Clear { address } | Self::Remove { address } => {
                *address
            },
        }
    }
}

#[derive(Clone, Debug)]
enum StagedValueEdit {
    Set(Value),
    Clear,
    Remove,
}

/// Owning source-backed editor for one exact XLSX artifact.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// Clone-staged atomic changes over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Vec<(Address, StagedValueEdit)>,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with explicit OPC ingress limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Ok(Self {
            package: SourceBackedPackage::from_read_at_with_limits(source, limits)?,
        })
    }

    /// Capture the exact safe value-only closure for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin one atomic value-only edit.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "value-only batch staging",
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
        self.package.cache_diagnostics()
    }

    /// Publish one exact-source-checked worksheet overlay to a sequential sink.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        let current =
            Snapshot::load_source_backed(&self.package, commit.patch().before().sheet_position())?;
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
        self.write_snapshot_overlay_to_stream(writer, &target)?;
        Ok(target)
    }

    pub(crate) fn write_snapshot_overlay_to_stream<W: Write>(
        self,
        writer: W,
        target: &Snapshot,
    ) -> Result<()> {
        self.package.write_part_overlay_to_stream(
            writer,
            target.worksheet_part_name(),
            target.source_xml().to_vec(),
        )?;
        Ok(())
    }
}

impl SourceEdit {
    /// Exact source snapshot captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Number of unique staged selectors.
    #[must_use]
    pub fn len(&self) -> usize {
        self.staged.len()
    }

    /// Whether no selectors have been staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.staged.is_empty()
    }

    /// Stage one value replacement. Repeated selectors are rejected.
    pub fn set(&mut self, address: Address, value: impl Into<Value>) -> Result<()> {
        self.apply_batch([CellValueEdit::set(address, value)])
    }

    /// Stage one scalar clear. Repeated selectors are rejected.
    pub fn clear(&mut self, address: Address) -> Result<()> {
        self.apply_batch([CellValueEdit::clear(address)])
    }

    /// Stage complete removal of one scalar `<c>` owner.
    ///
    /// The containing `<row>` is retained even when this is its last cell.
    /// Repeated selectors are rejected.
    pub fn remove(&mut self, address: Address) -> Result<()> {
        self.apply_batch([CellValueEdit::remove(address)])
    }

    /// Stage a bounded batch atomically.
    pub fn apply_batch(&mut self, edits: impl IntoIterator<Item = CellValueEdit>) -> Result<()> {
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "value-only pending batch",
                source,
            })?;
        for edit in edits {
            if self.staged.len() + pending.len() >= MAX_BATCH_EDITS {
                return Err(invalid(format!(
                    "value-only batch exceeds {MAX_BATCH_EDITS} unique cells"
                )));
            }
            let address = edit.address();
            if self
                .staged
                .iter()
                .chain(&pending)
                .any(|(stored, _)| *stored == address)
            {
                return Err(invalid(format!(
                    "duplicate value-only selector '{address}'"
                )));
            }
            let source = self.before.value(address).ok_or_else(|| {
                invalid(format!(
                    "value-only selector '{address}' is not an existing scalar value cell"
                ))
            })?;
            let value = match edit {
                CellValueEdit::Set { value, .. } => {
                    value.validate_for_write()?;
                    if matches!(value, Value::Date(_)) {
                        return Err(invalid("value-only batches currently refuse date cells"));
                    }
                    StagedValueEdit::Set(value)
                },
                CellValueEdit::Clear { .. } => StagedValueEdit::Clear,
                CellValueEdit::Remove { .. } => StagedValueEdit::Remove,
            };
            if matches!(source, Value::Date(_)) {
                return Err(invalid("value-only batches currently refuse date cells"));
            }
            pending.push((address, value));
        }
        self.staged.extend(pending);
        Ok(())
    }

    /// Validate, rewrite once, and freeze an exact reversible commit.
    pub fn commit(self) -> Result<Commit> {
        let mut actions = BTreeMap::new();
        for (address, value) in &self.staged {
            let action = match value {
                StagedValueEdit::Set(value) => {
                    if Some(value) == self.before.value(*address) {
                        continue;
                    }
                    Action::set(Content::Value(value.clone()))
                },
                StagedValueEdit::Clear => Action::clear(false),
                StagedValueEdit::Remove => Action::Remove,
            };
            actions.insert(*address, action);
        }
        if actions.is_empty() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, 0));
        }
        let changed = actions.len();
        let output = rewrite(self.before.source_xml(), self.before.sheet_name(), actions)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        for (address, expected) in &self.staged {
            let matches = match expected {
                StagedValueEdit::Set(value) => snapshot.value(*address) == Some(value),
                StagedValueEdit::Clear => {
                    snapshot.contains_cell(*address) && snapshot.value(*address).is_none()
                },
                StagedValueEdit::Remove => !snapshot.contains_cell(*address),
            };
            if !matches {
                return Err(invalid(
                    "value-only publication readback differs from staged state",
                ));
            }
        }
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, changed))
    }
}
