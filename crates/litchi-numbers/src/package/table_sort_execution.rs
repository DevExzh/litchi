//! Execution of a table's persisted sort order: the body rows move
//! physically, as Numbers' **Organize → Sort → Sort Now** does (record
//! 0421).
//!
//! The rewrite is the shared executor `litchi_iwa_table::sort` (tile row
//! records, row header buckets, the row UID map, stroke layers and
//! user-hidden row states; the host's scalar subset and feature refusals).
//! This facade resolves the table, reads its persisted order, checks the
//! scope and merges, supplies the user-hidden rows and verifies the order
//! and the hidden rows on the reopened candidate. Root previews are kept,
//! as the host keeps them, and the persisted sort order is unchanged.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::HashSet;
use std::fmt;

use litchi_iwa_common::table::axis::AxisIndex;
use litchi_iwa_common::table::sort::{Order, RowRange, Scope};
use litchi_iwa_table::sort::{self as engine, Hooks, Target};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{self, Working};
use crate::selector::{SheetSelector, TableSelector};

pub use litchi_iwa_table::sort::Unsupported;

structure::structural_family!("table sort execution");

/// A typed, content-redacted table sort-execution failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table sort-execution source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table sort items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the sorted table.
    #[error("table sort candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table sort patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The table has no persisted sort order.
    #[error("the table has no persisted sort order")]
    NoSortOrder,
    /// The persisted order's scope does not match the staged execution.
    #[error("the persisted sort order applies to another scope")]
    ScopeMismatch,
    /// The row range is outside the table's body rows.
    #[error("the sort row range is outside the table's body rows")]
    RowsOutOfRange,
    /// The table holds content this executor does not move.
    #[error("the table cannot be sorted: {reason:?}")]
    Unsupported {
        /// What the table holds.
        reason: Unsupported,
    },
    /// No execution was staged.
    #[error("no sort execution is staged")]
    NoStagedOperation,
    /// A second execution was staged on one edit.
    #[error("one sort edit stages one execution")]
    OperationAlreadyStaged,
    /// A surviving object depends on the edited graph.
    #[error("a surviving object depends on the table's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support exact table sorts")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound => RefusalKind::NotFound,
            Self::NoSortOrder | Self::ScopeMismatch => RefusalKind::UnsupportedTarget,
            Self::RowsOutOfRange => RefusalKind::InvalidValue,
            Self::Unsupported { .. } | Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::NoStagedOperation | Self::OperationAlreadyStaged => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Unsupported> for Error {
    fn from(reason: Unsupported) -> Self {
        Self::Refused(Refusal::Unsupported { reason })
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(engine::Refusal::RowsOutOfRange) => {
                Self::Refused(Refusal::RowsOutOfRange)
            },
            engine::Error::Refused(engine::Refusal::Unsupported(reason)) => Self::from(reason),
            engine::Error::Fault(fault) => Self::from(fault),
            engine::Error::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// One sort execution staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    staged: Option<Option<RowRange>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the execution of an entire-table sort order over every body
    /// row.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] for a second execution.
    pub fn apply(&mut self) -> Result<&mut Self, Error> {
        self.stage(None)
    }

    /// Stage the execution of a selected-rows sort order over a body-relative
    /// half-open row range (header and footer rows excluded). Numbers keeps
    /// the selected range in view state, so the caller supplies it.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] for a second execution.
    pub fn apply_to_rows(&mut self, rows: RowRange) -> Result<&mut Self, Error> {
        self.stage(Some(rows))
    }

    fn stage(&mut self, rows: Option<RowRange>) -> Result<&mut Self, Error> {
        if self.staged.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.staged = Some(rows);
        Ok(self)
    }

    /// Validate and publish the staged execution. Rows already in the
    /// requested stable order are an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; [`Refusal::NoSortOrder`] and
    /// [`Refusal::ScopeMismatch`]; [`Refusal::Unsupported`] for content the
    /// executor does not move; source, resource and verification errors.
    /// Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let rows = self.staged.ok_or(Refusal::NoStagedOperation)?;
        match execute(self.source, self.sheet, self.table, rows)? {
            Some(published) => Commit::new(self.source, published),
            None => Commit::unchanged(self.source),
        }
    }
}

/// Run one staged execution.
fn execute(
    source: &Package,
    sheet: usize,
    table: usize,
    rows: Option<RowRange>,
) -> Result<Option<structure::Published>, Error> {
    let order: Order = source
        .table_sort_order(sheet, table)
        .map_err(|_error| Error::InvalidSource)?
        .ok_or(Refusal::NoSortOrder)?;
    match (rows, order.scope()) {
        (None, Scope::EntireTable) | (Some(_), Scope::SelectedRows) => {},
        _ => return Err(Refusal::ScopeMismatch.into()),
    }
    if !source
        .table_merges(sheet, table)
        .map_err(|_error| Error::InvalidSource)?
        .is_empty()
    {
        return Err(Unsupported::MergedCells.into());
    }
    let target = {
        let working = Working::new(source)?;
        let located = super::table_topology::locate(&working, sheet, table)
            .map_err(|_error| Error::InvalidSource)?;
        Target {
            info: located.info,
            model: located.model,
        }
    };
    // User-hidden rows keep their positions in a native sort: the engine
    // asks for them only when the table declares hidden axes and rows move.
    let mut hidden = None;
    let executed = engine::execute(
        source,
        &target,
        &order,
        rows,
        || {
            let axes = source.table_hidden_axes(sheet, table).map_err(|_error| {
                engine::Error::Refused(engine::Refusal::Unsupported(Unsupported::HiddenAxes))
            })?;
            let rows = axes
                .iter()
                .filter_map(|axis| match axis {
                    AxisIndex::Row(index) => Some(index),
                    AxisIndex::Column(_) => None,
                })
                .collect::<HashSet<_>>();
            if !axes.is_empty() {
                hidden = Some(axes);
            }
            Ok(Some(rows))
        },
        Hooks::keeping_previews(),
    )?;
    let Some(published) = executed.published else {
        return Ok(None);
    };
    let candidate = &published.package;
    if candidate
        .table_sort_order(sheet, table)
        .map_err(|_error| Error::Verification)?
        .as_ref()
        != Some(&order)
    {
        return Err(Error::Verification);
    }
    if let Some(hidden) = hidden {
        if candidate
            .table_hidden_axes(sheet, table)
            .map_err(|_error| Error::Verification)?
            != hidden
        {
            return Err(Error::Verification);
        }
    }
    Ok(Some(published))
}

impl Package {
    /// Begin executing a table's persisted sort order (Numbers' **Sort
    /// Now**).
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] or [`Refusal::TableNotFound`] when a
    /// selector resolves nothing.
    pub fn edit_table_sort_execution<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Refusal::SheetNotFound)?;
        let sheet_position = sheet.index();
        let table = match table.into() {
            TableSelector::Position(position) => {
                (position.get() < sheet.table_count()).then_some(position.get())
            },
            TableSelector::Name(name) => {
                let mut matches = sheet
                    .tables()
                    .enumerate()
                    .filter(|(_, table)| table.name() == name);
                match (matches.next(), matches.next()) {
                    (Some((position, _)), None) => Some(position),
                    _ => None,
                }
            },
        }
        .ok_or(Refusal::TableNotFound)?;
        Ok(Edit {
            source: self,
            sheet: sheet_position,
            table,
            staged: None,
        })
    }

    /// Replay an exact-source sort-execution patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_sort_execution(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
