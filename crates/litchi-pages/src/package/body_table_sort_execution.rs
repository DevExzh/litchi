//! Execution of a body table's persisted sort order: the body rows move
//! physically, as Pages' **Sort Now** does (record 0371).
//!
//! The rewrite is the shared executor `litchi_iwa_table::sort` (lifted from
//! this module in record 0421): the facade resolves the table, reads its
//! persisted order, checks the scope and merges, supplies the user-hidden
//! rows, and verifies them after publication. Root previews are kept, as the
//! host keeps them.

use std::collections::HashSet;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::table::axis::AxisIndex;
use litchi_iwa_common::table::sort::{Order, RowRange, Scope};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_table::sort::{self as engine, Hooks, Target};
use litchi_iwa_table::structure::{Fault, FaultLimit};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{self as graph, Budget, GraphError, ObjectIndex};
use crate::selector::BodyTableSelector;

pub use litchi_iwa_table::sort::Unsupported;

const ROW_STORAGE_FIELD: u32 = 6;
const ROW_OFFSETS_FIELD: u32 = 7;
const ROW_WIDE_OFFSETS_FIELD: u32 = 8;

/// A finite resource charged by one sort execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
        })
    }
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector resolved no body table.
    #[error("the Pages body has no such table")]
    TableNotFound,
    /// The table has no persisted sort order.
    #[error("the Pages table has no persisted sort order")]
    NoSortOrder,
    /// The persisted order's scope does not match the staged execution.
    #[error("the persisted sort order applies to another scope")]
    ScopeMismatch,
    /// The row range is outside the table's body rows.
    #[error("the sort row range is outside the table's body rows")]
    RowsOutOfRange,
    /// The table holds content this executor does not move.
    #[error("the Pages table cannot be sorted: {0:?}")]
    Unsupported(Unsupported),
    /// No execution was staged.
    #[error("no sort execution is staged")]
    NoStagedOperation,
    /// A second execution was staged on one edit.
    #[error("one sort edit stages one execution")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact table sorts")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::NoSortOrder | Self::ScopeMismatch => RefusalKind::UnsupportedTarget,
            Self::RowsOutOfRange => RefusalKind::InvalidValue,
            Self::Unsupported(_) | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation | Self::OperationAlreadyStaged => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

/// Failure from a sort execution.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The table graph is malformed.
    #[error("the selected Pages table is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages table sort {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages table sort")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the sorted table.
    #[error("the sorted Pages table failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages table sort patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Unsupported> for Error {
    fn from(reason: Unsupported) -> Self {
        Self::Refused(Refusal::Unsupported(reason))
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Unsupported | Fault::Dependency => Self::Refused(Refusal::UnsupportedSource),
            Fault::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(engine::Refusal::RowsOutOfRange) => {
                Self::Refused(Refusal::RowsOutOfRange)
            },
            engine::Error::Refused(engine::Refusal::Unsupported(reason)) => {
                Self::Refused(Refusal::Unsupported(reason))
            },
            engine::Error::Fault(fault) => Self::from(fault),
            engine::Error::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// One sort execution staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    table: usize,
    staged: Option<Option<RowRange>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
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
    /// Refuses a second execution.
    pub fn apply(&mut self) -> Result<&mut Self, Error> {
        self.stage(None)
    }

    /// Stage the execution of a selected-rows sort order over a body-relative
    /// half-open row range (header and footer rows excluded).
    ///
    /// # Errors
    ///
    /// Refuses a second execution.
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

    /// Validate and publish the staged execution.
    pub fn commit(self) -> Result<Commit, Error> {
        let rows = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let (package, moved) = execute(self.source, self.table, rows)?;
        let target = package.state.source.shared_source();
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics: Diagnostics {
                changed: moved != 0,
                moved_rows: moved,
            },
        })
    }
}

/// Exact-source reversible sort-execution patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the exact artifact is unchanged (the body was already in
    /// order).
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics of one sort execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    moved_rows: usize,
}

impl Diagnostics {
    /// Whether any row moved.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Body rows that changed position.
    #[must_use]
    pub const fn moved_rows(self) -> usize {
        self.moved_rows
    }
}

/// The validated immutable result of one sort execution.
#[must_use = "a sort commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Begin executing a body table's persisted sort order.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selector resolves no table.
    pub fn edit_body_table_sort_execution<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let catalog = self.body_tables().map_err(|_error| Error::InvalidSource)?;
        let table = match selector {
            BodyTableSelector::Position(position) => {
                (position.get() < catalog.len()).then_some(position.get())
            },
            BodyTableSelector::Name(name) => {
                let matches = (0..catalog.len())
                    .filter(|index| {
                        catalog
                            .get(*index)
                            .is_some_and(|entry| entry.name() == name)
                    })
                    .collect::<Vec<_>>();
                match matches.as_slice() {
                    [index] => Some(*index),
                    _ => None,
                }
            },
        }
        .ok_or(Refusal::TableNotFound)?;
        Ok(Edit {
            source: self,
            table,
            staged: None,
        })
    }

    /// Replay an exact-source sort-execution patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_table_sort_execution(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            graph::reopen(self, patch.artifacts.target()).map_err(|_error| Error::Verification)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                moved_rows: 0,
            },
        })
    }
}

/// A single varint field of `payload`, if present.
pub(super) fn varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>, GraphError> {
    graph::unique_field(view, number, 0)?
        .map(|field| graph::canonical_varint(field.payload()))
        .transpose()
}

/// A single `TSP.Reference` field's identifier, if present.
pub(super) fn reference(
    view: &WireView<'_>,
    number: u32,
    budget: &mut Budget,
) -> Result<Option<NonZeroU64>, GraphError> {
    graph::unique_field(view, number, 2)?
        .map(|field| graph::parse_reference(field.payload(), budget))
        .transpose()
}

/// The single message of one of `types` in object `identifier`: its
/// component, index, type and payload.
pub(super) fn located_message<'a>(
    index: &ObjectIndex<'a>,
    identifier: NonZeroU64,
    types: &[u32],
) -> Result<(Arc<str>, usize, u32, &'a [u8]), GraphError> {
    let located = index.get(identifier)?;
    let mut found = None;
    for (message_index, message) in located.object.messages.iter().enumerate() {
        if types.contains(&message.type_)
            && found
                .replace((message_index, message.type_, message.data.as_slice()))
                .is_some()
        {
            return Err(GraphError::InvalidSource);
        }
    }
    let (message_index, message_type, payload) = found.ok_or(GraphError::InvalidSource)?;
    graph::validate_message_metadata(located.object, message_index)?;
    Ok((
        index.component_name(located.component_index)?,
        message_index,
        message_type,
        payload,
    ))
}

/// The BNC cells of one row record, by column; `None` for a legacy
/// (pre-BNC) row.
pub(super) fn row_cells<'a>(
    row: &'a [u8],
    budget: &mut Budget,
) -> Result<Option<Vec<Option<&'a [u8]>>>, GraphError> {
    let view = budget.parse(row, 4)?;
    let (Some(storage), Some(offsets)) = (
        graph::unique_field(&view, ROW_STORAGE_FIELD, 2)?,
        graph::unique_field(&view, ROW_OFFSETS_FIELD, 2)?,
    ) else {
        return Ok(None);
    };
    let (storage, offsets) = (storage.payload(), offsets.payload());
    let unit = if varint(&view, ROW_WIDE_OFFSETS_FIELD)?.unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    if offsets.len() % 2 != 0 {
        return Err(GraphError::InvalidSource);
    }
    let starts = offsets
        .chunks_exact(2)
        .map(|pair| {
            let raw = u16::from_le_bytes([pair[0], pair[1]]);
            (raw != u16::MAX).then(|| usize::from(raw) * unit)
        })
        .collect::<Vec<_>>();
    let mut present = starts.iter().flatten().copied().collect::<Vec<_>>();
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.last().is_some_and(|last| *last >= storage.len())
    {
        return Err(GraphError::InvalidSource);
    }
    present.push(storage.len());
    let mut cells = Vec::with_capacity(starts.len());
    for start in starts {
        cells.push(start.map(|start| {
            let end = present
                .iter()
                .copied()
                .find(|candidate| *candidate > start)
                .unwrap_or(storage.len());
            &storage[start..end]
        }));
    }
    Ok(Some(cells))
}

/// Run one staged execution: resolve, check the scope and merges, run the
/// shared executor and verify the user-hidden rows.
fn execute(
    source: &Package,
    table: usize,
    rows: Option<RowRange>,
) -> Result<(Package, usize), Error> {
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let selector = BodyTableSelector::index(table);
    let order: Order = source
        .body_table_sort_order(selector)
        .map_err(|_error| Error::InvalidSource)?
        .ok_or(Refusal::NoSortOrder)?;
    match (rows, order.scope()) {
        (None, Scope::EntireTable) | (Some(_), Scope::SelectedRows) => {},
        _ => return Err(Refusal::ScopeMismatch.into()),
    }
    if !source
        .body_table_merges(selector)
        .map_err(|_error| Error::InvalidSource)?
        .is_empty()
    {
        return Err(Unsupported::MergedCells.into());
    }
    let resolved = {
        let mut lock_budget = super::table_lock::WireBudget::new(source.state.source.limits())
            .map_err(|_error| Error::InvalidSource)?;
        source
            .resolve_body_table_with_budget(selector, &mut lock_budget)
            .map_err(|_error| Error::InvalidSource)?
    };
    let target = Target {
        info: resolved.drawable_identifier.get(),
        model: resolved.model_identifier.get(),
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
            let axes = source
                .body_table_hidden_axes(selector)
                .map_err(|_error| engine::Error::InvalidSource)?;
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
        return Ok((source.clone(), 0));
    };
    let candidate = published.package;
    if let Some(hidden) = hidden {
        if candidate
            .body_table_hidden_axes(selector)
            .map_err(|_error| Error::Verification)?
            != hidden
        {
            return Err(Error::Verification);
        }
    }
    Ok((candidate, executed.moved))
}
