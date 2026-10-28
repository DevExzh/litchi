//! Execution of a body table's persisted sort order: the body rows move
//! physically, as Pages' **Sort Now** does (record 0371).
//!
//! The port follows the migration host's scalar executor
//! (`numbers::editor::table_sort::apply`): every sort key must be a complete
//! column of one scalar kind (plain text, finite number, boolean, date or
//! duration); formula, error and unsupported body cells, merged cells,
//! filters, grouping, pivots, spill state, conditional styles and a missing
//! row UID map are refused. The stable row permutation is the shared
//! `litchi_iwa_common::table::sort::planning` planner. One sort rewrites, in
//! place, the tiles' row records (cells keep their bytes), the row header
//! buckets, the row UID map and the stroke layers (borders stay with their
//! cells); user-hidden rows keep their positions; header and footer rows stay
//! put.
//!
//! The payload rewrites below take only native payloads and a permutation,
//! so they can move to the shared table crate unchanged.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::table::sort::planning::{RowPermutation, stable_body_row_permutation};
use litchi_iwa_common::table::sort::{Order, RowRange, Scope};
use litchi_iwa_common::wire::{
    WireView, patch_varint_field, repeated_length_delimited_payloads, repeated_varint_values,
    rewrite_repeated_length_delimited_fields, rewrite_repeated_varint_fields,
};
use litchi_numbers_wire::{BncCellView, CachedScalar, StoredValue};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, MessageRewrite, ObjectIndex,
};
use crate::selector::BodyTableSelector;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const UID_MAP_MESSAGE_TYPES: [u32; 2] = [6_267, 6_200];
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const DATA_LIST_MESSAGE_TYPES: [u32; 2] = [6_005, 6_201];

const MODEL_DATA_STORE_FIELD: u32 = 4;
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
const MODEL_HEADER_ROWS_FIELD: u32 = 9;
const MODEL_FOOTER_ROWS_FIELD: u32 = 11;
const MODEL_HIDDEN_ROWS_FIELD: u32 = 14;
const MODEL_HIDDEN_COLUMNS_FIELD: u32 = 15;
const MODEL_HIDDEN_ROW_OWNER_FIELD: u32 = 35;
const MODEL_HIDDEN_COLUMN_OWNER_FIELD: u32 = 34;
const MODEL_ROW_FILTER_FIELD: u32 = 38;
const MODEL_FILTERED_ROWS_FIELD: u32 = 40;
const MODEL_USER_HIDDEN_ROWS_FIELD: u32 = 41;
const MODEL_USER_HIDDEN_COLUMNS_FIELD: u32 = 42;
const MODEL_UID_MAP_FIELD: u32 = 46;
const MODEL_STROKE_SIDECAR_FIELD: u32 = 49;
const MODEL_PIVOT_OWNER_FIELD: u32 = 85;
const MODEL_CATEGORY_OWNER_FIELD: u32 = 86;
const MODEL_HIDDEN_STATES_OWNER_FIELD: u32 = 70;
const MODEL_DEPRECATED_CATEGORY_OWNER_FIELD: u32 = 81;

const STORE_ROW_HEADERS_FIELD: u32 = 1;
const STORE_TILES_FIELD: u32 = 3;
const STORE_STRING_TABLE_FIELD: u32 = 4;
const STORE_MERGE_MAP_FIELD: u32 = 13;
const STORE_CONDITIONAL_STYLE_FIELD: u32 = 18;
const HEADER_STORAGE_BUCKETS_FIELD: u32 = 2;
const TILE_STORAGE_TILES_FIELD: u32 = 1;
const TILE_STORAGE_SIZE_FIELD: u32 = 2;
const TILE_REFERENCE_ID_FIELD: u32 = 1;
const TILE_REFERENCE_TILE_FIELD: u32 = 2;
const TILE_NUMROWS_FIELD: u32 = 4;
const TILE_ROWS_FIELD: u32 = 5;
const ROW_INDEX_FIELD: u32 = 1;
const ROW_STORAGE_FIELD: u32 = 6;
const ROW_OFFSETS_FIELD: u32 = 7;
const ROW_WIDE_OFFSETS_FIELD: u32 = 8;
const BUCKET_HEADERS_FIELD: u32 = 2;
const HEADER_INDEX_FIELD: u32 = 1;
const UID_ROW_INDEX_FOR_UID_FIELD: u32 = 5;
const UID_ROW_UID_FOR_INDEX_FIELD: u32 = 6;
const UID_SORTED_ROWS_FIELD: u32 = 4;
const LIST_ENTRIES_FIELD: u32 = 3;
const LIST_ENTRY_KEY_FIELD: u32 = 1;
const LIST_ENTRY_STRING_FIELD: u32 = 3;
const HEADER_BUCKET_ROWS: usize = 256;
const DEFAULT_TILE_SIZE: u64 = 256;

/// A finite resource charged by one sort execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payloads or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
        })
    }
}

/// What a sort execution cannot move safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Unsupported {
    /// A body cell holds a formula.
    Formula,
    /// A body cell holds an error or a formula error.
    ErrorCell,
    /// A body cell has a cell type this executor does not move.
    CellType,
    /// A body row has no stored key cell, or a key cell has no scalar.
    MissingKey,
    /// A key column mixes scalar kinds.
    MixedKinds,
    /// Rows are filtered, grouped by category, pivoted or carry spill state.
    FilterOrGrouping,
    /// The table has merged cells.
    MergedCells,
    /// The table has conditional styles.
    ConditionalStyles,
    /// The table has no stable row UID map.
    NoRowUids,
    /// The table has hidden rows or columns.
    HiddenAxes,
    /// The table stores legacy (pre-BNC) rows or an unexpected layout.
    Layout,
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

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

fn wire(error: litchi_iwa_common::Error) -> Error {
    graph::map_wire_error(error).into()
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
            graph::reopen(self, patch.artifacts.target())?
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

/// One stored row of a tile: its local index, raw record and cell slices.
struct TileRow<'a> {
    local: u64,
    raw: &'a [u8],
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

#[derive(Clone, Copy, Debug, PartialEq)]
enum Scalar {
    Number(f64),
    Text(u32),
    Boolean(bool),
    Date(f64),
    Duration(f64),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum ScalarKind {
    Number,
    Text,
    Boolean,
    Date,
    Duration,
}

impl Scalar {
    const fn kind(self) -> ScalarKind {
        match self {
            Self::Number(_) => ScalarKind::Number,
            Self::Text(_) => ScalarKind::Text,
            Self::Boolean(_) => ScalarKind::Boolean,
            Self::Date(_) => ScalarKind::Date,
            Self::Duration(_) => ScalarKind::Duration,
        }
    }

    fn compare(self, other: Self, texts: &HashMap<u32, String>) -> std::cmp::Ordering {
        use std::cmp::Ordering;
        let kinds = self.kind().cmp(&other.kind());
        if kinds != Ordering::Equal {
            return kinds;
        }
        let numbers = |left: f64, right: f64| {
            if left < right {
                Ordering::Less
            } else if left > right {
                Ordering::Greater
            } else {
                Ordering::Equal
            }
        };
        match (self, other) {
            (Self::Number(left), Self::Number(right))
            | (Self::Date(left), Self::Date(right))
            | (Self::Duration(left), Self::Duration(right)) => numbers(left, right),
            (Self::Text(left), Self::Text(right)) => match (texts.get(&left), texts.get(&right)) {
                (Some(left), Some(right)) => left.cmp(right),
                _ => left.cmp(&right),
            },
            (Self::Boolean(left), Self::Boolean(right)) => left.cmp(&right),
            _ => Ordering::Equal,
        }
    }
}

fn movable_scalar(cell: &[u8], key: bool) -> Result<Option<Scalar>, Error> {
    let limits = litchi_numbers_wire::Limits::default();
    let cell = BncCellView::parse(cell, &limits).map_err(|_error| Error::InvalidSource)?;
    if cell.formula_error_identifier().is_some() {
        return Err(Unsupported::ErrorCell.into());
    }
    match cell.stored_value() {
        StoredValue::Formula(_) => return Err(Unsupported::Formula.into()),
        StoredValue::Error => return Err(Unsupported::ErrorCell.into()),
        StoredValue::Unsupported(_) => return Err(Unsupported::CellType.into()),
        _ => {},
    }
    if !key {
        return Ok(None);
    }
    if let StoredValue::Text(identifier) = cell.stored_value() {
        return Ok(Some(Scalar::Text(identifier)));
    }
    Ok(Some(match cell.cached_scalar() {
        Some(CachedScalar::Number(value)) => Scalar::Number(value.get()),
        Some(CachedScalar::Boolean(value)) => Scalar::Boolean(value),
        Some(CachedScalar::Date(value)) => Scalar::Date(value.get()),
        Some(CachedScalar::Duration(value)) => Scalar::Duration(value.get()),
        Some(CachedScalar::Unsupported(_)) => return Err(Unsupported::CellType.into()),
        None => return Err(Unsupported::MissingKey.into()),
    }))
}

/// The strings of a string table data list, by key.
fn strings(payload: &[u8], budget: &mut Budget) -> Result<HashMap<u32, String>, Error> {
    let mut map = HashMap::new();
    for entry in repeated_length_delimited_payloads(payload, LIST_ENTRIES_FIELD).map_err(wire)? {
        let view = budget.parse(entry, 3)?;
        let key = varint(&view, LIST_ENTRY_KEY_FIELD)?.ok_or(Error::InvalidSource)?;
        let Some(text) = graph::unique_field(&view, LIST_ENTRY_STRING_FIELD, 2)? else {
            continue;
        };
        let text = std::str::from_utf8(text.payload()).map_err(|_error| Error::InvalidSource)?;
        map.insert(
            u32::try_from(key).map_err(|_error| Error::InvalidSource)?,
            text.to_owned(),
        );
    }
    Ok(map)
}

/// Whether a data list holds entries.
fn list_has_entries(payload: &[u8]) -> Result<bool, Error> {
    Ok(
        !repeated_length_delimited_payloads(payload, LIST_ENTRIES_FIELD)
            .map_err(wire)?
            .is_empty(),
    )
}

/// The destination row of `row` under the plan (rows outside the sorted
/// range keep their index).
fn relocated(row: usize, start: usize, destinations: &[usize]) -> usize {
    row.checked_sub(start)
        .and_then(|offset| destinations.get(offset).copied())
        .unwrap_or(row)
}

/// The destination of every source row of a permutation.
fn destinations_by_source(permutation: &RowPermutation, start: usize) -> Result<Vec<usize>, Error> {
    let rows = permutation.len();
    let mut destinations = vec![usize::MAX; rows];
    for (destination, source) in permutation.sources_by_destination().iter().enumerate() {
        let source = source.get();
        if destinations.get(source) != Some(&usize::MAX) {
            return Err(Error::InvalidSource);
        }
        destinations[source] = start + destination;
    }
    if destinations.contains(&usize::MAX) {
        return Err(Error::InvalidSource);
    }
    Ok(destinations)
}

/// A tile payload with its rows relocated (the host's
/// `reorder_body_table_tile_rows`, one tile at a time): every row record
/// moves to the tile and local index of its destination.
fn tile_payload(
    original: &[u8],
    rows: &[(usize, Vec<u8>)],
    tile_size: usize,
) -> Result<Vec<u8>, Error> {
    let numrows = rows
        .iter()
        .map(|(row, _)| (row % tile_size) as u64 + 1)
        .max()
        .unwrap_or(0);
    let data =
        patch_varint_field(original, TILE_NUMROWS_FIELD, true, Some(numrows)).map_err(wire)?;
    let records = rows.iter().map(|(_, raw)| raw.clone()).collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(&data, TILE_ROWS_FIELD, &records).map_err(wire)
}

#[allow(clippy::too_many_lines)]
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

    // The model.
    let mut budget = Budget::new(source)?;
    let index = ObjectIndex::new(source, &mut budget)?;
    let catalog = source
        .body_tables()
        .map_err(|_error| Error::InvalidSource)?;
    let model_identifier = {
        let mut lock_budget = super::table_lock::WireBudget::new(source.state.source.limits())
            .map_err(|_error| Error::InvalidSource)?;
        let target = source
            .resolve_body_table_with_budget(selector, &mut lock_budget)
            .map_err(|_error| Error::InvalidSource)?;
        target.model_identifier
    };
    let _ = catalog;
    let (_, _, _, model) = located_message(&index, model_identifier, &[TABLE_MODEL_MESSAGE_TYPE])?;
    let model_view = budget.parse(model, 1)?;
    let total_rows =
        usize::try_from(varint(&model_view, MODEL_ROWS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    let columns =
        usize::try_from(varint(&model_view, MODEL_COLUMNS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    let header_rows = varint(&model_view, MODEL_HEADER_ROWS_FIELD)?.unwrap_or(0) as usize;
    let footer_rows = varint(&model_view, MODEL_FOOTER_ROWS_FIELD)?.unwrap_or(0) as usize;
    let body_start = header_rows;
    let body_end = total_rows
        .checked_sub(footer_rows)
        .filter(|end| *end >= body_start)
        .ok_or(Error::InvalidSource)?;
    let (start, end) = match rows {
        None => (body_start, body_end),
        Some(range) => {
            if range.end() > body_end - body_start {
                return Err(Refusal::RowsOutOfRange.into());
            }
            (body_start + range.start(), body_start + range.end())
        },
    };
    if order
        .rules()
        .iter()
        .any(|rule| rule.column().get() >= columns)
    {
        return Err(Error::InvalidSource);
    }
    if end.saturating_sub(start) < 2 {
        return Ok((source.clone(), 0));
    }

    // Feature refusals (the host's `validate_sort_features`).
    let drawable = {
        let mut lock_budget = super::table_lock::WireBudget::new(source.state.source.limits())
            .map_err(|_error| Error::InvalidSource)?;
        source
            .resolve_body_table_with_budget(selector, &mut lock_budget)
            .map_err(|_error| Error::InvalidSource)?
            .drawable_identifier
    };
    if varint(&model_view, MODEL_FILTERED_ROWS_FIELD)?.unwrap_or(0) != 0
        || graph::unique_field(&model_view, MODEL_PIVOT_OWNER_FIELD, 2)?.is_some()
        || deprecated_grouping(&model_view, &mut budget)?
        || filter_has_rows(&index, &model_view, &mut budget)?
        || grouping_enabled(&index, &model_view, &mut budget)?
        || spill_state(source, drawable, &mut budget)?
    {
        return Err(Unsupported::FilterOrGrouping.into());
    }
    let declares_hidden = graph::unique_field(&model_view, MODEL_HIDDEN_STATES_OWNER_FIELD, 2)?
        .is_some()
        || graph::unique_field(&model_view, MODEL_HIDDEN_ROW_OWNER_FIELD, 2)?.is_some()
        || graph::unique_field(&model_view, MODEL_HIDDEN_COLUMN_OWNER_FIELD, 2)?.is_some()
        || [
            MODEL_HIDDEN_ROWS_FIELD,
            MODEL_HIDDEN_COLUMNS_FIELD,
            MODEL_USER_HIDDEN_ROWS_FIELD,
            MODEL_USER_HIDDEN_COLUMNS_FIELD,
        ]
        .into_iter()
        .map(|field| varint(&model_view, field))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .any(|count| count.unwrap_or(0) != 0);
    // User-hidden rows and columns keep their positions in a native sort:
    // their header records move with the rows and the positions are
    // restored afterwards, as the host does.
    let hidden = if declares_hidden {
        Some(
            source
                .body_table_hidden_axes(selector)
                .map_err(|_error| Error::InvalidSource)?,
        )
        .filter(|hidden| !hidden.is_empty())
    } else {
        None
    };
    let sidecar = reference(&model_view, MODEL_STROKE_SIDECAR_FIELD, &mut budget)?;
    let uid_map =
        reference(&model_view, MODEL_UID_MAP_FIELD, &mut budget)?.ok_or(Unsupported::NoRowUids)?;
    let store = graph::unique_field(&model_view, MODEL_DATA_STORE_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let store_view = budget.parse(store, 2)?;
    if graph::unique_field(&store_view, STORE_MERGE_MAP_FIELD, 2)?.is_some() {
        // A merge map without regions is accepted by the host only when it
        // is empty; the region read above proved there are none.
    }
    if let Some(conditional) = reference(&store_view, STORE_CONDITIONAL_STYLE_FIELD, &mut budget)? {
        let (_, _, _, payload) = located_message(&index, conditional, &DATA_LIST_MESSAGE_TYPES)?;
        if list_has_entries(payload)? {
            return Err(Unsupported::ConditionalStyles.into());
        }
    }

    // Tiles.
    let tiles = graph::unique_field(&store_view, STORE_TILES_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let tiles_view = budget.parse(tiles, 3)?;
    let tile_size = varint(&tiles_view, TILE_STORAGE_SIZE_FIELD)?.unwrap_or(DEFAULT_TILE_SIZE);
    let tile_size = usize::try_from(tile_size)
        .ok()
        .filter(|size| *size != 0)
        .ok_or(Error::InvalidSource)?;
    let mut tile_objects = Vec::new();
    for reference_payload in
        repeated_length_delimited_payloads(tiles, TILE_STORAGE_TILES_FIELD).map_err(wire)?
    {
        let view = budget.parse(reference_payload, 4)?;
        let key = varint(&view, TILE_REFERENCE_ID_FIELD)?.ok_or(Error::InvalidSource)?;
        let tile = reference(&view, TILE_REFERENCE_TILE_FIELD, &mut budget)?
            .ok_or(Error::InvalidSource)?;
        tile_objects.push((
            usize::try_from(key).map_err(|_error| Error::InvalidSource)?,
            tile,
        ));
    }
    let tile_keys = tile_objects
        .iter()
        .map(|(key, _)| *key)
        .collect::<HashSet<_>>();
    if tile_keys.len() != tile_objects.len() {
        return Err(Error::InvalidSource);
    }

    // Keys of every body row in the range.
    let rule_columns = order
        .rules()
        .iter()
        .map(|rule| rule.column().get())
        .collect::<Vec<_>>();
    let mut keys: Vec<Option<Vec<Scalar>>> = vec![None; end - start];
    let mut tile_rows: Vec<(usize, NonZeroU64, Vec<TileRow<'_>>)> = Vec::new();
    for (key, tile) in &tile_objects {
        let (_, _, _, payload) = located_message(&index, *tile, &[TILE_MESSAGE_TYPE])?;
        let mut rows_of_tile = Vec::new();
        let mut locals = HashSet::new();
        for raw in repeated_length_delimited_payloads(payload, TILE_ROWS_FIELD).map_err(wire)? {
            let view = budget.parse(raw, 3)?;
            let local = varint(&view, ROW_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
            if local >= tile_size as u64 || !locals.insert(local) {
                return Err(Error::InvalidSource);
            }
            let global = key * tile_size + local as usize;
            if global >= total_rows {
                return Err(Error::InvalidSource);
            }
            if (start..end).contains(&global) {
                let cells = row_cells(raw, &mut budget)?.ok_or(Unsupported::Layout)?;
                if cells.len() > columns {
                    return Err(Error::InvalidSource);
                }
                let mut values = vec![None; rule_columns.len()];
                for (column, cell) in cells.into_iter().enumerate() {
                    let Some(cell) = cell else {
                        continue;
                    };
                    let position = rule_columns.iter().position(|rule| *rule == column);
                    let scalar = movable_scalar(cell, position.is_some())?;
                    if let (Some(position), Some(scalar)) = (position, scalar) {
                        values[position] = Some(scalar);
                    }
                }
                let values = values
                    .into_iter()
                    .collect::<Option<Vec<_>>>()
                    .ok_or(Unsupported::MissingKey)?;
                if keys[global - start].replace(values).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
            rows_of_tile.push(TileRow { local, raw });
        }
        tile_rows.push((*key, *tile, rows_of_tile));
    }
    let keys = keys
        .into_iter()
        .collect::<Option<Vec<_>>>()
        .ok_or(Unsupported::MissingKey)?;
    let texts = if keys
        .iter()
        .flatten()
        .any(|key| matches!(key, Scalar::Text(_)))
    {
        let table = reference(&store_view, STORE_STRING_TABLE_FIELD, &mut budget)?
            .ok_or(Error::InvalidSource)?;
        let (_, _, _, payload) = located_message(&index, table, &DATA_LIST_MESSAGE_TYPES)?;
        strings(payload, &mut budget)?
    } else {
        HashMap::new()
    };
    if keys.iter().flatten().any(|key| match key {
        Scalar::Text(identifier) => !texts.contains_key(identifier),
        _ => false,
    }) {
        return Err(Error::InvalidSource);
    }
    let permutation = stable_body_row_permutation(
        &keys,
        &order,
        |scalar| scalar.kind(),
        |left, right| left.compare(*right, &texts),
    )
    .map_err(|error| match error {
        litchi_iwa_common::table::sort::planning::Error::MixedScalarKinds { .. } => {
            Error::from(Unsupported::MixedKinds)
        },
        litchi_iwa_common::table::sort::planning::Error::Allocation { rows } => {
            Error::Allocation { amount: rows }
        },
        _ => Error::InvalidSource,
    })?;
    if !permutation.changes_rows() {
        return Ok((source.clone(), 0));
    }
    let destinations = destinations_by_source(&permutation, start)?;
    let moved = destinations
        .iter()
        .enumerate()
        .filter(|(offset, destination)| start + offset != **destination)
        .count();

    let mut rewrites = Vec::new();

    // Tile rows.
    let mut relocated_rows: Vec<(usize, Vec<u8>)> = Vec::new();
    for (key, _, rows_of_tile) in &tile_rows {
        for row in rows_of_tile {
            let source_row = key * tile_size + row.local as usize;
            let destination = relocated(source_row, start, &destinations);
            let raw = patch_varint_field(
                row.raw,
                ROW_INDEX_FIELD,
                true,
                Some((destination % tile_size) as u64),
            )
            .map_err(wire)?;
            relocated_rows.push((destination, raw));
        }
    }
    relocated_rows.sort_by_key(|(row, _)| *row);
    if relocated_rows.windows(2).any(|pair| pair[0].0 == pair[1].0)
        || relocated_rows
            .iter()
            .any(|(row, _)| !tile_keys.contains(&(row / tile_size)))
    {
        return Err(Unsupported::Layout.into());
    }
    for (key, tile, _) in &tile_rows {
        let (component, message_index, message_type, payload) =
            located_message(&index, *tile, &[TILE_MESSAGE_TYPE])?;
        let rows_here = relocated_rows
            .iter()
            .filter(|(row, _)| row / tile_size == *key)
            .cloned()
            .collect::<Vec<_>>();
        let data = tile_payload(payload, &rows_here, tile_size)?;
        if data != payload {
            rewrites.push(MessageRewrite {
                component,
                object: *tile,
                message_index,
                message_type,
                data,
            });
        }
    }

    // Stroke layers: explicit cell borders stay with their cells.
    if let Some(sidecar) = sidecar {
        rewrites.extend(stroke_rewrites(
            &index,
            sidecar,
            total_rows,
            start,
            &destinations,
            &mut budget,
        )?);
    }

    // Row header buckets.
    let row_headers = graph::unique_field(&store_view, STORE_ROW_HEADERS_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let mut buckets = Vec::new();
    for bucket in repeated_length_delimited_payloads(row_headers, HEADER_STORAGE_BUCKETS_FIELD)
        .map_err(wire)?
    {
        buckets.push(graph::parse_reference(bucket, &mut budget)?);
    }
    let mut headers: Vec<(usize, Vec<u8>)> = Vec::new();
    for bucket in &buckets {
        let (_, _, _, payload) = located_message(&index, *bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        for raw in
            repeated_length_delimited_payloads(payload, BUCKET_HEADERS_FIELD).map_err(wire)?
        {
            let view = budget.parse(raw, 3)?;
            let row =
                usize::try_from(varint(&view, HEADER_INDEX_FIELD)?.ok_or(Error::InvalidSource)?)
                    .map_err(|_error| Error::InvalidSource)?;
            if row >= total_rows {
                return Err(Error::InvalidSource);
            }
            let destination = relocated(row, start, &destinations);
            headers.push((
                destination,
                patch_varint_field(raw, HEADER_INDEX_FIELD, true, Some(destination as u64))
                    .map_err(wire)?,
            ));
        }
    }
    headers.sort_by_key(|(row, _)| *row);
    if headers.windows(2).any(|pair| pair[0].0 == pair[1].0)
        || headers
            .iter()
            .any(|(row, _)| row / HEADER_BUCKET_ROWS >= buckets.len())
    {
        return Err(Unsupported::Layout.into());
    }
    for (position, bucket) in buckets.iter().enumerate() {
        let (component, message_index, message_type, payload) =
            located_message(&index, *bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let desired = headers
            .iter()
            .filter(|(row, _)| row / HEADER_BUCKET_ROWS == position)
            .map(|(_, raw)| raw.clone())
            .collect::<Vec<_>>();
        let data =
            rewrite_repeated_length_delimited_fields(payload, BUCKET_HEADERS_FIELD, &desired)
                .map_err(wire)?;
        if data != payload {
            rewrites.push(MessageRewrite {
                component,
                object: *bucket,
                message_index,
                message_type,
                data,
            });
        }
    }

    // The row UID map.
    let physical_row_uids: Vec<Vec<u8>>;
    {
        let (component, message_index, message_type, payload) =
            located_message(&index, uid_map, &UID_MAP_MESSAGE_TYPES)?;
        let uid_for_index =
            repeated_varint_values(payload, UID_ROW_UID_FOR_INDEX_FIELD).map_err(wire)?;
        let index_for_uid =
            repeated_varint_values(payload, UID_ROW_INDEX_FOR_UID_FIELD).map_err(wire)?;
        let sorted = repeated_length_delimited_payloads(payload, UID_SORTED_ROWS_FIELD)
            .map_err(wire)?
            .len();
        if uid_for_index.len() != total_rows
            || index_for_uid.len() != total_rows
            || sorted != total_rows
        {
            return Err(Unsupported::NoRowUids.into());
        }
        let mut current = uid_for_index.clone();
        for (offset, destination) in destinations.iter().copied().enumerate() {
            current[destination] = uid_for_index[start + offset];
        }
        let mut inverse = vec![0_u64; total_rows];
        for (row, uid) in current.iter().copied().enumerate() {
            let uid = usize::try_from(uid)
                .ok()
                .filter(|uid| *uid < total_rows)
                .ok_or(Error::InvalidSource)?;
            inverse[uid] = row as u64;
        }
        let sorted_uids =
            repeated_length_delimited_payloads(payload, UID_SORTED_ROWS_FIELD).map_err(wire)?;
        physical_row_uids = current
            .iter()
            .map(|uid| {
                usize::try_from(*uid)
                    .ok()
                    .and_then(|uid| sorted_uids.get(uid))
                    .map(|raw| raw.to_vec())
                    .ok_or(Error::InvalidSource)
            })
            .collect::<Result<Vec<_>, _>>()?;
        let data = rewrite_repeated_varint_fields(payload, UID_ROW_INDEX_FOR_UID_FIELD, &inverse)
            .map_err(wire)?;
        let data = rewrite_repeated_varint_fields(&data, UID_ROW_UID_FOR_INDEX_FIELD, &current)
            .map_err(wire)?;
        rewrites.push(MessageRewrite {
            component,
            object: uid_map,
            message_index,
            message_type,
            data,
        });
    }
    if let Some(hidden) = &hidden {
        let rows = hidden
            .iter()
            .filter_map(|axis| match axis {
                litchi_iwa_common::table::axis::AxisIndex::Row(index) => Some(index),
                litchi_iwa_common::table::axis::AxisIndex::Column(_) => None,
            })
            .collect::<HashSet<_>>();
        let (component, message_index, message_type, model) =
            located_message(&index, model_identifier, &[TABLE_MODEL_MESSAGE_TYPE])?;
        let (_, _, _, info) = located_message(&index, drawable, &[TABLE_INFO_MESSAGE_TYPE])?;
        if let Some(data) =
            restored_hidden_rows(model, info, &physical_row_uids, &rows, &mut budget)?
        {
            rewrites.push(MessageRewrite {
                component,
                object: model_identifier,
                message_index,
                message_type,
                data,
            });
        }
    }
    let candidate = graph::publish(source, rewrites)?;
    if let Some(hidden) = hidden {
        if candidate
            .body_table_hidden_axes(selector)
            .map_err(|_error| Error::Verification)?
            != hidden
        {
            return Err(Error::Verification);
        }
    }
    Ok((candidate, moved))
}

const FILTER_ENABLED_FIELD: u32 = 2;
const FILTER_RULES_PREPIVOT_FIELD: u32 = 3;
const FILTER_ENABLED_FLAGS_FIELD: u32 = 6;
const FILTER_RULES_FIELD: u32 = 7;
const CATEGORY_GROUP_BY_FIELD: u32 = 1;
const DEPRECATED_CATEGORY_GROUP_BY_FIELD: u32 = 2;
const GROUP_BY_ENABLED_FIELD: u32 = 6;
const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
const OWNER_FORMULA_OWNER_FIELD: u32 = 11;
const OWNER_SPILL_SIZES_FIELD: u32 = 16;
const SPILL_SIZES_SPILLS_FIELD: u32 = 1;

/// Whether the model's pre-pivot row filter set has row state.
fn filter_has_rows(
    index: &ObjectIndex<'_>,
    model: &WireView<'_>,
    budget: &mut Budget,
) -> Result<bool, Error> {
    let Some(filter) = reference(model, MODEL_ROW_FILTER_FIELD, budget)? else {
        return Ok(false);
    };
    let located = index.get(filter)?;
    let payload = located
        .object
        .messages
        .first()
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let view = budget.parse(payload, 2)?;
    let enabled = varint(&view, FILTER_ENABLED_FIELD)?.unwrap_or(1) != 0;
    let flags = repeated_varint_values(payload, FILTER_ENABLED_FLAGS_FIELD).map_err(wire)?;
    Ok(enabled
        && (view.fields().any(|field| {
            matches!(
                field.number(),
                FILTER_RULES_PREPIVOT_FIELD | FILTER_RULES_FIELD
            )
        }) || flags.iter().any(|flag| *flag != 0)))
}

/// Whether the category owner enables a group-by.
fn grouping_enabled(
    index: &ObjectIndex<'_>,
    model: &WireView<'_>,
    budget: &mut Budget,
) -> Result<bool, Error> {
    let Some(owner) = reference(model, MODEL_CATEGORY_OWNER_FIELD, budget)? else {
        return Ok(false);
    };
    let located = index.get(owner)?;
    let payload = located
        .object
        .messages
        .first()
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    for group in
        repeated_length_delimited_payloads(payload, CATEGORY_GROUP_BY_FIELD).map_err(wire)?
    {
        let group = graph::parse_reference(group, budget)?;
        let located = index.get(group)?;
        for message in &located.object.messages {
            let view = budget.parse(&message.data, 2)?;
            if varint(&view, GROUP_BY_ENABLED_FIELD)?.unwrap_or(0) != 0 {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// Whether the deprecated inline category owner enables a group-by.
fn deprecated_grouping(model: &WireView<'_>, budget: &mut Budget) -> Result<bool, Error> {
    let Some(owner) = graph::unique_field(model, MODEL_DEPRECATED_CATEGORY_OWNER_FIELD, 2)? else {
        return Ok(false);
    };
    for group in
        repeated_length_delimited_payloads(owner.payload(), DEPRECATED_CATEGORY_GROUP_BY_FIELD)
            .map_err(wire)?
    {
        let view = budget.parse(group, 3)?;
        if varint(&view, GROUP_BY_ENABLED_FIELD)?.unwrap_or(0) != 0 {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Whether a formula owner of the table drawable holds spill sizes.
fn spill_state(source: &Package, drawable: NonZeroU64, budget: &mut Budget) -> Result<bool, Error> {
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE {
                    continue;
                }
                let view = budget.parse(&message.data, 2)?;
                if reference(&view, OWNER_FORMULA_OWNER_FIELD, budget)? != Some(drawable) {
                    continue;
                }
                if let Some(spills) = graph::unique_field(&view, OWNER_SPILL_SIZES_FIELD, 2)? {
                    if !repeated_length_delimited_payloads(
                        spills.payload(),
                        SPILL_SIZES_SPILLS_FIELD,
                    )
                    .map_err(wire)?
                    .is_empty()
                    {
                        return Ok(true);
                    }
                }
            }
        }
    }
    Ok(false)
}

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const INFO_HIDDEN_STATES_UUID_FIELD: u32 = 8;
const OWNER_HIDDEN_STATES_FIELD: u32 = 2;
const STATES_UID_FIELD: u32 = 1;
const STATES_ROW_EXTENT_FIELD: u32 = 3;
const EXTENT_BASE_STATES_FIELD: u32 = 2;
const STATE_UID_FIELD: u32 = 1;
const STATE_USER_HIDDEN_FIELD: u32 = 2;
const STATE_FILTERED_FIELD: u32 = 3;
const STATE_PIVOT_HIDDEN_FIELD: u32 = 4;

/// The model payload with user-hidden rows back at `rows` after a row
/// permutation (the host's `restore_positional_user_hidden_axes`, rows
/// only): the active row extent's states are re-keyed by the rows' new
/// physical UIDs. `None` when nothing changes.
fn restored_hidden_rows(
    model: &[u8],
    info: &[u8],
    physical_uids: &[Vec<u8>],
    rows: &HashSet<usize>,
    budget: &mut Budget,
) -> Result<Option<Vec<u8>>, Error> {
    let model_view = budget.parse(model, 1)?;
    let Some(owner) = graph::unique_field(&model_view, MODEL_HIDDEN_STATES_OWNER_FIELD, 2)? else {
        return Err(Unsupported::HiddenAxes.into());
    };
    let owner = owner.payload();
    let info_view = budget.parse(info, 1)?;
    let active = graph::unique_field(&info_view, INFO_HIDDEN_STATES_UUID_FIELD, 2)?
        .ok_or(Unsupported::HiddenAxes)?
        .payload();
    let states =
        repeated_length_delimited_payloads(owner, OWNER_HIDDEN_STATES_FIELD).map_err(wire)?;
    if states.len() != 1 {
        return Err(Unsupported::HiddenAxes.into());
    }
    let state_view = budget.parse(states[0], 2)?;
    if graph::unique_field(&state_view, STATES_UID_FIELD, 2)?.map(|field| field.payload())
        != Some(active)
    {
        return Err(Error::InvalidSource);
    }
    let extent = graph::unique_field(&state_view, STATES_ROW_EXTENT_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let positions = physical_uids
        .iter()
        .enumerate()
        .map(|(index, uid)| (uid.as_slice(), index))
        .collect::<HashMap<_, _>>();
    let flag = |view: &WireView<'_>, number| -> Result<bool, Error> {
        Ok(varint(view, number)?.is_some_and(|value| value != 0))
    };
    let mut present = HashSet::new();
    let mut kept: Vec<(usize, Vec<u8>)> = Vec::new();
    for state in
        repeated_length_delimited_payloads(extent, EXTENT_BASE_STATES_FIELD).map_err(wire)?
    {
        let view = budget.parse(state, 3)?;
        let uid = graph::unique_field(&view, STATE_UID_FIELD, 2)?
            .ok_or(Error::InvalidSource)?
            .payload();
        if flag(&view, STATE_FILTERED_FIELD)? || flag(&view, STATE_PIVOT_HIDDEN_FIELD)? {
            return Err(Unsupported::FilterOrGrouping.into());
        }
        let Some(position) = positions.get(uid).copied() else {
            kept.push((usize::MAX, state.to_vec()));
            continue;
        };
        if rows.contains(&position) {
            present.insert(position);
            kept.push((
                position,
                patch_varint_field(
                    state,
                    STATE_USER_HIDDEN_FIELD,
                    varint(&view, STATE_USER_HIDDEN_FIELD)?.is_some(),
                    Some(1),
                )
                .map_err(wire)?,
            ));
        }
    }
    let mut desired = rows.iter().copied().collect::<Vec<_>>();
    desired.sort_unstable();
    for position in desired {
        if present.contains(&position) {
            continue;
        }
        let uid = physical_uids.get(position).ok_or(Error::InvalidSource)?;
        let mut state = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut state,
            STATE_UID_FIELD,
            uid,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(wire)?;
        litchi_iwa_common::wire::append_varint_field(&mut state, STATE_USER_HIDDEN_FIELD, 1)
            .map_err(wire)?;
        kept.push((position, state));
    }
    kept.sort_by_key(|(position, _)| *position);
    let records = kept.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    let new_extent =
        rewrite_repeated_length_delimited_fields(extent, EXTENT_BASE_STATES_FIELD, &records)
            .map_err(wire)?;
    if new_extent == extent {
        return Ok(None);
    }
    let new_state = litchi_iwa_common::wire::patch_length_delimited_field(
        states[0],
        STATES_ROW_EXTENT_FIELD,
        true,
        Some(&new_extent),
    )
    .map_err(wire)?;
    let new_owner =
        rewrite_repeated_length_delimited_fields(owner, OWNER_HIDDEN_STATES_FIELD, &[new_state])
            .map_err(wire)?;
    litchi_iwa_common::wire::patch_length_delimited_field(
        model,
        MODEL_HIDDEN_STATES_OWNER_FIELD,
        true,
        Some(&new_owner),
    )
    .map(Some)
    .map_err(wire)
}

const SIDECAR_LEFT_COLUMN_FIELD: u32 = 4;
const SIDECAR_TOP_ROW_FIELD: u32 = 6;
const SIDECAR_BOTTOM_ROW_FIELD: u32 = 7;
const STROKE_LAYER_MESSAGE_TYPE: u32 = 6_306;
const LAYER_INDEX_FIELD: u32 = 1;
const LAYER_RUNS_FIELD: u32 = 2;
const RUN_ORIGIN_FIELD: u32 = 1;
const RUN_LENGTH_FIELD: u32 = 2;

/// One run of a column layer split into the destination ranges of its rows.
fn split_run(
    run: &[u8],
    rows: usize,
    start: usize,
    destinations: &[usize],
    budget: &mut Budget,
) -> Result<Vec<Vec<u8>>, Error> {
    let view = budget.parse(run, 3)?;
    let origin = varint(&view, RUN_ORIGIN_FIELD)?.ok_or(Error::InvalidSource)?;
    let length = varint(&view, RUN_LENGTH_FIELD)?.ok_or(Error::InvalidSource)?;
    let origin = usize::try_from(origin).map_err(|_error| Error::InvalidSource)?;
    let length = usize::try_from(length).map_err(|_error| Error::InvalidSource)?;
    if length == 0 || origin.checked_add(length).is_none_or(|end| end > rows) {
        return Err(Error::InvalidSource);
    }
    let mut moved = (origin..origin + length)
        .map(|row| relocated(row, start, destinations))
        .collect::<Vec<_>>();
    moved.sort_unstable();
    let mut ranges = Vec::new();
    let mut first = moved[0];
    let mut last = moved[0];
    for destination in moved.into_iter().skip(1) {
        if last + 1 == destination {
            last = destination;
            continue;
        }
        ranges.push((first, last));
        first = destination;
        last = destination;
    }
    ranges.push((first, last));
    ranges
        .into_iter()
        .map(|(first, last)| {
            let patched = patch_varint_field(run, RUN_ORIGIN_FIELD, true, Some(first as u64))
                .map_err(wire)?;
            patch_varint_field(
                &patched,
                RUN_LENGTH_FIELD,
                true,
                Some((last - first + 1) as u64),
            )
            .map_err(wire)
        })
        .collect()
}

/// The stroke-layer rewrites of a row permutation (the host's
/// `stroke_layers::reorder_rows`): a row layer moves with its row, a column
/// layer's runs split into the destination ranges of their rows.
fn stroke_rewrites(
    index: &ObjectIndex<'_>,
    sidecar: NonZeroU64,
    rows: usize,
    start: usize,
    destinations: &[usize],
    budget: &mut Budget,
) -> Result<Vec<MessageRewrite>, Error> {
    let (_, _, _, payload) = located_message(index, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
    let mut rewrites = Vec::new();
    let mut seen = HashSet::new();
    for field in SIDECAR_LEFT_COLUMN_FIELD..=SIDECAR_BOTTOM_ROW_FIELD {
        for layer in repeated_length_delimited_payloads(payload, field).map_err(wire)? {
            let layer = graph::parse_reference(layer, budget)?;
            if !seen.insert(layer) {
                return Err(Error::InvalidSource);
            }
            let (component, message_index, message_type, data) =
                located_message(index, layer, &[STROKE_LAYER_MESSAGE_TYPE])?;
            let rewritten = if matches!(field, SIDECAR_TOP_ROW_FIELD | SIDECAR_BOTTOM_ROW_FIELD) {
                let view = budget.parse(data, 2)?;
                let current = varint(&view, LAYER_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
                let current = usize::try_from(current)
                    .ok()
                    .filter(|current| *current <= rows)
                    .ok_or(Error::InvalidSource)?;
                let moved = relocated(current, start, destinations);
                patch_varint_field(data, LAYER_INDEX_FIELD, true, Some(moved as u64))
                    .map_err(wire)?
            } else {
                let mut runs = Vec::new();
                for run in
                    repeated_length_delimited_payloads(data, LAYER_RUNS_FIELD).map_err(wire)?
                {
                    runs.extend(split_run(run, rows, start, destinations, budget)?);
                }
                rewrite_repeated_length_delimited_fields(data, LAYER_RUNS_FIELD, &runs)
                    .map_err(wire)?
            };
            if rewritten != data {
                rewrites.push(MessageRewrite {
                    component,
                    object: layer,
                    message_index,
                    message_type,
                    data: rewritten,
                });
            }
        }
    }
    Ok(rewrites)
}
