//! Implementation of the public [`crate::table::topology`] transactions:
//! inserting and removing table rows and columns (`IWORK_API_REQUESTS.md`
//! request 6, record 0309).
//!
//! An axis edit moves every cell after the edited index. One edit rewrites,
//! in one staged working set over the exact source:
//!
//! - the stored rows of every storage tile (a row edit renumbers the rows
//!   after the index; a column edit splices each row's current and pre-BNC
//!   cell-offset tables, so cell payloads keep their bytes);
//! - the row or column header buckets (sizes, hidden states and styles move
//!   with their row or column);
//! - the row/column UID map (an inserted axis gets a fresh identity above
//!   every existing one, as the host editor allocates it);
//! - the stroke sidecar's border layers (layers on the edited axis move,
//!   perpendicular runs shift, split or shorten);
//! - the table model's row or column count and, for a header or footer
//!   section edit, its section count;
//! - the table formula owner's table extents (the drawable's layout cache
//!   is left for the application to rebuild, as the host editor leaves it);
//! - every formula coordinate of the calculation engine that addresses the
//!   table ([`shift`], record 0318): the table's own formulas and their
//!   dependency records, and other tables' formulas that read it.
//!
//! A removed row or column first has its comments cleared through the
//! comment transaction and its stored cells cleared through the
//! [`crate::table::cells`] transaction (value, style and list references are
//! released there, and the caches of formulas that read them refreshed), so
//! the structural pass only drops empty storage. Where the cell writer
//! cannot clear them, the structural pass drops the cells itself, with
//! their strings, formulas and dependency records (record 0329). A
//! surviving formula that reads a removed cell reads `#REF!` there and
//! caches an error result; a range that loses a row or column shrinks; a
//! shared formula whose cells diverge is split into one entry per formula.
//!
//! Ported from the host editor's `row_insert`, `column_insert`,
//! `table_delete` and `formula_dependency_shift` phases. Refused with a
//! typed [`Unsupported`] reason so a failed edit never publishes: a chart or
//! other non-table owner reading the table, volatile, spanning, error or
//! spill dependency state, a stale-cache range over cells the cell writer
//! cannot clear, a removed range among several of one tract, shared
//! formulas that diverge in a segmented list, merged cells, filters,
//! categories, pivots, column edits under sort rules, removing a hidden row
//! or column, and removing a row or column that carries a border layer.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

mod shift;
pub(super) mod wire;

pub(super) use litchi_numbers_wire::table_topology::{extents as calc, lists, storage};

use std::collections::BTreeMap;
use std::fmt;

use litchi_iwa_core::RawMessage;
use litchi_numbers_wire::ListKind;
use thiserror::Error as ThisError;

use self::storage::{AxisEdit, HEADER_BUCKET_ROWS, StrokePlan, UidMap};
use self::wire::{repeated_bytes, single_bytes, single_reference, single_varint};
use super::Package;
use super::structure::{
    self, DOCUMENT_SHEETS_FIELD, Fault, FaultResult, Published, Working, map_core,
    metadata::Transition, references,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;
use crate::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};

structure::structural_family!("table-topology");

/// `TST.Tile`.
pub(super) const TILE_MESSAGE_TYPE: u32 = 6_002;
/// `TST.HeaderStorageBucket`.
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// `TST.TableDataList`.
pub(super) const LIST_MESSAGE_TYPE: u32 = 6_005;
/// `TST.TableDataListSegment`.
pub(super) const LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
/// `TST.FilterSetArchive`.
const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;
/// `TST.ColumnRowUIDMapArchive`.
const UID_MAP_MESSAGE_TYPES: [u32; 2] = [6_200, 6_267];
/// `TST.StrokeSidecarArchive`.
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
/// `TST.StrokeLayerArchive`.
const STROKE_LAYER_MESSAGE_TYPE: u32 = 6_306;
/// `TST.SummaryModelArchive`.
const SUMMARY_MODEL_MESSAGE_TYPE: u32 = 6_316;
/// `TST.CategoryOwnerRefArchive`.
const CATEGORY_OWNER_MESSAGE_TYPE: u32 = 6_372;
/// `TST.GroupByArchive`.
const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;
/// The most header or footer rows (or header columns) Numbers allows.
const MAX_SECTION: u32 = 5;
/// The rows of a storage tile when the tile storage does not say.
const DEFAULT_TILE_SIZE: u64 = 256;

/// Why a table-topology edit is refused before anything is published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Unsupported {
    /// Calculation-engine state the coordinate shift does not rewrite:
    /// volatile, spanning, whole-owner, error or spill dependencies, a
    /// category reference whose host moves, a packed coordinate, a
    /// dependency record that would leave its tile, or (record 0329) a
    /// range over a removed row or column whose cells the cell writer could
    /// not clear, so its cached result would go stale.
    Formulas,
    /// A chart or another non-table owner reads the table, or a formula
    /// reads it through legacy UUID ranges or UUID references.
    ReferencedByOtherOwners,
    /// The table has merged cells.
    Merges,
    /// The table has filtered rows or an active filter.
    Filters,
    /// The table has active categories (group-by) or a summary model.
    Categories,
    /// The table is, or feeds, a pivot table.
    Pivot,
    /// A column edit on a table with sort rules, or any edit on a table
    /// sorted by a row range.
    SortRules,
    /// The removed row or column is hidden.
    HiddenAxis,
    /// The removed row or column carries a border layer.
    Borders,
    /// A comment on the removed row or column that the comment transaction
    /// cannot clear.
    Comments,
    /// The table's storage layout is outside what this transaction rewrites
    /// (a missing UID map, a row that would cross a storage tile, a cell
    /// the clear left behind).
    StorageLayout,
    /// The removed row or column is one of several ranges of one colon
    /// tract a formula reads, or an endpoint of a merge range. (A single
    /// reference to it, or a range wholly on it, reads `#REF!`, and a range
    /// that loses it shrinks: record 0329.)
    DeletedReference,
    /// One formula shared by several cells would become different formulas
    /// in a segmented formula list (an unsegmented one is split: record
    /// 0329).
    SharedFormulas,
}

impl fmt::Display for Unsupported {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Formulas => "calculation dependencies the coordinate shift does not rewrite",
            Self::ReferencedByOtherOwners => "charts or other owners elsewhere that read the table",
            Self::Merges => "merged cells",
            Self::Filters => "filters",
            Self::Categories => "categories",
            Self::Pivot => "pivot tables",
            Self::SortRules => "sort rules",
            Self::HiddenAxis => "a hidden row or column",
            Self::Borders => "a border layer on the removed row or column",
            Self::Comments => "a commented cell on the removed row or column",
            Self::StorageLayout => "this table storage layout",
            Self::DeletedReference => "a formula that reads the removed row or column",
            Self::SharedFormulas => "a formula shared by cells that would diverge",
        })
    }
}

/// A typed, content-redacted table-topology failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table-topology source is invalid")]
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
    #[error("could not allocate {amount} table-topology items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("table-topology candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-topology patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The section-relative index is outside its section.
    #[error("index {index} is outside a section of {length}")]
    PositionOutOfRange {
        /// Requested section-relative index.
        index: usize,
        /// The section's length (an insertion may equal it).
        length: usize,
    },
    /// A header or footer section would exceed Numbers' five rows or
    /// columns.
    #[error("a header or footer section holds at most five rows or columns")]
    SectionFull,
    /// The edit would remove the table's last row or column, or a body row
    /// or column its header and footer sections need.
    #[error("the table must keep at least one row and one column")]
    LastAxis,
    /// The edit touches table content this transaction does not rewrite.
    #[error("table topology edits are not supported for {reason}")]
    Unsupported {
        /// What the table holds.
        reason: Unsupported,
    },
    /// Clearing the removed row's or column's stored cells failed.
    #[error("clearing the removed cells failed: {error:?}")]
    CellClear {
        /// The cell transaction's error.
        error: crate::table::cells::Error,
    },
    /// Commit was requested before staging an operation.
    #[error("no table-topology operation is staged")]
    NoStagedOperation,
    /// A surviving object depends on the edited graph.
    #[error("a surviving object depends on the table's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support table topology changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::LastAxis { .. } => RefusalKind::UnsupportedTarget,
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::PositionOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::SectionFull { .. }
            | Self::Unsupported { .. }
            | Self::CellClear { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

const fn refuse(reason: Unsupported) -> Error {
    Error::Refused(Refusal::Unsupported { reason })
}

/// A storage-phase fault: an unowned layout is a typed storage refusal.
fn storage_fault(fault: impl Into<Fault>) -> Error {
    match fault.into() {
        Fault::Unsupported => refuse(Unsupported::StorageLayout),
        other => other.into(),
    }
}

/// One staged axis edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operation {
    InsertRow(RowInsertion),
    InsertColumn(ColumnInsertion),
    RemoveRow(RowDeletion),
    RemoveColumn(ColumnDeletion),
}

/// A table's size and section counts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Geometry {
    pub(super) rows: u32,
    pub(super) columns: u32,
    pub(super) header_rows: u32,
    pub(super) header_columns: u32,
    pub(super) footer_rows: u32,
}

impl Geometry {
    fn extents(self) -> calc::Extents {
        calc::Extents::of(
            self.rows,
            self.columns,
            self.header_rows,
            self.header_columns,
            self.footer_rows,
        )
    }
}

/// One operation resolved against a geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Step {
    row_axis: bool,
    edit: AxisEdit,
    after: Geometry,
    /// A body-row insertion (footer ranges ending on the row before it
    /// grow).
    expand_footer: bool,
}

fn usize_of(value: u32) -> usize {
    usize::try_from(value).unwrap_or(usize::MAX)
}

fn checked(index: usize, length: u32, inclusive: bool) -> Result<u32, Error> {
    let fits = if inclusive {
        index <= usize_of(length)
    } else {
        index < usize_of(length)
    };
    if !fits {
        return Err(Error::Refused(Refusal::PositionOutOfRange {
            index,
            length: usize_of(length),
        }));
    }
    u32::try_from(index).map_err(|_error| {
        Error::Refused(Refusal::PositionOutOfRange {
            index,
            length: usize_of(length),
        })
    })
}

fn grow(count: u32) -> Result<u32, Error> {
    if count >= MAX_SECTION {
        return Err(Error::Refused(Refusal::SectionFull));
    }
    Ok(count + 1)
}

fn resolve(operation: Operation, geometry: Geometry) -> Result<Step, Error> {
    let mut after = geometry;
    let body_rows = geometry
        .rows
        .checked_sub(geometry.header_rows + geometry.footer_rows)
        .ok_or(Error::InvalidSource)?;
    let body_columns = geometry
        .columns
        .checked_sub(geometry.header_columns)
        .ok_or(Error::InvalidSource)?;
    let (row_axis, edit) = match operation {
        Operation::InsertRow(insertion) => {
            after.rows = geometry.rows.checked_add(1).ok_or(Error::InvalidSource)?;
            let physical = match insertion {
                RowInsertion::Header { index } => {
                    after.header_rows = grow(geometry.header_rows)?;
                    checked(index, geometry.header_rows, true)?
                },
                RowInsertion::Body { index } => {
                    geometry.header_rows + checked(index, body_rows, true)?
                },
                RowInsertion::Footer { index } => {
                    after.footer_rows = grow(geometry.footer_rows)?;
                    geometry.rows - geometry.footer_rows
                        + checked(index, geometry.footer_rows, true)?
                },
            };
            (true, AxisEdit::Insert(physical))
        },
        Operation::RemoveRow(deletion) => {
            after.rows = geometry
                .rows
                .checked_sub(1)
                .ok_or(Error::Refused(Refusal::LastAxis))?;
            let physical = match deletion {
                RowDeletion::Header { index } => {
                    let index = checked(index, geometry.header_rows, false)?;
                    after.header_rows -= 1;
                    index
                },
                RowDeletion::Body { index } => {
                    geometry.header_rows + checked(index, body_rows, false)?
                },
                RowDeletion::Footer { index } => {
                    let index = checked(index, geometry.footer_rows, false)?;
                    after.footer_rows -= 1;
                    geometry.rows - geometry.footer_rows + index
                },
            };
            if after.rows == 0 || after.header_rows + after.footer_rows > after.rows {
                return Err(Error::Refused(Refusal::LastAxis));
            }
            (true, AxisEdit::Remove(physical))
        },
        Operation::InsertColumn(insertion) => {
            after.columns = geometry
                .columns
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            let physical = match insertion {
                ColumnInsertion::Header { index } => {
                    after.header_columns = grow(geometry.header_columns)?;
                    checked(index, geometry.header_columns, true)?
                },
                ColumnInsertion::Body { index } => {
                    geometry.header_columns + checked(index, body_columns, true)?
                },
            };
            (false, AxisEdit::Insert(physical))
        },
        Operation::RemoveColumn(deletion) => {
            after.columns = geometry
                .columns
                .checked_sub(1)
                .ok_or(Error::Refused(Refusal::LastAxis))?;
            let physical = match deletion {
                ColumnDeletion::Header { index } => {
                    let index = checked(index, geometry.header_columns, false)?;
                    after.header_columns -= 1;
                    index
                },
                ColumnDeletion::Body { index } => {
                    geometry.header_columns + checked(index, body_columns, false)?
                },
            };
            if after.columns == 0 || after.header_columns > after.columns {
                return Err(Error::Refused(Refusal::LastAxis));
            }
            (false, AxisEdit::Remove(physical))
        },
    };
    Ok(Step {
        row_axis,
        edit,
        after,
        expand_footer: matches!(operation, Operation::InsertRow(RowInsertion::Body { .. })),
    })
}

/// One immutable edit that stages row and column insertions and removals
/// in one table, applied in order.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    geometry: Geometry,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, operation: Operation) -> Result<Self, Error> {
        let step = resolve(operation, self.geometry)?;
        self.operations
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.operations.push(operation);
        self.geometry = step.after;
        Ok(self)
    }

    /// Stage a blank row inserted at a section-relative position. An index
    /// equal to the section's length appends to it; a header or footer
    /// insertion grows that section.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_row(self, insertion: RowInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertRow(insertion))
    }

    /// Stage a blank column inserted at a section-relative position.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_column(self, insertion: ColumnInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertColumn(insertion))
    }

    /// Stage the removal of the row at a section-relative position, with
    /// its cells.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_row(self, deletion: RowDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveRow(deletion))
    }

    /// Stage the removal of the column at a section-relative position,
    /// with its cells.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_column(self, deletion: ColumnDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveColumn(deletion))
    }

    /// Validate and atomically publish the staged operations, in order.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::Unsupported`] for table content this transaction does not
    /// rewrite; [`Refusal::CellClear`] when a removed cell cannot be cleared;
    /// source, resource and verification errors. Failure never publishes a
    /// partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        let mut previews = 0_usize;
        for operation in &self.operations {
            let base = current.as_ref().unwrap_or(self.source);
            let published = run(base, self.sheet, self.table, *operation)?;
            touched = touched.saturating_add(published.touched_components);
            previews = previews.saturating_add(published.deleted_previews);
            current = Some(published.package);
        }
        let package = current.ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: previews,
            },
        )
    }
}

/// The table drawable and model of one table.
pub(super) struct Target {
    pub(super) info: u64,
    pub(super) model: u64,
}

pub(super) fn locate(working: &Working<'_>, sheet: usize, table: usize) -> Result<Target, Error> {
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let sheet = *sheets
        .get(sheet)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let mut tables = Vec::new();
    for drawable in references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)? {
        if structure::table_info_message(working.object(drawable)?).is_ok() {
            tables.push(drawable);
        }
    }
    let info = *tables
        .get(table)
        .ok_or(Error::Refused(Refusal::TableNotFound))?;
    let info_object = working.object(info)?;
    let info_index = structure::table_info_message(info_object)?;
    let model =
        single_reference(&info_object.messages[info_index].data, 2)?.ok_or(Error::InvalidSource)?;
    structure::unique_message(working.object(model)?, &[6_001])?;
    Ok(Target { info, model })
}

/// The unique message of `identifier` with one of `types`.
pub(super) fn message<'w>(
    working: &'w Working<'_>,
    identifier: u64,
    types: &[u32],
) -> FaultResult<(usize, &'w [u8])> {
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, types)?;
    Ok((index, &object.messages[index].data))
}

pub(super) fn geometry(model: &[u8]) -> Result<Geometry, Error> {
    let count = |number| -> Result<u32, Error> {
        u32::try_from(single_varint(model, number)?.unwrap_or(0))
            .map_err(|_error| Error::InvalidSource)
    };
    let geometry = Geometry {
        rows: u32::try_from(single_varint(model, 6)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
        columns: u32::try_from(single_varint(model, 7)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
        header_rows: count(9)?,
        header_columns: count(10)?,
        footer_rows: count(11)?,
    };
    if geometry.header_rows + geometry.footer_rows > geometry.rows
        || geometry.header_columns > geometry.columns
    {
        return Err(Error::InvalidSource);
    }
    Ok(geometry)
}

fn read_geometry(source: &Package, sheet: usize, table: usize) -> Result<Geometry, Error> {
    let working = Working::new(source)?;
    let target = locate(&working, sheet, table)?;
    let (_, model) = message(&working, target.model, &[6_001])?;
    geometry(model)
}

/// Whether a `TSP.UUID` payload names `(lower, upper)`.
fn uuid_is(payload: &[u8], uuid: (u64, u64)) -> FaultResult<bool> {
    Ok(single_varint(payload, 1)? == Some(uuid.0) && single_varint(payload, 2)? == Some(uuid.1))
}

/// Refuse the table features this transaction does not rewrite.
fn check_features(
    working: &Working<'_>,
    target: &Target,
    model: &[u8],
    info: &[u8],
    step: Step,
) -> Result<(), Error> {
    if single_varint(model, 40)?.unwrap_or(0) != 0 {
        return Err(refuse(Unsupported::Filters));
    }
    if let Some(filter) = single_reference(model, 38)? {
        let (_, data) = message(working, filter, &[FILTER_SET_MESSAGE_TYPE])?;
        let enabled = single_varint(data, 2)?.unwrap_or(1) != 0;
        let rules = !repeated_bytes(data, 3)?.is_empty() || !repeated_bytes(data, 7)?.is_empty();
        let toggles = wire::repeated_varints(data, 6)?
            .iter()
            .any(|value| *value != 0);
        if enabled && (rules || toggles) {
            return Err(refuse(Unsupported::Filters));
        }
    }
    if single_bytes(model, 85)?.is_some()
        || single_bytes(info, 15)?.is_some()
        || single_varint(info, 16)?.unwrap_or(0) != 0
        || single_bytes(info, 17)?.is_some()
    {
        return Err(refuse(Unsupported::Pivot));
    }
    if single_bytes(model, 83)?.is_some_and(|uid| !uid.is_empty())
        || single_bytes(info, 6)?.is_some()
    {
        return Err(refuse(Unsupported::Categories));
    }
    if let Some(deprecated) = single_bytes(model, 81)? {
        for group in repeated_bytes(deprecated, 2)? {
            if single_varint(group, 6)?.unwrap_or(0) != 0 {
                return Err(refuse(Unsupported::Categories));
            }
        }
    }
    if let Some(owner) = single_reference(model, 86)? {
        let (_, data) = message(working, owner, &[CATEGORY_OWNER_MESSAGE_TYPE])?;
        for reference in repeated_bytes(data, 1)? {
            let group = references::reference_identifier(reference)?;
            let (_, group) = message(working, group, &[GROUP_BY_MESSAGE_TYPE])?;
            if single_varint(group, 6)?.unwrap_or(0) != 0 {
                return Err(refuse(Unsupported::Categories));
            }
        }
    }
    if let Some(summary) = single_reference(info, 4)? {
        let (_, data) = message(working, summary, &[SUMMARY_MODEL_MESSAGE_TYPE])?;
        if let Some(uids) = single_reference(data, 3)? {
            let (_, uid_map) = message(working, uids, &UID_MAP_MESSAGE_TYPES)?;
            if !uid_map.is_empty() {
                return Err(refuse(Unsupported::Categories));
            }
        }
    }
    let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;
    if let Some(merges) = single_reference(store, 13)? {
        // `TST.MergeRegionMapArchive`, the object's only message.
        let object = working.object(merges)?;
        let [map] = object.messages.as_slice() else {
            return Err(Error::InvalidSource);
        };
        if !repeated_bytes(&map.data, 1)?.is_empty() {
            return Err(refuse(Unsupported::Merges));
        }
    }
    if let Some(owner) = single_bytes(model, 47)?
        && let Some(formulas) = single_bytes(owner, 2)?
        && !repeated_bytes(formulas, 3)?.is_empty()
    {
        return Err(refuse(Unsupported::Merges));
    }
    if let Some(sort) = single_bytes(model, 44)? {
        if single_varint(sort, 1)?.unwrap_or(0) != 0 {
            return Err(refuse(Unsupported::SortRules));
        }
        if !step.row_axis && !repeated_bytes(sort, 2)?.is_empty() {
            return Err(refuse(Unsupported::SortRules));
        }
    }
    let _ = target;
    Ok(())
}

/// Refuse removing a row or column that a hidden state names.
fn check_hidden(model: &[u8], uuid: (u64, u64), row_axis: bool) -> Result<(), Error> {
    let Some(owner) = single_bytes(model, 70)? else {
        return Ok(());
    };
    for state in repeated_bytes(owner, 2)? {
        let extent = single_bytes(state, if row_axis { 3 } else { 2 })?;
        let Some(extent) = extent else { continue };
        for base in repeated_bytes(extent, 2)? {
            if let Some(identity) = single_bytes(base, 1)?
                && uuid_is(identity, uuid)?
            {
                return Err(refuse(Unsupported::HiddenAxis));
            }
        }
    }
    Ok(())
}

/// The storage tiles of a data store: `(key, tile object)`, and the rows
/// per tile.
pub(super) fn tiles(store: &[u8]) -> Result<(Vec<(u32, u64)>, u64), Error> {
    let storage = single_bytes(store, 3)?.ok_or(Error::InvalidSource)?;
    let size = single_varint(storage, 2)?.unwrap_or(DEFAULT_TILE_SIZE);
    if size == 0 {
        return Err(Error::InvalidSource);
    }
    let mut output = Vec::new();
    for entry in repeated_bytes(storage, 1)? {
        let key = u32::try_from(single_varint(entry, 1)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
        let tile = single_reference(entry, 2)?.ok_or(Error::InvalidSource)?;
        if output.iter().any(|(existing, _)| *existing == key) {
            return Err(Error::InvalidSource);
        }
        output
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        output.push((key, tile));
    }
    Ok((output, size))
}

/// The stored cells on the removed row or column, as table positions.
fn cells_on_axis(
    working: &Working<'_>,
    store: &[u8],
    columns: u32,
    row_axis: bool,
    index: u32,
) -> Result<Vec<At>, Error> {
    let (tiles, size) = tiles(store)?;
    let mut output = Vec::new();
    for (key, tile) in tiles {
        let (_, data) = message(working, tile, &[TILE_MESSAGE_TYPE])?;
        for row in repeated_bytes(data, 5)? {
            let local = single_varint(row, 1)?.ok_or(Error::InvalidSource)?;
            let global = u32::try_from(u64::from(key) * size + local)
                .map_err(|_error| Error::InvalidSource)?;
            if row_axis && global != index {
                continue;
            }
            let mut stored = vec![false; usize_of(columns)];
            for number in [4, 7] {
                let Some(offsets) = single_bytes(row, number)? else {
                    continue;
                };
                for (column, pair) in offsets.chunks_exact(2).enumerate() {
                    if pair != [0xff, 0xff] && column < stored.len() {
                        stored[column] = true;
                    }
                }
            }
            for (column, present) in stored.into_iter().enumerate() {
                let column = u32::try_from(column).map_err(|_error| Error::InvalidSource)?;
                if present && (row_axis || column == index) {
                    output
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation { amount: 1 })?;
                    output.push(At::new(global, column));
                }
            }
        }
    }
    Ok(output)
}

/// Clear the comments on a removed row or column through the comment
/// transaction (record 0329); `None` when the axis holds none.
fn clear_comments(
    source: &Package,
    sheet: usize,
    table: usize,
    row_axis: bool,
    index: u32,
) -> Result<Option<Package>, Error> {
    let refused = |_error| refuse(Unsupported::Comments);
    let comments = source
        .table_cell_comments(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(refused)?;
    let mut current: Option<Package> = None;
    for (position, _) in comments {
        let on_axis = if row_axis {
            position.row() == index
        } else {
            position.column() == index
        };
        if !on_axis {
            continue;
        }
        let base = current.as_ref().unwrap_or(source);
        let commit = base
            .clear_table_cell_comment(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(refused)?;
        current = Some(commit.into_package());
    }
    Ok(current)
}

/// Clear the stored cells of a removed row or column through the cell
/// transaction, so value, style and list references are released.
fn clear_axis(
    source: &Package,
    sheet: usize,
    table: usize,
    cells: &[At],
) -> Result<Package, Error> {
    let cell_error = |error| Error::Refused(Refusal::CellClear { error });
    let mut edit = source
        .edit_table_cells(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(cell_error)?;
    for position in cells {
        edit = edit.clear(*position).map_err(cell_error)?;
    }
    Ok(edit.commit().map_err(cell_error)?.into_package())
}

/// The planned rewrites of one operation, and the objects it removes
/// (record 0329: dependency tiles emptied by dropped formula hosts).
struct Planned {
    replacements: Vec<Replacement>,
    removed: Vec<u64>,
}

/// One planned payload replacement.
struct Replacement {
    object: u64,
    index: usize,
    data: Vec<u8>,
    /// Object references the new payload no longer holds.
    pruned: Vec<u64>,
}

/// The data-store list field a released cell key belongs to, or `None` for
/// a key this transaction does not release (values, comments). With
/// `values` (record 0329: the removed cells were not cleared through the
/// cell writer) strings and formula errors are released too; formula keys
/// are counted for the coordinate shift, which owns the formula list.
const fn list_field(kind: ListKind, values: bool) -> Option<u32> {
    match kind {
        ListKind::Style => Some(5),
        ListKind::ConditionalStyle => Some(18),
        ListKind::ControlCellSpec => Some(21),
        ListKind::Format => Some(22),
        ListKind::String if values => Some(4),
        ListKind::FormulaError if values => Some(12),
        ListKind::String
        | ListKind::RichText
        | ListKind::Formula
        | ListKind::FormulaError
        | ListKind::Comment => None,
    }
}

/// The data-store field of a table's string list.
const STRING_LIST_FIELD: u32 = 4;

/// The cells a removal drops after the value clear: their columns (a row
/// removal) or rows (a column removal), and the list keys they hold, per
/// data-store list field.
#[derive(Default)]
struct Released {
    /// Source index on the other axis -> dropped cells there.
    per_index: BTreeMap<u32, u64>,
    /// Data-store list field -> key -> references released.
    keys: BTreeMap<u32, BTreeMap<u32, u32>>,
    /// Formula-list key -> references the dropped cells hold (record 0329).
    formulas: BTreeMap<u64, u64>,
}

fn released_cells(
    working: &Working<'_>,
    store: &[u8],
    row_axis: bool,
    index: u32,
    values: bool,
) -> Result<Released, Error> {
    let (tiles, size) = tiles(store)?;
    let limits = working
        .source()
        .numbers_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut released = Released::default();
    for (key, tile) in tiles {
        let (_, data) = message(working, tile, &[TILE_MESSAGE_TYPE])?;
        for row in repeated_bytes(data, 5)? {
            let local = single_varint(row, 1)?.ok_or(Error::InvalidSource)?;
            let global = u32::try_from(u64::from(key) * size + local)
                .map_err(|_error| Error::InvalidSource)?;
            if row_axis && global != index {
                continue;
            }
            for (column, bytes) in storage::bnc_cells(row).map_err(storage_fault)? {
                if !row_axis && column != index {
                    continue;
                }
                let other = if row_axis { column } else { global };
                *released.per_index.entry(other).or_insert(0) += 1;
                let cell = litchi_numbers_wire::BncCell::parse(bytes, &limits)
                    .map_err(|_error| Error::InvalidSource)?;
                for (kind, list_key) in cell.list_keys() {
                    if values && kind == ListKind::Formula {
                        *released.formulas.entry(u64::from(list_key)).or_insert(0) += 1;
                        continue;
                    }
                    let field =
                        list_field(kind, values).ok_or(refuse(if kind == ListKind::Comment {
                            Unsupported::Comments
                        } else {
                            Unsupported::StorageLayout
                        }))?;
                    *released
                        .keys
                        .entry(field)
                        .or_default()
                        .entry(list_key)
                        .or_insert(0) += 1;
                }
            }
        }
    }
    Ok(released)
}

/// Apply one formula cell patch (record 0329) to the tile of `store` that
/// holds the cell, keeping the rewritten tile in `patched`; returns the
/// string-list key a replaced text cache held.
fn patch_cell(
    working: &Working<'_>,
    store: &[u8],
    columns: u32,
    patch: shift::CellPatch,
    patched: &mut BTreeMap<u64, Vec<u8>>,
) -> Result<Option<u32>, Error> {
    let (tiles, size) = tiles(store)?;
    let limits = working
        .source()
        .numbers_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let key = u32::try_from(u64::from(patch.row) / size).map_err(|_error| Error::InvalidSource)?;
    let local =
        u32::try_from(u64::from(patch.row) % size).map_err(|_error| Error::InvalidSource)?;
    let tile = tiles
        .iter()
        .find(|(candidate, _)| *candidate == key)
        .map(|(_, tile)| *tile)
        .ok_or(Error::InvalidSource)?;
    let current = match patched.get(&tile) {
        Some(data) => data.clone(),
        None => message(working, tile, &[TILE_MESSAGE_TYPE])?.1.to_vec(),
    };
    let layout = |_error| refuse(Unsupported::StorageLayout);
    let bytes = litchi_numbers_wire::table_write::tile_cell(&current, local, patch.column, &limits)
        .map_err(layout)?
        .ok_or(Error::InvalidSource)?;
    let mut cell = litchi_numbers_wire::BncCell::parse(&bytes, &limits)
        .map_err(|_error| Error::InvalidSource)?;
    let mut released = None;
    if let Some(key) = patch.formula_key {
        cell.set_formula_identifier(key)
            .map_err(|_error| Error::InvalidSource)?;
    }
    if patch.error {
        released = cell
            .set_formula_cached_error()
            .map_err(|_error| refuse(Unsupported::StorageLayout))?;
    }
    let cell_bytes = crate::cell::wire::encode(&cell, &limits).map_err(|_| Error::InvalidSource)?;
    let written = litchi_numbers_wire::table_write::tile_set_cell(
        &current,
        local,
        patch.column,
        columns,
        Some(&cell_bytes),
        &limits,
    )
    .map_err(layout)?;
    patched.insert(tile, written.tile);
    Ok(released)
}

/// Plan every payload rewrite of `step` over `working`'s source.
///
/// With `cells_pending`, the removed row or column may still hold the cells
/// the caller clears before the real plan, so the storage tiles are not
/// planned. With `drop_cells` (record 0329), the removed cells were not
/// cleared through the cell writer: their values, strings and formulas are
/// released here.
fn plan(
    working: &Working<'_>,
    target: &Target,
    step: Step,
    before: Geometry,
    cells_pending: bool,
    drop_cells: bool,
) -> Result<Planned, Error> {
    let source = working.source();
    let (model_index, model) = message(working, target.model, &[6_001])?;
    let (_, info) = message(working, target.info, &[6_000, 6_003])?;
    check_features(working, target, model, info, step)?;
    let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;

    // The cells a removal drops with the axis, and their list references.
    let mut released = match step.edit {
        AxisEdit::Remove(index) if !cells_pending => {
            released_cells(working, store, step.row_axis, index, drop_cells)?
        },
        _ => Released::default(),
    };

    // The calculation graph: the formula-coordinate shift (record 0318).
    let (insert, position) = match step.edit {
        AxisEdit::Insert(index) => (true, index),
        AxisEdit::Remove(index) => (false, index),
    };
    let calculation = shift::plan(
        source,
        &shift::Table {
            info: target.info,
            model,
            rows: before.rows,
            footer_rows: before.footer_rows,
            dropped_formulas: &released.formulas,
        },
        shift::Shift {
            row_axis: step.row_axis,
            position,
            insert,
            expand_footer: step.expand_footer,
            pending_clear: cells_pending,
            drop_hosts: drop_cells,
        },
        cells_pending,
    )
    .map_err(|failure| match failure {
        shift::Failure::Fault(fault) => Error::from(fault),
        shift::Failure::Refused(reason) => refuse(reason),
    })?;

    let mut replacements = Vec::new();
    let mut push =
        |object: u64, index: usize, data: Vec<u8>, pruned: Vec<u64>| -> Result<(), Error> {
            replacements
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            replacements.push(Replacement {
                object,
                index,
                data,
                pruned,
            });
            Ok(())
        };

    // Cells the formula rewrite changes (record 0329): error caches of
    // `#REF!` formulas and cells of split shared formulas. The edited
    // table's tiles are patched before they move; other tables' in place.
    let mut patched: BTreeMap<u64, Vec<u8>> = BTreeMap::new();
    let mut readers: BTreeMap<u64, Vec<shift::CellPatch>> = BTreeMap::new();
    for patch in &calculation.cells {
        if patch.info == target.info {
            let strings = patch_cell(working, store, before.columns, *patch, &mut patched)?;
            if let Some(key) = strings {
                *released
                    .keys
                    .entry(STRING_LIST_FIELD)
                    .or_default()
                    .entry(key)
                    .or_insert(0) += 1;
            }
        } else {
            readers.entry(patch.info).or_default().push(*patch);
        }
    }
    for (reader, patches) in &readers {
        let reader_info = message(working, *reader, &[6_000, 6_003])?.1;
        let reader_model = single_reference(reader_info, 2)?.ok_or(Error::InvalidSource)?;
        let (_, reader_model) = message(working, reader_model, &[6_001])?;
        let columns = geometry(reader_model)?.columns;
        let reader_store = single_bytes(reader_model, 4)?.ok_or(Error::InvalidSource)?;
        let mut tiles_patched = BTreeMap::new();
        let mut strings: BTreeMap<u32, u32> = BTreeMap::new();
        for patch in patches {
            if let Some(key) =
                patch_cell(working, reader_store, columns, *patch, &mut tiles_patched)?
            {
                *strings.entry(key).or_insert(0) += 1;
            }
        }
        for (tile, data) in tiles_patched {
            let (index, _) = message(working, tile, &[TILE_MESSAGE_TYPE])?;
            push(tile, index, data, Vec::new())?;
        }
        if !strings.is_empty() {
            let list =
                single_reference(reader_store, STRING_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
            let (index, data) = message(working, list, &[LIST_MESSAGE_TYPE])?;
            let (payload, pruned) =
                lists::release_with(data, &mut strings, true).map_err(storage_fault)?;
            if !strings.is_empty() {
                return Err(refuse(Unsupported::StorageLayout));
            }
            if let Some(payload) = payload {
                push(list, index, payload, pruned)?;
            }
        }
    }
    for (field, keys) in &released.keys {
        let list = single_reference(store, *field)?.ok_or(Error::InvalidSource)?;
        let (index, data) = message(working, list, &[LIST_MESSAGE_TYPE])?;
        let mut remaining = keys.clone();
        let reclaim = drop_cells && (*field == STRING_LIST_FIELD || *field == 12);
        let (payload, pruned) =
            lists::release_with(data, &mut remaining, reclaim).map_err(storage_fault)?;
        if let Some(payload) = payload {
            push(list, index, payload, pruned)?;
        }
        for reference in repeated_bytes(data, 4)? {
            if remaining.is_empty() {
                break;
            }
            let segment = references::reference_identifier(reference)?;
            let (segment_index, segment_data) =
                message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])?;
            let (payload, pruned) =
                lists::release(segment_data, &mut remaining).map_err(storage_fault)?;
            if let Some(payload) = payload {
                push(segment, segment_index, payload, pruned)?;
            }
        }
        if !remaining.is_empty() {
            return Err(refuse(Unsupported::StorageLayout));
        }
    }
    let released = released;
    let released_at = |index: u32| released.per_index.get(&index).copied().unwrap_or(0);
    let none = |_index: u32| 0_u64;

    // The UID map.
    let uid_object = single_reference(model, 46)?.ok_or(refuse(Unsupported::StorageLayout))?;
    let (uid_index, uid_payload) = message(working, uid_object, &UID_MAP_MESSAGE_TYPES)?;
    let uids = UidMap::read(uid_payload, before.rows, before.columns).map_err(storage_fault)?;
    if let AxisEdit::Remove(index) = step.edit {
        let uuid = uids.uuid_at(step.row_axis, index).map_err(Fault::from)?;
        check_hidden(model, uuid, step.row_axis)?;
    }
    push(
        uid_object,
        uid_index,
        uids.rewrite(step.row_axis, step.edit)
            .map_err(storage_fault)?,
        Vec::new(),
    )?;

    // Storage tiles.
    let (tiles, size) = tiles(store)?;
    for (key, tile) in tiles.iter().filter(|_| !cells_pending) {
        let (index, original) = message(working, *tile, &[TILE_MESSAGE_TYPE])?;
        let data = patched.get(tile).map_or(original, Vec::as_slice);
        let rewritten = if step.row_axis {
            storage::tile_rows(
                data,
                *key,
                u32::try_from(size).map_err(|_error| Error::InvalidSource)?,
                step.edit,
            )
        } else {
            storage::tile_columns(data, before.columns, step.edit)
        }
        .map_err(storage_fault)?;
        match rewritten {
            Some(rewritten) => push(*tile, index, rewritten, Vec::new())?,
            None if data != original => push(*tile, index, data.to_vec(), Vec::new())?,
            None => {},
        }
    }

    // Header buckets: the edited axis's headers move (a removed one goes),
    // and the other axis's headers lose the removed cells from their counts.
    let headers = single_bytes(store, 1)?.ok_or(Error::InvalidSource)?;
    for (bucket, reference) in repeated_bytes(headers, 2)?.into_iter().enumerate() {
        let identifier = references::reference_identifier(reference)?;
        let (index, data) = message(working, identifier, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let bucket = u64::try_from(bucket).map_err(|_error| Error::InvalidSource)?;
        let rewritten = if step.row_axis {
            storage::header_bucket(data, bucket, Some(HEADER_BUCKET_ROWS), step.edit, &none)
        } else {
            storage::header_bucket(
                data,
                bucket,
                Some(HEADER_BUCKET_ROWS),
                AxisEdit::Insert(u32::MAX),
                &released_at,
            )
        }
        .map_err(storage_fault)?;
        if let Some(rewritten) = rewritten {
            push(identifier, index, rewritten, Vec::new())?;
        }
    }
    if let Some(identifier) = single_reference(store, 2)? {
        let (index, data) = message(working, identifier, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let rewritten = if step.row_axis {
            storage::header_bucket(data, 0, None, AxisEdit::Insert(u32::MAX), &released_at)
        } else {
            storage::header_bucket(data, 0, None, step.edit, &none)
        }
        .map_err(storage_fault)?;
        if let Some(rewritten) = rewritten {
            push(identifier, index, rewritten, Vec::new())?;
        }
    }

    // Border layers.
    if let Some(sidecar) = single_reference(model, 49)? {
        let (index, data) = message(working, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
        let mut layer_fault = None;
        let StrokePlan {
            sidecar: new_sidecar,
            layers,
        } = storage::strokes(
            data,
            before.rows,
            before.columns,
            step.row_axis,
            step.edit,
            |layer| {
                message(working, layer, &[STROKE_LAYER_MESSAGE_TYPE])
                    .map(|(_, data)| data)
                    .map_err(|fault| {
                        layer_fault.get_or_insert(fault);
                        litchi_numbers_wire::table_topology::Error::Invalid
                    })
            },
        )
        .map_err(
            |error| match layer_fault.take().unwrap_or_else(|| error.into()) {
                Fault::Unsupported => refuse(Unsupported::Borders),
                other => other.into(),
            },
        )?;
        for (layer, payload) in layers {
            let (layer_index, _) = message(working, layer, &[STROKE_LAYER_MESSAGE_TYPE])?;
            push(layer, layer_index, payload, Vec::new())?;
        }
        push(sidecar, index, new_sidecar, Vec::new())?;
    }

    // The model's counts.
    let mut new_model = structure::put_varint(model, 6, u64::from(step.after.rows))?;
    new_model = structure::put_varint(&new_model, 7, u64::from(step.after.columns))?;
    for (number, old, new) in [
        (9, before.header_rows, step.after.header_rows),
        (10, before.header_columns, step.after.header_columns),
        (11, before.footer_rows, step.after.footer_rows),
    ] {
        if old == new {
            continue;
        }
        new_model = if new == 0 {
            structure::strip_fields(&new_model, &[number])?
        } else {
            structure::put_varint(&new_model, number, u64::from(new))?
        };
    }
    push(target.model, model_index, new_model, Vec::new())?;

    // The drawable's layout cache stays as it is (record 0370): the
    // application rebuilds a cache that no longer describes the table, and
    // the host editor, whose output carries the native evidence, leaves it.

    // The calculation engine: the shifted payloads, and the table formula
    // owner's extents.
    if let Some((object, index, payload)) = calculation.own {
        let rewritten = calc::rewrite_extents(
            &payload,
            before.extents(),
            step.after.extents(),
            step.row_axis,
            insert,
        )
        .map_err(|fault| match Fault::from(fault) {
            Fault::Unsupported => refuse(Unsupported::Formulas),
            other => other.into(),
        })?
        .unwrap_or(payload);
        let (_, original) = message(working, object, &[4_008])?;
        if rewritten != original {
            push(object, index, rewritten, calculation.removed.clone())?;
        }
    }
    for rewrite in calculation.rewrites {
        push(rewrite.object, rewrite.message, rewrite.data, Vec::new())?;
    }
    Ok(Planned {
        replacements,
        removed: calculation.removed,
    })
}

fn replace(working: &mut Working<'_>, replacement: Replacement) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(replacement.object)?;
    if !replacement.pruned.is_empty() {
        return references::replace(
            object,
            replacement.index,
            replacement.data,
            &replacement.pruned,
            &[],
            &[],
            limits,
        );
    }
    let message_type = object
        .messages
        .get(replacement.index)
        .ok_or(Fault::Invalid)?
        .type_;
    object
        .replace_message_preserving_header(
            replacement.index,
            RawMessage {
                type_: message_type,
                data: replacement.data,
            },
            limits,
        )
        .map_err(|error| map_core(&error))?;
    Ok(())
}

/// The table's stored cells, by position.
fn table_cells(
    package: &Package,
    sheet: usize,
    table: usize,
) -> Result<Vec<(At, crate::cell::Value)>, Error> {
    let table = package
        .workbook_ref()
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .ok_or(Error::Verification)?;
    let mut output = Vec::new();
    output
        .try_reserve(table.cell_count())
        .map_err(|_error| Error::Allocation {
            amount: table.cell_count(),
        })?;
    for cell in table.iter_cells() {
        output.push((cell.position(), cell.value().clone()));
    }
    Ok(output)
}

/// Run one operation: check, clear a removed axis, rewrite, publish and
/// verify.
fn run(
    source: &Package,
    sheet: usize,
    table: usize,
    operation: Operation,
) -> Result<Published, Error> {
    let before = read_geometry(source, sheet, table)?;
    let step = resolve(operation, before)?;
    // Check everything before clearing anything.
    let cells = {
        let working = Working::new(source)?;
        let target = locate(&working, sheet, table)?;
        plan(&working, &target, step, before, true, false)?;
        match step.edit {
            AxisEdit::Remove(index) => {
                let (_, model) = message(&working, target.model, &[6_001])?;
                let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;
                cells_on_axis(&working, store, before.columns, step.row_axis, index)?
            },
            AxisEdit::Insert(_) => Vec::new(),
        }
    };
    // Comments on the removed axis go through the comment transaction
    // first, which removes their storage graphs (record 0329).
    let uncommented = match step.edit {
        AxisEdit::Remove(index) if !cells.is_empty() => {
            clear_comments(source, sheet, table, step.row_axis, index)?
        },
        _ => None,
    };
    let commented = uncommented.as_ref().unwrap_or(source);
    // The cell writer clears the removed cells and refreshes the caches of
    // the formulas that read them; where it cannot, the structural pass
    // drops them itself (record 0329).
    let (cleared, drop_cells) = if cells.is_empty() {
        (None, false)
    } else {
        match clear_axis(commented, sheet, table, &cells) {
            Ok(package) => (Some(package), false),
            Err(
                error @ Error::Refused(Refusal::CellClear {
                    error: crate::table::cells::Error::Allocation { .. },
                }),
            ) => return Err(error),
            Err(Error::Refused(Refusal::CellClear { .. })) => (None, true),
            Err(other) => return Err(other),
        }
    };
    let base = cleared.as_ref().unwrap_or(commented);
    let mut working = Working::new(base)?;
    working.keep_previews();
    let target = locate(&working, sheet, table)?;
    let planned = plan(&working, &target, step, before, false, drop_cells)?;
    for replacement in planned.replacements {
        replace(&mut working, replacement)?;
    }
    let mut transition = Transition::default();
    if !planned.removed.is_empty() {
        if let Some(registry) = structure::metadata::registry(&working)? {
            transition.remove_objects(&registry, &planned.removed)?;
        }
        for identifier in &planned.removed {
            working.remove(*identifier)?;
        }
    }
    let published = working.publish(&transition)?;
    verify(source, &published.package, sheet, table, step)?;
    Ok(published)
}

/// The candidate has the new size and every surviving cell at its shifted
/// position with its value.
fn verify(
    source: &Package,
    candidate: &Package,
    sheet: usize,
    table: usize,
    step: Step,
) -> Result<(), Error> {
    let candidate_table = candidate
        .workbook_ref()
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .ok_or(Error::Verification)?;
    if (candidate_table.row_count(), candidate_table.column_count())
        != (step.after.rows, step.after.columns)
    {
        return Err(Error::Verification);
    }
    if read_geometry(candidate, sheet, table)? != step.after {
        return Err(Error::Verification);
    }
    let mut expected = Vec::new();
    for (position, value) in table_cells(source, sheet, table)? {
        let (row, column) = if step.row_axis {
            match step.edit.shift(position.row()).map_err(Fault::from)? {
                Some(row) => (row, position.column()),
                None => continue,
            }
        } else {
            match step.edit.shift(position.column()).map_err(Fault::from)? {
                Some(column) => (position.row(), column),
                None => continue,
            }
        };
        expected
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        expected.push((At::new(row, column), value));
    }
    let mut actual = table_cells(candidate, sheet, table)?;
    let key = |cell: &(At, crate::cell::Value)| (cell.0.row(), cell.0.column());
    expected.sort_by_key(key);
    actual.sort_by_key(key);
    // A formula's text follows the shifted coordinates: only its presence is
    // compared here.
    for cell in expected.iter_mut().chain(actual.iter_mut()) {
        if let crate::cell::Value::Formula(text) = &mut cell.1 {
            text.clear();
        }
    }
    if expected != actual {
        return Err(Error::Verification);
    }
    Ok(())
}

pub(super) fn table_position(
    source: &Package,
    sheet: usize,
    selector: TableSelector<'_>,
) -> Result<usize, Error> {
    let sheet = source
        .workbook_ref()
        .sheets()
        .get(sheet)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    match selector {
        TableSelector::Position(position) => {
            let position = position.get();
            if position >= sheet.table_count() {
                return Err(Error::Refused(Refusal::TableNotFound));
            }
            Ok(position)
        },
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_, table)| table.name() == name);
            let (position, _) = matches
                .next()
                .ok_or(Error::Refused(Refusal::TableNotFound))?;
            if matches.next().is_some() {
                return Err(Error::InvalidSource);
            }
            Ok(position)
        },
    }
}

impl Package {
    /// Start an immutable row and column insertion and removal edit on one
    /// table.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::table::topology::{RowDeletion, RowInsertion};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package
    ///     .edit_table_topology(0, 0)?
    ///     .insert_row(RowInsertion::body(0))?
    ///     .remove_row(RowDeletion::body(3))?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_table_topology(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::TableNotFound`]; source errors
    /// for a semantic-only or malformed source.
    pub fn edit_table_topology<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = table_position(self, sheet, table.into())?;
        let geometry = read_geometry(self, sheet, table)?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            geometry,
            operations: Vec::new(),
        })
    }

    /// Apply a table-topology patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_topology(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

#[cfg(test)]
mod tests {
    use litchi_numbers_wire::{BncCell, CachedScalar};

    use super::{TILE_MESSAGE_TYPE, locate, message, tiles};
    use crate::formula::Expression;
    use crate::package::structure::Working;
    use crate::table::cells::Input;
    use crate::table::topology::RowDeletion;
    use crate::{At, Package};

    /// The stored BNC cell at `position` of the first table.
    fn stored(package: &Package, position: At) -> BncCell {
        let working = Working::new(package).expect("working set");
        let target = locate(&working, 0, 0).expect("table");
        let (_, model) = message(&working, target.model, &[6_001]).expect("model");
        let store = super::single_bytes(model, 4)
            .expect("store")
            .expect("store");
        let (tiles, size) = tiles(store).expect("tiles");
        let key = u32::try_from(u64::from(position.row()) / size).expect("key");
        let local = u32::try_from(u64::from(position.row()) % size).expect("row");
        let tile = tiles
            .iter()
            .find(|(candidate, _)| *candidate == key)
            .expect("tile")
            .1;
        let (_, data) = message(&working, tile, &[TILE_MESSAGE_TYPE]).expect("tile payload");
        let limits = crate::cell::wire::limits_for_source(data);
        let bytes =
            litchi_numbers_wire::table_write::tile_cell(data, local, position.column(), &limits)
                .expect("cell")
                .expect("stored");
        BncCell::parse(&bytes, &crate::cell::wire::limits_for_source(&bytes)).expect("cell")
    }

    /// Record 0329: a formula whose reference is removed reads `#REF!` and
    /// caches an error result.
    #[test]
    fn deleted_references_cache_an_error() {
        let base = Package::new();
        let mut edit = base.edit_table_cells(0, 0).expect("edit");
        edit = edit
            .set_a1("B2", Input::number(4.0).expect("number"))
            .expect("value")
            .set_a1("B3", Input::number(5.0).expect("number"))
            .expect("value");
        edit = edit
            .set_formula_a1("E3", Expression::parse("B2+B3").expect("formula"))
            .expect("formula");
        let package = edit.commit().expect("cells").into_package();
        let before = stored(&package, At::new(2, 4));
        assert!(matches!(
            before.cached_scalar(),
            Ok(Some(CachedScalar::Number(_)))
        ));
        let removed = package
            .edit_table_topology(0, 0)
            .expect("topology")
            .remove_row(RowDeletion::body(0))
            .expect("row")
            .commit()
            .expect("commit");
        let after = stored(removed.package(), At::new(1, 4));
        assert!(matches!(
            after.cached_scalar(),
            Ok(Some(CachedScalar::Unsupported(8)))
        ));
        assert!(matches!(
            after.stored_value(),
            litchi_numbers_wire::StoredValue::Formula(_)
        ));
    }
}
