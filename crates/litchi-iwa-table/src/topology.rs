//! The shared table-topology engine: inserting and removing one table row
//! or column (record 0309, lifted from `litchi-numbers` for Pages and Keynote
//! in record 0412e).
//!
//! An axis edit moves every cell after the edited index. [`publish`]
//! rewrites, in one staged working set over the exact source:
//!
//! - the stored rows of every storage tile (a row edit renumbers the rows
//!   after the index; a column edit splices each row's current and pre-BNC
//!   cell-offset tables, so cell payloads keep their bytes);
//! - the row or column header buckets (sizes, hidden states and styles move
//!   with their row or column);
//! - the row/column UID map (an inserted axis gets a fresh identity above
//!   every existing one, as the host editor allocates it);
//! - the stroke sidecar's border layers;
//! - the table model's row or column count and, for a header or footer
//!   section edit, its section count;
//! - the table formula owner's table extents;
//! - every formula coordinate of the calculation engine that addresses the
//!   table (`shift`, record 0318).
//!
//! A leaf locates the table ([`Target`]), resolves the operation against its
//! geometry ([`resolve`]), runs the deletion-safe preflight ([`check`]) and
//! clears what its own transactions own (comments, cells a formula reads)
//! before [`publish`]. Cells a removal drops without a clear (`drop_cells`)
//! have their list references, strings and formulas released here.
//!
//! Ported from the host editor's `row_insert`, `column_insert`,
//! `table_delete` and `formula_dependency_shift` phases. Refused with a
//! typed [`Unsupported`] reason so a failed edit never publishes.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    reason = "the engine keeps its phases in plan order"
)]

mod shift;

use std::collections::BTreeMap;
use std::fmt;

use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_common::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};
use litchi_iwa_core::RawMessage;
use litchi_numbers_wire::ListKind;
pub(crate) use litchi_numbers_wire::table_topology::{extents as calc, lists, storage};

use self::storage::{AxisEdit, HEADER_BUCKET_ROWS, StrokePlan, UidMap};
use crate::TableSource;
use crate::structure::wire::{self, repeated_bytes, single_bytes, single_reference, single_varint};
use crate::structure::{
    self, Fault, FaultResult, Published, TableLeaf, Working, map_core, metadata::Transition,
    references,
};

/// `TST.Tile`.
pub(crate) const TILE_MESSAGE_TYPE: u32 = 6_002;
/// `TST.HeaderStorageBucket`.
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// `TST.TableDataList`.
pub(crate) const LIST_MESSAGE_TYPE: u32 = 6_005;
/// `TST.TableDataListSegment`.
pub(crate) const LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
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

/// A typed table-topology engine failure; a leaf maps it into its family's
/// public error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request the engine refuses.
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    InvalidSource,
    /// A structural working-set fault (limits, allocation, verification).
    Fault(Fault),
    /// A bounded allocation failed.
    Allocation {
        /// Requested amount.
        amount: usize,
    },
}

/// A well-formed request the engine refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The section-relative index is outside its section.
    PositionOutOfRange {
        /// Requested section-relative index.
        index: usize,
        /// The section's length (an insertion may equal it).
        length: usize,
    },
    /// A header or footer section would exceed five rows or columns.
    SectionFull,
    /// The edit would remove the table's last row or column, or a body row
    /// or column its header and footer sections need.
    LastAxis,
    /// The edit touches table content this engine does not rewrite.
    Unsupported {
        /// What the table holds.
        reason: Unsupported,
    },
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Invalid => Self::InvalidSource,
            Fault::Allocation { amount } => Self::Allocation { amount },
            other => Self::Fault(other),
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// The checked BNC cell limits of a source's ingress profile.
fn wire_limits(source: &dyn TableSource) -> Result<litchi_numbers_wire::Limits, Error> {
    let maximum = source
        .archive_limits()
        .effective_archive_limits()
        .map_err(|_error| Error::InvalidSource)?
        .max_message_bytes();
    let wire = litchi_iwa_common::WireLimits::default()
        .with_input_bytes(maximum)
        .and_then(|limits| limits.with_output_bytes(maximum))
        .map_err(|_error| Error::InvalidSource)?;
    Ok(litchi_numbers_wire::Limits::from_wire_limits(wire))
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
pub enum Operation {
    /// Insert one blank row.
    InsertRow(RowInsertion),
    /// Insert one blank column.
    InsertColumn(ColumnInsertion),
    /// Remove one row with its cells.
    RemoveRow(RowDeletion),
    /// Remove one column with its cells.
    RemoveColumn(ColumnDeletion),
}

/// A table's size and section counts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Geometry {
    /// Rows.
    pub rows: u32,
    /// Columns.
    pub columns: u32,
    /// Header rows.
    pub header_rows: u32,
    /// Header columns.
    pub header_columns: u32,
    /// Footer rows.
    pub footer_rows: u32,
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
pub struct Step {
    row_axis: bool,
    edit: AxisEdit,
    after: Geometry,
    /// A body-row insertion (footer ranges ending on the row before it
    /// grow).
    expand_footer: bool,
}

impl Step {
    /// The geometry after the step.
    #[must_use]
    pub const fn after(self) -> Geometry {
        self.after
    }

    /// Whether the step edits a row (else a column).
    #[must_use]
    pub const fn row_axis(self) -> bool {
        self.row_axis
    }

    /// Where the cell at physical `index` of the edited axis lands after
    /// the step (`None` for the removed index).
    ///
    /// # Errors
    ///
    /// An index the step cannot move.
    pub fn shifted(self, index: u32) -> Result<Option<u32>, Error> {
        self.edit
            .shift(index)
            .map_err(|error| Error::from(Fault::from(error)))
    }

    /// The physical index a removal drops, if the step removes.
    #[must_use]
    pub const fn removed(self) -> Option<u32> {
        match self.edit {
            AxisEdit::Remove(index) => Some(index),
            AxisEdit::Insert(_) => None,
        }
    }
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

/// Resolve one operation against a table geometry.
///
/// # Errors
///
/// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`],
/// [`Refusal::LastAxis`], or a malformed geometry.
pub fn resolve(operation: Operation, geometry: Geometry) -> Result<Step, Error> {
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

/// The table drawable and model of one table.
pub struct Target {
    /// The table drawable (`TST.TableInfoArchive`).
    pub info: u64,
    /// The table model (`TST.TableModelArchive`).
    pub model: u64,
}

/// The unique message of `identifier` with one of `types`.
pub(crate) fn message<'w, P: TableLeaf>(
    working: &'w Working<'_, P>,
    identifier: u64,
    types: &[u32],
) -> FaultResult<(usize, &'w [u8])> {
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, types)?;
    Ok((index, &object.messages[index].data))
}

/// A table model payload's geometry.
///
/// # Errors
///
/// A malformed or inconsistent model.
pub fn geometry(model: &[u8]) -> Result<Geometry, Error> {
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

/// Whether a `TSP.UUID` payload names `(lower, upper)`.
fn uuid_is(payload: &[u8], uuid: (u64, u64)) -> FaultResult<bool> {
    Ok(single_varint(payload, 1)? == Some(uuid.0) && single_varint(payload, 2)? == Some(uuid.1))
}

/// Refuse the table features this transaction does not rewrite.
fn check_features<P: TableLeaf>(
    working: &Working<'_, P>,
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
/// The storage tiles of a data store: `(key, tile object)`, and the rows
/// per tile.
///
/// # Errors
///
/// A malformed tile storage.
pub fn tiles(store: &[u8]) -> Result<(Vec<(u32, u64)>, u64), Error> {
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
fn cells_on_axis<P: TableLeaf>(
    working: &Working<'_, P>,
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

/// `released` (a root list after its releases) with `next_list_id` handed
/// back over the keys whose entries the releases in `order` emptied, one
/// release at a time, against `original`'s reference counts.
fn reclaim_in_order(original: &[u8], released: Vec<u8>, order: &[u32]) -> Result<Vec<u8>, Error> {
    let mut counts: BTreeMap<u32, u64> = BTreeMap::new();
    for entry in repeated_bytes(original, 3)? {
        let key = single_varint(entry, 1)?.ok_or(Error::InvalidSource)?;
        let key = u32::try_from(key).map_err(|_error| Error::InvalidSource)?;
        counts.insert(key, single_varint(entry, 2)?.unwrap_or(0));
    }
    let current = single_varint(original, 2)?.ok_or(Error::InvalidSource)?;
    let mut next = current;
    for key in order {
        let Some(count) = counts.get_mut(key) else {
            continue;
        };
        *count = count.saturating_sub(1);
        if *count == 0 && u64::from(*key) + 1 == next {
            next = u64::from(*key);
        }
    }
    if next == current {
        return Ok(released);
    }
    Ok(structure::put_varint(&released, 2, next)?)
}

/// Whether no formula owner of the document holds formulas of, or reads,
/// the table whose drawable is `info`.
fn formula_inert<P: TableLeaf>(working: &Working<'_, P>, info: u64) -> Result<bool, Error> {
    let mut owners: Vec<&[u8]> = Vec::new();
    for object in working
        .source()
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
        for message in &object.messages {
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                owners
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation { amount: 1 })?;
                owners.push(message.data.as_slice());
            }
        }
    }
    let object = |identifier: u64| -> Option<&[u8]> {
        working
            .object(identifier)
            .ok()
            .and_then(|object| object.messages.first())
            .map(|message| message.data.as_slice())
    };
    let limits = wire_limits(working.source())?;
    let state =
        litchi_numbers_wire::table_write::table_formula_state(&owners, &object, info, &limits)
            .map_err(|_error| Error::InvalidSource)?;
    Ok(state == litchi_numbers_wire::table_write::FormulaState::Inert)
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
    /// Data-store list field -> keys in release order (cell by cell), for
    /// the host's per-release hand-back of a last key (record 0370b).
    order: BTreeMap<u32, Vec<u32>>,
    /// Formula-list key -> references the dropped cells hold (record 0329).
    formulas: BTreeMap<u64, u64>,
}

fn released_cells<P: TableLeaf>(
    working: &Working<'_, P>,
    store: &[u8],
    row_axis: bool,
    index: u32,
    values: bool,
) -> Result<Released, Error> {
    let (tiles, size) = tiles(store)?;
    let limits = wire_limits(working.source())?;
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
                    let order = released.order.entry(field).or_default();
                    order
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation { amount: 1 })?;
                    order.push(list_key);
                }
            }
        }
    }
    Ok(released)
}

/// Apply one formula cell patch (record 0329) to the tile of `store` that
/// holds the cell, keeping the rewritten tile in `patched`; returns the
/// string-list key a replaced text cache held.
fn patch_cell<P: TableLeaf>(
    working: &Working<'_, P>,
    store: &[u8],
    columns: u32,
    patch: shift::CellPatch,
    patched: &mut BTreeMap<u64, Vec<u8>>,
) -> Result<Option<u32>, Error> {
    let (tiles, size) = tiles(store)?;
    let limits = wire_limits(working.source())?;
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
    let cell_bytes = cell.encode(&limits).map_err(|_| Error::InvalidSource)?;
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
fn plan<P: TableLeaf>(
    working: &Working<'_, P>,
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
                let order = released.order.entry(STRING_LIST_FIELD).or_default();
                order
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation { amount: 1 })?;
                order.push(key);
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
        let (payload, pruned) =
            lists::release_with(data, &mut remaining, false).map_err(storage_fault)?;
        if let Some(payload) = payload {
            // A released last key goes back as each release empties it, as
            // the host's per-cell list decrements do (record 0370b).
            let order = released.order.get(field).map_or(&[][..], Vec::as_slice);
            let payload = reclaim_in_order(data, payload, order)?;
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

fn replace<P: TableLeaf>(
    working: &mut Working<'_, P>,
    replacement: Replacement,
) -> FaultResult<()> {
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

/// The current geometry of the located table.
///
/// # Errors
///
/// A malformed model or a structural fault.
pub fn read_geometry<P: TableLeaf>(source: &P, target: &Target) -> Result<Geometry, Error> {
    let working = Working::new(source)?;
    let (_, model) = message(&working, target.model, &[6_001])?;
    geometry(model)
}

/// The deletion-safe preflight: everything [`publish`] would refuse for
/// `step`, checked before a leaf clears comments or cells.
///
/// # Errors
///
/// A typed [`Refusal`], a malformed source or a structural fault.
pub fn check<P: TableLeaf>(
    source: &P,
    target: &Target,
    step: Step,
    before: Geometry,
) -> Result<(), Error> {
    let working = Working::new(source)?;
    plan(&working, target, step, before, true, false).map(|_planned| ())
}

/// The stored cells on `step`'s removed row or column, as table positions
/// (none for an insertion).
///
/// # Errors
///
/// A malformed source or a structural fault.
pub fn removed_cells<P: TableLeaf>(
    source: &P,
    target: &Target,
    step: Step,
    before: Geometry,
) -> Result<Vec<At>, Error> {
    let AxisEdit::Remove(index) = step.edit else {
        return Ok(Vec::new());
    };
    let working = Working::new(source)?;
    let (_, model) = message(&working, target.model, &[6_001])?;
    let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;
    cells_on_axis(&working, store, before.columns, step.row_axis, index)
}

/// Whether no formula owner of the document holds formulas of, or reads,
/// the table.
///
/// # Errors
///
/// A malformed calculation graph or a structural fault.
pub fn table_formula_inert<P: TableLeaf>(source: &P, target: &Target) -> Result<bool, Error> {
    let working = Working::new(source)?;
    formula_inert(&working, target.info)
}

/// How [`publish`] finishes one step.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Publish {
    /// The removed row's or column's cells were not cleared by the leaf:
    /// their values, strings and formulas are released here.
    pub drop_cells: bool,
    /// The table drawable's geometry size follows the table's new rendered
    /// size (the sum of its row heights and column widths), as Keynote's
    /// host edit moves the slide drawable bounds with the table.
    pub fit_drawable: bool,
}

/// Rewrite and publish one resolved step over `source` (root previews
/// kept, as the host's axis edit keeps them).
///
/// # Errors
///
/// A typed [`Refusal`], a malformed source or a structural fault.
pub fn publish<P: TableLeaf>(
    source: &P,
    target: &Target,
    step: Step,
    before: Geometry,
    options: Publish,
) -> Result<Published<P>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let planned = plan(&working, target, step, before, false, options.drop_cells)?;
    for replacement in planned.replacements {
        replace(&mut working, replacement)?;
    }
    if options.fit_drawable {
        let (width, height) = size_points(&working, target.model)?;
        let (index, info) = message(&working, target.info, &[6_000, 6_003])?;
        let data = fit_drawable_size(info, width, height)?;
        replace(
            &mut working,
            Replacement {
                object: target.info,
                index,
                data,
                pruned: Vec::new(),
            },
        )?;
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
    Ok(working.publish(&transition)?)
}

/// `TST.TableModelArchive.default_row_height`.
const MODEL_DEFAULT_ROW_HEIGHT_FIELD: u32 = 16;
/// `TST.TableModelArchive.default_column_width`.
const MODEL_DEFAULT_COLUMN_WIDTH_FIELD: u32 = 17;

/// One singular fixed-width field's raw bytes.
fn fixed_field(data: &[u8], number: u32, wire_type: u8) -> Result<Option<&[u8]>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(data, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| Error::from(structure::map_wire(&error)))?;
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type || found.replace(field.payload()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(found)
}

/// The explicit sizes of one axis's header buckets: header index to size
/// in points (a zero size is the model default).
fn explicit_sizes<P: TableLeaf>(
    working: &Working<'_, P>,
    buckets: &[u64],
) -> Result<BTreeMap<u32, f32>, Error> {
    let mut sizes = BTreeMap::new();
    for bucket in buckets {
        let (_, data) = message(working, *bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        for header in repeated_bytes(data, 2)? {
            let index = u32::try_from(single_varint(header, 1)?.ok_or(Error::InvalidSource)?)
                .map_err(|_error| Error::InvalidSource)?;
            let Some(bits) = fixed_field(header, 2, 5)? else {
                continue;
            };
            let bits = u32::from_le_bytes(bits.try_into().map_err(|_error| Error::InvalidSource)?);
            let size = f32::from_bits(bits);
            if !size.is_finite() || size < 0.0 {
                return Err(Error::InvalidSource);
            }
            if bits != 0 && sizes.insert(index, size).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(sizes)
}

/// The table's rendered size in points: every row's and column's explicit
/// header size or the model default, summed (the host's
/// `attached_table_size_points`).
fn size_points<P: TableLeaf>(working: &Working<'_, P>, model: u64) -> Result<(f32, f32), Error> {
    let (_, model) = message(working, model, &[6_001])?;
    let geometry = geometry(model)?;
    let default = |number| -> Result<f32, Error> {
        let bytes = fixed_field(model, number, 1)?.ok_or(Error::InvalidSource)?;
        let value = f64::from_bits(u64::from_le_bytes(
            bytes.try_into().map_err(|_error| Error::InvalidSource)?,
        ));
        Ok(value as f32)
    };
    let default_row = default(MODEL_DEFAULT_ROW_HEIGHT_FIELD)?;
    let default_column = default(MODEL_DEFAULT_COLUMN_WIDTH_FIELD)?;
    let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;
    let row_storage = single_bytes(store, 1)?.ok_or(Error::InvalidSource)?;
    let mut row_buckets = Vec::new();
    for reference in repeated_bytes(row_storage, 2)? {
        structure::reserve(&mut row_buckets, 1)?;
        row_buckets.push(references::reference_identifier(reference)?);
    }
    let column_bucket = single_reference(store, 2)?.ok_or(Error::InvalidSource)?;
    let rows = explicit_sizes(working, &row_buckets)?;
    let columns = explicit_sizes(working, &[column_bucket])?;
    let total = |count: u32, explicit: &BTreeMap<u32, f32>, default: f32| -> Result<f32, Error> {
        let mut sum = 0.0_f64;
        for index in 0..count {
            let points = explicit.get(&index).copied().unwrap_or(default);
            if !points.is_finite() || points <= 0.0 {
                return Err(Error::InvalidSource);
            }
            sum += f64::from(points);
        }
        let sum = sum as f32;
        if !sum.is_finite() || sum <= 0.0 {
            return Err(Error::InvalidSource);
        }
        Ok(sum)
    };
    Ok((
        total(geometry.columns, &columns, default_column)?,
        total(geometry.rows, &rows, default_row)?,
    ))
}

/// The table drawable payload with its geometry size replaced
/// (`TSD.DrawableArchive.geometry.size`, a `TSP.Size` of two floats).
fn fit_drawable_size(info: &[u8], width: f32, height: f32) -> Result<Vec<u8>, Error> {
    let wire_fault = |error: litchi_iwa_common::Error| Error::from(structure::map_wire(&error));
    let drawable = single_bytes(info, 1)?.ok_or(Error::InvalidSource)?;
    let geometry = single_bytes(drawable, 1)?.ok_or(Error::InvalidSource)?;
    let mut size = Vec::new();
    structure::reserve(&mut size, 10)?;
    size.push(0x0d);
    size.extend_from_slice(&width.to_le_bytes());
    size.push(0x15);
    size.extend_from_slice(&height.to_le_bytes());
    let present = single_bytes(geometry, 2)?.is_some();
    let geometry =
        litchi_iwa_common::wire::patch_length_delimited_field(geometry, 2, present, Some(&size))
            .map_err(wire_fault)?;
    let drawable =
        litchi_iwa_common::wire::patch_length_delimited_field(drawable, 1, true, Some(&geometry))
            .map_err(wire_fault)?;
    litchi_iwa_common::wire::patch_length_delimited_field(info, 1, true, Some(&drawable))
        .map_err(wire_fault)
}
