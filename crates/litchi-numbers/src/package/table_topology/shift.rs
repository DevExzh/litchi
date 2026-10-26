//! The formula-coordinate shift of a table-axis edit (record 0318).
//!
//! An inserted or removed row or column moves every cell after it, so every
//! coordinate the calculation engine holds for the edited table moves too.
//! This is a wire-level port of the retired host editor's
//! `formula_dependency_shift` (`litchi-iwa`, over Prost archives) onto the
//! exact payloads this crate rewrites. It uses the same bounded wire
//! primitives (`litchi_iwa_common::wire`) in the same order, so a payload it
//! rewrites carries the bytes the host editor writes; the differential test
//! `litchi-iwa/tests/numbers_topology_formula_differential.rs` checks that
//! over the corpus and over synthetic workbooks.
//!
//! What moves:
//!
//! - every formula AST of the table's own formula owner: local cell
//!   references and colon tracts (relative and absolute, both endpoints of a
//!   range), in nested thunks too; a relative reference into another table
//!   compensates for its host's move and keeps pointing at the same cell; a
//!   footer formula's range that ends on the last body row grows with a body
//!   row inserted after it and shrinks with that row's removal;
//! - the owner's cell dependency records (inline and tiled): the formula
//!   hosts and their precedent edges into the table;
//! - its range dependencies and range-precedent tiles, its UUID-reference
//!   coordinate sets, and the table extents of its spanning dependencies;
//! - the ranges a merge-proxy owner (an owner without a drawable) holds into
//!   the table.
//!
//! Refused, typed, before anything is published:
//!
//! Removed rows and columns (record 0329): a reference to a removed cell,
//! and a range wholly on the removed axis, become `REFERENCE_ERROR_NODE`s
//! (`#REF!`, as native Numbers writes them) and the formula caches an error
//! result; a range that loses a row or column shrinks; precedent edges and
//! range dependencies into the removed axis go with it. Where the cell
//! writer could not clear the removed cells, their hosts' dependency
//! records go here, and an emptied dependency tile with them.
//!
//! - [`Unsupported::DeletedReference`]: the removed axis is one of several
//!   ranges of one colon tract, or an endpoint of a merge range.
//! - [`Unsupported::ReferencedByOtherOwners`]: a formula in another table, or
//!   a chart, reads the table (its dependency edges, legacy UUID ranges or
//!   UUID references name the table).
//! - [`Unsupported::Formulas`]: volatile, spanning, whole-owner, error or
//!   spill dependency state in any owner, a category reference whose host
//!   moves, a packed coordinate, a dependency record that would cross its
//!   tile.
//! - [`Unsupported::SharedFormulas`]: one formula-list entry shared by cells
//!   whose rewritten formulas would differ, in a segmented formula list (an
//!   unsegmented list is split: record 0329).

use std::collections::{BTreeMap, HashMap, HashSet};

use litchi_iwa_common::wire as w;

use super::super::Package;
use super::super::structure::{Fault, map_wire, references, reserve};
use super::storage::bnc_cells;
use super::wire::{
    bytes_field, fields, push_varint, push_varint_field, rebuild, repeated_bytes, repeated_varints,
    single_bytes, single_reference, single_varint,
};
use super::{LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE, TILE_MESSAGE_TYPE, Unsupported};

/// `TSCE.FormulaOwnerDependenciesArchive`.
const OWNER_TYPE: u32 = 4_008;
/// `TSCE.CellRecordTileArchive`.
const CELL_TILE_TYPE: u32 = 4_009;
/// `TSCE.RangePrecedentsTileArchive`.
const RANGE_TILE_TYPE: u32 = 4_010;
/// `TSCE.ASTNodeArrayArchive.ASTNodeType.COLON_NODE`.
const COLON_NODE: u64 = 29;
/// Rows of one dependency-record tile.
const DEPENDENCY_TILE_ROWS: u32 = 128;
/// Columns of one dependency-record tile.
const DEPENDENCY_TILE_COLUMNS: u32 = 32;
/// The perpendicular-axis marker of a whole-row reference.
const WHOLE_ROW_COLUMN_SENTINEL: u64 = 0x7fff;
/// The perpendicular-axis marker of a whole-column reference.
const WHOLE_COLUMN_ROW_SENTINEL: u64 = 0x7fff_ffff;
/// `TST.TableDataList.ListType.FORMULA`.
const FORMULA_LIST: u64 = 3;

/// Why the shift failed: a malformed source, or a typed refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Failure {
    Fault(Fault),
    Refused(Unsupported),
}

impl From<Fault> for Failure {
    fn from(fault: Fault) -> Self {
        Self::Fault(fault)
    }
}

impl From<litchi_iwa_common::Error> for Failure {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self::Fault(map_wire(&error))
    }
}

type Out<T> = Result<T, Failure>;

const fn refuse<T>(reason: Unsupported) -> Out<T> {
    Err(Failure::Refused(reason))
}

/// One axis edit, as the calculation engine sees it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Shift {
    /// A row edit (else a column edit).
    pub(super) row_axis: bool,
    /// The physical index inserted before, or removed.
    pub(super) position: u32,
    /// An insertion (else a removal).
    pub(super) insert: bool,
    /// A body-row insertion: a footer range ending on the row before it
    /// grows.
    pub(super) expand_footer: bool,
    /// A check before the removed axis is cleared: formula hosts on it are
    /// left as they are (the clear removes them).
    pub(super) pending_clear: bool,
    /// The removed axis was not cleared through the cell writer (record
    /// 0329): the formula hosts on it are dropped here, with their
    /// dependency records and formula-list references, and no dependent
    /// cache is recomputed, so a surviving range over the removed axis is
    /// refused.
    pub(super) drop_hosts: bool,
}

impl Shift {
    /// The new coordinate of `value`; a removed coordinate is a typed
    /// refusal.
    fn coordinate(self, value: u32) -> Out<u32> {
        if value < self.position {
            return Ok(value);
        }
        if self.insert {
            return value.checked_add(1).ok_or(Failure::Fault(Fault::Invalid));
        }
        if value > self.position {
            return Ok(value - 1);
        }
        refuse(Unsupported::DeletedReference)
    }

    const fn axis(self, row: u32, column: u32) -> u32 {
        if self.row_axis { row } else { column }
    }

    /// Whether a formula host at `(row, column)` sits on the removed axis
    /// that a pending clear empties.
    const fn cleared(self, row: u32, column: u32) -> bool {
        (self.pending_clear || self.drop_hosts)
            && !self.insert
            && self.axis(row, column) == self.position
    }

    /// Whether `value` is the removed coordinate.
    const fn removed(self, value: u32) -> bool {
        !self.insert && value == self.position
    }

    /// The interval `[low, high]` after the edit: a removed endpoint moves
    /// inward, as native Numbers shrinks a range that loses a row or
    /// column (record 0329); `None` when the whole interval is removed.
    fn interval(self, low: u32, high: u32) -> Out<Option<(u32, u32)>> {
        if self.insert || !(low..=high).contains(&self.position) {
            return Ok(Some((self.coordinate(low)?, self.coordinate(high)?)));
        }
        if low == high {
            return Ok(None);
        }
        if self.drop_hosts {
            // The removed cells were not cleared through the cell writer, so
            // the range's cached result would go stale.
            return refuse(Unsupported::Formulas);
        }
        let low = if low > self.position { low - 1 } else { low };
        Ok(Some((low, high - 1)))
    }
}

/// One planned payload replacement.
pub(super) struct Rewrite {
    pub(super) object: u64,
    pub(super) message: usize,
    pub(super) data: Vec<u8>,
}

/// One formula owner of the calculation engine.
struct Owner<'p> {
    object: u64,
    message: usize,
    data: &'p [u8],
}

/// Footer context of one formula host.
#[derive(Debug, Clone, Copy)]
struct Footer {
    boundary: u32,
    expand_on_insert: bool,
}

/// Precedents a footer range gains or loses with the edit.
#[derive(Debug, Default)]
struct Adjustment {
    insert: Vec<(u32, u32)>,
    remove: Vec<(u32, u32)>,
}

impl Adjustment {
    fn normalize(&mut self) {
        self.insert.sort_unstable();
        self.insert.dedup();
        self.remove.sort_unstable();
        self.remove.dedup();
    }

    fn is_empty(&self) -> bool {
        self.insert.is_empty() && self.remove.is_empty()
    }
}

/// The table being edited.
pub(super) struct Table<'a> {
    /// The table drawable (`TST.TableInfoArchive`).
    pub(super) info: u64,
    /// The table model payload.
    pub(super) model: &'a [u8],
    /// Rows before the edit.
    pub(super) rows: u32,
    /// Footer rows before the edit.
    pub(super) footer_rows: u32,
    /// The references removed cells hold per formula-list key, when the
    /// cell writer could not clear them (record 0329).
    pub(super) dropped_formulas: &'a BTreeMap<u64, u64>,
}

fn u32_of(value: u64) -> Out<u32> {
    u32::try_from(value).map_err(|_error| Failure::Fault(Fault::Invalid))
}

/// Decode a protobuf `int32`/`sint32`-carrying varint as the host does.
fn int32_of(value: u64) -> Out<i32> {
    // `int32` is sign-extended to 64 bits on the wire.
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "protobuf int32 wire semantics: the low 32 bits carry the value"
    )]
    Ok(value as i64 as i32)
}

const fn zigzag_decode(value: u64) -> i64 {
    #[allow(clippy::cast_possible_wrap, reason = "zigzag decoding")]
    let decoded = ((value >> 1) as i64) ^ -((value & 1) as i64);
    decoded
}

const fn zigzag_i32(value: i32) -> u64 {
    #[allow(clippy::cast_sign_loss, reason = "zigzag encoding of a 32-bit value")]
    let encoded = ((value << 1) ^ (value >> 31)) as u32 as u64;
    encoded
}

fn int32_wire(value: i32) -> u64 {
    #[allow(clippy::cast_sign_loss, reason = "protobuf int32 wire semantics")]
    let encoded = i64::from(value) as u64;
    encoded
}

fn optional_varint(data: &[u8], number: u32) -> Out<Option<u64>> {
    Ok(single_varint(data, number)?)
}

fn required_u32(data: &[u8], number: u32) -> Out<u32> {
    u32_of(single_varint(data, number)?.ok_or(Fault::Invalid)?)
}

/// Every formula owner of the package, in catalog order.
fn owners(package: &Package) -> Out<Vec<Owner<'_>>> {
    let mut output = Vec::new();
    for component in package.state.components.catalog().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == OWNER_TYPE {
                    reserve(&mut output, 1)?;
                    output.push(Owner {
                        object: object.archive_info.identifier.ok_or(Fault::Invalid)?,
                        message: index,
                        data: &message.data,
                    });
                }
            }
        }
    }
    Ok(output)
}

/// The unique message of `identifier` with type `type_`.
fn object_message(package: &Package, identifier: u64, type_: u32) -> Out<(usize, &[u8])> {
    for component in package.state.components.catalog().iter() {
        if let Some(object) = component.archive().object(identifier) {
            let mut found = None;
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == type_
                    && found.replace((index, message.data.as_slice())).is_some()
                {
                    return Err(Fault::Invalid.into());
                }
            }
            return Ok(found.ok_or(Fault::Invalid)?);
        }
    }
    Err(Fault::Invalid.into())
}

fn tile_references(owner: &[u8], number: u32) -> Out<Vec<u64>> {
    let mut output = Vec::new();
    if let Some(tiled) = single_bytes(owner, number)? {
        for reference in repeated_bytes(tiled, 1)? {
            reserve(&mut output, 1)?;
            output.push(references::reference_identifier(reference)?);
        }
    }
    Ok(output)
}

/// Whether an `CellCoordSetArchive`-carrying payload holds column entries.
fn coord_set_populated(data: Option<&[u8]>) -> Out<bool> {
    Ok(match data {
        Some(set) => !repeated_bytes(set, 1)?.is_empty(),
        None => false,
    })
}

/// `validate_shiftable_formula_owner`: volatile, spanning, whole-owner,
/// error and spill state block an axis edit.
fn shiftable(owner: &[u8]) -> Out<bool> {
    if let Some(volatile) = single_bytes(owner, 6)? {
        for number in 1..=5 {
            if coord_set_populated(single_bytes(volatile, number)?)? {
                return Ok(false);
            }
        }
        if coord_set_populated(single_bytes(volatile, 7)?)? {
            return Ok(false);
        }
    }
    for number in [7, 8] {
        if let Some(spanning) = single_bytes(owner, number)?
            && !repeated_bytes(spanning, 1)?.is_empty()
        {
            return Ok(false);
        }
    }
    if let Some(whole) = single_bytes(owner, 9)?
        && coord_set_populated(single_bytes(whole, 1)?)?
    {
        return Ok(false);
    }
    if let Some(errors) = single_bytes(owner, 10)?
        && (!repeated_bytes(errors, 1)?.is_empty() || !repeated_bytes(errors, 2)?.is_empty())
    {
        return Ok(false);
    }
    if let Some(spills) = single_bytes(owner, 16)?
        && !repeated_bytes(spills, 1)?.is_empty()
    {
        return Ok(false);
    }
    Ok(true)
}

/// Rewrite the repeated varint field `number` of `data` to `values`, in its
/// own encoding: unpacked fields as the host editor rewrites them, one
/// packed field (as this crate's formula writer emits) in place.
fn write_varints(data: &[u8], number: u32, values: &[u64]) -> Out<Vec<u8>> {
    let occurrences: Vec<u8> = fields(data)?
        .iter()
        .filter(|field| field.number == number)
        .map(|field| field.wire_type)
        .collect();
    if !occurrences.contains(&2) {
        return Ok(w::rewrite_repeated_varint_fields(data, number, values)?);
    }
    if occurrences.len() != 1 {
        return Err(Fault::Invalid.into());
    }
    let mut packed = Vec::new();
    for value in values {
        push_varint(&mut packed, *value)?;
    }
    let field = bytes_field(number, &packed)?;
    Ok(rebuild(data, |candidate| {
        Ok((candidate.number == number).then(|| field.clone()))
    })?)
}

/// The cell records of a records container (`field` repeated records).
fn records(container: &[u8], field: u32) -> Out<Vec<&[u8]>> {
    Ok(repeated_bytes(container, field)?)
}

/// Whether one record's edges name `internal` as their owner.
fn record_has_external_owner(record: &[u8], internal: u64) -> Out<bool> {
    let Some(edges) = single_bytes(record, 6)? else {
        return Ok(false);
    };
    Ok(repeated_varints(edges, 5)?.contains(&internal))
}

/// Whether a UUID-references payload names `uid` as an owner.
fn uuid_references_owner(owner: &[u8], uid: &[u8]) -> Out<bool> {
    let Some(references) = single_bytes(owner, 14)? else {
        return Ok(false);
    };
    for number in [1, 2] {
        for reference in repeated_bytes(references, number)? {
            if single_bytes(reference, 1)? == Some(uid) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// The incoming-range state of another owner: its range-precedent tiles
/// into the table, and whether it holds an internal range into it.
struct Incoming {
    tiles: Vec<u64>,
    ranges: bool,
    /// The owner is another table's: its formulas read the table, and their
    /// references move with it (beyond the host editor, which refuses).
    reader: bool,
}

/// Check another owner for references into the table (`internal`, `uid`).
///
/// Returns the merge-proxy ranges to rewrite, if any; any other incoming
/// reference is a refusal.
fn incoming(package: &Package, owner: &[u8], internal: u64, uid: &[u8]) -> Out<Option<Incoming>> {
    let mut edges = false;
    if let Some(dependencies) = single_bytes(owner, 4)? {
        for record in records(dependencies, 1)? {
            edges |= record_has_external_owner(record, internal)?;
        }
    }
    let mut internal_ranges = false;
    if let Some(ranges) = single_bytes(owner, 5)? {
        for dependency in repeated_bytes(ranges, 2)? {
            if single_bytes(dependency, 3)?.is_some() {
                return refuse(Unsupported::ReferencedByOtherOwners);
            }
            if let Some(reference) = single_bytes(dependency, 4)?
                && optional_varint(reference, 1)? == Some(internal)
            {
                internal_ranges = true;
            }
        }
    }
    if uuid_references_owner(owner, uid)? {
        return refuse(Unsupported::ReferencedByOtherOwners);
    }
    for tile in tile_references(owner, 13)? {
        let (_, data) = object_message(package, tile, CELL_TILE_TYPE)?;
        for record in records(data, 4)? {
            edges |= record_has_external_owner(record, internal)?;
        }
    }
    let mut tiles = Vec::new();
    for tile in tile_references(owner, 15)? {
        let (_, data) = object_message(package, tile, RANGE_TILE_TYPE)?;
        if optional_varint(data, 1)? == Some(internal) && !repeated_bytes(data, 2)?.is_empty() {
            reserve(&mut tiles, 1)?;
            tiles.push(tile);
        }
    }
    if !edges && !internal_ranges && tiles.is_empty() {
        return Ok(None);
    }
    let reader = single_bytes(owner, 11)?.is_some();
    if edges && !reader {
        // A chart or another non-table owner.
        return refuse(Unsupported::ReferencedByOtherOwners);
    }
    Ok(Some(Incoming {
        tiles,
        ranges: internal_ranges,
        reader,
    }))
}

/// A `TSCE.RangeCoordinateArchive`: `[top-left column, top-left row,
/// bottom-right column, bottom-right row]`.
fn range_coordinate(data: &[u8]) -> Out<[u32; 4]> {
    Ok([
        required_u32(data, 1)?,
        required_u32(data, 2)?,
        required_u32(data, 3)?,
        required_u32(data, 4)?,
    ])
}

fn patch_range_coordinate(data: &[u8], values: [u32; 4]) -> Out<Vec<u8>> {
    let mut output = data.to_vec();
    for (number, value) in (1_u32..).zip(values) {
        output = w::patch_varint_field(&output, number, true, Some(u64::from(value)))?;
    }
    Ok(output)
}

/// `mutate_range_coordinate`, with the footer override; a range that loses
/// a row or column shrinks, and `None` is a range wholly removed (record
/// 0329).
fn mutate_range(
    range: [u32; 4],
    shift: Shift,
    adjustment: Option<&Adjustment>,
) -> Out<Option<[u32; 4]>> {
    let [left, top, right, bottom] = range;
    let footer = adjustment.and_then(|adjustment| {
        if !shift.row_axis {
            return None;
        }
        let in_columns =
            |&(row, column): &(u32, u32)| row == shift.position && (left..=right).contains(&column);
        if shift.insert {
            (bottom.checked_add(1) == Some(shift.position)
                && adjustment.insert.iter().any(in_columns))
            .then_some(shift.position)
        } else {
            (bottom == shift.position
                && top < shift.position
                && adjustment.remove.iter().any(in_columns))
            .then(|| shift.position.checked_sub(1))
            .flatten()
        }
    });
    if shift.row_axis {
        if let Some(value) = footer {
            return Ok(Some([left, shift.coordinate(top)?, right, value]));
        }
        Ok(shift
            .interval(top.min(bottom), top.max(bottom))?
            .map(|(top, bottom)| [left, top, right, bottom]))
    } else {
        Ok(shift
            .interval(left.min(right), left.max(right))?
            .map(|(left, right)| [left, top, right, bottom]))
    }
}

/// An explicit `TSCE.CellCoordinateArchive`: `(row, column)`.
fn explicit_coordinate(data: &[u8]) -> Out<(u32, u32)> {
    if fields(data)?.iter().any(|field| field.number == 1) {
        // A packed coordinate.
        return refuse(Unsupported::Formulas);
    }
    Ok((required_u32(data, 3)?, required_u32(data, 2)?))
}

fn patch_coordinate(data: &[u8], row: u32, column: u32) -> Out<Vec<u8>> {
    let has_column = optional_varint(data, 2)?.is_some();
    let has_row = optional_varint(data, 3)?.is_some();
    let output = w::patch_varint_field(data, 2, has_column, Some(u64::from(column)))?;
    Ok(w::patch_varint_field(
        &output,
        3,
        has_row,
        Some(u64::from(row)),
    )?)
}

/// `mutate_cell_rect` and its wire rewrite; `None` for a rect wholly
/// removed.
fn rewrite_rect(
    rect: &[u8],
    shift: Shift,
    adjustment: Option<&Adjustment>,
) -> Out<Option<Vec<u8>>> {
    let origin = single_bytes(rect, 1)?.ok_or(Fault::Invalid)?;
    let size = single_bytes(rect, 2)?.ok_or(Fault::Invalid)?;
    let (row, column) = explicit_coordinate(origin)?;
    let previous_columns = optional_varint(size, 1)?;
    let previous_rows = optional_varint(size, 2)?;
    let columns = u32_of(previous_columns.unwrap_or(1))?;
    let rows = u32_of(previous_rows.unwrap_or(1))?;
    let range = [
        column,
        row,
        column
            .checked_add(columns.saturating_sub(1))
            .ok_or(Fault::Invalid)?,
        row.checked_add(rows.saturating_sub(1))
            .ok_or(Fault::Invalid)?,
    ];
    let Some([left, top, right, bottom]) = mutate_range(range, shift, adjustment)? else {
        return Ok(None);
    };
    let columns = right
        .checked_sub(left)
        .and_then(|size| size.checked_add(1))
        .ok_or(Fault::Invalid)?;
    let rows = bottom
        .checked_sub(top)
        .and_then(|size| size.checked_add(1))
        .ok_or(Fault::Invalid)?;
    let output =
        w::transform_length_delimited_field(rect, 1, |origin| patch_coordinate(origin, top, left))?;
    Ok(Some(w::transform_length_delimited_field(
        &output,
        2,
        |size| -> Out<Vec<u8>> {
            let size = w::patch_varint_field(
                size,
                1,
                previous_columns.is_some(),
                (columns != 1).then_some(u64::from(columns)),
            )?;
            Ok(w::patch_varint_field(
                &size,
                2,
                previous_rows.is_some(),
                (rows != 1).then_some(u64::from(rows)),
            )?)
        },
    )?))
}

// ---------------------------------------------------------------------------
// Formula ASTs.
// ---------------------------------------------------------------------------

/// One decoded relative or absolute colon-tract range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TractRange {
    begin: i64,
    end: Option<i64>,
}

fn tract_ranges(tract: &[u8], number: u32, relative: bool) -> Out<Vec<TractRange>> {
    let mut output = Vec::new();
    for range in w::repeated_length_delimited_payloads(tract, number)? {
        let begin = optional_varint(range, 1)?.unwrap_or(0);
        let end = optional_varint(range, 2)?;
        reserve(&mut output, 1)?;
        output.push(if relative {
            TractRange {
                begin: i64::from(int32_of(begin)?),
                end: end.map(int32_of).transpose()?.map(i64::from),
            }
        } else {
            TractRange {
                begin: i64::from(u32_of(begin)?),
                end: end.map(u32_of).transpose()?.map(i64::from),
            }
        });
    }
    Ok(output)
}

fn resolve_relative(host: u32, offset: i64) -> Out<u32> {
    u32::try_from(i64::from(host) + offset)
        .map_err(|_error| Failure::Refused(Unsupported::Formulas))
}

/// `tract_axis_intervals`.
fn tract_intervals(tract: &[u8], row_axis: bool, host: u32) -> Out<Vec<(u32, u32)>> {
    let (relative, absolute) = if row_axis { (2, 4) } else { (1, 3) };
    let mut output = Vec::new();
    for range in tract_ranges(tract, relative, true)? {
        let begin = resolve_relative(host, range.begin)?;
        let end = match range.end {
            Some(end) => resolve_relative(host, end)?,
            None => begin,
        };
        reserve(&mut output, 1)?;
        output.push((begin.min(end), begin.max(end)));
    }
    for range in tract_ranges(tract, absolute, false)? {
        let begin = u32_of(u64::try_from(range.begin).map_err(|_error| Fault::Invalid)?)?;
        let end = match range.end {
            Some(end) => u32_of(u64::try_from(end).map_err(|_error| Fault::Invalid)?)?,
            None => begin,
        };
        reserve(&mut output, 1)?;
        output.push((begin.min(end), begin.max(end)));
    }
    Ok(output)
}

/// Whether the tract covers the whole of `row_axis` (a whole-row or
/// whole-column reference keeps its sentinel).
fn whole_axis(tract: &[u8], row_axis: bool) -> Out<bool> {
    let (relative, absolute, sentinel) = if row_axis {
        (2, 4, WHOLE_COLUMN_ROW_SENTINEL)
    } else {
        (1, 3, WHOLE_ROW_COLUMN_SENTINEL)
    };
    if !w::repeated_length_delimited_payloads(tract, relative)?.is_empty() {
        return Ok(false);
    }
    let ranges = tract_ranges(tract, absolute, false)?;
    Ok(
        matches!(ranges.as_slice(), [range] if u64::try_from(range.begin).ok() == Some(sentinel) && range.end.is_none()),
    )
}

/// `mutate_tract_interval`; `None` for an interval wholly removed (record
/// 0329).
#[allow(
    clippy::too_many_arguments,
    reason = "the host editor's interval rule takes the whole edit context"
)]
fn mutate_interval(
    begin: u32,
    end: u32,
    cross_table: bool,
    shift: Shift,
    footer: Option<Footer>,
    columns: &[(u32, u32)],
    adjustment: &mut Adjustment,
) -> Out<Option<(u32, u32)>> {
    if cross_table {
        return Ok(Some((begin, end)));
    }
    let low = begin.min(end);
    let high = begin.max(end);
    let footer_change = shift.row_axis
        && footer.is_some_and(|footer| {
            if shift.insert {
                footer.expand_on_insert
                    && shift.position == footer.boundary
                    && high.checked_add(1) == Some(shift.position)
            } else {
                shift.position.checked_add(1) == Some(footer.boundary)
                    && high == shift.position
                    && low < shift.position
            }
        });
    let (shifted_low, shifted_high) = if footer_change {
        (
            shift.coordinate(low)?,
            if shift.insert {
                shift.position
            } else {
                shift.position.checked_sub(1).ok_or(Fault::Invalid)?
            },
        )
    } else {
        match shift.interval(low, high)? {
            Some(interval) => interval,
            None => return Ok(None),
        }
    };
    if footer_change {
        if columns.is_empty() {
            return refuse(Unsupported::Formulas);
        }
        for &(start, finish) in columns {
            let target = if shift.insert {
                &mut adjustment.insert
            } else {
                &mut adjustment.remove
            };
            for column in start..=finish {
                reserve(target, 1)?;
                target.push((shift.position, column));
            }
        }
    }
    Ok(Some(if begin <= end {
        (shifted_low, shifted_high)
    } else {
        (shifted_high, shifted_low)
    }))
}

/// The edit context of one formula host.
#[derive(Clone, Copy)]
struct Host {
    row: u32,
    column: u32,
    shift: Shift,
    footer: Option<Footer>,
    /// A formula of another table: only its references into the edited
    /// table (whose formula-owner UUID this is) move, and its host stays.
    incoming: Option<(u64, u64)>,
}

impl Host {
    const fn axis(self) -> u32 {
        self.shift.axis(self.row, self.column)
    }

    fn shifted(self) -> Out<u32> {
        if self.incoming.is_some() {
            return Ok(self.axis());
        }
        self.shift.coordinate(self.axis())
    }

    fn cross_table_offset(self) -> Out<i64> {
        Ok(i64::from(self.axis()) - i64::from(self.shifted()?))
    }

    /// Whether the node addresses another table than the one whose
    /// coordinates move (for an own formula, any cross-table node; for an
    /// incoming formula, every node but those into the edited table).
    fn other_table(self, node: &[u8]) -> Out<bool> {
        let extra = single_bytes(node, 28)?;
        let Some(uid) = self.incoming else {
            return Ok(extra.is_some());
        };
        let Some(extra) = extra else {
            return Ok(true);
        };
        let table = single_bytes(extra, 1)?.ok_or(Fault::Invalid)?;
        Ok(cfuuid_key(table)? != Some(uid))
    }
}

/// The `(lower, upper)` key of a `TSP.CFUUIDArchive`, as the host editor
/// derives it.
fn cfuuid_key(uuid: &[u8]) -> Out<Option<(u64, u64)>> {
    let words = [
        optional_varint(uuid, 2)?,
        optional_varint(uuid, 3)?,
        optional_varint(uuid, 4)?,
        optional_varint(uuid, 5)?,
    ];
    if let [Some(w0), Some(w1), Some(w2), Some(w3)] = words {
        return Ok(Some((w0 | (w1 << 32), w2 | (w3 << 32))));
    }
    let Some(bytes) = single_bytes(uuid, 1)? else {
        return Ok(None);
    };
    let Ok(bytes) = <[u8; 16]>::try_from(bytes) else {
        return Ok(None);
    };
    let value = u128::from_be_bytes(bytes);
    #[allow(clippy::cast_possible_truncation, reason = "splitting a 128-bit value")]
    Ok(Some((value as u64, (value >> 64) as u64)))
}

/// What a colon tract becomes.
enum Tract {
    Unchanged,
    Rewritten(Vec<u8>),
    /// Its only range on the edited axis was removed: the node reads
    /// `#REF!` (record 0329).
    Removed,
}

/// `rewrite_colon_tract` and its wire rewrite (`rewrite_colon_tract_wire`).
fn rewrite_tract(
    tract: &[u8],
    host: Host,
    cross_table: bool,
    adjustment: &mut Adjustment,
) -> Out<Tract> {
    let shift = host.shift;
    let columns = tract_intervals(tract, false, host.column)?;
    let (relative_field, absolute_field) = if shift.row_axis { (2, 4) } else { (1, 3) };
    let anchor = host.axis();
    let shifted_host = host.shifted()?;
    let offset = host.cross_table_offset()?;
    let mut current: BTreeMap<u32, Vec<TractRange>> = BTreeMap::new();
    for number in 1..=4 {
        current.insert(number, tract_ranges(tract, number, number <= 2)?);
    }
    let previous = current.clone();
    let keep_sentinel = whole_axis(tract, shift.row_axis)?;
    let axis_ranges = current[&relative_field].len() + current[&absolute_field].len();
    let mut removed = false;
    // A range whose endpoints mix absolute and relative rows (or columns),
    // `C$2:$C6`, holds one single-point range of each kind: its endpoints
    // move as one interval's (record 0329).
    let mixed = match (
        current[&relative_field].as_slice(),
        current[&absolute_field].as_slice(),
    ) {
        ([relative], [absolute])
            if !cross_table
                && !shift.insert
                && !keep_sentinel
                && host.footer.is_none()
                && relative.end.is_none()
                && absolute.end.is_none() =>
        {
            let first = resolve_relative(anchor, relative.begin)?;
            let second = u32_of(u64::try_from(absolute.begin).map_err(|_error| Fault::Invalid)?)?;
            let (low, high) = (first.min(second), first.max(second));
            if (low..=high).contains(&shift.position) {
                if shift.interval(low, high)?.is_none() {
                    return Ok(Tract::Removed);
                }
                let moved = |value: u32| -> Out<u32> {
                    if value != shift.position {
                        return shift.coordinate(value);
                    }
                    Ok(if value == high {
                        shift.position - 1
                    } else {
                        shift.position
                    })
                };
                Some((moved(first)?, moved(second)?))
            } else {
                None
            }
        },
        _ => None,
    };
    if let Some((first, second)) = mixed {
        if let Some([range]) = current.get_mut(&relative_field).map(Vec::as_mut_slice) {
            range.begin = i64::from(first) - i64::from(shifted_host);
            check_i32(range.begin)?;
        }
        if let Some([range]) = current.get_mut(&absolute_field).map(Vec::as_mut_slice) {
            range.begin = i64::from(second);
        }
    }
    if let Some(ranges) = current.get_mut(&relative_field).filter(|_| mixed.is_none()) {
        for range in ranges.iter_mut() {
            let begin = resolve_relative(anchor, range.begin)?;
            let end = match range.end {
                Some(end) => resolve_relative(anchor, end)?,
                None => begin,
            };
            let Some((shifted_begin, shifted_end)) = mutate_interval(
                begin,
                end,
                cross_table,
                shift,
                host.footer,
                &columns,
                adjustment,
            )?
            else {
                removed = true;
                continue;
            };
            if cross_table {
                range.begin = shift_relative(range.begin, offset)?;
                range.end = range
                    .end
                    .map(|end| shift_relative(end, offset))
                    .transpose()?;
            } else {
                range.begin = i64::from(shifted_begin) - i64::from(shifted_host);
                range.end = (shifted_begin != shifted_end)
                    .then(|| i64::from(shifted_end) - i64::from(shifted_host));
                check_i32(range.begin)?;
                range.end.map(check_i32).transpose()?;
            }
        }
    }
    if !keep_sentinel
        && mixed.is_none()
        && let Some(ranges) = current.get_mut(&absolute_field)
    {
        for range in ranges.iter_mut() {
            let begin = u32_of(u64::try_from(range.begin).map_err(|_error| Fault::Invalid)?)?;
            let end = match range.end {
                Some(end) => u32_of(u64::try_from(end).map_err(|_error| Fault::Invalid)?)?,
                None => begin,
            };
            let Some((shifted_begin, shifted_end)) = mutate_interval(
                begin,
                end,
                cross_table,
                shift,
                host.footer,
                &columns,
                adjustment,
            )?
            else {
                removed = true;
                continue;
            };
            range.begin = i64::from(shifted_begin);
            range.end = (shifted_begin != shifted_end).then_some(i64::from(shifted_end));
        }
    }
    if removed {
        if axis_ranges == 1 {
            return Ok(Tract::Removed);
        }
        // A tract of several ranges, one of them removed.
        return refuse(Unsupported::DeletedReference);
    }
    if current == previous {
        return Ok(Tract::Unchanged);
    }
    let mut output = tract.to_vec();
    for number in 1..=4 {
        let relative = number <= 2;
        let raw = w::repeated_length_delimited_payloads(&output, number)?;
        let before = &previous[&number];
        let after = &current[&number];
        if raw.len() != before.len() || before.len() != after.len() {
            return Err(Fault::Invalid.into());
        }
        let mut replacements = Vec::new();
        reserve(&mut replacements, raw.len())?;
        for ((raw, before), after) in raw.into_iter().zip(before).zip(after) {
            let encode = |value: i64| -> Out<u64> {
                if relative {
                    Ok(int32_wire(check_i32(value)?))
                } else {
                    u64::try_from(value).map_err(|_error| Failure::Fault(Fault::Invalid))
                }
            };
            let rewritten = w::patch_varint_field(raw, 1, true, Some(encode(after.begin)?))?;
            let rewritten = w::patch_varint_field(
                &rewritten,
                2,
                before.end.is_some(),
                after.end.map(encode).transpose()?,
            )?;
            replacements.push(rewritten);
        }
        output = w::rewrite_repeated_length_delimited_fields(&output, number, &replacements)?;
    }
    Ok(Tract::Rewritten(output))
}

fn check_i32(value: i64) -> Out<i32> {
    i32::try_from(value).map_err(|_error| Failure::Refused(Unsupported::Formulas))
}

fn shift_relative(offset: i64, adjustment: i64) -> Out<i64> {
    let shifted = offset.checked_add(adjustment).ok_or(Fault::Invalid)?;
    check_i32(shifted)?;
    Ok(shifted)
}

/// One AST node's coordinate (`AST_column` 26 or `AST_row` 27):
/// `(encoded, absolute)`.
fn node_coordinate(node: &[u8], number: u32) -> Out<Option<(i64, bool)>> {
    let Some(coordinate) = single_bytes(node, number)? else {
        return Ok(None);
    };
    let encoded = zigzag_decode(optional_varint(coordinate, 1)?.unwrap_or(0));
    let absolute = optional_varint(coordinate, 2)?.unwrap_or(0) != 0;
    Ok(Some((encoded, absolute)))
}

/// `local_cell_coordinate`: a same-table reference's target, `(row,
/// column)`.
fn local_coordinate(node: &[u8], host: Host) -> Out<Option<(u32, u32)>> {
    if single_bytes(node, 28)?.is_some() {
        return Ok(None);
    }
    let (Some(row), Some(column)) = (node_coordinate(node, 27)?, node_coordinate(node, 26)?) else {
        return Ok(None);
    };
    let resolve = |(encoded, absolute): (i64, bool), anchor: u32| -> Out<u32> {
        let value = if absolute {
            encoded
        } else {
            i64::from(anchor) + encoded
        };
        u32::try_from(value).map_err(|_error| Failure::Refused(Unsupported::Formulas))
    };
    Ok(Some((
        resolve(row, host.row)?,
        resolve(column, host.column)?,
    )))
}

/// `footer_range_overrides`: the node indices whose row a footer range's
/// growth or shrink replaces.
fn footer_overrides(
    nodes: &[&[u8]],
    host: Host,
    adjustment: &mut Adjustment,
) -> Out<HashMap<usize, u32>> {
    let shift = host.shift;
    let Some(footer) = host.footer.filter(|_| shift.row_axis) else {
        return Ok(HashMap::new());
    };
    let expands = shift.insert && footer.expand_on_insert && shift.position == footer.boundary;
    let contracts = !shift.insert && shift.position.checked_add(1) == Some(footer.boundary);
    if !expands && !contracts {
        return Ok(HashMap::new());
    }
    let mut overrides = HashMap::new();
    for index in 0..nodes.len().saturating_sub(2) {
        if optional_varint(nodes[index + 2], 1)? != Some(COLON_NODE) {
            continue;
        }
        let Some(start) = local_coordinate(nodes[index], host)? else {
            continue;
        };
        let Some(end) = local_coordinate(nodes[index + 1], host)? else {
            continue;
        };
        let endpoint = if end.0 >= start.0 { index + 1 } else { index };
        let replacement = if expands && start.0.max(end.0).checked_add(1) == Some(shift.position) {
            shift.position
        } else if contracts
            && start.0.max(end.0) == shift.position
            && start.0.min(end.0) < shift.position
        {
            shift.position.checked_sub(1).ok_or(Fault::Invalid)?
        } else {
            continue;
        };
        if overrides.insert(endpoint, replacement).is_some() {
            return Err(Fault::Invalid.into());
        }
        let target = if shift.insert {
            &mut adjustment.insert
        } else {
            &mut adjustment.remove
        };
        for column in start.1.min(end.1)..=start.1.max(end.1) {
            reserve(target, 1)?;
            target.push((shift.position, column));
        }
    }
    Ok(overrides)
}

/// `TSCE.ASTNodeArrayArchive.ASTNodeType.REFERENCE_ERROR_NODE`.
const REFERENCE_ERROR_NODE: u64 = 30;

/// The node that replaces a reference to a removed row or column (or a
/// range wholly removed): `#REF!`, as native Numbers writes it (record
/// 0329). The replaced node's leading whitespace stays.
fn reference_error_node(node: &[u8]) -> Out<Vec<u8>> {
    let mut output = Vec::new();
    push_varint_field(&mut output, 1, REFERENCE_ERROR_NODE)?;
    if let Some(whitespace) = single_bytes(node, 25)? {
        output.extend(bytes_field(25, whitespace)?);
    }
    Ok(output)
}

/// The target, on the edited axis, of a node whose coordinate moves with
/// the edit (a same-table reference of an own formula, or a reference into
/// the edited table of another table's formula).
fn moving_target(node: &[u8], host: Host, coordinate_field: u32) -> Out<Option<u32>> {
    if host.other_table(node)? {
        return Ok(None);
    }
    let Some((encoded, absolute)) = node_coordinate(node, coordinate_field)? else {
        return Ok(None);
    };
    let target = if absolute {
        encoded
    } else {
        i64::from(host.axis()) + encoded
    };
    u32::try_from(target)
        .map(Some)
        .map_err(|_error| Failure::Refused(Unsupported::Formulas))
}

/// The ranges (`reference, reference, COLON_NODE`) a removal reaches: the
/// first-node indices of the ranges wholly removed, and the endpoints that
/// move inward (record 0329). Footer overrides keep their endpoints.
fn range_removals(
    nodes: &[&[u8]],
    host: Host,
    coordinate_field: u32,
    overrides: &HashMap<usize, u32>,
) -> Out<(HashSet<usize>, HashMap<usize, u32>)> {
    let shift = host.shift;
    let mut removed = HashSet::new();
    let mut shrunk = HashMap::new();
    if shift.insert {
        return Ok((removed, shrunk));
    }
    let mut index = 0;
    while index + 2 < nodes.len() {
        if optional_varint(nodes[index + 2], 1)? != Some(COLON_NODE) {
            index += 1;
            continue;
        }
        let (Some(first), Some(second)) = (
            moving_target(nodes[index], host, coordinate_field)?,
            moving_target(nodes[index + 1], host, coordinate_field)?,
        ) else {
            index += 1;
            continue;
        };
        let (low, high) = (first.min(second), first.max(second));
        if (low..=high).contains(&shift.position) {
            if overrides.contains_key(&index) || overrides.contains_key(&(index + 1)) {
                index += 3;
                continue;
            }
            match shift.interval(low, high)? {
                None => {
                    removed.insert(index);
                },
                Some(_) => {
                    for (endpoint, target) in [(index, first), (index + 1, second)] {
                        if target == shift.position {
                            let moved = if target == high {
                                shift.position - 1
                            } else {
                                shift.position
                            };
                            shrunk.insert(endpoint, moved);
                        }
                    }
                },
            }
        }
        index += 3;
    }
    Ok((removed, shrunk))
}

/// `rewrite_formula_nodes` and its wire rewrite (`rewrite_ast_array_wire`):
/// the rewritten `ASTNodeArrayArchive`, or `None` when nothing moves. A
/// reference to a removed row or column becomes `#REF!` and sets
/// `references_removed`; a range that loses a row or column shrinks.
fn rewrite_array(
    array: &[u8],
    host: Host,
    adjustment: &mut Adjustment,
    references_removed: &mut bool,
) -> Out<Option<Vec<u8>>> {
    let shift = host.shift;
    let nodes = w::repeated_length_delimited_payloads(array, 1)?;
    let anchor = host.axis();
    let shifted_host = host.shifted()?;
    let offset = host.cross_table_offset()?;
    let overrides = footer_overrides(&nodes, host, adjustment)?;
    let coordinate_field = if shift.row_axis { 27 } else { 26 };
    let (removed_ranges, shrunk) = range_removals(&nodes, host, coordinate_field, &overrides)?;
    let mut changed = false;
    let mut replacements = Vec::new();
    reserve(&mut replacements, nodes.len())?;
    let mut index = 0;
    while index < nodes.len() {
        let node: &[u8] = nodes[index];
        if removed_ranges.contains(&index) {
            replacements.push(reference_error_node(node)?);
            *references_removed = true;
            changed = true;
            index += 3;
            continue;
        }
        index += 1;
        if single_bytes(node, 39)?.is_some() && shifted_host != anchor {
            return refuse(Unsupported::Formulas);
        }
        let cross_table = host.other_table(node)?;
        let mut output = node.to_vec();
        if let Some((encoded, absolute)) = node_coordinate(node, coordinate_field)?
            && !(cross_table && host.incoming.is_some())
        {
            let shifted = if cross_table {
                if absolute {
                    encoded
                } else {
                    encoded.checked_add(offset).ok_or(Fault::Invalid)?
                }
            } else {
                let target = if absolute {
                    encoded
                } else {
                    i64::from(anchor) + encoded
                };
                let target = u32::try_from(target)
                    .map_err(|_error| Failure::Refused(Unsupported::Formulas))?;
                let shifted_target = match overrides
                    .get(&(index - 1))
                    .or_else(|| shrunk.get(&(index - 1)))
                {
                    Some(value) => *value,
                    None if shift.removed(target) => {
                        replacements.push(reference_error_node(node)?);
                        *references_removed = true;
                        changed = true;
                        continue;
                    },
                    None => shift.coordinate(target)?,
                };
                if absolute {
                    i64::from(shifted_target)
                } else {
                    i64::from(shifted_target) - i64::from(shifted_host)
                }
            };
            let shifted = check_i32(shifted)?;
            if i64::from(shifted) != encoded {
                output = w::patch_nested_varint_field(
                    &output,
                    &[coordinate_field, 1],
                    true,
                    Some(zigzag_i32(shifted)),
                )?;
            }
        }
        if let Some(thunk) = single_bytes(node, 14)?
            && let Some(rewritten) = rewrite_array(thunk, host, adjustment, references_removed)?
        {
            output = w::transform_length_delimited_field(&output, 14, |_| -> Out<Vec<u8>> {
                Ok(rewritten)
            })?;
        }
        if let Some(tract) = single_bytes(node, 40)? {
            match rewrite_tract(tract, host, cross_table, adjustment)? {
                Tract::Unchanged => {},
                Tract::Rewritten(rewritten) => {
                    output =
                        w::transform_length_delimited_field(&output, 40, |_| -> Out<Vec<u8>> {
                            Ok(rewritten)
                        })?;
                },
                Tract::Removed => {
                    replacements.push(reference_error_node(node)?);
                    *references_removed = true;
                    changed = true;
                    continue;
                },
            }
        }
        changed |= output.as_slice() != node;
        replacements.push(output);
    }
    if !changed {
        return Ok(None);
    }
    Ok(Some(w::rewrite_repeated_length_delimited_fields(
        array,
        1,
        &replacements,
    )?))
}

/// The rewritten `FormulaArchive`, or `None` when nothing moves.
fn rewrite_formula(
    formula: &[u8],
    host: Host,
    adjustment: &mut Adjustment,
    references_removed: &mut bool,
) -> Out<Option<Vec<u8>>> {
    let array = single_bytes(formula, 1)?.ok_or(Fault::Invalid)?;
    let Some(rewritten) = rewrite_array(array, host, adjustment, references_removed)? else {
        return Ok(None);
    };
    Ok(Some(w::transform_length_delimited_field(
        formula,
        1,
        |_| -> Out<Vec<u8>> { Ok(rewritten) },
    )?))
}

// ---------------------------------------------------------------------------
// The formula list.
// ---------------------------------------------------------------------------

/// One list container: the root list or a segment.
struct Container<'p> {
    object: u64,
    message: usize,
    data: &'p [u8],
}

/// The formula-list containers of the table.
fn formula_containers<'p>(package: &'p Package, store: &[u8]) -> Out<Vec<Container<'p>>> {
    let list = single_reference(store, 6)?.ok_or(Fault::Invalid)?;
    let (message, data) = object_message(package, list, LIST_MESSAGE_TYPE)?;
    if optional_varint(data, 1)? != Some(FORMULA_LIST) {
        return Err(Fault::Invalid.into());
    }
    let mut output = Vec::new();
    reserve(&mut output, 1)?;
    output.push(Container {
        object: list,
        message,
        data,
    });
    for reference in repeated_bytes(data, 4)? {
        let segment = references::reference_identifier(reference)?;
        let (message, data) = object_message(package, segment, LIST_SEGMENT_MESSAGE_TYPE)?;
        reserve(&mut output, 1)?;
        output.push(Container {
            object: segment,
            message,
            data,
        });
    }
    Ok(output)
}

/// One formula-list entry: `(container index, entry index, refcount,
/// formula payload)`.
struct Entry<'p> {
    container: usize,
    position: usize,
    refcount: u64,
    formula: &'p [u8],
}

fn formula_entries<'p>(containers: &[Container<'p>]) -> Out<BTreeMap<u64, Entry<'p>>> {
    let mut output = BTreeMap::new();
    for (container_index, container) in containers.iter().enumerate() {
        for (position, entry) in repeated_bytes(container.data, 3)?.into_iter().enumerate() {
            let key = single_varint(entry, 1)?.ok_or(Fault::Invalid)?;
            let refcount = single_varint(entry, 2)?.ok_or(Fault::Invalid)?;
            let Some(formula) = single_bytes(entry, 5)? else {
                continue;
            };
            let replaced = output.insert(
                key,
                Entry {
                    container: container_index,
                    position,
                    refcount,
                    formula,
                },
            );
            if replaced.is_some() {
                return Err(Fault::Invalid.into());
            }
        }
    }
    Ok(output)
}

/// The formula-list key of the stored formula cell `(row, column)`.
fn formula_key(package: &Package, store: &[u8], row: u32, column: u32) -> Out<u64> {
    let (tiles, size) = super::tiles(store).map_err(|_error| Failure::Fault(Fault::Invalid))?;
    let size = u32_of(size)?;
    let key = row / size;
    let local = u64::from(row % size);
    let tile = tiles
        .iter()
        .find(|(candidate, _)| *candidate == key)
        .map(|(_, tile)| *tile)
        .ok_or(Fault::Invalid)?;
    let (_, data) = object_message(package, tile, TILE_MESSAGE_TYPE)?;
    for stored_row in repeated_bytes(data, 5)? {
        if single_varint(stored_row, 1)? != Some(local) {
            continue;
        }
        for (stored_column, bytes) in bnc_cells(stored_row).map_err(Fault::from)? {
            if stored_column != column {
                continue;
            }
            let cell = litchi_numbers_wire::BncCell::parse(
                bytes,
                &crate::cell::wire::limits_for_source(bytes),
            )
            .map_err(|_error| Failure::Fault(Fault::Invalid))?;
            let key = cell
                .list_keys()
                .find(|(kind, _)| *kind == litchi_numbers_wire::ListKind::Formula)
                .map(|(_, key)| key)
                .ok_or(Fault::Invalid)?;
            return Ok(u64::from(key));
        }
    }
    Err(Fault::Invalid.into())
}

// ---------------------------------------------------------------------------
// Dependency records.
// ---------------------------------------------------------------------------

/// `mutate_dependency_record` and its wire rewrite
/// (`rewrite_shifted_dependency_records`, one record). `None` drops the
/// record of a dropped host; precedent edges into a removed row or column
/// go (record 0329: the formula reads `#REF!` there).
fn rewrite_record(
    record: &[u8],
    shift: Shift,
    internal: u64,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
    range_hosts: &HashSet<(u32, u32)>,
    host_moves: bool,
) -> Out<Option<Vec<u8>>> {
    let column = required_u32(record, 1)?;
    let row = required_u32(record, 2)?;
    if host_moves && shift.cleared(row, column) {
        return Ok((!shift.drop_hosts).then(|| record.to_vec()));
    }
    let (axis_field, axis_value) = if shift.row_axis {
        (2, row)
    } else {
        (1, column)
    };
    let mut output = if host_moves {
        let shifted = shift.coordinate(axis_value)?;
        w::patch_varint_field(record, axis_field, true, Some(u64::from(shifted)))?
    } else {
        record.to_vec()
    };
    let Some(edges) = single_bytes(record, 6)? else {
        return Ok(Some(output));
    };
    let mut local_rows = repeated_varints(edges, 1)?;
    let mut local_columns = repeated_varints(edges, 2)?;
    let mut owned_rows = repeated_varints(edges, 3)?;
    let mut owned_columns = repeated_varints(edges, 4)?;
    let mut owners = repeated_varints(edges, 5)?;
    if local_rows.len() != local_columns.len()
        || owned_rows.len() != owned_columns.len()
        || owned_rows.len() != owners.len()
    {
        return Err(Fault::Invalid.into());
    }
    let host = (row, column);
    if !host_moves {
        // Another table's record: its local edges stay.
    } else if let Some(adjustment) = adjustments
        .get(&host)
        .filter(|_| !range_hosts.contains(&host))
    {
        let mut existing = Vec::new();
        reserve(&mut existing, local_rows.len())?;
        for (row, column) in local_rows.iter().zip(&local_columns) {
            existing.push((u32_of(*row)?, u32_of(*column)?));
        }
        if adjustment
            .remove
            .iter()
            .any(|coordinate| !existing.contains(coordinate))
        {
            return Err(Fault::Invalid.into());
        }
        let mut retained = Vec::new();
        reserve(&mut retained, existing.len() + adjustment.insert.len())?;
        for coordinate in existing {
            if adjustment.remove.binary_search(&coordinate).is_ok() {
                continue;
            }
            let (row, column) = coordinate;
            if shift.removed(shift.axis(row, column)) {
                continue;
            }
            retained.push(if shift.row_axis {
                (shift.coordinate(row)?, column)
            } else {
                (row, shift.coordinate(column)?)
            });
        }
        retained.extend(adjustment.insert.iter().copied());
        retained.sort_unstable();
        retained.dedup();
        local_rows = retained.iter().map(|(row, _)| u64::from(*row)).collect();
        local_columns = retained
            .iter()
            .map(|(_, column)| u64::from(*column))
            .collect();
    } else {
        let mut rows = Vec::new();
        let mut columns = Vec::new();
        reserve(&mut rows, local_rows.len())?;
        reserve(&mut columns, local_columns.len())?;
        for (row, column) in local_rows.iter().zip(&local_columns) {
            let (row, column) = (u32_of(*row)?, u32_of(*column)?);
            if shift.removed(shift.axis(row, column)) {
                continue;
            }
            let (row, column) = if shift.row_axis {
                (shift.coordinate(row)?, column)
            } else {
                (row, shift.coordinate(column)?)
            };
            rows.push(u64::from(row));
            columns.push(u64::from(column));
        }
        local_rows = rows;
        local_columns = columns;
    }
    let mut kept_rows = Vec::new();
    let mut kept_columns = Vec::new();
    let mut kept_owners = Vec::new();
    reserve(&mut kept_rows, owned_rows.len())?;
    reserve(&mut kept_columns, owned_rows.len())?;
    reserve(&mut kept_owners, owned_rows.len())?;
    for ((row, column), owner) in owned_rows.iter().zip(&owned_columns).zip(&owners) {
        let (mut row, mut column) = (u32_of(*row)?, u32_of(*column)?);
        if *owner == internal {
            if shift.removed(shift.axis(row, column)) {
                continue;
            }
            if shift.row_axis {
                row = shift.coordinate(row)?;
            } else {
                column = shift.coordinate(column)?;
            }
        }
        kept_rows.push(u64::from(row));
        kept_columns.push(u64::from(column));
        kept_owners.push(*owner);
    }
    let owners_changed = kept_owners.len() != owners.len();
    owned_rows = kept_rows;
    owned_columns = kept_columns;
    owners = kept_owners;
    output = w::transform_length_delimited_field(&output, 6, |edges| -> Out<Vec<u8>> {
        let mut data = edges.to_vec();
        let mut lists = vec![
            (1, &local_rows),
            (2, &local_columns),
            (3, &owned_rows),
            (4, &owned_columns),
        ];
        if owners_changed {
            lists.push((5, &owners));
        }
        for (number, values) in lists {
            data = write_varints(&data, number, values)?;
        }
        Ok(data)
    })?;
    Ok(Some(output))
}

fn rewrite_records(
    container: &[u8],
    field: u32,
    mut map: impl FnMut(&[u8]) -> Out<Vec<u8>>,
) -> Out<Vec<u8>> {
    filter_records(container, field, |record| map(record).map(Some))
}

/// Rewrite the records of `container` (`field` repeated records); `None`
/// from `map` drops the record (record 0329: a dependency on removed cells
/// only, or a removed host's).
fn filter_records(
    container: &[u8],
    field: u32,
    mut map: impl FnMut(&[u8]) -> Out<Option<Vec<u8>>>,
) -> Out<Vec<u8>> {
    let raw = w::repeated_length_delimited_payloads(container, field)?;
    let mut mapped = Vec::new();
    reserve(&mut mapped, raw.len())?;
    for record in raw {
        mapped.push(map(record)?);
    }
    if mapped.iter().all(Option::is_some) {
        let replacements: Vec<Vec<u8>> = mapped.into_iter().flatten().collect();
        return Ok(w::rewrite_repeated_length_delimited_fields(
            container,
            field,
            &replacements,
        )?);
    }
    let mut next = mapped.into_iter();
    let mut output = Vec::new();
    reserve(&mut output, container.len())?;
    for candidate in fields(container)? {
        if candidate.number != field {
            output.extend_from_slice(candidate.raw);
            continue;
        }
        match next.next().ok_or(Fault::Invalid)? {
            Some(record) if record.as_slice() == candidate.payload => {
                output.extend_from_slice(candidate.raw);
            },
            Some(record) => output.extend(bytes_field(field, &record)?),
            None => {},
        }
    }
    Ok(output)
}

/// `rewrite_range_back_dependency` for one own range dependency; `None`
/// drops it (its host is dropped, or its range wholly removed).
fn rewrite_back_dependency(
    dependency: &[u8],
    shift: Shift,
    internal: u64,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
) -> Out<Option<Vec<u8>>> {
    let row = required_u32(dependency, 1)?;
    let column = required_u32(dependency, 2)?;
    if shift.cleared(row, column) {
        return Ok((!shift.drop_hosts).then(|| dependency.to_vec()));
    }
    let (row_shifted, column_shifted) = if shift.row_axis {
        (shift.coordinate(row)?, column)
    } else {
        (row, shift.coordinate(column)?)
    };
    let external = single_bytes(dependency, 3)?;
    let reference = single_bytes(dependency, 4)?;
    if external.is_some() && reference.is_some() {
        return Err(Fault::Invalid.into());
    }
    let mut output = w::patch_varint_field(dependency, 1, true, Some(u64::from(row_shifted)))?;
    output = w::patch_varint_field(&output, 2, true, Some(u64::from(column_shifted)))?;
    if let Some(reference) = reference
        && optional_varint(reference, 1)? == Some(internal)
    {
        let range = single_bytes(reference, 2)?.ok_or(Fault::Invalid)?;
        let Some(values) = mutate_range(
            range_coordinate(range)?,
            shift,
            adjustments.get(&(row, column)),
        )?
        else {
            return Ok(None);
        };
        output = w::transform_length_delimited_field(&output, 4, |reference| -> Out<Vec<u8>> {
            w::transform_length_delimited_field(reference, 2, |range| {
                patch_range_coordinate(range, values)
            })
        })?;
    }
    Ok(Some(output))
}

/// One range-precedent tile entry (`FromToRangeArchive`); `None` drops it.
fn rewrite_from_to(
    entry: &[u8],
    shift: Shift,
    to_own: bool,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
) -> Out<Option<Vec<u8>>> {
    let from = single_bytes(entry, 1)?.ok_or(Fault::Invalid)?;
    let (row, column) = explicit_coordinate(from)?;
    if shift.cleared(row, column) {
        return Ok((!shift.drop_hosts).then(|| entry.to_vec()));
    }
    let (row_shifted, column_shifted) = if shift.row_axis {
        (shift.coordinate(row)?, column)
    } else {
        (row, shift.coordinate(column)?)
    };
    let mut output = w::transform_length_delimited_field(entry, 1, |from| {
        patch_coordinate(from, row_shifted, column_shifted)
    })?;
    if to_own {
        let adjustment = adjustments.get(&(row, column));
        let rect = single_bytes(&output, 2)?.ok_or(Fault::Invalid)?;
        let Some(rewritten) = rewrite_rect(rect, shift, adjustment)? else {
            return Ok(None);
        };
        output =
            w::transform_length_delimited_field(&output, 2, |_| -> Out<Vec<u8>> { Ok(rewritten) })?;
    }
    Ok(Some(output))
}

/// `mutate_cell_coord_set` and its wire rewrite. With dropped hosts
/// (record 0329) the removed row or column leaves the set.
fn rewrite_coord_set(set: &[u8], shift: Shift) -> Out<Vec<u8>> {
    if shift.pending_clear && !shift.insert {
        // The hosts on the removed axis are cleared first; the coordinates
        // are checked again after the clear.
        return Ok(set.to_vec());
    }
    let dropping = shift.drop_hosts && !shift.insert;
    filter_records(set, 1, |entry| {
        let column = required_u32(entry, 1)?;
        if dropping && !shift.row_axis && column == shift.position {
            return Ok(None);
        }
        let column = if shift.row_axis {
            column
        } else {
            shift.coordinate(column)?
        };
        let output = w::patch_varint_field(entry, 1, true, Some(u64::from(column)))?;
        Ok(Some(w::transform_length_delimited_field(
            &output,
            2,
            |rows| {
                filter_records(rows, 1, |range| {
                    let begin = int32_of(optional_varint(range, 1)?.unwrap_or(0))?;
                    let end = optional_varint(range, 2)?.map(int32_of).transpose()?;
                    let (begin, new_end) = if shift.row_axis {
                        let row = |value: i32| -> Out<u32> {
                            u32::try_from(value).map_err(|_error| Failure::Fault(Fault::Invalid))
                        };
                        let low = row(begin)?;
                        let high = end.map(row).transpose()?.unwrap_or(low);
                        let shifted = if dropping
                            && (low.min(high)..=low.max(high)).contains(&shift.position)
                        {
                            if low == high {
                                return Ok(None);
                            }
                            let (low, high) = (low.min(high), low.max(high));
                            let low = if low > shift.position { low - 1 } else { low };
                            (low, high - 1)
                        } else {
                            (shift.coordinate(low)?, shift.coordinate(high)?)
                        };
                        let encode = |value: u32| -> Out<i32> {
                            i32::try_from(value).map_err(|_error| Failure::Fault(Fault::Invalid))
                        };
                        (
                            encode(shifted.0)?,
                            end.map(|_| encode(shifted.1)).transpose()?,
                        )
                    } else {
                        (begin, end)
                    };
                    let output = w::patch_varint_field(range, 1, true, Some(int32_wire(begin)))?;
                    let output =
                        w::patch_varint_field(&output, 2, end.is_some(), new_end.map(int32_wire))?;
                    Ok(Some(output))
                })
            },
        )?))
    })
}

fn rewrite_uuid_references(references: &[u8], shift: Shift) -> Out<Vec<u8>> {
    let output = rewrite_records(references, 1, |reference| {
        if single_bytes(reference, 2)?.is_none() {
            return Ok(reference.to_vec());
        }
        w::transform_length_delimited_field(reference, 2, |set| rewrite_coord_set(set, shift))
    })?;
    rewrite_records(&output, 2, |group| {
        rewrite_records(group, 2, |reference| {
            if single_bytes(reference, 2)?.is_none() {
                return Ok(reference.to_vec());
            }
            w::transform_length_delimited_field(reference, 2, |set| rewrite_coord_set(set, shift))
        })
    })
}

/// The hosts of the own owner's range dependencies and range tiles.
fn range_dependency_hosts(package: &Package, owner: &[u8]) -> Out<HashSet<(u32, u32)>> {
    let mut hosts = HashSet::new();
    if let Some(ranges) = single_bytes(owner, 5)? {
        for dependency in repeated_bytes(ranges, 2)? {
            hosts.insert((required_u32(dependency, 1)?, required_u32(dependency, 2)?));
        }
    }
    for tile in tile_references(owner, 15)? {
        let (_, data) = object_message(package, tile, RANGE_TILE_TYPE)?;
        for entry in repeated_bytes(data, 2)? {
            hosts.insert(explicit_coordinate(
                single_bytes(entry, 1)?.ok_or(Fault::Invalid)?,
            )?);
        }
    }
    Ok(hosts)
}

/// The dependency records of an owner: `(row, column)` of every record,
/// inline and tiled.
fn formula_hosts(package: &Package, owner: &[u8]) -> Out<Vec<(u32, u32)>> {
    let mut hosts = Vec::new();
    if let Some(dependencies) = single_bytes(owner, 4)? {
        for record in records(dependencies, 1)? {
            reserve(&mut hosts, 1)?;
            hosts.push((required_u32(record, 2)?, required_u32(record, 1)?));
        }
    }
    for tile in tile_references(owner, 13)? {
        let (_, data) = object_message(package, tile, CELL_TILE_TYPE)?;
        for record in records(data, 4)? {
            reserve(&mut hosts, 1)?;
            hosts.push((required_u32(record, 2)?, required_u32(record, 1)?));
        }
    }
    hosts.sort_unstable();
    hosts.dedup();
    Ok(hosts)
}

/// One cell rewrite the topology applies to a table's tiles (record 0329).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CellPatch {
    /// The drawable of the table holding the cell.
    pub(super) info: u64,
    /// The cell's row before the edit.
    pub(super) row: u32,
    /// The cell's column before the edit.
    pub(super) column: u32,
    /// The formula reads `#REF!`: its cached result becomes an error.
    pub(super) error: bool,
    /// The cell moves to this formula-list entry (a shared formula split).
    pub(super) formula_key: Option<u32>,
}

/// One formula cell of a list entry and its rewritten formula.
struct Moved {
    row: u32,
    column: u32,
    formula: Option<Vec<u8>>,
}

/// Rewrite the formulas of `hosts` (cells of the table whose drawable is
/// `info` and data store `store`) in its formula list. `dropped` counts the
/// references of removed cells per formula-list key (record 0329: a
/// removal whose cells the cell writer could not clear).
///
/// An entry whose surviving cells all take one formula is rewritten in
/// place. When they diverge, the entry is split (record 0329): it keeps the
/// formula of its unchanged cells (or the first new formula), and every
/// other new formula takes a new entry its cells move to.
#[allow(
    clippy::too_many_arguments,
    reason = "the formula list of one table, with the plan it feeds"
)]
fn rewrite_formulas(
    package: &Package,
    store: &[u8],
    info: u64,
    hosts: &[Host],
    dropped: &BTreeMap<u64, u64>,
    adjustments: &mut BTreeMap<(u32, u32), Adjustment>,
    output: &mut Plan,
) -> Out<()> {
    if hosts.is_empty() && dropped.is_empty() {
        return Ok(());
    }
    let containers = formula_containers(package, store)?;
    let entries = formula_entries(&containers)?;
    let mut groups: BTreeMap<u64, Vec<Moved>> = BTreeMap::new();
    for host in hosts {
        let key = formula_key(package, store, host.row, host.column)?;
        let entry = entries.get(&key).ok_or(Fault::Invalid)?;
        let mut adjustment = Adjustment::default();
        let mut removed = false;
        let rewritten = rewrite_formula(entry.formula, *host, &mut adjustment, &mut removed)?;
        adjustment.normalize();
        if !adjustment.is_empty() {
            adjustments.insert((host.row, host.column), adjustment);
        }
        if removed {
            output.cell(CellPatch {
                info,
                row: host.row,
                column: host.column,
                error: true,
                formula_key: None,
            })?;
        }
        let list = groups.entry(key).or_default();
        reserve(list, 1)?;
        list.push(Moved {
            row: host.row,
            column: host.column,
            formula: rewritten,
        });
    }
    for key in dropped.keys() {
        groups.entry(*key).or_default();
    }
    let root = &containers[0];
    let mut next_key = single_varint(root.data, 2)?.ok_or(Fault::Invalid)?;
    if let Some(maximum) = entries.keys().next_back() {
        next_key = next_key.max(maximum.checked_add(1).ok_or(Fault::Invalid)?);
    }
    let mut container_rewrites: BTreeMap<usize, BTreeMap<usize, Option<Vec<u8>>>> = BTreeMap::new();
    let mut appended: Vec<Vec<u8>> = Vec::new();
    for (key, cells) in &groups {
        let entry = entries.get(key).ok_or(Fault::Invalid)?;
        let dropped_count = dropped.get(key).copied().unwrap_or(0);
        let surviving = entry
            .refcount
            .checked_sub(dropped_count)
            .ok_or(Fault::Invalid)?;
        let mut outcomes: Vec<(&[u8], Vec<(u32, u32)>)> = Vec::new();
        for cell in cells {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            match outcomes.iter_mut().find(|(known, _)| *known == formula) {
                Some((_, members)) => {
                    reserve(members, 1)?;
                    members.push((cell.row, cell.column));
                },
                None => {
                    reserve(&mut outcomes, 1)?;
                    outcomes.push((formula, vec![(cell.row, cell.column)]));
                },
            }
        }
        let changed: u64 = outcomes
            .iter()
            .map(|(_, members)| members.len() as u64)
            .sum();
        // Cells that keep the entry's formula: unchanged hosts and cells
        // without a dependency record.
        let others = surviving.checked_sub(changed).ok_or(Fault::Invalid)?;
        let raw = w::repeated_length_delimited_payloads(containers[entry.container].data, 3)?;
        let entry_raw = raw.get(entry.position).ok_or(Fault::Invalid)?;
        let with = |formula: Option<&[u8]>, refcount: u64| -> Out<Vec<u8>> {
            let mut data = entry_raw.to_vec();
            if let Some(formula) = formula {
                data = w::transform_length_delimited_field(&data, 5, |_| -> Out<Vec<u8>> {
                    Ok(formula.to_vec())
                })?;
            }
            if refcount != entry.refcount {
                data = w::patch_varint_field(&data, 2, true, Some(refcount))?;
            }
            Ok(data)
        };
        let replacement = if outcomes.is_empty() {
            if dropped_count == 0 {
                continue;
            }
            (surviving > 0).then(|| with(None, surviving)).transpose()?
        } else if others == 0 && outcomes.len() == 1 {
            // Every surviving cell takes the same formula: the entry is
            // rewritten in place. (The host editor moves a cell whose formula
            // becomes another entry's onto that entry; with that entry also
            // rewritten, its cells end up with each other's formulas.
            // Rewriting in place keeps every cell's formula.)
            Some(with(Some(outcomes[0].0), surviving)?)
        } else {
            // The cells diverge: split the entry (record 0329).
            if containers.len() > 1 {
                // A segmented list: new keys would fall in a segment's range.
                return refuse(Unsupported::SharedFormulas);
            }
            let (kept, rest) = if others > 0 {
                (with(None, others)?, &outcomes[..])
            } else {
                let (formula, members) = &outcomes[0];
                (with(Some(formula), members.len() as u64)?, &outcomes[1..])
            };
            for (formula, members) in rest {
                let new_key = next_key;
                next_key = next_key.checked_add(1).ok_or(Fault::Invalid)?;
                let mut data = w::patch_varint_field(entry_raw, 1, true, Some(new_key))?;
                data = w::patch_varint_field(&data, 2, true, Some(members.len() as u64))?;
                data = w::transform_length_delimited_field(&data, 5, |_| -> Out<Vec<u8>> {
                    Ok(formula.to_vec())
                })?;
                reserve(&mut appended, 1)?;
                appended.push(data);
                for (row, column) in members {
                    output.cell(CellPatch {
                        info,
                        row: *row,
                        column: *column,
                        error: false,
                        formula_key: Some(u32_of(new_key)?),
                    })?;
                }
            }
            Some(kept)
        };
        container_rewrites
            .entry(entry.container)
            .or_default()
            .insert(entry.position, replacement);
    }
    if !appended.is_empty() {
        container_rewrites.entry(0).or_default();
    }
    for (index, rewrites) in container_rewrites {
        let container = &containers[index];
        let raw = w::repeated_length_delimited_payloads(container.data, 3)?;
        let mut replacements = Vec::new();
        reserve(&mut replacements, raw.len() + appended.len())?;
        for (position, entry) in raw.iter().enumerate() {
            match rewrites.get(&position) {
                Some(Some(rewritten)) => replacements.push(rewritten.clone()),
                Some(None) => {},
                None => replacements.push(entry.to_vec()),
            }
        }
        let mut data = container.data.to_vec();
        if index == 0 && !appended.is_empty() {
            replacements.extend(appended.iter().cloned());
            data = w::patch_varint_field(&data, 2, true, Some(next_key))?;
        } else if index == 0 {
            // A dropped last key is handed back, as the cell writer does.
            let mut next = single_varint(&data, 2)?.ok_or(Fault::Invalid)?;
            for (position, rewritten) in &rewrites {
                if rewritten.is_some() {
                    continue;
                }
                let entry = raw.get(*position).ok_or(Fault::Invalid)?;
                let key = single_varint(entry, 1)?.ok_or(Fault::Invalid)?;
                if key.checked_add(1) == Some(next) {
                    next = key;
                }
            }
            if Some(next) != single_varint(&data, 2)? {
                data = w::patch_varint_field(&data, 2, true, Some(next))?;
            }
        }
        let data = w::rewrite_repeated_length_delimited_fields(&data, 3, &replacements)?;
        output.push(container.object, container.message, data)?;
    }
    Ok(())
}

/// The data store of the table whose drawable is `info`.
fn table_store(package: &Package, info: u64) -> Out<&[u8]> {
    let info_data = match object_message(package, info, 6_000) {
        Ok((_, data)) => data,
        Err(_) => object_message(package, info, 6_003)?.1,
    };
    let model = single_reference(info_data, 2)?.ok_or(Fault::Invalid)?;
    let (_, model) = object_message(package, model, 6_001)?;
    Ok(single_bytes(model, 4)?.ok_or(Fault::Invalid)?)
}

/// The rewrites of another table's owner whose formulas read the edited
/// table (`internal`, `uid`): their references into it, the precedent edges
/// and ranges naming it; its hosts stay.
fn rewrite_reader(
    package: &Package,
    owner: &Owner<'_>,
    incoming: &Incoming,
    shift: Shift,
    internal: u64,
    uid: (u64, u64),
    output: &mut Plan,
) -> Out<()> {
    let info = single_reference(owner.data, 11)?.ok_or(Fault::Invalid)?;
    let store = table_store(package, info)?;
    let mut hosts = Vec::new();
    for (row, column) in formula_hosts(package, owner.data)? {
        reserve(&mut hosts, 1)?;
        hosts.push(Host {
            row,
            column,
            shift,
            footer: None,
            incoming: Some(uid),
        });
    }
    let mut ignored = BTreeMap::new();
    rewrite_formulas(
        package,
        store,
        info,
        &hosts,
        &BTreeMap::new(),
        &mut ignored,
        output,
    )?;
    let empty = HashSet::new();
    let mut data = owner.data.to_vec();
    if single_bytes(&data, 4)?.is_some() {
        data = w::transform_length_delimited_field(&data, 4, |dependencies| {
            filter_records(dependencies, 1, |record| {
                rewrite_record(record, shift, internal, &ignored, &empty, false)
            })
        })?;
    }
    if incoming.ranges {
        data = w::transform_length_delimited_field(&data, 5, |ranges| {
            filter_records(ranges, 2, |dependency| {
                let Some(reference) = single_bytes(dependency, 4)? else {
                    return Ok(Some(dependency.to_vec()));
                };
                if optional_varint(reference, 1)? != Some(internal) {
                    return Ok(Some(dependency.to_vec()));
                }
                let range = single_bytes(reference, 2)?.ok_or(Fault::Invalid)?;
                let Some(values) = mutate_range(range_coordinate(range)?, shift, None)? else {
                    return Ok(None);
                };
                Ok(Some(w::transform_length_delimited_field(
                    dependency,
                    4,
                    |reference| {
                        w::transform_length_delimited_field(reference, 2, |range| {
                            patch_range_coordinate(range, values)
                        })
                    },
                )?))
            })
        })?;
    }
    if data != owner.data {
        output.push(owner.object, owner.message, data)?;
    }
    for tile in tile_references(owner.data, 13)? {
        let (message, tile_data) = object_message(package, tile, CELL_TILE_TYPE)?;
        let rewritten = filter_records(tile_data, 4, |record| {
            rewrite_record(record, shift, internal, &ignored, &empty, false)
        })?;
        if rewritten != tile_data {
            output.push(tile, message, rewritten)?;
        }
    }
    for tile in &incoming.tiles {
        let (message, tile_data) = object_message(package, *tile, RANGE_TILE_TYPE)?;
        let rewritten = filter_records(tile_data, 2, |entry| {
            let rect = single_bytes(entry, 2)?.ok_or(Fault::Invalid)?;
            let Some(rect) = rewrite_rect(rect, shift, None)? else {
                return Ok(None);
            };
            Ok(Some(w::transform_length_delimited_field(
                entry,
                2,
                |_| -> Out<Vec<u8>> { Ok(rect) },
            )?))
        })?;
        if rewritten != tile_data {
            output.push(*tile, message, rewritten)?;
        }
    }
    Ok(())
}

/// Plan the calculation-engine rewrites of one axis edit.
///
/// With `check_only` (before a removed axis is cleared), formula hosts on the
/// removed axis are skipped, every other rewrite is checked, and nothing is
/// returned.
///
/// The own owner's table extents are not rewritten here; the caller applies
/// [`super::calc::rewrite_extents`] to the returned own-owner payload.
pub(super) fn plan(
    package: &Package,
    table: &Table<'_>,
    shift: Shift,
    check_only: bool,
) -> Out<Plan> {
    let owners = owners(package)?;
    let mut own = None;
    for (index, owner) in owners.iter().enumerate() {
        if let Some(reference) = single_bytes(owner.data, 11)?
            && references::reference_identifier(reference)? == table.info
            && own.replace(index).is_some()
        {
            return Err(Fault::Invalid.into());
        }
    }
    let Some(own_index) = own else {
        return Ok(Plan::default());
    };
    let own = &owners[own_index];
    if !shiftable(own.data)? {
        return refuse(Unsupported::Formulas);
    }
    let internal = single_varint(own.data, 2)?.ok_or(Fault::Invalid)?;
    let uid = single_bytes(own.data, 1)?.ok_or(Fault::Invalid)?;
    let mut proxies = Vec::new();
    for (index, owner) in owners.iter().enumerate() {
        if index == own_index {
            continue;
        }
        if !shiftable(owner.data)? {
            return refuse(Unsupported::Formulas);
        }
        if let Some(incoming) = incoming(package, owner.data, internal, uid)? {
            reserve(&mut proxies, 1)?;
            proxies.push((index, incoming));
        }
    }
    let shift = Shift {
        pending_clear: check_only,
        ..shift
    };
    let mut output = Plan::default();

    // Formula ASTs.
    let store = single_bytes(table.model, 4)?.ok_or(Fault::Invalid)?;
    let hosts = formula_hosts(package, own.data)?;
    let boundary = table.rows.saturating_sub(table.footer_rows);
    let mut adjustments: BTreeMap<(u32, u32), Adjustment> = BTreeMap::new();
    let mut own_hosts = Vec::new();
    for (row, column) in hosts {
        if shift.cleared(row, column) {
            continue;
        }
        reserve(&mut own_hosts, 1)?;
        own_hosts.push(Host {
            row,
            column,
            shift,
            footer: (table.footer_rows > 0 && row >= boundary).then_some(Footer {
                boundary,
                expand_on_insert: shift.expand_footer,
            }),
            incoming: None,
        });
    }
    rewrite_formulas(
        package,
        store,
        table.info,
        &own_hosts,
        table.dropped_formulas,
        &mut adjustments,
        &mut output,
    )?;

    // The own owner.
    let range_hosts = range_dependency_hosts(package, own.data)?;
    let mut owner = own.data.to_vec();
    if single_bytes(&owner, 4)?.is_some() {
        owner = w::transform_length_delimited_field(&owner, 4, |dependencies| {
            filter_records(dependencies, 1, |record| {
                rewrite_record(record, shift, internal, &adjustments, &range_hosts, true)
            })
        })?;
    }
    if single_bytes(&owner, 5)?.is_some() {
        owner = w::transform_length_delimited_field(&owner, 5, |ranges| {
            filter_records(ranges, 2, |dependency| {
                rewrite_back_dependency(dependency, shift, internal, &adjustments)
            })
        })?;
    }
    if single_bytes(&owner, 14)?.is_some() {
        owner = w::transform_length_delimited_field(&owner, 14, |references| {
            rewrite_uuid_references(references, shift)
        })?;
    }
    output.own = Some((own.object, own.message, owner));

    // Dependency-record tiles.
    let mut empty_tiles: Vec<(u32, u64)> = Vec::new();
    for tile in tile_references(own.data, 13)? {
        let (message, data) = object_message(package, tile, CELL_TILE_TYPE)?;
        if single_varint(data, 1)? != Some(internal) {
            return Err(Fault::Invalid.into());
        }
        let (begin_field, size) = if shift.row_axis {
            (3, DEPENDENCY_TILE_ROWS)
        } else {
            (2, DEPENDENCY_TILE_COLUMNS)
        };
        let begin = required_u32(data, begin_field)?;
        let rewritten = filter_records(data, 4, |record| {
            let Some(rewritten) =
                rewrite_record(record, shift, internal, &adjustments, &range_hosts, true)?
            else {
                return Ok(None);
            };
            let coordinate = required_u32(&rewritten, if shift.row_axis { 2 } else { 1 })?;
            if coordinate / size * size != begin {
                return refuse(Unsupported::Formulas);
            }
            Ok(Some(rewritten))
        })?;
        if repeated_bytes(&rewritten, 4)?.is_empty() && !repeated_bytes(data, 4)?.is_empty() {
            // Its last record went with a dropped host (record 0329): the
            // tile goes too, as the cell writer drops it.
            empty_tiles.push((13, tile));
        } else if rewritten != data {
            output.push(tile, message, rewritten)?;
        }
    }

    // Range-precedent tiles.
    for tile in tile_references(own.data, 15)? {
        let (message, data) = object_message(package, tile, RANGE_TILE_TYPE)?;
        let to_own = single_varint(data, 1)? == Some(internal);
        let rewritten = filter_records(data, 2, |entry| {
            rewrite_from_to(entry, shift, to_own, &adjustments)
        })?;
        if repeated_bytes(&rewritten, 2)?.is_empty() && !repeated_bytes(data, 2)?.is_empty() {
            empty_tiles.push((15, tile));
        } else if rewritten != data {
            output.push(tile, message, rewritten)?;
        }
    }
    if !empty_tiles.is_empty() {
        let mut owner = output.own.take().ok_or(Fault::Invalid)?;
        for (field, tile) in &empty_tiles {
            owner.2 = w::transform_length_delimited_field(&owner.2, *field, |tiles| {
                filter_records(tiles, 1, |reference| {
                    Ok((references::reference_identifier(reference)? != *tile)
                        .then(|| reference.to_vec()))
                })
            })?;
            reserve(&mut output.removed, 1)?;
            output.removed.push(*tile);
        }
        output.own = Some(owner);
    }

    // Other tables' formulas that read the table.
    let uid_key = (
        single_varint(uid, 1)?.ok_or(Fault::Invalid)?,
        single_varint(uid, 2)?.ok_or(Fault::Invalid)?,
    );
    for (index, incoming) in proxies.iter().filter(|(_, incoming)| incoming.reader) {
        rewrite_reader(
            package,
            &owners[*index],
            incoming,
            shift,
            internal,
            uid_key,
            &mut output,
        )?;
    }

    // Merge-proxy owners' ranges into the table.
    for (index, incoming) in proxies.into_iter().filter(|(_, incoming)| !incoming.reader) {
        let proxy = &owners[index];
        let mut data = proxy.data.to_vec();
        if incoming.ranges {
            data = w::transform_length_delimited_field(&data, 5, |ranges| {
                rewrite_records(ranges, 2, |dependency| {
                    let Some(reference) = single_bytes(dependency, 4)? else {
                        return Ok(dependency.to_vec());
                    };
                    if optional_varint(reference, 1)? != Some(internal) {
                        return Ok(dependency.to_vec());
                    }
                    let range = single_bytes(reference, 2)?.ok_or(Fault::Invalid)?;
                    let values = incoming_range(range_coordinate(range)?, shift)?;
                    w::transform_length_delimited_field(dependency, 4, |reference| {
                        w::transform_length_delimited_field(reference, 2, |range| {
                            patch_range_coordinate(range, values)
                        })
                    })
                })
            })?;
        }
        if data != proxy.data {
            output.push(proxy.object, proxy.message, data)?;
        }
        for tile in incoming.tiles {
            let (message, tile_data) = object_message(package, tile, RANGE_TILE_TYPE)?;
            let rewritten = rewrite_records(tile_data, 2, |entry| {
                w::transform_length_delimited_field(entry, 2, |rect| {
                    rewrite_incoming_rect(rect, shift)
                })
            })?;
            if rewritten != tile_data {
                output.push(tile, message, rewritten)?;
            }
        }
    }
    if check_only {
        return Ok(Plan::default());
    }
    Ok(output)
}

/// `mutate_incoming_target_range_coordinate` without retained hosts.
fn incoming_range(range: [u32; 4], shift: Shift) -> Out<[u32; 4]> {
    let [left, top, right, bottom] = range;
    let output = if shift.row_axis {
        [
            left,
            shift.coordinate(top)?,
            right,
            shift.coordinate(bottom)?,
        ]
    } else {
        [
            shift.coordinate(left)?,
            top,
            shift.coordinate(right)?,
            bottom,
        ]
    };
    if output[1] > output[3] || output[0] > output[2] {
        return Err(Fault::Invalid.into());
    }
    Ok(output)
}

/// `mutate_incoming_target_cell_rect` without retained hosts.
fn rewrite_incoming_rect(rect: &[u8], shift: Shift) -> Out<Vec<u8>> {
    let origin = single_bytes(rect, 1)?.ok_or(Fault::Invalid)?;
    let size = single_bytes(rect, 2)?.ok_or(Fault::Invalid)?;
    let (top, left) = explicit_coordinate(origin)?;
    let previous_columns = optional_varint(size, 1)?;
    let previous_rows = optional_varint(size, 2)?;
    let columns = u32_of(previous_columns.unwrap_or(1))?;
    let rows = u32_of(previous_rows.unwrap_or(1))?;
    if rows == 0 || columns == 0 {
        return Err(Fault::Invalid.into());
    }
    let [left, top, right, bottom] = incoming_range(
        [
            left,
            top,
            left.checked_add(columns - 1).ok_or(Fault::Invalid)?,
            top.checked_add(rows - 1).ok_or(Fault::Invalid)?,
        ],
        shift,
    )?;
    let columns = right - left + 1;
    let rows = bottom - top + 1;
    let output =
        w::transform_length_delimited_field(rect, 1, |origin| patch_coordinate(origin, top, left))?;
    w::transform_length_delimited_field(&output, 2, |size| -> Out<Vec<u8>> {
        let size = w::patch_varint_field(
            size,
            1,
            previous_columns.is_some(),
            (columns != 1).then_some(u64::from(columns)),
        )?;
        Ok(w::patch_varint_field(
            &size,
            2,
            previous_rows.is_some(),
            (rows != 1).then_some(u64::from(rows)),
        )?)
    })
}

/// The planned rewrites: the own owner's payload (before its extents), and
/// every other payload.
#[derive(Default)]
pub(super) struct Plan {
    pub(super) own: Option<(u64, usize, Vec<u8>)>,
    pub(super) rewrites: Vec<Rewrite>,
    /// Cells whose tiles change with the formulas (record 0329): error
    /// caches of `#REF!` formulas, cells of split shared formulas.
    pub(super) cells: Vec<CellPatch>,
    /// Dependency tiles emptied by dropped hosts, which go with their
    /// owner's references to them (record 0329).
    pub(super) removed: Vec<u64>,
}

impl Plan {
    fn cell(&mut self, patch: CellPatch) -> Out<()> {
        reserve(&mut self.cells, 1)?;
        self.cells.push(patch);
        Ok(())
    }

    fn push(&mut self, object: u64, message: usize, data: Vec<u8>) -> Out<()> {
        reserve(&mut self.rewrites, 1)?;
        self.rewrites.push(Rewrite {
            object,
            message,
            data,
        });
        Ok(())
    }
}
