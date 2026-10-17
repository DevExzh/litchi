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
//! - [`Unsupported::DeletedReference`]: a surviving formula reads a cell of
//!   the removed row or column, or a range starts or ends there. Native
//!   Numbers turns such a reference into `#REF!`; like the host editor, this
//!   transaction does not write reference-error nodes.
//! - [`Unsupported::ReferencedByOtherOwners`]: a formula in another table, or
//!   a chart, reads the table (its dependency edges, legacy UUID ranges or
//!   UUID references name the table).
//! - [`Unsupported::Formulas`]: volatile, spanning, whole-owner, error or
//!   spill dependency state in any owner, a category reference whose host
//!   moves, a packed coordinate, a dependency record that would cross its
//!   tile.
//! - [`Unsupported::SharedFormulas`]: one formula-list entry shared by cells
//!   whose rewritten formulas would differ (the host editor splits the entry;
//!   not ported).

use std::collections::{BTreeMap, HashMap, HashSet};

use litchi_iwa_common::wire as w;

use super::super::Package;
use super::super::structure::{Fault, map_wire, references, reserve};
use super::storage::bnc_cells;
use super::wire::{
    bytes_field, fields, push_varint, rebuild, repeated_bytes, repeated_varints, single_bytes,
    single_reference, single_varint,
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
        self.pending_clear && !self.insert && self.axis(row, column) == self.position
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

/// `mutate_range_coordinate`, with the footer override.
fn mutate_range(range: [u32; 4], shift: Shift, adjustment: Option<&Adjustment>) -> Out<[u32; 4]> {
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
    Ok(if shift.row_axis {
        [
            left,
            shift.coordinate(top)?,
            right,
            match footer {
                Some(value) => value,
                None => shift.coordinate(bottom)?,
            },
        ]
    } else {
        [
            shift.coordinate(left)?,
            top,
            shift.coordinate(right)?,
            bottom,
        ]
    })
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

/// `mutate_cell_rect` and its wire rewrite.
fn rewrite_rect(rect: &[u8], shift: Shift, adjustment: Option<&Adjustment>) -> Out<Vec<u8>> {
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
    let [left, top, right, bottom] = mutate_range(range, shift, adjustment)?;
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

/// `mutate_tract_interval`.
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
) -> Out<(u32, u32)> {
    if cross_table {
        return Ok((begin, end));
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
    let shifted_high = if footer_change {
        if shift.insert {
            shift.position
        } else {
            shift.position.checked_sub(1).ok_or(Fault::Invalid)?
        }
    } else {
        shift.coordinate(high)?
    };
    let shifted_low = shift.coordinate(low)?;
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
    Ok(if begin <= end {
        (shifted_low, shifted_high)
    } else {
        (shifted_high, shifted_low)
    })
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

/// `rewrite_colon_tract` and its wire rewrite (`rewrite_colon_tract_wire`).
fn rewrite_tract(
    tract: &[u8],
    host: Host,
    cross_table: bool,
    adjustment: &mut Adjustment,
) -> Out<Option<Vec<u8>>> {
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
    if let Some(ranges) = current.get_mut(&relative_field) {
        for range in ranges.iter_mut() {
            let begin = resolve_relative(anchor, range.begin)?;
            let end = match range.end {
                Some(end) => resolve_relative(anchor, end)?,
                None => begin,
            };
            let (shifted_begin, shifted_end) = mutate_interval(
                begin,
                end,
                cross_table,
                shift,
                host.footer,
                &columns,
                adjustment,
            )?;
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
    if !keep_sentinel && let Some(ranges) = current.get_mut(&absolute_field) {
        for range in ranges.iter_mut() {
            let begin = u32_of(u64::try_from(range.begin).map_err(|_error| Fault::Invalid)?)?;
            let end = match range.end {
                Some(end) => u32_of(u64::try_from(end).map_err(|_error| Fault::Invalid)?)?,
                None => begin,
            };
            let (shifted_begin, shifted_end) = mutate_interval(
                begin,
                end,
                cross_table,
                shift,
                host.footer,
                &columns,
                adjustment,
            )?;
            range.begin = i64::from(shifted_begin);
            range.end = (shifted_begin != shifted_end).then_some(i64::from(shifted_end));
        }
    }
    if current == previous {
        return Ok(None);
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
    Ok(Some(output))
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

/// `rewrite_formula_nodes` and its wire rewrite (`rewrite_ast_array_wire`):
/// the rewritten `ASTNodeArrayArchive`, or `None` when nothing moves.
fn rewrite_array(array: &[u8], host: Host, adjustment: &mut Adjustment) -> Out<Option<Vec<u8>>> {
    let shift = host.shift;
    let nodes = w::repeated_length_delimited_payloads(array, 1)?;
    let anchor = host.axis();
    let shifted_host = host.shifted()?;
    let offset = host.cross_table_offset()?;
    let overrides = footer_overrides(&nodes, host, adjustment)?;
    let coordinate_field = if shift.row_axis { 27 } else { 26 };
    let mut changed = false;
    let mut replacements = Vec::new();
    reserve(&mut replacements, nodes.len())?;
    for (index, node) in nodes.iter().enumerate() {
        let node: &[u8] = node;
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
                let shifted_target = match overrides.get(&index) {
                    Some(value) => *value,
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
            && let Some(rewritten) = rewrite_array(thunk, host, adjustment)?
        {
            output = w::transform_length_delimited_field(&output, 14, |_| -> Out<Vec<u8>> {
                Ok(rewritten)
            })?;
        }
        if let Some(tract) = single_bytes(node, 40)?
            && let Some(rewritten) = rewrite_tract(tract, host, cross_table, adjustment)?
        {
            output = w::transform_length_delimited_field(&output, 40, |_| -> Out<Vec<u8>> {
                Ok(rewritten)
            })?;
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
) -> Out<Option<Vec<u8>>> {
    let array = single_bytes(formula, 1)?.ok_or(Fault::Invalid)?;
    let Some(rewritten) = rewrite_array(array, host, adjustment)? else {
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
        for (stored_column, bytes) in bnc_cells(stored_row)? {
            if stored_column != column {
                continue;
            }
            let cell = litchi_numbers_wire::BncCell::parse(bytes)
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
/// (`rewrite_shifted_dependency_records`, one record).
fn rewrite_record(
    record: &[u8],
    shift: Shift,
    internal: u64,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
    range_hosts: &HashSet<(u32, u32)>,
    host_moves: bool,
) -> Out<Vec<u8>> {
    let column = required_u32(record, 1)?;
    let row = required_u32(record, 2)?;
    if host_moves && shift.cleared(row, column) {
        return Ok(record.to_vec());
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
        return Ok(output);
    };
    let mut local_rows = repeated_varints(edges, 1)?;
    let mut local_columns = repeated_varints(edges, 2)?;
    let mut owned_rows = repeated_varints(edges, 3)?;
    let mut owned_columns = repeated_varints(edges, 4)?;
    let owners = repeated_varints(edges, 5)?;
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
        let local = if shift.row_axis {
            &mut local_rows
        } else {
            &mut local_columns
        };
        for coordinate in local.iter_mut() {
            *coordinate = u64::from(shift.coordinate(u32_of(*coordinate)?)?);
        }
    }
    let owned = if shift.row_axis {
        &mut owned_rows
    } else {
        &mut owned_columns
    };
    for (coordinate, owner) in owned.iter_mut().zip(&owners) {
        if *owner == internal {
            *coordinate = u64::from(shift.coordinate(u32_of(*coordinate)?)?);
        }
    }
    output = w::transform_length_delimited_field(&output, 6, |edges| -> Out<Vec<u8>> {
        let mut data = edges.to_vec();
        for (number, values) in [
            (1, &local_rows),
            (2, &local_columns),
            (3, &owned_rows),
            (4, &owned_columns),
        ] {
            data = write_varints(&data, number, values)?;
        }
        Ok(data)
    })?;
    Ok(output)
}

fn rewrite_records(
    container: &[u8],
    field: u32,
    mut map: impl FnMut(&[u8]) -> Out<Vec<u8>>,
) -> Out<Vec<u8>> {
    let raw = w::repeated_length_delimited_payloads(container, field)?;
    let mut replacements = Vec::new();
    reserve(&mut replacements, raw.len())?;
    for record in raw {
        replacements.push(map(record)?);
    }
    Ok(w::rewrite_repeated_length_delimited_fields(
        container,
        field,
        &replacements,
    )?)
}

/// `rewrite_range_back_dependency` for one own range dependency.
fn rewrite_back_dependency(
    dependency: &[u8],
    shift: Shift,
    internal: u64,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
) -> Out<Vec<u8>> {
    let row = required_u32(dependency, 1)?;
    let column = required_u32(dependency, 2)?;
    if shift.cleared(row, column) {
        return Ok(dependency.to_vec());
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
        let values = mutate_range(
            range_coordinate(range)?,
            shift,
            adjustments.get(&(row, column)),
        )?;
        output = w::transform_length_delimited_field(&output, 4, |reference| -> Out<Vec<u8>> {
            w::transform_length_delimited_field(reference, 2, |range| {
                patch_range_coordinate(range, values)
            })
        })?;
    }
    Ok(output)
}

/// One range-precedent tile entry (`FromToRangeArchive`).
fn rewrite_from_to(
    entry: &[u8],
    shift: Shift,
    to_own: bool,
    adjustments: &BTreeMap<(u32, u32), Adjustment>,
) -> Out<Vec<u8>> {
    let from = single_bytes(entry, 1)?.ok_or(Fault::Invalid)?;
    let (row, column) = explicit_coordinate(from)?;
    if shift.cleared(row, column) {
        return Ok(entry.to_vec());
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
        output = w::transform_length_delimited_field(&output, 2, |rect| {
            rewrite_rect(rect, shift, adjustment)
        })?;
    }
    Ok(output)
}

/// `mutate_cell_coord_set` and its wire rewrite.
fn rewrite_coord_set(set: &[u8], shift: Shift) -> Out<Vec<u8>> {
    if shift.pending_clear && !shift.insert {
        // The hosts on the removed axis are cleared first; the coordinates
        // are checked again after the clear.
        return Ok(set.to_vec());
    }
    rewrite_records(set, 1, |entry| {
        let column = required_u32(entry, 1)?;
        let column = if shift.row_axis {
            column
        } else {
            shift.coordinate(column)?
        };
        let output = w::patch_varint_field(entry, 1, true, Some(u64::from(column)))?;
        w::transform_length_delimited_field(&output, 2, |rows| {
            rewrite_records(rows, 1, |range| {
                let begin = int32_of(optional_varint(range, 1)?.unwrap_or(0))?;
                let end = optional_varint(range, 2)?.map(int32_of).transpose()?;
                let (begin, new_end) = if shift.row_axis {
                    let shift_row = |value: i32| -> Out<i32> {
                        let value = u32::try_from(value).map_err(|_error| Fault::Invalid)?;
                        i32::try_from(shift.coordinate(value)?)
                            .map_err(|_error| Failure::Fault(Fault::Invalid))
                    };
                    (shift_row(begin)?, end.map(shift_row).transpose()?)
                } else {
                    (begin, end)
                };
                let output = w::patch_varint_field(range, 1, true, Some(int32_wire(begin)))?;
                Ok(w::patch_varint_field(
                    &output,
                    2,
                    end.is_some(),
                    new_end.map(int32_wire),
                )?)
            })
        })
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

/// Rewrite the formulas of `hosts` (cells of the table whose data store is
/// `store`) in its formula list.
fn rewrite_formulas(
    package: &Package,
    store: &[u8],
    hosts: &[Host],
    adjustments: &mut BTreeMap<(u32, u32), Adjustment>,
    output: &mut Plan,
) -> Out<()> {
    if hosts.is_empty() {
        return Ok(());
    }
    let containers = formula_containers(package, store)?;
    let entries = formula_entries(&containers)?;
    // key -> rewritten formulas, and the count of cells left as they are.
    let mut jobs: BTreeMap<u64, Vec<Vec<u8>>> = BTreeMap::new();
    let mut unchanged: BTreeMap<u64, usize> = BTreeMap::new();
    for host in hosts {
        let key = formula_key(package, store, host.row, host.column)?;
        let entry = entries.get(&key).ok_or(Fault::Invalid)?;
        let mut adjustment = Adjustment::default();
        let rewritten = rewrite_formula(entry.formula, *host, &mut adjustment)?;
        adjustment.normalize();
        if !adjustment.is_empty() {
            adjustments.insert((host.row, host.column), adjustment);
        }
        match rewritten {
            Some(current) => {
                let list = jobs.entry(key).or_default();
                reserve(list, 1)?;
                list.push(current);
            },
            None => *unchanged.entry(key).or_default() += 1,
        }
    }
    let mut container_rewrites: BTreeMap<usize, BTreeMap<usize, Vec<u8>>> = BTreeMap::new();
    for (key, changed) in &jobs {
        let entry = &entries[key];
        let kept = unchanged.get(key).copied().unwrap_or(0);
        let all_rewritten =
            u64::try_from(changed.len() + kept).ok() == Some(entry.refcount) && kept == 0;
        let first = &changed[0];
        let shared = changed.iter().all(|current| current == first);
        // Every cell holding the entry moves to the same formula: the entry
        // is rewritten in place. (The host editor moves a cell whose formula
        // becomes another entry's onto that entry; with that entry also
        // rewritten, its cells end up with each other's formulas. Rewriting
        // in place keeps every cell's formula.)
        if !(entry.refcount == 1 || (all_rewritten && shared)) {
            return refuse(Unsupported::SharedFormulas);
        }
        let raw = w::repeated_length_delimited_payloads(containers[entry.container].data, 3)?;
        let entry_raw = raw.get(entry.position).ok_or(Fault::Invalid)?;
        let rewritten = w::transform_length_delimited_field(entry_raw, 5, |_| -> Out<Vec<u8>> {
            Ok(first.clone())
        })?;
        container_rewrites
            .entry(entry.container)
            .or_default()
            .insert(entry.position, rewritten);
    }
    for (index, rewrites) in container_rewrites {
        let container = &containers[index];
        let raw = w::repeated_length_delimited_payloads(container.data, 3)?;
        let mut replacements = Vec::new();
        reserve(&mut replacements, raw.len())?;
        for (position, entry) in raw.into_iter().enumerate() {
            replacements.push(match rewrites.get(&position) {
                Some(rewritten) => rewritten.clone(),
                None => entry.to_vec(),
            });
        }
        let data = w::rewrite_repeated_length_delimited_fields(container.data, 3, &replacements)?;
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
    rewrite_formulas(package, store, &hosts, &mut ignored, output)?;
    let empty = HashSet::new();
    let mut data = owner.data.to_vec();
    if single_bytes(&data, 4)?.is_some() {
        data = w::transform_length_delimited_field(&data, 4, |dependencies| {
            rewrite_records(dependencies, 1, |record| {
                rewrite_record(record, shift, internal, &ignored, &empty, false)
            })
        })?;
    }
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
                let values = mutate_range(range_coordinate(range)?, shift, None)?;
                w::transform_length_delimited_field(dependency, 4, |reference| {
                    w::transform_length_delimited_field(reference, 2, |range| {
                        patch_range_coordinate(range, values)
                    })
                })
            })
        })?;
    }
    if data != owner.data {
        output.push(owner.object, owner.message, data)?;
    }
    for tile in tile_references(owner.data, 13)? {
        let (message, tile_data) = object_message(package, tile, CELL_TILE_TYPE)?;
        let rewritten = rewrite_records(tile_data, 4, |record| {
            rewrite_record(record, shift, internal, &ignored, &empty, false)
        })?;
        if rewritten != tile_data {
            output.push(tile, message, rewritten)?;
        }
    }
    for tile in &incoming.tiles {
        let (message, tile_data) = object_message(package, *tile, RANGE_TILE_TYPE)?;
        let rewritten = rewrite_records(tile_data, 2, |entry| {
            w::transform_length_delimited_field(entry, 2, |rect| rewrite_rect(rect, shift, None))
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
    rewrite_formulas(package, store, &own_hosts, &mut adjustments, &mut output)?;

    // The own owner.
    let range_hosts = range_dependency_hosts(package, own.data)?;
    let mut owner = own.data.to_vec();
    if single_bytes(&owner, 4)?.is_some() {
        owner = w::transform_length_delimited_field(&owner, 4, |dependencies| {
            rewrite_records(dependencies, 1, |record| {
                rewrite_record(record, shift, internal, &adjustments, &range_hosts, true)
            })
        })?;
    }
    if single_bytes(&owner, 5)?.is_some() {
        owner = w::transform_length_delimited_field(&owner, 5, |ranges| {
            rewrite_records(ranges, 2, |dependency| {
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
        let rewritten = rewrite_records(data, 4, |record| {
            let rewritten =
                rewrite_record(record, shift, internal, &adjustments, &range_hosts, true)?;
            let coordinate = required_u32(&rewritten, if shift.row_axis { 2 } else { 1 })?;
            if coordinate / size * size != begin {
                return refuse(Unsupported::Formulas);
            }
            Ok(rewritten)
        })?;
        if rewritten != data {
            output.push(tile, message, rewritten)?;
        }
    }

    // Range-precedent tiles.
    for tile in tile_references(own.data, 15)? {
        let (message, data) = object_message(package, tile, RANGE_TILE_TYPE)?;
        let to_own = single_varint(data, 1)? == Some(internal);
        let rewritten = rewrite_records(data, 2, |entry| {
            rewrite_from_to(entry, shift, to_own, &adjustments)
        })?;
        if rewritten != data {
            output.push(tile, message, rewritten)?;
        }
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
}

impl Plan {
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
