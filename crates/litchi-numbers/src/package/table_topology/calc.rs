//! The CalculationEngine side of a table-axis edit.
//!
//! An axis edit moves every cell after it, so every formula coordinate in
//! the table's own formula owner and every reference other owners hold into
//! it would move too. The host editor shifts formula ASTs and dependency
//! records (`formula_dependency_shift`); that port is a follow-up. This
//! phase admits the tables whose calculation graph needs no coordinate
//! rewrite: the table's own owner (the `TSCE.FormulaOwnerDependenciesArchive`
//! whose `formula_owner` is the table drawable) holds no cell, range,
//! volatile, spanning, whole-owner, error, UUID or spill dependency, and no
//! other owner holds an edge, range, whole-owner, spanning or UUID
//! reference into it. The owner's table-extent ranges (the spanning
//! dependencies' total and body ranges) follow the new geometry.
//!
//! The table's sub-owners (conditional styles, hidden states, merges, group
//! by, spill, haunted) address their own coordinate spaces; like the host
//! editor, this phase does not rewrite them. Merges are refused by the
//! caller.

use super::super::Package;
use super::super::structure::{Fault, FaultResult, references, reserve};
use super::wire::{
    bytes_field, fields, rebuild, repeated_bytes, repeated_varints, single_bytes, single_varint,
    varint_field,
};

/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `TSCE.CellRecordTileArchive`.
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;
/// `TSCE.RangePrecedentsTileArchive`.
const RANGE_PRECEDENTS_TILE_MESSAGE_TYPE: u32 = 4_010;
/// The column sentinel of an unset `TSCE.RangeCoordinateArchive`.
const SENTINEL_COLUMN: u64 = 0x7fff;
/// The row sentinel of an unset `TSCE.RangeCoordinateArchive`.
const SENTINEL_ROW: u64 = 0x7fff_ffff;

/// Why the calculation graph refuses an axis edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Refusal {
    /// The table's own owner holds formula or dependency state.
    OwnFormulas,
    /// Another owner (a formula in another table, a chart) reads the table.
    Incoming,
}

/// The table's own formula owner, located for its extent rewrite.
pub(super) struct OwnOwner {
    pub(super) object: u64,
    pub(super) message: usize,
}

/// One table's table-sized extents: `[first column, first row, last
/// column, last row]` of the whole table and of its body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Extents {
    pub(super) total: [u64; 4],
    pub(super) body: [u64; 4],
}

impl Extents {
    pub(super) fn of(
        rows: u32,
        columns: u32,
        header_rows: u32,
        header_columns: u32,
        footer_rows: u32,
    ) -> Self {
        let last_column = u64::from(columns.saturating_sub(1));
        let last_row = u64::from(rows.saturating_sub(1));
        Self {
            total: [0, 0, last_column, last_row],
            body: [
                u64::from(header_columns),
                u64::from(header_rows),
                last_column,
                u64::from(rows.saturating_sub(footer_rows).saturating_sub(1)),
            ],
        }
    }
}

/// Whether a payload holds nothing but empty sub-messages and no other
/// field (an empty `repeated` container).
fn inert_container(payload: &[u8]) -> FaultResult<bool> {
    for field in fields(payload)? {
        if field.wire_type != 2 || !field.payload.is_empty() {
            return Ok(false);
        }
    }
    Ok(true)
}

/// Every cell record the owner holds, inline or in its record tiles.
fn cell_records<'p>(
    package: &'p Package,
    owner: &'p [u8],
    object: &impl Fn(u64) -> FaultResult<Option<&'p [u8]>>,
) -> FaultResult<Vec<&'p [u8]>> {
    let mut records = Vec::new();
    if let Some(inline) = single_bytes(owner, 4)? {
        for record in repeated_bytes(inline, 1)? {
            reserve(&mut records, 1)?;
            records.push(record);
        }
    }
    if let Some(tiled) = single_bytes(owner, 13)? {
        for reference in repeated_bytes(tiled, 1)? {
            let identifier = references::reference_identifier(reference)?;
            let tile = object(identifier)?.ok_or(Fault::Invalid)?;
            for record in repeated_bytes(tile, 4)? {
                reserve(&mut records, 1)?;
                records.push(record);
            }
        }
    }
    let _ = package;
    Ok(records)
}

/// Whether one owner's records, ranges or references address `internal`
/// (the table owner's internal identifier) or `uuid`.
fn reads_table<'p>(
    package: &'p Package,
    owner: &'p [u8],
    internal: u64,
    uuid: &[u8],
    object: &impl Fn(u64) -> FaultResult<Option<&'p [u8]>>,
) -> FaultResult<bool> {
    for record in cell_records(package, owner, object)? {
        if let Some(edges) = single_bytes(record, 6)?
            && repeated_varints(edges, 5)?.contains(&internal)
        {
            return Ok(true);
        }
        // Legacy (unexpanded) edges carry their owner as a UUID.
        if let Some(edges) = single_bytes(record, 9)?
            && repeated_varints(edges, 12)?.contains(&internal)
        {
            return Ok(true);
        }
    }
    if let Some(ranges) = single_bytes(owner, 5)? {
        for dependency in repeated_bytes(ranges, 2)? {
            if single_bytes(dependency, 3)?.is_some() {
                // A legacy UUID-addressed range: conservatively a reader.
                return Ok(true);
            }
            if let Some(internal_range) = single_bytes(dependency, 4)?
                && single_varint(internal_range, 1)? == Some(internal)
            {
                return Ok(true);
            }
        }
    }
    if let Some(tiled) = single_bytes(owner, 15)? {
        for reference in repeated_bytes(tiled, 1)? {
            let identifier = references::reference_identifier(reference)?;
            let tile = object(identifier)?.ok_or(Fault::Invalid)?;
            if single_varint(tile, 1)? == Some(internal) {
                return Ok(true);
            }
        }
    }
    if let Some(whole) = single_bytes(owner, 9)?
        && let Some(cells) = single_bytes(whole, 1)?
    {
        for entry in repeated_bytes(cells, 1)? {
            if single_varint(entry, 1)? == Some(internal) {
                return Ok(true);
            }
        }
    }
    for number in [7, 8] {
        if let Some(spanning) = single_bytes(owner, number)? {
            for coordinate in repeated_bytes(spanning, 1)? {
                for by_table in repeated_bytes(coordinate, 2)? {
                    if single_varint(by_table, 1)? == Some(internal) {
                        return Ok(true);
                    }
                }
            }
        }
    }
    if let Some(uuids) = single_bytes(owner, 14)? {
        for number in [1, 2] {
            for table_ref in repeated_bytes(uuids, number)? {
                if single_bytes(table_ref, 1)? == Some(uuid) {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

/// Whether the table's own owner is free of formula and dependency state.
fn own_owner_is_inert<'p>(
    package: &'p Package,
    owner: &'p [u8],
    object: &impl Fn(u64) -> FaultResult<Option<&'p [u8]>>,
) -> FaultResult<bool> {
    if !cell_records(package, owner, object)?.is_empty() {
        return Ok(false);
    }
    for field in fields(owner)? {
        let inert = match field.number {
            // Identity, internal identifier, kind, owner reference, base
            // owner: no coordinates.
            1 | 2 | 3 | 11 | 12 => true,
            // Cell records (checked above): the containers themselves.
            4 => repeated_bytes(field.payload, 1)?.is_empty(),
            13 => true,
            // Range, volatile, whole-owner, error, UUID and spill state.
            5 | 10 | 14 | 16 => field.payload.is_empty(),
            6 => inert_container(field.payload)?,
            9 => match single_bytes(field.payload, 1)? {
                None => true,
                Some(cells) => repeated_bytes(cells, 1)?.is_empty(),
            },
            15 => repeated_bytes(field.payload, 1)?.is_empty(),
            // Spanning dependencies: only the table extents, no spans.
            7 | 8 => repeated_bytes(field.payload, 1)?.is_empty(),
            _ => false,
        };
        if !inert {
            return Ok(false);
        }
    }
    Ok(true)
}

/// Check the calculation graph of the table whose drawable is `info` and
/// locate its own owner.
///
/// Returns the own owner (`None` when the table has none) or the refusal.
pub(super) fn check(
    package: &Package,
    info: u64,
) -> FaultResult<Result<Option<OwnOwner>, Refusal>> {
    let mut owners: Vec<(u64, usize, &[u8])> = Vec::new();
    for component in package.state.components.catalog().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                    reserve(&mut owners, 1)?;
                    owners.push((
                        object.archive_info.identifier.ok_or(Fault::Invalid)?,
                        index,
                        &message.data,
                    ));
                }
            }
        }
    }
    let object = |identifier: u64| -> FaultResult<Option<&[u8]>> {
        for component in package.state.components.catalog().iter() {
            if let Some(object) = component.archive().object(identifier) {
                let mut found = None;
                for message in &object.messages {
                    if matches!(
                        message.type_,
                        CELL_RECORD_TILE_MESSAGE_TYPE | RANGE_PRECEDENTS_TILE_MESSAGE_TYPE
                    ) && found.replace(message.data.as_slice()).is_some()
                    {
                        return Err(Fault::Invalid);
                    }
                }
                return Ok(found);
            }
        }
        Ok(None)
    };
    let mut own = None;
    for (identifier, index, data) in &owners {
        if let Some(reference) = single_bytes(data, 11)?
            && references::reference_identifier(reference)? == info
            && own.replace((*identifier, *index, *data)).is_some()
        {
            return Err(Fault::Invalid);
        }
    }
    let Some((own_object, own_message, own_data)) = own else {
        return Ok(Ok(None));
    };
    if !own_owner_is_inert(package, own_data, &object)? {
        return Ok(Err(Refusal::OwnFormulas));
    }
    let internal = single_varint(own_data, 2)?.ok_or(Fault::Invalid)?;
    let uuid = single_bytes(own_data, 1)?.ok_or(Fault::Invalid)?;
    for (identifier, _, data) in &owners {
        if *identifier == own_object {
            continue;
        }
        if reads_table(package, data, internal, uuid, &object)? {
            return Ok(Err(Refusal::Incoming));
        }
    }
    Ok(Ok(Some(OwnOwner {
        object: own_object,
        message: own_message,
    })))
}

fn range(payload: &[u8]) -> FaultResult<[u64; 4]> {
    let mut output = [0; 4];
    for (slot, number) in output.iter_mut().zip(1_u32..) {
        *slot = single_varint(payload, number)?.ok_or(Fault::Invalid)?;
    }
    if fields(payload)?.len() != 4 {
        return Err(Fault::Unsupported);
    }
    Ok(output)
}

fn encode_range(values: [u64; 4]) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    for (number, value) in (1_u32..).zip(values) {
        output.extend(varint_field(number, value)?);
    }
    Ok(output)
}

fn is_sentinel(values: [u64; 4]) -> bool {
    values == [SENTINEL_COLUMN, SENTINEL_ROW, SENTINEL_COLUMN, SENTINEL_ROW]
}

/// Rewrite the own owner's table extents from `before` to `after`.
///
/// An unset (sentinel) extent stays; an extent that does not describe the
/// table before the edit is refused.
pub(super) fn rewrite_extents(
    owner: &[u8],
    before: Extents,
    after: Extents,
) -> FaultResult<Option<Vec<u8>>> {
    let mut changed = false;
    let output = rebuild(owner, |field| {
        if field.number != 7 && field.number != 8 {
            return Ok(None);
        }
        let spanning = rebuild(field.payload, |inner| {
            let (old, new) = match inner.number {
                2 => (before.total, after.total),
                3 => (before.body, after.body),
                _ => return Ok(None),
            };
            let values = range(inner.payload)?;
            if is_sentinel(values) {
                return Ok(None);
            }
            if values != old {
                return Err(Fault::Unsupported);
            }
            if old == new {
                return Ok(None);
            }
            changed = true;
            bytes_field(inner.number, &encode_range(new)?).map(Some)
        })?;
        if spanning == field.payload {
            return Ok(None);
        }
        bytes_field(field.number, &spanning).map(Some)
    })?;
    Ok(changed.then_some(output))
}
