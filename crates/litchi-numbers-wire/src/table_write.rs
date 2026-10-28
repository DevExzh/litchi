//! Payload-level cell writes shared by the table owners of Pages and
//! Keynote (record 0309).
//!
//! A TST table stores each row's cells in a `TST.TileRowInfo` (BNC storage
//! and a cell-offset table), its text in a string `TST.TableDataList`, and
//! per-row and per-column cell counts in header buckets. These functions
//! rewrite one payload each and keep every field they do not own byte for
//! byte; the caller locates the objects, replaces the payloads and
//! publishes. The calculation-engine check refuses tables whose formula
//! graph a value change would have to recalculate.

use litchi_iwa_common::wire::{RawWireFields, RawWireLimits, WireView};

use crate::{BncCell, LimitKind, Limits, ListKind};

/// A malformed payload, or a native shape these writes do not own.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WriteError {
    /// The payload is malformed.
    Invalid,
    /// The payload uses a shape these writes do not rewrite (a rich-text,
    /// formula or commented cell, a legacy tile, a shared string entry
    /// with an object).
    Unsupported,
    /// A caller-selected raw-wire ceiling was exceeded before output was
    /// published.
    LimitExceeded {
        /// Dimension that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Selected maximum.
        maximum: usize,
    },
    /// A fallible output or staging allocation failed.
    Allocation {
        /// Exact requested allocation size.
        requested: usize,
    },
}

impl std::fmt::Display for WriteError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid => formatter.write_str("malformed table payload"),
            Self::Unsupported => {
                formatter.write_str("table payload shape outside the shared cell writer")
            },
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "Numbers table {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { requested } => {
                write!(
                    formatter,
                    "Numbers table writer allocation failed for {requested} bytes"
                )
            },
        }
    }
}

impl std::error::Error for WriteError {}

type Result<T> = std::result::Result<T, WriteError>;

const MISSING: [u8; 2] = [0xff, 0xff];
/// `owner_kind` of a table's merge owner.
const MERGE_OWNER_KIND: u64 = 5;

fn map_limits(error: crate::limits::Error) -> WriteError {
    match error {
        crate::limits::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => WriteError::LimitExceeded {
            kind,
            observed: value,
            maximum,
        },
        crate::limits::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => WriteError::LimitExceeded {
            kind,
            observed,
            maximum,
        },
    }
}

fn reserve(output: &mut Vec<u8>, additional: usize) -> Result<()> {
    output
        .try_reserve(additional)
        .map_err(|_| WriteError::Allocation {
            requested: additional,
        })
}

fn append(output: &mut Vec<u8>, bytes: &[u8], limits: &Limits) -> Result<()> {
    let length = output
        .len()
        .checked_add(bytes.len())
        .ok_or(WriteError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: usize::MAX,
            maximum: limits.max_output_bytes(),
        })?;
    limits.check_output(length).map_err(map_limits)?;
    reserve(output, bytes.len())?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn push_varint(output: &mut Vec<u8>, mut value: u64, limits: &Limits) -> Result<()> {
    let mut encoded = [0u8; 10];
    let mut length = 0usize;
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        encoded[length] = if value == 0 { byte } else { byte | 0x80 };
        length += 1;
        if value == 0 {
            return append(output, &encoded[..length], limits);
        }
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64, limits: &Limits) -> Result<()> {
    push_varint(output, u64::from(number) << 3, limits)?;
    push_varint(output, value, limits)
}

fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8], limits: &Limits) -> Result<()> {
    push_varint(output, (u64::from(number) << 3) | 2, limits)?;
    push_varint(output, payload.len() as u64, limits)?;
    append(output, payload, limits)
}

fn fixed32_field(output: &mut Vec<u8>, number: u32, value: u32, limits: &Limits) -> Result<()> {
    push_varint(output, (u64::from(number) << 3) | 5, limits)?;
    append(output, &value.to_le_bytes(), limits)
}

pub(crate) struct Field<'a> {
    pub(crate) number: u32,
    pub(crate) wire_type: u8,
    pub(crate) raw: &'a [u8],
    pub(crate) payload: &'a [u8],
}

pub(crate) fn fields<'a>(data: &'a [u8], limits: &Limits) -> Result<Vec<Field<'a>>> {
    limits.check_input(data.len()).map_err(map_limits)?;
    let view = WireView::parse(data, limits.wire()).map_err(map_wire_error)?;
    let count = view.fields().count();
    limits.check_fields(count).map_err(map_limits)?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(count)
        .map_err(|_| WriteError::Allocation { requested: count })?;
    for field in view.fields() {
        fields.push(Field {
            number: field.number(),
            wire_type: field.wire_type(),
            raw: field.raw(),
            payload: field.payload(),
        });
    }
    Ok(fields)
}

fn map_wire_error(error: litchi_iwa_common::Error) -> WriteError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => {
            let kind = match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::Fields,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::RewriteWork
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells
                | litchi_iwa_common::LimitKind::Nesting => LimitKind::Items,
            };
            WriteError::LimitExceeded {
                kind,
                observed,
                maximum: limit,
            }
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            WriteError::Allocation { requested: amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => WriteError::Invalid,
        _ => WriteError::Invalid,
    }
}

fn varint(payload: &[u8]) -> Result<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| WriteError::Invalid)?;
    if length != payload.len() {
        return Err(WriteError::Invalid);
    }
    Ok(value)
}

pub(crate) fn single_varint(data: &[u8], number: u32, limits: &Limits) -> Result<Option<u64>> {
    let mut found = None;
    for field in fields(data, limits)? {
        if field.number == number {
            if field.wire_type != 0 || found.replace(varint(field.payload)?).is_some() {
                return Err(WriteError::Invalid);
            }
        }
    }
    Ok(found)
}

pub(crate) fn single_bytes<'a>(
    data: &'a [u8],
    number: u32,
    limits: &Limits,
) -> Result<Option<&'a [u8]>> {
    let mut found = None;
    for field in fields(data, limits)? {
        if field.number == number {
            if field.wire_type != 2 || found.replace(field.payload).is_some() {
                return Err(WriteError::Invalid);
            }
        }
    }
    Ok(found)
}

/// Rebuild `data`, replacing the fields `map` returns `Some` for (an empty
/// replacement drops the field).
fn rebuild(
    data: &[u8],
    limits: &Limits,
    mut map: impl FnMut(&Field<'_>) -> Result<Option<Vec<u8>>>,
) -> Result<Vec<u8>> {
    limits.check_input(data.len()).map_err(map_limits)?;
    let fields = fields(data, limits)?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(fields.len())
        .map_err(|_| WriteError::Allocation {
            requested: fields.len(),
        })?;
    if replacements.capacity() != fields.len() {
        return Err(WriteError::Allocation {
            requested: fields.len(),
        });
    }
    for field in &fields {
        replacements.push(map(field)?);
    }

    // Preflight the actual encoded representation before allocating the final
    // output.  This counts replacements and retained raw fields separately;
    // the old source-length reservation could under-reserve a larger rewrite
    // and over-reserve a shrinking one.
    let output_len =
        fields
            .iter()
            .zip(&replacements)
            .try_fold(0usize, |length, (field, replacement)| {
                length
                    .checked_add(replacement.as_ref().map_or(field.raw.len(), Vec::len))
                    .ok_or(WriteError::LimitExceeded {
                        kind: LimitKind::OutputBytes,
                        observed: usize::MAX,
                        maximum: limits.max_output_bytes(),
                    })
            })?;
    limits.check_output(output_len).map_err(map_limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| WriteError::Allocation {
            requested: output_len,
        })?;
    if output.capacity() != output_len {
        return Err(WriteError::Allocation {
            requested: output_len,
        });
    }
    for (field, replacement) in fields.iter().zip(replacements) {
        output.extend_from_slice(replacement.as_deref().unwrap_or(field.raw));
    }
    if output.len() != output_len {
        return Err(WriteError::Invalid);
    }
    Ok(output)
}

/// The cell slots of one BNC storage: `(start, end)` per column.
fn slots(
    storage: &[u8],
    offsets: &[u8],
    unit: usize,
    limits: &Limits,
) -> Result<Vec<Option<(usize, usize)>>> {
    if unit == 0 {
        return Err(WriteError::Invalid);
    }
    if offsets.len() % 2 != 0 {
        return Err(WriteError::Invalid);
    }
    let count = offsets.len() / 2;
    limits.check_items(count).map_err(map_limits)?;
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(count)
        .map_err(|_| WriteError::Allocation { requested: count })?;
    for pair in offsets.chunks_exact(2) {
        let raw = [pair[0], pair[1]];
        let start = if raw == MISSING {
            None
        } else {
            Some(
                usize::from(u16::from_le_bytes(raw))
                    .checked_mul(unit)
                    .ok_or(WriteError::Invalid)?,
            )
        };
        starts.push(start);
    }
    let present_count = starts.iter().flatten().count();
    limits.check_items(present_count).map_err(map_limits)?;
    let mut present = Vec::new();
    present
        .try_reserve_exact(present_count)
        .map_err(|_| WriteError::Allocation {
            requested: present_count,
        })?;
    present.extend(starts.iter().flatten().copied());
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.last().is_some_and(|last| *last >= storage.len())
        || present.first().is_some_and(|first| *first != 0)
    {
        return Err(WriteError::Unsupported);
    }
    let mut next = present.iter().skip(1);
    let mut output = Vec::new();
    output
        .try_reserve_exact(starts.len())
        .map_err(|_| WriteError::Allocation {
            requested: starts.len(),
        })?;
    output.extend(
        starts
            .into_iter()
            .map(|start| start.map(|start| (start, next.next().copied().unwrap_or(storage.len())))),
    );
    Ok(output)
}

/// Encode a cell list as BNC storage and offsets in `unit`-byte units.
fn encode(cells: &[Option<&[u8]>], unit: usize, limits: &Limits) -> Result<(Vec<u8>, Vec<u8>)> {
    if unit == 0 {
        return Err(WriteError::Invalid);
    }
    limits.check_items(cells.len()).map_err(map_limits)?;
    for cell in cells.iter().flatten() {
        limits.check_input(cell.len()).map_err(map_limits)?;
    }
    let storage_len = cells
        .iter()
        .try_fold(0usize, |length, cell| {
            let Some(cell) = cell else {
                return Some(length);
            };
            let padding = (unit - (length % unit)) % unit;
            length.checked_add(padding)?.checked_add(cell.len())
        })
        .ok_or(WriteError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: usize::MAX,
            maximum: limits.max_output_bytes(),
        })?;
    let storage_len = storage_len
        .checked_add((unit - (storage_len % unit)) % unit)
        .ok_or(WriteError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: usize::MAX,
            maximum: limits.max_output_bytes(),
        })?;
    limits.check_output(storage_len).map_err(map_limits)?;
    let offset_len = cells.len().checked_mul(2).ok_or(WriteError::Invalid)?;
    let combined_len = storage_len
        .checked_add(offset_len)
        .ok_or(WriteError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: usize::MAX,
            maximum: limits.max_output_bytes(),
        })?;
    limits.check_output(combined_len).map_err(map_limits)?;
    let mut storage = Vec::new();
    storage
        .try_reserve_exact(storage_len)
        .map_err(|_| WriteError::Allocation {
            requested: storage_len,
        })?;
    let mut offsets = Vec::new();
    offsets
        .try_reserve_exact(offset_len)
        .map_err(|_| WriteError::Allocation {
            requested: offset_len,
        })?;
    for cell in cells {
        let Some(cell) = cell else {
            offsets.extend_from_slice(&MISSING);
            continue;
        };
        while storage.len() % unit != 0 {
            storage.push(0);
        }
        let offset =
            u16::try_from(storage.len() / unit).map_err(|_error| WriteError::Unsupported)?;
        if offset.to_le_bytes() == MISSING {
            return Err(WriteError::Unsupported);
        }
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(cell);
    }
    while storage.len() % unit != 0 {
        storage.push(0);
    }
    Ok((storage, offsets))
}

/// The result of one tile cell write.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TileCellWrite {
    /// The rewritten tile payload.
    pub tile: Vec<u8>,
    /// The cell's bytes before the write, if it was stored.
    pub previous: Option<Vec<u8>>,
}

/// The stored BNC cell at `local_row`/`column` of a tile payload.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed tile.
pub fn tile_cell(
    tile: &[u8],
    local_row: u32,
    column: u32,
    limits: &Limits,
) -> Result<Option<Vec<u8>>> {
    for field in fields(tile, limits)? {
        if field.number != 5
            || single_varint(field.payload, 1, limits)? != Some(u64::from(local_row))
        {
            continue;
        }
        let (Some(storage), Some(offsets)) = (
            single_bytes(field.payload, 6, limits)?,
            single_bytes(field.payload, 7, limits)?,
        ) else {
            return Err(WriteError::Unsupported);
        };
        let unit = if single_varint(field.payload, 8, limits)?.unwrap_or(0) != 0 {
            4
        } else {
            1
        };
        let cells = slots(storage, offsets, unit, limits)?;
        let column = usize::try_from(column).map_err(|_| WriteError::Invalid)?;
        let Some((start, end)) = cells.get(column).copied().flatten() else {
            return Ok(None);
        };
        limits.check_output(end - start).map_err(map_limits)?;
        let mut copy = Vec::new();
        copy.try_reserve_exact(end - start)
            .map_err(|_| WriteError::Allocation {
                requested: end - start,
            })?;
        copy.extend_from_slice(&storage[start..end]);
        return Ok(Some(copy));
    }
    Ok(None)
}

/// `previous` without its formula: the formula reference, the cached result
/// and every other value field go, the cell's metadata stays (the
/// formula-removal half of a write into a formula cell, record 0376a).
///
/// # Errors
///
/// [`WriteError::Unsupported`] for a formula with a formula-error entry;
/// [`WriteError::Invalid`] for malformed BNC.
pub fn formula_free(previous: &[u8], limits: &Limits) -> Result<Vec<u8>> {
    let mut cell = BncCell::parse(previous, limits).map_err(map_bnc_error)?;
    if cell.formula_error_identifier().is_some() {
        return Err(WriteError::Unsupported);
    }
    cell.clear_value_preserving_metadata();
    cell.encode(limits).map_err(map_bnc_error)
}

/// `cell`, holding a formula's cached result, with the formula-list key
/// `key` attached (the host's order: cached value first, then the formula).
///
/// # Errors
///
/// [`WriteError::Invalid`] for malformed BNC.
pub fn with_formula(cell: &[u8], key: u32, limits: &Limits) -> Result<Vec<u8>> {
    let mut cell = BncCell::parse(cell, limits).map_err(map_bnc_error)?;
    cell.set_formula_reference(key);
    cell.encode(limits).map_err(map_bnc_error)
}

/// Every stored BNC cell of a modern tile, as `(local row, column, bytes)`
/// in row then column order (the formula-cache refresh's tile scan).
pub(crate) fn stored_cells<'a>(
    tile: &'a [u8],
    limits: &Limits,
) -> Result<Vec<(u32, u32, &'a [u8])>> {
    let mut cells = Vec::new();
    for field in fields(tile, limits)? {
        if field.number != 5 {
            continue;
        }
        let row = single_varint(field.payload, 1, limits)?.ok_or(WriteError::Invalid)?;
        let row = u32::try_from(row).map_err(|_| WriteError::Invalid)?;
        let (Some(storage), Some(offsets)) = (
            single_bytes(field.payload, 6, limits)?,
            single_bytes(field.payload, 7, limits)?,
        ) else {
            return Err(WriteError::Unsupported);
        };
        let unit = if single_varint(field.payload, 8, limits)?.unwrap_or(0) != 0 {
            4
        } else {
            1
        };
        for (column, slot) in slots(storage, offsets, unit, limits)?
            .into_iter()
            .enumerate()
        {
            let Some((start, end)) = slot else {
                continue;
            };
            let column = u32::try_from(column).map_err(|_| WriteError::Invalid)?;
            let next = cells.len().checked_add(1).ok_or(WriteError::Invalid)?;
            limits.check_items(next).map_err(map_limits)?;
            cells
                .try_reserve(1)
                .map_err(|_| WriteError::Allocation { requested: 1 })?;
            cells.push((row, column, &storage[start..end]));
        }
    }
    Ok(cells)
}

/// Set (`Some`) or remove (`None`) the BNC cell at `local_row`/`column` of
/// a modern tile with `columns` table columns; a missing row is inserted in
/// row order.
///
/// # Errors
///
/// [`WriteError::Unsupported`] for a legacy (pre-BNC only) tile or row.
pub fn tile_set_cell(
    tile: &[u8],
    local_row: u32,
    column: u32,
    columns: u32,
    cell: Option<&[u8]>,
    limits: &Limits,
) -> Result<TileCellWrite> {
    tile_set_cell_placed(tile, local_row, column, columns, cell, false, limits)
}

/// [`tile_set_cell`], writing a new row after every stored row in field
/// order: after the last row record, before the tile's storage-version
/// fields, as the Numbers host writer's encoder places it (record 0370b).
/// [`tile_set_cell`] appends such a row after them, as the Pages and
/// Keynote host writers do.
///
/// # Errors
///
/// As [`tile_set_cell`].
pub fn tile_set_cell_in_field_order(
    tile: &[u8],
    local_row: u32,
    column: u32,
    columns: u32,
    cell: Option<&[u8]>,
    limits: &Limits,
) -> Result<TileCellWrite> {
    tile_set_cell_placed(tile, local_row, column, columns, cell, true, limits)
}

#[allow(
    clippy::too_many_lines,
    reason = "one row rewrite: decode, splice, new row and row count"
)]
fn tile_set_cell_placed(
    tile: &[u8],
    local_row: u32,
    column: u32,
    columns: u32,
    cell: Option<&[u8]>,
    field_order: bool,
    limits: &Limits,
) -> Result<TileCellWrite> {
    if single_varint(tile, 6, limits)? != Some(5) || single_varint(tile, 7, limits)? != Some(1) {
        return Err(WriteError::Unsupported);
    }
    let column = usize::try_from(column).map_err(|_| WriteError::Invalid)?;
    let columns_usize = usize::try_from(columns).map_err(|_| WriteError::Invalid)?;
    limits.check_items(columns_usize).map_err(map_limits)?;
    if let Some(cell) = cell {
        limits.check_input(cell.len()).map_err(map_limits)?;
    }
    if column >= columns_usize {
        return Err(WriteError::Invalid);
    }
    let tile_fields = fields(tile, limits)?;
    let row_count = tile_fields.iter().filter(|field| field.number == 5).count();
    limits.check_items(row_count).map_err(map_limits)?;
    let mut rows = Vec::new();
    rows.try_reserve_exact(row_count)
        .map_err(|_| WriteError::Allocation {
            requested: row_count,
        })?;
    // A new row's offset table takes the widest existing row's capacity
    // (native rows keep spare slots), as the Numbers host writer sizes it
    // (record 0370).
    let mut capacity = columns_usize;
    for field in tile_fields.iter().filter(|field| field.number == 5) {
        rows.push(single_varint(field.payload, 1, limits)?.ok_or(WriteError::Invalid)?);
        if let Some(offsets) = single_bytes(field.payload, 7, limits)? {
            capacity = capacity.max(offsets.len() / 2);
        }
    }
    limits.check_items(capacity).map_err(map_limits)?;
    limits.check_items(rows.len()).map_err(map_limits)?;
    let numrows = single_varint(tile, 4, limits)?.ok_or(WriteError::Invalid)?;
    let counts_rows = numrows == rows.len() as u64;
    let mut previous = None;
    let mut found = false;
    let mut rewritten_row = |row: &[u8]| -> Result<Option<Vec<u8>>> {
        let (Some(storage), Some(offsets)) =
            (single_bytes(row, 6, limits)?, single_bytes(row, 7, limits)?)
        else {
            return Err(WriteError::Unsupported);
        };
        let unit = if single_varint(row, 8, limits)?.unwrap_or(0) != 0 {
            4
        } else {
            1
        };
        let slots = slots(storage, offsets, unit, limits)?;
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(slots.len())
            .map_err(|_| WriteError::Allocation {
                requested: slots.len(),
            })?;
        for slot in slots {
            cells.push(slot.map(|(start, end)| &storage[start..end]));
        }
        if cells.len() < columns_usize {
            limits.check_items(columns_usize).map_err(map_limits)?;
            let additional = columns_usize - cells.len();
            cells
                .try_reserve_exact(additional)
                .map_err(|_| WriteError::Allocation {
                    requested: additional,
                })?;
            cells.resize(columns_usize, None);
        }
        previous = match cells[column] {
            Some(bytes) => {
                limits.check_output(bytes.len()).map_err(map_limits)?;
                let mut copy = Vec::new();
                copy.try_reserve_exact(bytes.len())
                    .map_err(|_| WriteError::Allocation {
                        requested: bytes.len(),
                    })?;
                copy.extend_from_slice(bytes);
                Some(copy)
            },
            None => None,
        };
        cells[column] = cell;
        let count = cells.iter().filter(|cell| cell.is_some()).count() as u64;
        if count == 0 {
            // A row left without cells goes, as the host writer drops it
            // (record 0370; native tiles store no empty rows).
            return Ok(None);
        }
        let (storage, offsets) = encode(&cells, unit, limits)?;
        Ok(Some(rebuild(row, limits, |field| match field.number {
            2 => {
                let mut output = Vec::new();
                varint_field(&mut output, 2, count, limits)?;
                Ok(Some(output))
            },
            6 => {
                let mut output = Vec::new();
                bytes_field(&mut output, 6, &storage, limits)?;
                Ok(Some(output))
            },
            7 => {
                let mut output = Vec::new();
                bytes_field(&mut output, 7, &offsets, limits)?;
                Ok(Some(output))
            },
            _ => Ok(None),
        })?))
    };
    // A new row record for a missing row, in the canonical BNC shape.
    let new_row = |cell: &[u8]| -> Result<Vec<u8>> {
        limits.check_items(columns_usize).map_err(map_limits)?;
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(capacity)
            .map_err(|_| WriteError::Allocation {
                requested: capacity,
            })?;
        cells.resize(capacity, None);
        cells[column] = Some(cell);
        let (storage, offsets) = encode(&cells, 1, limits)?;
        let mut row = Vec::new();
        varint_field(&mut row, 1, u64::from(local_row), limits)?;
        varint_field(&mut row, 2, 1, limits)?;
        bytes_field(&mut row, 3, &[], limits)?;
        bytes_field(&mut row, 4, &[], limits)?;
        varint_field(&mut row, 5, 5, limits)?;
        bytes_field(&mut row, 6, &storage, limits)?;
        bytes_field(&mut row, 7, &offsets, limits)?;
        varint_field(&mut row, 8, 0, limits)?;
        let mut field = Vec::new();
        bytes_field(&mut field, 5, &row, limits)?;
        Ok(field)
    };
    let insert_before = rows.iter().position(|row| *row > u64::from(local_row));
    let exists = rows.contains(&u64::from(local_row));
    let mut row_index = 0_usize;
    let mut inserted = false;
    let mut removed = false;
    let mut output = Vec::new();
    reserve(&mut output, tile.len().min(limits.max_output_bytes()))?;
    // In field order, a row after every stored row follows the last row
    // record (or, in a tile without rows, precedes the storage-version
    // fields); otherwise it goes last.
    let all_fields = fields(tile, limits)?;
    let trailing = if field_order {
        all_fields
            .iter()
            .rposition(|field| field.number == 5)
            .map(|last| last + 1)
            .or_else(|| all_fields.iter().position(|field| field.number > 5))
            .unwrap_or(all_fields.len())
    } else {
        all_fields.len()
    };
    for (position, field) in all_fields.into_iter().enumerate() {
        if position == trailing
            && !exists
            && insert_before.is_none()
            && let Some(cell) = cell
        {
            let row = new_row(cell)?;
            append(&mut output, &row, limits)?;
            inserted = true;
        }
        if field.number == 5 {
            let this = row_index;
            row_index += 1;
            if !exists && Some(this) == insert_before {
                if let Some(cell) = cell {
                    let row = new_row(cell)?;
                    append(&mut output, &row, limits)?;
                    inserted = true;
                }
            }
            if single_varint(field.payload, 1, limits)? == Some(u64::from(local_row)) {
                found = true;
                match rewritten_row(field.payload)? {
                    Some(row) => bytes_field(&mut output, 5, &row, limits)?,
                    None => removed = true,
                }
                continue;
            }
        }
        append(&mut output, field.raw, limits)?;
    }
    if !exists && !inserted {
        if let Some(cell) = cell {
            let row = new_row(cell)?;
            append(&mut output, &row, limits)?;
            inserted = true;
        }
    }
    let _ = found;
    if inserted || removed {
        let new_numrows = if counts_rows {
            if inserted {
                numrows.checked_add(1).ok_or(WriteError::Invalid)?
            } else {
                numrows.checked_sub(1).ok_or(WriteError::Invalid)?
            }
        } else if inserted {
            numrows.max(
                u64::from(local_row)
                    .checked_add(1)
                    .ok_or(WriteError::Invalid)?,
            )
        } else {
            rows.iter()
                .filter(|row| **row != u64::from(local_row))
                .max()
                .map_or(0, |row| row + 1)
        };
        output = rebuild(&output, limits, |field| {
            if field.number != 4 {
                return Ok(None);
            }
            let mut replacement = Vec::new();
            varint_field(&mut replacement, 4, new_numrows, limits)?;
            Ok(Some(replacement))
        })?;
    }
    Ok(TileCellWrite {
        tile: output,
        previous,
    })
}

/// Whether a `TST.HeaderStorageBucket.Header` holds only its index, a zero
/// size, a visible hiding state and its cell count.
fn header_is_bare(header: &[u8], limits: &Limits) -> Result<bool> {
    for field in fields(header, limits)? {
        match field.number {
            1 | 4 => {},
            2 => {
                if field.payload.iter().any(|byte| *byte != 0) {
                    return Ok(false);
                }
            },
            3 => {
                if single_varint(header, 3, limits)? != Some(0) {
                    return Ok(false);
                }
            },
            _ => return Ok(false),
        }
    }
    Ok(true)
}

/// The most bytes [`header_adjust_many`] adds to a bucket per changed
/// header: a new header's key and length, and its index, size, hiding and
/// count fields with their keys.
pub const HEADER_BYTES: usize = 26;

/// The facts of one `TST.HeaderStorageBucket.Header` a count change needs.
struct HeaderFacts {
    index: u64,
    cells: Option<u64>,
    /// Encoded length of the count field.
    count_bytes: usize,
    /// Only an index, a zero size, a visible hiding state and a count.
    bare: bool,
}

fn header_facts(header: &[u8], limits: RawWireLimits) -> Result<HeaderFacts> {
    let mut index = None;
    let mut cells = None;
    let mut count_bytes = 0;
    let mut bare = true;
    let mut fields = RawWireFields::new(header, limits);
    while let Some(field) = fields.next().map_err(map_wire_error)? {
        match (field.number(), field.wire_type()) {
            (1, 0) => {
                if index.replace(varint(field.payload())?).is_some() {
                    return Err(WriteError::Invalid);
                }
            },
            (4, 0) => {
                if cells.replace(varint(field.payload())?).is_some() {
                    return Err(WriteError::Invalid);
                }
                count_bytes = field.raw().len();
            },
            (1 | 4, _) => return Err(WriteError::Invalid),
            (2, _) => bare &= field.payload().iter().all(|byte| *byte == 0),
            (3, 0) => bare &= varint(field.payload())? == 0,
            _ => bare = false,
        }
    }
    Ok(HeaderFacts {
        index: index.ok_or(WriteError::Unsupported)?,
        cells,
        count_bytes,
        bare,
    })
}

const fn varint_len(value: u64) -> usize {
    (64 - (value | 1).leading_zeros() as usize).div_ceil(7)
}

/// Append a header field for column or row `index` counting `cells` cells
/// (size 0, visible).
fn push_new_header(output: &mut Vec<u8>, index: u32, cells: u64, limits: &Limits) -> Result<()> {
    let length = varint_len(u64::from(index)) + varint_len(cells) + 9;
    push_varint(output, (2 << 3) | 2, limits)?;
    push_varint(output, length as u64, limits)?;
    varint_field(output, 1, u64::from(index), limits)?;
    fixed32_field(output, 2, 0, limits)?;
    varint_field(output, 3, 0, limits)?;
    varint_field(output, 4, cells, limits)
}

/// Append the header field `header` with its count replaced by `cells`.
fn push_counted_header(
    output: &mut Vec<u8>,
    header: &[u8],
    facts: &HeaderFacts,
    cells: u64,
    limits: &Limits,
    raw_limits: RawWireLimits,
) -> Result<()> {
    let length = header.len() - facts.count_bytes + 1 + varint_len(cells);
    push_varint(output, (2 << 3) | 2, limits)?;
    push_varint(output, length as u64, limits)?;
    let mut fields = RawWireFields::new(header, raw_limits);
    while let Some(field) = fields.next().map_err(map_wire_error)? {
        if field.number() == 4 {
            varint_field(output, 4, cells, limits)?;
        } else {
            append(output, field.raw(), limits)?;
        }
    }
    Ok(())
}

/// [`header_adjust`] of several headers in one pass over a bucket whose
/// headers are in increasing index order, growing it by at most
/// [`HEADER_BYTES`] per changed header. `deltas` holds `(index, delta)`
/// pairs in increasing index order. A removal the bucket does not count is
/// ignored, as the host ignores it, so no count drops below zero.
///
/// # Errors
///
/// [`WriteError::Invalid`] for unsorted `deltas`; [`WriteError::Unsupported`]
/// for a bucket whose headers are out of index order.
pub fn header_adjust_many(
    bucket: &[u8],
    deltas: &[(u32, i64)],
    limits: &Limits,
) -> Result<Vec<u8>> {
    if deltas.windows(2).any(|pair| pair[0].0 >= pair[1].0) {
        return Err(WriteError::Invalid);
    }
    limits.check_input(bucket.len()).map_err(map_limits)?;
    let raw_limits = RawWireLimits::new(
        bucket.len().max(1),
        limits.max_fields(),
        0,
        limits
            .wire()
            .max_rewrite_work()
            .min(RawWireLimits::MAX_WORK),
    )
    .map_err(map_wire_error)?;
    let bound = deltas
        .len()
        .checked_mul(HEADER_BYTES)
        .and_then(|growth| growth.checked_add(bucket.len()))
        .ok_or(WriteError::Invalid)?;
    let mut output = Vec::new();
    reserve(&mut output, bound.min(limits.max_output_bytes()))?;
    let mut pending = deltas
        .iter()
        .copied()
        .filter(|(_, delta)| *delta != 0)
        .peekable();
    let mut previous = None;
    let mut fields = RawWireFields::new(bucket, raw_limits);
    while let Some(field) = fields.next().map_err(map_wire_error)? {
        if field.number() == 2 {
            if field.wire_type() != 2 {
                return Err(WriteError::Invalid);
            }
            let facts = header_facts(field.payload(), raw_limits)?;
            if previous.is_some_and(|previous| previous >= facts.index) {
                return Err(WriteError::Unsupported);
            }
            previous = Some(facts.index);
            while let Some((index, delta)) =
                pending.next_if(|(index, _)| u64::from(*index) < facts.index)
            {
                if let Ok(cells) = u64::try_from(delta) {
                    push_new_header(&mut output, index, cells, limits)?;
                }
            }
            if let Some((_, delta)) = pending.next_if(|(index, _)| u64::from(*index) == facts.index)
            {
                let after = match facts.cells {
                    Some(cells) if delta < 0 => cells.saturating_sub(delta.unsigned_abs()),
                    Some(cells) => cells
                        .checked_add(delta.unsigned_abs())
                        .ok_or(WriteError::Invalid)?,
                    None if delta < 0 => {
                        append(&mut output, field.raw(), limits)?;
                        continue;
                    },
                    None => return Err(WriteError::Invalid),
                };
                if facts.cells == Some(after) {
                    append(&mut output, field.raw(), limits)?;
                } else if after != 0 || !facts.bare {
                    push_counted_header(
                        &mut output,
                        field.payload(),
                        &facts,
                        after,
                        limits,
                        raw_limits,
                    )?;
                }
                continue;
            }
        }
        append(&mut output, field.raw(), limits)?;
    }
    for (index, delta) in pending {
        if let Ok(cells) = u64::try_from(delta) {
            push_new_header(&mut output, index, cells, limits)?;
        }
    }
    Ok(output)
}

/// Add `delta` to the cell count of header `index` in a header bucket,
/// creating the header (size 0, visible) when it has none; a bare header
/// left counting no cell goes.
///
/// # Errors
///
/// [`WriteError::Unsupported`] when a count would drop below zero.
pub fn header_adjust(bucket: &[u8], index: u32, delta: i64, limits: &Limits) -> Result<Vec<u8>> {
    if delta == 0 {
        limits.check_input(bucket.len()).map_err(map_limits)?;
        limits.check_output(bucket.len()).map_err(map_limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(bucket.len())
            .map_err(|_| WriteError::Allocation {
                requested: bucket.len(),
            })?;
        output.extend_from_slice(bucket);
        return Ok(output);
    }
    let mut done = false;
    let mut output = rebuild(bucket, limits, |field| {
        if field.number != 2 || single_varint(field.payload, 1, limits)? != Some(u64::from(index)) {
            return Ok(None);
        }
        done = true;
        let cells = single_varint(field.payload, 4, limits)?.ok_or(WriteError::Invalid)?;
        let after = i64::try_from(cells)
            .ok()
            .and_then(|cells| cells.checked_add(delta))
            .filter(|after| *after >= 0)
            .ok_or(WriteError::Unsupported)?;
        if after == 0 && header_is_bare(field.payload, limits)? {
            // A header left counting no cell and carrying nothing else goes,
            // as the host writer drops it (record 0370).
            return Ok(Some(Vec::new()));
        }
        let header = rebuild(field.payload, limits, |inner| {
            if inner.number != 4 {
                return Ok(None);
            }
            let mut replacement = Vec::new();
            varint_field(&mut replacement, 4, after as u64, limits)?;
            Ok(Some(replacement))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 2, &header, limits)?;
        Ok(Some(replacement))
    })?;
    if !done {
        if delta < 0 {
            return Err(WriteError::Unsupported);
        }
        // Headers stay sorted by index.
        let mut header = Vec::new();
        varint_field(&mut header, 1, u64::from(index), limits)?;
        fixed32_field(&mut header, 2, 0, limits)?;
        varint_field(&mut header, 3, 0, limits)?;
        varint_field(&mut header, 4, delta as u64, limits)?;
        let mut entry = Vec::new();
        bytes_field(&mut entry, 2, &header, limits)?;
        let mut placed = false;
        let mut sorted = Vec::new();
        reserve(
            &mut sorted,
            output
                .len()
                .saturating_add(entry.len())
                .min(limits.max_output_bytes()),
        )?;
        for field in fields(&output, limits)? {
            if !placed
                && field.number == 2
                && single_varint(field.payload, 1, limits)?
                    .is_some_and(|other| other > u64::from(index))
            {
                append(&mut sorted, &entry, limits)?;
                placed = true;
            }
            append(&mut sorted, field.raw, limits)?;
        }
        if !placed {
            append(&mut sorted, &entry, limits)?;
        }
        output = sorted;
    }
    Ok(output)
}

/// Intern `text` in a string `TST.TableDataList` payload: an entry with the
/// same string gains a reference, else a new entry takes the next key.
///
/// # Errors
///
/// [`WriteError::Unsupported`] for a segmented list.
pub fn list_intern(list: &[u8], text: &str, limits: &Limits) -> Result<(Vec<u8>, u32)> {
    limits.check_input(text.len()).map_err(map_limits)?;
    if fields(list, limits)?.iter().any(|field| field.number == 4) {
        return Err(WriteError::Unsupported);
    }
    let mut key = None;
    let output = rebuild(list, limits, |field| {
        if field.number != 3
            || key.is_some()
            || single_bytes(field.payload, 3, limits)? != Some(text.as_bytes())
        {
            return Ok(None);
        }
        let entry_key = single_varint(field.payload, 1, limits)?.ok_or(WriteError::Invalid)?;
        key = Some(u32::try_from(entry_key).map_err(|_error| WriteError::Invalid)?);
        let count = single_varint(field.payload, 2, limits)?.ok_or(WriteError::Invalid)?;
        let entry = rebuild(field.payload, limits, |inner| {
            if inner.number != 2 {
                return Ok(None);
            }
            let mut replacement = Vec::new();
            varint_field(
                &mut replacement,
                2,
                count.checked_add(1).ok_or(WriteError::Invalid)?,
                limits,
            )?;
            Ok(Some(replacement))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 3, &entry, limits)?;
        Ok(Some(replacement))
    })?;
    if let Some(key) = key {
        return Ok((output, key));
    }
    let next = single_varint(list, 2, limits)?
        .ok_or(WriteError::Invalid)?
        .max(1);
    let key = u32::try_from(next).map_err(|_error| WriteError::Unsupported)?;
    let mut output = rebuild(list, limits, |field| {
        if field.number != 2 {
            return Ok(None);
        }
        let mut replacement = Vec::new();
        varint_field(
            &mut replacement,
            2,
            next.checked_add(1).ok_or(WriteError::Invalid)?,
            limits,
        )?;
        Ok(Some(replacement))
    })?;
    let mut entry = Vec::new();
    varint_field(&mut entry, 1, next, limits)?;
    varint_field(&mut entry, 2, 1, limits)?;
    bytes_field(&mut entry, 3, text.as_bytes(), limits)?;
    bytes_field(&mut output, 3, &entry, limits)?;
    Ok((output, key))
}

/// Drop one reference to string entry `key`; an entry left without
/// references goes, and when it held the list's last key, `next_list_id`
/// takes that key back.
///
/// # Errors
///
/// [`WriteError::Unsupported`] for an entry holding an object reference or
/// a key the list does not hold.
pub fn list_release(list: &[u8], key: u32, limits: &Limits) -> Result<Vec<u8>> {
    let mut done = false;
    let mut output = rebuild(list, limits, |field| {
        if field.number != 3 || single_varint(field.payload, 1, limits)? != Some(u64::from(key)) {
            return Ok(None);
        }
        if single_bytes(field.payload, 4, limits)?.is_some() {
            return Err(WriteError::Unsupported);
        }
        done = true;
        let count = single_varint(field.payload, 2, limits)?.ok_or(WriteError::Invalid)?;
        if count <= 1 {
            return Ok(Some(Vec::new()));
        }
        let entry = rebuild(field.payload, limits, |inner| {
            if inner.number != 2 {
                return Ok(None);
            }
            let mut replacement = Vec::new();
            varint_field(&mut replacement, 2, count - 1, limits)?;
            Ok(Some(replacement))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 3, &entry, limits)?;
        Ok(Some(replacement))
    })?;
    if !done {
        return Err(WriteError::Unsupported);
    }
    // A removed last key is handed back, as the Numbers cell writers do
    // (record 0370).
    if single_varint(&output, 2, limits)? == Some(u64::from(key) + 1)
        && !fields(&output, limits)?.iter().any(|field| {
            field.number == 3
                && single_varint(field.payload, 1, limits).ok().flatten() == Some(u64::from(key))
        })
    {
        output = rebuild(&output, limits, |field| {
            if field.number != 2 {
                return Ok(None);
            }
            let mut replacement = Vec::new();
            varint_field(&mut replacement, 2, u64::from(key), limits)?;
            Ok(Some(replacement))
        })?;
    }
    Ok(output)
}

/// A value a shared cell write stores.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum CellInput<'a> {
    /// Plain text (interned in the string list).
    Text(&'a str),
    /// A finite number.
    Number(f64),
    /// A boolean.
    Boolean(bool),
    /// No value: the cell's style and format stay, its value goes.
    Empty,
}

/// The new BNC bytes of a cell taking `input`, keeping the previous cell's
/// style, text style and format keys. Returns the bytes (`None` when the
/// cell is left without anything) and the string key the previous value
/// released.
///
/// # Errors
///
/// [`WriteError::Unsupported`] for a previous rich-text, formula, error or
/// commented cell, or a non-finite number.
pub fn cell_bytes(
    previous: Option<&[u8]>,
    input: &CellInput<'_>,
    string_key: Option<u32>,
    limits: &Limits,
) -> Result<(Option<Vec<u8>>, Option<u32>)> {
    let mut cell = match previous {
        Some(bytes) => BncCell::parse(bytes, limits).map_err(map_bnc_error)?,
        None => BncCell::minimal(),
    };
    let mut released = None;
    for (kind, key) in cell.list_keys() {
        match kind {
            ListKind::String => released = Some(key),
            ListKind::RichText | ListKind::Formula | ListKind::FormulaError | ListKind::Comment => {
                return Err(WriteError::Unsupported);
            },
            _ => {},
        }
    }
    match input {
        CellInput::Text(_) => {
            cell.set_string(string_key.ok_or(WriteError::Invalid)?);
        },
        CellInput::Number(value) => {
            cell.set_number(*value)
                .map_err(|_error| WriteError::Unsupported)?;
        },
        CellInput::Boolean(value) => cell.set_boolean(*value),
        CellInput::Empty => {
            cell.clear_value_preserving_metadata();
            if previous.is_some()
                && cell.list_keys().next().is_none()
                && cell_is_bare(&cell, limits)?
            {
                return Ok((None, released));
            }
            if previous.is_none() {
                return Ok((None, None));
            }
        },
    }
    Ok((Some(cell.encode(limits).map_err(map_bnc_error)?), released))
}

/// Whether a cleared cell carries nothing worth storing.
fn cell_is_bare(cell: &BncCell, limits: &Limits) -> Result<bool> {
    Ok(cell.encode(limits).map_err(map_bnc_error)?
        == BncCell::minimal().encode(limits).map_err(map_bnc_error)?)
}

fn map_bnc_error(error: crate::Error) -> WriteError {
    match error {
        crate::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => WriteError::LimitExceeded {
            kind,
            observed,
            maximum,
        },
        crate::Error::Allocation { requested } => WriteError::Allocation { requested },
        crate::Error::InvalidFormat(_) | crate::Error::ParseError(_) => WriteError::Invalid,
    }
}

/// Why a table's calculation graph refuses a value write.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum FormulaState {
    /// The table has no formula owner, or an inert one.
    Inert,
    /// The table holds formulas or dependency records.
    Formulas,
    /// Another owner (a formula elsewhere, a chart) reads the table.
    Referenced,
}

fn cell_records<'a>(
    owner: &'a [u8],
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    limits: &Limits,
) -> Result<Vec<&'a [u8]>> {
    let mut records = Vec::new();
    if let Some(inline) = single_bytes(owner, 4, limits)? {
        for field in fields(inline, limits)? {
            if field.number == 1 {
                let next = records
                    .len()
                    .checked_add(1)
                    .ok_or(WriteError::LimitExceeded {
                        kind: LimitKind::Items,
                        observed: usize::MAX,
                        maximum: limits.max_items(),
                    })?;
                limits.check_items(next).map_err(map_limits)?;
                records
                    .try_reserve(1)
                    .map_err(|_| WriteError::Allocation { requested: 1 })?;
                records.push(field.payload);
            }
        }
    }
    if let Some(tiled) = single_bytes(owner, 13, limits)? {
        for field in fields(tiled, limits)? {
            if field.number != 1 {
                continue;
            }
            let tile = single_varint(field.payload, 1, limits)?
                .and_then(object)
                .ok_or(WriteError::Invalid)?;
            for record in fields(tile, limits)? {
                if record.number == 4 {
                    let next = records
                        .len()
                        .checked_add(1)
                        .ok_or(WriteError::LimitExceeded {
                            kind: LimitKind::Items,
                            observed: usize::MAX,
                            maximum: limits.max_items(),
                        })?;
                    limits.check_items(next).map_err(map_limits)?;
                    records
                        .try_reserve(1)
                        .map_err(|_| WriteError::Allocation { requested: 1 })?;
                    records.push(record.payload);
                }
            }
        }
    }
    Ok(records)
}

fn packed_or_single(field: &Field<'_>, limits: &Limits) -> Result<Vec<u64>> {
    if field.wire_type == 2 {
        let mut values = Vec::new();
        let mut rest = field.payload;
        while !rest.is_empty() {
            let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(rest)
                .map_err(|_error| WriteError::Invalid)?;
            let next = values
                .len()
                .checked_add(1)
                .ok_or(WriteError::LimitExceeded {
                    kind: LimitKind::Items,
                    observed: usize::MAX,
                    maximum: limits.max_items(),
                })?;
            limits.check_items(next).map_err(map_limits)?;
            values
                .try_reserve(1)
                .map_err(|_| WriteError::Allocation { requested: 1 })?;
            values.push(value);
            rest = &rest[length..];
        }
        Ok(values)
    } else {
        let mut values = Vec::new();
        limits.check_items(1).map_err(map_limits)?;
        values
            .try_reserve_exact(1)
            .map_err(|_| WriteError::Allocation { requested: 1 })?;
        values.push(varint(field.payload)?);
        Ok(values)
    }
}

/// The repeated (packed or not) varint field `number` of `data`.
fn varints(data: &[u8], number: u32, limits: &Limits) -> Result<Vec<u64>> {
    let mut output = Vec::new();
    for field in fields(data, limits)? {
        if field.number == number {
            let values = packed_or_single(&field, limits)?;
            let next = output
                .len()
                .checked_add(values.len())
                .ok_or(WriteError::LimitExceeded {
                    kind: LimitKind::Items,
                    observed: usize::MAX,
                    maximum: limits.max_items(),
                })?;
            limits.check_items(next).map_err(map_limits)?;
            output
                .try_reserve(values.len())
                .map_err(|_| WriteError::Allocation {
                    requested: values.len(),
                })?;
            output.extend(values);
        }
    }
    Ok(output)
}

/// The formula state of one cell (`row`, `column`) of the table whose
/// drawable is `info`: whether it holds a formula or another formula
/// (in the table or elsewhere) reads it, over the document's formula owners
/// (`TSCE.FormulaOwnerDependenciesArchive` payloads); `object` resolves a
/// dependency tile's payload. Range dependencies on the table are treated
/// as reading every cell.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed owner.
pub fn cell_formula_state<'a>(
    owners: &[&'a [u8]],
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    info: u64,
    row: u32,
    column: u32,
    limits: &Limits,
) -> Result<FormulaState> {
    limits.check_items(owners.len()).map_err(map_limits)?;
    let mut own = None;
    for owner in owners {
        if let Some(reference) = single_bytes(owner, 11, limits)?
            && single_varint(reference, 1, limits)? == Some(info)
        {
            own = Some(*owner);
        }
    }
    let Some(own) = own else {
        return Ok(FormulaState::Inert);
    };
    let (row, column) = (u64::from(row), u64::from(column));
    if single_bytes(own, 5, limits)?.is_some_and(|ranges| !ranges.is_empty())
        || single_bytes(own, 15, limits)?.is_some_and(|tiles| !tiles.is_empty())
    {
        return Ok(FormulaState::Formulas);
    }
    for record in cell_records(own, object, limits)? {
        if single_varint(record, 1, limits)? == Some(column)
            && single_varint(record, 2, limits)? == Some(row)
        {
            return Ok(FormulaState::Formulas);
        }
        if let Some(edges) = single_bytes(record, 6, limits)? {
            let rows = varints(edges, 1, limits)?;
            let columns = varints(edges, 2, limits)?;
            if rows
                .iter()
                .zip(&columns)
                .any(|edge| edge == (&row, &column))
            {
                return Ok(FormulaState::Referenced);
            }
        }
    }
    let internal = single_varint(own, 2, limits)?.ok_or(WriteError::Invalid)?;
    let own_uid = single_bytes(own, 1, limits)?.ok_or(WriteError::Invalid)?;
    for owner in owners {
        if std::ptr::eq(*owner, own) {
            continue;
        }
        // The table's merge owner (kind 5, based on the table's owner)
        // ranges over merged regions, not over values.
        if single_varint(owner, 3, limits)? == Some(MERGE_OWNER_KIND)
            && single_bytes(owner, 12, limits)? == Some(own_uid)
        {
            continue;
        }
        for record in cell_records(owner, object, limits)? {
            if let Some(edges) = single_bytes(record, 6, limits)? {
                let rows = varints(edges, 3, limits)?;
                let columns = varints(edges, 4, limits)?;
                let targets = varints(edges, 5, limits)?;
                if targets
                    .iter()
                    .zip(rows.iter().zip(&columns))
                    .any(|(target, edge)| *target == internal && edge == (&row, &column))
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(ranges) = single_bytes(owner, 5, limits)? {
            for dependency in fields(ranges, limits)? {
                if dependency.number != 2 {
                    continue;
                }
                if single_bytes(dependency.payload, 3, limits)?.is_some() {
                    return Ok(FormulaState::Referenced);
                }
                if let Some(range) = single_bytes(dependency.payload, 4, limits)?
                    && single_varint(range, 1, limits)? == Some(internal)
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(tiled) = single_bytes(owner, 15, limits)? {
            for field in fields(tiled, limits)? {
                if field.number != 1 {
                    continue;
                }
                let tile = single_varint(field.payload, 1, limits)?
                    .and_then(object)
                    .ok_or(WriteError::Invalid)?;
                if single_varint(tile, 1, limits)? == Some(internal) {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
    }
    Ok(FormulaState::Inert)
}

/// The formula state of the table whose drawable is `info`, over the
/// formula owners (`TSCE.FormulaOwnerDependenciesArchive` payloads) of the
/// document; `object` resolves a dependency tile's payload.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed owner.
pub fn table_formula_state<'a>(
    owners: &[&'a [u8]],
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    info: u64,
    limits: &Limits,
) -> Result<FormulaState> {
    limits.check_items(owners.len()).map_err(map_limits)?;
    let mut own = None;
    for owner in owners {
        if let Some(reference) = single_bytes(owner, 11, limits)?
            && single_varint(reference, 1, limits)? == Some(info)
        {
            own = Some(*owner);
        }
    }
    let Some(own) = own else {
        return Ok(FormulaState::Inert);
    };
    if !cell_records(own, object, limits)?.is_empty()
        || single_bytes(own, 5, limits)?.is_some_and(|ranges| !ranges.is_empty())
        || single_bytes(own, 15, limits)?.is_some_and(|tiles| !tiles.is_empty())
    {
        return Ok(FormulaState::Formulas);
    }
    let internal = single_varint(own, 2, limits)?.ok_or(WriteError::Invalid)?;
    for owner in owners {
        if std::ptr::eq(*owner, own) {
            continue;
        }
        for record in cell_records(owner, object, limits)? {
            if let Some(edges) = single_bytes(record, 6, limits)? {
                for field in fields(edges, limits)? {
                    if field.number != 5 {
                        continue;
                    }
                    let owners = packed_or_single(&field, limits)?;
                    if owners.contains(&internal) {
                        return Ok(FormulaState::Referenced);
                    }
                }
            }
        }
        if let Some(ranges) = single_bytes(owner, 5, limits)? {
            for dependency in fields(ranges, limits)? {
                if dependency.number != 2 {
                    continue;
                }
                if single_bytes(dependency.payload, 3, limits)?.is_some() {
                    return Ok(FormulaState::Referenced);
                }
                if let Some(range) = single_bytes(dependency.payload, 4, limits)?
                    && single_varint(range, 1, limits)? == Some(internal)
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(tiled) = single_bytes(owner, 15, limits)? {
            for field in fields(tiled, limits)? {
                if field.number != 1 {
                    continue;
                }
                let tile = single_varint(field.payload, 1, limits)?
                    .and_then(object)
                    .ok_or(WriteError::Invalid)?;
                if single_varint(tile, 1, limits)? == Some(internal) {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
    }
    Ok(FormulaState::Inert)
}

#[cfg(test)]
mod tests {
    use super::{
        CellInput, WriteError, cell_bytes, header_adjust, header_adjust_many, list_intern,
        list_release, tile_cell, tile_set_cell,
    };
    use super::{bytes_field, fixed32_field, varint_field};
    use crate::{BncCell, Limits};

    const TEST_LIMITS: Limits = Limits::DEFAULT;

    fn tile_with_row(row: u64) -> Vec<u8> {
        let mut tile = Vec::new();
        for number in 1..=3 {
            varint_field(&mut tile, number, 0, &TEST_LIMITS).unwrap();
        }
        varint_field(&mut tile, 4, 1, &TEST_LIMITS).unwrap();
        let cell = BncCell::minimal().encode(&TEST_LIMITS).unwrap();
        let mut storage = cell.clone();
        storage.extend_from_slice(&[]);
        let mut entry = Vec::new();
        varint_field(&mut entry, 1, row, &TEST_LIMITS).unwrap();
        varint_field(&mut entry, 2, 1, &TEST_LIMITS).unwrap();
        bytes_field(&mut entry, 3, &[], &TEST_LIMITS).unwrap();
        bytes_field(&mut entry, 4, &[], &TEST_LIMITS).unwrap();
        varint_field(&mut entry, 5, 5, &TEST_LIMITS).unwrap();
        bytes_field(&mut entry, 6, &storage, &TEST_LIMITS).unwrap();
        bytes_field(&mut entry, 7, &[0, 0, 0xff, 0xff], &TEST_LIMITS).unwrap();
        varint_field(&mut entry, 8, 0, &TEST_LIMITS).unwrap();
        bytes_field(&mut tile, 5, &entry, &TEST_LIMITS).unwrap();
        varint_field(&mut tile, 6, 5, &TEST_LIMITS).unwrap();
        varint_field(&mut tile, 7, 1, &TEST_LIMITS).unwrap();
        tile
    }

    #[test]
    fn rows_are_inserted_in_order_and_cells_replaced() {
        let tile = tile_with_row(3);
        let mut number = BncCell::minimal();
        number.set_number(4.5).unwrap();
        let number = number.encode(&TEST_LIMITS).unwrap();
        let before = tile_set_cell(&tile, 1, 1, 2, Some(&number), &TEST_LIMITS).unwrap();
        assert_eq!(
            tile_cell(&before.tile, 1, 1, &TEST_LIMITS)
                .unwrap()
                .as_deref(),
            Some(number.as_slice())
        );
        assert!(
            tile_cell(&before.tile, 3, 0, &TEST_LIMITS)
                .unwrap()
                .is_some()
        );
        let after = tile_set_cell(&before.tile, 3, 1, 2, Some(&number), &TEST_LIMITS).unwrap();
        assert!(
            tile_cell(&after.tile, 3, 1, &TEST_LIMITS)
                .unwrap()
                .is_some()
        );
        assert!(
            tile_cell(&after.tile, 3, 0, &TEST_LIMITS)
                .unwrap()
                .is_some()
        );
        let removed = tile_set_cell(&after.tile, 3, 0, 2, None, &TEST_LIMITS).unwrap();
        assert!(removed.previous.is_some());
        assert!(
            tile_cell(&removed.tile, 3, 0, &TEST_LIMITS)
                .unwrap()
                .is_none()
        );

        let narrow = TEST_LIMITS.with_output_bytes(1).unwrap();
        assert!(matches!(
            tile_set_cell(&tile, 1, 1, 2, Some(&number), &narrow),
            Err(WriteError::LimitExceeded { .. })
        ));
    }

    #[test]
    fn strings_intern_and_release() {
        let mut list = Vec::new();
        varint_field(&mut list, 1, 1, &TEST_LIMITS).unwrap();
        varint_field(&mut list, 2, 1, &TEST_LIMITS).unwrap();
        let (list, key) = list_intern(&list, "hello", &TEST_LIMITS).unwrap();
        assert_eq!(key, 1);
        let (list, again) = list_intern(&list, "hello", &TEST_LIMITS).unwrap();
        assert_eq!(again, 1);
        let (list, other) = list_intern(&list, "world", &TEST_LIMITS).unwrap();
        assert_eq!(other, 2);
        let list = list_release(&list, 1, &TEST_LIMITS).unwrap();
        let list = list_release(&list, 1, &TEST_LIMITS).unwrap();
        assert!(list_release(&list, 1, &TEST_LIMITS).is_err());
        let (cell, released) =
            cell_bytes(None, &CellInput::Text("x"), Some(2), &TEST_LIMITS).unwrap();
        assert!(cell.is_some());
        assert_eq!(released, None);
        let (cleared, released) =
            cell_bytes(cell.as_deref(), &CellInput::Empty, None, &TEST_LIMITS).unwrap();
        assert_eq!(cleared, None);
        assert_eq!(released, Some(2));
    }

    #[test]
    fn header_counts_are_adjusted_or_created() {
        let bucket = {
            let mut bucket = Vec::new();
            varint_field(&mut bucket, 1, 1, &TEST_LIMITS).unwrap();
            bucket
        };
        let added = header_adjust(&bucket, 4, 1, &TEST_LIMITS).unwrap();
        let again = header_adjust(&added, 4, 1, &TEST_LIMITS).unwrap();
        let lowered = header_adjust(&again, 4, -2, &TEST_LIMITS).unwrap();
        assert!(header_adjust(&lowered, 4, -1, &TEST_LIMITS).is_err());
        assert!(header_adjust(&bucket, 2, -1, &TEST_LIMITS).is_err());
    }

    #[test]
    fn header_adjust_many_matches_one_header_at_a_time() {
        let mut bucket = Vec::new();
        varint_field(&mut bucket, 1, 1, &TEST_LIMITS).unwrap();
        for (index, cells) in [(1, 2), (3, 1), (5, 1)] {
            bucket = header_adjust(&bucket, index, cells, &TEST_LIMITS).unwrap();
        }
        let deltas = [(0, 1), (3, -1), (4, 2), (5, 1), (6, 0), (9, 1)];
        let mut sequential = bucket.clone();
        for (index, delta) in deltas {
            sequential = header_adjust(&sequential, index, delta, &TEST_LIMITS).unwrap();
        }
        assert_eq!(
            header_adjust_many(&bucket, &deltas, &TEST_LIMITS).unwrap(),
            sequential
        );
        assert_eq!(
            header_adjust_many(&bucket, &[(4, 1), (4, 1)], &TEST_LIMITS),
            Err(WriteError::Invalid)
        );
    }

    #[test]
    fn header_adjust_many_ignores_removals_the_bucket_does_not_count() {
        let mut bucket = Vec::new();
        varint_field(&mut bucket, 1, 1, &TEST_LIMITS).unwrap();
        bucket = header_adjust(&bucket, 3, 1, &TEST_LIMITS).unwrap();
        // A header carrying a size stays when it counts no cell.
        let mut sized = Vec::new();
        varint_field(&mut sized, 1, 5, &TEST_LIMITS).unwrap();
        fixed32_field(&mut sized, 2, 20.0f32.to_bits(), &TEST_LIMITS).unwrap();
        varint_field(&mut sized, 3, 0, &TEST_LIMITS).unwrap();
        varint_field(&mut sized, 4, 1, &TEST_LIMITS).unwrap();
        bytes_field(&mut bucket, 2, &sized, &TEST_LIMITS).unwrap();

        let unchanged = header_adjust_many(&bucket, &[(2, -1)], &TEST_LIMITS).unwrap();
        assert_eq!(unchanged, bucket);
        let lowered = header_adjust_many(&bucket, &[(3, -4), (5, -2)], &TEST_LIMITS).unwrap();
        let mut expected = header_adjust(&bucket, 3, -1, &TEST_LIMITS).unwrap();
        expected = header_adjust(&expected, 5, -1, &TEST_LIMITS).unwrap();
        assert_eq!(lowered, expected);
        assert_eq!(
            header_adjust_many(&lowered, &[(5, -1)], &TEST_LIMITS).unwrap(),
            lowered
        );
    }

    #[test]
    fn header_adjust_many_refuses_unordered_buckets_and_bounds_growth() {
        let mut header = Vec::new();
        varint_field(&mut header, 1, 4, &TEST_LIMITS).unwrap();
        varint_field(&mut header, 4, 1, &TEST_LIMITS).unwrap();
        let mut unordered = Vec::new();
        bytes_field(&mut unordered, 2, &header, &TEST_LIMITS).unwrap();
        let mut lower = Vec::new();
        varint_field(&mut lower, 1, 2, &TEST_LIMITS).unwrap();
        varint_field(&mut lower, 4, 1, &TEST_LIMITS).unwrap();
        bytes_field(&mut unordered, 2, &lower, &TEST_LIMITS).unwrap();
        assert_eq!(
            header_adjust_many(&unordered, &[(3, 1)], &TEST_LIMITS),
            Err(WriteError::Unsupported)
        );

        let deltas = [(1, 1), (4, i64::from(u32::MAX)), (u32::MAX, i64::MAX)];
        let mut bucket = Vec::new();
        bytes_field(&mut bucket, 2, &header, &TEST_LIMITS).unwrap();
        let grown = header_adjust_many(&bucket, &deltas, &TEST_LIMITS).unwrap();
        assert!(grown.len() <= bucket.len() + deltas.len() * super::HEADER_BYTES);
        let mut sequential = bucket.clone();
        for (index, delta) in deltas {
            sequential = header_adjust(&sequential, index, delta, &TEST_LIMITS).unwrap();
        }
        assert_eq!(grown, sequential);
    }

    #[test]
    fn packed_varints_charge_decoded_items_not_payload_bytes() {
        let limits = TEST_LIMITS.with_items(1).unwrap();
        let reference = {
            let mut value = Vec::new();
            varint_field(&mut value, 1, 42, &TEST_LIMITS).unwrap();
            value
        };
        let packed_ten_byte = [0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01];
        let edges = {
            let mut value = Vec::new();
            bytes_field(&mut value, 1, &packed_ten_byte, &TEST_LIMITS).unwrap();
            varint_field(&mut value, 2, 1, &TEST_LIMITS).unwrap();
            value
        };
        let record = {
            let mut value = Vec::new();
            varint_field(&mut value, 1, 0, &TEST_LIMITS).unwrap();
            varint_field(&mut value, 2, 0, &TEST_LIMITS).unwrap();
            bytes_field(&mut value, 6, &edges, &TEST_LIMITS).unwrap();
            value
        };
        let inline_records = {
            let mut value = Vec::new();
            bytes_field(&mut value, 1, &record, &TEST_LIMITS).unwrap();
            value
        };
        let owner = {
            let mut value = Vec::new();
            bytes_field(&mut value, 11, &reference, &TEST_LIMITS).unwrap();
            varint_field(&mut value, 2, 7, &TEST_LIMITS).unwrap();
            bytes_field(&mut value, 1, b"owner", &TEST_LIMITS).unwrap();
            bytes_field(&mut value, 4, &inline_records, &TEST_LIMITS).unwrap();
            value
        };
        let owners = [owner.as_slice()];
        let state = super::cell_formula_state(&owners, &|_object| None, 42, 1, 1, &limits).unwrap();
        assert_eq!(state, super::FormulaState::Inert);
    }
}
