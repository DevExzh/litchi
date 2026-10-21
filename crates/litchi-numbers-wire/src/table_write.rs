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

use litchi_iwa_common::wire::WireView;

use crate::{BncCell, ListKind};

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
}

impl std::fmt::Display for WriteError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Invalid => "malformed table payload",
            Self::Unsupported => "table payload shape outside the shared cell writer",
        })
    }
}

impl std::error::Error for WriteError {}

type Result<T> = std::result::Result<T, WriteError>;

const MISSING: [u8; 2] = [0xff, 0xff];
/// `owner_kind` of a table's merge owner.
const MERGE_OWNER_KIND: u64 = 5;

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return;
        }
        output.push(byte | 0x80);
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
}

fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) {
    push_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
}

struct Field<'a> {
    number: u32,
    wire_type: u8,
    raw: &'a [u8],
    payload: &'a [u8],
}

fn fields(data: &[u8]) -> Result<Vec<Field<'_>>> {
    let view = WireView::parse(data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_error| WriteError::Invalid)?;
    Ok(view
        .fields()
        .map(|field| Field {
            number: field.number(),
            wire_type: field.wire_type(),
            raw: field.raw(),
            payload: field.payload(),
        })
        .collect())
}

fn varint(payload: &[u8]) -> Result<u64> {
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|_error| WriteError::Invalid)?;
    if length != payload.len() {
        return Err(WriteError::Invalid);
    }
    Ok(value)
}

fn single_varint(data: &[u8], number: u32) -> Result<Option<u64>> {
    let mut found = None;
    for field in fields(data)? {
        if field.number == number {
            if field.wire_type != 0 || found.replace(varint(field.payload)?).is_some() {
                return Err(WriteError::Invalid);
            }
        }
    }
    Ok(found)
}

fn single_bytes(data: &[u8], number: u32) -> Result<Option<&[u8]>> {
    let mut found = None;
    for field in fields(data)? {
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
    mut map: impl FnMut(&Field<'_>) -> Result<Option<Vec<u8>>>,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(data.len());
    for field in fields(data)? {
        match map(&field)? {
            Some(replacement) => output.extend_from_slice(&replacement),
            None => output.extend_from_slice(field.raw),
        }
    }
    Ok(output)
}

/// The cell slots of one BNC storage: `(start, end)` per column.
fn slots(storage: &[u8], offsets: &[u8], unit: usize) -> Result<Vec<Option<(usize, usize)>>> {
    if offsets.len() % 2 != 0 {
        return Err(WriteError::Invalid);
    }
    let starts: Vec<Option<usize>> = offsets
        .chunks_exact(2)
        .map(|pair| {
            let raw = [pair[0], pair[1]];
            (raw != MISSING).then(|| usize::from(u16::from_le_bytes(raw)) * unit)
        })
        .collect();
    let present: Vec<usize> = starts.iter().flatten().copied().collect();
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.last().is_some_and(|last| *last >= storage.len())
        || present.first().is_some_and(|first| *first != 0)
    {
        return Err(WriteError::Unsupported);
    }
    let mut next = present.iter().skip(1);
    Ok(starts
        .into_iter()
        .map(|start| start.map(|start| (start, next.next().copied().unwrap_or(storage.len()))))
        .collect())
}

/// Encode a cell list as BNC storage and offsets in `unit`-byte units.
fn encode(cells: &[Option<&[u8]>], unit: usize) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len() * 2);
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
pub fn tile_cell(tile: &[u8], local_row: u32, column: u32) -> Result<Option<Vec<u8>>> {
    for field in fields(tile)? {
        if field.number != 5 || single_varint(field.payload, 1)? != Some(u64::from(local_row)) {
            continue;
        }
        let (Some(storage), Some(offsets)) = (
            single_bytes(field.payload, 6)?,
            single_bytes(field.payload, 7)?,
        ) else {
            return Err(WriteError::Unsupported);
        };
        let unit = if single_varint(field.payload, 8)?.unwrap_or(0) != 0 {
            4
        } else {
            1
        };
        let cells = slots(storage, offsets, unit)?;
        return Ok(cells
            .get(column as usize)
            .copied()
            .flatten()
            .map(|(start, end)| storage[start..end].to_vec()));
    }
    Ok(None)
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
) -> Result<TileCellWrite> {
    if single_varint(tile, 6)? != Some(5) || single_varint(tile, 7)? != Some(1) {
        return Err(WriteError::Unsupported);
    }
    let column = column as usize;
    let columns_usize = columns as usize;
    if column >= columns_usize {
        return Err(WriteError::Invalid);
    }
    let rows: Vec<u64> = fields(tile)?
        .iter()
        .filter(|field| field.number == 5)
        .map(|field| single_varint(field.payload, 1).and_then(|row| row.ok_or(WriteError::Invalid)))
        .collect::<Result<_>>()?;
    let numrows = single_varint(tile, 4)?.ok_or(WriteError::Invalid)?;
    let counts_rows = numrows == rows.len() as u64;
    let mut previous = None;
    let mut found = false;
    let mut rewritten_row = |row: &[u8]| -> Result<Option<Vec<u8>>> {
        let (Some(storage), Some(offsets)) = (single_bytes(row, 6)?, single_bytes(row, 7)?) else {
            return Err(WriteError::Unsupported);
        };
        let unit = if single_varint(row, 8)?.unwrap_or(0) != 0 {
            4
        } else {
            1
        };
        let mut cells: Vec<Option<&[u8]>> = slots(storage, offsets, unit)?
            .into_iter()
            .map(|slot| slot.map(|(start, end)| &storage[start..end]))
            .collect();
        if cells.len() < columns_usize {
            cells.resize(columns_usize, None);
        }
        previous = cells[column].map(<[u8]>::to_vec);
        cells[column] = cell;
        let count = cells.iter().filter(|cell| cell.is_some()).count() as u64;
        let (storage, offsets) = encode(&cells, unit)?;
        Ok(Some(rebuild(row, |field| {
            Ok(match field.number {
                2 => {
                    let mut output = Vec::new();
                    varint_field(&mut output, 2, count);
                    Some(output)
                },
                6 => {
                    let mut output = Vec::new();
                    bytes_field(&mut output, 6, &storage);
                    Some(output)
                },
                7 => {
                    let mut output = Vec::new();
                    bytes_field(&mut output, 7, &offsets);
                    Some(output)
                },
                _ => None,
            })
        })?))
    };
    // A new row record for a missing row, in the canonical BNC shape.
    let new_row = |cell: &[u8]| -> Result<Vec<u8>> {
        let mut cells: Vec<Option<&[u8]>> = vec![None; columns_usize];
        cells[column] = Some(cell);
        let (storage, offsets) = encode(&cells, 1)?;
        let mut row = Vec::new();
        varint_field(&mut row, 1, u64::from(local_row));
        varint_field(&mut row, 2, 1);
        bytes_field(&mut row, 3, &[]);
        bytes_field(&mut row, 4, &[]);
        varint_field(&mut row, 5, 5);
        bytes_field(&mut row, 6, &storage);
        bytes_field(&mut row, 7, &offsets);
        varint_field(&mut row, 8, 0);
        let mut field = Vec::new();
        bytes_field(&mut field, 5, &row);
        Ok(field)
    };
    let insert_before = rows.iter().position(|row| *row > u64::from(local_row));
    let exists = rows.contains(&u64::from(local_row));
    let mut row_index = 0_usize;
    let mut inserted = false;
    let mut output = Vec::with_capacity(tile.len() + 64);
    for field in fields(tile)? {
        if field.number == 5 {
            let this = row_index;
            row_index += 1;
            if !exists && Some(this) == insert_before {
                if let Some(cell) = cell {
                    output.extend_from_slice(&new_row(cell)?);
                    inserted = true;
                }
            }
            if single_varint(field.payload, 1)? == Some(u64::from(local_row)) {
                found = true;
                let row = rewritten_row(field.payload)?.ok_or(WriteError::Invalid)?;
                bytes_field(&mut output, 5, &row);
                continue;
            }
        }
        output.extend_from_slice(field.raw);
    }
    if !exists && !inserted {
        if let Some(cell) = cell {
            output.extend_from_slice(&new_row(cell)?);
            inserted = true;
        }
    }
    let _ = found;
    if inserted {
        let new_numrows = if counts_rows {
            numrows + 1
        } else {
            numrows.max(u64::from(local_row) + 1)
        };
        output = rebuild(&output, |field| {
            Ok((field.number == 4).then(|| {
                let mut replacement = Vec::new();
                varint_field(&mut replacement, 4, new_numrows);
                replacement
            }))
        })?;
    }
    Ok(TileCellWrite {
        tile: output,
        previous,
    })
}

/// Add `delta` to the cell count of header `index` in a header bucket,
/// creating the header (size 0, visible) when it has none.
///
/// # Errors
///
/// [`WriteError::Unsupported`] when a count would drop below zero.
pub fn header_adjust(bucket: &[u8], index: u32, delta: i64) -> Result<Vec<u8>> {
    if delta == 0 {
        return Ok(bucket.to_vec());
    }
    let mut done = false;
    let mut output = rebuild(bucket, |field| {
        if field.number != 2 || single_varint(field.payload, 1)? != Some(u64::from(index)) {
            return Ok(None);
        }
        done = true;
        let cells = single_varint(field.payload, 4)?.ok_or(WriteError::Invalid)?;
        let after = i64::try_from(cells)
            .ok()
            .and_then(|cells| cells.checked_add(delta))
            .filter(|after| *after >= 0)
            .ok_or(WriteError::Unsupported)?;
        let header = rebuild(field.payload, |inner| {
            Ok((inner.number == 4).then(|| {
                let mut replacement = Vec::new();
                varint_field(&mut replacement, 4, after as u64);
                replacement
            }))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 2, &header);
        Ok(Some(replacement))
    })?;
    if !done {
        if delta < 0 {
            return Err(WriteError::Unsupported);
        }
        // Headers stay sorted by index.
        let mut header = Vec::new();
        varint_field(&mut header, 1, u64::from(index));
        fixed32_field(&mut header, 2, 0);
        varint_field(&mut header, 3, 0);
        varint_field(&mut header, 4, delta as u64);
        let mut entry = Vec::new();
        bytes_field(&mut entry, 2, &header);
        let mut placed = false;
        let mut sorted = Vec::with_capacity(output.len() + entry.len());
        for field in fields(&output)? {
            if !placed
                && field.number == 2
                && single_varint(field.payload, 1)?.is_some_and(|other| other > u64::from(index))
            {
                sorted.extend_from_slice(&entry);
                placed = true;
            }
            sorted.extend_from_slice(field.raw);
        }
        if !placed {
            sorted.extend_from_slice(&entry);
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
pub fn list_intern(list: &[u8], text: &str) -> Result<(Vec<u8>, u32)> {
    if fields(list)?.iter().any(|field| field.number == 4) {
        return Err(WriteError::Unsupported);
    }
    let mut key = None;
    let output = rebuild(list, |field| {
        if field.number != 3
            || key.is_some()
            || single_bytes(field.payload, 3)? != Some(text.as_bytes())
        {
            return Ok(None);
        }
        let entry_key = single_varint(field.payload, 1)?.ok_or(WriteError::Invalid)?;
        key = Some(u32::try_from(entry_key).map_err(|_error| WriteError::Invalid)?);
        let count = single_varint(field.payload, 2)?.ok_or(WriteError::Invalid)?;
        let entry = rebuild(field.payload, |inner| {
            Ok((inner.number == 2).then(|| {
                let mut replacement = Vec::new();
                varint_field(&mut replacement, 2, count + 1);
                replacement
            }))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 3, &entry);
        Ok(Some(replacement))
    })?;
    if let Some(key) = key {
        return Ok((output, key));
    }
    let next = single_varint(list, 2)?.ok_or(WriteError::Invalid)?.max(1);
    let key = u32::try_from(next).map_err(|_error| WriteError::Unsupported)?;
    let mut output = rebuild(list, |field| {
        Ok((field.number == 2).then(|| {
            let mut replacement = Vec::new();
            varint_field(&mut replacement, 2, next + 1);
            replacement
        }))
    })?;
    let mut entry = Vec::new();
    varint_field(&mut entry, 1, next);
    varint_field(&mut entry, 2, 1);
    bytes_field(&mut entry, 3, text.as_bytes());
    bytes_field(&mut output, 3, &entry);
    Ok((output, key))
}

/// Drop one reference to string entry `key`; an entry left without
/// references goes.
///
/// # Errors
///
/// [`WriteError::Unsupported`] for an entry holding an object reference or
/// a key the list does not hold.
pub fn list_release(list: &[u8], key: u32) -> Result<Vec<u8>> {
    let mut done = false;
    let output = rebuild(list, |field| {
        if field.number != 3 || single_varint(field.payload, 1)? != Some(u64::from(key)) {
            return Ok(None);
        }
        if single_bytes(field.payload, 4)?.is_some() {
            return Err(WriteError::Unsupported);
        }
        done = true;
        let count = single_varint(field.payload, 2)?.ok_or(WriteError::Invalid)?;
        if count <= 1 {
            return Ok(Some(Vec::new()));
        }
        let entry = rebuild(field.payload, |inner| {
            Ok((inner.number == 2).then(|| {
                let mut replacement = Vec::new();
                varint_field(&mut replacement, 2, count - 1);
                replacement
            }))
        })?;
        let mut replacement = Vec::new();
        bytes_field(&mut replacement, 3, &entry);
        Ok(Some(replacement))
    })?;
    if !done {
        return Err(WriteError::Unsupported);
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
) -> Result<(Option<Vec<u8>>, Option<u32>)> {
    let mut cell = match previous {
        Some(bytes) => BncCell::parse(bytes).map_err(|_error| WriteError::Unsupported)?,
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
            if previous.is_some() && cell.list_keys().next().is_none() && cell_is_bare(&cell) {
                return Ok((None, released));
            }
            if previous.is_none() {
                return Ok((None, None));
            }
        },
    }
    Ok((Some(cell.encode()), released))
}

/// Whether a cleared cell carries nothing worth storing.
fn cell_is_bare(cell: &BncCell) -> bool {
    cell.encode() == BncCell::minimal().encode()
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
) -> Result<Vec<&'a [u8]>> {
    let mut records = Vec::new();
    if let Some(inline) = single_bytes(owner, 4)? {
        for field in fields(inline)? {
            if field.number == 1 {
                records.push(field.payload);
            }
        }
    }
    if let Some(tiled) = single_bytes(owner, 13)? {
        for field in fields(tiled)? {
            if field.number != 1 {
                continue;
            }
            let tile = single_varint(field.payload, 1)?
                .and_then(object)
                .ok_or(WriteError::Invalid)?;
            for record in fields(tile)? {
                if record.number == 4 {
                    records.push(record.payload);
                }
            }
        }
    }
    Ok(records)
}

fn packed_or_single(field: &Field<'_>) -> Result<Vec<u64>> {
    if field.wire_type == 2 {
        let mut values = Vec::new();
        let mut rest = field.payload;
        while !rest.is_empty() {
            let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(rest)
                .map_err(|_error| WriteError::Invalid)?;
            values.push(value);
            rest = &rest[length..];
        }
        Ok(values)
    } else {
        Ok(vec![varint(field.payload)?])
    }
}

/// The repeated (packed or not) varint field `number` of `data`.
fn varints(data: &[u8], number: u32) -> Result<Vec<u64>> {
    let mut output = Vec::new();
    for field in fields(data)? {
        if field.number == number {
            output.extend(packed_or_single(&field)?);
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
) -> Result<FormulaState> {
    let mut own = None;
    for owner in owners {
        if let Some(reference) = single_bytes(owner, 11)?
            && single_varint(reference, 1)? == Some(info)
        {
            own = Some(*owner);
        }
    }
    let Some(own) = own else {
        return Ok(FormulaState::Inert);
    };
    let (row, column) = (u64::from(row), u64::from(column));
    if single_bytes(own, 5)?.is_some_and(|ranges| !ranges.is_empty())
        || single_bytes(own, 15)?.is_some_and(|tiles| !tiles.is_empty())
    {
        return Ok(FormulaState::Formulas);
    }
    for record in cell_records(own, object)? {
        if single_varint(record, 1)? == Some(column) && single_varint(record, 2)? == Some(row) {
            return Ok(FormulaState::Formulas);
        }
        if let Some(edges) = single_bytes(record, 6)? {
            let rows = varints(edges, 1)?;
            let columns = varints(edges, 2)?;
            if rows
                .iter()
                .zip(&columns)
                .any(|edge| edge == (&row, &column))
            {
                return Ok(FormulaState::Referenced);
            }
        }
    }
    let internal = single_varint(own, 2)?.ok_or(WriteError::Invalid)?;
    let own_uid = single_bytes(own, 1)?.ok_or(WriteError::Invalid)?;
    for owner in owners {
        if std::ptr::eq(*owner, own) {
            continue;
        }
        // The table's merge owner (kind 5, based on the table's owner)
        // ranges over merged regions, not over values.
        if single_varint(owner, 3)? == Some(MERGE_OWNER_KIND)
            && single_bytes(owner, 12)? == Some(own_uid)
        {
            continue;
        }
        for record in cell_records(owner, object)? {
            if let Some(edges) = single_bytes(record, 6)? {
                let rows = varints(edges, 3)?;
                let columns = varints(edges, 4)?;
                let targets = varints(edges, 5)?;
                if targets
                    .iter()
                    .zip(rows.iter().zip(&columns))
                    .any(|(target, edge)| *target == internal && edge == (&row, &column))
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(ranges) = single_bytes(owner, 5)? {
            for dependency in fields(ranges)? {
                if dependency.number != 2 {
                    continue;
                }
                if single_bytes(dependency.payload, 3)?.is_some() {
                    return Ok(FormulaState::Referenced);
                }
                if let Some(range) = single_bytes(dependency.payload, 4)?
                    && single_varint(range, 1)? == Some(internal)
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(tiled) = single_bytes(owner, 15)? {
            for field in fields(tiled)? {
                if field.number != 1 {
                    continue;
                }
                let tile = single_varint(field.payload, 1)?
                    .and_then(object)
                    .ok_or(WriteError::Invalid)?;
                if single_varint(tile, 1)? == Some(internal) {
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
) -> Result<FormulaState> {
    let mut own = None;
    for owner in owners {
        if let Some(reference) = single_bytes(owner, 11)?
            && single_varint(reference, 1)? == Some(info)
        {
            own = Some(*owner);
        }
    }
    let Some(own) = own else {
        return Ok(FormulaState::Inert);
    };
    if !cell_records(own, object)?.is_empty()
        || single_bytes(own, 5)?.is_some_and(|ranges| !ranges.is_empty())
        || single_bytes(own, 15)?.is_some_and(|tiles| !tiles.is_empty())
    {
        return Ok(FormulaState::Formulas);
    }
    let internal = single_varint(own, 2)?.ok_or(WriteError::Invalid)?;
    for owner in owners {
        if std::ptr::eq(*owner, own) {
            continue;
        }
        for record in cell_records(owner, object)? {
            if let Some(edges) = single_bytes(record, 6)? {
                for field in fields(edges)? {
                    if field.number != 5 {
                        continue;
                    }
                    let owners: Vec<u64> = if field.wire_type == 2 {
                        let mut values = Vec::new();
                        let mut rest = field.payload;
                        while !rest.is_empty() {
                            let (value, length) =
                                litchi_iwa_common::varint::decode_varint_from_bytes(rest)
                                    .map_err(|_error| WriteError::Invalid)?;
                            values.push(value);
                            rest = &rest[length..];
                        }
                        values
                    } else {
                        vec![varint(field.payload)?]
                    };
                    if owners.contains(&internal) {
                        return Ok(FormulaState::Referenced);
                    }
                }
            }
        }
        if let Some(ranges) = single_bytes(owner, 5)? {
            for dependency in fields(ranges)? {
                if dependency.number != 2 {
                    continue;
                }
                if single_bytes(dependency.payload, 3)?.is_some() {
                    return Ok(FormulaState::Referenced);
                }
                if let Some(range) = single_bytes(dependency.payload, 4)?
                    && single_varint(range, 1)? == Some(internal)
                {
                    return Ok(FormulaState::Referenced);
                }
            }
        }
        if let Some(tiled) = single_bytes(owner, 15)? {
            for field in fields(tiled)? {
                if field.number != 1 {
                    continue;
                }
                let tile = single_varint(field.payload, 1)?
                    .and_then(object)
                    .ok_or(WriteError::Invalid)?;
                if single_varint(tile, 1)? == Some(internal) {
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
        CellInput, cell_bytes, header_adjust, list_intern, list_release, tile_cell, tile_set_cell,
    };
    use super::{bytes_field, varint_field};
    use crate::BncCell;

    fn tile_with_row(row: u64) -> Vec<u8> {
        let mut tile = Vec::new();
        for number in 1..=3 {
            varint_field(&mut tile, number, 0);
        }
        varint_field(&mut tile, 4, 1);
        let cell = BncCell::minimal().encode();
        let mut storage = cell.clone();
        storage.extend_from_slice(&[]);
        let mut entry = Vec::new();
        varint_field(&mut entry, 1, row);
        varint_field(&mut entry, 2, 1);
        bytes_field(&mut entry, 3, &[]);
        bytes_field(&mut entry, 4, &[]);
        varint_field(&mut entry, 5, 5);
        bytes_field(&mut entry, 6, &storage);
        bytes_field(&mut entry, 7, &[0, 0, 0xff, 0xff]);
        varint_field(&mut entry, 8, 0);
        bytes_field(&mut tile, 5, &entry);
        varint_field(&mut tile, 6, 5);
        varint_field(&mut tile, 7, 1);
        tile
    }

    #[test]
    fn rows_are_inserted_in_order_and_cells_replaced() {
        let tile = tile_with_row(3);
        let mut number = BncCell::minimal();
        number.set_number(4.5).unwrap();
        let number = number.encode();
        let before = tile_set_cell(&tile, 1, 1, 2, Some(&number)).unwrap();
        assert_eq!(
            tile_cell(&before.tile, 1, 1).unwrap().as_deref(),
            Some(number.as_slice())
        );
        assert!(tile_cell(&before.tile, 3, 0).unwrap().is_some());
        let after = tile_set_cell(&before.tile, 3, 1, 2, Some(&number)).unwrap();
        assert!(tile_cell(&after.tile, 3, 1).unwrap().is_some());
        assert!(tile_cell(&after.tile, 3, 0).unwrap().is_some());
        let removed = tile_set_cell(&after.tile, 3, 0, 2, None).unwrap();
        assert!(removed.previous.is_some());
        assert!(tile_cell(&removed.tile, 3, 0).unwrap().is_none());
    }

    #[test]
    fn strings_intern_and_release() {
        let mut list = Vec::new();
        varint_field(&mut list, 1, 1);
        varint_field(&mut list, 2, 1);
        let (list, key) = list_intern(&list, "hello").unwrap();
        assert_eq!(key, 1);
        let (list, again) = list_intern(&list, "hello").unwrap();
        assert_eq!(again, 1);
        let (list, other) = list_intern(&list, "world").unwrap();
        assert_eq!(other, 2);
        let list = list_release(&list, 1).unwrap();
        let list = list_release(&list, 1).unwrap();
        assert!(list_release(&list, 1).is_err());
        let (cell, released) = cell_bytes(None, &CellInput::Text("x"), Some(2)).unwrap();
        assert!(cell.is_some());
        assert_eq!(released, None);
        let (cleared, released) = cell_bytes(cell.as_deref(), &CellInput::Empty, None).unwrap();
        assert_eq!(cleared, None);
        assert_eq!(released, Some(2));
    }

    #[test]
    fn header_counts_are_adjusted_or_created() {
        let bucket = {
            let mut bucket = Vec::new();
            varint_field(&mut bucket, 1, 1);
            bucket
        };
        let added = header_adjust(&bucket, 4, 1).unwrap();
        let again = header_adjust(&added, 4, 1).unwrap();
        let lowered = header_adjust(&again, 4, -2).unwrap();
        assert!(header_adjust(&lowered, 4, -1).is_err());
        assert!(header_adjust(&bucket, 2, -1).is_err());
    }
}
