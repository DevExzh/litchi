//! Payload rewrites of one table-axis insertion or deletion: storage tiles,
//! row and column header buckets, the row/column UID map and the stroke
//! sidecar's border layers.
//!
//! Each function maps one native payload to its rewritten form (or `None`
//! when it is unchanged). Fields a rewrite does not own keep their exact
//! bytes. Ported from the host editor's `row_insert`, `column_insert`,
//! `table_delete` and `stroke_layers` storage phases.

use super::wire::{
    Field, bytes_field, fields, push_bytes_field, push_varint_field, rebuild, repeated_bytes,
    repeated_varints, single_bytes, single_varint, varint, varint_field,
};
use super::{Error, Result, reserve};

/// Rows one native header bucket covers.
pub const HEADER_BUCKET_ROWS: u64 = 65_536;
/// The "no cell" marker of a cell-offset table.
const MISSING: [u8; 2] = [0xff, 0xff];

/// One axis edit at a physical index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisEdit {
    /// A blank row or column is inserted before `index` (`index` may equal
    /// the axis length to append).
    Insert(u32),
    /// The row or column at `index` is removed.
    Remove(u32),
}

impl AxisEdit {
    /// The new position of index `value`, or `None` for the removed index.
    pub fn shift(self, value: u32) -> Result<Option<u32>> {
        Ok(match self {
            Self::Insert(index) if value >= index => {
                Some(value.checked_add(1).ok_or(Error::Invalid)?)
            },
            Self::Insert(_) => Some(value),
            Self::Remove(index) if value == index => None,
            Self::Remove(index) if value > index => Some(value - 1),
            Self::Remove(_) => Some(value),
        })
    }
}

/// Rewrite the rows of one storage tile for a row edit.
///
/// `key` is the tile's key and `size` the rows per tile. A removed row goes
/// with its cells (the caller releases their list references); a row that
/// would move into another tile is refused.
pub fn tile_rows(tile: &[u8], key: u32, size: u32, edit: AxisEdit) -> Result<Option<Vec<u8>>> {
    let base = u64::from(key) * u64::from(size);
    let original_rows = repeated_bytes(tile, 5)?;
    let numrows = single_varint(tile, 4)?.ok_or(Error::Invalid)?;
    let mut maximum = None;
    for row in &original_rows {
        let index = single_varint(row, 1)?.ok_or(Error::Invalid)?;
        maximum = maximum.max(Some(index));
    }
    let counts_rows = numrows == u64::try_from(original_rows.len()).map_err(|_e| Error::Invalid)?;
    let counts_extent = maximum.map_or(numrows == 0, |maximum| numrows == maximum + 1);
    if !counts_rows && !counts_extent {
        return Err(Error::Unsupported);
    }
    let mut changed = false;
    let mut kept = 0_u64;
    let mut new_maximum = None;
    let mut output = rebuild(tile, |field| {
        if field.number != 5 {
            return Ok(None);
        }
        let local = single_varint(field.payload, 1)?.ok_or(Error::Invalid)?;
        let global = u32::try_from(base + local).map_err(|_error| Error::Invalid)?;
        let Some(shifted) = edit.shift(global)? else {
            // The removed row goes with its cells; the caller releases their
            // list references.
            changed = true;
            return Ok(Some(Vec::new()));
        };
        let shifted = u64::from(shifted);
        if shifted / u64::from(size) != u64::from(key) {
            return Err(Error::Unsupported);
        }
        kept += 1;
        let local_after = shifted - base;
        new_maximum = new_maximum.max(Some(local_after));
        if local_after == local {
            return Ok(None);
        }
        changed = true;
        let row = rebuild(field.payload, |inner| {
            if inner.number == 1 {
                return varint_field(1, local_after).map(Some);
            }
            Ok(None)
        })?;
        bytes_field(5, &row).map(Some)
    })?;
    if !changed {
        return Ok(None);
    }
    let numrows_after = if counts_rows {
        kept
    } else {
        new_maximum.map_or(0, |maximum| maximum + 1)
    };
    if numrows_after != numrows {
        output = rebuild(&output, |field| {
            if field.number == 4 {
                return varint_field(4, numrows_after).map(Some);
            }
            Ok(None)
        })?;
    }
    Ok(Some(output))
}

/// Splice one cell-offset table (`2` bytes per column) for a column edit.
fn splice_offsets(offsets: &[u8], columns: u32, edit: AxisEdit) -> Result<Option<Vec<u8>>> {
    if offsets.is_empty() {
        return Ok(None);
    }
    if offsets.len() % 2 != 0 {
        return Err(Error::Invalid);
    }
    let entries = offsets.len() / 2;
    let columns = usize::try_from(columns).map_err(|_error| Error::Invalid)?;
    let mut table: Vec<[u8; 2]> = Vec::new();
    reserve(&mut table, entries + 1)?;
    for pair in offsets.chunks_exact(2) {
        table.push([pair[0], pair[1]]);
    }
    if table.iter().skip(columns).any(|entry| *entry != MISSING) {
        return Err(Error::Unsupported);
    }
    match edit {
        AxisEdit::Insert(index) => {
            let index = usize::try_from(index).map_err(|_error| Error::Invalid)?;
            if index >= entries {
                return Ok(None);
            }
            table.insert(index, MISSING);
            if entries > columns {
                // Spare capacity: keep the table's length.
                if table.pop() != Some(MISSING) {
                    return Err(Error::Unsupported);
                }
            }
        },
        AxisEdit::Remove(index) => {
            let index = usize::try_from(index).map_err(|_error| Error::Invalid)?;
            if index >= entries {
                return Ok(None);
            }
            if table.remove(index) != MISSING {
                // A stored cell is removed with its bytes (`remove_slot`).
                return Err(Error::Invalid);
            }
            if entries > columns {
                table.push(MISSING);
            }
        },
    }
    let mut output = Vec::new();
    reserve(&mut output, table.len() * 2)?;
    for entry in table {
        output.extend_from_slice(&entry);
    }
    Ok(Some(output))
}

/// The slots of one cell-offset table: `(start, end)` byte ranges of the
/// stored cells in `storage`, `None` for a missing cell. `unit` is the
/// offset unit in bytes (`4` for wide BNC offsets, else `1`).
fn slots(storage: &[u8], offsets: &[u8], unit: usize) -> Result<Vec<Option<(usize, usize)>>> {
    if offsets.len() % 2 != 0 {
        return Err(Error::Invalid);
    }
    let mut starts: Vec<Option<usize>> = Vec::new();
    reserve(&mut starts, offsets.len() / 2)?;
    for pair in offsets.chunks_exact(2) {
        let raw = [pair[0], pair[1]];
        starts.push((raw != MISSING).then(|| usize::from(u16::from_le_bytes(raw)) * unit));
    }
    let present: Vec<usize> = starts.iter().flatten().copied().collect();
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.last().is_some_and(|last| *last >= storage.len())
        || present.first().is_some_and(|first| *first != 0)
    {
        return Err(Error::Unsupported);
    }
    let mut output = Vec::new();
    reserve(&mut output, starts.len())?;
    let mut next = present.iter().skip(1);
    for start in starts {
        output.push(start.map(|start| (start, next.next().copied().unwrap_or(storage.len()))));
    }
    Ok(output)
}

/// The stored BNC cells of one tile row, as `(column, bytes)`.
pub fn bnc_cells(row: &[u8]) -> Result<Vec<(u32, &[u8])>> {
    let (Some(storage), Some(offsets)) = (single_bytes(row, 6)?, single_bytes(row, 7)?) else {
        return Ok(Vec::new());
    };
    let unit = if single_varint(row, 8)?.unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    let mut output = Vec::new();
    for (column, slot) in slots(storage, offsets, unit)?.into_iter().enumerate() {
        if let Some((start, end)) = slot {
            reserve(&mut output, 1)?;
            output.push((
                u32::try_from(column).map_err(|_error| Error::Invalid)?,
                storage.get(start..end).ok_or(Error::Invalid)?,
            ));
        }
    }
    Ok(output)
}

/// Rewrite one tile row so its BNC cell in `column` holds `cell`; every
/// other cell keeps its bytes (later cells move with the new length).
///
/// Returns `None` when the row stores no BNC cell in `column`.
pub fn replace_cell(row: &[u8], column: u32, cell: &[u8]) -> Result<Option<Vec<u8>>> {
    let (Some(storage), Some(offsets)) = (single_bytes(row, 6)?, single_bytes(row, 7)?) else {
        return Ok(None);
    };
    let unit = if single_varint(row, 8)?.unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    let column = usize::try_from(column).map_err(|_error| Error::Invalid)?;
    let cells = slots(storage, offsets, unit)?;
    if cells.get(column).copied().flatten().is_none() {
        return Ok(None);
    }
    let mut new_storage = Vec::new();
    reserve(&mut new_storage, storage.len() + cell.len())?;
    let mut new_offsets = Vec::new();
    reserve(&mut new_offsets, offsets.len())?;
    for (index, slot) in cells.into_iter().enumerate() {
        let Some((start, end)) = slot else {
            new_offsets.extend_from_slice(&MISSING);
            continue;
        };
        while new_storage.len() % unit != 0 {
            new_storage.push(0);
        }
        let offset =
            u16::try_from(new_storage.len() / unit).map_err(|_error| Error::Unsupported)?;
        if offset.to_le_bytes() == MISSING {
            return Err(Error::Unsupported);
        }
        new_offsets.extend_from_slice(&offset.to_le_bytes());
        if index == column {
            new_storage.extend_from_slice(cell);
        } else {
            new_storage.extend_from_slice(storage.get(start..end).ok_or(Error::Invalid)?);
        }
    }
    while new_storage.len() % unit != 0 {
        new_storage.push(0);
    }
    let rebuilt = rebuild(row, |field| match field.number {
        6 => bytes_field(6, &new_storage).map(Some),
        7 => bytes_field(7, &new_offsets).map(Some),
        _ => Ok(None),
    })?;
    Ok(Some(rebuilt))
}

/// Re-encode `storage`/`offsets` without the cell in `column`, keeping the
/// table's length (spare capacity) or shortening it (exact length).
fn remove_slot(
    storage: &[u8],
    offsets: &[u8],
    unit: usize,
    column: usize,
    columns: usize,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut cells = slots(storage, offsets, unit)?;
    if column >= cells.len() {
        return Err(Error::Invalid);
    }
    let spare = cells.len() > columns;
    cells.remove(column);
    if spare {
        cells.push(None);
    }
    let mut new_storage = Vec::new();
    reserve(&mut new_storage, storage.len())?;
    let mut new_offsets = Vec::new();
    reserve(&mut new_offsets, cells.len() * 2)?;
    for cell in cells {
        let Some((start, end)) = cell else {
            new_offsets.extend_from_slice(&MISSING);
            continue;
        };
        while new_storage.len() % unit != 0 {
            new_storage.push(0);
        }
        let offset =
            u16::try_from(new_storage.len() / unit).map_err(|_error| Error::Unsupported)?;
        if offset.to_le_bytes() == MISSING {
            return Err(Error::Unsupported);
        }
        new_offsets.extend_from_slice(&offset.to_le_bytes());
        new_storage.extend_from_slice(storage.get(start..end).ok_or(Error::Invalid)?);
    }
    while new_storage.len() % unit != 0 {
        new_storage.push(0);
    }
    Ok((new_storage, new_offsets))
}

/// Rewrite one tile row for a column edit. An insertion gives each
/// cell-offset table the new column's empty slot; a removal drops the
/// removed column's cell from the current (BNC) and pre-BNC storage (its
/// list references are released by the caller) and lowers the row's cell
/// count when a BNC cell went.
/// One row's pre-BNC cell storage and offset table after removing column
/// `column` of a `columns`-wide table, as the structural column removal
/// rewrites it (`None` when they are unchanged). Exposed for the host
/// editor's column removal (record 0370b).
///
/// # Errors
///
/// [`Error::Invalid`] for a malformed offset table.
pub fn pre_bnc_remove_column(
    storage: &[u8],
    offsets: &[u8],
    column: u32,
    columns: u32,
) -> Result<Option<(Vec<u8>, Vec<u8>)>> {
    if offsets.is_empty() {
        return Ok(None);
    }
    let index = usize::try_from(column).map_err(|_error| Error::Invalid)?;
    let columns_usize = usize::try_from(columns).map_err(|_error| Error::Invalid)?;
    let present = offsets
        .get(index * 2..index * 2 + 2)
        .is_some_and(|pair| pair != MISSING);
    if present {
        return remove_slot(storage, offsets, 1, index, columns_usize).map(Some);
    }
    Ok(splice_offsets(offsets, columns, AxisEdit::Remove(column))?
        .map(|offsets| (storage.to_vec(), offsets)))
}

fn row_columns(row: &[u8], columns: u32, edit: AxisEdit) -> Result<Option<Vec<u8>>> {
    let columns_usize = usize::try_from(columns).map_err(|_error| Error::Invalid)?;
    if let AxisEdit::Remove(index) = edit {
        let column = usize::try_from(index).map_err(|_error| Error::Invalid)?;
        let wide = single_varint(row, 8)?.unwrap_or(0) != 0;
        let mut replacements: Vec<(u32, Vec<u8>)> = Vec::new();
        let mut removed_bnc = false;
        for (storage_field, offsets_field, unit) in [(6, 7, if wide { 4 } else { 1 }), (3, 4, 1)] {
            let (Some(storage), Some(offsets)) = (
                single_bytes(row, storage_field)?,
                single_bytes(row, offsets_field)?,
            ) else {
                continue;
            };
            if offsets.is_empty() {
                continue;
            }
            let present = offsets
                .get(column * 2..column * 2 + 2)
                .is_some_and(|pair| pair != MISSING);
            let (new_storage, new_offsets) = if present {
                if storage_field == 6 {
                    removed_bnc = true;
                }
                remove_slot(storage, offsets, unit, column, columns_usize)?
            } else {
                match splice_offsets(offsets, columns, edit)? {
                    Some(new_offsets) => (storage.to_vec(), new_offsets),
                    None => continue,
                }
            };
            if new_storage == storage && new_offsets == offsets {
                continue;
            }
            replacements.push((storage_field, new_storage));
            replacements.push((offsets_field, new_offsets));
        }
        if replacements.is_empty() {
            return Ok(None);
        }
        // A BNC row the host's column shift rebuilds (no spare slot, or a
        // cell after the removed column) states its offset width (record
        // 0370b).
        let rebuilt_by_shift = match single_bytes(row, 7)? {
            Some(offsets) => {
                let slots = offsets.len() / 2;
                slots == columns_usize
                    || offsets
                        .chunks_exact(2)
                        .skip(column + 1)
                        .take(columns_usize.saturating_sub(column + 1))
                        .any(|pair| pair != MISSING)
            },
            None => false,
        };
        let state_width = rebuilt_by_shift
            && replacements.iter().any(|(number, _)| *number == 7)
            && single_varint(row, 8)?.is_none();
        let count = single_varint(row, 2)?.ok_or(Error::Invalid)?;
        let mut rebuilt = rebuild(row, |field| {
            if field.number == 2 && removed_bnc {
                return varint_field(2, count.checked_sub(1).ok_or(Error::Invalid)?).map(Some);
            }
            match replacements
                .iter()
                .find(|(number, _)| *number == field.number)
            {
                Some((number, payload)) => bytes_field(*number, payload).map(Some),
                None => Ok(None),
            }
        })?;
        if state_width {
            push_varint_field(&mut rebuilt, 8, 0)?;
        }
        return Ok(Some(rebuilt));
    }
    let mut changed = false;
    let rebuilt = rebuild(row, |field| {
        if field.number != 4 && field.number != 7 {
            return Ok(None);
        }
        if field.wire_type != 2 {
            return Err(Error::Invalid);
        }
        match splice_offsets(field.payload, columns, edit)? {
            Some(offsets) if offsets != field.payload => {
                changed = true;
                bytes_field(field.number, &offsets).map(Some)
            },
            _ => Ok(None),
        }
    })?;
    if !changed {
        return Ok(None);
    }
    // A rewritten row states its offset width, as the host editor's row
    // encoder does (record 0370); an absent `has_wide_offsets` reads as
    // narrow.
    let mut rebuilt = rebuilt;
    if single_varint(row, 8)?.is_none() {
        push_varint_field(&mut rebuilt, 8, 0)?;
    }
    Ok(Some(rebuilt))
}

/// Rewrite the rows of one storage tile for a column edit.
pub fn tile_columns(tile: &[u8], columns: u32, edit: AxisEdit) -> Result<Option<Vec<u8>>> {
    let original_rows = repeated_bytes(tile, 5)?;
    let numrows = single_varint(tile, 4)?.ok_or(Error::Invalid)?;
    let counts_rows = numrows == u64::try_from(original_rows.len()).map_err(|_e| Error::Invalid)?;
    let mut changed = false;
    let mut dropped = false;
    let mut kept = 0_u64;
    let mut maximum = None;
    let mut output = rebuild(tile, |field| {
        if field.number != 5 {
            return Ok(None);
        }
        let local = single_varint(field.payload, 1)?.ok_or(Error::Invalid)?;
        match row_columns(field.payload, columns, edit)? {
            // A row the removal leaves without cells goes, as the host's
            // cell removal drops it (record 0370b).
            Some(row)
                if single_varint(&row, 2)? == Some(0)
                    && single_varint(field.payload, 2)? != Some(0) =>
            {
                changed = true;
                dropped = true;
                Ok(Some(Vec::new()))
            },
            Some(row) => {
                changed = true;
                kept += 1;
                maximum = maximum.max(Some(local));
                bytes_field(5, &row).map(Some)
            },
            None => {
                kept += 1;
                maximum = maximum.max(Some(local));
                Ok(None)
            },
        }
    })?;
    if dropped {
        let numrows_after = if counts_rows {
            kept
        } else {
            maximum.map_or(0, |maximum| maximum + 1)
        };
        if numrows_after != numrows {
            output = rebuild(&output, |field| {
                if field.number == 4 {
                    return varint_field(4, numrows_after).map(Some);
                }
                Ok(None)
            })?;
        }
    }
    Ok(changed.then_some(output))
}

/// Rewrite the headers of one header bucket for an edit of its own axis.
///
/// `bucket` is the bucket's position (row buckets cover
/// [`HEADER_BUCKET_ROWS`] rows each; the column bucket is bucket `0`). A
/// header that would move into another bucket is refused. A removed index's
/// header must not hide its row or column. `released` gives, per source
/// index, how many of its stored cells the edit drops (a row removal drops
/// one cell from each column it stored a cell in, and the reverse).
pub fn header_bucket(
    payload: &[u8],
    bucket: u64,
    span: Option<u64>,
    edit: AxisEdit,
    released: &dyn Fn(u32) -> u64,
) -> Result<Option<Vec<u8>>> {
    let mut changed = false;
    let output = rebuild(payload, |field| {
        if field.number != 2 {
            return Ok(None);
        }
        let index = single_varint(field.payload, 1)?.ok_or(Error::Invalid)?;
        let index_u32 = u32::try_from(index).map_err(|_error| Error::Invalid)?;
        let release = released(index_u32);
        let cells_after = if release == 0 {
            None
        } else {
            let cells = single_varint(field.payload, 4)?.ok_or(Error::Invalid)?;
            Some(cells.checked_sub(release).ok_or(Error::Unsupported)?)
        };
        let Some(shifted) = edit.shift(index_u32)? else {
            if single_varint(field.payload, 3)?.unwrap_or(0) != 0 {
                return Err(Error::Unsupported);
            }
            changed = true;
            return Ok(Some(Vec::new()));
        };
        if let Some(span) = span
            && u64::from(shifted) / span != bucket
        {
            return Err(Error::Unsupported);
        }
        if u64::from(shifted) == index && cells_after.is_none() {
            return Ok(None);
        }
        changed = true;
        // A header left counting no cell and holding nothing else goes, as
        // the host's cell removal drops it (record 0370b).
        if cells_after == Some(0) && header_is_bare(field.payload)? {
            return Ok(Some(Vec::new()));
        }
        let header = rebuild(field.payload, |inner| match (inner.number, cells_after) {
            (1, _) => varint_field(1, u64::from(shifted)).map(Some),
            (4, Some(cells)) => varint_field(4, cells).map(Some),
            _ => Ok(None),
        })?;
        bytes_field(2, &header).map(Some)
    })?;
    Ok(changed.then_some(output))
}

/// Whether a `TST.HeaderStorageBucket.Header` holds only its index, a zero
/// size, a visible hiding state and its cell count.
fn header_is_bare(header: &[u8]) -> Result<bool> {
    for field in fields(header)? {
        match field.number {
            1 | 4 => {},
            2 => {
                if field.payload.iter().any(|byte| *byte != 0) {
                    return Ok(false);
                }
            },
            3 => {
                if single_varint(header, 3)? != Some(0) {
                    return Ok(false);
                }
            },
            _ => return Ok(false),
        }
    }
    Ok(true)
}

/// A `TSP.UUID` as its two halves.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Uuid {
    upper: u64,
    lower: u64,
}

fn decode_uuid(payload: &[u8]) -> Result<Uuid> {
    let mut lower = None;
    let mut upper = None;
    for field in fields(payload)? {
        let slot = match (field.number, field.wire_type) {
            (1, 0) => &mut lower,
            (2, 0) => &mut upper,
            _ => return Err(Error::Unsupported),
        };
        if slot.replace(varint(field.payload)?).is_some() {
            return Err(Error::Invalid);
        }
    }
    Ok(Uuid {
        lower: lower.ok_or(Error::Invalid)?,
        upper: upper.ok_or(Error::Invalid)?,
    })
}

fn encode_uuid(output: &mut Vec<u8>, number: u32, uuid: Uuid) -> Result<()> {
    let mut inner = Vec::new();
    push_varint_field(&mut inner, 1, uuid.lower)?;
    push_varint_field(&mut inner, 2, uuid.upper)?;
    push_bytes_field(output, number, &inner)
}

/// One axis of a `TST.ColumnRowUIDMapArchive`.
struct UidAxis {
    sorted: Vec<Uuid>,
    index_for_uid: Vec<u64>,
    uid_for_index: Vec<u64>,
}

impl UidAxis {
    fn read(payload: &[u8], first: u32) -> Result<Self> {
        let mut sorted = Vec::new();
        for record in repeated_bytes(payload, first)? {
            reserve(&mut sorted, 1)?;
            sorted.push(decode_uuid(record)?);
        }
        Ok(Self {
            sorted,
            index_for_uid: repeated_varints(payload, first + 1)?,
            uid_for_index: repeated_varints(payload, first + 2)?,
        })
    }

    fn write(&self, output: &mut Vec<u8>, first: u32) -> Result<()> {
        for uuid in &self.sorted {
            encode_uuid(output, first, *uuid)?;
        }
        for value in &self.index_for_uid {
            push_varint_field(output, first + 1, *value)?;
        }
        for value in &self.uid_for_index {
            push_varint_field(output, first + 2, *value)?;
        }
        Ok(())
    }

    fn check(&self, length: u32) -> Result<()> {
        let length = usize::try_from(length).map_err(|_error| Error::Invalid)?;
        if self.sorted.len() != length
            || self.index_for_uid.len() != length
            || self.uid_for_index.len() != length
        {
            return Err(Error::Unsupported);
        }
        Ok(())
    }

    /// The UUID at physical `index`.
    fn uuid_at(&self, index: u32) -> Result<Uuid> {
        let slot = self
            .uid_for_index
            .get(usize::try_from(index).map_err(|_error| Error::Invalid)?)
            .ok_or(Error::Invalid)?;
        self.sorted
            .get(usize::try_from(*slot).map_err(|_error| Error::Invalid)?)
            .copied()
            .ok_or(Error::Invalid)
    }

    fn apply(&mut self, edit: AxisEdit) -> Result<()> {
        match edit {
            AxisEdit::Insert(index) => {
                // A fresh identity above every existing one keeps the list
                // sorted, as the host editor allocates it.
                let upper = self
                    .sorted
                    .iter()
                    .map(|uuid| uuid.upper)
                    .max()
                    .unwrap_or(0)
                    .checked_add(1)
                    .ok_or(Error::Unsupported)?;
                let lower = self
                    .sorted
                    .iter()
                    .map(|uuid| uuid.lower)
                    .max()
                    .unwrap_or(0)
                    .checked_add(1)
                    .ok_or(Error::Unsupported)?;
                let slot = u64::try_from(self.sorted.len()).map_err(|_error| Error::Invalid)?;
                for value in &mut self.index_for_uid {
                    if *value >= u64::from(index) {
                        *value += 1;
                    }
                }
                reserve(&mut self.sorted, 1)?;
                self.sorted.push(Uuid { upper, lower });
                reserve(&mut self.index_for_uid, 1)?;
                self.index_for_uid.push(u64::from(index));
                reserve(&mut self.uid_for_index, 1)?;
                self.uid_for_index.insert(
                    usize::try_from(index).map_err(|_error| Error::Invalid)?,
                    slot,
                );
            },
            AxisEdit::Remove(index) => {
                let position = usize::try_from(index).map_err(|_error| Error::Invalid)?;
                let slot = self.uid_for_index.remove(position);
                let slot_index = usize::try_from(slot).map_err(|_error| Error::Invalid)?;
                if slot_index >= self.sorted.len() {
                    return Err(Error::Invalid);
                }
                self.sorted.remove(slot_index);
                self.index_for_uid.remove(slot_index);
                for value in &mut self.index_for_uid {
                    if *value > u64::from(index) {
                        *value -= 1;
                    }
                }
                for value in &mut self.uid_for_index {
                    if *value > slot {
                        *value -= 1;
                    }
                }
            },
        }
        Ok(())
    }
}

/// A decoded row/column UID map whose native encoding this module
/// reproduces exactly.
pub struct UidMap {
    columns: UidAxis,
    rows: UidAxis,
}

impl UidMap {
    /// Decode `payload`, refusing a map this module would not re-encode
    /// byte for byte or whose lengths disagree with the table.
    pub fn read(payload: &[u8], rows: u32, columns: u32) -> Result<Self> {
        if fields(payload)?
            .iter()
            .any(|field: &Field<'_>| !(1..=6).contains(&field.number))
        {
            return Err(Error::Unsupported);
        }
        let map = Self {
            columns: UidAxis::read(payload, 1)?,
            rows: UidAxis::read(payload, 4)?,
        };
        if map.encode()? != payload {
            return Err(Error::Unsupported);
        }
        map.columns.check(columns)?;
        map.rows.check(rows)?;
        Ok(map)
    }

    fn encode(&self) -> Result<Vec<u8>> {
        let mut output = Vec::new();
        self.columns.write(&mut output, 1)?;
        self.rows.write(&mut output, 4)?;
        Ok(output)
    }

    /// The UUID halves `(lower, upper)` of physical row or column `index`.
    pub fn uuid_at(&self, row: bool, index: u32) -> Result<(u64, u64)> {
        let uuid = if row {
            self.rows.uuid_at(index)?
        } else {
            self.columns.uuid_at(index)?
        };
        Ok((uuid.lower, uuid.upper))
    }

    /// Apply one edit and encode the result.
    /// The map for a table resized from `old_rows` × `old_columns` to
    /// `rows` × `columns`: trailing identities go, and new trailing ones
    /// are allocated above every existing one (the host editor's
    /// `resize_uid_axis`; record 0370e).
    ///
    /// # Errors
    ///
    /// As [`Self::rewrite`].
    pub fn resize(
        mut self,
        old_rows: u32,
        rows: u32,
        old_columns: u32,
        columns: u32,
    ) -> Result<Vec<u8>> {
        for (axis, old, new) in [
            (&mut self.rows, old_rows, rows),
            (&mut self.columns, old_columns, columns),
        ] {
            for index in (new..old).rev() {
                axis.apply(AxisEdit::Remove(index))?;
            }
            for index in old..new {
                axis.apply(AxisEdit::Insert(index))?;
            }
        }
        self.encode()
    }

    pub fn rewrite(mut self, row: bool, edit: AxisEdit) -> Result<Vec<u8>> {
        if row {
            self.rows.apply(edit)?;
        } else {
            self.columns.apply(edit)?;
        }
        self.encode()
    }
}

/// The new runs of one perpendicular stroke layer, or `None` when nothing
/// changes. A removal that would leave the layer without runs is refused.
fn stroke_runs(layer: &[u8], length: u32, edit: AxisEdit) -> Result<Option<Vec<u8>>> {
    let mut changed = false;
    let mut runs = 0_usize;
    let output = rebuild(layer, |field| {
        if field.number != 2 {
            return Ok(None);
        }
        let origin = single_varint(field.payload, 1)?.ok_or(Error::Invalid)?;
        let run_length = single_varint(field.payload, 2)?.ok_or(Error::Invalid)?;
        let start = u32::try_from(origin).map_err(|_error| Error::Unsupported)?;
        let run_length = u32::try_from(run_length).map_err(|_error| Error::Invalid)?;
        let end = start.checked_add(run_length).ok_or(Error::Invalid)?;
        if run_length == 0 || end > length {
            return Err(Error::Invalid);
        }
        let patch = |payload: &[u8], origin: u32, run_length: u32| -> Result<Vec<u8>> {
            rebuild(payload, |inner| match inner.number {
                1 => varint_field(1, u64::from(origin)).map(Some),
                2 => varint_field(2, u64::from(run_length)).map(Some),
                _ => Ok(None),
            })
        };
        let replacement = match edit {
            AxisEdit::Insert(index) if index <= start => Some(bytes_field(
                2,
                &patch(field.payload, start + 1, run_length)?,
            )?),
            AxisEdit::Insert(index) if index >= end => None,
            AxisEdit::Insert(index) => {
                let mut both = bytes_field(2, &patch(field.payload, start, index - start)?)?;
                let after = bytes_field(2, &patch(field.payload, index + 1, end - index)?)?;
                reserve(&mut both, after.len())?;
                both.extend_from_slice(&after);
                runs += 1;
                Some(both)
            },
            AxisEdit::Remove(index) if index < start => Some(bytes_field(
                2,
                &patch(field.payload, start - 1, run_length)?,
            )?),
            AxisEdit::Remove(index) if index >= end => None,
            AxisEdit::Remove(_) if run_length == 1 => Some(Vec::new()),
            AxisEdit::Remove(_) => Some(bytes_field(
                2,
                &patch(field.payload, start, run_length - 1)?,
            )?),
        };
        if replacement.as_ref().is_none_or(|bytes| !bytes.is_empty()) {
            runs += 1;
        }
        if replacement.is_some() {
            changed = true;
        }
        Ok(replacement)
    })?;
    if changed && runs == 0 {
        return Err(Error::Unsupported);
    }
    Ok(changed.then_some(output))
}

/// The stroke-sidecar rewrite of one axis edit: the sidecar's new payload
/// and each changed layer's new payload.
pub struct StrokePlan {
    pub sidecar: Vec<u8>,
    pub layers: Vec<(u64, Vec<u8>)>,
}

/// Plan the stroke-sidecar rewrite. `layer` returns a layer's payload.
///
/// Layers on the edited axis move with their index; perpendicular layers'
/// runs shift, split around an inserted index or shorten over a removed
/// one. A layer sitting on a removed index is refused (it would need its
/// object collected).
pub fn strokes<'p>(
    sidecar: &[u8],
    rows: u32,
    columns: u32,
    row_axis: bool,
    edit: AxisEdit,
    mut layer: impl FnMut(u64) -> Result<&'p [u8]>,
) -> Result<StrokePlan> {
    if single_varint(sidecar, 2)? != Some(u64::from(columns))
        || single_varint(sidecar, 3)? != Some(u64::from(rows))
    {
        return Err(Error::Unsupported);
    }
    let (rows_after, columns_after) = match (row_axis, edit) {
        (true, AxisEdit::Insert(_)) => (rows + 1, columns),
        (true, AxisEdit::Remove(_)) => (rows - 1, columns),
        (false, AxisEdit::Insert(_)) => (rows, columns + 1),
        (false, AxisEdit::Remove(_)) => (rows, columns - 1),
    };
    let mut layers = Vec::new();
    // Fields 4/5 hold column layers (fixed column, runs over rows); 6/7
    // hold row layers (fixed row, runs over columns).
    for number in 4..=7_u32 {
        let fixed_is_row = number >= 6;
        for reference in repeated_bytes(sidecar, number)? {
            let identifier = super::reference_identifier(reference)?;
            let payload = layer(identifier)?;
            let index = single_varint(payload, 1)?.ok_or(Error::Invalid)?;
            let index = u32::try_from(index).map_err(|_error| Error::Invalid)?;
            let rewritten = if fixed_is_row == row_axis {
                let Some(shifted) = edit.shift(index)? else {
                    return Err(Error::Unsupported);
                };
                (shifted != index)
                    .then(|| {
                        rebuild(payload, |inner| {
                            if inner.number == 1 {
                                return varint_field(1, u64::from(shifted)).map(Some);
                            }
                            Ok(None)
                        })
                    })
                    .transpose()?
            } else {
                let length = if fixed_is_row { columns } else { rows };
                stroke_runs(payload, length, edit)?
            };
            if let Some(rewritten) = rewritten {
                reserve(&mut layers, 1)?;
                layers.push((identifier, rewritten));
            }
        }
    }
    let sidecar = rebuild(sidecar, |field| match field.number {
        2 => varint_field(2, u64::from(columns_after)).map(Some),
        3 => varint_field(3, u64::from(rows_after)).map(Some),
        _ => Ok(None),
    })?;
    Ok(StrokePlan { sidecar, layers })
}

#[cfg(test)]
mod tests {
    use super::{AxisEdit, remove_slot, splice_offsets, stroke_runs};
    use crate::table_topology::wire::{bytes_field, push_varint_field};

    #[test]
    fn shift_moves_indices_at_or_after_the_edit() {
        assert_eq!(AxisEdit::Insert(2).shift(1).unwrap(), Some(1));
        assert_eq!(AxisEdit::Insert(2).shift(2).unwrap(), Some(3));
        assert_eq!(AxisEdit::Remove(2).shift(2).unwrap(), None);
        assert_eq!(AxisEdit::Remove(2).shift(3).unwrap(), Some(2));
    }

    #[test]
    fn offsets_keep_capacity_and_refuse_a_stored_removed_cell() {
        let offsets = [0, 0, 0xff, 0xff, 4, 0, 0xff, 0xff];
        let inserted = splice_offsets(&offsets, 3, AxisEdit::Insert(1))
            .unwrap()
            .unwrap();
        assert_eq!(inserted, [0, 0, 0xff, 0xff, 0xff, 0xff, 4, 0]);
        let removed = splice_offsets(&offsets, 3, AxisEdit::Remove(1))
            .unwrap()
            .unwrap();
        assert_eq!(removed, [0, 0, 4, 0, 0xff, 0xff, 0xff, 0xff]);
        assert!(splice_offsets(&offsets, 3, AxisEdit::Remove(0)).is_err());
        // A stored cell leaves with its bytes; the next cell moves up.
        let storage = [1, 2, 3, 4, 9, 9];
        let (storage, offsets) = remove_slot(&storage, &offsets, 1, 0, 3).unwrap();
        assert_eq!(storage, [9, 9]);
        assert_eq!(offsets, [0xff, 0xff, 0, 0, 0xff, 0xff, 0xff, 0xff]);
        let exact = [0, 0, 4, 0];
        let grown = splice_offsets(&exact, 2, AxisEdit::Insert(2));
        assert_eq!(grown.unwrap(), None);
        let grown = splice_offsets(&exact, 2, AxisEdit::Insert(0))
            .unwrap()
            .unwrap();
        assert_eq!(grown, [0xff, 0xff, 0, 0, 4, 0]);
    }

    fn run(origin: u64, length: u64) -> Vec<u8> {
        let mut payload = Vec::new();
        push_varint_field(&mut payload, 1, origin).unwrap();
        push_varint_field(&mut payload, 2, length).unwrap();
        bytes_field(2, &payload).unwrap()
    }

    #[test]
    fn perpendicular_runs_split_and_shorten() {
        let mut layer = Vec::new();
        push_varint_field(&mut layer, 1, 0).unwrap();
        layer.extend(run(1, 3));
        let split = stroke_runs(&layer, 5, AxisEdit::Insert(2))
            .unwrap()
            .unwrap();
        let mut expected = Vec::new();
        push_varint_field(&mut expected, 1, 0).unwrap();
        expected.extend(run(1, 1));
        expected.extend(run(3, 2));
        assert_eq!(split, expected);
        let shortened = stroke_runs(&layer, 5, AxisEdit::Remove(2))
            .unwrap()
            .unwrap();
        let mut expected = Vec::new();
        push_varint_field(&mut expected, 1, 0).unwrap();
        expected.extend(run(1, 2));
        assert_eq!(shortened, expected);
        assert_eq!(stroke_runs(&layer, 5, AxisEdit::Insert(4)).unwrap(), None);
    }
}
