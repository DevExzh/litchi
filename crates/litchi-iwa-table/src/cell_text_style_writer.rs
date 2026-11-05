//! Bounded, batch-capable BNC text-style-key publication.
//!
//! This module owns only the tile payload phase.  It does not resolve tables,
//! allocate native identifiers, update style lists, or publish a package.  A
//! leaf supplies the transaction ledger and the already-selected Numbers-wire
//! profile; this writer consumes both while rebuilding each affected row once.

use litchi_iwa_common::wire::{
    RawWireFields, RawWireLimits, WireFieldView, WireView, append_length_delimited_field,
    append_varint_field,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_numbers_wire::{BncCell, Limits as BncLimits};

use crate::cell_text_style::{Error, Limits};

const TILE_ROW_FIELD: u32 = 5;
const TILE_NUM_ROWS_FIELD: u32 = 4;
const ROW_INDEX_FIELD: u32 = 1;
const ROW_CELL_COUNT_FIELD: u32 = 2;
const ROW_STORAGE_FIELD: u32 = 6;
const ROW_OFFSETS_FIELD: u32 = 7;
const ROW_WIDE_OFFSETS_FIELD: u32 = 8;
const MISSING_OFFSET: [u8; 2] = [0xff, 0xff];
const MAX_STYLE_KEY_WIRE_GROWTH: usize = 8;

/// One cell whose text-style key is to be changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Update {
    /// Tile-local row index.
    pub row: u32,
    /// Table-local column index.
    pub column: u32,
    /// New text-style key. `None` clears the key on an existing cell; an empty
    /// slot with `None` remains absent.
    pub key: Option<u32>,
}

/// Placement policy for a newly materialized row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowPlacement {
    /// Append the new row after the existing fields.
    Append,
    /// Place it with the existing row fields, before the first larger row.
    FieldOrder,
}

/// Result of one bounded batch tile write.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Write {
    tile: Vec<u8>,
    previous: Vec<Option<Vec<u8>>>,
    changed: bool,
}

impl Write {
    /// Rewritten tile payload.
    #[must_use]
    pub fn tile(&self) -> &[u8] {
        &self.tile
    }

    /// Previous bytes aligned with the caller's update slice.
    #[must_use]
    pub fn previous(&self) -> &[Option<Vec<u8>>] {
        &self.previous
    }

    /// Whether any cell or row changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Consume the result and return the rewritten tile.
    #[must_use]
    pub fn into_tile(self) -> Vec<u8> {
        self.tile
    }
}

#[derive(Debug, Clone, Copy)]
struct IndexedUpdate {
    index: usize,
    update: Update,
}

#[derive(Debug, Clone)]
enum CellBytes<'a> {
    Borrowed(&'a [u8]),
    Owned(Vec<u8>),
}

impl CellBytes<'_> {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Borrowed(bytes) => bytes,
            Self::Owned(bytes) => bytes,
        }
    }
}

#[derive(Debug)]
struct RowRewrite {
    row: u32,
    bytes: Vec<u8>,
    changed: bool,
}

fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource(reason)
}

fn map_common(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind,
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => invalid("malformed tile wire payload"),
    }
}

fn map_bnc(error: litchi_numbers_wire::Error) -> Error {
    match error {
        litchi_numbers_wire::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes
                | litchi_numbers_wire::LimitKind::TailBytes => LimitKind::InputBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_numbers_wire::LimitKind::Fields => LimitKind::Fields,
                litchi_numbers_wire::LimitKind::Items => LimitKind::RewriteWork,
                _ => LimitKind::RewriteWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_numbers_wire::Error::Allocation { requested } => {
            Error::Allocation { amount: requested }
        },
        _ => invalid("malformed BNC cell payload"),
    }
}

fn map_bnc_limits(error: litchi_numbers_wire::limits::Error) -> Error {
    match error {
        litchi_numbers_wire::limits::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes
                | litchi_numbers_wire::LimitKind::TailBytes => LimitKind::InputBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_numbers_wire::LimitKind::Fields => LimitKind::Fields,
                litchi_numbers_wire::LimitKind::Items => LimitKind::RewriteWork,
                _ => LimitKind::RewriteWork,
            },
            observed: value as u64,
            maximum: maximum as u64,
        },
        litchi_numbers_wire::limits::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes
                | litchi_numbers_wire::LimitKind::TailBytes => LimitKind::InputBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_numbers_wire::LimitKind::Fields => LimitKind::Fields,
                litchi_numbers_wire::LimitKind::Items => LimitKind::RewriteWork,
                _ => LimitKind::RewriteWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => invalid("invalid BNC limits"),
    }
}

fn view<'a>(data: &'a [u8], transaction: &Limits) -> Result<WireView<'a>, Error> {
    let residual = transaction.residual_wire()?;
    let raw_limits = RawWireLimits::new(
        residual.max_input_bytes(),
        residual.max_fields(),
        residual.max_nesting(),
        residual.max_rewrite_work(),
    )
    .map_err(map_common)?;
    let mut raw = RawWireFields::new(data, raw_limits);
    while raw.next().map_err(map_common)?.is_some() {}
    let field_count = raw.fields();
    let raw_work = raw.work();
    transaction.charge_input(data.len())?;
    transaction.charge_fields(field_count)?;
    transaction.admit_work(raw_work)?;
    transaction.charge_input(data.len())?;
    transaction.charge_fields(field_count)?;
    transaction.admit_work(data.len())?;
    transaction.admit_allocations(field_count)?;
    let material = WireLimits::default()
        .with_input_bytes(data.len().max(1))
        .and_then(|limits| limits.with_fields(field_count.max(1)))
        .and_then(|limits| limits.with_nesting(WireLimits::MAX_NESTING))
        .and_then(|limits| {
            limits.with_rewrite_work(field_count.clamp(1, WireLimits::MAX_REWRITE_WORK))
        })
        .map_err(map_common)?;
    WireView::parse(data, material).map_err(map_common)
}

fn charge_view_scans(view: &WireView<'_>, transaction: &Limits, scans: usize) -> Result<(), Error> {
    transaction.admit_work(view.len().saturating_mul(scans))
}

fn decode_varint(field: WireFieldView<'_>) -> Result<u64, Error> {
    if field.wire_type() != 0 {
        return Err(invalid("tile field has the wrong wire type"));
    }
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
        .map_err(|_| invalid("invalid tile varint"))?;
    if length != field.payload().len() {
        return Err(invalid("tile varint has trailing bytes"));
    }
    Ok(value)
}

fn one_varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>, Error> {
    let mut found = None;
    for field in view.fields() {
        if field.number() == number {
            if found.replace(decode_varint(field)?).is_some() {
                return Err(invalid("duplicate tile scalar field"));
            }
        }
    }
    Ok(found)
}

fn one_bytes<'a>(view: &'a WireView<'a>, number: u32) -> Result<Option<&'a [u8]>, Error> {
    let mut found = None;
    for field in view.fields() {
        if field.number() == number {
            if field.wire_type() != 2 || found.replace(field.payload()).is_some() {
                return Err(invalid("duplicate tile bytes field"));
            }
        }
    }
    Ok(found)
}

fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn length_field_len(number: u32, payload: usize) -> Result<usize, Error> {
    let key = (u64::from(number) << 3) | 2;
    key.checked_add(0)
        .and_then(|_| {
            encoded_varint_len(key)
                .checked_add(encoded_varint_len(payload as u64))
                .and_then(|value| value.checked_add(payload))
        })
        .ok_or_else(|| invalid("tile output length overflow"))
}

fn append_raw(output: &mut Vec<u8>, raw: &[u8], limits: WireLimits) -> Result<(), Error> {
    let length = output
        .len()
        .checked_add(raw.len())
        .ok_or_else(|| invalid("tile output length overflow"))?;
    if length > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: length as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    output
        .try_reserve(raw.len())
        .map_err(|_| Error::Allocation { amount: raw.len() })?;
    output.extend_from_slice(raw);
    Ok(())
}

fn append_bytes(
    output: &mut Vec<u8>,
    number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<(), Error> {
    append_length_delimited_field(output, number, payload, limits).map_err(map_common)
}

fn append_varint(
    output: &mut Vec<u8>,
    number: u32,
    value: u64,
    _limits: WireLimits,
) -> Result<(), Error> {
    append_varint_field(output, number, value).map_err(map_common)
}

fn build_bnc_limits(wire: &BncLimits, columns: usize) -> Result<BncLimits, Error> {
    let items = wire.max_items().min(columns);
    if items == 0 {
        return Err(Error::Limit {
            kind: LimitKind::RewriteWork,
            observed: 1,
            maximum: 0,
        });
    }
    wire.with_items(items).map_err(map_bnc_limits)
}

fn intersect_bnc_limits(caller: &BncLimits, transaction: &Limits) -> Result<BncLimits, Error> {
    let residual = transaction.residual_wire()?;
    let caller_wire = caller.wire();
    let wire = caller_wire
        .with_input_bytes(
            caller_wire
                .max_input_bytes()
                .min(residual.max_input_bytes()),
        )
        .and_then(|value| {
            value.with_output_bytes(
                caller_wire
                    .max_output_bytes()
                    .min(residual.max_output_bytes()),
            )
        })
        .and_then(|value| value.with_fields(caller_wire.max_fields().min(residual.max_fields())))
        .and_then(|value| {
            value.with_rewrite_work(
                caller_wire
                    .max_rewrite_work()
                    .min(residual.max_rewrite_work()),
            )
        })
        .map_err(map_common)?;
    BncLimits::from_wire_limits(wire)
        .with_tail_bytes(caller.max_tail_bytes().min(residual.max_input_bytes()))
        .and_then(|value| value.with_items(caller.max_items().min(residual.max_rewrite_work())))
        .map_err(map_bnc_limits)
}

fn parse_slots<'a>(
    storage: &'a [u8],
    offsets: &[u8],
    unit: usize,
) -> Result<Vec<Option<&'a [u8]>>, Error> {
    if unit == 0 || offsets.len() % 2 != 0 {
        return Err(invalid("invalid tile cell offsets"));
    }
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(offsets.len() / 2)
        .map_err(|_| Error::Allocation {
            amount: offsets.len() / 2,
        })?;
    for pair in offsets.chunks_exact(2) {
        starts.push(if pair == MISSING_OFFSET {
            None
        } else {
            Some(
                usize::from(u16::from_le_bytes([pair[0], pair[1]]))
                    .checked_mul(unit)
                    .ok_or_else(|| invalid("cell offset overflow"))?,
            )
        });
    }
    let present_count = starts.iter().flatten().count();
    let mut present = Vec::new();
    present
        .try_reserve_exact(present_count)
        .map_err(|_| Error::Allocation {
            amount: present_count,
        })?;
    present.extend(starts.iter().flatten().copied());
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.first().is_some_and(|value| *value != 0)
        || present.last().is_some_and(|value| *value >= storage.len())
    {
        return Err(Error::Refused(
            crate::cell_text_style::Refusal::UnsupportedSource,
        ));
    }
    let mut next = present.iter().skip(1);
    let mut output = Vec::new();
    output
        .try_reserve_exact(starts.len())
        .map_err(|_| Error::Allocation {
            amount: starts.len(),
        })?;
    for start in starts {
        output.push(start.map(|start| (start, next.next().copied().unwrap_or(storage.len()))));
    }
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(output.len())
        .map_err(|_| Error::Allocation {
            amount: output.len(),
        })?;
    for slot in output {
        cells.push(slot.map(|(start, end)| &storage[start..end]));
    }
    Ok(cells)
}

fn encode_cells<'a>(
    cells: &[CellBytes<'a>],
    unit: usize,
    limits: &BncLimits,
    transaction: &Limits,
) -> Result<(Vec<u8>, Vec<u8>), Error> {
    if unit == 0 {
        return Err(invalid("invalid cell offset unit"));
    }
    let mut storage_len = 0usize;
    for cell in cells {
        let bytes = cell.as_bytes();
        if bytes.is_empty() {
            continue;
        }
        let padding = (unit - (storage_len % unit)) % unit;
        storage_len = storage_len
            .checked_add(padding)
            .and_then(|value| value.checked_add(bytes.len()))
            .ok_or_else(|| invalid("cell storage length overflow"))?;
    }
    storage_len = storage_len
        .checked_add((unit - (storage_len % unit)) % unit)
        .ok_or_else(|| invalid("cell storage padding overflow"))?;
    let offset_len = cells
        .len()
        .checked_mul(2)
        .ok_or_else(|| invalid("cell offset length overflow"))?;
    let combined = storage_len
        .checked_add(offset_len)
        .ok_or_else(|| invalid("cell output length overflow"))?;
    limits
        .with_output_bytes(combined.max(1))
        .map_err(map_bnc_limits)?;
    transaction.admit_work(combined)?;
    transaction.admit_output(combined)?;
    transaction.admit_copies(combined)?;
    transaction.admit_allocations(2)?;
    let mut storage = Vec::new();
    storage
        .try_reserve_exact(storage_len)
        .map_err(|_| Error::Allocation {
            amount: storage_len,
        })?;
    let mut offsets = Vec::new();
    offsets
        .try_reserve_exact(offset_len)
        .map_err(|_| Error::Allocation { amount: offset_len })?;
    for cell in cells {
        let bytes = cell.as_bytes();
        if bytes.is_empty() {
            offsets.extend_from_slice(&MISSING_OFFSET);
            continue;
        }
        while storage.len() % unit != 0 {
            storage.push(0);
        }
        let offset = u16::try_from(storage.len() / unit)
            .map_err(|_| Error::Refused(crate::cell_text_style::Refusal::UnsupportedSource))?;
        if offset.to_le_bytes() == MISSING_OFFSET {
            return Err(Error::Refused(
                crate::cell_text_style::Refusal::UnsupportedSource,
            ));
        }
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(bytes);
    }
    while storage.len() % unit != 0 {
        storage.push(0);
    }
    Ok((storage, offsets))
}

fn rewrite_row(
    row: &[u8],
    updates: &[IndexedUpdate],
    columns: usize,
    bnc_limits: &BncLimits,
    transaction: &Limits,
    previous: &mut [Option<Vec<u8>>],
) -> Result<RowRewrite, Error> {
    let wire = bnc_limits.wire();
    let view = view(row, transaction)?;
    // Five schema lookups plus the sizing and emission walks below.  Admit
    // every borrowed-view pass before the first lookup starts.
    charge_view_scans(&view, transaction, 7)?;
    let row_number =
        one_varint(&view, ROW_INDEX_FIELD)?.ok_or_else(|| invalid("tile row index is missing"))?;
    let row_number = u32::try_from(row_number).map_err(|_| invalid("tile row index overflow"))?;
    let count = one_varint(&view, ROW_CELL_COUNT_FIELD)?
        .ok_or_else(|| invalid("tile row cell count is missing"))?;
    let storage = one_bytes(&view, ROW_STORAGE_FIELD)?.ok_or(Error::Refused(
        crate::cell_text_style::Refusal::UnsupportedSource,
    ))?;
    let offsets = one_bytes(&view, ROW_OFFSETS_FIELD)?.ok_or(Error::Refused(
        crate::cell_text_style::Refusal::UnsupportedSource,
    ))?;
    let unit = if one_varint(&view, ROW_WIDE_OFFSETS_FIELD)?.unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    transaction.admit_allocations(5)?;
    let source_slots = parse_slots(storage, offsets, unit)?;
    let mut slots = Vec::new();
    slots
        .try_reserve_exact(source_slots.len())
        .map_err(|_| Error::Allocation {
            amount: source_slots.len(),
        })?;
    for cell in source_slots {
        slots.push(match cell {
            Some(bytes) => CellBytes::Borrowed(bytes),
            None => CellBytes::Owned(Vec::new()),
        });
    }
    let actual_count = slots
        .iter()
        .filter(|cell| !cell.as_bytes().is_empty())
        .count() as u64;
    if count != actual_count {
        return Err(invalid("tile row cell count does not match offsets"));
    }
    if slots.len() < columns {
        let additional = columns - slots.len();
        transaction.admit_allocations(additional)?;
        slots
            .try_reserve_exact(additional)
            .map_err(|_| Error::Allocation { amount: additional })?;
        slots.resize_with(columns, || CellBytes::Owned(Vec::new()));
    }
    let mut changed = false;
    for indexed in updates {
        let column =
            usize::try_from(indexed.update.column).map_err(|_| invalid("tile column overflow"))?;
        if column >= slots.len() {
            return Err(invalid("tile column is outside row offsets"));
        }
        let old = match &slots[column] {
            CellBytes::Borrowed(bytes) if !bytes.is_empty() => Some(*bytes),
            CellBytes::Owned(bytes) if !bytes.is_empty() => Some(bytes.as_slice()),
            _ => None,
        };
        if let Some(old) = old {
            transaction.admit_work(old.len())?;
            transaction.admit_copies(old.len())?;
            transaction.admit_allocations(litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS)?;
            let mut cell = BncCell::parse(old, bnc_limits).map_err(map_bnc)?;
            let previous_key = cell.text_style_identifier();
            if previous_key == indexed.update.key {
                transaction.admit_copies(old.len())?;
                transaction.admit_allocations(1)?;
                previous[indexed.index] = Some(old.to_vec());
                continue;
            }
            transaction.admit_copies(old.len())?;
            transaction.admit_allocations(1)?;
            previous[indexed.index] = Some(old.to_vec());
            cell.set_text_style_identifier(indexed.update.key);
            transaction.admit_allocations(1)?;
            transaction.admit_copies(old.len().saturating_add(MAX_STYLE_KEY_WIRE_GROWTH))?;
            transaction.admit_output(old.len().saturating_add(MAX_STYLE_KEY_WIRE_GROWTH))?;
            let encoded = cell.encode(bnc_limits).map_err(map_bnc)?;
            slots[column] = CellBytes::Owned(encoded);
            changed = true;
        } else if let Some(key) = indexed.update.key {
            let mut cell = BncCell::minimal();
            cell.set_text_style_identifier(Some(key));
            transaction.admit_allocations(1)?;
            transaction.admit_copies(MAX_STYLE_KEY_WIRE_GROWTH)?;
            transaction.admit_output(MAX_STYLE_KEY_WIRE_GROWTH)?;
            let encoded = cell.encode(bnc_limits).map_err(map_bnc)?;
            slots[column] = CellBytes::Owned(encoded);
            changed = true;
        }
    }
    if !changed {
        transaction.admit_copies(row.len())?;
        transaction.admit_output(row.len())?;
        transaction.admit_allocations(1)?;
        return Ok(RowRewrite {
            row: row_number,
            bytes: row.to_vec(),
            changed: false,
        });
    }
    let occupied = slots.iter().filter(|cell| match cell {
        CellBytes::Borrowed(bytes) => !bytes.is_empty(),
        CellBytes::Owned(bytes) => !bytes.is_empty(),
    });
    let new_count = occupied.count();
    let (new_storage, new_offsets) = encode_cells(&slots, unit, bnc_limits, transaction)?;
    let count_bytes = length_field_len(ROW_CELL_COUNT_FIELD, encoded_varint_len(new_count as u64))?;
    let storage_bytes = length_field_len(ROW_STORAGE_FIELD, new_storage.len())?;
    let offsets_bytes = length_field_len(ROW_OFFSETS_FIELD, new_offsets.len())?;
    let replacement_bytes = count_bytes
        .checked_add(storage_bytes)
        .and_then(|value| value.checked_add(offsets_bytes))
        .ok_or_else(|| invalid("row replacement length overflow"))?;
    transaction.admit_output(replacement_bytes)?;
    transaction.admit_copies(replacement_bytes)?;
    transaction.admit_allocations(8)?;
    let mut replacements: Vec<(u32, Vec<u8>)> = Vec::new();
    replacements
        .try_reserve_exact(3)
        .map_err(|_| Error::Allocation { amount: 3 })?;
    let mut count_field = Vec::new();
    append_varint(
        &mut count_field,
        ROW_CELL_COUNT_FIELD,
        new_count as u64,
        wire,
    )?;
    replacements.push((ROW_CELL_COUNT_FIELD, count_field));
    let mut storage_field = Vec::new();
    append_bytes(&mut storage_field, ROW_STORAGE_FIELD, &new_storage, wire)?;
    replacements.push((ROW_STORAGE_FIELD, storage_field));
    let mut offsets_field = Vec::new();
    append_bytes(&mut offsets_field, ROW_OFFSETS_FIELD, &new_offsets, wire)?;
    replacements.push((ROW_OFFSETS_FIELD, offsets_field));
    let field_lookup_work = view
        .len()
        .saturating_mul(replacements.len().saturating_add(2));
    transaction.admit_work(field_lookup_work)?;
    let output_len = view
        .fields()
        .map(|field| {
            replacements
                .iter()
                .find(|(number, _)| *number == field.number())
                .map_or(field.raw().len(), |(_, bytes)| bytes.len())
        })
        .try_fold(0usize, |total, length| total.checked_add(length))
        .ok_or_else(|| invalid("row output length overflow"))?;
    transaction.admit_work(output_len)?;
    transaction.admit_output(output_len)?;
    transaction.admit_copies(output_len)?;
    transaction.admit_allocations(1)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation { amount: output_len })?;
    for field in view.fields() {
        if let Some((_, replacement)) = replacements
            .iter()
            .find(|(number, _)| *number == field.number())
        {
            output.extend_from_slice(replacement);
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if output.len() != output_len {
        return Err(invalid("row output length mismatch"));
    }
    Ok(RowRewrite {
        row: row_number,
        bytes: output,
        changed: true,
    })
}

/// Rewrite all requested style keys in one tile. The tile is scanned once and
/// each affected row is rebuilt once; unchanged fields and cells remain raw.
pub fn set_text_style_keys(
    tile: &[u8],
    columns: u32,
    updates: &[Update],
    placement: RowPlacement,
    transaction: &Limits,
    bnc_limits: &BncLimits,
) -> Result<Write, Error> {
    if updates.is_empty() {
        transaction.admit_allocations(1)?;
        transaction.admit_copies(tile.len())?;
        transaction.admit_output(tile.len())?;
        return Ok(Write {
            tile: tile.to_vec(),
            previous: Vec::new(),
            changed: false,
        });
    }
    let columns = usize::try_from(columns).map_err(|_| invalid("table column count overflow"))?;
    if columns == 0 {
        return Err(invalid("table has no columns"));
    }
    transaction.admit_allocations(updates.len())?;
    let mut indexed = Vec::new();
    indexed
        .try_reserve_exact(updates.len())
        .map_err(|_| Error::Allocation {
            amount: updates.len(),
        })?;
    transaction.admit_work(updates.len())?;
    for (index, update) in updates.iter().copied().enumerate() {
        if usize::try_from(update.column).map_err(|_| invalid("tile column overflow"))? >= columns {
            return Err(Error::Refused(
                crate::cell_text_style::Refusal::OutsideTable,
            ));
        }
        if update.key == Some(0) {
            return Err(invalid("text-style key zero is invalid"));
        }
        indexed.push(IndexedUpdate { index, update });
    }
    indexed.sort_by_key(|item| (item.update.row, item.update.column));
    transaction.admit_work(updates.len().saturating_mul(usize::BITS as usize))?;
    transaction.admit_work(updates.len())?;
    if indexed.windows(2).any(|pair| {
        pair[0].update.row == pair[1].update.row && pair[0].update.column == pair[1].update.column
    }) {
        return Err(invalid("duplicate tile cell update"));
    }
    let bnc_limits = intersect_bnc_limits(bnc_limits, transaction)?;
    let wire = bnc_limits.wire();
    let root = view(tile, transaction)?;
    let bnc_limits = build_bnc_limits(&bnc_limits, columns.max(updates.len()))?;
    // The root is walked for row counting, row rewrites, and capacity
    // discovery before the changed/no-op decision is known.
    charge_view_scans(&root, transaction, 3)?;
    transaction.admit_allocations(1)?;
    let mut previous = Vec::new();
    previous
        .try_reserve_exact(updates.len())
        .map_err(|_| Error::Allocation {
            amount: updates.len(),
        })?;
    previous.resize_with(updates.len(), || None);
    let row_count = root
        .fields()
        .filter(|field| field.number() == TILE_ROW_FIELD)
        .count();
    transaction.admit_allocations(row_count.saturating_add(updates.len()))?;
    let mut rewrites = Vec::new();
    rewrites
        .try_reserve_exact(row_count.saturating_add(updates.len()))
        .map_err(|_| Error::Allocation {
            amount: row_count.saturating_add(updates.len()),
        })?;
    let mut existing_rows = Vec::new();
    existing_rows
        .try_reserve_exact(row_count)
        .map_err(|_| Error::Allocation { amount: row_count })?;
    for field in root
        .fields()
        .filter(|field| field.number() == TILE_ROW_FIELD)
    {
        let row = field.payload();
        let row_view = view(row, transaction)?;
        charge_view_scans(&row_view, transaction, 1)?;
        let row_number = one_varint(&row_view, ROW_INDEX_FIELD)?
            .ok_or_else(|| invalid("tile row index is missing"))?;
        let row_number =
            u32::try_from(row_number).map_err(|_| invalid("tile row index overflow"))?;
        existing_rows.push(row_number);
        let start = indexed.partition_point(|item| item.update.row < row_number);
        let end = indexed.partition_point(|item| item.update.row <= row_number);
        if start != end {
            let rewritten = rewrite_row(
                row,
                &indexed[start..end],
                columns,
                &bnc_limits,
                transaction,
                &mut previous,
            )?;
            rewrites.push(rewritten);
        }
    }
    transaction.admit_work(existing_rows.len())?;
    if existing_rows.windows(2).any(|pair| pair[0] >= pair[1]) {
        return Err(invalid("tile rows are not strictly ordered"));
    }
    let mut missing: Vec<Vec<IndexedUpdate>> = Vec::new();
    transaction.admit_allocations(1)?;
    missing
        .try_reserve_exact(updates.len())
        .map_err(|_| Error::Allocation {
            amount: updates.len(),
        })?;
    let mut index = 0usize;
    while index < indexed.len() {
        let row = indexed[index].update.row;
        transaction.admit_work(indexed.len())?;
        let end = indexed[index..]
            .iter()
            .position(|item| item.update.row != row)
            .map_or(indexed.len(), |offset| index + offset);
        transaction.admit_work(existing_rows.len())?;
        if !existing_rows.contains(&row) {
            transaction.admit_allocations(1)?;
            let mut group = Vec::new();
            group
                .try_reserve_exact(end - index)
                .map_err(|_| Error::Allocation {
                    amount: end - index,
                })?;
            group.extend_from_slice(&indexed[index..end]);
            missing.push(group);
        }
        index = end;
    }
    let mut capacity = columns;
    for field in root
        .fields()
        .filter(|field| field.number() == TILE_ROW_FIELD)
    {
        let row_view = view(field.payload(), transaction)?;
        charge_view_scans(&row_view, transaction, 1)?;
        let offsets = one_bytes(&row_view, ROW_OFFSETS_FIELD)?
            .ok_or_else(|| invalid("tile row offsets are missing"))?;
        capacity = capacity.max(offsets.len() / 2);
    }
    for chunk in missing {
        let row_number = chunk[0].update.row;
        if chunk.iter().all(|item| item.update.key.is_none()) {
            continue;
        }
        transaction.admit_allocations(capacity)?;
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(capacity)
            .map_err(|_| Error::Allocation { amount: capacity })?;
        cells.resize_with(capacity, || CellBytes::Owned(Vec::new()));
        for item in &chunk {
            let mut cell = BncCell::minimal();
            cell.set_text_style_identifier(item.update.key);
            transaction.admit_allocations(1)?;
            transaction.admit_copies(MAX_STYLE_KEY_WIRE_GROWTH)?;
            transaction.admit_output(MAX_STYLE_KEY_WIRE_GROWTH)?;
            let encoded = cell.encode(&bnc_limits).map_err(map_bnc)?;
            cells[item.update.column as usize] = CellBytes::Owned(encoded);
        }
        let (storage, offsets) = encode_cells(&cells, 1, &bnc_limits, transaction)?;
        let new_count = chunk
            .iter()
            .filter(|item| item.update.key.is_some())
            .count();
        let row_len = encoded_varint_len(u64::from(ROW_INDEX_FIELD) << 3)
            .checked_add(encoded_varint_len(u64::from(row_number)))
            .and_then(|value| {
                value.checked_add(encoded_varint_len(u64::from(ROW_CELL_COUNT_FIELD) << 3))
            })
            .and_then(|value| value.checked_add(encoded_varint_len(new_count as u64)))
            .and_then(|value| value.checked_add(length_field_len(3, 0).ok()?))
            .and_then(|value| value.checked_add(length_field_len(4, 0).ok()?))
            .and_then(|value| value.checked_add(encoded_varint_len(u64::from(5_u32) << 3)))
            .and_then(|value| value.checked_add(encoded_varint_len(5)))
            .and_then(|value| {
                value.checked_add(length_field_len(ROW_STORAGE_FIELD, storage.len()).ok()?)
            })
            .and_then(|value| {
                value.checked_add(length_field_len(ROW_OFFSETS_FIELD, offsets.len()).ok()?)
            })
            .and_then(|value| {
                value.checked_add(encoded_varint_len(u64::from(ROW_WIDE_OFFSETS_FIELD) << 3))
            })
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| invalid("new row length overflow"))?;
        transaction.admit_output(row_len)?;
        transaction.admit_copies(row_len)?;
        transaction.admit_allocations(1)?;
        let mut row = Vec::new();
        append_varint(&mut row, ROW_INDEX_FIELD, u64::from(row_number), wire)?;
        append_varint(&mut row, ROW_CELL_COUNT_FIELD, new_count as u64, wire)?;
        append_bytes(&mut row, 3, &[], wire)?;
        append_bytes(&mut row, 4, &[], wire)?;
        append_varint(&mut row, 5, 5, wire)?;
        append_bytes(&mut row, ROW_STORAGE_FIELD, &storage, wire)?;
        append_bytes(&mut row, ROW_OFFSETS_FIELD, &offsets, wire)?;
        append_varint(&mut row, ROW_WIDE_OFFSETS_FIELD, 0, wire)?;
        rewrites.push(RowRewrite {
            row: row_number,
            bytes: row,
            changed: true,
        });
        for item in chunk {
            previous[item.index] = None;
        }
    }
    transaction.admit_work(rewrites.len())?;
    let mut changed = false;
    for rewrite in &rewrites {
        changed |= rewrite.changed;
    }
    if !changed {
        transaction.admit_allocations(1)?;
        transaction.admit_copies(tile.len())?;
        transaction.admit_output(tile.len())?;
        return Ok(Write {
            tile: tile.to_vec(),
            previous,
            changed: false,
        });
    }
    // The remaining root passes are: old row-count lookup, output sizing, and
    // final emission.  Admit all three before any of them starts.
    charge_view_scans(&root, transaction, 3)?;
    let old_numrows = one_varint(&root, TILE_NUM_ROWS_FIELD)?
        .ok_or_else(|| invalid("tile row count is missing"))?;
    let row_lookup_work = rewrites
        .len()
        .saturating_mul(existing_rows.len())
        .saturating_mul(2);
    transaction.admit_work(row_lookup_work)?;
    transaction.admit_work(
        rewrites
            .len()
            .saturating_add(existing_rows.len())
            .saturating_mul(2),
    )?;
    let inserted_rows = rewrites
        .iter()
        .map(|rewrite| rewrite.row)
        .filter(|row| !existing_rows.contains(row))
        .count();
    let counts_rows = old_numrows == existing_rows.len() as u64;
    let new_numrows = if counts_rows {
        old_numrows
            .checked_add(inserted_rows as u64)
            .ok_or_else(|| invalid("tile row count overflow"))?
    } else {
        old_numrows.max(
            rewrites
                .iter()
                .map(|rewrite| rewrite.row)
                .filter(|row| !existing_rows.contains(row))
                .filter_map(|row| u64::from(row).checked_add(1))
                .max()
                .unwrap_or(old_numrows),
        )
    };
    let mut output_len = 0usize;
    transaction.admit_work(
        (root.len().saturating_add(1)).saturating_mul(
            rewrites
                .len()
                .saturating_mul(4)
                .saturating_add(existing_rows.len().saturating_mul(2)),
        ),
    )?;
    for field in root.fields() {
        let length = if field.number() == TILE_ROW_FIELD {
            let row_view = view(field.payload(), transaction)?;
            charge_view_scans(&row_view, transaction, 1)?;
            let row_number = one_varint(&row_view, ROW_INDEX_FIELD)?.unwrap_or(u64::MAX);
            if let Some(rewrite) = rewrites
                .iter()
                .find(|rewrite| rewrite.row == row_number as u32)
            {
                length_field_len(TILE_ROW_FIELD, rewrite.bytes.len())?
            } else {
                field.raw().len()
            }
        } else if field.number() == TILE_NUM_ROWS_FIELD && inserted_rows != 0 {
            encoded_varint_len(u64::from(TILE_NUM_ROWS_FIELD) << 3)
                .checked_add(encoded_varint_len(new_numrows))
                .ok_or_else(|| invalid("tile output length overflow"))?
        } else {
            field.raw().len()
        };
        output_len = output_len
            .checked_add(length)
            .ok_or_else(|| invalid("tile output length overflow"))?;
    }
    for rewrite in &rewrites {
        if !existing_rows.contains(&rewrite.row) {
            output_len = output_len
                .checked_add(length_field_len(TILE_ROW_FIELD, rewrite.bytes.len())?)
                .ok_or_else(|| invalid("tile output length overflow"))?;
        }
    }
    transaction.admit_output(output_len)?;
    transaction.admit_copies(output_len)?;
    transaction.admit_work(output_len)?;
    transaction.admit_allocations(1)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation { amount: output_len })?;
    transaction.admit_allocations(1)?;
    let mut inserted_row_set = Vec::new();
    inserted_row_set
        .try_reserve_exact(rewrites.len())
        .map_err(|_| Error::Allocation {
            amount: rewrites.len(),
        })?;
    for field in root.fields() {
        if matches!(placement, RowPlacement::FieldOrder) && field.number() > TILE_ROW_FIELD {
            transaction.admit_allocations(1)?;
            transaction.admit_work(rewrites.len().saturating_add(existing_rows.len()))?;
            let mut rows_to_insert = Vec::new();
            rows_to_insert
                .try_reserve_exact(rewrites.len())
                .map_err(|_| Error::Allocation {
                    amount: rewrites.len(),
                })?;
            rows_to_insert.extend(
                rewrites
                    .iter()
                    .filter(|rewrite| {
                        !existing_rows.contains(&rewrite.row)
                            && !inserted_row_set.contains(&rewrite.row)
                    })
                    .map(|rewrite| rewrite.row),
            );
            for row in rows_to_insert {
                let rewrite = rewrites
                    .iter()
                    .find(|rewrite| rewrite.row == row)
                    .ok_or_else(|| invalid("missing inserted row"))?;
                append_bytes(&mut output, TILE_ROW_FIELD, &rewrite.bytes, wire)?;
                inserted_row_set.push(row);
            }
        }
        if field.number() == TILE_ROW_FIELD {
            let row_view = view(field.payload(), transaction)?;
            charge_view_scans(&row_view, transaction, 1)?;
            let row_number = u32::try_from(one_varint(&row_view, ROW_INDEX_FIELD)?.unwrap_or(0))
                .map_err(|_| invalid("tile row index overflow"))?;
            if matches!(placement, RowPlacement::FieldOrder) {
                transaction.admit_allocations(1)?;
                transaction.admit_work(rewrites.len().saturating_add(existing_rows.len()))?;
                let mut rows_to_insert = Vec::new();
                rows_to_insert
                    .try_reserve_exact(rewrites.len())
                    .map_err(|_| Error::Allocation {
                        amount: rewrites.len(),
                    })?;
                rows_to_insert.extend(
                    rewrites
                        .iter()
                        .filter(|rewrite| {
                            !existing_rows.contains(&rewrite.row)
                                && rewrite.row < row_number
                                && !inserted_row_set.contains(&rewrite.row)
                        })
                        .map(|rewrite| rewrite.row),
                );
                for row in rows_to_insert {
                    let rewrite = rewrites
                        .iter()
                        .find(|rewrite| rewrite.row == row)
                        .ok_or_else(|| invalid("missing inserted row"))?;
                    append_bytes(&mut output, TILE_ROW_FIELD, &rewrite.bytes, wire)?;
                    inserted_row_set.push(row);
                }
            }
            if let Some(rewrite) = rewrites.iter().find(|rewrite| rewrite.row == row_number) {
                append_bytes(&mut output, TILE_ROW_FIELD, &rewrite.bytes, wire)?;
            } else {
                append_raw(&mut output, field.raw(), wire)?;
            }
        } else if field.number() == TILE_NUM_ROWS_FIELD && inserted_rows != 0 {
            append_varint(&mut output, TILE_NUM_ROWS_FIELD, new_numrows, wire)?;
        } else {
            append_raw(&mut output, field.raw(), wire)?;
        }
    }
    transaction.admit_allocations(1)?;
    transaction.admit_work(rewrites.len().saturating_add(existing_rows.len()))?;
    let mut rows_to_insert = Vec::new();
    rows_to_insert
        .try_reserve_exact(rewrites.len())
        .map_err(|_| Error::Allocation {
            amount: rewrites.len(),
        })?;
    rows_to_insert.extend(
        rewrites
            .iter()
            .filter(|rewrite| {
                !existing_rows.contains(&rewrite.row) && !inserted_row_set.contains(&rewrite.row)
            })
            .map(|rewrite| rewrite.row),
    );
    for row in rows_to_insert {
        let rewrite = rewrites
            .iter()
            .find(|rewrite| rewrite.row == row)
            .ok_or_else(|| invalid("missing inserted row"))?;
        append_bytes(&mut output, TILE_ROW_FIELD, &rewrite.bytes, wire)?;
        inserted_row_set.push(row);
    }
    if output.len() != output_len {
        return Err(invalid("tile output length mismatch"));
    }
    Ok(Write {
        tile: output,
        previous,
        changed: true,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{ResolveError, ResolvedObject, TableSource};
    use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    use litchi_numbers_wire::BncCell;

    struct Source {
        catalog: SourceCatalog,
    }

    impl TableSource for Source {
        fn catalog(&self) -> &ComponentCatalog {
            self.catalog.components()
        }

        fn physical(&self) -> Option<&SourceCatalog> {
            Some(&self.catalog)
        }

        fn resolve_object(
            &self,
            identifier: u64,
        ) -> Result<Option<ResolvedObject<'_>>, ResolveError> {
            for (component_index, component) in self.catalog.components().iter().enumerate() {
                if let Some((object_index, object)) = component
                    .archive()
                    .objects
                    .iter()
                    .enumerate()
                    .find(|(_, object)| object.archive_info.identifier == Some(identifier))
                {
                    return Ok(Some(ResolvedObject {
                        messages: &object.messages,
                        component_index,
                        object_index,
                    }));
                }
            }
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            self.catalog.source_bytes()
        }

        fn archive_limits(&self) -> litchi_iwa_archive::Limits {
            litchi_iwa_archive::Limits::default()
        }

        fn max_objects(&self) -> usize {
            1_000_000
        }

        fn max_references(&self) -> usize {
            1_000_000
        }
    }

    fn transaction_limits() -> Limits {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/numbers/basic.numbers"),
        )
        .expect("basic Numbers fixture");
        let source = Source {
            catalog: SourceCatalog::from_bytes(&bytes, litchi_iwa_archive::Limits::default())
                .expect("source catalog"),
        };
        Limits::from_source(&source).expect("transaction limits")
    }

    fn field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, number, value).expect("varint field");
        output
    }

    fn bytes(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_length_delimited_field(&mut output, number, payload, WireLimits::DEFAULT)
            .expect("bytes field");
        output
    }

    fn row(row: u32, cells: &[Option<&[u8]>], unknown: &[u8]) -> Vec<u8> {
        let mut storage = Vec::new();
        let mut offsets = Vec::new();
        let mut count = 0u64;
        for cell in cells {
            let Some(cell) = cell else {
                offsets.extend_from_slice(&MISSING_OFFSET);
                continue;
            };
            offsets.extend_from_slice(&(storage.len() as u16).to_le_bytes());
            storage.extend_from_slice(cell);
            count += 1;
        }
        let mut output = Vec::new();
        output.extend_from_slice(&field(ROW_INDEX_FIELD, u64::from(row)));
        output.extend_from_slice(&field(ROW_CELL_COUNT_FIELD, count));
        output.extend_from_slice(&bytes(3, &[]));
        output.extend_from_slice(&bytes(4, &[]));
        output.extend_from_slice(&field(5, 5));
        output.extend_from_slice(&bytes(ROW_STORAGE_FIELD, &storage));
        output.extend_from_slice(&bytes(ROW_OFFSETS_FIELD, &offsets));
        output.extend_from_slice(&field(ROW_WIDE_OFFSETS_FIELD, 0));
        output.extend_from_slice(unknown);
        output
    }

    fn tile(rows: &[Vec<u8>], unknown: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        output.extend_from_slice(&field(TILE_NUM_ROWS_FIELD, rows.len() as u64));
        for row in rows {
            output.extend_from_slice(&bytes(TILE_ROW_FIELD, row));
        }
        output.extend_from_slice(&field(6, 5));
        output.extend_from_slice(&field(7, 1));
        output.extend_from_slice(unknown);
        output
    }

    fn bnc(key: Option<u32>) -> Vec<u8> {
        let mut cell = BncCell::minimal();
        cell.set_text_style_identifier(key);
        cell.encode(&BncLimits::DEFAULT).expect("BNC cell")
    }

    fn tracked_tile(path: &str) -> Vec<u8> {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(path),
        )
        .expect("tracked iWork fixture");
        let catalog = SourceCatalog::from_bytes(&bytes, litchi_iwa_archive::Limits::default())
            .expect("tracked iWork catalog");
        for component in catalog.components().iter() {
            for object in &component.archive().objects {
                for message in &object.messages {
                    if message.type_ != 6_002 {
                        continue;
                    }
                    let view = WireView::parse(message.data.as_slice(), WireLimits::DEFAULT)
                        .expect("tile wire");
                    if one_varint(&view, 6).expect("tile storage version") == Some(5)
                        && one_varint(&view, 7).expect("tile BNC version") == Some(1)
                    {
                        return message.data.clone();
                    }
                }
            }
        }
        panic!("tracked fixture has no modern tile: {path}");
    }

    fn first_safe_update(tile: &[u8]) -> Update {
        let view = WireView::parse(tile, WireLimits::DEFAULT).expect("tile");
        if let Some(field) = view.fields().find(|field| field.number() == TILE_ROW_FIELD) {
            let row_view = WireView::parse(field.payload(), WireLimits::DEFAULT).expect("row");
            let row = one_varint(&row_view, ROW_INDEX_FIELD)
                .expect("row index")
                .expect("row index value") as u32;
            let storage = one_bytes(&row_view, ROW_STORAGE_FIELD)
                .expect("storage")
                .expect("storage value");
            let offsets = one_bytes(&row_view, ROW_OFFSETS_FIELD)
                .expect("offsets")
                .expect("offsets value");
            for (column, cell) in parse_slots(storage, offsets, 1)
                .expect("slots")
                .into_iter()
                .enumerate()
            {
                if let Some(cell) = cell
                    && let Ok(cell) = BncCell::parse(cell, &BncLimits::DEFAULT)
                    && !matches!(
                        cell.stored_value(),
                        litchi_numbers_wire::StoredValue::Formula(_)
                    )
                {
                    return Update {
                        row,
                        column: column as u32,
                        key: Some(cell.text_style_identifier().unwrap_or(1).saturating_add(1)),
                    };
                }
            }
            return Update {
                row,
                column: 0,
                key: Some(1),
            };
        }
        Update {
            row: 0,
            column: 0,
            key: Some(1),
        }
    }

    fn read_keys(tile: &[u8]) -> Vec<(u32, Vec<Option<u32>>)> {
        let view = WireView::parse(tile, WireLimits::DEFAULT).expect("tile view");
        let mut output = Vec::new();
        for field in view
            .fields()
            .filter(|field| field.number() == TILE_ROW_FIELD)
        {
            let row_view = WireView::parse(field.payload(), WireLimits::DEFAULT).expect("row");
            let row_number = one_varint(&row_view, ROW_INDEX_FIELD)
                .expect("row index")
                .expect("row index value") as u32;
            let storage = one_bytes(&row_view, ROW_STORAGE_FIELD)
                .expect("storage")
                .expect("storage value");
            let offsets = one_bytes(&row_view, ROW_OFFSETS_FIELD)
                .expect("offsets")
                .expect("offsets value");
            let cells = parse_slots(storage, offsets, 1).expect("slots");
            output.push((
                row_number,
                cells
                    .into_iter()
                    .map(|cell| {
                        cell.and_then(|bytes| {
                            BncCell::parse(bytes, &BncLimits::DEFAULT)
                                .expect("cell")
                                .text_style_identifier()
                        })
                    })
                    .collect(),
            ));
        }
        output
    }

    #[test]
    fn batch_rewrites_existing_and_empty_cells_once_and_preserves_unknowns() {
        let source_cell = bnc(Some(3));
        let unknown_tile = bytes(99, b"tile-tail");
        let unknown_row = bytes(99, b"row-tail");
        let source = tile(
            &[row(0, &[Some(&source_cell), None], &unknown_row)],
            &unknown_tile,
        );
        let transaction = transaction_limits();
        let wire = BncLimits::from_wire_limits(transaction.wire);
        let result = set_text_style_keys(
            &source,
            2,
            &[
                Update {
                    row: 0,
                    column: 0,
                    key: Some(7),
                },
                Update {
                    row: 0,
                    column: 1,
                    key: Some(8),
                },
            ],
            RowPlacement::FieldOrder,
            &transaction,
            &wire,
        )
        .expect("batch write");
        assert!(result.changed());
        assert_eq!(result.previous().len(), 2);
        assert_eq!(
            result.previous()[0].as_deref(),
            Some(source_cell.as_slice())
        );
        assert_eq!(result.previous()[1], None);
        assert_eq!(read_keys(result.tile()), vec![(0, vec![Some(7), Some(8)])]);
        assert!(
            result
                .tile()
                .windows(unknown_tile.len())
                .any(|window| window == unknown_tile)
        );
        assert!(
            result
                .tile()
                .windows(unknown_row.len())
                .any(|window| window == unknown_row)
        );
    }

    #[test]
    fn empty_row_materialization_and_exact_noop_are_bounded() {
        let source = tile(&[row(0, &[Some(&bnc(Some(3))), None], &[])], &[]);
        let transaction = transaction_limits();
        let wire = BncLimits::from_wire_limits(transaction.wire);
        let materialized = set_text_style_keys(
            &source,
            2,
            &[Update {
                row: 1,
                column: 1,
                key: Some(9),
            }],
            RowPlacement::FieldOrder,
            &transaction,
            &wire,
        )
        .expect("empty row write");
        assert_eq!(
            read_keys(materialized.tile()),
            vec![(0, vec![Some(3), None]), (1, vec![None, Some(9)])]
        );

        let empty_tile = tile(&[], &[]);
        let materialized_empty = set_text_style_keys(
            &empty_tile,
            2,
            &[Update {
                row: 0,
                column: 0,
                key: Some(4),
            }],
            RowPlacement::FieldOrder,
            &transaction_limits(),
            &wire,
        )
        .expect("empty tile write");
        assert_eq!(
            read_keys(materialized_empty.tile()),
            vec![(0, vec![Some(4), None])]
        );

        let no_op = set_text_style_keys(
            &source,
            2,
            &[Update {
                row: 0,
                column: 0,
                key: Some(3),
            }],
            RowPlacement::Append,
            &transaction_limits(),
            &wire,
        )
        .expect("exact no-op");
        assert!(!no_op.changed());
        assert_eq!(no_op.tile(), source.as_slice());
        assert_eq!(
            no_op.previous()[0].as_deref(),
            Some(bnc(Some(3)).as_slice())
        );
    }

    #[test]
    fn caller_wire_ceiling_remains_typed_before_publication() {
        let source = tile(&[row(0, &[Some(&bnc(Some(3))), None], &[])], &[]);
        let transaction = transaction_limits();
        let narrow = BncLimits::from_wire_limits(transaction.wire)
            .with_output_bytes(1)
            .expect("narrow BNC output profile");
        assert!(matches!(
            set_text_style_keys(
                &source,
                2,
                &[Update {
                    row: 0,
                    column: 0,
                    key: Some(7),
                }],
                RowPlacement::FieldOrder,
                &transaction,
                &narrow,
            ),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn transaction_ledgers_fail_closed_on_zero_residual_dimensions() {
        let source = tile(&[row(0, &[Some(&bnc(Some(3))), None], &[])], &[]);
        let update = Update {
            row: 0,
            column: 0,
            key: Some(7),
        };
        let wire = BncLimits::from_wire_limits(transaction_limits().wire);

        let transaction = transaction_limits();
        transaction
            .admit_work(transaction.max_work)
            .expect("debit work");
        assert!(matches!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            ),
            Err(Error::Limit { .. })
        ));

        let transaction = transaction_limits();
        transaction
            .charge_fields(transaction.wire.max_fields())
            .expect("debit fields");
        assert!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            )
            .is_err()
        );

        let transaction = transaction_limits();
        transaction
            .admit_allocations(transaction.max_allocations)
            .expect("debit allocations");
        assert!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            )
            .is_err()
        );

        let transaction = transaction_limits();
        transaction
            .admit_copies(transaction.max_copies)
            .expect("debit copies");
        assert!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            )
            .is_err()
        );

        let transaction = transaction_limits();
        transaction
            .admit_output(transaction.wire.max_output_bytes())
            .expect("debit output");
        assert!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            )
            .is_err()
        );
    }

    #[test]
    fn materialized_second_wire_pass_has_its_own_input_debit() {
        let source = tile(&[row(0, &[Some(&bnc(Some(3))), None], &[])], &[]);
        let transaction = transaction_limits();
        let remaining = transaction.wire.max_input_bytes();
        transaction
            .charge_input(remaining.saturating_sub(source.len()))
            .expect("leave exactly one raw-pass input budget");
        let wire = BncLimits::from_wire_limits(transaction.wire);
        assert!(matches!(
            set_text_style_keys(
                &source,
                2,
                &[Update {
                    row: 0,
                    column: 0,
                    key: Some(7),
                }],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            ),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                ..
            })
        ));
    }

    #[test]
    fn residual_bnc_item_cap_is_not_widened_by_target_columns() {
        let transaction = transaction_limits();
        let narrow = BncLimits::from_wire_limits(transaction.wire)
            .with_items(1)
            .expect("narrow items");
        let bounded = build_bnc_limits(&narrow, 100).expect("intersected items");
        assert_eq!(bounded.max_items(), 1);
    }

    #[test]
    fn cumulative_writer_work_refuses_one_unit_below_full_receipt() {
        let source = tile(&[row(0, &[Some(&bnc(Some(3))), None], &[])], &[]);
        let update = Update {
            row: 0,
            column: 0,
            key: Some(7),
        };
        let full = transaction_limits();
        let wire = BncLimits::from_wire_limits(full.wire);
        set_text_style_keys(
            &source,
            2,
            &[update],
            RowPlacement::FieldOrder,
            &full,
            &wire,
        )
        .expect("full writer budget");
        let used = full.report().work;
        assert!(used > 1, "writer must charge multiple work units");

        let limited = transaction_limits();
        let leave = used - 1;
        let debit = limited.max_work.saturating_sub(leave);
        limited.admit_work(debit).expect("leave residual work");
        assert!(matches!(
            set_text_style_keys(
                &source,
                2,
                &[update],
                RowPlacement::FieldOrder,
                &limited,
                &wire,
            ),
            Err(Error::Limit {
                kind: LimitKind::RewriteWork,
                ..
            })
        ));
    }

    #[test]
    fn tracked_numbers_pages_and_keynote_tiles_share_the_writer_contract() {
        for path in [
            "iwork/numbers/basic.numbers",
            "iwork/pages/body-table-cells-native.pages",
            "iwork/keynote/slide-table-read-native.key",
        ] {
            let tile = tracked_tile(path);
            let update = first_safe_update(&tile);
            let transaction = transaction_limits();
            let wire = BncLimits::from_wire_limits(transaction.wire);
            let output = set_text_style_keys(
                &tile,
                update.column.saturating_add(1).max(1),
                &[update],
                RowPlacement::FieldOrder,
                &transaction,
                &wire,
            )
            .expect("tracked tile style-key write");
            assert!(output.changed(), "fixture unexpectedly no-op: {path}");
            assert!(!output.tile().is_empty(), "empty tile output: {path}");
        }
    }
}
