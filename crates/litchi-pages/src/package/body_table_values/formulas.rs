//! The formula side of a body-table value edit (record 0376a): writes that
//! author a formula or replace one, staged on the structural working set,
//! because the shared engine
//! (`litchi_numbers_wire::formula_cache::{author, refresh}`) may create or
//! delete a cell-record tile.
//!
//! The order is the host's: every formula a write replaces is removed and
//! every authored formula is planned first, then the dependent caches of the
//! plain value writes are refreshed over that state, then the cells are
//! written (an authored cell takes its cached value and its formula key).

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_numbers_wire::{
    BncCellView, Limits, StoredValue,
    formula_cache::{self, DependencyPayload, author, refresh},
    table_write,
};

use super::{
    CALCULATION_ENGINE, DATA_LIST, Edit, Error, FORMULA_OWNER, HEADER_BUCKET, HEADER_BUCKET_ROWS,
    Package, Refusal, TABLE_MODEL, TILE, Value, cell_input, find_object, message, reference,
    storage_error, varint,
};
use crate::package::structure::{Fault, Transition, Working};

const CELL_RECORD_TILE: u32 = 4_009;

fn fault(fault: Fault) -> Error {
    match fault {
        Fault::Unsupported => Error::Refused(Refusal::UnsupportedSource),
        Fault::Allocation { amount } => Error::Allocation { amount },
        Fault::Verification => Error::Verification,
        Fault::Invalid | Fault::Dependency | Fault::Limit { .. } => Error::InvalidSource,
    }
}

fn refused(error: formula_cache::Error) -> Error {
    match error {
        formula_cache::Error::UnsupportedDependency(_) => Error::Refused(Refusal::Formulas),
        _ => Error::Refused(Refusal::UnsupportedStorage),
    }
}

/// Whether the edit authors a formula or writes over a cell that holds one.
pub(super) fn involved(edit: &Edit<'_>, limits: &Limits) -> Result<bool, Error> {
    if !edit.formulas.is_empty() {
        return Ok(true);
    }
    let target = edit.target.as_ref().ok_or(Error::InvalidSource)?;
    let package = edit.source;
    let payload = |identifier: u64, types: &[u32]| -> Result<&[u8], Error> {
        let object = find_object(package.state.source.components(), identifier)
            .ok_or(Error::InvalidSource)?;
        object
            .messages
            .iter()
            .find(|message| types.contains(&message.type_))
            .map(|message| message.data.as_slice())
            .ok_or(Error::InvalidSource)
    };
    let layout = Layout::read(payload(target.model_identifier.get(), &[TABLE_MODEL])?)?;
    for (row, column, _) in &edit.cells {
        let Some(tile) = layout.tile(*row) else {
            continue;
        };
        let cell = table_write::tile_cell(
            payload(tile, &[TILE])?,
            layout.local(*row)?,
            *column,
            limits,
        )
        .map_err(storage_error)?;
        if holds_formula(cell.as_deref(), limits)? {
            return Ok(true);
        }
    }
    Ok(false)
}

fn holds_formula(cell: Option<&[u8]>, limits: &Limits) -> Result<bool, Error> {
    Ok(cell
        .map(|cell| BncCellView::parse(cell, limits))
        .transpose()
        .map_err(|_error| Error::Refused(Refusal::UnsupportedStorage))?
        .is_some_and(|cell| matches!(cell.stored_value(), StoredValue::Formula(_))))
}

/// The tile layout of the table's data store.
struct Layout {
    tiles: Vec<(u64, u64)>,
    size: u64,
    strings: u64,
    buckets: Vec<u64>,
    column_headers: Option<u64>,
}

impl Layout {
    fn read(model: &[u8]) -> Result<Self, Error> {
        let store = message(model, 4)?
            .first()
            .copied()
            .ok_or(Error::InvalidSource)?;
        let tiles = message(store, 3)?
            .first()
            .copied()
            .ok_or(Error::InvalidSource)?;
        let size = varint(tiles, 2)?.unwrap_or(256).max(1);
        let mut entries = Vec::new();
        for entry in message(tiles, 1)? {
            let key = varint(entry, 1)?.ok_or(Error::InvalidSource)?;
            let tile = reference(entry, 2)?.ok_or(Error::InvalidSource)?;
            entries.push((key, tile));
        }
        let mut buckets = Vec::new();
        if let Some(headers) = message(store, 1)?.first() {
            for bucket in message(headers, 2)? {
                buckets.push(reference_value(bucket)?);
            }
        }
        Ok(Self {
            tiles: entries,
            size,
            strings: reference(store, 4)?.ok_or(Error::InvalidSource)?,
            buckets,
            column_headers: reference(store, 2)?,
        })
    }

    fn tile(&self, row: u32) -> Option<u64> {
        let key = u64::from(row) / self.size;
        self.tiles
            .iter()
            .find(|(candidate, _)| *candidate == key)
            .map(|(_, tile)| *tile)
    }

    fn local(&self, row: u32) -> Result<u32, Error> {
        u32::try_from(u64::from(row) % self.size).map_err(|_error| Error::InvalidSource)
    }
}

/// The identifier of a bare `TSP.Reference` payload.
fn reference_value(payload: &[u8]) -> Result<u64, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
    let field = view.fields().next().ok_or(Error::InvalidSource)?;
    litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
        .map(|(value, _)| value)
        .map_err(|_error| Error::InvalidSource)
}

/// The current first message of object `identifier`, with its type.
fn current(working: &Working<'_>, identifier: u64, types: &[u32]) -> Result<(u32, Vec<u8>), Error> {
    let object = working.object(identifier).map_err(fault)?;
    let message = object
        .messages
        .iter()
        .find(|message| types.contains(&message.type_))
        .ok_or(Error::InvalidSource)?;
    Ok((message.type_, message.data.clone()))
}

/// Replace the message of `message_type` in object `identifier`.
fn replace(
    working: &mut Working<'_>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
) -> Result<(), Error> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier).map_err(fault)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    object.replace_message_preserving_header(
        index,
        RawMessage {
            type_: message_type,
            data,
        },
        limits,
    )?;
    Ok(())
}

/// The source identifiers of the engine and every formula owner.
fn graph_identifiers(package: &Package) -> Result<(Option<u64>, Vec<u64>), Error> {
    let mut engine = None;
    let mut owners = Vec::new();
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let Some(message) = object.messages.first() else {
                continue;
            };
            let identifier = || object.archive_info.identifier.ok_or(Error::InvalidSource);
            if message.type_ == FORMULA_OWNER {
                owners.push(identifier()?);
            } else if message.type_ == CALCULATION_ENGINE && engine.replace(identifier()?).is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok((engine, owners))
}

fn staged_object<'w>(working: &'w Working<'_>, identifier: u64) -> Option<author::Object<'w>> {
    let object = working.object(identifier).ok()?;
    let message = object.messages.first()?;
    Some(author::Object {
        identifier,
        payload: &message.data,
        references: object
            .archive_info
            .message_infos
            .first()
            .map_or(&[][..], |info| info.object_references.as_slice()),
    })
}

#[allow(clippy::too_many_lines, reason = "one staged publication")]
pub(super) fn publish(edit: &Edit<'_>) -> Result<Package, Error> {
    let target = edit.target.as_ref().ok_or(Error::InvalidSource)?;
    let package = edit.source;
    let limits = package
        .numbers_wire_limits()
        .map_err(|_error| Error::InvalidSource)?;
    let mut working = Working::new(package).map_err(fault)?;
    let mut transition = Transition::default();
    let (engine_id, owner_ids) = graph_identifiers(package)?;
    let model = target.model_identifier.get();
    let info = target.drawable_identifier.get();
    let layout = Layout::read(&current(&working, model, &[TABLE_MODEL])?.1)?;

    // 1. The formula side of every cell, in staging order.
    let mut keys = Vec::new();
    for (row, column, _) in &edit.cells {
        let expression = edit
            .formulas
            .iter()
            .find(|(r, c, _)| (*r, *c) == (*row, *column))
            .map(|(_, _, expression)| expression);
        let tile = layout
            .tile(*row)
            .ok_or(Error::Refused(Refusal::UnsupportedStorage))?;
        let local = layout.local(*row)?;
        let (tile_type, tile_payload) = current(&working, tile, &[TILE])?;
        let previous = table_write::tile_cell(&tile_payload, local, *column, &limits)
            .map_err(storage_error)?;
        let formula_cell = holds_formula(previous.as_deref(), &limits)?;
        if !formula_cell && expression.is_none() {
            keys.push(None);
            continue;
        }
        let engine_id = engine_id.ok_or(Error::Refused(Refusal::Formulas))?;
        let next = working.next_identifier();
        let edit_plan = {
            let engine = staged_object(&working, engine_id).ok_or(Error::InvalidSource)?;
            let mut owners = Vec::new();
            for identifier in &owner_ids {
                owners.push(staged_object(&working, *identifier).ok_or(Error::InvalidSource)?);
            }
            let source = author::Source {
                engine,
                owners: &owners,
                table: info,
            };
            let object = |identifier: u64| staged_object(&working, identifier);
            let encoded = match expression {
                Some(expression) => {
                    let (owner, rows, columns) = author::table_owner(source, &object, &limits)
                        .map_err(refused)?
                        .ok_or(Error::Refused(Refusal::Formulas))?;
                    Some(
                        author::encode(expression, owner, *row, *column, rows, columns, &limits)
                            .map_err(refused)?,
                    )
                },
                None => None,
            };
            let change = match &encoded {
                Some(encoded) => author::Change::Author(encoded),
                None => author::Change::Remove,
            };
            author::plan(source, *row, *column, change, next, &object, &limits).map_err(refused)?
        };
        keys.push(edit_plan.key);
        let owner_member = working
            .locate(edit_plan.owner)
            .map_err(fault)?
            .ok_or(Error::InvalidSource)?;
        let registry = working.registry().map_err(fault)?;
        let mut payloads = edit_plan.payloads;
        if let Some(tile) = edit_plan.created_tile {
            let position = payloads
                .iter()
                .position(|(identifier, _)| *identifier == tile)
                .ok_or(Error::InvalidSource)?;
            let (_, payload) = payloads.remove(position);
            if working.allocate(1).map_err(fault)?.first() != Some(&tile) {
                return Err(Error::InvalidSource);
            }
            let object = ArchiveObject::new(
                tile,
                vec![RawMessage {
                    type_: CELL_RECORD_TILE,
                    data: payload,
                }],
                working.archive_limits(),
            )?;
            working.insert(&owner_member, object).map_err(fault)?;
            if let Some(registry) = &registry {
                let component = registry
                    .component_for_member(&owner_member)
                    .map_err(fault)?
                    .identifier;
                let seed = working.identity_seed(tile).map_err(fault)?;
                transition
                    .register_uuid(registry, component, tile, seed)
                    .map_err(fault)?;
            }
            let owner = working.object_mut(edit_plan.owner).map_err(fault)?;
            let message_info = owner
                .archive_info
                .message_infos
                .first_mut()
                .ok_or(Error::InvalidSource)?;
            if !message_info.object_references.contains(&tile) {
                message_info.object_references.push(tile);
            }
        }
        for (identifier, payload) in payloads {
            let (message_type, _) = current(
                &working,
                identifier,
                &[
                    CALCULATION_ENGINE,
                    FORMULA_OWNER,
                    CELL_RECORD_TILE,
                    4_010,
                    DATA_LIST[0],
                    DATA_LIST[1],
                    6_011,
                ],
            )?;
            replace(&mut working, identifier, message_type, payload)?;
        }
        if let Some(tile) = edit_plan.removed_tile {
            working.remove(tile).map_err(fault)?;
            if let Some(registry) = &registry {
                transition
                    .remove_objects(registry, &[tile])
                    .map_err(fault)?;
            }
            let owner = working.object_mut(edit_plan.owner).map_err(fault)?;
            let message_info = owner
                .archive_info
                .message_infos
                .first_mut()
                .ok_or(Error::InvalidSource)?;
            message_info
                .object_references
                .retain(|reference| *reference != tile);
        }
        if formula_cell {
            let stripped = table_write::formula_free(
                previous.as_deref().ok_or(Error::InvalidSource)?,
                &limits,
            )
            .map_err(|_error| Error::Refused(Refusal::Formulas))?;
            let (_, tile_payload) = current(&working, tile, &[TILE])?;
            let written = table_write::tile_set_cell(
                &tile_payload,
                local,
                *column,
                target.table_columns,
                Some(&stripped),
                &limits,
            )
            .map_err(storage_error)?;
            replace(&mut working, tile, tile_type, written.tile)?;
        }
    }

    // 2. The caches the plain value writes refresh, over that state.
    let mut inputs = Vec::new();
    for ((row, column, value), key) in edit.cells.iter().zip(&keys) {
        if key.is_none() {
            inputs.push((*row, *column, cell_input(value)));
        }
    }
    inputs.sort_unstable_by_key(|(row, column, _)| (*row, *column));
    let refresh = match engine_id {
        Some(engine_id) if !owner_ids.is_empty() && !inputs.is_empty() => {
            let engine = current(&working, engine_id, &[CALCULATION_ENGINE])?.1;
            let mut owner_payloads = Vec::new();
            for identifier in &owner_ids {
                owner_payloads.push((
                    *identifier,
                    current(&working, *identifier, &[FORMULA_OWNER])?.1,
                ));
            }
            let owners: Vec<DependencyPayload<'_>> = owner_payloads
                .iter()
                .map(|(identifier, bytes)| DependencyPayload {
                    identifier: *identifier,
                    bytes,
                })
                .collect();
            let writes: Vec<refresh::Write<'_>> = inputs
                .iter()
                .map(|(row, column, value)| refresh::Write {
                    row: *row,
                    column: *column,
                    value,
                })
                .collect();
            let object = |identifier: u64| -> Option<&[u8]> {
                working
                    .object(identifier)
                    .ok()
                    .and_then(|object| object.messages.first())
                    .map(|message| message.data.as_slice())
            };
            refresh::plan(
                refresh::Source {
                    engine: &engine,
                    owners: &owners,
                    table: info,
                },
                &writes,
                &object,
                &limits,
            )
            .map_err(refused)?
        },
        _ => refresh::Plan::default(),
    };

    // 3. The cells.
    for ((row, column, value), key) in edit.cells.iter().zip(&keys) {
        let tile = layout
            .tile(*row)
            .ok_or(Error::Refused(Refusal::UnsupportedStorage))?;
        let local = layout.local(*row)?;
        let (tile_type, tile_payload) = current(&working, tile, &[TILE])?;
        let previous = table_write::tile_cell(&tile_payload, local, *column, &limits)
            .map_err(storage_error)?;
        let (list_type, mut list) = current(&working, layout.strings, &DATA_LIST)?;
        let mut string_key = None;
        if let Value::Text(text) = value {
            let (interned, interned_key) =
                table_write::list_intern(&list, text, &limits).map_err(storage_error)?;
            list = interned;
            string_key = Some(interned_key);
        }
        let input = cell_input(value);
        let (cell, released) =
            table_write::cell_bytes(previous.as_deref(), &input, string_key, &limits)
                .map_err(storage_error)?;
        let cell = match key {
            Some(key) => Some(
                table_write::with_formula(
                    cell.as_deref().ok_or(Error::Refused(Refusal::Formulas))?,
                    *key,
                    &limits,
                )
                .map_err(storage_error)?,
            ),
            None => cell,
        };
        if let Some(released) = released {
            list = table_write::list_release(&list, released, &limits).map_err(storage_error)?;
        }
        let written = table_write::tile_set_cell(
            &tile_payload,
            local,
            *column,
            target.table_columns,
            cell.as_deref(),
            &limits,
        )
        .map_err(storage_error)?;
        let delta = i64::from(cell.is_some()) - i64::from(written.previous.is_some());
        replace(&mut working, tile, tile_type, written.tile)?;
        replace(&mut working, layout.strings, list_type, list)?;
        if delta != 0 {
            let bucket = *layout
                .buckets
                .get((*row / HEADER_BUCKET_ROWS) as usize)
                .ok_or(Error::Refused(Refusal::UnsupportedStorage))?;
            let (bucket_type, payload) = current(&working, bucket, &[HEADER_BUCKET])?;
            let adjusted = table_write::header_adjust(&payload, *row, delta, &limits)
                .map_err(storage_error)?;
            replace(&mut working, bucket, bucket_type, adjusted)?;
            if let Some(column_headers) = layout.column_headers {
                let (header_type, payload) = current(&working, column_headers, &[HEADER_BUCKET])?;
                let adjusted = table_write::header_adjust(&payload, *column, delta, &limits)
                    .map_err(storage_error)?;
                replace(&mut working, column_headers, header_type, adjusted)?;
            }
        }
    }

    // 4. The refreshed caches.
    for caches in &refresh.tiles {
        let (tile_type, payload) = current(&working, caches.tile, &[TILE])?;
        let refreshed = caches
            .apply(&payload, &limits)
            .map_err(|_error| Error::Refused(Refusal::Formulas))?;
        replace(&mut working, caches.tile, tile_type, refreshed)?;
    }
    let (candidate, _) = working.publish(&transition).map_err(fault)?;
    Ok(candidate)
}
