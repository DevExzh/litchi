//! A new empty table, cloned from the structure of the workbook's first
//! table (the template).
//!
//! The template keeps the new table native: its drawable settings, styles,
//! default row and column sizes and storage layout are the workbook's own.
//! The clone takes the template's table drawable and model, and one fresh
//! copy of each storage object the model owns, emptied: tiles hold no rows,
//! header buckets no headers, data lists no entries, the merge map no
//! ranges, and the row/column UID map and stroke sidecar are rebuilt for the
//! requested size. Styles stay shared. Everything the template owns beyond
//! its storage (sort, filter, category, pivot and hidden-state owners,
//! title, caption, comment) is left out. The table gets a new identity (a
//! table UUID, object UUIDs in the registry), a CalculationEngine owner so
//! formulas can be authored in it, a place at the end of its sheet's
//! drawables and a sidebar node. No clock or randomness is consulted: fresh
//! identities derive from the registry payload and the allocated
//! identifiers, so the same edit of the same source is byte-identical.

use std::collections::{BTreeMap, HashMap};

use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, parse_wire_view,
    repeated_length_delimited_payloads, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};

use super::super::Package;
use super::super::structure::{
    self, DOCUMENT_SHEETS_FIELD, DOCUMENT_SIDEBAR_FIELD, Fault, FaultResult, Published,
    TREE_CHILDREN_FIELD, TREE_NODE_MESSAGE_TYPE, TREE_OBJECT_FIELD, Working,
    metadata::{self, Transition},
    references,
};
use super::{Error, rewrite_list};

/// `TST.TableModelArchive`.
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TSCE.FormulaOwnerDependenciesArchive`.
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// The header versions native Numbers writes on a formula owner.
const FORMULA_OWNER_VERSIONS: [u32; 3] = [3, 2, 10];
/// `owner_kind` of a table's own formula owner.
const TABLE_FORMULA_OWNER_KIND: u64 = 1;
/// The most row and column identities one table may need (the native
/// editor's ceiling).
const MAX_TABLE_UIDS: u64 = 1_100_000;
/// How far a new table is offset from a template on the same sheet.
const POSITION_OFFSET: f32 = 40.0;

/// Model fields the clone drops: the template's identity of origin and
/// every owner or provider object beyond its storage.
const DROPPED_MODEL_FIELDS: [u32; 21] = [
    43, 5, 34, 35, 38, 39, 44, 45, 47, 52, 70, 81, 82, 83, 84, 85, 86, 90, 91, 92, 93,
];
/// Model counters reset to zero (hidden, user-hidden and filtered rows and
/// columns).
const ZEROED_MODEL_FIELDS: [u32; 5] = [14, 15, 41, 42, 40];
/// Table-info fields the clone drops, as the host editor clears them:
/// editing state, summary and category models, view UIDs and pivot state
/// (record 0370f; the group and hidden-state UUIDs, the pasteboard mapping
/// and the pivot flag are kept, as native tables carry them).
const DROPPED_INFO_FIELDS: [u32; 6] = [3, 4, 5, 6, 15, 17];
/// Drawable fields the clone drops: comment, pencil annotations, title and
/// caption.
const DROPPED_DRAWABLE_FIELDS: [u32; 4] = [6, 9, 10, 11];

/// The role of one storage object the template's model owns.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Role {
    Tile,
    Header,
    Data,
    MergeMap,
    UidMap,
    Stroke,
}

/// The data-store reference fields and the role of the object each holds.
const DATA_STORE_FIELDS: [(u32, Role); 14] = [
    (4, Role::Data),
    (5, Role::Data),
    (6, Role::Data),
    (11, Role::Data),
    (12, Role::Data),
    (13, Role::MergeMap),
    (15, Role::Data),
    (16, Role::Data),
    (17, Role::Data),
    (18, Role::Data),
    (19, Role::Data),
    (20, Role::Data),
    (21, Role::Data),
    (22, Role::Data),
];

/// The requested shape of the new table.
pub(super) struct Request<'n> {
    pub(super) sheet: usize,
    pub(super) name: &'n str,
    pub(super) rows: u32,
    pub(super) columns: u32,
}

/// Validate the requested size.
pub(super) fn check_size(rows: u32, columns: u32) -> Result<(), Error> {
    if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::InvalidSize { rows, columns },
        ));
    }
    Ok(())
}

/// The first table drawable of the workbook, in sheet and drawable order.
fn template(working: &Working<'_>) -> FaultResult<Option<(u64, u64)>> {
    let document = working.object(1)?;
    let index = structure::document_message(document)?;
    for sheet in references::identifiers(&document.messages[index].data, &[DOCUMENT_SHEETS_FIELD])?
    {
        let object = working.object(sheet)?;
        let (sheet_index, path) = structure::sheet_message(object)?;
        for drawable in references::identifiers(&object.messages[sheet_index].data, path)? {
            if structure::table_info_message(working.object(drawable)?).is_ok() {
                return Ok(Some((sheet, drawable)));
            }
        }
    }
    Ok(None)
}

/// The single reference payload in field `number` of `data`.
fn single_reference(data: &[u8], number: u32) -> FaultResult<Option<u64>> {
    match references::identifiers(data, &[number])?.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Fault::Invalid),
    }
}

/// Collect the storage objects the model owns, by role, in a stable order.
fn owned_storage(model: &[u8]) -> FaultResult<BTreeMap<u64, Role>> {
    let mut owned = BTreeMap::new();
    let mut insert = |identifier: u64, role: Role| -> FaultResult<()> {
        if let Some(previous) = owned.insert(identifier, role)
            && previous != role
        {
            return Err(Fault::Invalid);
        }
        Ok(())
    };
    let stores = repeated_length_delimited_payloads(model, 4)
        .map_err(|error| structure::map_wire(&error))?;
    let [store] = stores.as_slice() else {
        return Err(Fault::Invalid);
    };
    for bucket in references::identifiers(store, &[1, 2])? {
        insert(bucket, Role::Header)?;
    }
    if let Some(column_headers) = single_reference(store, 2)? {
        insert(column_headers, Role::Header)?;
    }
    let tiles = repeated_length_delimited_payloads(store, 3)
        .map_err(|error| structure::map_wire(&error))?;
    for storage in tiles {
        for entry in repeated_length_delimited_payloads(storage, 1)
            .map_err(|error| structure::map_wire(&error))?
        {
            if let Some(tile) = single_reference(entry, 2)? {
                insert(tile, Role::Tile)?;
            }
        }
    }
    for (field, role) in DATA_STORE_FIELDS {
        if let Some(identifier) = single_reference(store, field)? {
            insert(identifier, role)?;
        }
    }
    if let Some(identifier) = single_reference(model, 46)? {
        insert(identifier, Role::UidMap)?;
    }
    if let Some(identifier) = single_reference(model, 49)? {
        insert(identifier, Role::Stroke)?;
    }
    Ok(owned)
}

/// Replace every `TSP.Reference` at `path` of `data` whose identifier is in
/// `remap`.
fn remap_at(data: &[u8], path: &[u32], remap: &HashMap<u64, u64>) -> FaultResult<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |reference| {
        let identifier = references::reference_identifier(reference).map_err(WireFault)?;
        Ok::<_, WireFault>(match remap.get(&identifier) {
            Some(replacement) => references::reference_payload(*replacement).map_err(WireFault)?,
            None => structure::copy_bytes(reference).map_err(WireFault)?,
        })
    })
    .map_err(|error| error.0)
}

/// A [`Fault`] carried through the common wire transform callbacks.
struct WireFault(Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(structure::map_wire(&error))
    }
}

/// Encode `value` as a protobuf string payload of a canonical UUID.
fn table_uuid(seed: u64) -> (String, u64, u64) {
    let high = metadata::mix(seed ^ 0x7461_626c_6555_5549);
    let low = metadata::mix(high ^ seed);
    // RFC 4122 version 4 / variant 1 bit layout, as the app writes.
    let high = (high & 0xffff_ffff_ffff_0fff) | 0x0000_0000_0000_4000;
    let low = (low & 0x3fff_ffff_ffff_ffff) | 0x8000_0000_0000_0000;
    let text = format!(
        "{:08X}-{:04X}-{:04X}-{:04X}-{:012X}",
        high >> 32,
        (high >> 16) & 0xffff,
        high & 0xffff,
        low >> 48,
        low & 0xffff_ffff_ffff
    );
    // The CalculationEngine reads the 128-bit value with `lower` holding the
    // low half of the hexadecimal text.
    (text, low, high)
}

/// The model of the new table.
fn model_payload(
    template: &[u8],
    remap: &HashMap<u64, u64>,
    uuid: &str,
    request: &Request<'_>,
) -> FaultResult<Vec<u8>> {
    let mut data = structure::strip_fields(template, &DROPPED_MODEL_FIELDS)?;
    data = structure::put_bytes(&data, 1, uuid.as_bytes())?;
    data = structure::put_varint(&data, 50, 0)?;
    data = structure::put_bytes(&data, 8, request.name.as_bytes())?;
    data = structure::put_varint(&data, 6, u64::from(request.rows))?;
    data = structure::put_varint(&data, 7, u64::from(request.columns))?;
    for (field, limit) in [(9, request.rows), (10, request.columns), (11, request.rows)] {
        let values = litchi_iwa_common::wire::repeated_varint_values(&data, field)
            .map_err(|error| structure::map_wire(&error))?;
        if let [value] = values.as_slice() {
            data = structure::put_varint(&data, field, (*value).min(u64::from(limit)))?;
        }
    }
    for field in ZEROED_MODEL_FIELDS {
        data = structure::put_varint(&data, field, 0)?;
    }
    data = transform_length_delimited_field(&data, 4, |store| {
        let mut store = remap_at(store, &[1, 2], remap).map_err(WireFault)?;
        store = remap_at(&store, &[2], remap).map_err(WireFault)?;
        store = remap_at(&store, &[3, 1, 2], remap).map_err(WireFault)?;
        for (field, _) in DATA_STORE_FIELDS {
            store = remap_at(&store, &[field], remap).map_err(WireFault)?;
        }
        store = structure::put_varint(&store, 7, 1).map_err(WireFault)?;
        store = structure::put_varint(&store, 8, 0).map_err(WireFault)?;
        store = structure::put_bytes(&store, 9, &[]).map_err(WireFault)?;
        store = structure::put_bytes(&store, 10, &[]).map_err(WireFault)?;
        Ok::<_, WireFault>(store)
    })
    .map_err(|error| error.0)?;
    data = remap_at(&data, &[46], remap)?;
    remap_at(&data, &[49], remap)
}

/// The drawable of the new table.
fn info_payload(
    template: &[u8],
    template_sheet: u64,
    sheet: u64,
    model: u64,
) -> FaultResult<Vec<u8>> {
    let mut data = structure::strip_fields(template, &DROPPED_INFO_FIELDS)?;
    data = transform_length_delimited_field(&data, 1, |drawable| {
        let mut drawable =
            structure::strip_fields(drawable, &DROPPED_DRAWABLE_FIELDS).map_err(WireFault)?;
        drawable = structure::put_bytes(
            &drawable,
            2,
            &references::reference_payload(sheet).map_err(WireFault)?,
        )
        .map_err(WireFault)?;
        if template_sheet == sheet {
            drawable = transform_length_delimited_field(&drawable, 1, |geometry| {
                transform_length_delimited_field(geometry, 1, |position| {
                    offset_point(position).map_err(WireFault)
                })
            })?;
        }
        Ok::<_, WireFault>(drawable)
    })
    .map_err(|error| error.0)?;
    structure::put_bytes(&data, 2, &references::reference_payload(model)?)
}

/// Move a `TSP.Point` by [`POSITION_OFFSET`] on both axes.
fn offset_point(point: &[u8]) -> FaultResult<Vec<u8>> {
    let view = parse_wire_view(point, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| structure::map_wire(&error))?;
    let mut output = Vec::new();
    structure::reserve(&mut output, point.len())?;
    for field in view.fields() {
        let payload = field.payload();
        if (field.number() == 1 || field.number() == 2)
            && field.wire_type() == 5
            && payload.len() == 4
        {
            let value = f32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let moved = value + POSITION_OFFSET;
            if !moved.is_finite() {
                return Err(Fault::Invalid);
            }
            output.extend_from_slice(field.key());
            output.extend_from_slice(&moved.to_le_bytes());
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

fn encode_uuid(lower: u64, upper: u64) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    append_varint_field(&mut output, 1, lower).map_err(|error| structure::map_wire(&error))?;
    append_varint_field(&mut output, 2, upper).map_err(|error| structure::map_wire(&error))?;
    Ok(output)
}

/// An emptied copy of one storage object's payload.
fn empty_storage(
    role: Role,
    payload: &[u8],
    rows: u32,
    columns: u32,
    seed: u64,
) -> FaultResult<Vec<u8>> {
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    Ok(match role {
        Role::Tile => {
            let mut data = structure::strip_fields(payload, &[5])?;
            for field in [1, 2, 3, 4] {
                data = structure::put_varint(&data, field, 0)?;
            }
            data
        },
        Role::Header => structure::strip_fields(payload, &[2])?,
        Role::Data => {
            let data = structure::strip_fields(payload, &[3, 4])?;
            structure::put_varint(&data, 2, 1)?
        },
        Role::MergeMap => structure::strip_fields(payload, &[1])?,
        Role::UidMap => {
            // The host editor's fresh map (record 0370f): identities in
            // creation order derived from `seed` (the new model's
            // identifier), so the sorted lists and both index permutations
            // are the identity; the index lists are unpacked, as native maps
            // store them.
            let mut output = Vec::new();
            for index in 0..columns {
                let uuid = encode_uuid(
                    seed.wrapping_mul(0x9e37_79b9)
                        .wrapping_add(u64::from(index)),
                    seed.rotate_left(29) ^ 0x434f_4c55_4d4e_0000,
                )?;
                append_length_delimited_field(
                    &mut output,
                    1,
                    &uuid,
                    litchi_iwa_common::WireLimits::DEFAULT,
                )
                .map_err(wire)?;
            }
            for field in [2, 3] {
                for index in 0..columns {
                    append_varint_field(&mut output, field, u64::from(index)).map_err(wire)?;
                }
            }
            for index in 0..rows {
                let uuid = encode_uuid(
                    seed.wrapping_mul(0x517c_c1b7)
                        .wrapping_add(u64::from(index)),
                    seed.rotate_left(31) ^ 0x524f_5700_0000_0000,
                )?;
                append_length_delimited_field(
                    &mut output,
                    4,
                    &uuid,
                    litchi_iwa_common::WireLimits::DEFAULT,
                )
                .map_err(wire)?;
            }
            for field in [5, 6] {
                for index in 0..rows {
                    append_varint_field(&mut output, field, u64::from(index)).map_err(wire)?;
                }
            }
            output
        },
        Role::Stroke => {
            // The size alone, as the host editor writes a new sidecar.
            let mut output = Vec::new();
            append_varint_field(&mut output, 2, u64::from(columns)).map_err(wire)?;
            append_varint_field(&mut output, 3, u64::from(rows)).map_err(wire)?;
            output
        },
    })
}

/// The CalculationEngine root (component, object and message).
fn engine(working: &Working<'_>) -> FaultResult<Option<(String, u64, usize)>> {
    let mut found = None;
    for component in working.source().state.components.catalog().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                    if found
                        .replace((component.name().to_owned(), identifier, index))
                        .is_some()
                    {
                        return Err(Fault::Unsupported);
                    }
                }
            }
        }
    }
    Ok(found)
}

/// The next free internal formula-owner identifier of `engine`'s owners.
fn next_internal_owner(
    working: &Working<'_>,
    component: &str,
    engine_data: &[u8],
) -> FaultResult<u32> {
    let mut maximum = 0_u64;
    for object in &working.archive(component)?.objects {
        for message in &object.messages {
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                for value in litchi_iwa_common::wire::repeated_varint_values(&message.data, 2)
                    .map_err(|error| structure::map_wire(&error))?
                {
                    maximum = maximum.max(value);
                }
            }
        }
    }
    // Registered and unregistered map entries also hold internal identifiers.
    for tracker in repeated_length_delimited_payloads(engine_data, 2)
        .map_err(|error| structure::map_wire(&error))?
    {
        for map in repeated_length_delimited_payloads(tracker, 3)
            .map_err(|error| structure::map_wire(&error))?
        {
            for entry in repeated_length_delimited_payloads(map, 1)
                .map_err(|error| structure::map_wire(&error))?
            {
                for value in litchi_iwa_common::wire::repeated_varint_values(entry, 1)
                    .map_err(|error| structure::map_wire(&error))?
                {
                    maximum = maximum.max(value);
                }
            }
        }
    }
    u32::try_from(maximum.checked_add(1).ok_or(Fault::Invalid)?).map_err(|_error| Fault::Invalid)
}

/// The geometry a table's own formula owner records.
#[derive(Debug, Clone, Copy)]
struct Geometry {
    rows: u32,
    columns: u32,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
}

impl Geometry {
    /// The header and footer counts of the new model.
    fn of(model: &[u8], rows: u32, columns: u32) -> FaultResult<Self> {
        let count = |field: u32| -> FaultResult<u32> {
            let values = litchi_iwa_common::wire::repeated_varint_values(model, field)
                .map_err(|error| structure::map_wire(&error))?;
            Ok(values
                .first()
                .map_or(0, |value| u32::try_from(*value).unwrap_or(u32::MAX)))
        };
        Ok(Self {
            rows,
            columns,
            header_rows: count(9)?,
            header_columns: count(10)?,
            footer_rows: count(11)?,
        })
    }

    /// The spanning ranges of a table's own owner: the whole table
    /// (`RangeContext` total) and its body, `[first column, first row, last
    /// column, last row]`.
    fn spanning(self) -> FaultResult<Vec<u8>> {
        let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
        let last_column = u64::from(self.columns.saturating_sub(1));
        let last_row = u64::from(self.rows.saturating_sub(1));
        let body_last_row = u64::from(self.rows.saturating_sub(self.footer_rows).saturating_sub(1));
        let mut output = Vec::new();
        for (field, range) in [
            (2, [0, 0, last_column, last_row]),
            (
                3,
                [
                    u64::from(self.header_columns),
                    u64::from(self.header_rows),
                    last_column,
                    body_last_row,
                ],
            ),
        ] {
            let mut record = Vec::new();
            for (number, value) in (1_u32..).zip(range) {
                append_varint_field(&mut record, number, value).map_err(wire)?;
            }
            append_length_delimited_field(
                &mut output,
                field,
                &record,
                litchi_iwa_common::WireLimits::DEFAULT,
            )
            .map_err(wire)?;
        }
        Ok(output)
    }
}

/// A table's own empty formula owner, field for field as the native writer
/// emits it.
fn owner_payload(
    uid: (u64, u64),
    internal: u32,
    info: u64,
    geometry: Geometry,
) -> FaultResult<Vec<u8>> {
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    let mut output = Vec::new();
    append_length_delimited_field(
        &mut output,
        1,
        &encode_uuid(uid.0, uid.1)?,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    append_varint_field(&mut output, 2, u64::from(internal)).map_err(wire)?;
    append_varint_field(&mut output, 3, TABLE_FORMULA_OWNER_KIND).map_err(wire)?;
    append_length_delimited_field(&mut output, 4, &[], litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(wire)?;
    append_length_delimited_field(&mut output, 5, &[], litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(wire)?;
    let mut volatile = Vec::new();
    for field in [1, 2, 3, 4, 5, 7] {
        append_length_delimited_field(
            &mut volatile,
            field,
            &[],
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(wire)?;
    }
    append_length_delimited_field(
        &mut output,
        6,
        &volatile,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    let spanning = geometry.spanning()?;
    append_length_delimited_field(
        &mut output,
        7,
        &spanning,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    append_length_delimited_field(
        &mut output,
        8,
        &spanning,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    let mut whole = Vec::new();
    append_length_delimited_field(&mut whole, 1, &[], litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(wire)?;
    append_length_delimited_field(
        &mut output,
        9,
        &whole,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    append_length_delimited_field(&mut output, 10, &[], litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(wire)?;
    append_length_delimited_field(
        &mut output,
        11,
        &references::reference_payload(info)?,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    for field in [13, 14, 15, 16] {
        append_length_delimited_field(
            &mut output,
            field,
            &[],
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(wire)?;
    }
    Ok(output)
}

/// Register the new owner in the engine's dependency tracker: its reference
/// and its internal-identifier map entry.
fn engine_payload(
    engine: &[u8],
    owner: u64,
    internal: u32,
    uid: (u64, u64),
) -> FaultResult<Vec<u8>> {
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    let mut cfuuid = Vec::new();
    for (field, word) in [
        (2, uid.0 & 0xffff_ffff),
        (3, uid.0 >> 32),
        (4, uid.1 & 0xffff_ffff),
        (5, uid.1 >> 32),
    ] {
        append_varint_field(&mut cfuuid, field, word).map_err(wire)?;
    }
    let mut entry = Vec::new();
    append_varint_field(&mut entry, 1, u64::from(internal)).map_err(wire)?;
    append_length_delimited_field(
        &mut entry,
        2,
        &cfuuid,
        litchi_iwa_common::WireLimits::DEFAULT,
    )
    .map_err(wire)?;
    transform_length_delimited_field(engine, 2, |tracker| {
        let mut tracker = litchi_iwa_common::wire::append_repeated_length_delimited_field(
            tracker,
            6,
            &references::reference_payload(owner).map_err(WireFault)?,
        )?;
        if structure::has_field(&tracker, 3).map_err(WireFault)? {
            tracker = transform_length_delimited_field(&tracker, 3, |map| {
                litchi_iwa_common::wire::append_repeated_length_delimited_field(map, 1, &entry)
            })?;
        } else {
            let mut map = Vec::new();
            append_length_delimited_field(
                &mut map,
                1,
                &entry,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            tracker = structure::put_bytes(&tracker, 3, &map).map_err(WireFault)?;
        }
        Ok::<_, WireFault>(tracker)
    })
    .map_err(|error| error.0)
}

/// Add the new table and return the published candidate.
pub(super) fn add_table(source: &Package, request: &Request<'_>) -> Result<Published, Error> {
    check_size(request.rows, request.columns)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let sidebar_root = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SIDEBAR_FIELD],
    )?;
    let sheet = *sheets.get(request.sheet).ok_or(Error::Refused(
        crate::package::table_lifecycle::Refusal::SheetNotFound,
    ))?;
    let (template_sheet, template_info) = template(&working)?.ok_or(Error::Refused(
        crate::package::table_lifecycle::Refusal::NoTemplate,
    ))?;

    // The template's drawable, model and owned storage.
    let info_object = working.object(template_info)?.clone();
    let info_component = working.locate(template_info)?.ok_or(Error::InvalidSource)?;
    if info_object.messages.len() != 1 {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::UnsupportedSource,
        ));
    }
    let model = single_reference(&info_object.messages[0].data, 2)?.ok_or(Error::InvalidSource)?;
    let model_object = working.object(model)?.clone();
    let model_component = working.locate(model)?.ok_or(Error::InvalidSource)?;
    if model_object.messages.len() != 1
        || model_object.messages[0].type_ != TABLE_MODEL_MESSAGE_TYPE
    {
        return Err(Error::Refused(
            crate::package::table_lifecycle::Refusal::UnsupportedSource,
        ));
    }
    let owned = owned_storage(&model_object.messages[0].data)?;

    // Fresh identities.
    // The host editor's order: the drawable, the model and the owned
    // storage in source-identifier order; the formula owner and the sidebar
    // node follow (record 0370f).
    let fresh = working.allocate(owned.len() + 2)?;
    let (new_info, new_model) = (fresh[0], fresh[1]);
    let mut remap: HashMap<u64, u64> = HashMap::new();
    for ((identifier, _), replacement) in owned.iter().zip(&fresh[2..]) {
        remap.insert(*identifier, *replacement);
    }
    let seed = metadata::identity_seed(&working, new_info)?;
    let (uuid_text, uuid_lower, uuid_upper) = table_uuid(seed);
    let limits = working.archive_limits();

    // The storage clones, in their sources' components.
    let mut inserted: Vec<(u64, String)> = Vec::new();
    for (identifier, role) in &owned {
        let object = working.object(*identifier)?.clone();
        if object.messages.len() != 1 {
            return Err(Error::Refused(
                crate::package::table_lifecycle::Refusal::UnsupportedSource,
            ));
        }
        let component = working.locate(*identifier)?.ok_or(Error::InvalidSource)?;
        let data = empty_storage(
            *role,
            &object.messages[0].data,
            request.rows,
            request.columns,
            new_model,
        )?;
        let clone = structure::new_object(
            remap[identifier],
            Some(&object),
            object.messages[0].type_,
            data,
            Vec::new(),
            limits,
        )?;
        working.insert(&component, clone)?;
        inserted.push((remap[identifier], component));
    }

    // The model, with its styles shared and its storage remapped.
    let model_data = model_payload(&model_object.messages[0].data, &remap, &uuid_text, request)?;
    let geometry = Geometry::of(&model_data, request.rows, request.columns)?;
    let model_references = model_references(&model_data)?;
    let new_model_object = structure::new_object(
        new_model,
        Some(&model_object),
        TABLE_MODEL_MESSAGE_TYPE,
        model_data,
        model_references.clone(),
        limits,
    )?;
    working.insert(&model_component, new_model_object)?;
    inserted.push((new_model, model_component.clone()));

    // The drawable, parented to the target sheet.
    let info_data = info_payload(
        &info_object.messages[0].data,
        template_sheet,
        sheet,
        new_model,
    )?;
    let mut info_references: Vec<u64> = info_object.archive_info.message_infos[0]
        .object_references
        .iter()
        .copied()
        .filter(|identifier| *identifier != model && *identifier != template_sheet)
        .filter(|identifier| references_payload(&info_data, *identifier))
        .collect();
    info_references.push(sheet);
    info_references.push(new_model);
    let new_info_object = structure::new_object(
        new_info,
        Some(&info_object),
        info_object.messages[0].type_,
        info_data,
        info_references.clone(),
        limits,
    )?;
    working.insert(&info_component, new_info_object)?;
    inserted.push((new_info, info_component.clone()));

    // Attach to the sheet's drawables.
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let drawables = references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)?;
    let mut after = drawables.clone();
    after.push(new_info);
    rewrite_list(
        &mut working,
        sheet,
        sheet_index,
        sheet_path,
        &drawables,
        &after,
    )?;

    // A formula owner for the new table, when the workbook has an engine.
    let mut owner_component = None;
    let mut new_owner = None;
    if let Some((component, engine_id, engine_index)) = engine(&working)? {
        let owner = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
        new_owner = Some(owner);
        let new_owner = owner;
        let engine_data = working.object(engine_id)?.messages[engine_index]
            .data
            .clone();
        let internal = next_internal_owner(&working, &component, &engine_data)?;
        let uid = (uuid_upper.swap_bytes(), uuid_lower.swap_bytes());
        let owner_data = owner_payload(uid, internal, new_info, geometry)?;
        let mut owner_object = structure::new_object(
            new_owner,
            None,
            FORMULA_OWNER_MESSAGE_TYPE,
            owner_data,
            Vec::new(),
            limits,
        )?;
        owner_object.archive_info.message_infos[0].versions =
            FORMULA_OWNER_VERSIONS.to_vec().into();
        working.insert(&component, owner_object)?;
        inserted.push((new_owner, component.clone()));
        let engine_data = engine_payload(&engine_data, new_owner, internal, uid)?;
        let engine_object = working.object_mut(engine_id)?;
        references::replace(
            engine_object,
            engine_index,
            engine_data,
            &[],
            &[new_owner],
            &[2, 6],
            limits,
        )?;
        owner_component = Some(component);
    }

    // The sheet's sidebar node lists its tables.
    let mut node_parent = None;
    if let [root] = sidebar_root.as_slice() {
        let root_object = working.object(*root)?;
        let root_index = structure::unique_message(root_object, &[TREE_NODE_MESSAGE_TYPE])?;
        for node in references::identifiers(
            &root_object.messages[root_index].data,
            &[TREE_CHILDREN_FIELD],
        )? {
            let node_object = working.object(node)?;
            let node_index = structure::unique_message(node_object, &[TREE_NODE_MESSAGE_TYPE])?;
            if references::identifiers(
                &node_object.messages[node_index].data,
                &[TREE_OBJECT_FIELD],
            )?
            .as_slice()
                == [sheet]
            {
                node_parent = Some((node, node_index));
            }
        }
    }
    if let Some((parent, parent_index)) = node_parent {
        let new_node = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
        let parent_component = working.locate(parent)?.ok_or(Error::InvalidSource)?;
        let mut data = Vec::new();
        append_length_delimited_field(
            &mut data,
            TREE_OBJECT_FIELD,
            &references::reference_payload(new_info)?,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| structure::map_wire(&error))?;
        let template_node = working.object(parent)?.clone();
        let node_object = structure::new_object(
            new_node,
            Some(&template_node),
            TREE_NODE_MESSAGE_TYPE,
            data,
            vec![new_info],
            limits,
        )?;
        working.insert(&parent_component, node_object)?;
        inserted.push((new_node, parent_component));
        let children = references::identifiers(
            &working.object(parent)?.messages[parent_index].data,
            &[TREE_CHILDREN_FIELD],
        )?;
        let mut after = children.clone();
        after.push(new_node);
        rewrite_list(
            &mut working,
            parent,
            parent_index,
            &[TREE_CHILDREN_FIELD],
            &children,
            &after,
        )?;
    }

    // Registry: UUIDs for the clones of registered objects, and every new
    // cross-component reference.
    let mut transition = Transition::default();
    // In identifier order, as the host editor registers them.
    inserted.sort_by_key(|(identifier, _)| *identifier);
    if let Some(registry) = metadata::registry(&working)? {
        let registered = |identifier: u64| {
            registry
                .uuids
                .iter()
                .any(|binding| binding.object == identifier)
        };
        for (identifier, component) in &inserted {
            let source_identifier = if *identifier == new_info {
                Some(template_info)
            } else if *identifier == new_model {
                Some(model)
            } else {
                remap
                    .iter()
                    .find(|(_, new)| **new == *identifier)
                    .map(|(old, _)| *old)
            };
            let component_id = registry.component_for_entry(component)?.identifier;
            let needs_uuid = match source_identifier {
                Some(source_identifier) => registered(source_identifier),
                None => Some(*identifier) == new_owner,
            };
            if needs_uuid {
                transition.register_uuid(&registry, component_id, *identifier, seed)?;
            }
            let referenced: Vec<u64> = working.object(*identifier)?.archive_info.message_infos[0]
                .object_references
                .clone();
            for target in referenced {
                let target_component = working.locate(target)?.ok_or(Error::InvalidSource)?;
                let target_id = registry.component_for_entry(&target_component)?.identifier;
                transition.reference(&registry, component_id, target_id, target)?;
            }
        }
        // The sheet now addresses the drawable; the engine's owner addresses
        // the drawable through its `formula_owner`.
        let sheet_component = working.locate(sheet)?.ok_or(Error::InvalidSource)?;
        let sheet_component = registry.component_for_entry(&sheet_component)?.identifier;
        let info_component_id = registry.component_for_entry(&info_component)?.identifier;
        transition.reference(&registry, sheet_component, info_component_id, new_info)?;
        if let Some(component) = owner_component {
            let engine_component = registry.component_for_entry(&component)?.identifier;
            transition.reference(&registry, engine_component, info_component_id, new_info)?;
        }
    }
    Ok(working.publish(&transition)?)
}

/// The objects a model payload references, sorted, as the host editor's
/// `table_model_references` lists them (record 0370f): its style and owner
/// references, and its data store's headers, tiles and lists.
fn model_references(model: &[u8]) -> FaultResult<Vec<u64>> {
    const MODEL_FIELDS: [u32; 42] = [
        3, 24, 25, 26, 27, 18, 19, 20, 21, 30, 36, 48, 5, 34, 35, 38, 46, 49, 60, 61, 62, 63, 64,
        65, 66, 67, 68, 69, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 85, 86, 87, 88,
    ];
    const STORE_FIELDS: [u32; 15] = [2, 4, 5, 6, 11, 12, 16, 13, 15, 17, 18, 19, 20, 21, 22];
    let mut output = Vec::new();
    for field in MODEL_FIELDS.into_iter().chain([89]) {
        output.extend(references::identifiers(model, &[field])?);
    }
    output.extend(references::identifiers(model, &[4, 1, 2])?);
    output.extend(references::identifiers(model, &[4, 3, 1, 2])?);
    for field in STORE_FIELDS {
        output.extend(references::identifiers(model, &[4, field])?);
    }
    output.retain(|identifier| *identifier != 0);
    output.sort_unstable();
    output.dedup();
    Ok(output)
}

/// Whether `data` holds a `TSP.Reference` to `identifier` anywhere in its
/// nested messages. Only identifiers the header already declares are asked,
/// so a coincidental non-reference record is not a concern for correctness
/// of the declared set.
fn references_payload(data: &[u8], identifier: u64) -> bool {
    let Ok(view) = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT) else {
        return false;
    };
    view.fields().any(|field| {
        field.wire_type() == 2
            && (references::reference_identifier(field.payload()).ok() == Some(identifier)
                || references_payload(field.payload(), identifier))
    })
}
