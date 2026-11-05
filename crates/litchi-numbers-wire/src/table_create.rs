//! Payload builders for a new, empty table cloned from the structure of an
//! existing one (the template), shared by the Keynote and Pages table
//! owners (record 0329).
//!
//! This is the payload half of `litchi-numbers`' `table::lifecycle::add`
//! (record 0301), which the Keynote and Pages owners pair with their own
//! object-graph plumbing (identifier allocation, component placement,
//! registry identities and references). The clone takes the template's
//! table drawable and model and one emptied copy of each storage object the
//! model owns: tiles hold no rows, header buckets no headers, data lists no
//! entries, the merge map no ranges, and the row/column UID map and stroke
//! sidecar are rebuilt for the requested size. Styles stay shared.
//! Everything the template owns beyond its storage (sort, filter, category,
//! pivot and hidden-state owners, summary model, title, caption, comment)
//! is left out. No clock or randomness is consulted: fresh identities
//! derive from caller-supplied seeds.

use std::collections::BTreeMap;

use litchi_iwa_common::wire::{
    append_length_delimited_field, append_repeated_length_delimited_field, append_varint_field,
    parse_wire_view, patch_fixed64_field, patch_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads, repeated_varint_values, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};

use crate::table_write::WriteError;

type Result<T> = std::result::Result<T, WriteError>;

pub mod fresh;

/// `TST.TableInfoArchive`.
pub const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
/// `TST.TableModelArchive`.
pub const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TSCE.CalculationEngineArchive`.
pub const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TSCE.FormulaOwnerDependenciesArchive`.
pub const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// The header versions native Numbers writes on a formula owner.
pub const FORMULA_OWNER_VERSIONS: [u32; 3] = [3, 2, 10];
/// `owner_kind` of a table's own formula owner.
const TABLE_FORMULA_OWNER_KIND: u64 = 1;
/// The most row and column identities one table may need (the native
/// editor's ceiling).
pub const MAX_TABLE_UIDS: u64 = 1_100_000;

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

/// The role of one storage object a table model owns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Role {
    /// A `TST.Tile`.
    Tile,
    /// A `TST.HeaderStorageBucket`.
    Header,
    /// A `TST.TableDataList`.
    Data,
    /// The merge-region map.
    MergeMap,
    /// The row/column UID map.
    UidMap,
    /// The stroke sidecar.
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

/// The requested shape of a new table.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Shape<'n> {
    /// The table name.
    pub name: &'n str,
    /// Rows, header and footer rows included.
    pub rows: u32,
    /// Columns, header columns included.
    pub columns: u32,
    /// A uniform row height in points, replacing the template's default.
    pub row_height: Option<f64>,
    /// A uniform column width in points, replacing the template's default.
    pub column_width: Option<f64>,
}

/// Whether `rows` × `columns` is a size a new table may take.
#[must_use]
pub fn admits_size(rows: u32, columns: u32) -> bool {
    rows > 0 && columns > 0 && u64::from(rows) + u64::from(columns) <= MAX_TABLE_UIDS
}

fn wire(_error: litchi_iwa_common::Error) -> WriteError {
    WriteError::Invalid
}

/// A [`WriteError`] carried through the common wire transform callbacks.
struct Fault(WriteError);

impl From<litchi_iwa_common::Error> for Fault {
    fn from(_error: litchi_iwa_common::Error) -> Self {
        Self(WriteError::Invalid)
    }
}

/// The encoding of a `TSP.Reference` to `identifier`.
#[must_use]
pub fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut output = Vec::with_capacity(11);
    output.push(0x08);
    litchi_iwa_common::varint::encode_varint_into(&mut output, identifier);
    output
}

/// The identifier of a `TSP.Reference` payload.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a payload other than one non-zero identifier.
pub fn reference_identifier(payload: &[u8]) -> Result<u64> {
    let view = parse_wire_view(payload, litchi_iwa_common::WireLimits::DEFAULT).map_err(wire)?;
    if view
        .fields()
        .any(|field| field.number() != 1 || field.wire_type() != 0)
    {
        return Err(WriteError::Unsupported);
    }
    match repeated_varint_values(payload, 1).map_err(wire)?.as_slice() {
        [identifier] if *identifier != 0 => Ok(*identifier),
        _ => Err(WriteError::Invalid),
    }
}

/// The identifiers of the repeated reference field `number` of `data`.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed payload.
pub fn references(data: &[u8], number: u32) -> Result<Vec<u64>> {
    repeated_length_delimited_payloads(data, number)
        .map_err(wire)?
        .into_iter()
        .map(reference_identifier)
        .collect()
}

/// The single reference in field `number` of `data`, if any.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed payload or a repeated field.
pub fn single_reference(data: &[u8], number: u32) -> Result<Option<u64>> {
    match references(data, number)?.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(WriteError::Invalid),
    }
}

/// The parent a `TST.TableInfoArchive` payload's drawable names, if any.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed payload.
pub fn drawable_parent(info: &[u8]) -> Result<Option<u64>> {
    match repeated_length_delimited_payloads(info, 1)
        .map_err(wire)?
        .as_slice()
    {
        [] => Ok(None),
        [drawable] => single_reference(drawable, 2),
        _ => Err(WriteError::Invalid),
    }
}

fn strip_fields(data: &[u8], excluded: &[u32]) -> Result<Vec<u8>> {
    let view = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT).map_err(wire)?;
    let mut output = Vec::with_capacity(data.len());
    for field in view.fields() {
        if !excluded.contains(&field.number()) {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

fn has_field(data: &[u8], number: u32) -> Result<bool> {
    let view = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT).map_err(wire)?;
    Ok(view.fields().any(|field| field.number() == number))
}

fn put_varint(data: &[u8], number: u32, value: u64) -> Result<Vec<u8>> {
    let present = has_field(data, number)?;
    patch_varint_field(data, number, present, Some(value)).map_err(wire)
}

fn put_bytes(data: &[u8], number: u32, payload: &[u8]) -> Result<Vec<u8>> {
    let present = has_field(data, number)?;
    patch_length_delimited_field(data, number, present, Some(payload)).map_err(wire)
}

fn put_double(data: &[u8], number: u32, value: f64) -> Result<Vec<u8>> {
    if !value.is_finite() || value <= 0.0 {
        return Err(WriteError::Invalid);
    }
    let present = has_field(data, number)?;
    patch_fixed64_field(data, number, present, Some(value.to_bits())).map_err(wire)
}

/// The storage objects a template's model owns, by identifier, with their
/// roles.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed model or an object named in two
/// roles.
pub fn owned_storage(model: &[u8]) -> Result<BTreeMap<u64, Role>> {
    let mut owned = BTreeMap::new();
    let mut insert = |identifier: u64, role: Role| -> Result<()> {
        if let Some(previous) = owned.insert(identifier, role)
            && previous != role
        {
            return Err(WriteError::Invalid);
        }
        Ok(())
    };
    let stores = repeated_length_delimited_payloads(model, 4).map_err(wire)?;
    let [store] = stores.as_slice() else {
        return Err(WriteError::Invalid);
    };
    for header in repeated_length_delimited_payloads(store, 1).map_err(wire)? {
        for bucket in references(header, 2)? {
            insert(bucket, Role::Header)?;
        }
    }
    if let Some(column_headers) = single_reference(store, 2)? {
        insert(column_headers, Role::Header)?;
    }
    for storage in repeated_length_delimited_payloads(store, 3).map_err(wire)? {
        for entry in repeated_length_delimited_payloads(storage, 1).map_err(wire)? {
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

/// Replace every `TSP.Reference` at `path` of `data` that `remap` maps.
fn remap_at(data: &[u8], path: &[u32], remap: &dyn Fn(u64) -> Option<u64>) -> Result<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |reference| {
        let identifier = reference_identifier(reference).map_err(Fault)?;
        Ok::<_, Fault>(match remap(identifier) {
            Some(replacement) => reference_payload(replacement),
            None => reference.to_vec(),
        })
    })
    .map_err(|error| error.0)
}

/// A deterministic 64-bit mix (SplitMix64's finalizer).
#[must_use]
pub const fn mix(mut value: u64) -> u64 {
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

/// A table's identity: the canonical UUID text of its model (`table_id`)
/// and the 128-bit value the calculation engine reads.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableUuid {
    /// The upper-case hyphenated text.
    pub text: String,
    /// The low half of the hexadecimal text.
    pub lower: u64,
    /// The high half of the hexadecimal text.
    pub upper: u64,
}

impl TableUuid {
    /// The identity derived from `seed`, with the RFC 4122 version 4 and
    /// variant 1 bits the app writes.
    #[must_use]
    pub fn derive(seed: u64) -> Self {
        let high = mix(seed ^ 0x7461_626c_6555_5549);
        Self::from_halves(high, mix(high ^ seed))
    }

    /// The identity of a 128-bit value (`high` the first half of the text),
    /// with the RFC 4122 version 4 and variant 1 bits the app writes; the
    /// owners' seed policies derive the halves.
    #[must_use]
    pub fn from_halves(high: u64, low: u64) -> Self {
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
        Self {
            text,
            lower: low,
            upper: high,
        }
    }

    /// The formula owner's UID for a table of this identity, as the
    /// calculation engine pairs them.
    #[must_use]
    pub const fn owner_uid(&self) -> (u64, u64) {
        (self.upper.swap_bytes(), self.lower.swap_bytes())
    }
}

/// The model of the new table: the template's model with its dropped
/// fields removed, its storage remapped through `remap`, its counters
/// reset and its identity, name and size replaced.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed template or an invalid size.
pub fn model_payload(
    template: &[u8],
    remap: &dyn Fn(u64) -> Option<u64>,
    uuid: &TableUuid,
    shape: &Shape<'_>,
) -> Result<Vec<u8>> {
    if !admits_size(shape.rows, shape.columns) {
        return Err(WriteError::Invalid);
    }
    let mut data = strip_fields(template, &DROPPED_MODEL_FIELDS)?;
    data = put_bytes(&data, 1, uuid.text.as_bytes())?;
    data = put_varint(&data, 50, 0)?;
    data = put_bytes(&data, 8, shape.name.as_bytes())?;
    data = put_varint(&data, 6, u64::from(shape.rows))?;
    data = put_varint(&data, 7, u64::from(shape.columns))?;
    for (field, limit) in [(9, shape.rows), (10, shape.columns), (11, shape.rows)] {
        let values = repeated_varint_values(&data, field).map_err(wire)?;
        if let [value] = values.as_slice() {
            data = put_varint(&data, field, (*value).min(u64::from(limit)))?;
        }
    }
    // Header and footer rows together never exceed the table.
    let header_rows = repeated_varint_values(&data, 9)
        .map_err(wire)?
        .first()
        .copied()
        .unwrap_or(0);
    let footer_rows = repeated_varint_values(&data, 11)
        .map_err(wire)?
        .first()
        .copied()
        .unwrap_or(0);
    if header_rows + footer_rows > u64::from(shape.rows) && has_field(&data, 11)? {
        data = put_varint(&data, 11, u64::from(shape.rows).saturating_sub(header_rows))?;
    }
    for field in ZEROED_MODEL_FIELDS {
        data = put_varint(&data, field, 0)?;
    }
    if let Some(height) = shape.row_height {
        data = put_double(&data, 16, height)?;
    }
    if let Some(width) = shape.column_width {
        data = put_double(&data, 17, width)?;
    }
    data = transform_length_delimited_field(&data, 4, |store| {
        let mut store = remap_at(store, &[1, 2], remap).map_err(Fault)?;
        store = remap_at(&store, &[2], remap).map_err(Fault)?;
        store = remap_at(&store, &[3, 1, 2], remap).map_err(Fault)?;
        for (field, _) in DATA_STORE_FIELDS {
            store = remap_at(&store, &[field], remap).map_err(Fault)?;
        }
        store = put_varint(&store, 7, 1).map_err(Fault)?;
        store = put_varint(&store, 8, 0).map_err(Fault)?;
        store = put_bytes(&store, 9, &[]).map_err(Fault)?;
        store = put_bytes(&store, 10, &[]).map_err(Fault)?;
        Ok::<_, Fault>(store)
    })
    .map_err(|error| error.0)?;
    data = remap_at(&data, &[46], remap)?;
    remap_at(&data, &[49], remap)
}

/// A frame in points for the new drawable.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Frame {
    /// Left edge.
    pub x: f32,
    /// Top edge.
    pub y: f32,
    /// Width.
    pub width: f32,
    /// Height.
    pub height: f32,
}

fn point(x: f32, y: f32) -> Result<Vec<u8>> {
    if !x.is_finite() || !y.is_finite() {
        return Err(WriteError::Invalid);
    }
    let mut output = Vec::with_capacity(10);
    for (number, value) in [(1_u32, x), (2, y)] {
        litchi_iwa_common::varint::encode_varint_into(&mut output, u64::from(number << 3 | 5));
        output.extend_from_slice(&value.to_le_bytes());
    }
    Ok(output)
}

/// The drawable of the new table: the template's `TST.TableInfoArchive`
/// without its dropped fields, parented to `parent`, naming `model`, at
/// `frame` (position and size; the template's other geometry fields stay).
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed template or a non-finite frame.
pub fn info_payload(template: &[u8], parent: u64, model: u64, frame: Frame) -> Result<Vec<u8>> {
    if !(frame.width.is_finite() && frame.height.is_finite())
        || frame.width <= 0.0
        || frame.height <= 0.0
    {
        return Err(WriteError::Invalid);
    }
    let position = point(frame.x, frame.y)?;
    let size = point(frame.width, frame.height)?;
    let mut data = strip_fields(template, &DROPPED_INFO_FIELDS)?;
    data = transform_length_delimited_field(&data, 1, |drawable| {
        let mut drawable = strip_fields(drawable, &DROPPED_DRAWABLE_FIELDS).map_err(Fault)?;
        drawable = put_bytes(&drawable, 2, &reference_payload(parent)).map_err(Fault)?;
        if has_field(&drawable, 1).map_err(Fault)? {
            drawable = transform_length_delimited_field(&drawable, 1, |geometry| {
                let geometry = put_bytes(geometry, 1, &position).map_err(Fault)?;
                put_bytes(&geometry, 2, &size).map_err(Fault)
            })?;
        } else {
            let mut geometry = Vec::new();
            append_length_delimited_field(
                &mut geometry,
                1,
                &position,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            append_length_delimited_field(
                &mut geometry,
                2,
                &size,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            drawable = put_bytes(&drawable, 1, &geometry).map_err(Fault)?;
        }
        Ok::<_, Fault>(drawable)
    })
    .map_err(|error| error.0)?;
    put_bytes(&data, 2, &reference_payload(model))
}

fn encode_uuid(lower: u64, upper: u64) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    append_varint_field(&mut output, 1, lower).map_err(wire)?;
    append_varint_field(&mut output, 2, upper).map_err(wire)?;
    Ok(output)
}

/// An emptied copy of one storage object's payload for a table of `rows` ×
/// `columns`; `seed` derives the row and column identities of a UID map.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed payload.
pub fn empty_storage(
    role: Role,
    payload: &[u8],
    rows: u32,
    columns: u32,
    seed: u64,
) -> Result<Vec<u8>> {
    Ok(match role {
        Role::Tile => {
            let mut data = strip_fields(payload, &[5])?;
            for field in [1, 2, 3, 4] {
                data = put_varint(&data, field, 0)?;
            }
            data
        },
        Role::Header => strip_fields(payload, &[2])?,
        Role::Data => {
            let data = strip_fields(payload, &[3, 4])?;
            put_varint(&data, 2, 1)?
        },
        Role::MergeMap => strip_fields(payload, &[1])?,
        Role::UidMap => {
            // The host editor's fresh map (record 0370f): identities in
            // creation order derived from `seed` (the new model's
            // identifier), so the sorted lists and both index permutations
            // are the identity; the index lists are unpacked, as native maps
            // store them.
            // Encoded in one linear pass (record 0370i): appending each
            // field through the checked appender re-parsed the growing
            // output, quadratic in the table size.
            use litchi_iwa_common::varint::encode_varint_into;
            let mut output = Vec::new();
            output
                .try_reserve(26 * (columns as usize + rows as usize) + 64)
                .map_err(|_error| WriteError::Invalid)?;
            let uuid = |output: &mut Vec<u8>, number: u32, lower: u64, upper: u64| {
                let mut payload = Vec::with_capacity(22);
                encode_varint_into(&mut payload, 1 << 3);
                encode_varint_into(&mut payload, lower);
                encode_varint_into(&mut payload, 2 << 3);
                encode_varint_into(&mut payload, upper);
                encode_varint_into(output, (u64::from(number) << 3) | 2);
                encode_varint_into(output, payload.len() as u64);
                output.extend_from_slice(&payload);
            };
            let varint = |output: &mut Vec<u8>, number: u32, value: u64| {
                encode_varint_into(output, u64::from(number) << 3);
                encode_varint_into(output, value);
            };
            for index in 0..columns {
                uuid(
                    &mut output,
                    1,
                    seed.wrapping_mul(0x9e37_79b9)
                        .wrapping_add(u64::from(index)),
                    seed.rotate_left(29) ^ 0x434f_4c55_4d4e_0000,
                );
            }
            for field in [2, 3] {
                for index in 0..columns {
                    varint(&mut output, field, u64::from(index));
                }
            }
            for index in 0..rows {
                uuid(
                    &mut output,
                    4,
                    seed.wrapping_mul(0x517c_c1b7)
                        .wrapping_add(u64::from(index)),
                    seed.rotate_left(31) ^ 0x524f_5700_0000_0000,
                );
            }
            for field in [5, 6] {
                for index in 0..rows {
                    varint(&mut output, field, u64::from(index));
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

/// The next free internal formula-owner identifier, over the engine's
/// formula owners (`TSCE.FormulaOwnerDependenciesArchive` payloads) and the
/// engine's dependency-tracker maps.
///
/// # Errors
///
/// [`WriteError::Invalid`] for malformed payloads or an exhausted space.
pub fn next_internal_owner(owners: &[&[u8]], engine: &[u8]) -> Result<u32> {
    let mut maximum = 0_u64;
    for owner in owners {
        for value in repeated_varint_values(owner, 2).map_err(wire)? {
            maximum = maximum.max(value);
        }
    }
    for tracker in repeated_length_delimited_payloads(engine, 2).map_err(wire)? {
        for map in repeated_length_delimited_payloads(tracker, 3).map_err(wire)? {
            for entry in repeated_length_delimited_payloads(map, 1).map_err(wire)? {
                for value in repeated_varint_values(entry, 1).map_err(wire)? {
                    maximum = maximum.max(value);
                }
            }
        }
    }
    u32::try_from(maximum.checked_add(1).ok_or(WriteError::Invalid)?)
        .map_err(|_error| WriteError::Invalid)
}

/// The geometry a table's own formula owner records.
#[derive(Debug, Clone, Copy)]
pub struct OwnerGeometry {
    rows: u32,
    columns: u32,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
}

impl OwnerGeometry {
    /// The geometry of a new model payload of `rows` × `columns`.
    ///
    /// # Errors
    ///
    /// [`WriteError::Invalid`] for a malformed model.
    pub fn of(model: &[u8], rows: u32, columns: u32) -> Result<Self> {
        let count = |field: u32| -> Result<u32> {
            let values = repeated_varint_values(model, field).map_err(wire)?;
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

    /// The spanning ranges of a table's own owner: the whole table and its
    /// body, `[first column, first row, last column, last row]`.
    fn spanning(self) -> Result<Vec<u8>> {
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
///
/// # Errors
///
/// [`WriteError::Invalid`] when an encoding fails.
pub fn owner_payload(
    uid: (u64, u64),
    internal: u32,
    info: u64,
    geometry: OwnerGeometry,
) -> Result<Vec<u8>> {
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
        &reference_payload(info),
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

/// Register a new owner in the engine's dependency tracker: its reference
/// and its internal-identifier map entry.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed engine.
pub fn engine_payload(
    engine: &[u8],
    owner: u64,
    internal: u32,
    uid: (u64, u64),
) -> Result<Vec<u8>> {
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
        let mut tracker =
            append_repeated_length_delimited_field(tracker, 6, &reference_payload(owner))?;
        if has_field(&tracker, 3).map_err(Fault)? {
            tracker = transform_length_delimited_field(&tracker, 3, |map| {
                append_repeated_length_delimited_field(map, 1, &entry)
            })?;
        } else {
            let mut map = Vec::new();
            append_length_delimited_field(
                &mut map,
                1,
                &entry,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            tracker = put_bytes(&tracker, 3, &map).map_err(Fault)?;
        }
        Ok::<_, Fault>(tracker)
    })
    .map_err(|error| error.0)
}

/// Whether `data` holds a `TSP.Reference` to `identifier` anywhere in its
/// nested messages. Only identifiers a header already declares should be
/// asked, so a coincidental non-reference record cannot add a declaration.
#[must_use]
pub fn references_identifier(data: &[u8], identifier: u64) -> bool {
    let Ok(view) = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT) else {
        return false;
    };
    view.fields().any(|field| {
        field.wire_type() == 2
            && (reference_identifier(field.payload()).ok() == Some(identifier)
                || references_identifier(field.payload(), identifier))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sizes_are_bounded() {
        assert!(admits_size(1, 1));
        assert!(!admits_size(0, 3));
        assert!(!admits_size(3, 0));
        assert!(!admits_size(1_000_000, 100_001));
    }

    #[test]
    fn identities_are_deterministic_and_versioned() {
        let first = TableUuid::derive(42);
        assert_eq!(first, TableUuid::derive(42));
        assert_ne!(first, TableUuid::derive(43));
        assert_eq!(first.text.len(), 36);
        assert_eq!(&first.text[14..15], "4");
        assert_eq!(first.upper & 0xf000, 0x4000);
    }

    #[test]
    fn references_round_trip() {
        let payload = reference_payload(300);
        assert_eq!(reference_identifier(&payload), Ok(300));
        let mut data = Vec::new();
        append_length_delimited_field(
            &mut data,
            3,
            &payload,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut data,
            3,
            &reference_payload(7),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        assert_eq!(references(&data, 3), Ok(vec![300, 7]));
        assert_eq!(single_reference(&data, 3), Err(WriteError::Invalid));
        assert!(references_identifier(&data, 7));
        assert!(!references_identifier(&data, 8));
    }

    #[test]
    fn info_takes_the_frame_parent_and_model() {
        let mut geometry = Vec::new();
        append_length_delimited_field(
            &mut geometry,
            1,
            &point(1.0, 2.0).unwrap(),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut geometry,
            2,
            &point(3.0, 4.0).unwrap(),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_varint_field(&mut geometry, 3, 3).unwrap();
        let mut drawable = Vec::new();
        append_length_delimited_field(
            &mut drawable,
            1,
            &geometry,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut drawable,
            2,
            &reference_payload(5),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut drawable,
            10,
            &reference_payload(6),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        let mut info = Vec::new();
        append_length_delimited_field(
            &mut info,
            1,
            &drawable,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut info,
            2,
            &reference_payload(9),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        append_length_delimited_field(
            &mut info,
            4,
            &reference_payload(10),
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .unwrap();
        let frame = Frame {
            x: 10.0,
            y: 20.0,
            width: 300.0,
            height: 90.0,
        };
        let output = info_payload(&info, 77, 88, frame).unwrap();
        assert_eq!(single_reference(&output, 2), Ok(Some(88)));
        assert!(!has_field(&output, 4).unwrap());
        let drawable = repeated_length_delimited_payloads(&output, 1).unwrap()[0];
        assert_eq!(single_reference(drawable, 2), Ok(Some(77)));
        assert!(!has_field(drawable, 10).unwrap());
        let geometry = repeated_length_delimited_payloads(drawable, 1).unwrap()[0];
        assert_eq!(
            repeated_length_delimited_payloads(geometry, 2).unwrap()[0],
            point(300.0, 90.0).unwrap().as_slice()
        );
        assert_eq!(repeated_varint_values(geometry, 3), Ok(vec![3]));
        assert!(
            info_payload(
                &info,
                77,
                88,
                Frame {
                    width: 0.0,
                    ..frame
                }
            )
            .is_err()
        );
    }
}
