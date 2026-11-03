//! Table layout: the trailing-edge resize (record 0370e, lifted from
//! `litchi-numbers` in record 0412g), uniform row and column sizes and the
//! table drawable's frame, as Keynote's host moves a slide table's bounds
//! with its size.

use litchi_iwa_common::wire as common_wire;
use litchi_iwa_protos::table_dimension_codec as dimension_codec;

use super::{
    Error, HEADER_BUCKET_MESSAGE_TYPE, Refusal, STROKE_SIDECAR_MESSAGE_TYPE, TILE_MESSAGE_TYPE,
    Target, UID_MAP_MESSAGE_TYPES, calc, geometry, message, replace_payload, storage, tiles,
};
use crate::structure::wire::{
    rebuild, repeated_bytes, single_bytes, single_reference, single_varint, varint_field,
};
use crate::structure::{self, Fault, TableLeaf, Working};

const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `TableModelArchive.number_of_rows` / `number_of_columns`.
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
/// `TableModelArchive.base_column_row_uids`.
const MODEL_UID_MAP_FIELD: u32 = 46;
/// `TableModelArchive.stroke_sidecar`.
const MODEL_STROKE_FIELD: u32 = 49;
/// The most row and column identities one table may need.
pub const MAX_TABLE_UIDS: u64 = 1_100_000;
/// Rows of one row-header bucket.
const HEADER_BUCKET_ROWS: u32 = 65_536;
const HEADER_CODEC_RECURSION_LIMIT: u32 = 64;
const HEADER_REWRITE_OUTPUT_SLACK: usize = 64;

fn storage_fault(error: litchi_numbers_wire::table_topology::Error) -> Error {
    match Fault::from(error) {
        Fault::Unsupported => Error::Refused(Refusal::UnsupportedLayout),
        fault => fault.into(),
    }
}

fn replace<P: TableLeaf>(
    working: &mut Working<'_, P>,
    object: u64,
    index: usize,
    data: Vec<u8>,
) -> Result<(), Error> {
    replace_payload(working, object, index, data)
}

/// Stage resizing the table to `rows` x `columns` at its trailing edges
/// (the host editor's `resize_table`); `false` when it already has that
/// size. Growth adds blank rows and columns; a shrink is refused when the
/// removed region holds a stored cell.
///
/// # Errors
///
/// [`Refusal::InvalidSize`], [`Refusal::StoredCells`],
/// [`Refusal::UnsupportedLayout`], source and resource errors.
#[allow(
    clippy::too_many_lines,
    reason = "one resize: tiles, headers, UID map, strokes, owner extents and model"
)]
pub fn stage_resize<P: TableLeaf>(
    working: &mut Working<'_, P>,
    target: &Target,
    rows: u32,
    columns: u32,
) -> Result<bool, Error> {
    if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
        return Err(Error::Refused(Refusal::InvalidSize));
    }
    let (model_index, model) = message(working, target.model, &[6_001])?;
    let model = model.to_vec();
    let before = geometry(&model)?;
    if (before.rows, before.columns) == (rows, columns) {
        return Ok(false);
    }
    if rows <= before.header_rows + before.footer_rows || columns <= before.header_columns {
        return Err(Error::Refused(Refusal::InvalidSize));
    }
    let store = single_bytes(&model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();

    // Tiles: removed rows go; no stored cell may lie outside the new size.
    let (tiles, size) = tiles(&store)?;
    for (key, tile) in &tiles {
        let (index, data) = message(working, *tile, &[TILE_MESSAGE_TYPE])?;
        let base = u64::from(*key) * size;
        let mut removed = false;
        let mut maximum: Option<u64> = None;
        let mut fault = None;
        let rewritten = rebuild(data, |field| {
            if field.number != 5 {
                return Ok(None);
            }
            let local = single_varint(field.payload, 1)?.ok_or(Fault::Invalid)?;
            if single_bytes(field.payload, 6)?.is_none() {
                fault.get_or_insert(Refusal::UnsupportedLayout);
                return Ok(None);
            }
            let cells = storage::bnc_cells(field.payload).map_err(Fault::from)?;
            if base + local >= u64::from(rows) {
                if !cells.is_empty() {
                    fault.get_or_insert(Refusal::StoredCells);
                    return Ok(None);
                }
                removed = true;
                return Ok(Some(Vec::new()));
            }
            if cells.iter().any(|(column, _)| *column >= columns) {
                fault.get_or_insert(Refusal::StoredCells);
            }
            maximum = maximum.max(Some(local));
            Ok(None)
        })?;
        if let Some(refusal) = fault {
            return Err(Error::Refused(refusal));
        }
        if removed {
            // The host editor counts the tile's rows by extent here.
            let numrows = maximum.map_or(0, |maximum| maximum + 1);
            let rewritten = rebuild(&rewritten, |field| {
                if field.number == 4 {
                    return varint_field(4, numrows).map(Some);
                }
                Ok(None)
            })?;
            replace(working, *tile, index, rewritten)?;
        }
    }

    // Header buckets keep the headers inside the new size.
    let headers = single_bytes(&store, 1)?.ok_or(Error::InvalidSource)?;
    let mut buckets: Vec<(u64, u32)> = Vec::new();
    for reference in repeated_bytes(headers, 2)? {
        buckets.push((
            single_varint(reference, 1)?.ok_or(Error::InvalidSource)?,
            rows,
        ));
    }
    if let Some(column_bucket) = single_reference(&store, 2)? {
        buckets.push((column_bucket, columns));
    }
    for (bucket, limit) in buckets {
        if bucket == 0 {
            continue;
        }
        let (index, data) = message(working, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let mut changed = false;
        let rewritten = rebuild(data, |field| {
            if field.number != 2 || single_varint(field.payload, 1)?.unwrap_or(0) < u64::from(limit)
            {
                return Ok(None);
            }
            changed = true;
            Ok(Some(Vec::new()))
        })?;
        if changed {
            replace(working, bucket, index, rewritten)?;
        }
    }

    // The row/column UID map.
    if let Some(uids) = single_reference(&model, MODEL_UID_MAP_FIELD)? {
        let (index, data) = message(working, uids, &UID_MAP_MESSAGE_TYPES)?;
        let map =
            storage::UidMap::read(data, before.rows, before.columns).map_err(storage_fault)?;
        let rewritten = map
            .resize(before.rows, rows, before.columns, columns)
            .map_err(storage_fault)?;
        replace(working, uids, index, rewritten)?;
    }

    // The stroke sidecar's size.
    if let Some(sidecar) = single_reference(&model, MODEL_STROKE_FIELD)? {
        let (index, data) = message(working, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
        let mut rewritten = data.to_vec();
        for (field, value) in [(2, columns), (3, rows)] {
            rewritten = litchi_iwa_common::wire::patch_varint_field(
                &rewritten,
                field,
                single_varint(&rewritten, field)?.is_some(),
                Some(u64::from(value)),
            )
            .map_err(|error| Error::from(structure::map_wire(&error)))?;
        }
        replace(working, sidecar, index, rewritten)?;
    }

    // The table's own formula owner: its table and body extents.
    let extents = |rows: u32, columns: u32| {
        calc::Extents::of(
            rows,
            columns,
            before.header_rows,
            before.header_columns,
            before.footer_rows,
        )
    };
    let mut owner = None;
    for object in working
        .source()
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE
                && single_reference(&message.data, 11)? == Some(target.info)
            {
                owner = Some((
                    object.archive_info.identifier.ok_or(Error::InvalidSource)?,
                    index,
                    message.data.clone(),
                ));
            }
        }
    }
    if let Some((identifier, index, data)) = owner
        && let Some(rewritten) = calc::resize_extents(
            &data,
            extents(before.rows, before.columns),
            extents(rows, columns),
        )
        .map_err(storage_fault)?
    {
        replace(working, identifier, index, rewritten)?;
    }

    // The model's size.
    let mut data = model;
    for (field, value) in [(MODEL_ROWS_FIELD, rows), (MODEL_COLUMNS_FIELD, columns)] {
        data = rebuild(&data, |candidate| {
            if candidate.number == field {
                return varint_field(field, u64::from(value)).map(Some);
            }
            Ok(None)
        })?;
    }
    replace(working, target.model, model_index, data)?;
    Ok(true)
}

/// Stage every row height as `height / rows` and every column width as
/// `width / columns` (Keynote's host `set_uniform_table_dimensions`, through
/// the shared header-size codec).
///
/// # Errors
///
/// A malformed model or header bucket, resource errors.
pub fn stage_uniform_dimensions<P: TableLeaf>(
    working: &mut Working<'_, P>,
    target: &Target,
    width: f32,
    height: f32,
) -> Result<(), Error> {
    let (_, model) = message(working, target.model, &[6_001])?;
    let model = model.to_vec();
    let table = geometry(&model)?;
    let row_height = uniform(height, table.rows)?;
    let column_width = uniform(width, table.columns)?;
    let store = single_bytes(&model, 4)?.ok_or(Error::InvalidSource)?;
    let row_storage = single_bytes(store, 1)?.ok_or(Error::InvalidSource)?;
    let mut slot = 0_u32;
    for reference in repeated_bytes(row_storage, 2)? {
        let bucket = single_varint(reference, 1)?.ok_or(Error::InvalidSource)?;
        let first = slot.saturating_mul(HEADER_BUCKET_ROWS);
        let last = first.saturating_add(HEADER_BUCKET_ROWS).min(table.rows);
        set_bucket_sizes(working, bucket, first, last, row_height)?;
        slot = slot.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    let column_bucket = single_reference(store, 2)?.ok_or(Error::InvalidSource)?;
    set_bucket_sizes(working, column_bucket, 0, table.columns, column_width)
}

fn uniform(total: f32, count: u32) -> Result<f32, Error> {
    if count == 0 {
        return Err(Error::InvalidSource);
    }
    let size = total / count as f32;
    if !size.is_finite() || size <= 0.0 {
        return Err(Error::Refused(Refusal::InvalidSize));
    }
    Ok(size)
}

fn dimension_options(max_bytes: usize) -> dimension_codec::DecodeOptions {
    let bounded = max_bytes.max(1);
    dimension_codec::DecodeOptions::new(
        bounded.clamp(1, litchi_iwa_common::WireLimits::MAX_INPUT_BYTES),
        bounded.clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS),
        bounded
            .saturating_mul(32)
            .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK),
        HEADER_CODEC_RECURSION_LIMIT,
        bounded,
        bounded,
    )
}

/// Set the size of every header index in `first..last` of one bucket.
fn set_bucket_sizes<P: TableLeaf>(
    working: &mut Working<'_, P>,
    bucket: u64,
    first: u32,
    last: u32,
    size: f32,
) -> Result<(), Error> {
    let (index, data) = message(working, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
    let data = data.to_vec();
    let mut edits = Vec::new();
    structure::reserve(
        &mut edits,
        usize::try_from(last.saturating_sub(first)).unwrap_or(0),
    )?;
    for header in first..last {
        edits.push(if size == 0.0 {
            dimension_codec::HeaderSizeEdit::remove(header)
        } else {
            dimension_codec::HeaderSizeEdit::set(header, size.to_bits())
        });
    }
    // Header bounds grow with the edits: each inserted minimal header adds a
    // bounded record.
    let slack = edits
        .len()
        .saturating_mul(16)
        .saturating_add(HEADER_REWRITE_OUTPUT_SLACK);
    let options = dimension_options(data.len().saturating_add(slack));
    let plan = dimension_codec::plan_header_storage_bucket_sizes(&data, last, &edits, options)
        .map_err(|_error| Error::InvalidSource)?;
    let requirements = plan.requirements();
    let (rewritten, _report) = dimension_codec::execute_header_storage_bucket_size_plan(
        plan,
        dimension_options(requirements.output_bytes()),
    )
    .map_err(|_error| Error::InvalidSource)?;
    if rewritten != data {
        replace(working, bucket, index, rewritten)?;
    }
    Ok(())
}

/// A table drawable's frame (`TSD.GeometryArchive`): position, size, flags
/// and rotation, each optional as the archive stores it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Frame {
    /// The origin (`x`, `y`).
    pub position: Option<(f32, f32)>,
    /// The size (`width`, `height`).
    pub size: Option<(f32, f32)>,
    /// The geometry flags.
    pub flags: Option<u32>,
    /// The rotation in degrees.
    pub angle: Option<f32>,
}

fn pair(data: Option<&[u8]>) -> Result<Option<(f32, f32)>, Error> {
    let Some(data) = data else {
        return Ok(None);
    };
    let read = |number| -> Result<f32, Error> {
        let view = common_wire::WireView::parse(data, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| Error::from(structure::map_wire(&error)))?;
        for field in view.fields().filter(|field| field.number() == number) {
            if field.wire_type() == 5 {
                let bytes: [u8; 4] = field
                    .payload()
                    .try_into()
                    .map_err(|_error| Error::InvalidSource)?;
                return Ok(f32::from_le_bytes(bytes));
            }
        }
        Err(Error::InvalidSource)
    };
    Ok(Some((read(1)?, read(2)?)))
}

/// The table drawable's current frame.
///
/// # Errors
///
/// A drawable without geometry or a malformed one.
pub fn frame<P: TableLeaf>(working: &Working<'_, P>, target: &Target) -> Result<Frame, Error> {
    let (_, info) = message(working, target.info, &[6_000, 6_003])?;
    let drawable = single_bytes(info, 1)?.ok_or(Error::InvalidSource)?;
    let geometry = single_bytes(drawable, 1)?.ok_or(Error::InvalidSource)?;
    let angle = {
        let view = common_wire::WireView::parse(geometry, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|error| Error::from(structure::map_wire(&error)))?;
        let mut angle = None;
        for field in view.fields().filter(|field| field.number() == 4) {
            let bytes: [u8; 4] = field
                .payload()
                .try_into()
                .map_err(|_error| Error::InvalidSource)?;
            angle = Some(f32::from_le_bytes(bytes));
        }
        angle
    };
    Ok(Frame {
        position: pair(single_bytes(geometry, 1)?)?,
        size: pair(single_bytes(geometry, 2)?)?,
        flags: single_varint(geometry, 3)?
            .map(u32::try_from)
            .transpose()
            .map_err(|_error| Error::InvalidSource)?,
        angle,
    })
}

fn patch_pair(
    data: &[u8],
    number: u32,
    current: Option<(f32, f32)>,
    replacement: Option<(f32, f32)>,
) -> Result<Vec<u8>, litchi_iwa_common::Error> {
    match (current, replacement) {
        (Some(_), Some((first, second))) => {
            common_wire::transform_length_delimited_field(data, number, |value| {
                let value =
                    common_wire::patch_fixed32_field(value, 1, true, Some(first.to_bits()))?;
                common_wire::patch_fixed32_field(&value, 2, true, Some(second.to_bits()))
            })
        },
        (Some(_), None) => common_wire::patch_length_delimited_field(data, number, true, None),
        (None, Some((first, second))) => {
            let mut value = Vec::with_capacity(10);
            value.push(0x0d);
            value.extend_from_slice(&first.to_le_bytes());
            value.push(0x15);
            value.extend_from_slice(&second.to_le_bytes());
            common_wire::patch_length_delimited_field(data, number, false, Some(&value))
        },
        (None, None) => Ok(data.to_vec()),
    }
}

/// Stage the table drawable's frame (the host's `patch_drawable_geometry`
/// field by field: present fields patched in place, absent ones appended).
///
/// # Errors
///
/// A drawable without geometry or a malformed one.
pub fn stage_frame<P: TableLeaf>(
    working: &mut Working<'_, P>,
    target: &Target,
    replacement: Frame,
) -> Result<(), Error> {
    let current = frame(working, target)?;
    let (index, info) = message(working, target.info, &[6_000, 6_003])?;
    let info = info.to_vec();
    let wire = |error: litchi_iwa_common::Error| Error::from(structure::map_wire(&error));
    let rewritten = common_wire::transform_length_delimited_field(&info, 1, |drawable| {
        common_wire::transform_length_delimited_field(drawable, 1, |geometry| {
            let data = patch_pair(geometry, 1, current.position, replacement.position)?;
            let data = patch_pair(&data, 2, current.size, replacement.size)?;
            let data = common_wire::patch_varint_field(
                &data,
                3,
                current.flags.is_some(),
                replacement.flags.map(u64::from),
            )?;
            common_wire::patch_fixed32_field(
                &data,
                4,
                current.angle.is_some(),
                replacement.angle.map(f32::to_bits),
            )
        })
    })
    .map_err(wire)?;
    if rewritten != info {
        replace(working, target.info, index, rewritten)?;
    }
    Ok(())
}
