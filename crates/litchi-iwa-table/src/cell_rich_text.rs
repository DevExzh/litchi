//! Format-neutral rich text of whole table cells (record 0412): the
//! character and paragraph properties of a cell's text style
//! ([`text`](crate::cell_rich_text::text)) and the paragraph lists of its rich-text storage ([`list`](crate::cell_rich_text::list)).
//!
//! Numbers sheets, Pages body tables and Keynote slide tables store a cell's
//! text style the same way: a key in the table's style list
//! (`DataStore.styleTable`) whose entry names a `TSWP.ParagraphStyleArchive`;
//! a cell without a key uses its section's base style. A cell's lists live in
//! a `TSWP.StorageArchive` named by a `TST.RichTextPayloadArchive` of the
//! table's rich-text list. The engines here were lifted from `litchi-numbers`
//! (records 0370 and 0370b) and run on any [`TableLeaf`](crate::structure::TableLeaf): the leaf resolves
//! its selector to the table model's identifier, opens a
//! [`Working`](crate::structure::Working) set over its exact source and keeps its own public
//! transaction vocabulary, patch and candidate verification. The property,
//! list and character planning is the shared text engines'
//! ([`litchi_iwa_text_wire::paragraph_style`], [`litchi_iwa_text_wire::list`]);
//! nothing here re-plans a style.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the engine keeps each graph helper beside the phase that uses it"
)]

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_core::RawMessage;
use thiserror::Error as ThisError;

use crate::structure::wire::{
    bytes_field, fields, push_bytes_field, rebuild, repeated_bytes, single_bytes, single_reference,
    single_varint,
};
use crate::structure::{self, Fault, FaultLimit, FaultResult, TableLeaf, Working, references};

/// Paragraph lists of one cell's text.
pub mod list;
/// One character or paragraph property of whole cells' text.
pub mod text;

/// `TST.TableModelArchive`.
pub const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TST.Tile`.
pub const TILE_MESSAGE_TYPE: u32 = 6_002;
/// `TST.TableDataList`.
pub const LIST_MESSAGE_TYPE: u32 = 6_005;
/// `TST.TableDataListSegment`.
pub const LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
/// `TST.HeaderStorageBucket`.
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// `TSWP.ParagraphStyleArchive`.
pub const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
/// `TSS.StylesheetArchive`.
pub const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `DataStore.styleTable`.
const STYLE_LIST_FIELD: u32 = 5;
/// `TSWP.ParagraphStyleArchive.para_properties`.
const PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
/// The rows of a storage tile when the tile storage does not say.
const DEFAULT_TILE_SIZE: u64 = 256;

/// A finite resource a rich-text transaction governs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// A bounded wire operation of the shared text engines.
    Wire(litchi_iwa_common::LimitKind),
}

/// A typed, content-redacted rich-text engine failure. Each leaf maps it
/// onto its public family error.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request the engine refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell rich-text source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind:?} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} cell rich-text items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell rich-text candidate verification failed")]
    Verification,
}

/// A well-formed request the engine refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// The cell's style or rich text lives outside what the engine rewrites
    /// (no stylesheet, a segmented list, a shared rich-text payload, no
    /// `None` list preset).
    #[error("a surviving object depends on the cell's rich-text graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell rich text")]
    UnsupportedSource,
    /// The position is not the start of a paragraph of the cell's text.
    #[error("position is not a paragraph start")]
    NotAParagraphStart,
    /// The paragraph's list kind has no such property.
    #[error("the paragraph's list has no such property")]
    NotApplicable,
    /// Placements must start at zero and strictly increase.
    #[error("list placements must start at zero and strictly increase")]
    InvalidPlacements,
    /// The cell's paragraphs do not share one list.
    #[error("the cell's paragraphs do not share one list")]
    NotUniform,
    /// The cell holds a value other than text.
    #[error("the cell does not hold text")]
    NotText,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

/// A table's size and section counts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Geometry {
    /// Rows of the table.
    pub rows: u32,
    /// Columns of the table.
    pub columns: u32,
    /// Header rows.
    pub header_rows: u32,
    /// Header columns.
    pub header_columns: u32,
    /// Footer rows.
    pub footer_rows: u32,
}

impl Geometry {
    /// Whether `position` lies inside the table.
    #[must_use]
    pub const fn contains(self, position: At) -> bool {
        position.row() < self.rows && position.column() < self.columns
    }
}

/// The size and sections of a `TST.TableModelArchive` payload.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a model without a size or with sections
/// larger than the table.
pub fn geometry(model: &[u8]) -> Result<Geometry, Error> {
    let count = |number| -> Result<u32, Error> {
        u32::try_from(single_varint(model, number)?.unwrap_or(0))
            .map_err(|_error| Error::InvalidSource)
    };
    let geometry = Geometry {
        rows: u32::try_from(single_varint(model, 6)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
        columns: u32::try_from(single_varint(model, 7)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
        header_rows: count(9)?,
        header_columns: count(10)?,
        footer_rows: count(11)?,
    };
    if geometry.header_rows + geometry.footer_rows > geometry.rows
        || geometry.header_columns > geometry.columns
    {
        return Err(Error::InvalidSource);
    }
    Ok(geometry)
}

/// The unique message of `identifier` with one of `types`.
///
/// # Errors
///
/// A [`Fault`] when the object is missing or holds no unique such message.
pub fn message<'w, P: TableLeaf>(
    working: &'w Working<'_, P>,
    identifier: u64,
    types: &[u32],
) -> FaultResult<(usize, &'w [u8])> {
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, types)?;
    Ok((index, &object.messages[index].data))
}

/// One resolved table: its model identifier and payload, data store and
/// geometry.
pub struct Table {
    /// The `TST.TableModelArchive` identifier.
    pub model: u64,
    /// The model payload.
    pub payload: Vec<u8>,
    /// The model's `DataStore` payload.
    pub store: Vec<u8>,
    /// The table's size and sections.
    pub geometry: Geometry,
}

/// The table whose model is `model`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a model without a data store or size.
pub fn table<P: TableLeaf>(working: &Working<'_, P>, model: u64) -> Result<Table, Error> {
    let (_, payload) = message(working, model, &[TABLE_MODEL_MESSAGE_TYPE])?;
    let geometry = geometry(payload)?;
    let store = single_bytes(payload, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();
    Ok(Table {
        model,
        payload: payload.to_vec(),
        store,
        geometry,
    })
}

/// The storage tiles of a data store: `(key, tile object)`, and the rows
/// per tile.
fn tiles(store: &[u8]) -> Result<(Vec<(u32, u64)>, u64), Error> {
    let storage = single_bytes(store, 3)?.ok_or(Error::InvalidSource)?;
    let size = single_varint(storage, 2)?.unwrap_or(DEFAULT_TILE_SIZE);
    if size == 0 {
        return Err(Error::InvalidSource);
    }
    let mut output = Vec::new();
    for entry in repeated_bytes(storage, 1)? {
        let key = u32::try_from(single_varint(entry, 1)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
        let tile = single_reference(entry, 2)?.ok_or(Error::InvalidSource)?;
        if output.iter().any(|(existing, _)| *existing == key) {
            return Err(Error::InvalidSource);
        }
        output
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        output.push((key, tile));
    }
    Ok((output, size))
}

/// One stored BNC cell located in the table's tiles.
pub struct Located {
    /// The tile object.
    pub tile: u64,
    /// The tile message's index in its object.
    pub tile_index: usize,
    /// The field index of the cell's row record among the tile's fields.
    pub row_field: usize,
    /// The cell's BNC bytes.
    pub bytes: Vec<u8>,
}

/// The stored cell at `position` of the data store `store`, if any.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed tile storage.
pub fn locate_cell<P: TableLeaf>(
    working: &Working<'_, P>,
    store: &[u8],
    position: At,
) -> Result<Option<Located>, Error> {
    let (tiles, size) = tiles(store)?;
    let key = u64::from(position.row()) / size;
    let local = u64::from(position.row()) % size;
    let Some((_, tile)) = tiles
        .iter()
        .find(|(tile_key, _)| u64::from(*tile_key) == key)
    else {
        return Ok(None);
    };
    let (tile_index, data) = message(working, *tile, &[TILE_MESSAGE_TYPE])?;
    for (row_field, field) in fields(data)?.iter().enumerate() {
        if field.number != 5 || single_varint(field.payload, 1)? != Some(local) {
            continue;
        }
        for (column, bytes) in
            litchi_numbers_wire::table_topology::storage::bnc_cells(field.payload)
                .map_err(Fault::from)?
        {
            if column == position.column() {
                return Ok(Some(Located {
                    tile: *tile,
                    tile_index,
                    row_field,
                    bytes: bytes.to_vec(),
                }));
            }
        }
    }
    Ok(None)
}

/// The raw-cell profile derived from the cell's own source extent.
#[must_use]
pub fn cell_limits(source: &[u8]) -> litchi_numbers_wire::Limits {
    if source.is_empty() {
        return litchi_numbers_wire::Limits::DEFAULT;
    }
    litchi_numbers_wire::Limits::DEFAULT
        .with_input_bytes(source.len())
        .and_then(|limits| limits.with_tail_bytes(source.len()))
        .unwrap_or(litchi_numbers_wire::Limits::DEFAULT)
}

/// The entry of `key` in the style list (root or segments): its reference.
///
/// # Errors
///
/// [`Error::InvalidSource`] when no entry holds `key`.
pub fn list_reference<P: TableLeaf>(
    working: &Working<'_, P>,
    list: &[u8],
    key: u32,
) -> Result<Option<u64>, Error> {
    let find = |data: &[u8]| -> FaultResult<Option<Option<u64>>> {
        for entry in repeated_bytes(data, 3)? {
            if single_varint(entry, 1)? == Some(u64::from(key)) {
                return Ok(Some(single_reference(entry, 4)?));
            }
        }
        Ok(None)
    };
    if let Some(found) = find(list)? {
        return Ok(found);
    }
    for reference in repeated_bytes(list, 4)? {
        let segment = references::reference_identifier(reference)?;
        let (_, data) = message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])?;
        if let Some(found) = find(data)? {
            return Ok(found);
        }
    }
    Err(Error::InvalidSource)
}

/// The section base text style of `position` in the model `model`.
///
/// # Errors
///
/// [`Refusal::UnsupportedDependency`] when the model names no base style
/// for the section.
pub fn base_style(model: &[u8], geometry: Geometry, position: At) -> Result<u64, Error> {
    let field = if position.row() < geometry.header_rows {
        25
    } else if position.row() >= geometry.rows.saturating_sub(geometry.footer_rows) {
        27
    } else if position.column() < geometry.header_columns {
        26
    } else {
        24
    };
    single_reference(model, field)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

/// The cell's text-style key, if it has a stored cell with one.
fn style_key(located: Option<&Located>) -> Result<Option<u32>, Error> {
    let Some(located) = located else {
        return Ok(None);
    };
    let cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &cell_limits(&located.bytes))
        .map_err(|_error| Error::InvalidSource)?;
    Ok(cell.text_style_identifier())
}

/// The current text style object of the cell at `position` (its key's
/// entry, else its section's base style).
///
/// # Errors
///
/// [`Refusal::CellNotFound`] outside the table; source errors.
pub fn current_style<P: TableLeaf>(
    working: &Working<'_, P>,
    model: u64,
    position: At,
) -> Result<u64, Error> {
    let (_, payload) = message(working, model, &[TABLE_MODEL_MESSAGE_TYPE])?;
    let geometry = geometry(payload)?;
    if !geometry.contains(position) {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let store = single_bytes(payload, 4)?.ok_or(Error::InvalidSource)?;
    let base = base_style(payload, geometry, position)?;
    let Some(key) = style_key(locate_cell(working, store, position)?.as_ref())? else {
        return Ok(base);
    };
    let list = single_reference(store, STYLE_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
    let (_, list_data) = message(working, list, &[LIST_MESSAGE_TYPE, 6_201])?;
    list_reference(working, list_data, key)?.ok_or(Error::InvalidSource)
}

/// The wire profile of the shared text engines, bounded by the working
/// set's message ceiling.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the ceiling cannot form a profile.
pub fn style_wire_limits<P: TableLeaf>(working: &Working<'_, P>) -> Result<WireLimits, Error> {
    let maximum = working.archive_limits().max_message_bytes();
    WireLimits::default()
        .with_input_bytes(maximum.min(WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| limits.with_output_bytes(maximum.min(WireLimits::MAX_OUTPUT_BYTES)))
        .map_err(|_| Error::InvalidSource)
}

/// `style` with an empty `para_properties` when it has none.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed style.
pub fn with_paragraph_properties(mut style: Vec<u8>) -> Result<Vec<u8>, Error> {
    if single_bytes(&style, PARAGRAPH_PROPERTIES_FIELD)?.is_none() {
        push_bytes_field(&mut style, PARAGRAPH_PROPERTIES_FIELD, &[])?;
    }
    Ok(style)
}

/// The stylesheet a paragraph style belongs to (`super.stylesheet`).
///
/// # Errors
///
/// [`Refusal::UnsupportedDependency`] for a style outside a stylesheet.
pub fn stylesheet_of<P: TableLeaf>(working: &Working<'_, P>, style: u64) -> Result<u64, Error> {
    let (_, data) = message(working, style, &[PARAGRAPH_STYLE_MESSAGE_TYPE])?;
    let base = single_bytes(data, 1)?.ok_or(Error::InvalidSource)?;
    single_reference(base, 5)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

/// Replace one message, transitioning its aggregate object references.
///
/// # Errors
///
/// A [`Fault`] of the working set.
pub fn replace_with_references<P: TableLeaf>(
    working: &mut Working<'_, P>,
    identifier: u64,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    references::replace(object, index, data, removed, appended, &[], limits)
}

/// Replace one message, keeping its header.
///
/// # Errors
///
/// A [`Fault`] of the working set.
pub fn replace_message<P: TableLeaf>(
    working: &mut Working<'_, P>,
    identifier: u64,
    index: usize,
    data: Vec<u8>,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    Ok(())
}

/// Store `cell` at the empty position `position`, counting it in its row
/// and column headers.
///
/// # Errors
///
/// [`Refusal::UnsupportedDependency`] when the position's tile or header
/// bucket is missing; [`Refusal::UnsupportedSource`] for a tile the writer
/// does not own; source and allocation errors.
pub fn materialize_cell<P: TableLeaf>(
    working: &mut Working<'_, P>,
    store: &[u8],
    columns: u32,
    position: At,
    cell: &litchi_numbers_wire::BncCell,
) -> Result<(), Error> {
    let (tiles, size) = tiles(store)?;
    let tile_key = u64::from(position.row()) / size;
    let local =
        u32::try_from(u64::from(position.row()) % size).map_err(|_| Error::InvalidSource)?;
    let (_, tile) = tiles
        .iter()
        .find(|(candidate, _)| u64::from(*candidate) == tile_key)
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let limits = litchi_numbers_wire::Limits::DEFAULT;
    let encoded = cell.encode(&limits).map_err(|_| Error::InvalidSource)?;
    let (tile_index, tile_data) = message(working, *tile, &[TILE_MESSAGE_TYPE])?;
    let written = litchi_numbers_wire::table_write::tile_set_cell_in_field_order(
        tile_data,
        local,
        position.column(),
        columns,
        Some(&encoded),
        &limits,
    )
    .map_err(|error| match error {
        litchi_numbers_wire::table_write::WriteError::Unsupported => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        litchi_numbers_wire::table_write::WriteError::Allocation { requested } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    })?;
    replace_message(working, *tile, tile_index, written.tile)?;
    // The row's header counts the new cell.
    let headers = single_bytes(store, 1)?.ok_or(Error::InvalidSource)?;
    let bucket = usize::try_from(
        u64::from(position.row())
            / litchi_numbers_wire::table_topology::storage::HEADER_BUCKET_ROWS,
    )
    .map_err(|_| Error::InvalidSource)?;
    let reference = *repeated_bytes(headers, 2)?
        .get(bucket)
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let bucket = references::reference_identifier(reference)?;
    let (bucket_index, bucket_data) = message(working, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
    let adjusted =
        litchi_numbers_wire::table_write::header_adjust(bucket_data, position.row(), 1, &limits)
            .map_err(|_| Error::InvalidSource)?;
    replace_message(working, bucket, bucket_index, adjusted)?;
    // So does its column's header.
    let column_bucket =
        single_reference(store, 2)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let (column_index, column_data) =
        message(working, column_bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
    let adjusted =
        litchi_numbers_wire::table_write::header_adjust(column_data, position.column(), 1, &limits)
            .map_err(|_| Error::InvalidSource)?;
    replace_message(working, column_bucket, column_index, adjusted).map_err(Error::from)
}

/// Rewrite the stored cell `located` at `position` with `update` applied to
/// its parsed BNC record.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed cell or tile row.
pub fn rewrite_cell<P: TableLeaf>(
    working: &mut Working<'_, P>,
    located: &Located,
    position: At,
    update: impl FnOnce(&mut litchi_numbers_wire::BncCell),
) -> Result<(), Error> {
    let limits = cell_limits(&located.bytes);
    let mut cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &limits)
        .map_err(|_error| Error::InvalidSource)?;
    update(&mut cell);
    let encoded = cell.encode(&limits).map_err(|_| Error::InvalidSource)?;
    let (_, tile_data) = message(working, located.tile, &[TILE_MESSAGE_TYPE])?;
    let mut field_index = 0_usize;
    let rewritten = rebuild(tile_data, |field| {
        let this = field_index;
        field_index += 1;
        if this != located.row_field {
            return Ok(None);
        }
        let row = litchi_numbers_wire::table_topology::storage::replace_cell(
            field.payload,
            position.column(),
            &encoded,
        )
        .map_err(Fault::from)?
        .ok_or(Fault::Invalid)?;
        bytes_field(5, &row).map(Some)
    })?;
    replace_message(working, located.tile, located.tile_index, rewritten)?;
    Ok(())
}
