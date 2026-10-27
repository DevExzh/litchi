//! Cell fill and text layout of one table cell, through the copy-on-write
//! cell-style graph (record 0371; the migration host's `cell_style`,
//! `cell_fill` and `cell_layout`).
//!
//! A cell's appearance is the `TST.CellStyleArchive` its BNC style key names
//! in the table's style `TST.TableDataList`, or, without a key, the model's
//! body, header-row, header-column or footer-row style. Properties are
//! inherited along the style's parents. Setting a property gives the cell a
//! private style variation (a child of its current style, registered in the
//! stylesheet) or rewrites the cell's own variation in place when the cell
//! is its only user and the variation holds only supported overrides.
//! Resetting removes the override, dropping a variation that becomes empty.
//!
//! The engine reads through a [`TableSource`] and returns a [`Plan`]: the
//! payload rewrites, the new or removed style object and the registry
//! changes. The leaf applies the plan, allocates nothing itself beyond the
//! identifier the caller supplies, and owns publication.

use std::collections::HashSet;
use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::table::cell::layout::{Inset, Insets, TextWrap, VerticalAlignment};
use litchi_iwa_common::wire::{
    WireView, append_repeated_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
};
use litchi_numbers_wire::{BncCell, Limits, table_write};

use crate::TableSource;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const DATA_LIST_MESSAGE_TYPES: [u32; 2] = [6_005, 6_201];
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// The `TST.CellStyleArchive` versions the host writes.
pub const CELL_STYLE_VERSIONS: [u32; 3] = [1, 0, 5];
const STYLE_LIST_TYPE: u64 = 4;
const HEADER_BUCKET_ROWS: u32 = 256;
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

// `TST.TableModelArchive`.
const MODEL_DATA_STORE: u32 = 4;
const MODEL_ROWS: u32 = 6;
const MODEL_COLUMNS: u32 = 7;
const MODEL_HEADER_ROWS: u32 = 9;
const MODEL_HEADER_COLUMNS: u32 = 10;
const MODEL_FOOTER_ROWS: u32 = 11;
const MODEL_BODY_STYLE: u32 = 18;
const MODEL_HEADER_ROW_STYLE: u32 = 19;
const MODEL_HEADER_COLUMN_STYLE: u32 = 20;
const MODEL_FOOTER_ROW_STYLE: u32 = 21;
// `TST.DataStore`.
const STORE_ROW_HEADERS: u32 = 1;
const STORE_COLUMN_HEADERS: u32 = 2;
const STORE_TILES: u32 = 3;
const STORE_STYLE_TABLE: u32 = 5;
// `TST.TableDataList`.
const LIST_TYPE: u32 = 1;
const LIST_NEXT_ID: u32 = 2;
const LIST_ENTRIES: u32 = 3;
const LIST_SEGMENTS: u32 = 4;
const ENTRY_KEY: u32 = 1;
const ENTRY_REFCOUNT: u32 = 2;
const ENTRY_REFERENCE: u32 = 4;
// `TST.CellStyleArchive` and `TSS.StyleArchive`.
const STYLE_SUPER: u32 = 1;
const STYLE_OVERRIDE_COUNT: u32 = 10;
const STYLE_CELL_PROPERTIES: u32 = 11;
const STYLE_NAME: u32 = 1;
const STYLE_IDENTIFIER: u32 = 2;
const STYLE_PARENT: u32 = 3;
const STYLE_IS_VARIATION: u32 = 4;
const STYLE_STYLESHEET: u32 = 5;
// `TST.CellStylePropertiesArchive`.
const PROPERTY_FILL: u32 = 1;
const PROPERTY_TEXT_WRAP: u32 = 3;
const PROPERTY_VERTICAL_ALIGNMENT: u32 = 8;
const PROPERTY_PADDING: u32 = 9;
// `TSS.StylesheetArchive`.
const STYLESHEET_STYLES: u32 = 1;
const STYLESHEET_CHILDREN: u32 = 5;
const CHILDREN_PARENT: u32 = 1;
const CHILDREN_CHILDREN: u32 = 2;

/// One table cell of a resolved table model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Target {
    /// Object identifier of the `TST.TableModelArchive`.
    pub model_identifier: u64,
    /// Zero-based row.
    pub row: u32,
    /// Zero-based column.
    pub column: u32,
}

/// One cell-style property.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// The cell fill.
    Fill,
    /// Text wrapping.
    TextWrap,
    /// Vertical text alignment.
    VerticalAlignment,
    /// Text insets.
    Padding,
}

impl Kind {
    const fn field(self) -> u32 {
        match self {
            Self::Fill => PROPERTY_FILL,
            Self::TextWrap => PROPERTY_TEXT_WRAP,
            Self::VerticalAlignment => PROPERTY_VERTICAL_ALIGNMENT,
            Self::Padding => PROPERTY_PADDING,
        }
    }

    /// The value a cell shows when no style in its chain sets the property.
    #[must_use]
    pub const fn default_property(self) -> Property {
        match self {
            Self::Fill => Property::Fill(Fill::None),
            Self::TextWrap => Property::TextWrap(TextWrap::Unwrapped),
            Self::VerticalAlignment => Property::VerticalAlignment(VerticalAlignment::Top),
            Self::Padding => Property::Padding(Insets::ZERO),
        }
    }
}

/// One cell-style property value.
#[derive(Debug, Clone, PartialEq)]
pub enum Property {
    /// The cell fill.
    Fill(Fill),
    /// Text wrapping.
    TextWrap(TextWrap),
    /// Vertical text alignment.
    VerticalAlignment(VerticalAlignment),
    /// Text insets.
    Padding(Insets),
}

impl Property {
    /// The property's kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Fill(_) => Kind::Fill,
            Self::TextWrap(_) => Kind::TextWrap,
            Self::VerticalAlignment(_) => Kind::VerticalAlignment,
            Self::Padding(_) => Kind::Padding,
        }
    }

    fn encode(&self, output: &mut Vec<u8>) {
        match self {
            Self::Fill(fill) => {
                push_message(
                    output,
                    PROPERTY_FILL,
                    &litchi_iwa_style::drawing::encode_fill(fill),
                );
            },
            Self::TextWrap(wrap) => {
                push_varint_field(
                    output,
                    PROPERTY_TEXT_WRAP,
                    u64::from(matches!(wrap, TextWrap::Wrapped)),
                );
            },
            Self::VerticalAlignment(alignment) => {
                push_varint_field(
                    output,
                    PROPERTY_VERTICAL_ALIGNMENT,
                    match alignment {
                        VerticalAlignment::Top => 0,
                        VerticalAlignment::Middle => 1,
                        VerticalAlignment::Bottom => 2,
                    },
                );
            },
            Self::Padding(insets) => {
                let mut padding = Vec::with_capacity(20);
                for (number, inset) in [
                    (1, insets.left()),
                    (2, insets.top()),
                    (3, insets.right()),
                    (4, insets.bottom()),
                ] {
                    push_key(&mut padding, number, 5);
                    padding.extend_from_slice(&inset.points().to_le_bytes());
                }
                push_message(output, PROPERTY_PADDING, &padding);
            },
        }
    }
}

/// Why a cell-appearance plan was refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Refusal {
    /// The cell lies outside the table.
    OutsideTable,
    /// The cell's row has no tile (the host creates one; the engine does not).
    MissingTile,
    /// The style list is segmented or its entries are not plain style
    /// references.
    UnsupportedStyleList,
    /// An image fill (data-reference bookkeeping is not ported).
    ImageFill,
    /// A legacy or non-modern cell record.
    UnsupportedCell,
    /// The same property kind was given twice.
    RepeatedProperty,
}

/// A cell-appearance failure.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The source graph is malformed or inconsistent.
    InvalidSource(&'static str),
    /// The edit is refused.
    Refused(Refusal),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSource(reason) => write!(formatter, "invalid table cell style: {reason}"),
            Self::Refused(refusal) => {
                write!(formatter, "table cell style edit refused: {refusal:?}")
            },
        }
    }
}

impl std::error::Error for Error {}

type Result<T> = std::result::Result<T, Error>;

const fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource(reason)
}

fn wire<E>(reason: &'static str) -> impl FnOnce(E) -> Error {
    move |_error| Error::InvalidSource(reason)
}

/// A message payload replacement on an existing object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    /// The object.
    pub identifier: u64,
    /// The message type to replace (the object's unique message of it).
    pub message_type: u32,
    /// The new payload.
    pub data: Vec<u8>,
    /// Object references to add to the message's header.
    pub add_references: Vec<u64>,
    /// Object references to drop from the message's header (message and
    /// field references).
    pub remove_references: Vec<u64>,
}

/// A new cell-style variation object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewStyle {
    /// Its identifier (the caller's).
    pub identifier: u64,
    /// The stylesheet object whose component receives the object (appended).
    pub stylesheet: u64,
    /// The `TST.CellStyleArchive` payload.
    pub data: Vec<u8>,
    /// Its message's object references (the parent style).
    pub object_references: Vec<u64>,
}

/// One package-metadata registry change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Registration {
    /// Register a UUID for the new object in its component.
    Uuid {
        /// The new object.
        object: u64,
    },
    /// Register a cross-component reference from the component of `from` to
    /// `object` (when their components differ).
    ExternalReference {
        /// An object of the referring component.
        from: u64,
        /// The referenced object.
        object: u64,
    },
    /// Raise the issued-identifier watermark to `0`.
    Watermark(u64),
    /// Forget a removed object: its UUID, every external reference to it and
    /// its identifier-suffix reservation.
    Removed {
        /// The removed object.
        object: u64,
    },
}

/// What one cell-appearance edit changes.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Plan {
    /// Payload rewrites, in application order (one per object and type; a
    /// later rewrite of the same message replaces an earlier one).
    pub rewrites: Vec<Rewrite>,
    /// A new style variation.
    pub insertion: Option<NewStyle>,
    /// A style variation to delete (with its stylesheet entries, already in
    /// [`Self::rewrites`]).
    pub removal: Option<u64>,
    /// Registry changes, in host order.
    pub registrations: Vec<Registration>,
}

// ----------------------------------------------------------------------------
// Wire helpers

fn view(data: &[u8]) -> Result<WireView<'_>> {
    WireView::parse(data, WireLimits::DEFAULT).map_err(wire("malformed payload"))
}

fn varint(data: &[u8], number: u32) -> Result<Option<u64>> {
    let mut found = None;
    for field in view(data)?.fields() {
        if field.number() == number {
            if field.wire_type() != 0 {
                return Err(invalid("a varint field has another wire type"));
            }
            let (value, _) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                .map_err(wire("malformed varint"))?;
            found = Some(value);
        }
    }
    Ok(found)
}

fn messages(data: &[u8], number: u32) -> Result<Vec<&[u8]>> {
    repeated_length_delimited_payloads(data, number).map_err(wire("malformed message field"))
}

fn message(data: &[u8], number: u32) -> Result<Option<&[u8]>> {
    let found = messages(data, number)?;
    match found.as_slice() {
        [] => Ok(None),
        [one] => Ok(Some(one)),
        _ => Err(invalid("a singular message repeats")),
    }
}

fn reference(data: &[u8], number: u32) -> Result<Option<u64>> {
    message(data, number)?
        .map(|payload| varint(payload, 1)?.ok_or(invalid("a reference has no identifier")))
        .transpose()
}

fn field_numbers(data: &[u8]) -> Result<Vec<u32>> {
    let mut numbers = view(data)?
        .fields()
        .map(|field| field.number())
        .collect::<Vec<_>>();
    numbers.sort_unstable();
    Ok(numbers)
}

fn push_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    litchi_iwa_common::varint::encode_varint_into(
        output,
        (u64::from(number) << 3) | u64::from(wire_type),
    );
}

fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_key(output, number, 0);
    litchi_iwa_common::varint::encode_varint_into(output, value);
}

fn push_message(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_key(output, number, 2);
    litchi_iwa_common::varint::encode_varint_into(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut output = Vec::with_capacity(10);
    push_varint_field(&mut output, 1, identifier);
    output
}

// ----------------------------------------------------------------------------
// Source access

fn unique_payload<'a>(
    source: &'a dyn TableSource,
    identifier: u64,
    types: &[u32],
) -> Result<(u32, &'a [u8])> {
    let resolved = source
        .resolve_object(identifier)
        .map_err(wire("inconsistent object index"))?
        .ok_or(invalid("a referenced object is missing"))?;
    let mut found = None;
    for message in resolved.messages {
        if types.contains(&message.type_) {
            if found.is_some() {
                return Err(invalid("an object repeats its payload"));
            }
            found = Some((message.type_, message.data.as_slice()));
        }
    }
    found.ok_or(invalid("an object lacks its payload"))
}

fn component_of(source: &dyn TableSource, identifier: u64) -> Result<usize> {
    Ok(source
        .resolve_object(identifier)
        .map_err(wire("inconsistent object index"))?
        .ok_or(invalid("a referenced object is missing"))?
        .component_index)
}

/// The facts of one table the engine needs.
struct Table {
    rows: u32,
    columns: u32,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
    base: [u64; 4],
    style_list: u64,
    tile_size: u64,
    tiles: Vec<(u64, u64)>,
    row_buckets: Vec<u64>,
    column_bucket: Option<u64>,
}

impl Table {
    fn read(source: &dyn TableSource, model: u64) -> Result<Self> {
        let (_, payload) = unique_payload(source, model, &[TABLE_MODEL_MESSAGE_TYPE])?;
        let count = |number: u32| -> Result<u32> {
            u32::try_from(varint(payload, number)?.unwrap_or(0))
                .map_err(|_error| invalid("a table count exceeds u32"))
        };
        let style = |number: u32| -> Result<u64> {
            reference(payload, number)?.ok_or(invalid("a table base style is missing"))
        };
        let store = message(payload, MODEL_DATA_STORE)?.ok_or(invalid("no data store"))?;
        let tiles_payload = message(store, STORE_TILES)?.ok_or(invalid("no tile storage"))?;
        let mut tiles = Vec::new();
        for entry in messages(tiles_payload, 1)? {
            let key = varint(entry, 1)?.ok_or(invalid("a tile has no key"))?;
            let tile = reference(entry, 2)?.ok_or(invalid("a tile has no object"))?;
            tiles.push((key, tile));
        }
        let mut row_buckets = Vec::new();
        if let Some(headers) = message(store, STORE_ROW_HEADERS)? {
            for bucket in messages(headers, 2)? {
                row_buckets.push(varint(bucket, 1)?.ok_or(invalid("a bucket has no object"))?);
            }
        }
        Ok(Self {
            rows: count(MODEL_ROWS)?,
            columns: count(MODEL_COLUMNS)?,
            header_rows: count(MODEL_HEADER_ROWS)?,
            header_columns: count(MODEL_HEADER_COLUMNS)?,
            footer_rows: count(MODEL_FOOTER_ROWS)?,
            base: [
                style(MODEL_HEADER_ROW_STYLE)?,
                style(MODEL_FOOTER_ROW_STYLE)?,
                style(MODEL_HEADER_COLUMN_STYLE)?,
                style(MODEL_BODY_STYLE)?,
            ],
            style_list: reference(store, STORE_STYLE_TABLE)?.ok_or(invalid("no style list"))?,
            tile_size: varint(tiles_payload, 2)?.unwrap_or(256).max(1),
            tiles,
            row_buckets,
            column_bucket: reference(store, STORE_COLUMN_HEADERS)?,
        })
    }

    fn check(&self, target: Target) -> Result<()> {
        if target.row >= self.rows || target.column >= self.columns {
            return Err(Error::Refused(Refusal::OutsideTable));
        }
        Ok(())
    }

    /// The model's style for a cell without a key: header row, footer row,
    /// header column, then body.
    const fn base_style(&self, row: u32, column: u32) -> u64 {
        if row < self.header_rows {
            self.base[0]
        } else if row >= self.rows.saturating_sub(self.footer_rows) {
            self.base[1]
        } else if column < self.header_columns {
            self.base[2]
        } else {
            self.base[3]
        }
    }

    fn tile(&self, row: u32) -> Option<(u64, u32)> {
        let key = u64::from(row) / self.tile_size;
        let local = u32::try_from(u64::from(row) % self.tile_size).ok()?;
        self.tiles
            .iter()
            .find(|(candidate, _)| *candidate == key)
            .map(|(_, tile)| (*tile, local))
    }
}

/// The stored cell bytes, if any.
fn stored_cell(
    source: &dyn TableSource,
    table: &Table,
    target: Target,
    limits: &Limits,
) -> Result<Option<Vec<u8>>> {
    let Some((tile, local)) = table.tile(target.row) else {
        return Ok(None);
    };
    let (_, payload) = unique_payload(source, tile, &[TILE_MESSAGE_TYPE])?;
    table_write::tile_cell(payload, local, target.column, limits)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))
}

fn style_key(cell: Option<&[u8]>, limits: &Limits) -> Result<Option<u32>> {
    cell.map(|data| {
        BncCell::parse(data, limits)
            .map(|cell| cell.style_identifier())
            .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))
    })
    .transpose()
    .map(Option::flatten)
}

/// One style-list entry.
#[derive(Debug, Clone, Copy)]
struct Entry {
    key: u32,
    refcount: u32,
    style: u64,
}

struct StyleList<'a> {
    identifier: u64,
    message_type: u32,
    payload: &'a [u8],
    next: u32,
    entries: Vec<Entry>,
}

impl<'a> StyleList<'a> {
    fn read(source: &'a dyn TableSource, identifier: u64) -> Result<Self> {
        let (message_type, payload) = unique_payload(source, identifier, &DATA_LIST_MESSAGE_TYPES)?;
        if varint(payload, LIST_TYPE)? != Some(STYLE_LIST_TYPE)
            || !messages(payload, LIST_SEGMENTS)?.is_empty()
        {
            return Err(Error::Refused(Refusal::UnsupportedStyleList));
        }
        let next = u32::try_from(varint(payload, LIST_NEXT_ID)?.ok_or(invalid("no next list id"))?)
            .map_err(|_error| invalid("next list id exceeds u32"))?;
        let mut entries = Vec::new();
        for raw in messages(payload, LIST_ENTRIES)? {
            let numbers = field_numbers(raw)?;
            if numbers != [ENTRY_KEY, ENTRY_REFCOUNT, ENTRY_REFERENCE] {
                return Err(Error::Refused(Refusal::UnsupportedStyleList));
            }
            let entry = Entry {
                key: u32::try_from(varint(raw, ENTRY_KEY)?.ok_or(invalid("an entry has no key"))?)
                    .map_err(|_error| invalid("a key exceeds u32"))?,
                refcount: u32::try_from(
                    varint(raw, ENTRY_REFCOUNT)?.ok_or(invalid("an entry has no refcount"))?,
                )
                .map_err(|_error| invalid("a refcount exceeds u32"))?,
                style: reference(raw, ENTRY_REFERENCE)?.ok_or(invalid("an entry has no style"))?,
            };
            if entries.iter().any(|other: &Entry| other.key == entry.key) {
                return Err(invalid("the style list repeats a key"));
            }
            entries.push(entry);
        }
        Ok(Self {
            identifier,
            message_type,
            payload,
            next,
            entries,
        })
    }

    fn entry(&self, key: u32) -> Result<Entry> {
        self.entries
            .iter()
            .find(|entry| entry.key == key)
            .copied()
            .ok_or(invalid("a cell names a missing style key"))
    }

    /// The host's allocation: one above every key and at least the cursor.
    fn next_key(&self) -> Result<u32> {
        let after = self
            .entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .map_or(Some(1), |key| key.checked_add(1))
            .ok_or(invalid("style key overflow"))?;
        Ok(self.next.max(after).max(1))
    }
}

/// Rewrite a style list's entries and cursor; `edit` maps each entry to its
/// new refcount (`None` removes it); `append` adds an entry.
fn rewrite_list(
    payload: &[u8],
    entries: &[Entry],
    next: u32,
    change: Option<(u32, Option<u32>)>,
    append: Option<Entry>,
) -> Result<(Vec<u8>, Option<u64>)> {
    let mut data = patch_varint_field(payload, LIST_NEXT_ID, true, Some(u64::from(next)))
        .map_err(wire("style list cursor"))?;
    let raw = messages(&data, LIST_ENTRIES)?
        .into_iter()
        .map(<[u8]>::to_vec)
        .collect::<Vec<_>>();
    let mut replacements = Vec::with_capacity(raw.len() + 1);
    let mut removed = None;
    for (entry, bytes) in entries.iter().zip(raw) {
        match change {
            Some((key, None)) if key == entry.key => {
                removed = Some(entry.style);
            },
            Some((key, Some(refcount))) if key == entry.key => {
                replacements.push(
                    patch_varint_field(&bytes, ENTRY_REFCOUNT, true, Some(u64::from(refcount)))
                        .map_err(wire("style list refcount"))?,
                );
            },
            _ => replacements.push(bytes),
        }
    }
    if let Some(entry) = append {
        let mut bytes = Vec::new();
        push_varint_field(&mut bytes, ENTRY_KEY, u64::from(entry.key));
        push_varint_field(&mut bytes, ENTRY_REFCOUNT, u64::from(entry.refcount));
        push_message(&mut bytes, ENTRY_REFERENCE, &reference_payload(entry.style));
        replacements.push(bytes);
    }
    data = rewrite_repeated_length_delimited_fields(&data, LIST_ENTRIES, &replacements)
        .map_err(wire("style list entries"))?;
    Ok((data, removed))
}

/// One located cell style.
struct Style<'a> {
    payload: &'a [u8],
    parent: Option<u64>,
    stylesheet: Option<u64>,
}

fn style<'a>(source: &'a dyn TableSource, identifier: u64) -> Result<Style<'a>> {
    let (_, payload) = unique_payload(source, identifier, &[CELL_STYLE_MESSAGE_TYPE])?;
    let base = message(payload, STYLE_SUPER)?.ok_or(invalid("a cell style has no base"))?;
    Ok(Style {
        payload,
        parent: reference(base, STYLE_PARENT)?,
        stylesheet: reference(base, STYLE_STYLESHEET)?,
    })
}

fn decode_property(properties: &[u8], kind: Kind) -> Result<Option<Property>> {
    Ok(match kind {
        Kind::Fill => message(properties, PROPERTY_FILL)?
            .map(|fill| {
                let stored = litchi_iwa_protos::shape_style_codec::native::fill(
                    fill,
                    litchi_iwa_protos::chart_style_codec::Limits::DEFAULT,
                )
                .map_err(wire("malformed cell fill"))?;
                litchi_iwa_style::drawing::fill(&stored)
                    .map(Property::Fill)
                    .map_err(wire("unsupported cell fill"))
            })
            .transpose()?,
        Kind::TextWrap => varint(properties, PROPERTY_TEXT_WRAP)?.map(|value| {
            Property::TextWrap(if value != 0 {
                TextWrap::Wrapped
            } else {
                TextWrap::Unwrapped
            })
        }),
        Kind::VerticalAlignment => varint(properties, PROPERTY_VERTICAL_ALIGNMENT)?
            .map(|value| {
                Ok(Property::VerticalAlignment(match value {
                    0 => VerticalAlignment::Top,
                    1 => VerticalAlignment::Middle,
                    2 => VerticalAlignment::Bottom,
                    _ => return Err(invalid("unknown vertical alignment")),
                }))
            })
            .transpose()?,
        Kind::Padding => message(properties, PROPERTY_PADDING)?
            .map(|padding| {
                let mut values = [0.0_f32; 4];
                for field in view(padding)?.fields() {
                    let slot = match field.number() {
                        1..=4 => &mut values[field.number() as usize - 1],
                        _ => continue,
                    };
                    let bytes: [u8; 4] = field
                        .payload()
                        .try_into()
                        .map_err(|_error| invalid("a padding value is not a float"))?;
                    *slot = f32::from_le_bytes(bytes);
                }
                let inset = |value: f32| {
                    Inset::from_points(value).map_err(|_error| invalid("invalid padding"))
                };
                Ok(Property::Padding(Insets::new(
                    inset(values[0])?,
                    inset(values[1])?,
                    inset(values[2])?,
                    inset(values[3])?,
                )))
            })
            .transpose()?,
    })
}

fn direct_property(style: &Style<'_>, kind: Kind) -> Result<Option<Property>> {
    match message(style.payload, STYLE_CELL_PROPERTIES)? {
        Some(properties) => decode_property(properties, kind),
        None => Ok(None),
    }
}

fn effective_from(source: &dyn TableSource, first: u64, kind: Kind) -> Result<Option<Property>> {
    let mut identifier = first;
    let mut seen = HashSet::new();
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        if !seen.insert(identifier) {
            return Err(invalid("cell-style inheritance cycles"));
        }
        let located = style(source, identifier)?;
        if let Some(property) = direct_property(&located, kind)? {
            return Ok(Some(property));
        }
        let Some(parent) = located.parent else {
            return Ok(None);
        };
        identifier = parent;
    }
    Err(invalid("cell-style inheritance is too deep"))
}

/// The supported overrides of a private variation the engine may rewrite in
/// place: a parented, unnamed variation in a stylesheet whose payload holds
/// exactly its base, override count and the four supported properties.
fn direct_supported(style: &Style<'_>) -> Result<Option<Vec<Property>>> {
    let Some(properties) = message(style.payload, STYLE_CELL_PROPERTIES)? else {
        return Ok(None);
    };
    let mut overrides = Vec::new();
    for kind in [
        Kind::Fill,
        Kind::TextWrap,
        Kind::VerticalAlignment,
        Kind::Padding,
    ] {
        if let Some(property) = decode_property(properties, kind)? {
            overrides.push(property);
        }
    }
    if overrides.is_empty() {
        return Ok(None);
    }
    let base = message(style.payload, STYLE_SUPER)?.ok_or(invalid("a cell style has no base"))?;
    let count = varint(style.payload, STYLE_OVERRIDE_COUNT)?;
    let semantic = count == Some(overrides.len() as u64)
        && message(base, STYLE_NAME)?.is_none()
        && message(base, STYLE_IDENTIFIER)?.is_none()
        && style.parent.is_some()
        && varint(base, STYLE_IS_VARIATION)? == Some(1)
        && style.stylesheet.is_some();
    if !semantic {
        return Ok(None);
    }
    let mut expected = overrides
        .iter()
        .map(|property| property.kind().field())
        .collect::<Vec<_>>();
    expected.sort_unstable();
    let exact = field_numbers(style.payload)?
        == [STYLE_SUPER, STYLE_OVERRIDE_COUNT, STYLE_CELL_PROPERTIES]
        && field_numbers(base)? == [STYLE_PARENT, STYLE_IS_VARIATION, STYLE_STYLESHEET]
        && field_numbers(properties)? == expected;
    Ok(exact.then_some(overrides))
}

fn apply(overrides: &mut Vec<Property>, property: &Property) {
    if let Some(slot) = overrides
        .iter_mut()
        .find(|current| current.kind() == property.kind())
    {
        *slot = property.clone();
    } else {
        overrides.push(property.clone());
    }
}

/// The host's variation payload: base (parent, variation flag, stylesheet),
/// override count, properties in field order.
fn variation_payload(parent: u64, stylesheet: u64, overrides: &[Property]) -> Result<Vec<u8>> {
    if overrides.is_empty() {
        return Err(invalid("a variation needs an override"));
    }
    let mut base = Vec::new();
    push_message(&mut base, STYLE_PARENT, &reference_payload(parent));
    push_varint_field(&mut base, STYLE_IS_VARIATION, 1);
    push_message(&mut base, STYLE_STYLESHEET, &reference_payload(stylesheet));
    let mut ordered = overrides.to_vec();
    ordered.sort_by_key(|property| property.kind().field());
    let mut properties = Vec::new();
    for property in &ordered {
        property.encode(&mut properties);
    }
    let mut output = Vec::new();
    push_message(&mut output, STYLE_SUPER, &base);
    push_varint_field(&mut output, STYLE_OVERRIDE_COUNT, ordered.len() as u64);
    push_message(&mut output, STYLE_CELL_PROPERTIES, &properties);
    Ok(output)
}

fn refuse_image(property: &Property) -> Result<()> {
    if matches!(property, Property::Fill(Fill::Image(_))) {
        return Err(Error::Refused(Refusal::ImageFill));
    }
    Ok(())
}

fn refuse_images(overrides: &[Property]) -> Result<()> {
    overrides.iter().try_for_each(refuse_image)
}

fn equal(left: Option<&Property>, right: Option<&Property>, kind: Kind) -> bool {
    let default = kind.default_property();
    left.unwrap_or(&default) == right.unwrap_or(&default)
}

// ----------------------------------------------------------------------------
// Stylesheet entries

fn stylesheet_append(data: &[u8], parent: u64, style: u64) -> Result<Vec<u8>> {
    for existing in messages(data, STYLESHEET_STYLES)? {
        if varint(existing, 1)? == Some(style) {
            return Err(invalid("the stylesheet already holds the new style"));
        }
    }
    let mut data =
        append_repeated_length_delimited_field(data, STYLESHEET_STYLES, &reference_payload(style))
            .map_err(wire("stylesheet styles"))?;
    let entries = messages(&data, STYLESHEET_CHILDREN)?
        .into_iter()
        .map(<[u8]>::to_vec)
        .collect::<Vec<_>>();
    let mut matches = 0;
    let mut rewritten = Vec::with_capacity(entries.len() + 1);
    for entry in entries {
        if reference(&entry, CHILDREN_PARENT)? == Some(parent) {
            matches += 1;
            rewritten.push(
                append_repeated_length_delimited_field(
                    &entry,
                    CHILDREN_CHILDREN,
                    &reference_payload(style),
                )
                .map_err(wire("stylesheet children"))?,
            );
        } else {
            rewritten.push(entry);
        }
    }
    if matches > 1 {
        return Err(invalid("the stylesheet repeats a parent"));
    }
    if matches == 0 {
        let mut entry = Vec::new();
        push_message(&mut entry, CHILDREN_PARENT, &reference_payload(parent));
        push_message(&mut entry, CHILDREN_CHILDREN, &reference_payload(style));
        rewritten.push(entry);
    }
    data = rewrite_repeated_length_delimited_fields(&data, STYLESHEET_CHILDREN, &rewritten)
        .map_err(wire("stylesheet children"))?;
    Ok(data)
}

fn stylesheet_remove(data: &[u8], parent: u64, style: u64) -> Result<Vec<u8>> {
    let styles = messages(data, STYLESHEET_STYLES)?;
    let mut kept = Vec::with_capacity(styles.len());
    let mut found = 0;
    for entry in styles {
        if varint(entry, 1)? == Some(style) {
            found += 1;
        } else {
            kept.push(entry.to_vec());
        }
    }
    if found != 1 {
        return Err(invalid(
            "the stylesheet does not hold the removed style once",
        ));
    }
    let data = rewrite_repeated_length_delimited_fields(data, STYLESHEET_STYLES, &kept)
        .map_err(wire("stylesheet styles"))?;
    let mut removed = 0;
    let mut entries = Vec::new();
    for payload in messages(&data, STYLESHEET_CHILDREN)? {
        let entry_parent = reference(payload, CHILDREN_PARENT)?;
        let children = messages(payload, CHILDREN_CHILDREN)?;
        let mut kept = Vec::with_capacity(children.len());
        for child in children {
            if varint(child, 1)? == Some(style) {
                if entry_parent != Some(parent) {
                    return Err(invalid(
                        "the stylesheet maps the style under another parent",
                    ));
                }
                removed += 1;
            } else {
                kept.push(child.to_vec());
            }
        }
        if !kept.is_empty() {
            entries.push(
                rewrite_repeated_length_delimited_fields(payload, CHILDREN_CHILDREN, &kept)
                    .map_err(wire("stylesheet children"))?,
            );
        }
    }
    if removed != 1 {
        return Err(invalid(
            "the stylesheet child map does not hold the style once",
        ));
    }
    rewrite_repeated_length_delimited_fields(&data, STYLESHEET_CHILDREN, &entries)
        .map_err(wire("stylesheet children"))
}

fn has_children(source: &dyn TableSource, style: u64) -> Result<bool> {
    let component = component_of(source, style)?;
    let Some(component) = source.catalog().get_index(component) else {
        return Err(invalid("a style component is missing"));
    };
    for object in &component.archive().objects {
        for message_entry in &object.messages {
            if message_entry.type_ != CELL_STYLE_MESSAGE_TYPE {
                continue;
            }
            if let Some(base) = message(&message_entry.data, STYLE_SUPER)?
                && reference(base, STYLE_PARENT)? == Some(style)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

// ----------------------------------------------------------------------------
// Cell writes

/// The rewrites that give the cell `key` (`None` drops the key).
fn write_key(
    source: &dyn TableSource,
    table: &Table,
    target: Target,
    key: Option<u32>,
    limits: &Limits,
) -> Result<Vec<Rewrite>> {
    let (tile, local) = table
        .tile(target.row)
        .ok_or(Error::Refused(Refusal::MissingTile))?;
    let (tile_type, tile_payload) = unique_payload(source, tile, &[TILE_MESSAGE_TYPE])?;
    let previous = table_write::tile_cell(tile_payload, local, target.column, limits)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?;
    let mut cell = match previous.as_deref() {
        Some(data) => BncCell::parse(data, limits)
            .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?,
        None => BncCell::minimal(),
    };
    cell.set_style_identifier(key);
    let encoded = cell
        .encode(limits)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?;
    let written = table_write::tile_set_cell(
        tile_payload,
        local,
        target.column,
        table.columns,
        Some(&encoded),
        limits,
    )
    .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?;
    let mut rewrites = vec![Rewrite {
        identifier: tile,
        message_type: tile_type,
        data: written.tile,
        add_references: Vec::new(),
        remove_references: Vec::new(),
    }];
    if written.previous.is_none() {
        let bucket = *table
            .row_buckets
            .get((target.row / HEADER_BUCKET_ROWS) as usize)
            .ok_or(Error::Refused(Refusal::MissingTile))?;
        let (bucket_type, payload) = unique_payload(source, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        rewrites.push(Rewrite {
            identifier: bucket,
            message_type: bucket_type,
            data: table_write::header_adjust(payload, target.row, 1, limits)
                .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?,
            add_references: Vec::new(),
            remove_references: Vec::new(),
        });
        if let Some(column_bucket) = table.column_bucket {
            let (bucket_type, payload) =
                unique_payload(source, column_bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
            rewrites.push(Rewrite {
                identifier: column_bucket,
                message_type: bucket_type,
                data: table_write::header_adjust(payload, target.column, 1, limits)
                    .map_err(|_error| Error::Refused(Refusal::UnsupportedCell))?,
                add_references: Vec::new(),
                remove_references: Vec::new(),
            });
        }
    }
    Ok(rewrites)
}

// ----------------------------------------------------------------------------
// Public engine

/// The effective value of `kind` on one cell (`None`: no style in the
/// chain sets it; see [`Kind::default_property`]).
///
/// # Errors
///
/// [`Error::Refused`] for a cell outside the table or an unsupported cell or
/// style list; [`Error::InvalidSource`] for a malformed graph.
pub fn effective(
    source: &dyn TableSource,
    target: Target,
    kind: Kind,
    limits: &Limits,
) -> Result<Option<Property>> {
    let table = Table::read(source, target.model_identifier)?;
    table.check(target)?;
    let cell = stored_cell(source, &table, target, limits)?;
    let style_identifier = match style_key(cell.as_deref(), limits)? {
        Some(key) => StyleList::read(source, table.style_list)?.entry(key)?.style,
        None => table.base_style(target.row, target.column),
    };
    effective_from(source, style_identifier, kind)
}

/// The value `kind` shows on one cell, defaults included.
///
/// # Errors
///
/// As [`effective`].
pub fn value(
    source: &dyn TableSource,
    target: Target,
    kind: Kind,
    limits: &Limits,
) -> Result<Property> {
    Ok(effective(source, target, kind, limits)?.unwrap_or_else(|| kind.default_property()))
}

/// Plan setting `properties` on one cell; `fresh` is the identifier a new
/// variation gets. `None` when every property already holds.
///
/// # Errors
///
/// [`Error::Refused`] for repeated kinds, image fills, cells outside the
/// table or without a tile, and unsupported style lists;
/// [`Error::InvalidSource`] for a malformed graph.
pub fn plan_set(
    source: &dyn TableSource,
    target: Target,
    properties: &[Property],
    fresh: u64,
    limits: &Limits,
) -> Result<Option<Plan>> {
    let table = Table::read(source, target.model_identifier)?;
    table.check(target)?;
    let mut kinds = HashSet::new();
    for property in properties {
        if !kinds.insert(property.kind()) {
            return Err(Error::Refused(Refusal::RepeatedProperty));
        }
        refuse_image(property)?;
    }
    let mut changed = false;
    for property in properties {
        let current = effective(source, target, property.kind(), limits)?;
        changed |= !equal(current.as_ref(), Some(property), property.kind());
    }
    if !changed {
        return Ok(None);
    }
    let cell = stored_cell(source, &table, target, limits)?;
    let old_key = style_key(cell.as_deref(), limits)?;
    let list = StyleList::read(source, table.style_list)?;
    let old_entry = old_key.map(|key| list.entry(key)).transpose()?;
    let parent = old_entry.map_or_else(
        || table.base_style(target.row, target.column),
        |entry| entry.style,
    );

    // The cell's own variation is rewritten in place.
    if let Some(entry) = old_entry
        && entry.refcount == 1
    {
        let located = style(source, parent)?;
        if let Some(mut overrides) = direct_supported(&located)? {
            refuse_images(&overrides)?;
            for property in properties {
                apply(&mut overrides, property);
            }
            let data = variation_payload(
                located.parent.ok_or(invalid("a variation has no parent"))?,
                located
                    .stylesheet
                    .ok_or(invalid("a variation has no stylesheet"))?,
                &overrides,
            )?;
            return Ok(Some(Plan {
                rewrites: vec![Rewrite {
                    identifier: parent,
                    message_type: CELL_STYLE_MESSAGE_TYPE,
                    data,
                    add_references: Vec::new(),
                    remove_references: Vec::new(),
                }],
                ..Plan::default()
            }));
        }
    }

    // A new private variation of the cell's current style.
    let located = style(source, parent)?;
    let stylesheet = located
        .stylesheet
        .ok_or(invalid("a cell style has no stylesheet"))?;
    let data = variation_payload(parent, stylesheet, properties)?;
    let mut rewrites = Vec::new();
    let (sheet_type, sheet) = unique_payload(source, stylesheet, &[STYLESHEET_MESSAGE_TYPE])?;
    rewrites.push(Rewrite {
        identifier: stylesheet,
        message_type: sheet_type,
        data: stylesheet_append(sheet, parent, fresh)?,
        add_references: vec![fresh],
        remove_references: Vec::new(),
    });
    let key = list.next_key()?;
    let next = key.checked_add(1).ok_or(invalid("style key overflow"))?;
    let appended = Entry {
        key,
        refcount: 1,
        style: fresh,
    };
    let (mut list_data, _) = rewrite_list(list.payload, &list.entries, next, None, Some(appended))?;
    let mut entries = list.entries.clone();
    entries.push(appended);
    let mut remove_references = Vec::new();
    rewrites.extend(write_key(source, &table, target, Some(key), limits)?);
    if let Some(entry) = old_entry {
        let mut cursor = next;
        let change = if entry.refcount > 1 {
            (entry.key, Some(entry.refcount - 1))
        } else {
            if entry.key.checked_add(1) == Some(cursor) {
                cursor = entry.key;
            }
            (entry.key, None)
        };
        let (data, removed) = rewrite_list(&list_data, &entries, cursor, Some(change), None)?;
        list_data = data;
        if let Some(style) = removed
            && !entries
                .iter()
                .any(|other| other.key != entry.key && other.style == style)
        {
            remove_references.push(style);
        }
    }
    rewrites.push(Rewrite {
        identifier: list.identifier,
        message_type: list.message_type,
        data: list_data,
        add_references: vec![fresh],
        remove_references,
    });
    Ok(Some(Plan {
        rewrites,
        insertion: Some(NewStyle {
            identifier: fresh,
            stylesheet,
            data,
            object_references: vec![parent],
        }),
        removal: None,
        registrations: vec![
            Registration::Uuid { object: fresh },
            Registration::ExternalReference {
                from: list.identifier,
                object: fresh,
            },
            Registration::Watermark(fresh),
        ],
    }))
}

/// Plan resetting `kind` on one cell to the value its parent style gives.
/// `None` when the cell's own style does not set it.
///
/// A cell sharing its style, or whose style holds unsupported overrides,
/// gets a private variation restoring the inherited value (the host's
/// fallback), allocated as `fresh`.
///
/// # Errors
///
/// As [`plan_set`].
pub fn plan_reset(
    source: &dyn TableSource,
    target: Target,
    kind: Kind,
    fresh: u64,
    limits: &Limits,
) -> Result<Option<Plan>> {
    let table = Table::read(source, target.model_identifier)?;
    table.check(target)?;
    let cell = stored_cell(source, &table, target, limits)?;
    let Some(key) = style_key(cell.as_deref(), limits)? else {
        return Ok(None);
    };
    let list = StyleList::read(source, table.style_list)?;
    let entry = list.entry(key)?;
    let located = style(source, entry.style)?;
    let Some(direct) = direct_property(&located, kind)? else {
        return Ok(None);
    };
    refuse_image(&direct)?;
    let base = table.base_style(target.row, target.column);
    let parent = located.parent.unwrap_or(base);
    let inherited = effective_from(source, parent, kind)?;

    if entry.refcount == 1
        && let Some(mut overrides) = direct_supported(&located)?
    {
        refuse_images(&overrides)?;
        overrides.retain(|property| property.kind() != kind);
        if !overrides.is_empty() {
            let data = variation_payload(
                parent,
                located
                    .stylesheet
                    .ok_or(invalid("a variation has no stylesheet"))?,
                &overrides,
            )?;
            return Ok(Some(Plan {
                rewrites: vec![Rewrite {
                    identifier: entry.style,
                    message_type: CELL_STYLE_MESSAGE_TYPE,
                    data,
                    add_references: Vec::new(),
                    remove_references: Vec::new(),
                }],
                ..Plan::default()
            }));
        }
        // The variation becomes empty: the cell falls back to its parent.
        let mut rewrites = Vec::new();
        let mut entries = list.entries.clone();
        let mut list_data = list.payload.to_vec();
        let mut next = list.next;
        let mut add_references = Vec::new();
        if parent == base {
            rewrites.extend(write_key(source, &table, target, None, limits)?);
        } else {
            let parent_key = list.next_key()?;
            next = parent_key
                .checked_add(1)
                .ok_or(invalid("style key overflow"))?;
            let appended = Entry {
                key: parent_key,
                refcount: 1,
                style: parent,
            };
            list_data = rewrite_list(&list_data, &entries, next, None, Some(appended))?.0;
            entries.push(appended);
            add_references.push(parent);
            rewrites.extend(write_key(source, &table, target, Some(parent_key), limits)?);
        }
        if entry.key.checked_add(1) == Some(next) {
            next = entry.key;
        }
        let (data, _) = rewrite_list(&list_data, &entries, next, Some((entry.key, None)), None)?;
        let remove_references = if entries
            .iter()
            .any(|other| other.key != entry.key && other.style == entry.style)
        {
            Vec::new()
        } else {
            vec![entry.style]
        };
        rewrites.push(Rewrite {
            identifier: list.identifier,
            message_type: list.message_type,
            data,
            add_references,
            remove_references,
        });
        let mut registrations = Vec::new();
        let mut removal = None;
        if !has_children(source, entry.style)? {
            let stylesheet = located
                .stylesheet
                .ok_or(invalid("a variation has no stylesheet"))?;
            let (sheet_type, sheet) =
                unique_payload(source, stylesheet, &[STYLESHEET_MESSAGE_TYPE])?;
            rewrites.push(Rewrite {
                identifier: stylesheet,
                message_type: sheet_type,
                data: stylesheet_remove(sheet, parent, entry.style)?,
                add_references: Vec::new(),
                remove_references: vec![entry.style],
            });
            removal = Some(entry.style);
            registrations.push(Registration::Removed {
                object: entry.style,
            });
        }
        return Ok(Some(Plan {
            rewrites,
            insertion: None,
            removal,
            registrations,
        }));
    }

    // A shared or richer style: a private child restores the inherited value.
    plan_set(
        source,
        target,
        &[inherited.unwrap_or_else(|| kind.default_property())],
        fresh,
        limits,
    )
}
