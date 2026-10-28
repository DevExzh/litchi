//! Shared whole-cell character-style graph planning.
//!
//! This module is the format-neutral substrate for the Numbers whole-cell
//! `CharacterProperties` family.  It deliberately stops at a neutral graph
//! plan: the leaf owns identifier allocation, archive publication, metadata
//! transitions, exact-source reassembly and candidate reopen/readback.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;

use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, patch_varint_field, repeated_length_delimited_payloads,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text_wire::CharacterProperties;
use litchi_numbers_wire::{BncCell, Limits as BncLimits, table_topology::storage};

use crate::TableSource;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const LIST_MESSAGE_TYPE: u32 = 6_005;
const LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_201;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLE_LIST_FIELD: u32 = 5;
const PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
const MODEL_DATA_STORE: u32 = 4;
const MODEL_ROWS: u32 = 6;
const MODEL_COLUMNS: u32 = 7;
const MODEL_HEADER_ROWS: u32 = 9;
const MODEL_HEADER_COLUMNS: u32 = 10;
const MODEL_FOOTER_ROWS: u32 = 11;
const STORE_TILES: u32 = 3;
const MAX_STYLE_DEPTH: usize = 16;

/// One table cell selected by a leaf facade.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Target {
    /// The table-model object identifier.
    pub model_identifier: u64,
    /// Zero-based row.
    pub row: u32,
    /// Zero-based column.
    pub column: u32,
}

/// Refusal from the shared character-style graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected cell is outside the table geometry.
    OutsideTable,
    /// The style list is segmented or otherwise not owned by this planner.
    UnsupportedStyleList,
    /// The style graph has an external or ambiguous owner.
    UnsupportedDependency,
    /// The source lacks the required table structure.
    UnsupportedSource,
}

/// A typed, content-redacted shared planner failure.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The graph is malformed or inconsistent.
    InvalidSource(&'static str),
    /// A well-formed source is outside this engine's ownership profile.
    Refused(Refusal),
    /// A bounded wire/resource operation exceeded its caller's profile.
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The plan requires a leaf-supplied fresh identifier.
    NeedsIdentifier,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSource(reason) => {
                write!(formatter, "invalid cell text-style source: {reason}")
            },
            Self::Refused(reason) => write!(formatter, "cell text-style graph refused: {reason:?}"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "{kind} limit exceeded: observed {observed}, maximum {maximum}"
                )
            },
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "could not allocate {amount} cell text-style items"
                )
            },
            Self::NeedsIdentifier => {
                formatter.write_str("cell text-style plan needs a fresh identifier")
            },
        }
    }
}

impl std::error::Error for Error {}

fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource(reason)
}

fn map_wire(error: litchi_iwa_common::Error) -> Error {
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
        _ => invalid("malformed character-style wire payload"),
    }
}

fn view(data: &[u8]) -> Result<WireView<'_>, Error> {
    WireView::parse(data, WireLimits::DEFAULT).map_err(map_wire)
}

fn varint(data: &[u8], number: u32) -> Result<Option<u64>, Error> {
    let mut found = None;
    for field in view(data)?.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 || found.is_some() {
            return Err(invalid(
                "a singular varint field repeats or has the wrong wire type",
            ));
        }
        let (value, _) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
            .map_err(|_| invalid("malformed varint"))?;
        found = Some(value);
    }
    Ok(found)
}

fn bytes(data: &[u8], number: u32) -> Result<Option<&[u8]>, Error> {
    let mut found = None;
    for field in view(data)?.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 || found.is_some() {
            return Err(invalid(
                "a singular bytes field repeats or has the wrong wire type",
            ));
        }
        found = Some(field.payload());
    }
    Ok(found)
}

fn reference(data: &[u8], number: u32) -> Result<Option<u64>, Error> {
    let Some(payload) = bytes(data, number)? else {
        return Ok(None);
    };
    varint(payload, 1)
}

fn unique_payload<'a>(
    source: &'a dyn TableSource,
    identifier: u64,
    types: &[u32],
) -> Result<(u32, &'a [u8]), Error> {
    let object = source
        .resolve_object(identifier)
        .map_err(|_| invalid("object index is inconsistent"))?
        .ok_or(invalid("referenced object is missing"))?;
    let mut found = None;
    for message in object.messages {
        if types.contains(&message.type_) {
            if found.is_some() {
                return Err(invalid("object repeats a required payload"));
            }
            found = Some((message.type_, message.data.as_slice()));
        }
    }
    found.ok_or(invalid("object lacks a required payload"))
}

#[derive(Debug, Clone, Copy)]
struct Geometry {
    rows: u32,
    columns: u32,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
}

impl Geometry {
    fn read(model: &[u8]) -> Result<Self, Error> {
        let value = |field| -> Result<u32, Error> {
            u32::try_from(varint(model, field)?.ok_or(invalid("table geometry field missing"))?)
                .map_err(|_| invalid("table geometry exceeds u32"))
        };
        let geometry = Self {
            rows: value(MODEL_ROWS)?,
            columns: value(MODEL_COLUMNS)?,
            header_rows: u32::try_from(varint(model, MODEL_HEADER_ROWS)?.unwrap_or(0))
                .map_err(|_| invalid("header rows exceed u32"))?,
            header_columns: u32::try_from(varint(model, MODEL_HEADER_COLUMNS)?.unwrap_or(0))
                .map_err(|_| invalid("header columns exceed u32"))?,
            footer_rows: u32::try_from(varint(model, MODEL_FOOTER_ROWS)?.unwrap_or(0))
                .map_err(|_| invalid("footer rows exceed u32"))?,
        };
        if geometry.header_rows + geometry.footer_rows > geometry.rows
            || geometry.header_columns > geometry.columns
        {
            return Err(invalid("table geometry ranges overlap"));
        }
        Ok(geometry)
    }
}

struct Table {
    model: Vec<u8>,
    store: Vec<u8>,
    style_list: u64,
    geometry: Geometry,
}

impl Table {
    fn read(source: &dyn TableSource, target: Target) -> Result<Self, Error> {
        let (_, model) =
            unique_payload(source, target.model_identifier, &[TABLE_MODEL_MESSAGE_TYPE])?;
        let geometry = Geometry::read(model)?;
        if target.row >= geometry.rows || target.column >= geometry.columns {
            return Err(Error::Refused(Refusal::OutsideTable));
        }
        let store = bytes(model, MODEL_DATA_STORE)?.ok_or(invalid("table data store missing"))?;
        let store = store.to_vec();
        let style_list = reference(&store, STYLE_LIST_FIELD)?
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        Ok(Self {
            model: model.to_vec(),
            store,
            style_list,
            geometry,
        })
    }
}

struct LocatedCell {
    bytes: Vec<u8>,
}

fn locate_cell(
    source: &dyn TableSource,
    table: &Table,
    target: Target,
) -> Result<Option<LocatedCell>, Error> {
    let tile_storage = bytes(&table.store, STORE_TILES)?.ok_or(invalid("tile storage missing"))?;
    let mut tiles = Vec::new();
    let tile_size = varint(tile_storage, 2)?.ok_or(invalid("tile size missing"))?;
    for entry in repeated_length_delimited_payloads(tile_storage, 1).map_err(map_wire)? {
        let key = u32::try_from(varint(entry, 1)?.ok_or(invalid("tile key missing"))?)
            .map_err(|_| invalid("tile key exceeds u32"))?;
        let tile = reference(entry, 2)?.ok_or(invalid("tile reference missing"))?;
        if tiles.iter().any(|(existing, _)| *existing == key) {
            return Err(invalid("tile key repeats"));
        }
        tiles.push((key, tile));
    }
    if tile_size == 0 {
        return Err(invalid("tile size is zero"));
    }
    let tile_key = u64::from(target.row) / tile_size;
    let local = u64::from(target.row) % tile_size;
    let tile = tiles
        .iter()
        .find(|(key, _)| u64::from(*key) == tile_key)
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?
        .1;
    let (_, tile_data) = unique_payload(source, tile, &[TILE_MESSAGE_TYPE])?;
    for row in repeated_length_delimited_payloads(tile_data, 5)
        .map_err(map_wire)?
        .into_iter()
    {
        if varint(row, 1)? != Some(local) {
            continue;
        }
        for (column, cell) in storage::bnc_cells(row).map_err(|_| invalid("malformed BNC row"))? {
            if column == target.column {
                return Ok(Some(LocatedCell {
                    bytes: cell.to_vec(),
                }));
            }
        }
    }
    Ok(None)
}

fn bnc_limits(data: &[u8]) -> BncLimits {
    if data.is_empty() {
        return BncLimits::DEFAULT;
    }
    BncLimits::DEFAULT
        .with_input_bytes(data.len())
        .and_then(|v| v.with_tail_bytes(data.len()))
        .unwrap_or(BncLimits::DEFAULT)
}

fn style_key(cell: Option<&[u8]>) -> Result<Option<u32>, Error> {
    let Some(cell) = cell else { return Ok(None) };
    let cell =
        BncCell::parse(cell, &bnc_limits(cell)).map_err(|_| invalid("malformed BNC cell"))?;
    Ok(cell.text_style_identifier())
}

#[derive(Clone, Copy)]
struct Entry {
    key: u32,
    refcount: u32,
    style: u64,
    owner: u64,
}

struct StyleList {
    root: u64,
    root_type: u32,
    root_data: Vec<u8>,
    entries: Vec<Entry>,
    segments: Vec<(u64, u32, Vec<u8>)>,
    next: u32,
}

impl StyleList {
    fn read(source: &dyn TableSource, root: u64) -> Result<Self, Error> {
        let (root_type, root_data) = unique_payload(
            source,
            root,
            &[LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE],
        )?;
        if varint(root_data, 1)? != Some(4) {
            return Err(Error::Refused(Refusal::UnsupportedStyleList));
        }
        let next =
            u32::try_from(varint(root_data, 2)?.ok_or(invalid("style-list cursor missing"))?)
                .map_err(|_| invalid("style-list cursor exceeds u32"))?;
        let mut entries = Vec::new();
        for raw in repeated_length_delimited_payloads(root_data, 3).map_err(map_wire)? {
            let key = u32::try_from(varint(raw, 1)?.ok_or(invalid("style-list key missing"))?)
                .map_err(|_| invalid("style-list key exceeds u32"))?;
            let refcount =
                u32::try_from(varint(raw, 2)?.ok_or(invalid("style-list refcount missing"))?)
                    .map_err(|_| invalid("style-list refcount exceeds u32"))?;
            let style = reference(raw, 4)?.ok_or(invalid("style-list style reference missing"))?;
            if entries.iter().any(|entry: &Entry| entry.key == key) {
                return Err(invalid("style-list key repeats"));
            }
            entries.push(Entry {
                key,
                refcount,
                style,
                owner: root,
            });
        }
        let mut segments = Vec::new();
        for raw in repeated_length_delimited_payloads(root_data, 4).map_err(map_wire)? {
            let identifier =
                reference(raw, 1)?.ok_or(invalid("style-list segment reference missing"))?;
            let (segment_type, segment_data) =
                unique_payload(source, identifier, &[LIST_SEGMENT_MESSAGE_TYPE])?;
            segments.push((identifier, segment_type, segment_data.to_vec()));
            for segment_entry in
                repeated_length_delimited_payloads(segment_data, 3).map_err(map_wire)?
            {
                let key =
                    u32::try_from(varint(segment_entry, 1)?.ok_or(invalid("segment key missing"))?)
                        .map_err(|_| invalid("segment key exceeds u32"))?;
                let refcount = u32::try_from(
                    varint(segment_entry, 2)?.ok_or(invalid("segment refcount missing"))?,
                )
                .map_err(|_| invalid("segment refcount exceeds u32"))?;
                let style = reference(segment_entry, 4)?
                    .ok_or(invalid("segment style reference missing"))?;
                if entries.iter().any(|entry| entry.key == key) {
                    return Err(invalid("style-list key repeats across segments"));
                }
                entries.push(Entry {
                    key,
                    refcount,
                    style,
                    owner: identifier,
                });
            }
        }
        Ok(Self {
            root,
            root_type,
            root_data: root_data.to_vec(),
            entries,
            segments,
            next,
        })
    }

    fn entry(&self, key: u32) -> Result<Entry, Error> {
        self.entries
            .iter()
            .find(|entry| entry.key == key)
            .copied()
            .ok_or(invalid("style-list key is not present"))
    }

    fn next_key(&self) -> Result<u32, Error> {
        let max = self
            .entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .unwrap_or(0);
        self.next
            .max(max.checked_add(1).ok_or(invalid("style key overflow"))?)
            .max(1)
            .checked_add(0)
            .ok_or(invalid("style key overflow"))
    }
}

fn style_payload(source: &dyn TableSource, identifier: u64) -> Result<&[u8], Error> {
    unique_payload(source, identifier, &[PARAGRAPH_STYLE_MESSAGE_TYPE]).map(|(_, data)| data)
}

fn stylesheet_of(source: &dyn TableSource, style: u64) -> Result<u64, Error> {
    let base =
        bytes(style_payload(source, style)?, 1)?.ok_or(invalid("style base payload missing"))?;
    reference(base, 5)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

/// Read effective character properties through the parent chain.
pub fn effective(source: &dyn TableSource, style: u64) -> Result<CharacterProperties, Error> {
    let wire = WireLimits::default();
    let mut properties = CharacterProperties::new();
    let mut current = Some(style);
    let mut seen = BTreeSet::new();
    for _ in 0..MAX_STYLE_DEPTH {
        let Some(identifier) = current else { break };
        if !seen.insert(identifier) {
            return Err(invalid("character-style parent cycle"));
        }
        let payload = style_payload(source, identifier)?;
        properties.inherit(
            &litchi_iwa_text_wire::character_style_properties(payload, wire)
                .map_err(|_| invalid("malformed character-style properties"))?,
        );
        current = litchi_iwa_text_wire::character_style_parent(payload, wire)
            .map_err(|_| invalid("malformed character-style parent"))?;
    }
    if current.is_some() {
        return Err(Error::Limit {
            kind: LimitKind::Nesting,
            observed: (MAX_STYLE_DEPTH + 1) as u64,
            maximum: MAX_STYLE_DEPTH as u64,
        });
    }
    Ok(properties)
}

/// Read the effective character properties of one target cell.
pub fn read(source: &dyn TableSource, target: Target) -> Result<CharacterProperties, Error> {
    let table = Table::read(source, target)?;
    let cell = locate_cell(source, &table, target)?;
    let key = style_key(cell.as_ref().map(|cell| cell.bytes.as_slice()))?;
    let style = match key {
        Some(key) => StyleList::read(source, table.style_list)?.entry(key)?.style,
        None => {
            let field = if target.row < table.geometry.header_rows {
                25
            } else if target.row
                >= table
                    .geometry
                    .rows
                    .saturating_sub(table.geometry.footer_rows)
            {
                27
            } else if target.column < table.geometry.header_columns {
                26
            } else {
                24
            };
            reference(&table.model, field)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))?
        },
    };
    effective(source, style)
}

/// The reduced direct character overrides to apply to a current style.
pub fn minimal(
    requested: &CharacterProperties,
    inherited: &CharacterProperties,
) -> CharacterProperties {
    let flag = |value: Option<bool>| value.unwrap_or(false);
    let mut output = CharacterProperties::new();
    output.bold = requested
        .bold
        .filter(|value| flag(inherited.bold) != *value);
    output.italic = requested
        .italic
        .filter(|value| flag(inherited.italic) != *value);
    output.underline = requested
        .underline
        .filter(|value| flag(inherited.underline) != *value);
    output.strikethrough = requested
        .strikethrough
        .filter(|value| flag(inherited.strikethrough) != *value);
    output.size = requested
        .size
        .filter(|value| inherited.size != Some(*value));
    output.font_name = requested
        .font_name
        .clone()
        .filter(|value| inherited.font_name.as_ref() != Some(value));
    output.color = requested
        .color
        .filter(|value| inherited.color != Some(*value));
    output
}

/// A style payload rewrite planned by the shared graph engine.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    /// Object identifier.
    pub identifier: u64,
    /// Unique message type to replace.
    pub message_type: u32,
    /// Replacement payload.
    pub data: Vec<u8>,
    /// References added to the message header.
    pub add_references: Vec<u64>,
    /// References removed from the message header.
    pub remove_references: Vec<u64>,
}

/// A new paragraph-style variation planned by the shared engine.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewObject {
    /// Fresh identifier supplied by the leaf.
    pub identifier: u64,
    /// Component/object beside which the object is inserted.
    pub beside: u64,
    /// Message type.
    pub message_type: u32,
    /// Message versions, when required.
    pub versions: Option<[u32; 3]>,
    /// Payload.
    pub data: Vec<u8>,
    /// Message object references.
    pub object_references: Vec<u64>,
}

/// Metadata transition requested by the shared plan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Registration {
    /// Register a fresh UUID.
    Uuid { object: u64 },
    /// Register a cross-component reference from `from` to `object`.
    ExternalReference { from: u64, object: u64 },
    /// Remove a style object and its metadata records.
    Removed { object: u64 },
}

/// One neutral whole-cell style graph plan.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Plan {
    /// Message payload rewrites, including the cell/tile/header payloads.
    pub rewrites: Vec<Rewrite>,
    /// Optional new style variation.
    pub insertion: Option<NewObject>,
    /// Optional style object to remove after rewrites.
    pub removal: Option<u64>,
    /// Metadata registry operations.
    pub registrations: Vec<Registration>,
    /// The style key to write to the selected cell (`None` removes it).
    pub cell_key: Option<u32>,
    /// The selected cell's model/coordinates, for leaf publication.
    pub target: Target,
}

/// Plan a direct CharacterProperties set. The caller supplies `fresh` only
/// when the current graph cannot be rewritten in place.
pub fn plan_set(
    source: &dyn TableSource,
    target: Target,
    requested: &CharacterProperties,
    fresh: Option<u64>,
) -> Result<Option<Plan>, Error> {
    let table = Table::read(source, target)?;
    let cell = locate_cell(source, &table, target)?;
    let old_key = style_key(cell.as_ref().map(|cell| cell.bytes.as_slice()))?;
    let list = StyleList::read(source, table.style_list)?;
    let parent = match old_key {
        Some(key) => list.entry(key)?.style,
        None => {
            let field = if target.row < table.geometry.header_rows {
                25
            } else if target.row
                >= table
                    .geometry
                    .rows
                    .saturating_sub(table.geometry.footer_rows)
            {
                27
            } else if target.column < table.geometry.header_columns {
                26
            } else {
                24
            };
            reference(&table.model, field)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))?
        },
    };
    let inherited = effective(source, parent)?;
    let reduced = minimal(requested, &inherited);
    if reduced == CharacterProperties::new() {
        return Ok(None);
    }
    let wire = WireLimits::default();
    let stylesheet = stylesheet_of(source, parent)?;
    let exclusive = old_key.is_some_and(|key| {
        list.entry(key)
            .ok()
            .is_some_and(|entry| entry.refcount == 1)
    }) && style_is_exclusive(source, parent, table.style_list)?;
    if exclusive {
        let data = litchi_iwa_text_wire::character_style_variation(
            Some(parent),
            stylesheet,
            &reduced,
            wire,
        )
        .map_err(|_| invalid("character-style variation encoding failed"))?;
        let data = with_paragraph_properties(data)?;
        return Ok(Some(Plan {
            rewrites: vec![Rewrite {
                identifier: parent,
                message_type: PARAGRAPH_STYLE_MESSAGE_TYPE,
                data,
                add_references: Vec::new(),
                remove_references: Vec::new(),
            }],
            target,
            cell_key: None,
            ..Plan::default()
        }));
    }
    let fresh = fresh.ok_or(Error::NeedsIdentifier)?;
    let data =
        litchi_iwa_text_wire::character_style_variation(Some(parent), stylesheet, &reduced, wire)
            .map_err(|_| invalid("character-style variation encoding failed"))?;
    let data = with_paragraph_properties(data)?;
    let (sheet_type, sheet_payload) =
        unique_payload(source, stylesheet, &[STYLESHEET_MESSAGE_TYPE])?;
    let sheet_data =
        litchi_iwa_text_wire::stylesheet_with_style(sheet_payload, fresh, Some(parent), wire)
            .map_err(|_| invalid("stylesheet variation encoding failed"))?;
    let key = list.next_key()?;
    let next = key.checked_add(1).ok_or(invalid("style key overflow"))?;
    let mut list_data =
        patch_varint_field(&list.root_data, 2, true, Some(u64::from(next))).map_err(map_wire)?;
    let mut entry = Vec::new();
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 1_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut entry, u64::from(key));
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 2_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut entry, 1);
    let mut reference_payload = Vec::new();
    litchi_iwa_common::varint::encode_varint_into(&mut reference_payload, 1_u64 << 3);
    litchi_iwa_common::varint::encode_varint_into(&mut reference_payload, fresh);
    append_length_delimited_field(&mut entry, 4, &reference_payload, WireLimits::DEFAULT)
        .map_err(map_wire)?;
    append_length_delimited_field(&mut list_data, 3, &entry, WireLimits::DEFAULT)
        .map_err(map_wire)?;
    let mut rewrites = vec![
        Rewrite {
            identifier: stylesheet,
            message_type: sheet_type,
            data: sheet_data,
            add_references: vec![fresh],
            remove_references: Vec::new(),
        },
        Rewrite {
            identifier: list.root,
            message_type: list.root_type,
            data: list_data,
            add_references: vec![fresh],
            remove_references: Vec::new(),
        },
    ];
    if let Some(old) = old_key {
        let mut remaining = BTreeMap::from([(old, 1_u32)]);
        let owner = list.entry(old)?.owner;
        let (released, dropped) = if owner == list.root {
            litchi_numbers_wire::table_topology::lists::release(&rewrites[1].data, &mut remaining)
                .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?
        } else {
            let (_, segment_type, segment_data) = list
                .segments
                .iter()
                .find(|(identifier, _, _)| *identifier == owner)
                .ok_or(invalid("style-list segment disappeared"))?;
            let (released, dropped) =
                litchi_numbers_wire::table_topology::lists::release(segment_data, &mut remaining)
                    .map_err(|_| Error::Refused(Refusal::UnsupportedStyleList))?;
            if let Some(released) = released {
                rewrites.push(Rewrite {
                    identifier: owner,
                    message_type: *segment_type,
                    data: released,
                    add_references: Vec::new(),
                    remove_references: dropped.clone(),
                });
            }
            (None, dropped)
        };
        if !remaining.is_empty() {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        if let Some(released) = released {
            rewrites[1].data = released;
        }
        if owner == list.root {
            rewrites[1].remove_references = dropped;
        }
    }
    Ok(Some(Plan {
        rewrites,
        insertion: Some(NewObject {
            identifier: fresh,
            beside: stylesheet,
            message_type: PARAGRAPH_STYLE_MESSAGE_TYPE,
            versions: None,
            data,
            object_references: vec![parent],
        }),
        registrations: vec![
            Registration::Uuid { object: fresh },
            Registration::ExternalReference {
                from: list.root,
                object: fresh,
            },
        ],
        cell_key: Some(key),
        target,
        ..Plan::default()
    }))
}

fn with_paragraph_properties(mut data: Vec<u8>) -> Result<Vec<u8>, Error> {
    if bytes(&data, PARAGRAPH_PROPERTIES_FIELD)?.is_none() {
        append_length_delimited_field(
            &mut data,
            PARAGRAPH_PROPERTIES_FIELD,
            &[],
            WireLimits::DEFAULT,
        )
        .map_err(map_wire)?;
    }
    Ok(data)
}

fn style_is_exclusive(source: &dyn TableSource, style: u64, list: u64) -> Result<bool, Error> {
    let mut found = false;
    for component in source.catalog().iter() {
        for object in &component.archive().objects {
            let identifier = object.archive_info.identifier.unwrap_or_default();
            if identifier == style || identifier == list {
                continue;
            }
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&style))
            {
                return Ok(false);
            }
            for message in &object.messages {
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && litchi_iwa_text_wire::character_style_parent(
                        &message.data,
                        WireLimits::default(),
                    )
                    .ok()
                    .flatten()
                        == Some(style)
                {
                    return Ok(false);
                }
            }
            found = true;
        }
    }
    let _ = found;
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::minimal;
    use litchi_iwa_text_wire::CharacterProperties;

    #[test]
    fn minimal_keeps_only_values_that_differ_from_inheritance() {
        let mut requested = CharacterProperties::new();
        requested.bold = Some(true);
        requested.italic = Some(false);
        requested.size = Some(18.0);
        requested.font_name = Some("Helvetica".to_owned());

        let mut inherited = CharacterProperties::new();
        inherited.bold = Some(true);
        inherited.italic = Some(true);
        inherited.size = Some(12.0);
        inherited.font_name = Some("Helvetica".to_owned());

        let reduced = minimal(&requested, &inherited);
        assert_eq!(reduced.bold, None);
        assert_eq!(reduced.italic, Some(false));
        assert_eq!(reduced.size, Some(18.0));
        assert_eq!(reduced.font_name, None);
    }

    #[test]
    fn unset_emphasis_is_compared_as_host_default_off() {
        let mut requested = CharacterProperties::new();
        requested.bold = Some(false);
        let inherited = CharacterProperties::new();
        assert!(minimal(&requested, &inherited).is_empty());
    }
}
