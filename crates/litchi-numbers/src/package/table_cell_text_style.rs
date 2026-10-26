//! Implementation of the public [`crate::cell::text_style`] transactions:
//! character properties (bold, italic, underline, size, font, colour) of
//! table cells (`IWORK_API_REQUESTS.md` request 7, record 0309).
//!
//! A Numbers cell's text style is a key in its table's style list
//! (`DataStore.styleTable`) whose entry references a
//! `TSWP.ParagraphStyleArchive`; a cell without a key uses its section's
//! base text style (header rows, footer rows, header columns, body). A set
//! creates a paragraph-style variation of the cell's current style that
//! overrides the requested character properties (every other property
//! stays inherited), inserts it beside the stylesheet and lists it there,
//! appends a style-list entry for it and points the cell at the new key;
//! the old key's entry loses the cell's reference. Each cell takes its own
//! variation, and a property equal to the inherited one stays inherited,
//! as the host editor writes them (record 0370). Ported from the host
//! editor's `cell_paragraph_style` (new-style path).

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeMap;
use std::fmt;

use litchi_iwa_core::RawMessage;
use litchi_iwa_text_wire::CharacterProperties;
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, Fault, FaultResult, Working,
    metadata::{self, Transition},
    references,
};
use super::table_topology::wire::{
    bytes_field, push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_bytes,
    single_reference, single_varint, varint_field,
};
use super::table_topology::{
    self as topology, LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE, TILE_MESSAGE_TYPE, lists,
    storage,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("cell-text-style", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.ParagraphStyleArchive`.
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
/// `TST.HeaderStorageBucket`.
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `DataStore.styleTable`.
const STYLE_LIST_FIELD: u32 = 5;
/// The longest parent chain a cell's text style is resolved through.
const MAX_STYLE_DEPTH: usize = 16;

/// A typed, content-redacted cell text-style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell text-style source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} cell text-style items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell text-style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell text-style patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// The style sets no property.
    #[error("the style sets no property")]
    EmptyStyle,
    /// Commit was requested before staging a cell.
    #[error("no cell text style is staged")]
    NoStagedOperation,
    /// The cell's style lives outside what this transaction rewrites (no
    /// stylesheet, a segmented list, a style in another component).
    #[error("a surviving object depends on the cell's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell text styles")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } | Self::CellNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::EmptyStyle { .. } => RefusalKind::InvalidValue,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency { .. } | Self::UnsupportedSource { .. } => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

pub(super) fn from_topology(error: topology::Error) -> Error {
    use topology::Error as E;
    match error {
        E::Refused(crate::package::table_topology::Refusal::SheetNotFound) => {
            Error::Refused(Refusal::SheetNotFound)
        },
        E::Refused(crate::package::table_topology::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        E::Allocation { amount } => Error::Allocation { amount },
        E::Refused(crate::package::table_topology::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        _ => Error::InvalidSource,
    }
}

fn from_style(error: litchi_iwa_text_wire::CharStyleError) -> Error {
    match error {
        litchi_iwa_text_wire::CharStyleError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_text_wire::CharStyleError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

pub(super) fn style_wire_limits(
    working: &Working<'_>,
) -> Result<litchi_iwa_common::WireLimits, Error> {
    let maximum = working.archive_limits().max_message_bytes();
    litchi_iwa_common::WireLimits::default()
        .with_input_bytes(maximum.min(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| {
            limits.with_output_bytes(maximum.min(litchi_iwa_common::WireLimits::MAX_OUTPUT_BYTES))
        })
        .map_err(|_| Error::InvalidSource)
}

/// One immutable edit that sets the character properties of cells of one
/// table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    cells: Vec<(At, CharacterProperties)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("cells", &self.cells.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `properties` on the stored cell at `position`; a later stage
    /// of the same cell adds to the earlier one.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Refusal::EmptyStyle`].
    pub fn set(mut self, position: At, properties: CharacterProperties) -> Result<Self, Error> {
        if properties.is_empty() {
            return Err(Error::Refused(Refusal::EmptyStyle));
        }
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        if let Some((_, staged)) = self.cells.iter_mut().find(|(cell, _)| *cell == position) {
            let mut merged = properties;
            merged.inherit(staged);
            *staged = merged;
            return Ok(self);
        }
        self.cells
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.cells.push((position, properties));
        Ok(self)
    }

    /// Validate and atomically publish the staged cell styles.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] (an empty cell takes a style-only
    /// cell); [`Refusal::UnsupportedDependency`] for a style graph
    /// this transaction does not rewrite; source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.cells.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let published = publish(self.source, self.sheet, self.table, &self.cells)?;
        for (position, properties) in &self.cells {
            let effective = text_style(&published.package, self.sheet, self.table, *position)?;
            let mut expected = properties.clone();
            expected.inherit(&effective);
            if with_flags_off(expected) != with_flags_off(effective) {
                return Err(Error::Verification);
            }
        }
        // Every stored cell survives; an empty cell may gain a style-only
        // cell (record 0370).
        let count = |package: &Package| {
            package.workbook_ref().sheets()[self.sheet]
                .tables()
                .nth(self.table)
                .map(|table| table.iter_cells().count())
        };
        let (Some(before), Some(after)) = (count(self.source), count(&published.package)) else {
            return Err(Error::Verification);
        };
        if after < before || after > before.saturating_add(self.cells.len()) {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

/// One stored BNC cell located in the table's tiles.
pub(super) struct Located {
    pub(super) tile: u64,
    pub(super) tile_index: usize,
    /// The field index of the cell's row record among the tile's fields.
    pub(super) row_field: usize,
    pub(super) bytes: Vec<u8>,
}

pub(super) fn locate_cell(
    working: &Working<'_>,
    store: &[u8],
    position: At,
) -> Result<Option<Located>, Error> {
    let (tiles, size) = topology::tiles(store).map_err(from_topology)?;
    let key = u64::from(position.row()) / size;
    let local = u64::from(position.row()) % size;
    let Some((_, tile)) = tiles
        .iter()
        .find(|(tile_key, _)| u64::from(*tile_key) == key)
    else {
        return Ok(None);
    };
    let (tile_index, data) =
        topology::message(working, *tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
    for (row_field, field) in super::table_topology::wire::fields(data)?
        .iter()
        .enumerate()
    {
        if field.number != 5 || single_varint(field.payload, 1)? != Some(local) {
            continue;
        }
        for (column, bytes) in storage::bnc_cells(field.payload).map_err(Fault::from)? {
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

/// The entry of `key` in the style list (root or segments): its reference.
pub(super) fn list_reference(
    working: &Working<'_>,
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
        let (_, data) = topology::message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])
            .map_err(Error::from)?;
        if let Some(found) = find(data)? {
            return Ok(found);
        }
    }
    Err(Error::InvalidSource)
}

/// The section base text style of `position`.
pub(super) fn base_style(
    model: &[u8],
    geometry: topology::Geometry,
    position: At,
) -> Result<u64, Error> {
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

/// The current text style object of the stored cell at `position` (its
/// key's entry, else its section's base style).
pub(super) fn current_style(
    working: &Working<'_>,
    sheet: usize,
    table: usize,
    position: At,
) -> Result<u64, Error> {
    let target = topology::locate(working, sheet, table).map_err(from_topology)?;
    let (_, model) = topology::message(working, target.model, &[6_001]).map_err(Error::from)?;
    let geometry = topology::geometry(model).map_err(from_topology)?;
    if position.row() >= geometry.rows || position.column() >= geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let store = single_bytes(model, 4)?.ok_or(Error::InvalidSource)?;
    let base = base_style(model, geometry, position)?;
    let Some(cell) = locate_cell(working, store, position)? else {
        return Ok(base);
    };
    let cell = litchi_numbers_wire::BncCell::parse(
        &cell.bytes,
        &crate::cell::wire::limits_for_source(&cell.bytes),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let Some(key) = cell.text_style_identifier() else {
        return Ok(base);
    };
    let list = single_reference(store, STYLE_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
    let (_, list_data) =
        topology::message(working, list, &[LIST_MESSAGE_TYPE, 6_201]).map_err(Error::from)?;
    list_reference(working, list_data, key)?.ok_or(Error::InvalidSource)
}

/// The character properties the style `style` and its parent chain set.
fn effective(working: &Working<'_>, style: u64) -> Result<CharacterProperties, Error> {
    let wire = style_wire_limits(working)?;
    let mut properties = CharacterProperties::new();
    let mut current = Some(style);
    for _ in 0..MAX_STYLE_DEPTH {
        let Some(identifier) = current else { break };
        let Ok(object) = working.object(identifier) else {
            break;
        };
        let Some(message) = object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
        else {
            break;
        };
        properties.inherit(
            &litchi_iwa_text_wire::character_style_properties(&message.data, wire)
                .map_err(from_style)?,
        );
        current = litchi_iwa_text_wire::character_style_parent(&message.data, wire)
            .map_err(from_style)?;
    }
    Ok(properties)
}

fn text_style(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
) -> Result<CharacterProperties, Error> {
    let working = Working::new(package)?;
    let style = current_style(&working, sheet, table, position)?;
    effective(&working, style)
}

/// `TSWP.ParagraphStyleArchive.para_properties`.
const PARAGRAPH_PROPERTIES_FIELD: u32 = 12;

/// `style` with an empty `para_properties` when it has none.
pub(super) fn with_paragraph_properties(mut style: Vec<u8>) -> Result<Vec<u8>, Error> {
    if single_bytes(&style, PARAGRAPH_PROPERTIES_FIELD)?.is_none() {
        push_bytes_field(&mut style, PARAGRAPH_PROPERTIES_FIELD, &[])?;
    }
    Ok(style)
}

/// `properties` with every unset emphasis flag read as off.
fn with_flags_off(mut properties: CharacterProperties) -> CharacterProperties {
    for flag in [
        &mut properties.bold,
        &mut properties.italic,
        &mut properties.underline,
        &mut properties.strikethrough,
    ] {
        flag.get_or_insert(false);
    }
    properties
}

/// The overrides of `requested` that differ from `inherited`.
fn minimal(
    requested: &CharacterProperties,
    inherited: &CharacterProperties,
) -> CharacterProperties {
    // An unset emphasis flag reads as off.
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

/// The stylesheet a paragraph style belongs to (`super.stylesheet`).
pub(super) fn stylesheet_of(working: &Working<'_>, style: u64) -> Result<u64, Error> {
    let (_, data) =
        topology::message(working, style, &[PARAGRAPH_STYLE_MESSAGE_TYPE]).map_err(Error::from)?;
    let base = single_bytes(data, 1)?.ok_or(Error::InvalidSource)?;
    single_reference(base, 5)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

/// Replace one message, transitioning its aggregate object references.
pub(super) fn replace_with_references(
    working: &mut Working<'_>,
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

pub(super) fn replace_message(
    working: &mut Working<'_>,
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

#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: styles, stylesheet, list, cells, registry"
)]
fn publish(
    source: &Package,
    sheet: usize,
    table: usize,
    cells: &[(At, CharacterProperties)],
) -> Result<structure::Published, Error> {
    let mut working = Working::new(source)?;
    // A cell-level edit keeps the root previews, as the host does (record 0370).
    working.keep_previews();
    let wire = style_wire_limits(&working)?;
    let target = topology::locate(&working, sheet, table).map_err(from_topology)?;
    let (_, model) = topology::message(&working, target.model, &[6_001]).map_err(Error::from)?;
    let model_columns = topology::geometry(model).map_err(from_topology)?.columns;
    let store = single_bytes(model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();
    let list = single_reference(&store, STYLE_LIST_FIELD)?
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let list_member = working.locate(list)?.ok_or(Error::InvalidSource)?;
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();

    let mut list_data = topology::message(&working, list, &[LIST_MESSAGE_TYPE, 6_201])
        .map_err(Error::from)?
        .1
        .to_vec();
    let list_index = topology::message(&working, list, &[LIST_MESSAGE_TYPE, 6_201])
        .map_err(Error::from)?
        .0;
    let mut releases: BTreeMap<u32, u32> = BTreeMap::new();
    let mut appended: Vec<u64> = Vec::new();
    let mut cell_updates: Vec<(At, u32)> = Vec::new();
    for (position, properties) in cells {
        // An empty cell takes a style-only cell, as the host editor
        // materializes it (record 0370).
        let old_key = match locate_cell(&working, &store, *position)? {
            Some(located) => litchi_numbers_wire::BncCell::parse(
                &located.bytes,
                &crate::cell::wire::limits_for_source(&located.bytes),
            )
            .map_err(|_error| Error::InvalidSource)?
            .text_style_identifier(),
            None => None,
        };
        let current = current_style(&working, sheet, table, *position)?;
        // A property equal to the inherited value stays inherited, as the
        // host editor's minimal style variation leaves it (record 0370).
        let inherited = effective(&working, current)?;
        let properties = &minimal(properties, &inherited);
        if properties.is_empty() {
            continue;
        }
        // One variation per cell, as the host editor writes one per cell
        // (record 0370); a later stage of the same cell already merged into
        // this one.
        let key = {
            let stylesheet = stylesheet_of(&working, current)?;
            let sheet_member = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
            if working.locate(current)?.as_deref() != Some(sheet_member.as_str()) {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let identifier = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
            let data = litchi_iwa_text_wire::character_style_variation(
                Some(current),
                stylesheet,
                properties,
                wire,
            )
            .map_err(from_style)?;
            // The variation states its (empty) paragraph properties, as the
            // host editor's variation does (record 0370).
            let data = with_paragraph_properties(data)?;
            let template = working.object(current)?;
            let object = structure::new_object(
                identifier,
                Some(template),
                PARAGRAPH_STYLE_MESSAGE_TYPE,
                data,
                // Native styles list their parent, not their stylesheet
                // back-reference (record 0370; every tracked native style).
                vec![current],
                working.archive_limits(),
            )?;
            working.insert(&sheet_member, object)?;
            // The stylesheet lists the variation.
            let (sheet_index, sheet_data) =
                topology::message(&working, stylesheet, &[STYLESHEET_MESSAGE_TYPE])
                    .map_err(Error::from)?;
            let sheet_data = litchi_iwa_text_wire::stylesheet_with_style(
                sheet_data,
                identifier,
                Some(current),
                wire,
            )
            .map_err(from_style)?;
            // Some native stylesheets carry field-local references their
            // aggregate list does not repeat; the core refuses to transition
            // such a header, so the payload is then replaced with the
            // header kept as it is (the style list's entry declares the
            // reference the census follows).
            let declared = replace_with_references(
                &mut working,
                stylesheet,
                sheet_index,
                sheet_data.clone(),
                &[],
                &[identifier],
            );
            match declared {
                Ok(()) => {},
                Err(Fault::Invalid) => {
                    replace_message(&mut working, stylesheet, sheet_index, sheet_data)?;
                    // The aggregate list still names the new style, as the
                    // host editor's stylesheet write does (record 0370).
                    let object = working.object_mut(stylesheet)?;
                    let references = &mut object
                        .archive_info
                        .message_infos
                        .get_mut(sheet_index)
                        .ok_or(Fault::Invalid)?
                        .object_references;
                    if !references.contains(&identifier) {
                        references
                            .try_reserve(1)
                            .map_err(|_error| Error::Allocation { amount: 1 })?;
                        references.push(identifier);
                    }
                },
                Err(other) => return Err(other.into()),
            }
            let sheet_component = registry.component_for_entry(&sheet_member)?.identifier;
            let list_component = registry.component_for_entry(&list_member)?.identifier;
            let seed = metadata::identity_seed(&working, identifier)?;
            transition.register_uuid(&registry, sheet_component, identifier, seed)?;
            transition.reference(&registry, list_component, sheet_component, identifier)?;
            // A new style-list entry for it.
            let key = u32::try_from(single_varint(&list_data, 2)?.unwrap_or(1))
                .map_err(|_error| Error::InvalidSource)?
                .max(1);
            let mut entry = Vec::new();
            push_varint_field(&mut entry, 1, u64::from(key))?;
            push_varint_field(&mut entry, 2, 0)?;
            let mut reference = Vec::new();
            push_varint_field(&mut reference, 1, identifier)?;
            push_bytes_field(&mut entry, 4, &reference)?;
            let mut rebuilt = rebuild(&list_data, |field| {
                if field.number == 2 {
                    return varint_field(2, u64::from(key) + 1).map(Some);
                }
                Ok(None)
            })?;
            if single_varint(&rebuilt, 2)?.is_none() {
                return Err(Error::InvalidSource);
            }
            push_bytes_field(&mut rebuilt, 3, &entry)?;
            list_data = rebuilt;
            appended.push(identifier);
            key
        };
        // One more reference to the entry.
        list_data = rebuild(&list_data, |field| {
            if field.number != 3 || single_varint(field.payload, 1)? != Some(u64::from(key)) {
                return Ok(None);
            }
            let count = single_varint(field.payload, 2)?.unwrap_or(0);
            let entry = rebuild(field.payload, |inner| {
                if inner.number == 2 {
                    return varint_field(2, count + 1).map(Some);
                }
                Ok(None)
            })?;
            bytes_field(3, &entry).map(Some)
        })?;
        if let Some(old) = old_key {
            *releases.entry(old).or_insert(0) += 1;
        }
        cell_updates.push((*position, key));
    }
    // Release the cells' old keys (root entries only).
    let mut remaining = releases;
    let (released, pruned) = lists::release(&list_data, &mut remaining).map_err(Fault::from)?;
    if !remaining.is_empty() {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    if let Some(released) = released {
        list_data = released;
    }
    replace_with_references(
        &mut working,
        list,
        list_index,
        list_data,
        &pruned,
        &appended,
    )?;

    // Point each cell at its new key.
    for (position, key) in cell_updates {
        let Some(located) = locate_cell(&working, &store, position)? else {
            materialize(&mut working, &store, model_columns, position, key)?;
            continue;
        };
        let limits = crate::cell::wire::limits_for_source(&located.bytes);
        let mut cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &limits)
            .map_err(|_error| Error::InvalidSource)?;
        cell.set_text_style_identifier(Some(key));
        let encoded =
            crate::cell::wire::encode(&cell, &limits).map_err(|_| Error::InvalidSource)?;
        let (_, tile_data) =
            topology::message(&working, located.tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
        let mut field_index = 0_usize;
        let rewritten = rebuild(tile_data, |field| {
            let this = field_index;
            field_index += 1;
            if this != located.row_field {
                return Ok(None);
            }
            let row = storage::replace_cell(field.payload, position.column(), &encoded)
                .map_err(Fault::from)?
                .ok_or(Fault::Invalid)?;
            bytes_field(5, &row).map(Some)
        })?;
        replace_message(&mut working, located.tile, located.tile_index, rewritten)?;
    }
    Ok(working.publish(&transition)?)
}

/// Store a style-only cell for the empty cell at `position` (its row's and
/// column's headers count it), pointing at text-style key `key`.
pub(super) fn materialize(
    working: &mut Working<'_>,
    store: &[u8],
    columns: u32,
    position: At,
    key: u32,
) -> Result<(), Error> {
    let mut cell = litchi_numbers_wire::BncCell::minimal();
    cell.set_text_style_identifier(Some(key));
    materialize_cell(working, store, columns, position, &cell)
}

/// Store `cell` at the empty position `position`, counting it in its row
/// and column headers.
pub(super) fn materialize_cell(
    working: &mut Working<'_>,
    store: &[u8],
    columns: u32,
    position: At,
    cell: &litchi_numbers_wire::BncCell,
) -> Result<(), Error> {
    let (tiles, size) = topology::tiles(store).map_err(from_topology)?;
    let tile_key = u64::from(position.row()) / size;
    let local =
        u32::try_from(u64::from(position.row()) % size).map_err(|_| Error::InvalidSource)?;
    let (_, tile) = tiles
        .iter()
        .find(|(candidate, _)| u64::from(*candidate) == tile_key)
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let limits = litchi_numbers_wire::Limits::DEFAULT;
    let encoded = crate::cell::wire::encode(cell, &limits).map_err(|_| Error::InvalidSource)?;
    let (tile_index, tile_data) =
        topology::message(working, *tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
    let written = litchi_numbers_wire::table_write::tile_set_cell(
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
    let bucket = usize::try_from(u64::from(position.row()) / storage::HEADER_BUCKET_ROWS)
        .map_err(|_| Error::InvalidSource)?;
    let reference = *repeated_bytes(headers, 2)?
        .get(bucket)
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let bucket = references::reference_identifier(reference)?;
    let (bucket_index, bucket_data) =
        topology::message(working, bucket, &[HEADER_BUCKET_MESSAGE_TYPE]).map_err(Error::from)?;
    let adjusted =
        litchi_numbers_wire::table_write::header_adjust(bucket_data, position.row(), 1, &limits)
            .map_err(|_| Error::InvalidSource)?;
    replace_message(working, bucket, bucket_index, adjusted)?;
    // So does its column's header.
    let column_bucket =
        single_reference(store, 2)?.ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let (column_index, column_data) =
        topology::message(working, column_bucket, &[HEADER_BUCKET_MESSAGE_TYPE])
            .map_err(Error::from)?;
    let adjusted =
        litchi_numbers_wire::table_write::header_adjust(column_data, position.column(), 1, &limits)
            .map_err(|_| Error::InvalidSource)?;
    replace_message(working, column_bucket, column_index, adjusted).map_err(Error::from)
}

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        from_topology(error)
    }
}

impl Package {
    /// The effective character properties of one cell's text (its text
    /// style and that style's parents; a cell without its own style uses
    /// its section's base style).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], source errors.
    pub fn table_cell_text_style<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<CharacterProperties, Error> {
        let (sheet, table) = self.resolve_text_style_target(sheet.into(), table.into())?;
        text_style(self, sheet, table, position)
    }

    /// Start an immutable cell text-style edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::text_style::CharacterProperties;
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut bold = CharacterProperties::new();
    /// bold.bold = Some(true);
    /// let commit = package
    ///     .edit_table_cell_text_style(0, 0)?
    ///     .set(At::new(1, 1), bold)?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_table_cell_text_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_text_style<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_text_style_target(sheet.into(), table.into())?;
        let leaf = self.workbook_ref().sheets()[sheet]
            .tables()
            .nth(table)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            rows: leaf.row_count(),
            columns: leaf.column_count(),
            cells: Vec::new(),
        })
    }

    /// Apply a cell text-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_text_style(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_text_style_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize), Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table).map_err(from_topology)?;
        Ok((sheet, table))
    }
}

#[cfg(test)]
mod wire_error_tests {
    use super::{Error, LimitKind, from_style};

    #[test]
    fn character_style_resource_errors_survive_the_owner_boundary() {
        let error = from_style(litchi_iwa_text_wire::CharStyleError::Limit {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed: 11,
            maximum: 10,
        });
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Wire(litchi_iwa_common::LimitKind::RewriteWork),
                observed: 11,
                maximum: 10,
            }
        ));
        assert!(matches!(
            from_style(litchi_iwa_text_wire::CharStyleError::Allocation { amount: 42 }),
            Error::Allocation { amount: 42 }
        ));
        assert!(matches!(
            from_style(litchi_iwa_text_wire::CharStyleError::Invalid),
            Error::InvalidSource
        ));
    }
}
