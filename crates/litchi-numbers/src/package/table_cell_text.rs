//! Implementation of the public [`crate::cell::text`] transactions: one
//! character or paragraph property of whole cells (record 0370).
//!
//! A cell's text style is a key in its table's style list whose entry names
//! a `TSWP.ParagraphStyleArchive`; a cell without a key uses its section's
//! base style. Each staged change runs the shared paragraph-style engine
//! ([`litchi_iwa_text_wire::paragraph_style`], record 0374) over the cell's
//! style, with the cell's own variation counted exclusive when its list
//! entry is its only referrer, nothing else references the style and no
//! style names it as parent. The engine's plan is published here:
//!
//! - **Rewrite**: the cell's own variation takes the new payload in place;
//! - **Derive**: a new child variation of the cell's style is stored beside
//!   its stylesheet and registered there, a list entry names it and the cell
//!   points at the entry (an empty cell takes a style-only cell);
//! - **Collapse**: the cell's variation no longer overrides anything, so the
//!   cell points back at the parent (no key when the parent is the section's
//!   base style), the variation's entry is released and the variation is
//!   removed with its stylesheet registration.
//!
//! This is the behaviour of the host editor's `cell_paragraph_style`
//! `set_table_cell_*`/`reset_table_cell_*` families, collapsed into one
//! property verb.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{FieldInfo, FieldPath, FieldType, UnknownFieldRule};
use litchi_iwa_text_wire::paragraph_style::{
    self as engine, Change, ManagedFields, Plan, Property, PropertyKind, Styles, Target,
};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, Fault, Published, Working,
    metadata::{self, Transition},
};
use super::table_cell_text_style as cells;
use super::table_topology::wire::{
    bytes_field, push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_bytes,
    single_reference, single_varint, varint_field,
};
use super::table_topology::{self as topology, LIST_MESSAGE_TYPE, TILE_MESSAGE_TYPE, storage};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("cell-text", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.ParagraphStyleArchive`.
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `DataStore.styleTable`.
const STYLE_LIST_FIELD: u32 = 5;

/// A typed, content-redacted cell text failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell text source is invalid")]
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
    #[error("could not allocate {amount} cell text items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell text candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell text patch conflicts with this package")]
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
    /// Commit was requested before staging a change.
    #[error("no cell text change is staged")]
    NoStagedOperation,
    /// The cell's style lives outside what this transaction rewrites (no
    /// stylesheet, a segmented list, a style in another component).
    #[error("a surviving object depends on the cell's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell text properties")]
    UnsupportedSource,
    /// The value cannot be represented by the native style.
    #[error("the property value cannot be stored")]
    InvalidValue,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::InvalidValue => RefusalKind::InvalidValue,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
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

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        from_cells(cells::from_topology(error))
    }
}

/// Map the shared cell-style helpers' error (the text-style family's).
fn from_cells(error: super::table_cell_text_style::Error) -> Error {
    use super::table_cell_text_style::{Error as E, LimitKind as K, Refusal as R};
    match error {
        E::Refused(R::SheetNotFound) => Error::Refused(Refusal::SheetNotFound),
        E::Refused(R::TableNotFound) => Error::Refused(Refusal::TableNotFound),
        E::Refused(R::CellNotFound) => Error::Refused(Refusal::CellNotFound),
        E::Refused(R::UnsupportedDependency) => Error::Refused(Refusal::UnsupportedDependency),
        E::Refused(R::UnsupportedSource) => Error::Refused(Refusal::UnsupportedSource),
        E::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                K::PackageBytes => LimitKind::PackageBytes,
                K::PayloadBytes => LimitKind::PayloadBytes,
                K::Identifiers => LimitKind::Identifiers,
                K::ScratchBytes => LimitKind::ScratchBytes,
                K::AllocationEvents => LimitKind::AllocationEvents,
                K::Wire(kind) => LimitKind::Wire(kind),
            },
            observed,
            maximum,
        },
        E::Allocation { amount } => Error::Allocation { amount },
        E::Verification => Error::Verification,
        E::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// One staged change of one cell.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind),
}

/// One immutable edit that changes text properties of cells of one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, position: At, change: Staged) -> Result<Self, Error> {
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((position, change));
        Ok(self)
    }

    /// Stage `property` as the effective value of the cell at `position`.
    /// Changes apply in staging order.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(self, position: At, property: Property) -> Result<Self, Error> {
        self.stage(position, Staged::Set(property))
    }

    /// Stage removing the cell variation's own override of `kind`, so the
    /// value is inherited again.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(self, position: At, kind: PropertyKind) -> Result<Self, Error> {
        self.stage(position, Staged::Reset(kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves a property as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; [`Refusal::UnsupportedDependency`]
    /// for a style graph this transaction does not rewrite; source,
    /// resource and verification errors. Failure never publishes a partial
    /// package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let change = match change {
                Staged::Set(property) => Change::Set(property.clone()),
                Staged::Reset(kind) => Change::Reset(*kind),
            };
            if let Some(published) = apply_one(base, self.sheet, self.table, *position, &change)? {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        // Every staged set reads back.
        for (position, change) in &self.changes {
            if let Staged::Set(property) = change
                && let Some(later) = self
                    .changes
                    .iter()
                    .rposition(|(other, _)| other == position)
                && self.changes[later].1 == *change
                && &text(&package, self.sheet, self.table, *position, property.kind())? != property
            {
                return Err(Error::Verification);
            }
        }
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The package's paragraph styles, as the engine reads them.
struct WorkingStyles<'w, 'p>(&'w Working<'p>);

impl Styles for WorkingStyles<'_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn wire_limits(working: &Working<'_>) -> Result<WireLimits, Error> {
    cells::style_wire_limits(working).map_err(from_cells)
}

/// The cell's table: its model payload, style list and geometry.
struct Table {
    model: Vec<u8>,
    store: Vec<u8>,
    list: u64,
    geometry: topology::Geometry,
}

fn table(working: &Working<'_>, sheet: usize, table: usize) -> Result<Table, Error> {
    let target = topology::locate(working, sheet, table).map_err(Error::from)?;
    let (_, model) = topology::message(working, target.model, &[6_001]).map_err(Error::from)?;
    let geometry = topology::geometry(model).map_err(Error::from)?;
    let store = single_bytes(model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();
    let list = single_reference(&store, STYLE_LIST_FIELD)?
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    Ok(Table {
        model: model.to_vec(),
        store,
        list,
        geometry,
    })
}

/// The effective `kind` of the cell at `position`.
fn text(
    package: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let working = Working::new(package)?;
    let style = cells::current_style(&working, sheet, table_index, position).map_err(from_cells)?;
    engine::effective(
        &WorkingStyles(&working),
        style,
        kind,
        wire_limits(&working)?,
    )
    .map_err(from_engine)
}

/// The `refcount` of the root entry `key` of the style list, or `None`
/// when the key lives in a segment.
fn entry_refcount(list: &[u8], key: u32) -> Result<Option<u64>, Error> {
    for entry in repeated_bytes(list, 3)? {
        if single_varint(entry, 1)? == Some(u64::from(key)) {
            return Ok(Some(single_varint(entry, 2)?.unwrap_or(0)));
        }
    }
    Ok(None)
}

/// Whether nothing but the style list `list`, its stylesheet and the style
/// itself references `style`, and no paragraph style names it as parent
/// (the host editor's `style_is_exclusive_to_list` and
/// `style_has_children`).
fn exclusive(working: &Working<'_>, style: u64, stylesheet: u64, list: u64) -> Result<bool, Error> {
    let limits = wire_limits(working)?;
    for object in working.source().state.components.iter_objects() {
        let identifier = object.archive_info.identifier.unwrap_or_default();
        if identifier == style || identifier == stylesheet || identifier == list {
            continue;
        }
        if object.archive_info.message_infos.iter().any(|info| {
            info.object_references.contains(&style)
                || info
                    .field_infos
                    .iter()
                    .any(|field| field.object_references.contains(&style))
        }) {
            return Ok(false);
        }
        for message in &object.messages {
            if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                && litchi_iwa_text_wire::character_style_parent(&message.data, limits)
                    .ok()
                    .flatten()
                    == Some(style)
            {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

/// Bring `info`'s field infos in line with the variation's managed paths
/// (the host editor's `sync_managed_field_info`).
fn sync_managed(
    infos: &mut Vec<FieldInfo>,
    managed: ManagedFields,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    for (present, path) in [
        (managed.text_fill(), [11, 46]),
        (managed.text_stroke(), [11, 44]),
        (managed.text_shadow(), [11, 21]),
        (managed.text_background(), [11, 26]),
    ] {
        let is_path = |field: &FieldInfo| field.path.as_slice() == path;
        if present {
            if !infos.iter().any(is_path) {
                let mut info = FieldInfo::new(
                    FieldPath::new(path.to_vec(), limits).map_err(|_| Error::InvalidSource)?,
                    limits,
                )
                .map_err(|_| Error::InvalidSource)?;
                info.r#type = Some(FieldType::Message);
                info.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserve);
                infos
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation { amount: 1 })?;
                infos.push(info);
            }
        } else {
            infos.retain(|field| !is_path(field));
        }
    }
    Ok(())
}

/// Point the stored cell at `position` at style key `key` (`None` removes
/// the cell's key).
fn write_key(
    working: &mut Working<'_>,
    store: &[u8],
    columns: u32,
    position: At,
    key: Option<u32>,
) -> Result<(), Error> {
    let Some(located) = cells::locate_cell(working, store, position).map_err(from_cells)? else {
        let Some(key) = key else {
            return Ok(());
        };
        return cells::materialize(working, store, columns, position, key).map_err(from_cells);
    };
    let limits = crate::cell::wire::limits_for_source(&located.bytes);
    let mut cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &limits)
        .map_err(|_error| Error::InvalidSource)?;
    cell.set_text_style_identifier(key);
    let encoded = crate::cell::wire::encode(&cell, &limits).map_err(|_| Error::InvalidSource)?;
    let (_, tile_data) =
        topology::message(working, located.tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
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
    cells::replace_message(working, located.tile, located.tile_index, rewritten)?;
    Ok(())
}

/// A style list with one more reference to the entry naming `style`
/// (appended with a fresh key when there is none), and that key.
fn list_with_reference(list: &[u8], style: u64) -> Result<(Vec<u8>, u32, bool), Error> {
    for entry in repeated_bytes(list, 3)? {
        if single_reference(entry, 4)? == Some(style) {
            let key = u32::try_from(single_varint(entry, 1)?.ok_or(Error::InvalidSource)?)
                .map_err(|_| Error::InvalidSource)?;
            let rewritten = rebuild(list, |field| {
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
            return Ok((rewritten, key, false));
        }
    }
    let key = u32::try_from(single_varint(list, 2)?.unwrap_or(1))
        .map_err(|_error| Error::InvalidSource)?
        .max(1);
    let mut entry = Vec::new();
    push_varint_field(&mut entry, 1, u64::from(key))?;
    push_varint_field(&mut entry, 2, 1)?;
    let mut reference = Vec::new();
    push_varint_field(&mut reference, 1, style)?;
    push_bytes_field(&mut entry, 4, &reference)?;
    let mut rebuilt = rebuild(list, |field| {
        if field.number == 2 {
            return varint_field(2, u64::from(key) + 1).map(Some);
        }
        Ok(None)
    })?;
    if single_varint(&rebuilt, 2)?.is_none() {
        return Err(Error::InvalidSource);
    }
    push_bytes_field(&mut rebuilt, 3, &entry)?;
    Ok((rebuilt, key, true))
}

/// Replace the stylesheet payload, appending or removing `style` in its
/// aggregate references (a header the core cannot transition keeps its
/// field-local references and takes the aggregate change directly).
fn replace_stylesheet(
    working: &mut Working<'_>,
    stylesheet: u64,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
) -> Result<(), Error> {
    match cells::replace_with_references(
        working,
        stylesheet,
        index,
        data.clone(),
        removed,
        appended,
    ) {
        Ok(()) => Ok(()),
        Err(Fault::Invalid) => {
            cells::replace_message(working, stylesheet, index, data)?;
            let object = working.object_mut(stylesheet)?;
            let references = &mut object
                .archive_info
                .message_infos
                .get_mut(index)
                .ok_or(Error::InvalidSource)?
                .object_references;
            references.retain(|identifier| !removed.contains(identifier));
            for identifier in appended {
                if !references.contains(identifier) {
                    references
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation { amount: 1 })?;
                    references.push(*identifier);
                }
            }
            Ok(())
        },
        Err(other) => Err(other.into()),
    }
}

#[allow(
    clippy::too_many_lines,
    reason = "one staged change: plan, then the rewrite, derive or collapse publication"
)]
fn apply_one(
    source: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    change: &Change,
) -> Result<Option<Published>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let wire = wire_limits(&working)?;
    let table = table(&working, sheet, table_index)?;
    if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let (list_index, list_data) =
        topology::message(&working, table.list, &[LIST_MESSAGE_TYPE, 6_201])
            .map_err(Error::from)?;
    let list_data = list_data.to_vec();
    let old_key = match cells::locate_cell(&working, &table.store, position).map_err(from_cells)? {
        Some(located) => litchi_numbers_wire::BncCell::parse(
            &located.bytes,
            &crate::cell::wire::limits_for_source(&located.bytes),
        )
        .map_err(|_error| Error::InvalidSource)?
        .text_style_identifier(),
        None => None,
    };
    let base = cells::base_style(&table.model, table.geometry, position).map_err(from_cells)?;
    let current = match old_key {
        Some(key) => cells::list_reference(&working, &list_data, key)
            .map_err(from_cells)?
            .ok_or(Error::InvalidSource)?,
        None => base,
    };
    let stylesheet = cells::stylesheet_of(&working, current).map_err(from_cells)?;
    let is_exclusive = match old_key {
        Some(key) => {
            entry_refcount(&list_data, key)? == Some(1)
                && exclusive(&working, current, stylesheet, table.list)?
        },
        None => false,
    };
    let plan = engine::plan(
        &WorkingStyles(&working),
        Target::new(current, is_exclusive),
        change,
        wire,
    )
    .map_err(from_engine)?;
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    match plan {
        Plan::Unchanged => return Ok(None),
        Plan::Rewrite(variation) => {
            let index = structure::unique_message(
                working.object(current)?,
                &[PARAGRAPH_STYLE_MESSAGE_TYPE],
            )?;
            let before = working.object(current)?.archive_info.message_infos[index]
                .object_references
                .clone();
            let after = variation.references().to_vec();
            let removed: Vec<u64> = before
                .iter()
                .copied()
                .filter(|reference| *reference != stylesheet && !after.contains(reference))
                .collect();
            let appended: Vec<u64> = after
                .iter()
                .copied()
                .filter(|reference| !before.contains(reference))
                .collect();
            let managed = variation.managed_fields();
            cells::replace_with_references(
                &mut working,
                current,
                index,
                variation.into_payload(),
                &removed,
                &appended,
            )?;
            let limits = working.archive_limits();
            let object = working.object_mut(current)?;
            sync_managed(
                &mut object.archive_info.message_infos[index].field_infos,
                managed,
                limits,
            )?;
        },
        Plan::Derive(variation) => {
            let sheet_member = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
            if working.locate(current)?.as_deref() != Some(sheet_member.as_str()) {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let list_member = working.locate(table.list)?.ok_or(Error::InvalidSource)?;
            let identifier = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
            let managed = variation.managed_fields();
            let references = variation.references().to_vec();
            let payload =
                cells::with_paragraph_properties(variation.into_payload()).map_err(from_cells)?;
            let template = working.object(current)?;
            let mut object = structure::new_object(
                identifier,
                Some(template),
                PARAGRAPH_STYLE_MESSAGE_TYPE,
                payload,
                references,
                working.archive_limits(),
            )?;
            sync_managed(
                &mut object.archive_info.message_infos[0].field_infos,
                managed,
                working.archive_limits(),
            )?;
            working.insert(&sheet_member, object)?;
            let (sheet_index, sheet_data) =
                topology::message(&working, stylesheet, &[STYLESHEET_MESSAGE_TYPE])
                    .map_err(Error::from)?;
            let sheet_data =
                engine::stylesheet_with_variation(sheet_data, current, identifier, wire)
                    .map_err(from_engine)?;
            replace_stylesheet(
                &mut working,
                stylesheet,
                sheet_index,
                sheet_data,
                &[],
                &[identifier],
            )?;
            let sheet_component = registry.component_for_entry(&sheet_member)?.identifier;
            let list_component = registry.component_for_entry(&list_member)?.identifier;
            let seed = metadata::identity_seed(&working, identifier)?;
            transition.register_uuid(&registry, sheet_component, identifier, seed)?;
            transition.reference(&registry, list_component, sheet_component, identifier)?;
            let (mut list, key, _) = list_with_reference(&list_data, identifier)?;
            let mut pruned = Vec::new();
            if let Some(old) = old_key {
                let mut remaining = std::collections::BTreeMap::from([(old, 1)]);
                let (released, dropped) =
                    topology::lists::release(&list, &mut remaining).map_err(Fault::from)?;
                if !remaining.is_empty() {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                if let Some(released) = released {
                    list = released;
                }
                pruned = dropped;
            }
            cells::replace_with_references(
                &mut working,
                table.list,
                list_index,
                list,
                &pruned,
                &[identifier],
            )?;
            write_key(
                &mut working,
                &table.store,
                table.geometry.columns,
                position,
                Some(key),
            )?;
        },
        Plan::Collapse { parent, stylesheet } => {
            let old = old_key.ok_or(Error::InvalidSource)?;
            // The cell points at the parent; the base style needs no key.
            let (mut list, new_key, appended) = if parent == base {
                (list_data.clone(), None, false)
            } else {
                let (list, key, appended) = list_with_reference(&list_data, parent)?;
                (list, Some(key), appended)
            };
            let mut remaining = std::collections::BTreeMap::from([(old, 1)]);
            let (released, pruned) =
                topology::lists::release_with(&list, &mut remaining, true).map_err(Fault::from)?;
            if !remaining.is_empty() {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            if let Some(released) = released {
                list = released;
            }
            let added: &[u64] = if appended { &[parent] } else { &[] };
            if new_key.is_some() {
                // The list names the parent across components, as the host
                // registers it.
                let list_member = working.locate(table.list)?.ok_or(Error::InvalidSource)?;
                let parent_member = working.locate(parent)?.ok_or(Error::InvalidSource)?;
                let list_component = registry.component_for_entry(&list_member)?.identifier;
                let parent_component = registry.component_for_entry(&parent_member)?.identifier;
                transition.reference(&registry, list_component, parent_component, parent)?;
            }
            cells::replace_with_references(
                &mut working,
                table.list,
                list_index,
                list,
                &pruned,
                added,
            )?;
            write_key(
                &mut working,
                &table.store,
                table.geometry.columns,
                position,
                new_key,
            )?;
            // The variation goes with its stylesheet registration.
            let (sheet_index, sheet_data) =
                topology::message(&working, stylesheet, &[STYLESHEET_MESSAGE_TYPE])
                    .map_err(Error::from)?;
            let sheet_data =
                engine::stylesheet_without_variation(sheet_data, parent, current, wire)
                    .map_err(from_engine)?;
            replace_stylesheet(
                &mut working,
                stylesheet,
                sheet_index,
                sheet_data,
                &[current],
                &[],
            )?;
            transition.remove_objects(&registry, &[current])?;
            transition.release_suffix(&[current])?;
            working.remove(current)?;
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(Some(working.publish(&transition)?))
}

impl Package {
    /// The effective value of `kind` for one cell's text (its text style and
    /// that style's parents; a cell without its own style uses its section's
    /// base style).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], source errors.
    pub fn table_cell_text<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let (sheet, table) = self.resolve_cell_text_target(sheet.into(), table.into())?;
        text(self, sheet, table, position, kind)
    }

    /// Start an immutable cell text-property edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::text::{Property, PropertyKind};
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let current = package.table_cell_text(0, 0, At::new(1, 1), PropertyKind::Alignment)?;
    /// let commit = package
    ///     .edit_table_cell_text(0, 0)?
    ///     .reset(At::new(1, 1), PropertyKind::Alignment)?
    ///     .commit()?;
    /// # let _ = (current, commit);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_text<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_cell_text_target(sheet.into(), table.into())?;
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
            changes: Vec::new(),
        })
    }

    /// Apply a cell text patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_text(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_cell_text_target(
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
        let table = topology::table_position(self, sheet, table).map_err(Error::from)?;
        Ok((sheet, table))
    }
}
