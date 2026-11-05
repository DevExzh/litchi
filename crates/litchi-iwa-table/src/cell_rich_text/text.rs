//! One character or paragraph property of whole cells' text (lifted from
//! `litchi-numbers`, record 0370).
//!
//! A staged change runs the shared paragraph-style engine
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
//! This is the host editor's `cell_paragraph_style` behaviour, which the
//! host runs on Numbers sheets, Pages body tables and Keynote slide tables
//! alike.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_core::{FieldInfo, FieldPath, FieldType, UnknownFieldRule};
pub use litchi_iwa_text_wire::paragraph_style::Property;
use litchi_iwa_text_wire::paragraph_style::{
    self as engine, Change, ManagedFields, Plan, PropertyKind, Styles, Target,
};

use super::{
    Error, LIST_MESSAGE_TYPE, PARAGRAPH_STYLE_MESSAGE_TYPE, Refusal, STYLE_LIST_FIELD,
    STYLESHEET_MESSAGE_TYPE, Table, base_style, cell_limits, current_style, list_reference,
    locate_cell, materialize_cell, message, replace_message, replace_with_references, rewrite_cell,
    style_wire_limits, stylesheet_of, table, with_paragraph_properties,
};
use crate::structure::metadata::{self, Transition};
use crate::structure::wire::{
    bytes_field, push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_reference,
    single_varint, varint_field,
};
use crate::structure::{self, Fault, Published, TableLeaf, Working};

/// The working set's paragraph styles, as the engine reads them.
struct WorkingStyles<'w, 'p, P>(&'w Working<'p, P>);

impl<P: TableLeaf> Styles for WorkingStyles<'_, '_, P> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: super::LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// The effective `kind` of the cell at `position` of the table whose model
/// is `model`.
///
/// # Errors
///
/// [`Refusal::CellNotFound`] outside the table; source and wire errors.
pub fn read<P: TableLeaf>(
    working: &Working<'_, P>,
    model: u64,
    position: At,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let style = current_style(working, model, position)?;
    engine::effective(
        &WorkingStyles(working),
        style,
        kind,
        style_wire_limits(working)?,
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
fn exclusive<P: TableLeaf>(
    working: &Working<'_, P>,
    style: u64,
    stylesheet: u64,
    list: u64,
) -> Result<bool, Error> {
    let limits = style_wire_limits(working)?;
    for object in working
        .source()
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
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

/// Point the cell at `position` at style key `key` (`None` removes the
/// cell's key); an empty cell takes a style-only cell.
fn write_key<P: TableLeaf>(
    working: &mut Working<'_, P>,
    table: &Table,
    position: At,
    key: Option<u32>,
) -> Result<(), Error> {
    let Some(located) = locate_cell(working, &table.store, position)? else {
        let Some(key) = key else {
            return Ok(());
        };
        let mut cell = litchi_numbers_wire::BncCell::minimal();
        cell.set_text_style_identifier(Some(key));
        return materialize_cell(
            working,
            &table.store,
            table.geometry.columns,
            position,
            &cell,
        );
    };
    rewrite_cell(working, &located, position, |cell| {
        cell.set_text_style_identifier(key);
    })
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
fn replace_stylesheet<P: TableLeaf>(
    working: &mut Working<'_, P>,
    stylesheet: u64,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
) -> Result<(), Error> {
    match replace_with_references(working, stylesheet, index, data.clone(), removed, appended) {
        Ok(()) => Ok(()),
        Err(Fault::Invalid) => {
            replace_message(working, stylesheet, index, data)?;
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

/// Apply one staged change to the cell at `position` of the table whose
/// model is `model`, publishing the working set. `None` is an exact no-op.
///
/// The working set must be fresh over the leaf's exact source; it keeps the
/// root previews, as the host editor does for cell edits.
///
/// # Errors
///
/// [`Refusal::CellNotFound`]; [`Refusal::UnsupportedDependency`] for a
/// style graph the engine does not rewrite; [`Refusal::UnsupportedSource`]
/// without a package metadata registry; source, resource and candidate
/// errors.
#[allow(
    clippy::too_many_lines,
    reason = "one staged change: plan, then the rewrite, derive or collapse publication"
)]
pub fn apply<P: TableLeaf>(
    mut working: Working<'_, P>,
    model: u64,
    position: At,
    change: &Change,
) -> Result<Option<Published<P>>, Error> {
    working.keep_previews();
    let wire: WireLimits = style_wire_limits(&working)?;
    let table = table(&working, model)?;
    if !table.geometry.contains(position) {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let list = single_reference(&table.store, STYLE_LIST_FIELD)?
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let (list_index, list_data) = message(&working, list, &[LIST_MESSAGE_TYPE, 6_201])?;
    let list_data = list_data.to_vec();
    let old_key = match locate_cell(&working, &table.store, position)? {
        Some(located) => {
            litchi_numbers_wire::BncCell::parse(&located.bytes, &cell_limits(&located.bytes))
                .map_err(|_error| Error::InvalidSource)?
                .text_style_identifier()
        },
        None => None,
    };
    let base = base_style(&table.payload, table.geometry, position)?;
    let current = match old_key {
        Some(key) => list_reference(&working, &list_data, key)?.ok_or(Error::InvalidSource)?,
        None => base,
    };
    let stylesheet = stylesheet_of(&working, current)?;
    let is_exclusive = match old_key {
        Some(key) => {
            entry_refcount(&list_data, key)? == Some(1)
                && exclusive(&working, current, stylesheet, list)?
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
            replace_with_references(
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
            let list_member = working.locate(list)?.ok_or(Error::InvalidSource)?;
            let identifier = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
            let managed = variation.managed_fields();
            let references = variation.references().to_vec();
            let payload = with_paragraph_properties(variation.into_payload())?;
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
                message(&working, stylesheet, &[STYLESHEET_MESSAGE_TYPE])?;
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
            let (mut rewritten, key, _) = list_with_reference(&list_data, identifier)?;
            let mut pruned = Vec::new();
            if let Some(old) = old_key {
                let mut remaining = std::collections::BTreeMap::from([(old, 1)]);
                let (released, dropped) =
                    litchi_numbers_wire::table_topology::lists::release(&rewritten, &mut remaining)
                        .map_err(Fault::from)?;
                if !remaining.is_empty() {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                if let Some(released) = released {
                    rewritten = released;
                }
                pruned = dropped;
            }
            replace_with_references(
                &mut working,
                list,
                list_index,
                rewritten,
                &pruned,
                &[identifier],
            )?;
            write_key(&mut working, &table, position, Some(key))?;
        },
        Plan::Collapse { parent, stylesheet } => {
            let old = old_key.ok_or(Error::InvalidSource)?;
            // The cell points at the parent; the base style needs no key.
            let (mut rewritten, new_key, appended) = if parent == base {
                (list_data.clone(), None, false)
            } else {
                let (rewritten, key, appended) = list_with_reference(&list_data, parent)?;
                (rewritten, Some(key), appended)
            };
            let mut remaining = std::collections::BTreeMap::from([(old, 1)]);
            let (released, pruned) = litchi_numbers_wire::table_topology::lists::release_with(
                &rewritten,
                &mut remaining,
                true,
            )
            .map_err(Fault::from)?;
            if !remaining.is_empty() {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            if let Some(released) = released {
                rewritten = released;
            }
            let added: &[u64] = if appended { &[parent] } else { &[] };
            if new_key.is_some() {
                // The list names the parent across components, as the host
                // registers it.
                let list_member = working.locate(list)?.ok_or(Error::InvalidSource)?;
                let parent_member = working.locate(parent)?.ok_or(Error::InvalidSource)?;
                let list_component = registry.component_for_entry(&list_member)?.identifier;
                let parent_component = registry.component_for_entry(&parent_member)?.identifier;
                transition.reference(&registry, list_component, parent_component, parent)?;
            }
            replace_with_references(&mut working, list, list_index, rewritten, &pruned, added)?;
            write_key(&mut working, &table, position, new_key)?;
            // The variation goes with its stylesheet registration.
            let (sheet_index, sheet_data) =
                message(&working, stylesheet, &[STYLESHEET_MESSAGE_TYPE])?;
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
