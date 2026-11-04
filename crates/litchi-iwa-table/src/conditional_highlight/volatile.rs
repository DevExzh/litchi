//! The calculation-engine owner family of a volatile rule (record 0376a):
//! the host's `conditional_highlight::dependencies::{ensure_volatile_owner,
//! remove_volatile_host}` through the shared
//! `litchi_numbers_wire::conditional_volatile` planner.

use litchi_iwa_core::{FieldInfo, FieldPath, FieldType, UnknownFieldRule};
use litchi_numbers_wire::conditional_volatile::{self as planner, Object, Source};
use litchi_numbers_wire::formula_owner::Uuid;

use super::{Change, Error, FORMULA_OWNER_MESSAGE_TYPE, Result, Table, conditional_owner};
use crate::structure::{self, TableLeaf, Working, metadata};
use litchi_iwa_common::table::coordinate::At;

const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;

fn first_message<'w, P: TableLeaf>(
    working: &'w Working<'_, P>,
    identifier: u64,
) -> Option<&'w [u8]> {
    working
        .object(identifier)
        .ok()
        .and_then(|object| object.messages.first())
        .map(|message| message.data.as_slice())
}

/// The engine and owner object identifiers of the source.
fn graph<P: TableLeaf>(working: &Working<'_, P>) -> Result<Option<(u64, Vec<u64>)>> {
    let mut engine = None;
    let mut owners = Vec::new();
    for object in working
        .source()
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
        let Some(message) = object.messages.first() else {
            continue;
        };
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
            structure::reserve(&mut owners, 1)?;
            owners.push(identifier);
        } else if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE
            && engine.replace(identifier).is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(engine.map(|engine| (engine, owners)))
}

fn plan<P: TableLeaf>(
    change: &Change<'_, '_, P>,
    table: &Table,
    position: At,
    ensure: bool,
) -> Result<Option<(u64, planner::Edit)>> {
    let working = &*change.working;
    let (_, model) =
        super::topology::message(working, table.model, &[6_001]).map_err(Error::from)?;
    let Some(owner) = conditional_owner(model)? else {
        return Ok(None);
    };
    let Some((engine_id, owner_ids)) = graph(working)? else {
        return Ok(None);
    };
    let mut owners = Vec::new();
    for identifier in &owner_ids {
        structure::reserve(&mut owners, 1)?;
        owners.push(Object {
            identifier: *identifier,
            payload: first_message(working, *identifier).ok_or(Error::InvalidSource)?,
        });
    }
    let source = Source {
        engine: Object {
            identifier: engine_id,
            payload: first_message(working, engine_id).ok_or(Error::InvalidSource)?,
        },
        owners: &owners,
        table: table.info,
        conditional_owner: Uuid::new(owner.lower, owner.upper),
        row: position.row(),
        column: position.column(),
        creates_first_tile: P::CREATES_FIRST_VOLATILE_TILE,
    };
    let object = |identifier: u64| first_message(working, identifier);
    let limits = litchi_numbers_wire::Limits::DEFAULT;
    let edit = if ensure {
        // The host draws the global-name owner's UUID at random; the leaf
        // derives it from the table's conditional owner.
        let global_name = Uuid::new(
            metadata::mix(owner.lower ^ 0x7661_6c61_7469_6c65),
            metadata::mix(owner.upper ^ 0x6772_6f75_7069_6e67),
        );
        planner::ensure(
            source,
            working.next_identifier(),
            global_name,
            &object,
            &limits,
        )
    } else {
        planner::remove(source, &object, &limits)
    }
    .map_err(|error| match error {
        planner::Error::MissingTableOwner | planner::Error::IncompleteFamily => {
            Error::Refused(super::Refusal::UnsupportedDependency)
        },
        _ => Error::InvalidSource,
    })?;
    Ok(Some((engine_id, edit)))
}

fn publish<P: TableLeaf>(
    change: &mut Change<'_, '_, P>,
    engine_id: u64,
    edit: planner::Edit,
) -> Result<()> {
    let member = change
        .working
        .locate(engine_id)?
        .ok_or(Error::InvalidSource)?;
    if !edit.created.is_empty() {
        let identifiers = change.working.allocate(edit.created.len())?;
        if identifiers.first()
            != edit
                .created
                .first()
                .map(|created| created.identifier)
                .as_ref()
        {
            return Err(Error::InvalidSource);
        }
        let limits = change.working.archive_limits();
        let component = change.component(&member)?;
        for created in edit.created {
            let version: &[u32] = if created.message_type == FORMULA_OWNER_MESSAGE_TYPE {
                &planner::OWNER_MESSAGE_VERSION
            } else {
                &planner::TILE_MESSAGE_VERSION
            };
            let mut object = structure::new_object(
                created.identifier,
                None,
                created.message_type,
                created.payload,
                created.references,
                limits,
            )?;
            object
                .archive_info
                .message_infos
                .first_mut()
                .ok_or(Error::InvalidSource)?
                .versions = version.to_vec().into();
            change.working.insert(&member, object)?;
            let seed = metadata::identity_seed(change.working, created.identifier)?;
            change.transition.register_uuid(
                change.registry,
                component,
                created.identifier,
                seed,
            )?;
        }
    }
    for (identifier, payload) in edit.payloads {
        let index = 0;
        super::cells::replace_message(change.working, identifier, index, payload)?;
    }
    for (owner, tile) in edit.owner_references {
        let info = change
            .working
            .object_mut(owner)?
            .archive_info
            .message_infos
            .first_mut()
            .ok_or(Error::InvalidSource)?;
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(tile);
        structure::reserve(&mut info.field_infos, 1)?;
        info.field_infos.push(FieldInfo {
            path: FieldPath { path: vec![13, 1] },
            r#type: Some(FieldType::ObjectReference),
            unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserveUntilModified),
            ..Default::default()
        });
    }
    if !edit.engine_references.is_empty() {
        let info = change
            .working
            .object_mut(engine_id)?
            .archive_info
            .message_infos
            .first_mut()
            .ok_or(Error::InvalidSource)?;
        for identifier in edit.engine_references {
            if !info.object_references.contains(&identifier) {
                structure::reserve(&mut info.object_references, 1)?;
                info.object_references.push(identifier);
            }
        }
        if !info
            .field_infos
            .iter()
            .any(|field| field.path.path.as_slice() == [2, 6])
        {
            structure::reserve(&mut info.field_infos, 1)?;
            info.field_infos.push(FieldInfo {
                path: FieldPath { path: vec![2, 6] },
                r#type: Some(FieldType::ObjectReference),
                unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserveUntilModified),
                ..Default::default()
            });
        }
    }
    Ok(())
}

/// Register the volatile cell at `position` (the host's
/// `ensure_volatile_owner`).
pub(super) fn ensure<P: TableLeaf>(
    change: &mut Change<'_, '_, P>,
    table: &Table,
    position: At,
) -> Result<()> {
    if let Some((engine, edit)) = plan(change, table, position, true)? {
        publish(change, engine, edit)?;
    }
    Ok(())
}

/// Unregister the volatile cell at `position`, when it is registered (the
/// host's `remove_volatile_host`).
pub(super) fn remove<P: TableLeaf>(
    change: &mut Change<'_, '_, P>,
    table: &Table,
    position: At,
) -> Result<()> {
    if let Some((engine, edit)) = plan(change, table, position, false)? {
        publish(change, engine, edit)?;
    }
    Ok(())
}
