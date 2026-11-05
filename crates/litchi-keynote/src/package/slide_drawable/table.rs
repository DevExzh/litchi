//! A new, empty slide table cloned from the structure of the deck's first
//! slide table (the template), record 0329.
//!
//! The payloads come from `litchi_numbers_wire::table_create` (the payload
//! half of `litchi-numbers`' table creation, record 0301): the template's
//! drawable and model, and one emptied copy of each storage object the
//! model owns, with styles shared. This module places them: every clone
//! goes into its source's package member, the drawable is appended to the
//! slide's owned drawables and z-order, the table gets its own formula
//! owner in the calculation engine, the clones of registered objects get
//! registry UUIDs, and every new cross-member reference is registered as an
//! external reference. No clock or randomness is consulted.

use std::collections::{BTreeMap, HashMap};

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_numbers_wire::table_create::{
    self, CALCULATION_ENGINE_MESSAGE_TYPE, FORMULA_OWNER_MESSAGE_TYPE, FORMULA_OWNER_VERSIONS,
    Frame, OwnerGeometry, Role, Shape, TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE,
    TableUuid,
};
use sha1::{Digest, Sha1};

use super::super::structure::metadata::ExternalReference;
use super::super::structure::workspace::Workspace;
use super::{Error, Geometry, Unsupported, lift};

/// What a new table needs from the template and the request.
pub(super) struct Request<'r> {
    pub(super) slide: u64,
    pub(super) template: u64,
    pub(super) name: &'r str,
    pub(super) rows: u32,
    pub(super) columns: u32,
    pub(super) geometry: Geometry,
}

fn storage_error(_error: litchi_numbers_wire::table_write::WriteError) -> Error {
    Error::Refused(crate::slide::drawable::Refusal::Unsupported {
        reason: Unsupported::TableStorage,
    })
}

/// The single payload of `message_type` of `object`.
fn single(object: &ArchiveObject, message_type: u32) -> Result<&[u8], Error> {
    match object.messages.as_slice() {
        [message] if message.type_ == message_type => Ok(&message.data),
        _ => Err(Error::Refused(
            crate::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::TableStorage,
            },
        )),
    }
}

/// A new single-message object whose header copies `template`'s message
/// versions and merge flag and declares `references`.
fn fresh(
    workspace: &Workspace<'_>,
    identifier: u64,
    template: Option<&ArchiveObject>,
    message_type: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        workspace.archive_limits(),
    )
    .map_err(|error| lift(error.into()))?;
    if let Some(template) = template {
        object.archive_info.should_merge = template.archive_info.should_merge;
        if let Some(info) = template.archive_info.message_infos.first() {
            object.archive_info.message_infos[0].versions = info.versions.clone();
        }
    }
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

/// A deterministic 64-bit seed for the new table's identities.
fn seed(info: u64, registry_watermark: u64) -> u64 {
    let mut digest = Sha1::new();
    digest.update(b"litchi.keynote.slide.table.create.v1");
    digest.update(info.to_le_bytes());
    digest.update(registry_watermark.to_le_bytes());
    let digest = digest.finalize();
    let mut bytes = [0_u8; 8];
    bytes.copy_from_slice(&digest[..8]);
    u64::from_le_bytes(bytes)
}

/// The calculation engine (identifier, message index), when the deck has
/// one.
fn engine(workspace: &Workspace<'_>) -> Result<Option<(u64, usize)>, Error> {
    let mut found = None;
    for component in workspace.package().state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
                    if found.replace((identifier, index)).is_some() {
                        return Err(Error::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(found)
}

/// Every formula owner payload staged in `member`.
fn owners_in<'w>(workspace: &'w Workspace<'_>, member: &str) -> Vec<&'w [u8]> {
    workspace
        .member_objects(member)
        .into_iter()
        .filter_map(|identifier| workspace.object(identifier))
        .flat_map(|object| object.messages.iter())
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .collect()
}

/// Register `target` as referenced from the component of `from`, unless
/// both live in one member.
fn reference(workspace: &mut Workspace<'_>, from: u64, target: u64) -> Result<(), Error> {
    let from_member = workspace.component_of(from).ok_or(Error::InvalidSource)?;
    let target_member = workspace.component_of(target).ok_or(Error::InvalidSource)?;
    if from_member == target_member {
        return Ok(());
    }
    let source = workspace
        .registry
        .component_for_member(&from_member)
        .ok_or(Error::Refused(
            crate::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::RegistryShape,
            },
        ))?
        .identifier;
    let destination = workspace
        .registry
        .component_for_member(&target_member)
        .ok_or(Error::Refused(
            crate::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::RegistryShape,
            },
        ))?
        .identifier;
    let entry = workspace
        .registry
        .component_mut(source)
        .ok_or(Error::InvalidSource)?;
    // Native packages name the component alone when it is the target's own.
    let covered = entry.external_references().iter().any(|reference| {
        reference.component == destination
            && (reference.object == Some(target)
                || (reference.object.is_none() && destination == target))
    });
    if !covered {
        entry.add_external_reference(ExternalReference {
            component: destination,
            object: (destination != target).then_some(target),
            is_weak: None,
        });
    }
    Ok(())
}

/// Add the table and return its drawable identifier.
pub(super) fn add_table(
    workspace: &mut Workspace<'_>,
    request: &Request<'_>,
) -> Result<u64, Error> {
    if !table_create::admits_size(request.rows, request.columns) {
        return Err(Error::Refused(
            crate::slide::drawable::Refusal::InvalidTableSize {
                rows: request.rows,
                columns: request.columns,
            },
        ));
    }
    let info_object = workspace.require(request.template).map_err(lift)?.clone();
    let info_template = single(&info_object, TABLE_INFO_MESSAGE_TYPE)?.to_vec();
    let model = table_create::single_reference(&info_template, 2)
        .map_err(storage_error)?
        .ok_or(Error::InvalidSource)?;
    let template_parent = table_create::drawable_parent(&info_template).map_err(storage_error)?;
    let model_object = workspace.require(model).map_err(lift)?.clone();
    let model_template = single(&model_object, TABLE_MODEL_MESSAGE_TYPE)?.to_vec();
    let owned: BTreeMap<u64, Role> =
        table_create::owned_storage(&model_template).map_err(storage_error)?;

    // Fresh identities.
    let watermark = workspace.registry.last_object_identifier();
    let info = workspace.allocate().map_err(lift)?;
    let new_model = workspace.allocate().map_err(lift)?;
    let owner = workspace.allocate().map_err(lift)?;
    let mut remap: HashMap<u64, u64> = HashMap::with_capacity(owned.len() + 2);
    for identifier in owned.keys() {
        remap.insert(*identifier, workspace.allocate().map_err(lift)?);
    }
    let seed = seed(info, watermark);
    let uuid = TableUuid::derive(seed);

    // The storage clones, each in its source's member.
    let mut inserted: Vec<(u64, u64)> = Vec::with_capacity(owned.len() + 3);
    for (identifier, role) in &owned {
        let object = workspace.require(*identifier).map_err(lift)?.clone();
        let [message] = object.messages.as_slice() else {
            return Err(Error::Refused(
                crate::slide::drawable::Refusal::Unsupported {
                    reason: Unsupported::TableStorage,
                },
            ));
        };
        let member = workspace
            .component_of(*identifier)
            .ok_or(Error::InvalidSource)?;
        let data = table_create::empty_storage(
            *role,
            &message.data,
            request.rows,
            request.columns,
            table_create::mix(seed ^ remap[identifier]),
        )
        .map_err(storage_error)?;
        let clone = fresh(
            workspace,
            remap[identifier],
            Some(&object),
            message.type_,
            data,
            Vec::new(),
        )?;
        workspace.insert(&member, clone).map_err(lift)?;
        inserted.push((remap[identifier], *identifier));
    }

    // The model, with its styles shared and its storage remapped.
    let shape = Shape {
        name: request.name,
        rows: request.rows,
        columns: request.columns,
        row_height: Some(f64::from(request.geometry.height) / f64::from(request.rows)),
        column_width: Some(f64::from(request.geometry.width) / f64::from(request.columns)),
    };
    let model_data = table_create::model_payload(
        &model_template,
        &|identifier| remap.get(&identifier).copied(),
        &uuid,
        &shape,
    )
    .map_err(storage_error)?;
    let geometry =
        OwnerGeometry::of(&model_data, request.rows, request.columns).map_err(storage_error)?;
    let model_references = declared(&model_object, &model_data, &|identifier| {
        remap.get(&identifier).copied()
    });
    let model_member = workspace.component_of(model).ok_or(Error::InvalidSource)?;
    let clone = fresh(
        workspace,
        new_model,
        Some(&model_object),
        TABLE_MODEL_MESSAGE_TYPE,
        model_data,
        model_references,
    )?;
    workspace.insert(&model_member, clone).map_err(lift)?;
    inserted.push((new_model, model));

    // The drawable, parented to the slide.
    let frame = Frame {
        x: request.geometry.x,
        y: request.geometry.y,
        width: request.geometry.width,
        height: request.geometry.height,
    };
    let info_data = table_create::info_payload(&info_template, request.slide, new_model, frame)
        .map_err(|_error| Error::Refused(crate::slide::drawable::Refusal::InvalidGeometry))?;
    let info_references = declared(&info_object, &info_data, &|identifier| {
        if identifier == model {
            Some(new_model)
        } else if Some(identifier) == template_parent {
            Some(request.slide)
        } else {
            None
        }
    });
    let info_member = workspace
        .component_of(request.template)
        .ok_or(Error::InvalidSource)?;
    let clone = fresh(
        workspace,
        info,
        Some(&info_object),
        TABLE_INFO_MESSAGE_TYPE,
        info_data,
        info_references,
    )?;
    workspace.insert(&info_member, clone).map_err(lift)?;
    inserted.push((info, request.template));
    super::author::attach(workspace, request.slide, info).map_err(lift)?;
    reference(workspace, request.slide, info)?;

    // The table's own formula owner, when the deck has an engine.
    if let Some((engine_id, engine_index)) = engine(workspace)? {
        let engine_member = workspace
            .component_of(engine_id)
            .ok_or(Error::InvalidSource)?;
        let engine_object = workspace.require(engine_id).map_err(lift)?;
        let engine_data = engine_object.messages[engine_index].data.clone();
        let aggregate = engine_object.archive_info.message_infos[engine_index]
            .object_references
            .clone();
        let internal =
            table_create::next_internal_owner(&owners_in(workspace, &engine_member), &engine_data)
                .map_err(storage_error)?;
        let uid = uuid.owner_uid();
        let owner_data =
            table_create::owner_payload(uid, internal, info, geometry).map_err(storage_error)?;
        let mut owner_object = fresh(
            workspace,
            owner,
            None,
            FORMULA_OWNER_MESSAGE_TYPE,
            owner_data,
            Vec::new(),
        )?;
        owner_object.archive_info.message_infos[0].versions =
            FORMULA_OWNER_VERSIONS.to_vec().into();
        workspace
            .insert(&engine_member, owner_object)
            .map_err(lift)?;
        let engine_data = table_create::engine_payload(&engine_data, owner, internal, uid)
            .map_err(storage_error)?;
        let mut after = aggregate.clone();
        after.push(owner);
        let limits = workspace.archive_limits();
        workspace
            .object_mut(engine_id)
            .map_err(lift)?
            .replace_message_transitioning_object_references_preserving_header(
                engine_index,
                RawMessage {
                    type_: CALCULATION_ENGINE_MESSAGE_TYPE,
                    data: engine_data,
                },
                litchi_iwa_core::archive::ObjectReferenceTransition {
                    aggregate_before: &aggregate,
                    aggregate_after: &after,
                    fields: &[],
                },
                limits,
            )
            .map_err(|error| lift(error.into()))?;
        // The owner names the drawable through its `formula_owner`.
        reference(workspace, owner, info)?;
    }

    // Registry: UUIDs for the clones of registered objects, and every new
    // cross-member reference.
    for (clone, source) in &inserted {
        let member = workspace.component_of(*clone).ok_or(Error::InvalidSource)?;
        let component = workspace
            .registry
            .component_for_member(&member)
            .ok_or(Error::Refused(
                crate::slide::drawable::Refusal::Unsupported {
                    reason: Unsupported::RegistryShape,
                },
            ))?;
        let (identifier, source_uuid) = (component.identifier, component.uuid_of(*source));
        if let Some(source_uuid) = source_uuid {
            let uuid = workspace
                .clone_uuid(Some(source_uuid), *clone)
                .map_err(lift)?;
            workspace
                .registry
                .component_mut(identifier)
                .ok_or(Error::InvalidSource)?
                .add_object_uuid(*clone, uuid);
        }
        let targets = workspace
            .require(*clone)
            .map_err(lift)?
            .archive_info
            .message_infos[0]
            .object_references
            .clone();
        for target in targets {
            reference(workspace, *clone, target)?;
        }
    }
    Ok(info)
}

/// The references a clone declares: its template's declared references,
/// mapped through `remap` (the template's own storage, model or parent),
/// kept only where the clone's payload holds them.
fn declared(template: &ArchiveObject, data: &[u8], remap: &dyn Fn(u64) -> Option<u64>) -> Vec<u64> {
    let mut output = Vec::new();
    for identifier in template
        .archive_info
        .message_infos
        .first()
        .map(|info| info.object_references.as_slice())
        .unwrap_or_default()
    {
        let mapped = remap(*identifier).unwrap_or(*identifier);
        if table_create::references_identifier(data, mapped) && !output.contains(&mapped) {
            output.push(mapped);
        }
    }
    output
}
