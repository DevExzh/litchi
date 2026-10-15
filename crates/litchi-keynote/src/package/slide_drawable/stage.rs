//! Staging of geometry, removal and z-order operations in a structural
//! workspace.

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::RawMessage;

use super::super::structure::StructureError;
use super::super::structure::refs;
use super::super::structure::wire::{self, Fields, LENGTH_DELIMITED};
use super::super::structure::workspace::Workspace;
use super::inventory::{
    self, SLIDE, SLIDE_OWNED_DRAWABLES, SLIDE_Z_ORDER, drawable_path, unique_message,
};
use super::{Error, Geometry, Unsupported, lift};

/// Re-encode `payload` with `edit` applied to the message at `path`.
pub(super) fn edit_nested(
    payload: &[u8],
    path: &[u32],
    wire: WireLimits,
    edit: &mut dyn FnMut(&mut Fields) -> Result<(), StructureError>,
) -> Result<Vec<u8>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    if let Some((first, rest)) = path.split_first() {
        let inner = fields
            .unique(*first)?
            .ok_or(StructureError::Unexpected("nested message"))?
            .message()?
            .to_vec();
        let edited = edit_nested(&inner, rest, wire, edit)?;
        fields.set_message(*first, &edited);
    } else {
        edit(&mut fields)?;
    }
    Ok(fields.encode())
}

/// Replace one payload, keeping every untouched header byte.
pub(super) fn replace_payload(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    index: usize,
    message_type: u32,
    data: Vec<u8>,
) -> Result<(), StructureError> {
    let limits = workspace.archive_limits();
    workspace
        .object_mut(identifier)?
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )?;
    Ok(())
}

/// Patch a `TSD.GeometryArchive` field list in place.
pub(super) fn patch_geometry(
    fields: &mut Fields,
    geometry: Geometry,
    wire: WireLimits,
) -> Result<(), StructureError> {
    for (number, first, second) in [
        (1, geometry.x, geometry.y),
        (2, geometry.width, geometry.height),
    ] {
        let mut inner = match fields.unique(number)? {
            Some(field) => Fields::parse(field.message()?, wire)?,
            None => Fields::parse(&[], wire)?,
        };
        inner.set_float(1, first);
        inner.set_float(2, second);
        fields.set_message(number, &inner.encode());
    }
    match geometry.rotation {
        Some(rotation) => fields.set_float(4, rotation),
        None => {
            fields.remove(4);
        },
    }
    Ok(())
}

/// Set the geometry of one drawable payload.
pub(super) fn set_geometry(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    geometry: Geometry,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let payload = workspace
        .require(drawable)
        .map_err(lift)?
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let mut path = drawable_path(message_type).to_vec();
    path.push(1);
    let data = edit_nested(&payload, &path, wire, &mut |fields| {
        patch_geometry(fields, geometry, wire)
    })
    .map_err(lift)?;
    replace_payload(workspace, drawable, message_index, message_type, data).map_err(lift)
}

/// Every object reference held by one staged object: the header's aggregate
/// and field lists, and the schema walk of admitted payloads.
fn outgoing(workspace: &Workspace<'_>, identifier: u64) -> Result<Vec<u64>, StructureError> {
    let object = workspace.require(identifier)?;
    let mut found = BTreeSet::new();
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        found.extend(info.object_references.iter().copied());
        for field in &info.field_infos {
            found.extend(field.object_references.iter().copied());
        }
        if refs::admitted(message.type_) {
            found.extend(refs::object_references(
                message.type_,
                &message.data,
                workspace.wire,
            )?);
        }
    }
    found.remove(&identifier);
    Ok(found.into_iter().collect())
}

const SLIDE_TAGS: u32 = 28;
const SLIDE_INSTRUCTIONAL_TEXT: u32 = 45;

/// Whether message `entry` holds a reference to `drawable` in field `number`.
fn entry_names(
    entry: &[u8],
    number: u32,
    drawable: u64,
    wire: WireLimits,
) -> Result<bool, StructureError> {
    match Fields::parse(entry, wire)?.unique(number)? {
        Some(field) => Ok(wire::reference_identifier(field.message()?, wire)? == drawable),
        None => Ok(false),
    }
}

/// What a removal took out of the package.
pub(super) struct Removal {
    pub(super) objects: usize,
}

/// Remove one drawable and its private object closure from its slide.
pub(super) fn remove(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
) -> Result<Removal, Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    if workspace.component_of(drawable).as_deref() != Some(member.as_str()) {
        return Err(Error::Unsupported {
            reason: Unsupported::ForeignComponent,
        });
    }
    let members = workspace.member_objects(&member);
    let mut graph = HashMap::with_capacity(members.len());
    for identifier in &members {
        graph.insert(*identifier, outgoing(workspace, *identifier).map_err(lift)?);
    }
    // Closure of objects reachable from the drawable inside its component,
    // shrunk until no retained member outside it references a member.
    let mut shared = BTreeSet::from([slide]);
    let closure = loop {
        let mut closure = BTreeSet::from([drawable]);
        let mut pending = vec![drawable];
        while let Some(current) = pending.pop() {
            for target in graph.get(&current).map(Vec::as_slice).unwrap_or(&[]) {
                if graph.contains_key(target) && !shared.contains(target) && closure.insert(*target)
                {
                    pending.push(*target);
                }
            }
        }
        let mut grew = false;
        for (source, targets) in &graph {
            if closure.contains(source) {
                continue;
            }
            for target in targets {
                if !closure.contains(target) {
                    continue;
                }
                if *target == drawable {
                    if *source != slide {
                        return Err(Error::Unsupported {
                            reason: Unsupported::ExternalReferences,
                        });
                    }
                } else if shared.insert(*target) {
                    grew = true;
                }
            }
        }
        if !grew {
            break closure;
        }
    };
    let closure = closure.into_iter().collect::<Vec<_>>();
    let component_id = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    for other in workspace.registry.components() {
        if other.identifier == component_id {
            continue;
        }
        if other.external_references().iter().any(|reference| {
            reference.component == component_id
                && reference
                    .object
                    .is_some_and(|object| closure.contains(&object))
        }) {
            return Err(Error::Unsupported {
                reason: Unsupported::ExternalReferences,
            });
        }
    }

    // The slide may reference the drawable only from its owned list and
    // z-order.
    let slide_object = workspace.require(slide).map_err(lift)?;
    let (slide_index, slide_payload) = unique_message(slide_object, SLIDE).map_err(lift)?;
    let slide_payload = slide_payload.to_vec();
    let mut fields = Fields::parse(&slide_payload, wire).map_err(lift)?;
    let owned = inventory::references(&fields, SLIDE_OWNED_DRAWABLES, wire).map_err(lift)?;
    let z_order = inventory::references(&fields, SLIDE_Z_ORDER, wire).map_err(lift)?;
    // Template tags (`sage_tag_to_info_map`, field 28) and instructional
    // text entries (field 45) that name the drawable go with it.
    let tag_entries = fields
        .all(SLIDE_TAGS)
        .map(|field| field.message().map(<[u8]>::to_vec))
        .collect::<Result<Vec<_>, _>>()
        .map_err(lift)?;
    let mut kept_tags = Vec::with_capacity(tag_entries.len());
    for entry in tag_entries {
        if entry_names(&entry, 2, drawable, wire).map_err(lift)? {
            continue;
        }
        kept_tags.push(entry);
    }
    if kept_tags.len() != fields.all(SLIDE_TAGS).count() {
        fields.replace_all(SLIDE_TAGS, LENGTH_DELIMITED, &kept_tags);
    }
    if let Some(map) = fields.unique(SLIDE_INSTRUCTIONAL_TEXT).map_err(lift)? {
        let mut map = Fields::parse(map.message().map_err(lift)?, wire).map_err(lift)?;
        let entries = map
            .all(1)
            .map(|field| field.message().map(<[u8]>::to_vec))
            .collect::<Result<Vec<_>, _>>()
            .map_err(lift)?;
        let mut kept = Vec::with_capacity(entries.len());
        for entry in &entries {
            if !entry_names(entry, 1, drawable, wire).map_err(lift)? {
                kept.push(entry.clone());
            }
        }
        if kept.len() != entries.len() {
            map.replace_all(1, LENGTH_DELIMITED, &kept);
            fields.set_message(SLIDE_INSTRUCTIONAL_TEXT, &map.encode());
        }
    }
    let walked = refs::object_references(SLIDE, &fields.encode(), wire)
        .map_err(StructureError::from)
        .map_err(lift)?;
    let listed = owned.iter().filter(|id| **id == drawable).count()
        + z_order.iter().filter(|id| **id == drawable).count();
    if owned.iter().filter(|id| **id == drawable).count() != 1
        || walked.iter().filter(|id| **id == drawable).count() != listed
    {
        return Err(Error::Unsupported {
            reason: Unsupported::ExternalReferences,
        });
    }
    let encode = |values: &[u64]| {
        values
            .iter()
            .filter(|id| **id != drawable)
            .map(|id| wire::reference(*id))
            .collect::<Vec<_>>()
    };
    fields.replace_all(SLIDE_OWNED_DRAWABLES, LENGTH_DELIMITED, &encode(&owned));
    fields.replace_all(SLIDE_Z_ORDER, LENGTH_DELIMITED, &encode(&z_order));
    let limits = workspace.archive_limits();
    workspace
        .object_mut(slide)
        .map_err(lift)?
        .replace_message_pruning_object_references_preserving_header_with_limits(
            slide_index,
            RawMessage {
                type_: SLIDE,
                data: fields.encode(),
            },
            &[drawable],
            limits,
        )
        .map_err(StructureError::from)
        .map_err(lift)?;

    {
        let component = workspace
            .registry
            .component_mut(component_id)
            .ok_or(Error::InvalidSource)?;
        component.remove_object_uuids(&closure);
        for identifier in &closure {
            component.remove_data_owner(*identifier);
        }
    }
    for identifier in &closure {
        workspace.remove_object(*identifier).map_err(lift)?;
    }
    Ok(Removal {
        objects: closure.len(),
    })
}

/// Move one drawable to `position` in its slide's back-to-front z-order.
pub(super) fn move_to(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
    position: usize,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let slide_object = workspace.require(slide).map_err(lift)?;
    let (index, payload) = unique_message(slide_object, SLIDE).map_err(lift)?;
    let mut fields = Fields::parse(payload, wire).map_err(lift)?;
    let mut order = inventory::references(&fields, SLIDE_Z_ORDER, wire).map_err(lift)?;
    let current = order
        .iter()
        .position(|id| *id == drawable)
        .ok_or(Error::Unsupported {
            reason: Unsupported::ExternalReferences,
        })?;
    if position >= order.len() {
        return Err(Error::ZPositionOutOfRange {
            position: litchi_core::Position::new(position),
            count: order.len(),
        });
    }
    let moved = order.remove(current);
    order.insert(position, moved);
    let encoded = order
        .iter()
        .map(|id| wire::reference(*id))
        .collect::<Vec<_>>();
    fields.replace_all(SLIDE_Z_ORDER, LENGTH_DELIMITED, &encoded);
    replace_payload(workspace, slide, index, SLIDE, fields.encode()).map_err(lift)
}
