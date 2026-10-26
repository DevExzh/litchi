//! Duplicate, Arrange Flip and property staging of text boxes, shapes and
//! lines (record 0372; a port of the host's `duplicate_slide_shape`,
//! `duplicate_slide_text_box`, `flip_slide_shape`,
//! `set_slide_shape_properties` and `set_slide_text_box_properties`).

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_common::shape::geometry::FlipAxis;

use super::super::structure::StructureError;
use super::super::structure::wire::Fields;
use super::super::structure::workspace::Workspace;
use super::inventory::drawable_path;
use super::stage::{edit_nested, replace_payload};
use super::{Error, Geometry, lift};
use crate::slide::media::MediaProperties;

const DUPLICATE_OFFSET: f32 = 10.0;
const GEOMETRY: u32 = 1;
const GEOMETRY_FLAGS: u32 = 3;
const GEOMETRY_ANGLE: u32 = 4;
/// The `TSD.GeometryArchive.flags` bit Keynote's Arrange Flip toggles.
const REFLECTION_FLAG: u64 = 1 << 2;
const HYPERLINK_URL: u32 = 4;
const LOCKED: u32 = 5;
const ASPECT_RATIO_LOCKED: u32 = 7;
const ACCESSIBILITY_DESCRIPTION: u32 = 8;
const HALF_TURN_DEGREES: f32 = 180.0;
const FULL_TURN_DEGREES: f32 = 360.0;

fn payload(
    workspace: &Workspace<'_>,
    drawable: u64,
    message_index: usize,
) -> Result<Vec<u8>, Error> {
    Ok(workspace
        .require(drawable)
        .map_err(lift)?
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone())
}

/// Edit the drawable's `TSD.DrawableArchive` (plus `suffix`) and store it.
fn edit_drawable(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    suffix: &[u32],
    edit: &mut dyn FnMut(&mut Fields) -> Result<(), StructureError>,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let source = payload(workspace, drawable, message_index)?;
    let mut path = drawable_path(message_type).to_vec();
    path.extend_from_slice(suffix);
    let data = edit_nested(&source, &path, wire, edit).map_err(lift)?;
    replace_payload(workspace, drawable, message_index, message_type, data).map_err(lift)
}

/// The native flags and angle of a drawable's frame.
fn flags_and_angle(
    workspace: &Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
) -> Result<(Option<u64>, Option<f32>), Error> {
    let wire = workspace.wire;
    let mut fields = payload(workspace, drawable, message_index)?;
    for number in drawable_path(message_type).iter().chain([&GEOMETRY]) {
        let parsed = Fields::parse(&fields, wire).map_err(lift)?;
        fields = parsed
            .unique(*number)
            .map_err(lift)?
            .ok_or(Error::InvalidSource)?
            .message()
            .map_err(lift)?
            .to_vec();
    }
    let geometry = Fields::parse(&fields, wire).map_err(lift)?;
    let flags = geometry
        .unique(GEOMETRY_FLAGS)
        .map_err(lift)?
        .map(|field| field.varint().map_err(lift))
        .transpose()?;
    let angle = geometry
        .unique(GEOMETRY_ANGLE)
        .map_err(lift)?
        .map(|field| field.float().map_err(lift))
        .transpose()?;
    Ok((flags, angle))
}

fn vertical_flip_angle(angle: Option<f32>) -> f32 {
    let current = angle.unwrap_or_default().rem_euclid(FULL_TURN_DEGREES);
    let flipped = (current + HALF_TURN_DEGREES).rem_euclid(FULL_TURN_DEGREES);
    if flipped == 0.0 { 0.0 } else { flipped }
}

/// Apply one native Arrange Flip: the reflection flag toggles; a vertical
/// flip also adds a half turn.
pub(super) fn flip(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    axis: FlipAxis,
) -> Result<(), Error> {
    let (flags, angle) = flags_and_angle(workspace, drawable, message_index, message_type)?;
    let flags = flags.ok_or(Error::Refused(super::Refusal::UnflippableGeometry))?;
    let angle = match axis {
        FlipAxis::Horizontal => angle,
        FlipAxis::Vertical => Some(vertical_flip_angle(angle)),
    };
    edit_drawable(
        workspace,
        drawable,
        message_index,
        message_type,
        &[GEOMETRY],
        &mut |fields| {
            fields.set_varint(GEOMETRY_FLAGS, flags ^ REFLECTION_FLAG);
            if let Some(angle) = angle {
                fields.set_float(GEOMETRY_ANGLE, angle);
            }
            Ok(())
        },
    )
}

/// Whether a drawable's frame is reflected.
pub(super) fn reflected(
    workspace: &Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
) -> Result<bool, Error> {
    Ok(
        flags_and_angle(workspace, drawable, message_index, message_type)?
            .0
            .is_some_and(|flags| flags & REFLECTION_FLAG != 0),
    )
}

/// Set a drawable's hyperlink, lock and accessibility properties; `None`
/// omits a property.
pub(super) fn set_properties(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    properties: &MediaProperties,
) -> Result<(), Error> {
    edit_drawable(
        workspace,
        drawable,
        message_index,
        message_type,
        &[],
        &mut |fields| {
            let text = |fields: &mut Fields, number, value: Option<&str>| match value {
                Some(value) => fields.set_message(number, value.as_bytes()),
                None => {
                    fields.remove(number);
                },
            };
            let flag = |fields: &mut Fields, number, value: Option<bool>| match value {
                Some(value) => fields.set_varint(number, u64::from(value)),
                None => {
                    fields.remove(number);
                },
            };
            text(fields, HYPERLINK_URL, properties.hyperlink_url());
            flag(fields, LOCKED, properties.locked());
            flag(
                fields,
                ASPECT_RATIO_LOCKED,
                properties.aspect_ratio_locked(),
            );
            text(
                fields,
                ACCESSIBILITY_DESCRIPTION,
                properties.accessibility_description(),
            );
            Ok(())
        },
    )
}

/// Read a drawable's hyperlink, lock and accessibility properties.
pub(super) fn properties(
    workspace: &Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
) -> Result<MediaProperties, Error> {
    let wire = workspace.wire;
    let mut bytes = payload(workspace, drawable, message_index)?;
    for number in drawable_path(message_type) {
        bytes = Fields::parse(&bytes, wire)
            .map_err(lift)?
            .unique(*number)
            .map_err(lift)?
            .ok_or(Error::InvalidSource)?
            .message()
            .map_err(lift)?
            .to_vec();
    }
    let fields = Fields::parse(&bytes, wire).map_err(lift)?;
    let text = |number| -> Result<Option<String>, Error> {
        fields
            .unique(number)
            .map_err(lift)?
            .map(|field| {
                std::str::from_utf8(field.message().map_err(lift)?)
                    .map(str::to_owned)
                    .map_err(|_| Error::InvalidSource)
            })
            .transpose()
    };
    let flag = |number| -> Result<Option<bool>, Error> {
        fields
            .unique(number)
            .map_err(lift)?
            .map(|field| field.varint().map(|value| value != 0).map_err(lift))
            .transpose()
    };
    Ok(MediaProperties::from_parts(
        text(HYPERLINK_URL)?,
        flag(LOCKED)?,
        flag(ASPECT_RATIO_LOCKED)?,
        text(ACCESSIBILITY_DESCRIPTION)?,
    ))
}

/// The drawable's private graph inside its component, in archive order (the
/// host's `private_clone_object_ids` order).
fn private_graph(workspace: &Workspace<'_>, slide: u64, drawable: u64) -> Result<Vec<u64>, Error> {
    let member = workspace
        .component_of(drawable)
        .ok_or(Error::InvalidSource)?;
    let order = workspace.member_objects(&member);
    let internal = order.iter().copied().collect::<BTreeSet<_>>();
    let mut closure = BTreeSet::from([drawable]);
    let mut pending = vec![drawable];
    while let Some(current) = pending.pop() {
        for target in super::stage::outgoing(workspace, current).map_err(lift)? {
            if target != slide && internal.contains(&target) && closure.insert(target) {
                pending.push(target);
            }
        }
    }
    Ok(order
        .into_iter()
        .filter(|object| closure.contains(object))
        .collect())
}

/// Duplicate one drawable with its private graph on the same slide, offset
/// by ten points; returns the copy and the identifier remapping.
pub(super) fn duplicate(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    geometry: Geometry,
) -> Result<(u64, HashMap<u64, u64>), Error> {
    let member = workspace
        .component_of(drawable)
        .ok_or(Error::InvalidSource)?;
    let objects = private_graph(workspace, slide, drawable)?;
    let mut remap = HashMap::with_capacity(objects.len());
    for object in &objects {
        remap.insert(*object, workspace.allocate().map_err(lift)?);
    }
    for object in &objects {
        let original = workspace.require(*object).map_err(lift)?.clone();
        let clone = workspace
            .clone_object(&original, remap[object], &remap)
            .map_err(lift)?;
        workspace.insert(&member, clone).map_err(lift)?;
    }
    let copy = remap[&drawable];
    let offset = Geometry {
        x: geometry.x + DUPLICATE_OFFSET,
        y: geometry.y + DUPLICATE_OFFSET,
        ..geometry
    };
    let wire = workspace.wire;
    edit_drawable(
        workspace,
        copy,
        message_index,
        message_type,
        &[GEOMETRY],
        &mut |fields| super::stage::patch_geometry(fields, offset, wire),
    )?;
    super::author::attach(workspace, slide, copy).map_err(lift)?;
    let owner = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .clone();
    let mut fresh = Vec::new();
    for object in &objects {
        if let Some(uuid) = owner.uuid_of(*object) {
            fresh.push((
                remap[object],
                workspace
                    .clone_uuid(Some(uuid), remap[object])
                    .map_err(lift)?,
            ));
        }
    }
    let mut owners = Vec::new();
    for object in &objects {
        for datum in owner.data_owned_by(&[*object]) {
            owners.push((datum, remap[object]));
        }
    }
    let entry = workspace
        .registry
        .component_mut(owner.identifier)
        .ok_or(Error::InvalidSource)?;
    for (object, uuid) in fresh {
        entry.add_object_uuid(object, uuid);
    }
    for (datum, object) in owners {
        entry.add_data_owner(datum, object);
    }
    Ok((copy, remap))
}

/// `TSWP.ShapeInfoArchive.super` (`TSD.ShapeArchive`) and its path source.
const SHAPE_ARCHIVE: u32 = 1;
const SHAPE_PATH_SOURCE: u32 = 3;
/// The path-family fields of `TSD.PathSourceArchive` (point, scalar,
/// Bézier, callout, connection line, editable Bézier).
const PATH_FAMILIES: std::ops::RangeInclusive<u32> = 3..=8;

/// Replace a shape's path family with `preset`'s at the frame size, keeping
/// every other path-source field (host `set_shape_preset`).
pub(super) fn set_preset(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    preset: super::Preset,
    geometry: Geometry,
) -> Result<(), Error> {
    use litchi_iwa_common::wire::parse_wire_fields;
    let wire = workspace.wire;
    let replacement = super::author::path_source(preset, geometry.width, geometry.height)?.finish();
    let replacement_fields =
        parse_wire_fields(&replacement, wire).map_err(|_| Error::InvalidSource)?;
    let family = replacement_fields
        .iter()
        .find(|field| PATH_FAMILIES.contains(&field.number()))
        .ok_or(Error::InvalidSource)?
        .raw(&replacement)
        .map_err(|_| Error::InvalidSource)?
        .to_vec();
    let source = payload(workspace, drawable, message_index)?;
    let data = edit_nested(&source, &[SHAPE_ARCHIVE], wire, &mut |shape| {
        let original = shape
            .unique(SHAPE_PATH_SOURCE)?
            .ok_or(StructureError::Unexpected("path source"))?
            .message()?
            .to_vec();
        let fields = parse_wire_fields(&original, wire)
            .map_err(|_| StructureError::Unexpected("path source"))?;
        if fields
            .iter()
            .filter(|field| PATH_FAMILIES.contains(&field.number()))
            .count()
            > 1
        {
            return Err(StructureError::Unexpected("path families"));
        }
        let mut output = Vec::with_capacity(original.len() + family.len());
        let mut inserted = false;
        for field in &fields {
            if PATH_FAMILIES.contains(&field.number()) {
                if !inserted {
                    output.extend_from_slice(&family);
                    inserted = true;
                }
                continue;
            }
            output.extend_from_slice(
                field
                    .raw(&original)
                    .map_err(|_| StructureError::Unexpected("path source"))?,
            );
        }
        if !inserted {
            output.extend_from_slice(&family);
        }
        shape.set_message(SHAPE_PATH_SOURCE, &output);
        Ok(())
    })
    .map_err(lift)?;
    replace_payload(workspace, drawable, message_index, message_type, data).map_err(lift)
}

#[cfg(test)]
mod tests {
    use super::vertical_flip_angle;

    #[test]
    fn vertical_flip_adds_a_half_turn() {
        assert_eq!(vertical_flip_angle(None), 180.0);
        assert_eq!(vertical_flip_angle(Some(270.0)), 90.0);
    }
}
