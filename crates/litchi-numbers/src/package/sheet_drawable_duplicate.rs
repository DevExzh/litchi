//! The private graph of a sheet drawable and its copy under fresh
//! identifiers, for [`crate::sheet::drawable::Edit::duplicate`] (record
//! 0378): the host editor's `clone_numbers_drawable_graph_object` with its
//! per-type `TSP.Reference` routes.
//!
//! The graph is the drawable and the objects of its title and caption
//! stand-ins, in that order; every other reference (the sheet, the style,
//! the comment, embedded data) is shared with the copy.

use std::collections::HashMap;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    WireView, patch_varint_field, transform_length_delimited_field,
    transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::shape_properties_codec as codec;

use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::structure::{self, Fault, FaultResult, Working, references};

/// `TSD.ImageArchive`, `TSD.MovieArchive`, `TSWP.StorageArchive`,
/// `TSWP.ParagraphStyleArchive`, `TSA.CaptionInfoArchive` and the
/// title/caption stand-in.
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
/// `TSD.DrawableArchive.title` and `.caption`.
const DRAWABLE_TITLE_FIELD: u32 = 10;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
/// `TSWP.ShapeInfoArchive.owned_storage`.
const OWNED_STORAGE_FIELD: u32 = 4;
/// The duplicate's offset from its source (the host's
/// `DRAWABLE_DUPLICATE_OFFSET`).
pub(super) const OFFSET: f32 = 10.0;

/// The `TSP.Reference` routes of a cloned message (the host's
/// `remap_numbers_*_wire` reference paths).
fn reference_paths(message_type: u32) -> Option<&'static [&'static [u32]]> {
    const SHAPE: &[&[u32]] = &[
        &[1, 1, 2],
        &[1, 1, 6],
        &[1, 1, 9],
        &[1, 1, 10],
        &[1, 1, 11],
        &[1, 2],
        &[2],
        &[3],
        &[4],
    ];
    const IMAGE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    const MOVIE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    const STORAGE: &[&[u32]] = &[
        &[2],
        &[5, 1, 2],
        &[7, 1, 2],
        &[8, 1, 2],
        &[9, 1, 2],
        &[11, 1, 2],
        &[12, 1, 2],
        &[15, 1, 2],
        &[16, 1, 2],
        &[17, 1, 2],
        &[18, 1, 2],
        &[21, 1, 2],
        &[22, 1, 2],
        &[23, 1, 2],
        &[27, 1, 2],
        &[28, 1, 2],
        &[25, 1, 2],
        &[26, 1, 2],
    ];
    match message_type {
        SHAPE_INFO_MESSAGE_TYPE => Some(SHAPE),
        IMAGE_MESSAGE_TYPE => Some(IMAGE),
        MOVIE_MESSAGE_TYPE => Some(MOVIE),
        2_001 | 2_022 => Some(STORAGE),
        _ => super::sheet_drawable_chart::reference_paths(message_type),
    }
}

fn remap_paths(data: &[u8], paths: &[&[u32]], remap: &HashMap<u64, u64>) -> FaultResult<Vec<u8>> {
    let mut output = data.to_vec();
    for path in paths {
        output = transform_length_delimited_fields_at_path(&output, path, |reference| {
            let view = WireView::parse(reference, WireLimits::DEFAULT)?;
            let current = view
                .fields()
                .find(|field| field.number() == 1)
                .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
                .map(|(value, _)| value);
            let Some(replacement) = current.and_then(|current| remap.get(&current)) else {
                return Ok(reference.to_vec());
            };
            patch_varint_field(reference, 1, true, Some(*replacement))
        })
        .map_err(|error: litchi_iwa_common::Error| structure::map_wire(&error))?;
    }
    Ok(output)
}

/// The drawable payload's title or caption stand-in, if it references one.
fn label(working: &Working<'_>, drawable: &[u8], field: u32) -> FaultResult<Option<u64>> {
    let records = references::records(drawable, &[field]).unwrap_or_default();
    let [record] = records.as_slice() else {
        return Ok(None);
    };
    let identifier = references::reference_identifier(record)?;
    let object = working.object(identifier)?;
    match object.messages.as_slice() {
        [message] if message.type_ == STANDIN_MESSAGE_TYPE => Ok(Some(identifier)),
        // A materialized caption graph (CaptionInfo, storage, style) is not
        // copied yet.
        _ => Err(Fault::Unsupported),
    }
}

/// The private graph of `drawable` (its message `index` of `message_type`),
/// in the host's order: an image, movie or audio clip, then its title and
/// caption stand-ins; a text box, then its caption and title stand-ins and
/// its owned storage.
pub(super) fn graph_objects(
    working: &Working<'_>,
    drawable: u64,
    index: usize,
    message_type: u32,
) -> FaultResult<Vec<u64>> {
    let payload = &working.object(drawable)?.messages[index].data;
    let inner = graph::drawable_payload(message_type, payload)?;
    let shape = message_type == SHAPE_INFO_MESSAGE_TYPE;
    let fields = if shape {
        [DRAWABLE_CAPTION_FIELD, DRAWABLE_TITLE_FIELD]
    } else {
        [DRAWABLE_TITLE_FIELD, DRAWABLE_CAPTION_FIELD]
    };
    let mut objects = vec![drawable];
    let push = |objects: &mut Vec<u64>, identifier: u64| {
        if objects.contains(&identifier) {
            return Err(Fault::Invalid);
        }
        structure::reserve(objects, 1)?;
        objects.push(identifier);
        Ok(())
    };
    for field in fields {
        if let Some(label) = label(working, inner, field)? {
            push(&mut objects, label)?;
        }
    }
    if shape {
        let records = references::records(payload, &[OWNED_STORAGE_FIELD])?;
        let [record] = records.as_slice() else {
            return Err(Fault::Invalid);
        };
        push(&mut objects, references::reference_identifier(record)?)?;
    }
    Ok(objects)
}

/// A copy of `object` under `remap` (the host's
/// `clone_numbers_drawable_graph_object` and
/// `clone_numbers_object_metadata`).
pub(super) fn clone_object(
    object: &ArchiveObject,
    remap: &HashMap<u64, u64>,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<ArchiveObject> {
    let old = object.archive_info.identifier.ok_or(Fault::Invalid)?;
    let new = *remap.get(&old).ok_or(Fault::Invalid)?;
    let mut messages = Vec::new();
    structure::reserve(&mut messages, object.messages.len())?;
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        let data = match message.type_ {
            STANDIN_MESSAGE_TYPE => structure::copy_bytes(&message.data)?,
            message_type => match reference_paths(message_type) {
                Some(paths) => remap_paths(&message.data, paths, remap)?,
                None => {
                    if info
                        .object_references
                        .iter()
                        .any(|identifier| remap.contains_key(identifier))
                    {
                        return Err(Fault::Dependency);
                    }
                    structure::copy_bytes(&message.data)?
                },
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    let mut cloned =
        ArchiveObject::new(new, messages, limits).map_err(|error| structure::map_core(&error))?;
    cloned.archive_info.should_merge = object.archive_info.should_merge;
    for ((target, original), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&object.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_error| Fault::Invalid)?;
        *target = original.clone();
        target.length = length;
        for identifier in &mut target.object_references {
            if let Some(replacement) = remap.get(identifier) {
                *identifier = *replacement;
            }
        }
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

/// `payload` (a drawable message of `message_type`) moved by `offset`
/// points on both axes.
pub(super) fn offset(payload: &[u8], message_type: u32, offset: f32) -> FaultResult<Vec<u8>> {
    let inner = graph::drawable_payload(message_type, payload)?;
    let (geometry, _) = graph::frame_and_properties(inner)?;
    let position = geometry
        .and_then(|geometry| geometry.position)
        .ok_or(Fault::Invalid)?;
    let (x, y) = (position.x + offset, position.y + offset);
    if !x.is_finite() || !y.is_finite() {
        return Err(Fault::Invalid);
    }
    let rewrite = |drawable: &[u8]| {
        codec::rewrite_drawable_archive(
            drawable,
            codec::GeometryWrite::preserve().with_position(Some(codec::Point::new(x, y))),
            codec::DrawablePropertiesWrite::preserve(),
            graph::codec_options(drawable),
        )
        .map(codec::RewriteOutput::into_bytes)
        .map_err(|_error| litchi_iwa_common::Error::InvalidFormat("drawable".to_owned()))
    };
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    if message_type == SHAPE_INFO_MESSAGE_TYPE {
        transform_length_delimited_field(payload, 1, |shape| {
            transform_length_delimited_field(shape, 1, rewrite)
        })
        .map_err(wire)
    } else {
        transform_length_delimited_field(payload, 1, rewrite).map_err(wire)
    }
}
