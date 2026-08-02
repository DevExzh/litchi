//! Wire-preserving movie drawable geometry edits.

use super::*;
use crate::shapes::{
    DrawableProperties, drawable_properties, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};

const MOVIE_DRAWABLE_FIELD: u32 = 1;

pub(in crate::keynote::editor) fn set_movie_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let message_index = *message_index;
        let data = transform_length_delimited_field(
            object.messages[message_index].data.as_slice(),
            MOVIE_DRAWABLE_FIELD,
            |drawable| patch_drawable_geometry(drawable, geometry),
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(in crate::keynote::editor) fn set_movie_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    properties: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::MovieArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::MovieArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Keynote movie properties patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            *message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn offset_movie(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    offset: f32,
) -> Result<()> {
    let geometry = {
        let archive = package.archive(archive_name)?;
        let object = archive.object(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote movie object {movie_id} is missing"))
        })?;
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote movie {movie_id} payload is missing"))
            })?;
        let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
        geometry_from_drawable(&movie.super_)?
    };
    let position = geometry
        .position
        .ok_or_else(|| Error::InvalidFormat(format!("Keynote movie {movie_id} has no position")))?;
    set_movie_geometry(
        package,
        archive_name,
        movie_id,
        DrawableGeometry {
            position: Some(crate::shapes::DrawablePoint {
                x: position.x + offset,
                y: position.y + offset,
            }),
            ..geometry
        },
    )
}
