//! Root-object materialization for layout-owned images and live video.

use super::*;
use graph::LayoutMediaRoots;

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const IMAGE_FLAGS_FIELD: u32 = 7;
const TEMPLATE_MEDIA_FLAG: u32 = 1 << 1;

pub(in crate::keynote::editor) fn materialize_cloned_media(
    archive: &mut Archive,
    source_roots: &LayoutMediaRoots,
    remap: &HashMap<u64, u64>,
    slide_id: u64,
) -> Result<()> {
    for source in source_roots.identifiers() {
        let identifier = remap[&source];
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Cloned Keynote layout media object {identifier} is missing"
            ))
        })?;
        materialize_media_object(object, slide_id)?;
    }
    Ok(())
}

pub(super) fn materialize_media_object(object: &mut ArchiveObject, slide_id: u64) -> Result<()> {
    let image_count = object
        .messages
        .iter()
        .filter(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .count();
    let movie_count = object
        .messages
        .iter()
        .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .count();
    match (image_count, movie_count) {
        (1, 0) => materialize_image_object(object, slide_id),
        (0, 1) => materialize_live_video_object(object, slide_id),
        _ => Err(Error::InvalidFormat(
            "Keynote layout media root must contain exactly one image or live-video payload"
                .to_owned(),
        )),
    }
}

fn materialize_image_object(object: &mut ArchiveObject, slide_id: u64) -> Result<()> {
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote layout image payload is missing".to_owned())
        })?;
    let original = object.messages[index].data.as_slice();
    let image = tsd::ImageArchive::decode(original)?;
    validate_parent(&image.super_, slide_id, "image")?;
    let flags = image.flags.map(|flags| flags & !TEMPLATE_MEDIA_FLAG);
    let data = patch_varint_field(
        original,
        IMAGE_FLAGS_FIELD,
        image.flags.is_some(),
        flags.map(u64::from),
    )?;
    let mut expected = image;
    expected.flags = flags;
    if tsd::ImageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote layout image materialization failed validation".to_owned(),
        ));
    }
    object.replace_message(
        index,
        RawMessage {
            type_: IMAGE_MESSAGE_TYPE,
            data,
        },
    )?;
    Ok(())
}

fn materialize_live_video_object(object: &mut ArchiveObject, slide_id: u64) -> Result<()> {
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote layout live-video payload is missing".to_owned())
        })?;
    let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
    validate_parent(&movie.super_, slide_id, "live video")?;
    if movie.is_live_video != Some(true) {
        return Err(Error::InvalidFormat(
            "Keynote layout movie is not marked as live video".to_owned(),
        ));
    }
    Ok(())
}

fn validate_parent(drawable: &tsd::DrawableArchive, slide_id: u64, kind: &str) -> Result<()> {
    if drawable.parent.as_ref().map(|parent| parent.identifier) != Some(slide_id) {
        return Err(Error::InvalidFormat(format!(
            "Keynote materialized {kind} has the wrong slide parent"
        )));
    }
    Ok(())
}
