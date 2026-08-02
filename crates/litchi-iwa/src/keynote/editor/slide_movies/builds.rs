//! Movie build and timing-chunk graph cloning.

use super::*;

pub(super) fn clone_movie_build(
    source: &ArchiveObject,
    new_identifier: u64,
    source_drawable_id: u64,
    new_drawable_id: u64,
) -> Result<ArchiveObject> {
    let indexes = source
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == BUILD_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(
            "Keynote movie build must have exactly one BuildArchive payload".to_owned(),
        ));
    };
    let message_index = *message_index;
    let original = source.messages[message_index].data.as_slice();
    let mut expected = kn::BuildArchive::decode(original)?;
    if expected
        .drawable
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(source_drawable_id)
    {
        return Err(Error::InvalidFormat(
            "Keynote movie build targets the wrong source drawable".to_owned(),
        ));
    }
    expected.drawable = Some(tsp::Reference {
        identifier: new_drawable_id,
        ..Default::default()
    });
    let data = patch_length_delimited_field(
        original,
        1,
        true,
        Some(
            &tsp::Reference {
                identifier: new_drawable_id,
                ..Default::default()
            }
            .encode_to_vec(),
        ),
    )?;
    if kn::BuildArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote movie-build remap failed validation".to_owned(),
        ));
    }
    let mut messages = source.messages.clone();
    messages[message_index] = RawMessage {
        type_: BUILD_MESSAGE_TYPE,
        data,
    };
    clone_object_metadata(
        source,
        new_identifier,
        messages,
        &HashMap::from([(source_drawable_id, new_drawable_id)]),
        false,
    )
}

pub(super) fn clone_movie_build_chunk(
    source: &ArchiveObject,
    new_identifier: u64,
    source_build_id: u64,
    new_build_id: u64,
    new_build_uuid: tsp::Uuid,
) -> Result<ArchiveObject> {
    let indexes = source
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == BUILD_CHUNK_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(
            "Keynote movie build chunk must have exactly one BuildChunkArchive payload".to_owned(),
        ));
    };
    let message_index = *message_index;
    let original = source.messages[message_index].data.as_slice();
    let mut expected = kn::BuildChunkArchive::decode(original)?;
    if expected
        .build
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(source_build_id)
        || expected.build_id.is_none()
        || expected
            .build_chunk_identifier
            .as_ref()
            .and_then(|identifier| identifier.build_id)
            .is_none()
    {
        return Err(Error::InvalidFormat(
            "Keynote movie build chunk has an incomplete build identity".to_owned(),
        ));
    }
    expected.build = Some(tsp::Reference {
        identifier: new_build_id,
        ..Default::default()
    });
    expected.build_id = Some(new_build_uuid);
    if let Some(identifier) = &mut expected.build_chunk_identifier {
        identifier.build_id = Some(new_build_uuid);
    }

    let data = patch_length_delimited_field(
        original,
        1,
        true,
        Some(
            &tsp::Reference {
                identifier: new_build_id,
                ..Default::default()
            }
            .encode_to_vec(),
        ),
    )?;
    let data = transform_length_delimited_field(&data, 7, |identifier| {
        patch_length_delimited_field(identifier, 1, true, Some(&new_build_uuid.encode_to_vec()))
    })?;
    let data = patch_length_delimited_field(&data, 8, true, Some(&new_build_uuid.encode_to_vec()))?;
    if kn::BuildChunkArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Keynote movie build-chunk remap failed validation".to_owned(),
        ));
    }
    let mut messages = source.messages.clone();
    messages[message_index] = RawMessage {
        type_: BUILD_CHUNK_MESSAGE_TYPE,
        data,
    };
    clone_object_metadata(
        source,
        new_identifier,
        messages,
        &HashMap::from([(source_build_id, new_build_id)]),
        false,
    )
}

pub(super) fn append_slide_build_references(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    build_ids: &[u64],
    chunk_ids: &[u64],
) -> Result<()> {
    if build_ids.is_empty() && chunk_ids.is_empty() {
        return Ok(());
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        };
        let message_index = *message_index;
        let mut data = object.messages[message_index].data.clone();
        for (field, identifiers) in [
            (SLIDE_BUILDS_FIELD, build_ids),
            (SLIDE_BUILD_CHUNKS_FIELD, chunk_ids),
        ] {
            for identifier in identifiers {
                if repeated_length_delimited_payloads(&data, field)?
                    .into_iter()
                    .filter_map(|payload| tsp::Reference::decode(payload).ok())
                    .any(|reference| reference.identifier == *identifier)
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} already references build object {identifier}"
                    )));
                }
                data = append_repeated_length_delimited_field(
                    &data,
                    field,
                    &tsp::Reference {
                        identifier: *identifier,
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?;
            }
        }
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        if !build_ids.iter().all(|identifier| {
            verified
                .builds
                .iter()
                .any(|reference| reference.identifier == *identifier)
        }) || !chunk_ids.iter().all(|identifier| {
            verified
                .build_chunks
                .iter()
                .any(|reference| reference.identifier == *identifier)
        }) {
            return Err(Error::InvalidFormat(
                "Keynote movie-build insertion failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[message_index];
        for identifier in build_ids.iter().chain(chunk_ids) {
            if !info.object_references.contains(identifier) {
                info.object_references.push(*identifier);
            }
        }
        Ok(())
    })
}
