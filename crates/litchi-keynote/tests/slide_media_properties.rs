//! Selector-first, lossless properties transactions for Keynote slide media.
//!
//! The native fixture contains two zero-size audio controls followed by two
//! file-backed movies.  These tests intentionally select by source position
//! and inspect the native graph only for physical-locality assertions; native
//! identifiers never enter the public transaction calls.

use std::{collections::BTreeMap, env, fs, io, path::PathBuf, time::Duration};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsd};
use litchi_keynote::{
    MediaPart, MediaProperties, MovieKind, MovieSelector, Package, ReadOptions, SemanticLimits,
    SlideAudioPositionError, SlideMediaPropertiesError, SlideSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_AUDIO_FOCUSED: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-properties-audio-focused-native.key");
const NATIVE_FILE_FOCUSED: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-properties-file-focused-native.key");
const NATIVE_PLACEHOLDER: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-properties-placeholder-native.key");
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const MOVIE_AUDIO_ONLY_FIELD: u32 = 9;
const MOVIE_FLAGS_FIELD: u32 = 13;
const MOVIE_LIVE_VIDEO_FIELD: u32 = 30;
const OPAQUE_AUDIO_MESSAGE_TYPE: u32 = 65_000;
const OPAQUE_AUDIO_MESSAGE: &[u8] = b"opaque Audio B extension";
const NATIVE_MOVIE_DATA_MEMBER: &str = "Data/keynote-selfauthored-coral-mjpeg-9085.mov";

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn properties(package: &Package, movie: usize) -> TestResult<MediaProperties> {
    Ok(package.slide_media_properties(SlideSelector::index(0), MovieSelector::index(movie))?)
}

fn edit_properties(
    package: &Package,
    movie: usize,
    replacement: MediaProperties,
) -> TestResult<litchi_keynote::SlideMediaPropertiesCommit> {
    Ok(package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(movie))?
        .set(replacement)?
        .commit()?)
}

fn changed_properties() -> MediaProperties {
    MediaProperties::new()
        .with_hyperlink_url(Some("https://example.test/audio-a".to_owned()))
        .with_locked(Some(true))
        .with_aspect_ratio_locked(Some(true))
        .with_accessibility_description(Some("Audio A — accessible 北区".to_owned()))
}

fn changed_file_properties() -> MediaProperties {
    MediaProperties::new()
        .with_hyperlink_url(Some("https://example.test/file-a".to_owned()))
        .with_locked(Some(true))
        .with_aspect_ratio_locked(Some(true))
        .with_accessibility_description(Some("File A — accessible 北区".to_owned()))
}

fn explicit_defaults() -> MediaProperties {
    MediaProperties::new()
        .with_hyperlink_url(Some(String::new()))
        .with_locked(Some(false))
        .with_aspect_ratio_locked(Some(false))
        .with_accessibility_description(Some(String::new()))
}

fn export_candidate(name: &str, bytes: &[u8]) -> TestResult {
    let Ok(directory) = env::var("LITCHI_KEYNOTE_MEDIA_PROPERTIES_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    fs::create_dir_all(&directory)?;
    fs::write(directory.join(name), bytes)?;
    Ok(())
}

fn catalog_entries(source: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

fn without_materialized_member(source: &[u8], member_name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    if !catalog.iter().any(|entry| entry.name() == member_name) {
        return Err(io::Error::other(format!(
            "native fixture is missing materialized member {member_name}"
        ))
        .into());
    }
    Ok(catalog.reassemble_with_deletions_to_bytes(&[], &[member_name], Limits::default())?)
}

fn native_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        archives.push((entry.name().to_owned(), Archive::parse(&stream)?));
    }
    Ok(archives)
}

fn native_object(source: &[u8], identifier: u64) -> TestResult<ArchiveObject> {
    native_archives(source)?
        .into_iter()
        .find_map(|(_, archive)| archive.object(identifier).cloned())
        .ok_or_else(|| io::Error::other("native object is missing").into())
}

fn movie_archive(archives: &[(String, Archive)], identifier: u64) -> Option<tsd::MovieArchive> {
    archives.iter().find_map(|(_, archive)| {
        archive.object(identifier).and_then(|object| {
            object.messages.iter().find_map(|message| {
                (message.type_ == MOVIE_MESSAGE_TYPE)
                    .then(|| tsd::MovieArchive::decode(message.data.as_slice()).ok())
                    .flatten()
            })
        })
    })
}

fn native_audio_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    for (_, archive) in &archives {
        let Some(slide_identifier) = archive.objects.iter().find_map(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
                .then_some(object.archive_info.identifier?)
        }) else {
            continue;
        };
        let Some(slide_object) = archive.object(slide_identifier) else {
            continue;
        };
        let Some(slide_message) = slide_object
            .messages
            .iter()
            .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        else {
            continue;
        };
        let slide = kn::SlideArchive::decode(slide_message.data.as_slice())?;
        let media = slide
            .owned_drawables
            .into_iter()
            .filter_map(|reference| {
                let movie = movie_archive(&archives, reference.identifier)?;
                let parent = movie.super_.parent.as_ref()?.identifier;
                (parent == slide_identifier).then_some((reference.identifier, movie.audio_only))
            })
            .filter_map(|(identifier, audio_only)| audio_only.eq(&Some(true)).then_some(identifier))
            .collect::<Vec<_>>();
        if media.iter().any(|identifier| {
            movie_archive(&archives, *identifier)
                .is_some_and(|movie| movie.audio_only == Some(true))
        }) {
            return Ok(media);
        }
    }
    Err(io::Error::other("native package has no slide-owned audio").into())
}

fn native_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    for (_, archive) in &archives {
        let Some(slide_identifier) = archive.objects.iter().find_map(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
                .then_some(object.archive_info.identifier?)
        }) else {
            continue;
        };
        let Some(slide_object) = archive.object(slide_identifier) else {
            continue;
        };
        let Some(slide_message) = slide_object
            .messages
            .iter()
            .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        else {
            continue;
        };
        let slide = kn::SlideArchive::decode(slide_message.data.as_slice())?;
        let media = slide
            .owned_drawables
            .into_iter()
            .filter_map(|reference| {
                let movie = movie_archive(&archives, reference.identifier)?;
                let parent = movie.super_.parent.as_ref()?.identifier;
                (parent == slide_identifier).then_some(reference.identifier)
            })
            .collect::<Vec<_>>();
        if media.iter().any(|identifier| {
            movie_archive(&archives, *identifier)
                .is_some_and(|movie| movie.audio_only == Some(true))
        }) {
            return Ok(media);
        }
    }
    Err(io::Error::other("native package has no slide-owned media").into())
}

fn duplicate_first_audio_owned_drawable(source: &[u8]) -> TestResult<Vec<u8>> {
    let target = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    for (component_name, mut archive) in native_archives(source)? {
        for object in &mut archive.objects {
            let Some(message_index) = object
                .messages
                .iter()
                .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            else {
                continue;
            };
            let mut slide =
                kn::SlideArchive::decode(object.messages[message_index].data.as_slice())?;
            let Some(duplicate) = slide
                .owned_drawables
                .iter()
                .find(|reference| reference.identifier == target)
                .cloned()
            else {
                continue;
            };
            slide.owned_drawables.push(duplicate);
            object.messages[message_index].data = slide.encode_to_vec();
            let info = object
                .archive_info
                .message_infos
                .get_mut(message_index)
                .ok_or_else(|| io::Error::other("native slide message metadata is missing"))?;
            if !info.object_references.contains(&target) {
                info.object_references.push(target);
            }
            for field in &mut info.field_infos {
                if field.path.as_slice() == [7] {
                    field.object_references.push(target);
                }
            }
            return replace_component(source, &component_name, &archive);
        }
    }
    Err(io::Error::other("native slide does not own the first audio").into())
}

fn replace_component(source: &[u8], name: &str, archive: &Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == name {
                (entry.name(), compressed.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn mutate_first_audio_payload(
    source: &[u8],
    edit: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let target = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(target).is_some())
        .ok_or_else(|| io::Error::other("first audio component is missing"))?;
    let object = archive
        .object_mut(target)
        .ok_or_else(|| io::Error::other("first audio object is missing"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("first audio movie message is missing"))?;
    message.data = edit(&message.data)?;
    replace_component(source, &component_name, &archive)
}

fn mutate_movie_header(
    source: &[u8],
    movie: usize,
    edit: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let target = *native_media_ids(source)?
        .get(movie)
        .ok_or_else(|| io::Error::other("native package has no selected movie"))?;
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(target).is_some())
        .ok_or_else(|| io::Error::other("selected movie component is missing"))?;
    let object = archive
        .object_mut(target)
        .ok_or_else(|| io::Error::other("selected movie object is missing"))?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("selected movie message is missing"))?;
    edit(object, message_index)?;
    replace_component(source, &component_name, &archive)
}

fn mutate_movie_kind(
    source: &[u8],
    movie: usize,
    target_field: u32,
    value: u64,
) -> TestResult<Vec<u8>> {
    mutate_movie_header(source, movie, |object, message_index| {
        let payload = object.messages[message_index].data.clone();
        let root = WireView::parse(&payload)?;
        let replacement_field = if root.fields().any(|field| field.number() == target_field) {
            Some(target_field)
        } else if root
            .fields()
            .any(|field| field.number() == MOVIE_FLAGS_FIELD)
        {
            Some(MOVIE_FLAGS_FIELD)
        } else if root
            .fields()
            .any(|field| field.number() == MOVIE_AUDIO_ONLY_FIELD)
        {
            Some(MOVIE_AUDIO_ONLY_FIELD)
        } else {
            None
        };
        let mut output = Vec::with_capacity(payload.len().saturating_add(2));
        let mut replaced = false;
        for field in root.fields() {
            if Some(field.number()) != replacement_field {
                output.extend_from_slice(field.raw());
                continue;
            }
            if replaced || field.wire_type() != 0 {
                return Err(io::Error::other(
                    "selected movie classification is not singular varint",
                )
                .into());
            }
            append_varint_field(&mut output, target_field, value)?;
            replaced = true;
        }
        if !replaced {
            append_varint_field(&mut output, target_field, value)?;
        }
        object.messages[message_index].data = output;
        object.archive_info.message_infos[message_index].length =
            u32::try_from(object.messages[message_index].data.len())?;
        Ok(())
    })
}

fn without_movie_properties(source: &[u8], movie: usize) -> TestResult<Vec<u8>> {
    mutate_movie_header(source, movie, |object, message_index| {
        let payload = object.messages[message_index].data.clone();
        let root = WireView::parse(&payload)?;
        let super_field = root
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("selected movie drawable envelope is missing"))?;
        let drawable = WireView::parse(super_field.payload())?;
        let mut drawable_payload = Vec::with_capacity(super_field.payload().len());
        for field in drawable.fields() {
            if matches!(field.number(), 4 | 5 | 7 | 8) {
                continue;
            }
            drawable_payload.extend_from_slice(field.raw());
        }
        let output = rewrite_unique_length_field(&payload, 1, Some(&drawable_payload))?;
        object.messages[message_index].data = output;
        object.archive_info.message_infos[message_index].length =
            u32::try_from(object.messages[message_index].data.len())?;
        Ok(())
    })
}

fn stale_movie_field_info(source: &[u8], movie: usize) -> TestResult<Vec<u8>> {
    mutate_movie_header(source, movie, |object, message_index| {
        let payload = object.messages[message_index].data.as_slice();
        let movie = tsd::MovieArchive::decode(payload)?;
        let movie_data = movie
            .movie_data
            .as_ref()
            .map(|reference| reference.identifier);
        let poster_data = movie
            .poster_image_data
            .as_ref()
            .map(|reference| reference.identifier);
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or_else(|| io::Error::other("selected movie metadata is missing"))?;
        let field_index = info.field_infos.iter().position(|field| {
            field.data_references.iter().any(|identifier| {
                Some(*identifier) == movie_data || Some(*identifier) == poster_data
            })
        });
        let field_index = if let Some(field_index) = field_index {
            field_index
        } else if let Some(field_index) = info.field_infos.iter().position(|field| {
            !field.object_references.is_empty() || !field.data_references.is_empty()
        }) {
            field_index
        } else {
            let mut field = FieldInfo::new(vec![if movie_data.is_some() { 14 } else { 15 }]);
            field.data_references.push(u64::MAX);
            info.field_infos.push(field);
            info.field_infos.len().saturating_sub(1)
        };
        let field = &mut info.field_infos[field_index];
        if let Some(identifier) = field.data_references.first_mut() {
            *identifier = u64::MAX;
        } else if let Some(identifier) = field.object_references.first_mut() {
            *identifier = u64::MAX;
        } else {
            return Err(io::Error::other("selected movie field metadata has no edge").into());
        }
        Ok(())
    })
}

fn without_movie_data_reference(source: &[u8], movie: usize, poster: bool) -> TestResult<Vec<u8>> {
    mutate_movie_header(source, movie, |object, message_index| {
        let payload = object.messages[message_index].data.as_slice();
        let movie = tsd::MovieArchive::decode(payload)?;
        let target = if poster {
            movie.poster_image_data
        } else {
            movie.movie_data
        }
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            io::Error::other(if poster {
                "selected movie has no poster data reference"
            } else {
                "selected movie has no movie data reference"
            })
        })?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or_else(|| io::Error::other("selected movie metadata is missing"))?;
        let before = info.data_references.len();
        info.data_references
            .retain(|identifier| *identifier != target);
        if info.data_references.len().saturating_add(1) != before {
            return Err(io::Error::other("selected movie data reference is not unique").into());
        }
        Ok(())
    })
}

fn add_opaque_message_to_second_audio(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_movie_header(source, 1, |object, _message_index| {
        object.push_message(RawMessage {
            type_: OPAQUE_AUDIO_MESSAGE_TYPE,
            data: OPAQUE_AUDIO_MESSAGE.to_vec(),
        })?;
        Ok(())
    })
}

fn rewrite_unique_length_field(
    payload: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len());
    let mut found = false;
    for field in view.fields() {
        if field.number() != number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if found {
            return Err(io::Error::other("duplicate length-delimited field").into());
        }
        found = true;
        if let Some(replacement) = replacement {
            append_length_delimited_field(&mut output, number, replacement)?;
        }
    }
    if !found {
        return Err(io::Error::other("missing length-delimited field").into());
    }
    Ok(output)
}

fn without_audio_geometry(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_first_audio_payload(source, |payload| {
        let root = WireView::parse(payload)?;
        let super_field = root
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("audio drawable envelope is missing"))?;
        let super_payload = rewrite_unique_length_field(super_field.payload(), 1, None)?;
        rewrite_unique_length_field(payload, 1, Some(&super_payload))
    })
}

fn duplicate_audio_geometry(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_first_audio_payload(source, |payload| {
        let root = WireView::parse(payload)?;
        let super_field = root
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("audio drawable envelope is missing"))?;
        let drawable = WireView::parse(super_field.payload())?;
        let geometry = drawable
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("audio geometry is missing"))?;
        let mut drawable_payload =
            Vec::with_capacity(super_field.payload().len() + geometry.raw().len());
        let mut duplicated = false;
        for field in drawable.fields() {
            drawable_payload.extend_from_slice(field.raw());
            if field.number() == 1 {
                drawable_payload.extend_from_slice(field.raw());
                duplicated = true;
            }
        }
        if !duplicated {
            return Err(io::Error::other("audio geometry was not duplicated").into());
        }
        rewrite_unique_length_field(payload, 1, Some(&drawable_payload))
    })
}

fn mutate_drawable_varint(
    source: &[u8],
    field_number: u32,
    value: u64,
    duplicate: bool,
) -> TestResult<Vec<u8>> {
    mutate_first_audio_payload(source, |payload| {
        let root = WireView::parse(payload)?;
        let super_field = root
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("audio drawable envelope is missing"))?;
        let drawable = WireView::parse(super_field.payload())?;
        let mut drawable_payload = Vec::with_capacity(super_field.payload().len() + 8);
        let mut found = false;
        for field in drawable.fields() {
            if field.number() != field_number {
                drawable_payload.extend_from_slice(field.raw());
                continue;
            }
            if found {
                return Err(io::Error::other("duplicate drawable property field").into());
            }
            found = true;
            if duplicate {
                drawable_payload.extend_from_slice(field.raw());
                drawable_payload.extend_from_slice(field.raw());
            } else {
                append_varint_field(&mut drawable_payload, field_number, value)?;
            }
        }
        if !found {
            append_varint_field(&mut drawable_payload, field_number, value)?;
            if duplicate {
                append_varint_field(&mut drawable_payload, field_number, value)?;
            }
        }
        rewrite_unique_length_field(payload, 1, Some(&drawable_payload))
    })
}

fn mutate_drawable_wrong_wire(source: &[u8], field_number: u32) -> TestResult<Vec<u8>> {
    mutate_first_audio_payload(source, |payload| {
        let root = WireView::parse(payload)?;
        let super_field = root
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("audio drawable envelope is missing"))?;
        let drawable = WireView::parse(super_field.payload())?;
        let mut drawable_payload = Vec::with_capacity(super_field.payload().len() + 8);
        let mut found = false;
        for field in drawable.fields() {
            if field.number() != field_number {
                drawable_payload.extend_from_slice(field.raw());
                continue;
            }
            if found {
                return Err(io::Error::other("duplicate drawable property field").into());
            }
            found = true;
            append_length_delimited_field(&mut drawable_payload, field_number, b"wrong-wire")?;
        }
        if !found {
            append_length_delimited_field(&mut drawable_payload, field_number, b"wrong-wire")?;
        }
        rewrite_unique_length_field(payload, 1, Some(&drawable_payload))
    })
}

fn archive_payloads(source: &[u8]) -> TestResult<BTreeMap<(String, u64, u32), Vec<u8>>> {
    let mut payloads = BTreeMap::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&stream)?;
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            for message in object.messages {
                payloads.insert(
                    (entry.name().to_owned(), identifier, message.type_),
                    message.data,
                );
            }
        }
    }
    Ok(payloads)
}

fn assert_media_bytes_unchanged(before: &Package, after: &Package) -> TestResult {
    for movie in 0..4 {
        assert_eq!(
            after.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie),
                MediaPart::Content
            )?,
            before.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie),
                MediaPart::Content
            )?,
            "content bytes changed for source-order media {movie}"
        );
        if movie >= 2 {
            assert_eq!(
                after.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(movie),
                    MediaPart::Poster
                )?,
                before.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(movie),
                    MediaPart::Poster
                )?,
                "poster bytes changed for source-order movie {movie}"
            );
        }
    }
    Ok(())
}

fn assert_physical_locality(before: &[u8], after: &[u8]) -> TestResult {
    let before_entries = catalog_entries(before)?;
    let after_entries = catalog_entries(after)?;
    let before_objects = archive_payloads(before)?;
    let after_objects = archive_payloads(after)?;
    assert_eq!(
        before_objects.keys().collect::<Vec<_>>(),
        after_objects.keys().collect::<Vec<_>>()
    );
    let changed = before_objects
        .iter()
        .filter_map(|(key, payload)| (after_objects.get(key) != Some(payload)).then_some(key))
        .collect::<Vec<_>>();
    assert_eq!(
        changed.len(),
        1,
        "media-properties rewrite should touch one native movie payload"
    );
    assert_eq!(changed[0].2, MOVIE_MESSAGE_TYPE);

    let changed_component = &changed[0].0;
    let before_non_preview = before_entries
        .keys()
        .filter(|name| !name.starts_with("preview"))
        .collect::<Vec<_>>();
    let after_non_preview = after_entries
        .keys()
        .filter(|name| !name.starts_with("preview"))
        .collect::<Vec<_>>();
    assert_eq!(before_non_preview, after_non_preview);
    for (name, bytes) in &before_entries {
        if name.starts_with("preview") || name == changed_component {
            continue;
        }
        assert_eq!(
            after_entries.get(name),
            Some(bytes),
            "unrelated package member {name} changed"
        );
    }

    assert!(
        after_entries
            .keys()
            .all(|name| !name.starts_with("preview")),
        "a published media-properties edit must invalidate stale previews"
    );
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
enum CommentIdentity {
    StorageUuid { lower: u64, upper: u64 },
    Identifier(u64),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentSignature {
    identity: CommentIdentity,
    text: String,
    has_author: bool,
    replies: Vec<CommentIdentity>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentGraphSignature {
    root: CommentIdentity,
    nodes: Vec<CommentSignature>,
}

fn native_comment_graph_signature(
    source: &[u8],
    movie: usize,
) -> TestResult<Option<CommentGraphSignature>> {
    const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;

    let archives = native_archives(source)?;
    let media_identifier = *native_media_ids(source)?
        .get(movie)
        .ok_or_else(|| io::Error::other("native package has no selected media"))?;
    let movie = movie_archive(&archives, media_identifier)
        .ok_or_else(|| io::Error::other("native package has no selected movie payload"))?;
    let Some(root_identifier) = movie.super_.comment.map(|reference| reference.identifier) else {
        return Ok(None);
    };

    let mut pending = vec![root_identifier];
    let mut nodes = BTreeMap::new();
    while let Some(identifier) = pending.pop() {
        if nodes.contains_key(&identifier) {
            continue;
        }
        let object = native_object(source, identifier)?;
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("native comment storage payload is missing"))?;
        let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
        let identity = comment
            .storage_uuid
            .as_ref()
            .map(|uuid| CommentIdentity::StorageUuid {
                lower: uuid.lower,
                upper: uuid.upper,
            })
            .unwrap_or(CommentIdentity::Identifier(identifier));
        let text = comment
            .text
            .ok_or_else(|| io::Error::other("native comment text is missing"))?;
        let replies = comment
            .replies
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        pending.extend(replies.iter().rev().copied());
        nodes.insert(
            identifier,
            (identity, text, comment.author.is_some(), replies),
        );
    }

    let root = nodes
        .get(&root_identifier)
        .map(|node| node.0.clone())
        .ok_or_else(|| io::Error::other("native comment root is missing"))?;
    let identities = nodes
        .iter()
        .map(|(identifier, (identity, _, _, _))| (*identifier, identity.clone()))
        .collect::<BTreeMap<_, _>>();
    let mut signatures = nodes
        .into_iter()
        .map(|(_identifier, (identity, text, has_author, replies))| {
            let replies = replies
                .into_iter()
                .map(|reply| {
                    identities
                        .get(&reply)
                        .cloned()
                        .ok_or_else(|| io::Error::other("native comment reply is missing"))
                })
                .collect::<Result<Vec<_>, _>>()?;
            Ok(CommentSignature {
                identity,
                text,
                has_author,
                replies,
            })
        })
        .collect::<TestResult<Vec<_>>>()?;
    signatures.sort_by(|left, right| left.identity.cmp(&right.identity));
    Ok(Some(CommentGraphSignature {
        root,
        nodes: signatures,
    }))
}

fn assert_media_properties_candidate(
    saved_bytes: &[u8],
    selected_movie: usize,
    expected: &MediaProperties,
    context: &str,
) -> TestResult {
    let saved = Package::from_bytes(saved_bytes)?;
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    baseline.validate()?;
    saved.validate()?;
    let baseline_movies = baseline.show()?.slides()[0].movies();
    let saved_movies = saved.show()?.slides()[0].movies();
    assert_eq!(
        saved_movies.len(),
        baseline_movies.len(),
        "{context} changed the source-order media count"
    );
    assert_eq!(
        properties(&saved, selected_movie)?,
        *expected,
        "{context} media properties differ at source-order movie {selected_movie}"
    );
    for movie in 0..baseline_movies.len() {
        if movie != selected_movie {
            assert_eq!(
                properties(&saved, movie)?,
                properties(&baseline, movie)?,
                "{context} unselected media properties changed at source-order movie {movie}"
            );
        }
        assert_eq!(
            saved_movies[movie].kind(),
            baseline_movies[movie].kind(),
            "{context} media kind changed at source-order movie {movie}"
        );
        assert_eq!(
            saved_movies[movie].position(),
            baseline_movies[movie].position(),
            "{context} media position changed at source-order movie {movie}"
        );
        assert_eq!(
            saved_movies[movie].size(),
            baseline_movies[movie].size(),
            "{context} media size changed at source-order movie {movie}"
        );
        assert_eq!(
            saved_movies[movie].original_size(),
            baseline_movies[movie].original_size(),
            "{context} original media size changed at source-order movie {movie}"
        );
        assert_eq!(
            saved_movies[movie].natural_size(),
            baseline_movies[movie].natural_size(),
            "{context} natural media size changed at source-order movie {movie}"
        );
        assert_eq!(
            saved_movies[movie].playback(),
            baseline_movies[movie].playback(),
            "{context} playback changed at source-order movie {movie}"
        );
        assert_eq!(
            native_comment_graph_signature(NATIVE_BASELINE, movie)?,
            native_comment_graph_signature(saved_bytes, movie)?,
            "{context} comment/reply graph changed at source-order movie {movie}"
        );
    }
    assert_eq!(
        native_media_ids(saved_bytes)?.len(),
        baseline_movies.len(),
        "{context} changed the slide-owned media count"
    );
    assert_eq!(
        native_audio_ids(saved_bytes)?.len(),
        native_audio_ids(NATIVE_BASELINE)?.len(),
        "{context} changed the slide-owned audio count"
    );
    assert_media_bytes_unchanged(&baseline, &saved)?;
    Ok(())
}

fn assert_saved_candidate_if_requested(
    environment_variable: &str,
    selected_movie: usize,
    expected: &MediaProperties,
) -> TestResult {
    let Ok(path) = env::var(environment_variable) else {
        return Ok(());
    };
    let saved_bytes = fs::read(path)?;
    assert_media_properties_candidate(&saved_bytes, selected_movie, expected, "native-saved")
}

#[test]
fn native_saved_audio_properties_focused_candidate_is_stable() -> TestResult {
    assert_media_properties_candidate(
        NATIVE_AUDIO_FOCUSED,
        0,
        &changed_properties(),
        "native audio-focused saved candidate",
    )?;
    Ok(())
}

#[test]
fn native_saved_file_properties_focused_candidate_is_stable() -> TestResult {
    assert_media_properties_candidate(
        NATIVE_FILE_FOCUSED,
        2,
        &changed_file_properties(),
        "native file-focused saved candidate",
    )?;
    Ok(())
}

#[test]
fn native_media_properties_read_placeholder_fixture_and_refuse_mutation() -> TestResult {
    let package = Package::from_bytes(NATIVE_PLACEHOLDER)?;
    package.validate()?;
    let movies = package.show()?.slides()[0].movies();
    assert_eq!(movies.len(), 5);
    assert_eq!(
        movies.iter().map(|movie| movie.kind()).collect::<Vec<_>>(),
        vec![
            MovieKind::Audio,
            MovieKind::Audio,
            MovieKind::File,
            MovieKind::File,
            MovieKind::Placeholder,
        ]
    );

    let placeholder = movies[4];
    assert_eq!(
        placeholder
            .position()
            .map(|position| (position.x, position.y)),
        Some((321.0, 42.0))
    );
    assert_eq!(
        placeholder.size().map(|size| (size.width, size.height)),
        Some((640.0, 360.0))
    );
    assert_eq!(placeholder.duration(), Some(Duration::from_millis(1_250)));
    assert_eq!(
        properties(&package, 4)?.accessibility_description(),
        Some("Native movie placeholder — accessible 北区")
    );

    let before = exact_bytes(&package)?;
    let error =
        package.edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(4));
    assert!(matches!(
        error,
        Err(SlideMediaPropertiesError::WrongMediaKind)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn media_properties_retain_absence_explicit_defaults_and_unicode() -> TestResult {
    let omitted = MediaProperties::new();
    assert_eq!(omitted.hyperlink_url(), None);
    assert_eq!(omitted.locked(), None);
    assert_eq!(omitted.aspect_ratio_locked(), None);
    assert_eq!(omitted.accessibility_description(), None);

    let explicit = explicit_defaults().with_accessibility_description(Some("空欄 🎵".to_owned()));
    assert_eq!(explicit.hyperlink_url(), Some(""));
    assert_eq!(explicit.locked(), Some(false));
    assert_eq!(explicit.aspect_ratio_locked(), Some(false));
    assert_eq!(explicit.accessibility_description(), Some("空欄 🎵"));
    assert_ne!(omitted, explicit);
    Ok(())
}

#[test]
fn native_media_properties_read_all_media_in_source_order_and_reject_wrong_selectors() -> TestResult
{
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    package.validate()?;
    assert_eq!(package.show()?.slides()[0].movies().len(), 4);
    assert!(package.show()?.slides()[0].movies()[0].is_audio());
    assert!(package.show()?.slides()[0].movies()[1].is_audio());
    assert!(!package.show()?.slides()[0].movies()[2].is_audio());
    assert!(!package.show()?.slides()[0].movies()[3].is_audio());
    for movie in 0..4 {
        properties(&package, movie)?;
    }
    assert!(
        package
            .slide_media_properties(SlideSelector::index(9), MovieSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(4))
            .is_err()
    );
    Ok(())
}

#[test]
fn placeholder_and_live_video_properties_are_read_only_and_preserve_sparse_fields() -> TestResult {
    let placeholder_source = without_movie_properties(
        &mutate_movie_kind(NATIVE_BASELINE, 2, MOVIE_FLAGS_FIELD, 1)?,
        2,
    )?;
    let live_video_source = without_movie_properties(
        &mutate_movie_kind(NATIVE_BASELINE, 3, MOVIE_LIVE_VIDEO_FIELD, 1)?,
        3,
    )?;

    for (source, movie, expected_kind) in [
        (&placeholder_source, 2, MovieKind::Placeholder),
        (&live_video_source, 3, MovieKind::LiveVideo),
    ] {
        let package = Package::from_bytes(source)?;
        package.validate()?;
        assert_eq!(
            package.show()?.slides()[0].movies()[movie].kind(),
            expected_kind
        );
        assert_eq!(properties(&package, movie)?, MediaProperties::new());

        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(movie));
        assert!(matches!(
            error,
            Err(SlideMediaPropertiesError::WrongMediaKind)
        ));
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn native_media_properties_read_sparse_materialized_data_without_loading_content() -> TestResult {
    let source = without_materialized_member(NATIVE_BASELINE, NATIVE_MOVIE_DATA_MEMBER)?;
    let package = Package::from_bytes(&source)?;
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let before = exact_bytes(&package)?;

    assert_eq!(properties(&package, 2)?, properties(&baseline, 2)?);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_media_properties_edit_rejects_sparse_materialized_data_atomically() -> TestResult {
    let source = without_materialized_member(NATIVE_BASELINE, NATIVE_MOVIE_DATA_MEMBER)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(2))
        .and_then(|edit| edit.set(changed_file_properties()))
        .and_then(|edit| edit.commit());

    assert!(
        matches!(
            error,
            Err(SlideMediaPropertiesError::InvalidSource)
                | Err(SlideMediaPropertiesError::UnsupportedDependency)
        ),
        "editing a media item without its materialized Data member must be rejected"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_media_properties_noop_is_byte_exact() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let before = properties(&package, 0)?;
    let commit = edit_properties(&package, 0, before.clone())?;
    assert_eq!(commit.patch().before(), &before);
    assert_eq!(commit.patch().after(), &before);
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(exact_bytes(commit.package())?, NATIVE_BASELINE);
    Ok(())
}

#[test]
fn native_audio_properties_do_not_require_positive_geometry_and_preserve_locality() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let first = package.show()?.slides()[0].movies()[0];
    assert!(
        first
            .size()
            .is_none_or(|size| size.width == 0.0 && size.height == 0.0)
    );

    let target = changed_properties();
    let commit = edit_properties(&package, 0, target.clone())?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(properties(commit.package(), 0)?, target);
    assert_eq!(properties(commit.package(), 1)?, properties(&package, 1)?);
    assert_media_bytes_unchanged(&package, commit.package())?;
    assert_physical_locality(NATIVE_BASELINE, &candidate)?;
    assert_media_properties_candidate(&candidate, 0, &target, "generated transaction")?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().deleted_previews() > 0);
    export_candidate("media-properties-baseline.key", NATIVE_BASELINE)?;
    export_candidate("media-properties-audio-changed.key", &candidate)?;
    assert_saved_candidate_if_requested(
        "LITCHI_KEYNOTE_MEDIA_PROPERTIES_NATIVE_SAVED_PATH",
        0,
        &target,
    )?;
    Ok(())
}

#[test]
fn native_unselected_audio_opaque_message_does_not_block_audio_a_edit() -> TestResult {
    let source = add_opaque_message_to_second_audio(NATIVE_BASELINE)?;
    let audio_ids = native_audio_ids(&source)?;
    let audio_a = *audio_ids
        .first()
        .ok_or_else(|| io::Error::other("native package has no AudioA"))?;
    let audio_b = *audio_ids
        .get(1)
        .ok_or_else(|| io::Error::other("native package has no AudioB"))?;
    let before_audio_b = native_object(&source, audio_b)?;
    assert!(before_audio_b.messages.iter().any(|message| {
        message.type_ == OPAQUE_AUDIO_MESSAGE_TYPE
            && message.data.as_slice() == OPAQUE_AUDIO_MESSAGE
    }));

    let package = Package::from_bytes(&source)?;
    let target = changed_properties();
    let commit = edit_properties(&package, 0, target.clone())?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(properties(commit.package(), 0)?, target);
    assert_physical_locality(&source, &candidate)?;

    let after_audio_b = native_object(&candidate, audio_b)?;
    assert_eq!(after_audio_b.archive_info, before_audio_b.archive_info);
    assert_eq!(after_audio_b.messages, before_audio_b.messages);
    assert_eq!(
        native_object(&candidate, audio_a)?.messages.len(),
        native_object(&source, audio_a)?.messages.len()
    );

    let before_objects = archive_payloads(&source)?;
    let after_objects = archive_payloads(&candidate)?;
    for (key, payload) in before_objects.iter().filter(|(key, _)| key.1 != audio_a) {
        assert_eq!(
            after_objects.get(key),
            Some(payload),
            "unselected object {key:?} changed"
        );
    }
    Ok(())
}

#[test]
fn native_audio_properties_accept_a_missing_geometry_envelope() -> TestResult {
    let source = without_audio_geometry(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let target = changed_properties();
    let commit = edit_properties(&package, 0, target.clone())?;
    assert_eq!(properties(commit.package(), 0)?, target);
    assert_eq!(properties(commit.package(), 1)?, properties(&package, 1)?);
    assert_media_bytes_unchanged(&package, commit.package())?;
    Ok(())
}

#[test]
fn native_audio_properties_reject_a_duplicate_geometry_envelope() -> TestResult {
    let source = duplicate_audio_geometry(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
        .and_then(|edit| edit.set(changed_properties()))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        error,
        Err(SlideMediaPropertiesError::InvalidSource)
            | Err(SlideMediaPropertiesError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_missing_geometry_allows_properties_but_rejects_audio_position() -> TestResult {
    let source = without_audio_geometry(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let target = changed_properties();
    let commit = edit_properties(&package, 0, target.clone())?;
    assert_eq!(properties(commit.package(), 0)?, target);
    assert!(matches!(
        package.edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0)),
        Err(SlideAudioPositionError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn native_media_properties_reject_duplicate_slide_owned_movie_references_atomically() -> TestResult
{
    let source = duplicate_first_audio_owned_drawable(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
        .and_then(|edit| edit.set(changed_properties()))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        error,
        Err(SlideMediaPropertiesError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn media_properties_reject_merge_and_base_message_metadata_atomically() -> TestResult {
    // Properties reuse the media-asset selector's metadata guard. Keep this
    // end-to-end check so a future selector refactor cannot bypass that guard.
    for movie in [0, 2] {
        for merge_object in [false, true] {
            let source = mutate_movie_header(NATIVE_BASELINE, movie, |object, index| {
                if merge_object {
                    object.archive_info.should_merge = Some(true);
                } else {
                    object.archive_info.message_infos[index].base_message_index = Some(0);
                }
                Ok(())
            })?;
            let package = Package::from_bytes(&source)?;
            let before = exact_bytes(&package)?;
            assert!(matches!(
                package
                    .slide_media_properties(SlideSelector::index(0), MovieSelector::index(movie)),
                Err(SlideMediaPropertiesError::InvalidSource)
            ));
            assert!(matches!(
                package.edit_slide_media_properties(
                    SlideSelector::index(0),
                    MovieSelector::index(movie)
                ),
                Err(SlideMediaPropertiesError::InvalidSource)
            ));
            assert_eq!(exact_bytes(&package)?, before);
        }
    }
    Ok(())
}

#[test]
fn native_media_properties_reject_stale_selected_field_info_atomically() -> TestResult {
    let source = stale_movie_field_info(NATIVE_BASELINE, 2)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(2))
        .and_then(|edit| edit.set(changed_properties()))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        error,
        Err(SlideMediaPropertiesError::InvalidSource)
            | Err(SlideMediaPropertiesError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_media_properties_reject_missing_movie_and_poster_data_refs_atomically() -> TestResult {
    let candidates = [
        (
            "missing movie data reference",
            without_movie_data_reference(NATIVE_BASELINE, 2, false)?,
        ),
        (
            "missing poster data reference",
            without_movie_data_reference(NATIVE_BASELINE, 2, true)?,
        ),
    ];
    for (label, source) in candidates {
        let package = Package::from_bytes(&source)?;
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(2))
            .and_then(|edit| edit.set(changed_properties()))
            .and_then(|edit| edit.commit());
        assert!(
            matches!(
                error,
                Err(SlideMediaPropertiesError::InvalidSource)
                    | Err(SlideMediaPropertiesError::UnsupportedDependency)
            ),
            "{label} must be rejected"
        );
        assert_eq!(exact_bytes(&package)?, before, "{label} mutated its source");
    }
    Ok(())
}

#[test]
fn native_audio_properties_reject_duplicate_wrong_wire_and_noncanonical_fields_atomically()
-> TestResult {
    let candidates = [
        (
            "duplicate locked field",
            mutate_drawable_varint(NATIVE_BASELINE, 5, 1, true)?,
        ),
        (
            "wrong locked wire type",
            mutate_drawable_wrong_wire(NATIVE_BASELINE, 5)?,
        ),
        (
            "noncanonical locked value",
            mutate_drawable_varint(NATIVE_BASELINE, 5, 2, false)?,
        ),
    ];
    for (label, source) in candidates {
        let package = Package::from_bytes(&source)?;
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .and_then(|edit| edit.set(changed_properties()))
            .and_then(|edit| edit.commit());
        assert!(error.is_err(), "{label} must be rejected");
        assert_eq!(exact_bytes(&package)?, before, "{label} mutated its source");
    }
    Ok(())
}

#[test]
fn native_media_properties_tight_reference_budget_preempts_publication_atomically() -> TestResult {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let options = ReadOptions::new(Limits::default(), semantic);
    let package = Package::from_bytes_with_options(NATIVE_BASELINE, options)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
        .and_then(|edit| edit.set(changed_properties()))
        .and_then(|edit| edit.commit());
    assert!(
        error.is_err(),
        "a one-reference profile must reject the media-properties graph"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_media_properties_allow_both_lock_states_and_explicit_empty_values() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let explicit = explicit_defaults().with_accessibility_description(Some("".to_owned()));
    let false_commit = edit_properties(&package, 0, explicit.clone())?;
    assert_eq!(properties(false_commit.package(), 0)?, explicit);

    let unicode = explicit_defaults()
        .with_hyperlink_url(Some("https://例.example/音声".to_owned()))
        .with_locked(Some(true))
        .with_aspect_ratio_locked(Some(true))
        .with_accessibility_description(Some("説明 🎵".to_owned()));
    let true_commit = edit_properties(false_commit.package(), 0, unicode.clone())?;
    assert_eq!(properties(true_commit.package(), 0)?, unicode);
    assert_eq!(
        properties(true_commit.package(), 1)?,
        properties(&package, 1)?
    );
    Ok(())
}

#[test]
fn native_media_properties_accept_long_unicode_accessibility_text_without_losing_other_fields()
-> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let long_description = "北".repeat(8_192);
    let target = MediaProperties::new()
        .with_hyperlink_url(Some("https://example.test/audio-a".to_owned()))
        .with_locked(Some(true))
        .with_aspect_ratio_locked(Some(false))
        .with_accessibility_description(Some(long_description));
    let commit = edit_properties(&package, 0, target.clone())?;
    assert_eq!(properties(commit.package(), 0)?, target);
    assert_eq!(properties(commit.package(), 1)?, properties(&package, 1)?);
    assert_eq!(properties(commit.package(), 2)?, properties(&package, 2)?);
    assert_eq!(properties(commit.package(), 3)?, properties(&package, 3)?);
    assert_media_bytes_unchanged(&package, commit.package())?;
    Ok(())
}

#[test]
fn native_media_properties_inverse_replay_double_inverse_and_conflict_are_exact() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let target = changed_properties();
    let changed = edit_properties(&package, 0, target)?;
    let changed_bytes = exact_bytes(changed.package())?;
    assert_eq!(&changed.patch().inverse().inverse(), changed.patch());

    let replay = package.apply_slide_media_properties(changed.patch())?;
    assert_eq!(exact_bytes(replay.package())?, changed_bytes);
    let restored = changed
        .package()
        .apply_slide_media_properties(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE_BASELINE);
    export_candidate(
        "media-properties-inverse.key",
        &exact_bytes(restored.package())?,
    )?;

    let foreign = edit_properties(&package, 1, MediaProperties::new().with_locked(Some(true)))?;
    let foreign_bytes = exact_bytes(foreign.package())?;
    assert!(
        foreign
            .package()
            .apply_slide_media_properties(changed.patch())
            .is_err()
    );
    assert_eq!(exact_bytes(foreign.package())?, foreign_bytes);
    Ok(())
}

#[test]
fn native_file_media_properties_are_selector_typed_and_do_not_move_audio() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let before_audio = properties(&package, 0)?;
    let file_target = changed_file_properties();
    let commit = edit_properties(&package, 2, file_target.clone())?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(properties(commit.package(), 2)?, file_target);
    assert_eq!(properties(commit.package(), 0)?, before_audio);
    export_candidate("media-properties-file-changed.key", &candidate)?;
    assert_media_properties_candidate(&candidate, 2, &file_target, "generated transaction")?;
    assert_saved_candidate_if_requested(
        "LITCHI_KEYNOTE_MEDIA_PROPERTIES_NATIVE_SAVED_FILE_PATH",
        2,
        &file_target,
    )?;
    Ok(())
}
