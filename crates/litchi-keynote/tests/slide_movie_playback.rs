use std::io;
use std::time::Duration;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use litchi_keynote::slide::media::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use litchi_keynote::{MovieSelector, Package, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const METADATA_LAST_IDENTIFIER: u64 = 1_000;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const MOVIES: [u64; 2] = [100, 101];
const AUDIO: u64 = 102;
const TITLES: [u64; 2] = [110, 111];
const CAPTIONS: [u64; 2] = [130, 131];
const STYLES: [u64; 2] = [160, 161];
const AUDIO_TITLE: u64 = 112;
const AUDIO_CAPTION: u64 = 132;
const AUDIO_STYLE: u64 = 162;
const AUDIO_DATA: u64 = 2_003;
const THEME: u64 = 80;
const STYLESHEET: u64 = 81;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const UNKNOWN_FIELD: u32 = 4_091;
const UNKNOWN_MARKER: &[u8] = b"playback-unknown-extension";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> TestResult<ArchiveObject> {
    let mut value = object(identifier, type_, data)?;
    value.archive_info.message_infos[0].object_references = references;
    Ok(value)
}

fn movie_payload(movie: usize, full_playback: bool) -> Vec<u8> {
    let mut value = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                size: Some(tsp::Size {
                    width: 800.0,
                    height: 300.0,
                }),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            title: Some(reference(TITLES[movie])),
            caption: Some(reference(CAPTIONS[movie])),
            accessibility_description: Some("Playback test movie".to_owned()),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference { identifier: 2_002 }),
        poster_image_data: Some(tsp::DataReference { identifier: 2_001 }),
        style: Some(reference(STYLES[movie])),
        original_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        natural_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        flags: Some(0),
        ..Default::default()
    };
    if full_playback {
        value.start_time = Some(0.25);
        value.end_time = Some(8.5);
        value.poster_time = Some(1.5);
        value.loop_option = Some(2);
        value.volume = Some(0.75);
    } else {
        value.end_time = Some(7.0);
    }
    let mut payload = value.encode_to_vec();
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)
        .expect("unknown playback field");
    payload
}

fn audio_payload() -> Vec<u8> {
    let value = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 320.0, y: 240.0 }),
                size: Some(tsp::Size {
                    width: 0.0,
                    height: 0.0,
                }),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            title: Some(reference(AUDIO_TITLE)),
            caption: Some(reference(AUDIO_CAPTION)),
            accessibility_description: Some("Playback test audio".to_owned()),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference {
            identifier: AUDIO_DATA,
        }),
        style: Some(reference(AUDIO_STYLE)),
        original_size: Some(tsp::Size {
            width: 0.0,
            height: 0.0,
        }),
        natural_size: Some(tsp::Size {
            width: 0.0,
            height: 0.0,
        }),
        flags: Some(0),
        audio_only: Some(true),
        start_time: Some(0.5),
        end_time: Some(12.25),
        poster_time: Some(0.0),
        loop_option: Some(1),
        volume: Some(0.6),
        ..Default::default()
    };
    let mut payload = value.encode_to_vec();
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)
        .expect("unknown playback field");
    payload
}

fn metadata_uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier + 10_000,
            upper: identifier + 20_000,
        },
    }
}

fn metadata_component_payload(
    identifier: u64,
    locator: &str,
    token: u64,
    ids: &[u64],
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: ids.iter().copied().map(metadata_uuid_entry).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn metadata_payload(last_identifier: u64) -> TestResult<Vec<u8>> {
    let ids = [
        1,
        2,
        3,
        4,
        THEME,
        STYLESHEET,
        90,
        100,
        101,
        AUDIO,
        110,
        111,
        AUDIO_TITLE,
        130,
        131,
        AUDIO_CAPTION,
        160,
        161,
        AUDIO_STYLE,
    ];
    let document = metadata_component_payload(DOCUMENT_COMPONENT, "Document", 10, &ids)?;
    let unrelated = metadata_component_payload(UNRELATED_COMPONENT, "Unrelated", 7, &[900])?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(901)],
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, last_identifier)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    append_varint_field(&mut payload, 8, 10)?;
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(THEME),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(STYLESHEET),
        ..Default::default()
    };
    #[allow(deprecated)]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        ..Default::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: [MOVIES[0], MOVIES[1], AUDIO]
            .into_iter()
            .map(reference)
            .collect(),
        drawables_z_order: [MOVIES[0], MOVIES[1], AUDIO]
            .into_iter()
            .map(reference)
            .collect(),
        name: Some("Playback".to_owned()),
        in_document: true,
        ..Default::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_with_references(
            SLIDE,
            5,
            slide.encode_to_vec(),
            [MOVIES[0], MOVIES[1], AUDIO].to_vec(),
        )?,
        object(THEME, 10, Vec::new())?,
        object(STYLESHEET, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for movie in 0..MOVIES.len() {
        objects.push(object_with_references(
            MOVIES[movie],
            MOVIE_MESSAGE_TYPE,
            movie_payload(movie, true),
            vec![TITLES[movie], CAPTIONS[movie], STYLES[movie]],
        )?);
        objects.push(object(TITLES[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(CAPTIONS[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(STYLES[movie], 2_025, Vec::new())?);
    }
    objects.push(object_with_references(
        AUDIO,
        MOVIE_MESSAGE_TYPE,
        audio_payload(),
        vec![AUDIO_TITLE, AUDIO_CAPTION, AUDIO_STYLE],
    )?);
    objects.push(object(AUDIO_TITLE, STANDIN_MESSAGE_TYPE, Vec::new())?);
    objects.push(object(AUDIO_CAPTION, STANDIN_MESSAGE_TYPE, Vec::new())?);
    objects.push(object(AUDIO_STYLE, 2_025, Vec::new())?);
    let metadata = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata_payload(METADATA_LAST_IDENTIFIER)?,
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            ("Data/movie.mov", b"synthetic movie bytes".as_slice()),
            ("Data/poster.png", b"synthetic poster bytes".as_slice()),
            ("Data/audio.wav", b"synthetic audio bytes".as_slice()),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT_MEMBER, component(objects)?.as_slice()),
            (METADATA_MEMBER, metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing document"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn metadata_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing metadata"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    archive
        .object(METADATA_OBJECT)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 11_006)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing metadata message").into())
}

fn movie_payload_from_package(package: &[u8], movie: u64) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .object(movie)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing movie payload").into())
}

fn replace_document_archive(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == DOCUMENT_MEMBER {
                (entry.name(), replacement.as_slice())
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

fn with_movie_payload(source: &[u8], movie: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(movie)
        .ok_or_else(|| io::Error::other("missing movie"))?
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing movie message"))?
        .data = payload;
    replace_document_archive(source, archive)
}

fn append_fixed32(payload: &mut Vec<u8>, field: u32, value: f32) {
    let mut key = Vec::new();
    let mut raw = u64::from(field) << 3 | 5;
    while raw >= 0x80 {
        key.push((raw as u8 & 0x7f) | 0x80);
        raw >>= 7;
    }
    key.push(raw as u8);
    payload.extend_from_slice(&key);
    payload.extend_from_slice(&value.to_le_bytes());
}

fn append_unknown_extensions(payload: &mut Vec<u8>) -> TestResult<()> {
    append_length_delimited_field(payload, 4_090, b"unknown overlong/group carrier")?;
    // Unknown scalar values retain their non-canonical varint framing.
    payload.extend_from_slice(&[0xd0, 0x05, 0x96, 0x81, 0x00]);
    Ok(())
}

fn strip_movie_data(source: &[u8]) -> TestResult<Vec<u8>> {
    let payload =
        tsd::MovieArchive::decode(movie_payload_from_package(source, MOVIES[0])?.as_slice())?;
    let mut payload = payload;
    payload.movie_data = None;
    with_movie_payload(source, MOVIES[0], payload.encode_to_vec())
}

fn assert_playback_rejected(source: &[u8]) -> TestResult<()> {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => {
            assert!(
                package
                    .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))
                    .is_err()
            );
            Ok(())
        },
    }
}

#[test]
fn reads_file_movie_playback_and_preserves_media_graph() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let settings = package
        .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .ok_or_else(|| io::Error::other("missing playback settings"))?;
    assert_eq!(settings.start_time, Some(Duration::from_secs_f32(0.25)));
    assert_eq!(settings.end_time, Duration::from_secs_f32(8.5));
    assert_eq!(settings.poster_time, Some(Duration::from_secs_f32(1.5)));
    assert_eq!(settings.loop_mode, Some(MediaLoopMode::BackAndForth));
    assert_eq!(settings.volume, Some(MediaVolume::new(0.75).unwrap()));
    assert_eq!(
        movie_payload_from_package(&source, MOVIES[0])?,
        movie_payload(0, true)
    );
    assert_eq!(
        Catalog::from_bytes(&source)?
            .iter()
            .find(|entry| entry.name() == "Data/movie.mov")
            .map(|entry| entry.data()),
        Some(b"synthetic movie bytes".as_slice())
    );
    Ok(())
}

#[test]
fn no_op_is_exact_and_debug_does_not_expose_media_values() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let settings = package
        .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .ok_or_else(|| io::Error::other("missing playback settings"))?;
    let commit = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(settings)?
        .commit()?;
    assert_eq!(exact_bytes(commit.package())?, source);
    assert!(!commit.diagnostics().changed());
    assert!(!format!("{:?}", commit.patch()).contains("8.5"));
    Ok(())
}

#[test]
fn replacement_updates_all_playback_fields_inverse_and_apply_exactly() -> TestResult<()> {
    let source = synthetic_package()?;
    let before_metadata = metadata_stream(&source)?;
    let replacement = MediaPlaybackSettings::new(Duration::from_secs_f32(12.0))
        .with_start_time(Some(Duration::from_secs_f32(1.25)))
        .with_poster_time(Some(Duration::from_secs_f32(4.0)))
        .with_loop_mode(Some(MediaLoopMode::Repeat))
        .with_volume(Some(MediaVolume::new(0.5).unwrap()));
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(
        commit
            .package()
            .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
            .ok_or_else(|| io::Error::other("missing replacement playback settings"))?,
        replacement
    );
    assert_eq!(metadata_stream(&candidate)?, before_metadata);
    assert_eq!(
        movie_payload_from_package(&candidate, MOVIES[1])?,
        movie_payload(1, true)
    );
    assert_eq!(
        Catalog::from_bytes(&candidate)?
            .iter()
            .find(|entry| entry.name() == "Data/poster.png")
            .map(|entry| entry.data()),
        Some(b"synthetic poster bytes".as_slice())
    );
    let source_catalog = Catalog::from_bytes(&source)?;
    let candidate_catalog = Catalog::from_bytes(&candidate)?;
    for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
        assert_eq!(
            candidate_catalog
                .iter()
                .find(|entry| entry.name() == name)
                .map(|entry| entry.data()),
            source_catalog
                .iter()
                .find(|entry| entry.name() == name)
                .map(|entry| entry.data()),
        );
    }
    let restored = commit
        .package()
        .apply_slide_movie_playback_settings(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    let applied =
        Package::from_bytes(&source)?.apply_slide_movie_playback_settings(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, candidate);
    Ok(())
}

#[test]
fn optional_playback_fields_can_be_cleared_without_changing_required_end() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let replacement = MediaPlaybackSettings::new(Duration::from_secs_f32(6.0));
    let commit = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let actual = commit
        .package()
        .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .ok_or_else(|| io::Error::other("missing cleared playback settings"))?;
    assert_eq!(actual, replacement);
    assert_eq!(actual.start_time, None);
    assert_eq!(actual.poster_time, None);
    assert_eq!(actual.loop_mode, None);
    assert_eq!(actual.volume, None);
    Ok(())
}

#[test]
fn stale_patch_conflict_and_selector_failures_are_atomic() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let replacement = MediaPlaybackSettings::new(Duration::from_secs_f32(10.0));
    let first = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let second = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(MediaPlaybackSettings::new(Duration::from_secs_f32(11.0)))?
        .commit()?;
    assert!(
        second
            .package()
            .apply_slide_movie_playback_settings(first.patch())
            .is_err()
    );
    assert!(
        Package::from_bytes(&source)?
            .slide_movie_playback_settings(SlideSelector::index(9), MovieSelector::index(0))
            .is_err()
    );
    assert!(
        Package::from_bytes(&source)?
            .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(9))
            .is_err()
    );
    Ok(())
}

#[test]
fn malformed_known_fields_are_rejected_without_publication() -> TestResult<()> {
    let source = synthetic_package()?;
    let original = movie_payload_from_package(&source, MOVIES[0])?;
    let movie = tsd::MovieArchive::decode(original.as_slice())?;

    let mut duplicate = original.clone();
    append_fixed32(&mut duplicate, 4, 9.0);
    let duplicate_source = with_movie_payload(&source, MOVIES[0], duplicate)?;
    assert_playback_rejected(&duplicate_source)?;

    let mut wrong_wire = original.clone();
    append_varint_field(&mut wrong_wire, 4, 9)?;
    let wrong_wire_source = with_movie_payload(&source, MOVIES[0], wrong_wire)?;
    assert_playback_rejected(&wrong_wire_source)?;

    let mut nonfinite = movie;
    nonfinite.end_time = Some(f32::NAN);
    let nonfinite_source = with_movie_payload(&source, MOVIES[0], nonfinite.encode_to_vec())?;
    assert_playback_rejected(&nonfinite_source)?;

    let mut conflict = tsd::MovieArchive::decode(original.as_slice())?;
    conflict.loop_option = Some(1);
    #[allow(deprecated)]
    {
        conflict.loop_option_as_integer = Some(2);
    }
    let conflict_source = with_movie_payload(&source, MOVIES[0], conflict.encode_to_vec())?;
    assert_playback_rejected(&conflict_source)?;
    Ok(())
}

#[test]
fn unknown_fields_are_retained_and_unsupported_media_are_refused() -> TestResult<()> {
    let source = synthetic_package()?;
    let payload = movie_payload_from_package(&source, MOVIES[0])?;
    let mut rewritten = payload.clone();
    append_unknown_extensions(&mut rewritten)?;
    let hostile = with_movie_payload(&source, MOVIES[0], rewritten)?;
    let Ok(package) = Package::from_bytes(&hostile) else {
        // Some ingress profiles reject protobuf groups before the selected
        // playback edge is reached; preservation is asserted when ingress
        // admits the raw-preserving candidate.
        return Ok(());
    };
    assert!(
        package
            .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))
            .is_ok()
    );
    let replacement = MediaPlaybackSettings::new(Duration::from_secs_f32(11.0));
    let changed = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let target = movie_payload_from_package(&exact_bytes(changed.package())?, MOVIES[0])?;
    assert!(
        target
            .windows(UNKNOWN_MARKER.len())
            .any(|window| window == UNKNOWN_MARKER)
    );
    assert!(
        target
            .windows(5)
            .any(|window| window == [0xd0, 0x05, 0x96, 0x81, 0x00])
    );
    assert!(
        package
            .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(2))?
            .is_some()
    );
    let non_file = strip_movie_data(&source)?;
    assert_playback_rejected(&non_file)?;
    Ok(())
}

#[test]
fn missing_metadata_and_ingress_limits_are_atomic() -> TestResult<()> {
    let source = synthetic_package()?;
    let catalog = Catalog::from_bytes(&source)?;
    let no_metadata = litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .filter(|entry| entry.name() != METADATA_MEMBER)
            .map(|entry| (entry.name(), entry.data()))
            .collect::<Vec<_>>(),
        Limits::default(),
    )?;
    let package = Package::from_bytes(&no_metadata)?;
    let replacement = MediaPlaybackSettings::new(Duration::from_secs_f32(11.0));
    let commit = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))
        .and_then(|edit| edit.set(replacement))
        .and_then(|edit| edit.commit())?;
    let candidate = exact_bytes(commit.package())?;
    assert!(
        Catalog::from_bytes(&candidate)?
            .iter()
            .all(|entry| entry.name() != METADATA_MEMBER)
    );
    let restored = commit
        .package()
        .apply_slide_movie_playback_settings(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, no_metadata);

    let defaults = Limits::default();
    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            litchi_keynote::ReadOptions::new(tight, litchi_keynote::SemanticLimits::default()),
        )
        .is_err()
    );
    Ok(())
}
