//! Small synthetic Keynote package used by the slide-media fuzz target.
//!
//! Keeping the source in the fuzz crate makes the target independent of a
//! native `.key` fixture.  The package contains two file movies sharing one
//! content and one poster record, plus one audio drawable with its own
//! content record.  Its complete source is only a few kilobytes after ZIP
//! framing, so every fuzz iteration exercises the semantic transaction at a
//! bounded cost.

use litchi_iwa_archive::{Limits, package::to_bytes};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use prost::Message as _;
use sha1::{Digest as _, Sha1};
use std::sync::OnceLock;

pub const MOVIE_BYTES: &[u8] = b"\0\0\0\x18ftypavc1tiny-fuzz-movie-v1";
pub const POSTER_BYTES: &[u8] = b"\x89PNG\r\n\x1a\ntiny-fuzz-poster-v1";
pub const AUDIO_BYTES: &[u8] = b"RIFFxxxxWAVEtiny-fuzz-audio-v1";

pub const REPLACED_MOVIE_BYTES: &[u8] = b"\0\0\0\x18ftypavc1tiny-fuzz-movie-v2";
pub const REPLACED_POSTER_BYTES: &[u8] = b"\x89PNG\r\n\x1a\ntiny-fuzz-poster-v2";
pub const REPLACED_AUDIO_BYTES: &[u8] = b"RIFFxxxxWAVEtiny-fuzz-audio-v2";

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const DOCUMENT_COMPONENT: u64 = 1;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const MOVIES: [u64; 2] = [100, 101];
const AUDIO: u64 = 102;
const POSTER_DATA: u64 = 2_001;
const CONTENT_DATA: u64 = 2_002;
const AUDIO_DATA: u64 = 2_003;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const UNKNOWN_FIELD: u32 = 4_091;
const UNKNOWN_MARKER: &[u8] = b"fuzz-media-unknown-extension";

type SeedResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> SeedResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    object_references: Vec<u64>,
    data_references: Vec<u64>,
) -> SeedResult<ArchiveObject> {
    let mut value = object(identifier, type_, data)?;
    let info = &mut value.archive_info.message_infos[0];
    info.object_references = object_references;
    info.data_references = data_references;
    Ok(value)
}

fn movie_payload(movie: usize) -> SeedResult<Vec<u8>> {
    let value = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point {
                    x: 80.0 + movie as f32 * 420.0,
                    y: 120.0,
                }),
                size: Some(tsp::Size {
                    width: 320.0,
                    height: 180.0,
                }),
                flags: Some(0x20),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference {
            identifier: CONTENT_DATA,
        }),
        poster_image_data: Some(tsp::DataReference {
            identifier: POSTER_DATA,
        }),
        original_size: Some(tsp::Size {
            width: 1_920.0,
            height: 1_080.0,
        }),
        natural_size: Some(tsp::Size {
            width: 1_920.0,
            height: 1_080.0,
        }),
        flags: Some(0),
        start_time: Some(0.0),
        end_time: Some(2.0),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = value;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn audio_payload() -> SeedResult<Vec<u8>> {
    let value = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                size: Some(tsp::Size {
                    width: 0.0,
                    height: 0.0,
                }),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference {
            identifier: AUDIO_DATA,
        }),
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
        start_time: Some(0.0),
        end_time: Some(2.0),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = value;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn metadata_uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn metadata_payload() -> SeedResult<Vec<u8>> {
    let data = |identifier: u64, name: &str, bytes: &[u8]| tsp::DataInfo {
        identifier,
        digest: Sha1::digest(bytes).to_vec(),
        preferred_file_name: name.to_owned(),
        file_name: Some(name.to_owned()),
        materialized_length: Some(bytes.len() as u64),
        ..Default::default()
    };
    let owner = |data_identifier: u64, identifiers: &[u64]| tsp::ComponentDataReference {
        data_identifier,
        object_reference_list: identifiers
            .iter()
            .copied()
            .map(
                |object_identifier| tsp::component_data_reference::ObjectReference {
                    object_identifier,
                    count: 1,
                },
            )
            .collect(),
    };
    let object_identifiers = [
        1, 2, SLIDE_NODE, SLIDE, 80, 81, 90, MOVIES[0], MOVIES[1], AUDIO,
    ];
    let component = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(1),
        object_uuid_map_entries: object_identifiers
            .into_iter()
            .map(metadata_uuid_entry)
            .collect(),
        data_references: vec![
            owner(POSTER_DATA, &MOVIES),
            owner(CONTENT_DATA, &MOVIES),
            owner(AUDIO_DATA, &[AUDIO]),
        ],
        ..Default::default()
    };
    let metadata = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        components: vec![component],
        datas: vec![
            data(POSTER_DATA, "poster.png", POSTER_BYTES),
            data(CONTENT_DATA, "movie.mov", MOVIE_BYTES),
            data(AUDIO_DATA, "audio.m4a", AUDIO_BYTES),
        ],
        save_token: Some(1),
        ..Default::default()
    };
    let mut payload = metadata.encode_to_vec();
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> SeedResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn build() -> SeedResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_920.0,
            height: 1_080.0,
        },
        stylesheet: reference(81),
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
        name: Some("Fuzz media".to_owned()),
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
            MOVIES.into_iter().chain([AUDIO]).collect(),
            Vec::new(),
        )?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for (index, movie) in MOVIES.into_iter().enumerate() {
        objects.push(object_with_references(
            movie,
            MOVIE_MESSAGE_TYPE,
            movie_payload(index)?,
            Vec::new(),
            vec![CONTENT_DATA, POSTER_DATA],
        )?);
    }
    objects.push(object_with_references(
        AUDIO,
        MOVIE_MESSAGE_TYPE,
        audio_payload()?,
        Vec::new(),
        vec![AUDIO_DATA],
    )?);
    let document_component = component(objects)?;
    let metadata_component =
        component(vec![object(METADATA_OBJECT, 11_006, metadata_payload()?)?])?;
    Ok(to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated fuzz sentinel".as_slice()),
            ("Data/movie.mov", MOVIE_BYTES),
            ("Data/poster.png", POSTER_BYTES),
            ("Data/audio.m4a", AUDIO_BYTES),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

pub fn bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            build()
                .unwrap_or_else(|error| panic!("tiny Keynote media fuzz seed must build: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}
