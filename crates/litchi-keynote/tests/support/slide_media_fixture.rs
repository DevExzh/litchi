//! Synthetic Keynote package with shared movie/audio media records.
//!
//! The fixture deliberately contains the smallest complete media closure that
//! a package-level content replacement owner should accept: two file movies
//! share both a movie payload and a poster payload, while a separate audio
//! drawable owns a third payload.  PackageMetadata carries matching DataInfo
//! and component ownership records, and every modeled message includes an
//! opaque extension so tests can prove source-preserving edits.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use litchi_keynote::Package;
use prost::Message as _;
use sha1::{Digest, Sha1};

pub(super) const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
pub(super) const METADATA_MEMBER: &str = "Index/Metadata.iwa";
pub(super) const METADATA_OBJECT: u64 = 300;
pub(super) const DOCUMENT_COMPONENT: u64 = 1;
pub(super) const UNRELATED_COMPONENT: u64 = 2;
pub(super) const METADATA_LAST_IDENTIFIER: u64 = 1_000;
pub(super) const SLIDE_NODE: u64 = 3;
pub(super) const SLIDE: u64 = 4;
pub(super) const MOVIES: [u64; 2] = [100, 101];
pub(super) const AUDIO: u64 = 102;
pub(super) const SHARED_IMAGE: u64 = 150;
pub(super) const TITLES: [u64; 2] = [110, 111];
pub(super) const AUDIO_TITLE: u64 = 112;
pub(super) const CAPTIONS: [u64; 2] = [130, 131];
pub(super) const AUDIO_CAPTION: u64 = 132;
pub(super) const STYLES: [u64; 2] = [160, 161];
pub(super) const AUDIO_STYLE: u64 = 162;

/// One shared poster record is referenced by both file movies.
pub(super) const POSTER_DATA: u64 = 2_001;
/// One shared movie record is referenced by both file movies.
pub(super) const CONTENT_DATA: u64 = 2_002;
/// The audio drawable owns this independent content record.
pub(super) const AUDIO_DATA: u64 = 2_003;

pub(super) const MOVIE_MESSAGE_TYPE: u32 = 3_007;
pub(super) const IMAGE_MESSAGE_TYPE: u32 = 3_005;
pub(super) const STANDIN_MESSAGE_TYPE: u32 = 3_097;
pub(super) const UNKNOWN_FIELD: u32 = 4_091;
pub(super) const UNKNOWN_MARKER: &[u8] = b"slide-media-unknown-extension";
pub(super) const METADATA_UNKNOWN_MARKER: &[u8] = b"slide-media-metadata-extension";

pub(super) const MOVIE_BYTES: &[u8] = b"\0\0\0\x18ftypavc1synthetic movie source bytes v1";
pub(super) const POSTER_BYTES: &[u8] = b"\x89PNG\r\n\x1a\nsynthetic poster source bytes v1";
pub(super) const AUDIO_BYTES: &[u8] = b"RIFFxxxxWAVEsynthetic audio source bytes v1";
pub(super) const REPLACED_MOVIE_BYTES: &[u8] =
    b"\0\0\0\x18ftypavc1replacement movie source bytes v2";
pub(super) const REPLACED_POSTER_BYTES: &[u8] =
    b"\x89PNG\r\n\x1a\nreplacement poster source bytes v2";
pub(super) const REPLACED_AUDIO_BYTES: &[u8] = b"RIFFxxxxWAVEreplacement audio source bytes v2";

pub(super) type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

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
    object_references: Vec<u64>,
    data_references: Vec<u64>,
) -> TestResult<ArchiveObject> {
    let mut value = object(identifier, type_, data)?;
    let info = &mut value.archive_info.message_infos[0];
    info.object_references = object_references;
    info.data_references = data_references;
    Ok(value)
}

fn movie_payload(movie: usize) -> TestResult<Vec<u8>> {
    let value = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                size: Some(tsp::Size {
                    width: 800.0,
                    height: 300.0,
                }),
                flags: Some(0x20),
                angle: Some(17.5),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            title: Some(reference(TITLES[movie])),
            caption: Some(reference(CAPTIONS[movie])),
            accessibility_description: Some("Slide media locality movie".to_owned()),
            locked: Some(false),
            aspect_ratio_locked: Some(true),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference {
            identifier: CONTENT_DATA,
        }),
        poster_image_data: Some(tsp::DataReference {
            identifier: POSTER_DATA,
        }),
        style: Some(reference(STYLES[movie])),
        original_size: Some(tsp::Size {
            width: 1_920.0,
            height: 1_080.0,
        }),
        natural_size: Some(tsp::Size {
            width: 1_920.0,
            height: 1_080.0,
        }),
        flags: Some(0),
        start_time: Some(0.25),
        end_time: Some(8.5),
        poster_time: Some(1.5),
        loop_option: Some(2),
        volume: Some(0.75),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = value;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn audio_payload() -> TestResult<Vec<u8>> {
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
            accessibility_description: Some("Slide media locality audio".to_owned()),
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
    }
    .encode_to_vec();
    let mut payload = value;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn image_payload() -> TestResult<Vec<u8>> {
    let value = tsd::ImageArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 40.0, y: 60.0 }),
                size: Some(tsp::Size {
                    width: 120.0,
                    height: 80.0,
                }),
                ..Default::default()
            }),
            parent: Some(reference(SLIDE)),
            accessibility_description: Some("Shared poster image owner".to_owned()),
            ..Default::default()
        },
        data: Some(tsp::DataReference {
            identifier: POSTER_DATA,
        }),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = value;
    append_length_delimited_field(
        &mut payload,
        UNKNOWN_FIELD,
        b"shared-poster-image-unknown-extension",
    )?;
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

fn metadata_component_payload(
    identifier: u64,
    locator: &str,
    token: u64,
    object_identifiers: &[u64],
    data_references: Vec<tsp::ComponentDataReference>,
) -> TestResult<Vec<u8>> {
    let component = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: object_identifiers
            .iter()
            .copied()
            .map(metadata_uuid_entry)
            .collect(),
        data_references,
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = component;
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, METADATA_UNKNOWN_MARKER)?;
    Ok(payload)
}

fn metadata_payload() -> TestResult<Vec<u8>> {
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
        1,
        2,
        SLIDE_NODE,
        SLIDE,
        THEME,
        STYLESHEET,
        90,
        MOVIES[0],
        MOVIES[1],
        AUDIO,
        TITLES[0],
        TITLES[1],
        AUDIO_TITLE,
        CAPTIONS[0],
        CAPTIONS[1],
        AUDIO_CAPTION,
        STYLES[0],
        STYLES[1],
        AUDIO_STYLE,
    ];
    let document = metadata_component_payload(
        DOCUMENT_COMPONENT,
        "Document",
        10,
        &object_identifiers,
        vec![
            owner(POSTER_DATA, &MOVIES),
            owner(CONTENT_DATA, &MOVIES),
            owner(AUDIO_DATA, &[AUDIO]),
        ],
    )?;
    let unrelated =
        metadata_component_payload(UNRELATED_COMPONENT, "Unrelated", 7, &[900], Vec::new())?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(901)],
        ..Default::default()
    }
    .encode_to_vec();
    let metadata = tsp::PackageMetadata {
        last_object_identifier: METADATA_LAST_IDENTIFIER,
        components: vec![tsp::ComponentInfo::decode(document.as_slice())?],
        datas: vec![
            data(POSTER_DATA, "poster.png", POSTER_BYTES),
            data(CONTENT_DATA, "movie.mov", MOVIE_BYTES),
            data(AUDIO_DATA, "audio.m4a", AUDIO_BYTES),
        ],
        save_token: Some(10),
        versioned_components: vec![tsp::ComponentInfo::decode(versioned.as_slice())?],
        ..Default::default()
    };
    // Encode the typed metadata first, then append an opaque root extension.
    // This keeps the fixture's component/data shape easy to inspect while
    // retaining a byte that changed writers must preserve.
    let mut payload = metadata.encode_to_vec();
    append_length_delimited_field(&mut payload, UNKNOWN_FIELD, METADATA_UNKNOWN_MARKER)?;
    // Keep the unrelated component in the same root payload after the typed
    // projection so selected-component ownership is tested against a second
    // locator and UUID namespace.
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

/// Build the exact source used by the focused media replacement tests.
pub(super) fn synthetic_package() -> TestResult<Vec<u8>> {
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
            width: 1_920.0,
            height: 1_080.0,
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
        name: Some("Media replacement".to_owned()),
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
            Vec::new(),
        )?,
        object(THEME, 10, Vec::new())?,
        object(STYLESHEET, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for movie in 0..MOVIES.len() {
        objects.push(object_with_references(
            MOVIES[movie],
            MOVIE_MESSAGE_TYPE,
            movie_payload(movie)?,
            vec![TITLES[movie], CAPTIONS[movie], STYLES[movie]],
            vec![CONTENT_DATA, POSTER_DATA],
        )?);
        objects.push(object(TITLES[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(CAPTIONS[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(STYLES[movie], 2_025, Vec::new())?);
    }
    objects.push(object_with_references(
        AUDIO,
        MOVIE_MESSAGE_TYPE,
        audio_payload()?,
        vec![AUDIO_TITLE, AUDIO_CAPTION, AUDIO_STYLE],
        vec![AUDIO_DATA],
    )?);
    objects.push(object(AUDIO_TITLE, STANDIN_MESSAGE_TYPE, Vec::new())?);
    objects.push(object(AUDIO_CAPTION, STANDIN_MESSAGE_TYPE, Vec::new())?);
    objects.push(object(AUDIO_STYLE, 2_025, Vec::new())?);
    let metadata = component(vec![object(METADATA_OBJECT, 11_006, metadata_payload()?)?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            ("Data/movie.mov", MOVIE_BYTES),
            ("Data/poster.png", POSTER_BYTES),
            ("Data/audio.m4a", AUDIO_BYTES),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT_MEMBER, component(objects)?.as_slice()),
            (METADATA_MEMBER, metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}

/// Add a legitimate non-Movie image drawable that shares the file movies'
/// poster DataInfo record.  The base fixture stays minimal; only this scoped
/// variant exercises the replacement owner's complete current owner census.
pub(super) fn synthetic_package_with_shared_image_poster_owner() -> TestResult<Vec<u8>> {
    let source = synthetic_package()?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let slide = archive
        .object_mut(SLIDE)
        .ok_or_else(|| io::Error::other("missing synthetic slide object"))?;
    let message = slide
        .messages
        .iter_mut()
        .find(|message| message.type_ == 5)
        .ok_or_else(|| io::Error::other("missing synthetic slide message"))?;
    let mut slide_archive = kn::SlideArchive::decode(message.data.as_slice())?;
    slide_archive.owned_drawables.push(reference(SHARED_IMAGE));
    slide_archive
        .drawables_z_order
        .push(reference(SHARED_IMAGE));
    message.data = slide_archive.encode_to_vec();
    let info = slide
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing synthetic slide metadata"))?;
    info.object_references.extend([SHARED_IMAGE, SHARED_IMAGE]);
    archive.objects.push(object_with_references(
        SHARED_IMAGE,
        IMAGE_MESSAGE_TYPE,
        image_payload()?,
        Vec::new(),
        vec![POSTER_DATA],
    )?);
    let source = replace_document_archive(&source, archive)?;

    let mut metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;
    let component = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing synthetic document metadata"))?;
    component
        .object_uuid_map_entries
        .push(metadata_uuid_entry(SHARED_IMAGE));
    let poster = component
        .data_references
        .iter_mut()
        .find(|reference| reference.data_identifier == POSTER_DATA)
        .ok_or_else(|| io::Error::other("missing synthetic poster ownership"))?;
    poster
        .object_reference_list
        .push(tsp::component_data_reference::ObjectReference {
            object_identifier: SHARED_IMAGE,
            count: 1,
        });
    replace_metadata_payload(&source, metadata.encode_to_vec())
}

pub(super) fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

pub(super) fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| io::Error::other(format!("missing package member {name}")))
        .map_err(Into::into)
}

pub(super) fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::decompress(&member_bytes(package, DOCUMENT_MEMBER)?)?.into_bytes())
}

pub(super) fn metadata_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(
        &SnappyStream::decompress(&member_bytes(package, METADATA_MEMBER)?)?.into_bytes(),
    )?;
    archive
        .object(METADATA_OBJECT)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 11_006)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing metadata payload").into())
}

pub(super) fn movie_payload_from_package(package: &[u8], movie: u64) -> TestResult<Vec<u8>> {
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
        .ok_or_else(|| io::Error::other(format!("missing movie payload {movie}")).into())
}

pub(super) fn image_payload_from_package(package: &[u8], image: u64) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .object(image)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == IMAGE_MESSAGE_TYPE)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other(format!("missing image payload {image}")).into())
}

pub(super) fn replace_document_archive(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
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

pub(super) fn replace_metadata_payload(source: &[u8], payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let replacement = SnappyStream::compress(
        &Archive {
            objects: vec![object(METADATA_OBJECT, 11_006, payload)?],
        }
        .to_bytes()?,
    )?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == METADATA_MEMBER {
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

pub(super) fn with_movie_payload(
    source: &[u8],
    movie: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(movie)
        .ok_or_else(|| io::Error::other(format!("missing movie object {movie}")))?
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing movie message"))?
        .data = payload;
    replace_document_archive(source, archive)
}

const THEME: u64 = 80;
const STYLESHEET: u64 = 81;
