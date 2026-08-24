use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tswp};
use litchi_keynote::{MovieSelector, Package, SlideMovieCaptionError, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const METADATA_LAST_IDENTIFIER: u64 = 1_000;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const MOVIES: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const CAPTIONS: [u64; 2] = [130, 131];
const STORAGES: [u64; 2] = [140, 141];
const PLACEMENTS: [u64; 2] = [150, 151];
const STYLES: [u64; 2] = [160, 161];
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const METADATA_ROOT_UNKNOWN_FIELD: u32 = 4_001;
const METADATA_COMPONENT_UNKNOWN_FIELD: u32 = 4_002;
const METADATA_UNKNOWN_MARKER: &[u8] = b"movie-caption metadata extension";
const STORAGE_UNKNOWN_FIELD: u32 = 4_003;
const STORAGE_UNKNOWN_MARKER: &[u8] = b"movie-caption storage extension";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
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
    let mut object = object(identifier, type_, data)?;
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

fn movie_payload(movie: usize, caption_identifier: Option<u64>) -> Vec<u8> {
    tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                size: Some(tsp::Size {
                    width: 800.0,
                    height: 300.0,
                }),
                ..tsd::GeometryArchive::default()
            }),
            parent: Some(reference(SLIDE)),
            title: Some(reference(TITLES[movie])),
            caption: caption_identifier.map(reference),
            accessibility_description: Some("Test movie".to_owned()),
            ..tsd::DrawableArchive::default()
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
        ..tsd::MovieArchive::default()
    }
    .encode_to_vec()
}

fn caption_info_payload(movie: usize) -> Vec<u8> {
    #[allow(
        deprecated,
        reason = "the native caption graph retains the legacy storage edge"
    )]
    let info = tsa::CaptionInfoArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    parent: Some(reference(MOVIES[movie])),
                    caption_hidden: Some(false),
                    ..tsd::DrawableArchive::default()
                },
                style: Some(reference(STYLES[movie])),
                ..tsd::ShapeArchive::default()
            },
            deprecated_storage: Some(reference(STORAGES[movie])),
            owned_storage: Some(reference(STORAGES[movie])),
            is_text_box: Some(true),
            ..tswp::ShapeInfoArchive::default()
        },
        placement: Some(reference(PLACEMENTS[movie])),
        child_info_kind: Some(1),
    };
    info.encode_to_vec()
}

fn storage_payload(text: &str) -> Vec<u8> {
    let mut payload = tswp::StorageArchive {
        kind: Some(3),
        text: vec![text.to_owned()],
        in_document: Some(true),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, STORAGE_UNKNOWN_FIELD, STORAGE_UNKNOWN_MARKER)
        .expect("storage fixture unknown field must encode");
    payload
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
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

fn metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    object_identifiers: &[u64],
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: object_identifiers
            .iter()
            .copied()
            .map(metadata_uuid_entry)
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(
        &mut payload,
        METADATA_COMPONENT_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let ids = [
        1, 2, 3, 4, 80, 81, 90, 100, 101, 110, 111, 130, 131, 140, 141, 150, 151, 160, 161,
    ];
    let document = metadata_component(DOCUMENT_COMPONENT, "Document", 10, &ids)?;
    let unrelated = metadata_component(UNRELATED_COMPONENT, "Unrelated", 7, &[900])?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(901)],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, METADATA_LAST_IDENTIFIER)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    append_varint_field(&mut payload, 8, 10)?;
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(
        &mut payload,
        METADATA_ROOT_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_captions([Some("North"), None])
}

fn synthetic_package_with_captions(captions: [Option<&str>; 2]) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: MOVIES.iter().copied().map(reference).collect(),
        drawables_z_order: MOVIES.iter().copied().map(reference).collect(),
        name: Some("Movies".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_with_references(SLIDE, 5, slide.encode_to_vec(), MOVIES.to_vec())?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for movie in 0..MOVIES.len() {
        let caption = captions[movie].map(|_| CAPTIONS[movie]);
        let refs = [
            TITLES[movie],
            caption.unwrap_or(CAPTIONS[movie]),
            STYLES[movie],
        ];
        objects.push(object_with_references(
            MOVIES[movie],
            MOVIE_MESSAGE_TYPE,
            movie_payload(movie, caption),
            refs.to_vec(),
        )?);
        objects.push(object(TITLES[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        if let Some(text) = captions[movie] {
            objects.push(object_with_references(
                CAPTIONS[movie],
                CAPTION_INFO_MESSAGE_TYPE,
                caption_info_payload(movie),
                vec![STYLES[movie], STORAGES[movie], PLACEMENTS[movie]],
            )?);
            objects.push(object(
                STORAGES[movie],
                STORAGE_MESSAGE_TYPE,
                storage_payload(text),
            )?);
            objects.push(object(
                PLACEMENTS[movie],
                CAPTION_PLACEMENT_MESSAGE_TYPE,
                tsa::CaptionPlacementArchive::default().encode_to_vec(),
            )?);
            objects.push(object(STYLES[movie], SHAPE_STYLE_MESSAGE_TYPE, Vec::new())?);
        } else {
            objects.push(object(CAPTIONS[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        }
    }
    let document_component = component(objects)?;
    let metadata = component(vec![object(METADATA_OBJECT, 11_006, metadata_payload()?)?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            ("Data/movie.mov", b"synthetic movie bytes".as_slice()),
            ("Data/poster.png", b"synthetic poster bytes".as_slice()),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
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

fn metadata_message_payload(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing metadata member"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(METADATA_OBJECT))
        .and_then(|object| object.messages.first())
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing metadata object"))
        .map_err(Into::into)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn replace_document_stream(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
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

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing requested movie message").into())
}

fn with_document_message_payload(
    source: &[u8],
    identifier: u64,
    type_: u32,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other("missing requested document object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing requested document message"))?;
    message.data = payload;
    replace_document_stream(source, archive)
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn varint_field(number: u32, value: u64, key_width: usize, value_width: usize) -> Vec<u8> {
    let mut output = Vec::new();
    let mut key = Vec::new();
    push_varint(u64::from(number) << 3, &mut key);
    if key_width > key.len() {
        let last = key.len() - 1;
        key[last] |= 0x80;
        key.extend(std::iter::repeat_n(0x80, key_width - key.len() - 1));
        key.push(0);
    }
    output.extend_from_slice(&key);
    let mut value_bytes = Vec::new();
    push_varint(value, &mut value_bytes);
    if value_width > value_bytes.len() {
        let last = value_bytes.len() - 1;
        value_bytes[last] |= 0x80;
        value_bytes.extend(std::iter::repeat_n(
            0x80,
            value_width - value_bytes.len() - 1,
        ));
        value_bytes.push(0);
    }
    output.extend_from_slice(&value_bytes);
    output
}

fn append_duplicate_field(payload: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let duplicate = WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("missing field to duplicate"))?;
    let mut output = payload.to_vec();
    output.extend_from_slice(&duplicate);
    Ok(output)
}

fn append_nested_duplicate(payload: &[u8], path: &[u32]) -> TestResult<Vec<u8>> {
    let number = *path
        .first()
        .ok_or_else(|| io::Error::other("empty nested path"))?;
    let view = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len());
    let mut found = false;
    for field in view.fields() {
        output.extend_from_slice(field.raw());
        if !found && field.number() == number {
            if path.len() == 1 {
                output.extend_from_slice(field.raw());
            } else if field.wire_type() == 2 {
                let nested = append_nested_duplicate(field.payload(), &path[1..])?;
                let mut replacement = Vec::new();
                append_length_delimited_field(&mut replacement, number, &nested)?;
                output.truncate(output.len() - field.raw().len());
                output.extend_from_slice(&replacement);
                found = true;
            } else {
                return Err(io::Error::other("nested field is not length-delimited").into());
            }
        }
    }
    if found || path.len() == 1 && output.len() > payload.len() {
        Ok(output)
    } else {
        Err(io::Error::other("nested field not found").into())
    }
}

fn replace_nested_field_raw(
    payload: &[u8],
    path: &[u32],
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let number = *path
        .first()
        .ok_or_else(|| io::Error::other("empty nested path"))?;
    let view = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len() + replacement.len());
    let mut replaced = false;
    for field in view.fields() {
        if !replaced && field.number() == number {
            if path.len() == 1 {
                output.extend_from_slice(replacement);
            } else {
                if field.wire_type() != 2 {
                    return Err(io::Error::other("nested field is not length-delimited").into());
                }
                let nested = replace_nested_field_raw(field.payload(), &path[1..], replacement)?;
                append_length_delimited_field(&mut output, number, &nested)?;
            }
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if replaced {
        Ok(output)
    } else {
        Err(io::Error::other("nested field not found").into())
    }
}

fn metadata_component_raw_payload(
    payload: &[u8],
    identifier: u64,
    versioned: bool,
) -> TestResult<Vec<u8>> {
    let field_number = if versioned { 11 } else { 3 };
    for field in WireView::parse(payload)?.fields() {
        if field.number() != field_number {
            continue;
        }
        let matches = WireView::parse(field.payload())?.fields().any(|nested| {
            nested.number() == 1
                && decode_varint_from_bytes(nested.payload())
                    .ok()
                    .map(|(value, _)| value)
                    == Some(identifier)
        });
        if matches {
            return Ok(field.payload().to_vec());
        }
    }
    Err(io::Error::other("missing metadata component").into())
}

fn decoded_metadata(package: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        metadata_message_payload(package)?.as_slice(),
    )?)
}

#[test]
fn reads_active_and_standin_movie_captions() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    assert_eq!(
        package.slide_movie_caption(SlideSelector::index(0), MovieSelector::index(0))?,
        Some("North".to_owned())
    );
    assert_eq!(
        package.slide_movie_caption(SlideSelector::index(0), MovieSelector::index(1))?,
        None
    );
    Ok(())
}

#[test]
fn no_op_and_graph_create_clear_refuse_atomically() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let no_op = package
        .edit_slide_movie_caption(0usize, 0usize)?
        .set("North")?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);

    let create = package
        .edit_slide_movie_caption(0usize, 1usize)?
        .set("new graph")?
        .commit();
    assert!(matches!(
        create,
        Err(SlideMovieCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let clear = package
        .edit_slide_movie_caption(0usize, 0usize)?
        .clear()?
        .commit();
    assert!(matches!(
        clear,
        Err(SlideMovieCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let catalog = Catalog::from_bytes(&source)?;
    let no_metadata_source = litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .filter(|entry| entry.name() != METADATA_MEMBER)
            .map(|entry| (entry.name(), entry.data()))
            .collect::<Vec<_>>(),
        Limits::default(),
    )?;
    let no_metadata_package = Package::from_bytes(&no_metadata_source)?;
    let changed_without_metadata = no_metadata_package
        .edit_slide_movie_caption(0usize, 0usize)?
        .set("changed without metadata")?
        .commit();
    assert!(matches!(
        changed_without_metadata,
        Err(SlideMovieCaptionError::UnsupportedSource)
            | Err(SlideMovieCaptionError::InvalidSource)
            | Err(SlideMovieCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&no_metadata_package)?, no_metadata_source);
    Ok(())
}

#[test]
fn active_replacement_updates_selected_caption_and_inverse_exactly() -> TestResult<()> {
    let source = synthetic_package_with_captions([Some("North"), Some("South")])?;
    let package = Package::from_bytes(&source)?;
    let before_metadata = metadata_message_payload(&source)?;
    let before_movie = message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?;
    let before_title_edge = raw_fields(&before_movie, 1)?;
    let before_field_one = raw_fields(&before_metadata, 1)?;
    let before_root_unknown = raw_fields(&before_metadata, METADATA_ROOT_UNKNOWN_FIELD)?;
    let before_selected_component =
        metadata_component_raw_payload(&before_metadata, DOCUMENT_COMPONENT, false)?;
    let before_selected_unknown =
        raw_fields(&before_selected_component, METADATA_COMPONENT_UNKNOWN_FIELD)?;
    let before_unrelated =
        metadata_component_raw_payload(&before_metadata, UNRELATED_COMPONENT, false)?;
    let before_versioned =
        metadata_component_raw_payload(&before_metadata, DOCUMENT_COMPONENT, true)?;
    let commit = package
        .edit_slide_movie_caption(SlideSelector::index(0), MovieSelector::index(0))?
        .set("East")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_movie_caption(0usize, MovieSelector::index(0))?,
        Some("East".to_owned())
    );
    assert_eq!(
        commit
            .package()
            .slide_movie_caption(0usize, MovieSelector::index(1))?,
        Some("South".to_owned())
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert_eq!(
        decoded_metadata(&exact_bytes(commit.package())?)?.save_token,
        Some(11)
    );
    let target = exact_bytes(commit.package())?;
    let target_metadata = metadata_message_payload(&target)?;
    let target_movie = message_payload(&target, MOVIES[0], MOVIE_MESSAGE_TYPE)?;
    assert_eq!(raw_fields(&target_movie, 1)?, before_title_edge);
    let before_storage = message_payload(&source, STORAGES[0], STORAGE_MESSAGE_TYPE)?;
    let target_storage = message_payload(&target, STORAGES[0], STORAGE_MESSAGE_TYPE)?;
    assert_eq!(
        raw_fields(&target_storage, STORAGE_UNKNOWN_FIELD)?,
        raw_fields(&before_storage, STORAGE_UNKNOWN_FIELD)?
    );
    assert_eq!(raw_fields(&target_metadata, 1)?, before_field_one);
    assert_eq!(
        raw_fields(&target_metadata, METADATA_ROOT_UNKNOWN_FIELD)?,
        before_root_unknown
    );
    assert_eq!(
        raw_fields(
            &metadata_component_raw_payload(&target_metadata, DOCUMENT_COMPONENT, false)?,
            METADATA_COMPONENT_UNKNOWN_FIELD,
        )?,
        before_selected_unknown
    );
    assert_eq!(
        metadata_component_raw_payload(&target_metadata, UNRELATED_COMPONENT, false)?,
        before_unrelated
    );
    assert_eq!(
        metadata_component_raw_payload(&target_metadata, DOCUMENT_COMPONENT, true)?,
        before_versioned
    );
    let restored = commit
        .package()
        .apply_slide_movie_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selector_errors_are_typed_and_source_stays_unchanged() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_movie_caption(SlideSelector::index(9), MovieSelector::index(0)),
        Err(SlideMovieCaptionError::SlidePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_movie_caption(SlideSelector::index(0), MovieSelector::index(9)),
        Err(SlideMovieCaptionError::MoviePositionNotFound { .. })
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn malformed_movie_edge_and_caption_storage_are_refused_atomically() -> TestResult<()> {
    let source = synthetic_package()?;
    let malformed_movie = with_document_message_payload(
        &source,
        MOVIES[0],
        MOVIE_MESSAGE_TYPE,
        vec![0x0a, 0x01, 0x5a],
    )?;
    match Package::from_bytes(&malformed_movie) {
        Err(_) => {},
        Ok(package) => {
            assert!(matches!(
                package.slide_movie_caption(0usize, 0usize),
                Err(SlideMovieCaptionError::InvalidSource)
            ));
            assert_eq!(exact_bytes(&package)?, malformed_movie);
        },
    }

    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(STORAGES[0]))
        .ok_or_else(|| io::Error::other("missing caption storage"))?
        .messages[0]
        .data = vec![0x18, 0x01];
    let malformed_storage = replace_document_stream(&source, archive)?;
    let package = Package::from_bytes(&malformed_storage)?;
    assert!(matches!(
        package.slide_movie_caption(0usize, 0usize),
        Err(SlideMovieCaptionError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, malformed_storage);
    Ok(())
}

#[test]
fn shared_caption_storage_is_rejected_without_mutation() -> TestResult<()> {
    let source = synthetic_package_with_captions([Some("North"), Some("South")])?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let second = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(CAPTIONS[1]))
        .ok_or_else(|| io::Error::other("missing second caption info"))?;
    second.messages[0].data = caption_info_payload(0);
    second.archive_info.message_infos[0].object_references =
        vec![STYLES[1], STORAGES[0], PLACEMENTS[1]];
    let hostile = replace_document_stream(&source, archive)?;
    let package = Package::from_bytes(&hostile)?;
    assert!(matches!(
        package.slide_movie_caption(0usize, 0usize),
        Err(SlideMovieCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, hostile);
    Ok(())
}

#[test]
fn patch_conflict_is_detected_without_exposing_caption_text() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_movie_caption(0usize, 0usize)?
        .set("private movie caption")?
        .commit()?;
    let other = Package::from_bytes(&synthetic_package_with_captions([Some("Other"), None])?)?;
    assert!(matches!(
        other.apply_slide_movie_caption(commit.patch()),
        Err(SlideMovieCaptionError::PatchConflict)
    ));
    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains("private movie caption"));
    assert!(!debug.contains("North"));
    Ok(())
}

#[test]
fn movie_edge_duplicates_wrong_wire_and_noncanonical_keys_fail_closed() -> TestResult<()> {
    let source = synthetic_package()?;
    let movie = message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?;
    let variants = vec![
        append_duplicate_field(&movie, 1)?,
        varint_field(1, 1, 1, 1),
        {
            let mut field = Vec::new();
            field.extend_from_slice(&[0x8a, 0x00, 0x00]);
            field
        },
        append_nested_duplicate(&movie, &[1, 11])?,
        replace_nested_field_raw(&movie, &[1, 11], &varint_field(11, CAPTIONS[0], 1, 1))?,
    ];
    for payload in variants {
        let hostile =
            with_document_message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE, payload)?;
        match Package::from_bytes(&hostile) {
            Err(_) => {},
            Ok(package) => {
                assert!(package.slide_movie_caption(0usize, 0usize).is_err());
                assert_eq!(exact_bytes(&package)?, hostile);
            },
        }
    }
    Ok(())
}

#[test]
fn caption_graph_parent_and_nested_wire_fail_closed() -> TestResult<()> {
    let source = synthetic_package()?;
    let info = caption_info_payload(0);
    let wrong_parent = caption_info_payload(1);
    for payload in [vec![0x0a, 0x01, 0x5a], wrong_parent] {
        let hostile = with_document_message_payload(
            &source,
            CAPTIONS[0],
            CAPTION_INFO_MESSAGE_TYPE,
            payload,
        )?;
        match Package::from_bytes(&hostile) {
            Err(_) => {},
            Ok(package) => {
                assert!(package.slide_movie_caption(0usize, 0usize).is_err());
                assert_eq!(exact_bytes(&package)?, hostile);
            },
        }
    }
    let nested_duplicate = append_nested_duplicate(&info, &[1, 1])?;
    let hostile = with_document_message_payload(
        &source,
        CAPTIONS[0],
        CAPTION_INFO_MESSAGE_TYPE,
        nested_duplicate,
    )?;
    match Package::from_bytes(&hostile) {
        Err(_) => {},
        Ok(package) => {
            assert!(package.slide_movie_caption(0usize, 0usize).is_err());
            assert_eq!(exact_bytes(&package)?, hostile);
        },
    }
    Ok(())
}

#[test]
fn non_file_movie_kinds_are_not_caption_targets() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut movie = tsd::MovieArchive::decode(
        message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    movie.audio_only = Some(true);
    let hostile = with_document_message_payload(
        &source,
        MOVIES[0],
        MOVIE_MESSAGE_TYPE,
        movie.encode_to_vec(),
    )?;
    let package = Package::from_bytes(&hostile)?;
    assert!(package.slide_movie_caption(0usize, 0usize).is_err());
    assert_eq!(exact_bytes(&package)?, hostile);
    Ok(())
}

#[test]
fn movie_parent_drawable_ownership_and_title_alias_fail_closed() -> TestResult<()> {
    let source = synthetic_package()?;

    let mut movie = tsd::MovieArchive::decode(
        message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    movie.super_.parent = Some(reference(999));
    let wrong_parent = with_document_message_payload(
        &source,
        MOVIES[0],
        MOVIE_MESSAGE_TYPE,
        movie.encode_to_vec(),
    )?;

    let mut movie = tsd::MovieArchive::decode(
        message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    movie.super_.title = Some(reference(CAPTIONS[0]));
    let aliased_title = with_document_message_payload(
        &source,
        MOVIES[0],
        MOVIE_MESSAGE_TYPE,
        movie.encode_to_vec(),
    )?;

    let mut slide = kn::SlideArchive::decode(message_payload(&source, SLIDE, 5)?.as_slice())?;
    slide.owned_drawables.push(reference(MOVIES[0]));
    let duplicate_owner = with_document_message_payload(&source, SLIDE, 5, slide.encode_to_vec())?;

    for hostile in [wrong_parent, aliased_title, duplicate_owner] {
        let package = Package::from_bytes(&hostile)?;
        assert!(package.slide_movie_caption(0usize, 0usize).is_err());
        assert_eq!(exact_bytes(&package)?, hostile);
    }
    Ok(())
}

#[test]
fn aggregate_and_field_info_caption_owners_fail_closed() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(80))
        .ok_or_else(|| io::Error::other("missing theme object"))?
        .archive_info
        .message_infos[0]
        .object_references
        .push(CAPTIONS[0]);
    let aggregate = replace_document_stream(&source, archive)?;
    let package = Package::from_bytes(&aggregate)?;
    assert!(package.slide_movie_caption(0usize, 0usize).is_err());
    assert_eq!(exact_bytes(&package)?, aggregate);

    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(81))
        .ok_or_else(|| io::Error::other("missing stylesheet object"))?
        .archive_info
        .message_infos[0]
        .field_infos
        .push(FieldInfo {
            path: FieldPath::new(vec![77, 1]),
            object_references: vec![CAPTIONS[0]],
            ..FieldInfo::default()
        });
    let field_owner = replace_document_stream(&source, archive)?;
    let package = Package::from_bytes(&field_owner)?;
    assert!(package.slide_movie_caption(0usize, 0usize).is_err());
    assert_eq!(exact_bytes(&package)?, field_owner);
    Ok(())
}

#[test]
fn tight_input_limit_rejects_ingress_before_publication() -> TestResult<()> {
    let source = synthetic_package()?;
    let defaults = Limits::default();
    let archive_limits = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let result = Package::from_bytes_with_options(
        &source,
        litchi_keynote::ReadOptions::new(archive_limits, litchi_keynote::SemanticLimits::default()),
    );
    assert!(result.is_err());
    Ok(())
}
