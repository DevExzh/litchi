use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tswp};
use litchi_keynote::{MovieSelector, Package, SlideMovieTitleError, SlideSelector};
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
const TITLES: [u64; 2] = [110, 111];
const TITLE_STORAGES: [u64; 2] = [120, 121];
const TITLE_PLACEMENTS: [u64; 2] = [122, 123];
const TITLE_STYLES: [u64; 2] = [124, 125];
const CAPTIONS: [u64; 2] = [130, 131];
const CAPTION_STORAGES: [u64; 2] = [140, 141];
const CAPTION_STYLES: [u64; 2] = [160, 161];
const THEME: u64 = 80;
const STYLESHEET: u64 = 81;
const PARAGRAPH_STYLE: u64 = 82;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const ROOT_UNKNOWN: u32 = 4_001;
const COMPONENT_UNKNOWN: u32 = 4_002;
const STORAGE_UNKNOWN: u32 = 4_003;
const UNKNOWN_MARKER: &[u8] = b"movie-title metadata extension";
const STORAGE_MARKER: &[u8] = b"movie-title storage extension";

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

fn movie_payload(movie: usize) -> Vec<u8> {
    tsd::MovieArchive {
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
            accessibility_description: Some("Test movie".to_owned()),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference { identifier: 2_002 }),
        poster_image_data: Some(tsp::DataReference { identifier: 2_001 }),
        // The drawable's style edge is the style shared by its title graph;
        // caption remains a separate stand-in edge in this fixture.
        style: Some(reference(TITLE_STYLES[movie])),
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
    }
    .encode_to_vec()
}

fn title_info_payload(movie: usize, parent: u64) -> Vec<u8> {
    #[allow(deprecated)]
    let info = tsa::CaptionInfoArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    parent: Some(reference(parent)),
                    caption_hidden: Some(false),
                    ..Default::default()
                },
                style: Some(reference(TITLE_STYLES[movie])),
                ..Default::default()
            },
            deprecated_storage: Some(reference(TITLE_STORAGES[movie])),
            owned_storage: Some(reference(TITLE_STORAGES[movie])),
            is_text_box: Some(true),
            ..Default::default()
        },
        placement: Some(reference(TITLE_PLACEMENTS[movie])),
        child_info_kind: Some(2),
    };
    let mut payload = info.encode_to_vec();
    append_length_delimited_field(&mut payload, COMPONENT_UNKNOWN, UNKNOWN_MARKER)
        .expect("title unknown field");
    payload
}

fn storage_payload(text: &str) -> TestResult<Vec<u8>> {
    let mut payload = tswp::StorageArchive {
        kind: Some(3),
        text: vec![text.to_owned()],
        in_document: Some(true),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, STORAGE_UNKNOWN, STORAGE_MARKER)?;
    Ok(payload)
}

fn title_theme_payload() -> TestResult<Vec<u8>> {
    let mut presets = Vec::new();
    append_length_delimited_field(&mut presets, 1, &reference(PARAGRAPH_STYLE).encode_to_vec())?;
    let mut theme_super = Vec::new();
    append_length_delimited_field(&mut theme_super, 210, &presets)?;
    let mut theme = Vec::new();
    append_length_delimited_field(&mut theme, 1, &theme_super)?;
    Ok(theme)
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
    let component = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: ids.iter().copied().map(metadata_uuid_entry).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = component;
    append_length_delimited_field(&mut payload, COMPONENT_UNKNOWN, UNKNOWN_MARKER)?;
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
        PARAGRAPH_STYLE,
        90,
        100,
        101,
        110,
        111,
        120,
        121,
        122,
        123,
        124,
        125,
        130,
        131,
        140,
        141,
        150,
        151,
        160,
        161,
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
    append_length_delimited_field(&mut payload, ROOT_UNKNOWN, UNKNOWN_MARKER)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn synthetic_package_with_titles(
    titles: [Option<&str>; 2],
    last_identifier: u64,
) -> TestResult<Vec<u8>> {
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
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: MOVIES.iter().copied().map(reference).collect(),
        drawables_z_order: MOVIES.iter().copied().map(reference).collect(),
        name: Some("Movies".to_owned()),
        in_document: true,
        ..Default::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_with_references(SLIDE, 5, slide.encode_to_vec(), MOVIES.to_vec())?,
        object(THEME, 10, title_theme_payload()?)?,
        object(STYLESHEET, 9_002, Vec::new())?,
        object(PARAGRAPH_STYLE, 9_003, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for movie in 0..MOVIES.len() {
        let movie_refs = vec![TITLES[movie], CAPTIONS[movie], TITLE_STYLES[movie]];
        objects.push(object_with_references(
            MOVIES[movie],
            MOVIE_MESSAGE_TYPE,
            movie_payload(movie),
            movie_refs,
        )?);
        if let Some(text) = titles[movie] {
            objects.push(object_with_references(
                TITLES[movie],
                CAPTION_INFO_MESSAGE_TYPE,
                title_info_payload(movie, MOVIES[movie]),
                vec![
                    TITLE_STYLES[movie],
                    TITLE_STORAGES[movie],
                    TITLE_PLACEMENTS[movie],
                ],
            )?);
            objects.push(object(
                TITLE_STORAGES[movie],
                STORAGE_MESSAGE_TYPE,
                storage_payload(text)?,
            )?);
            objects.push(object(
                TITLE_PLACEMENTS[movie],
                CAPTION_PLACEMENT_MESSAGE_TYPE,
                tsa::CaptionPlacementArchive::default().encode_to_vec(),
            )?);
            objects.push(object(
                TITLE_STYLES[movie],
                SHAPE_STYLE_MESSAGE_TYPE,
                Vec::new(),
            )?);
        } else {
            objects.push(object(TITLES[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        }
        // Keep every caption graph distinct from the title graph. The owner
        // must change only the selected MovieArchive.title edge.
        objects.push(object(CAPTIONS[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        // MovieArchive.style is a separate required graph edge even when the
        // caption itself is a stand-in; keep that object independent of both
        // title and caption styles.
        objects.push(object(
            CAPTION_STYLES[movie],
            SHAPE_STYLE_MESSAGE_TYPE,
            Vec::new(),
        )?);
    }
    let document_component = component(objects)?;
    let metadata = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata_payload(last_identifier)?,
    )?])?;
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
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
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
        .ok_or_else(|| io::Error::other("missing message").into())
}

fn title_identifier(package: &[u8], movie: u64) -> TestResult<Option<u64>> {
    Ok(
        tsd::MovieArchive::decode(message_payload(package, movie, MOVIE_MESSAGE_TYPE)?.as_slice())?
            .super_
            .title
            .map(|reference| reference.identifier),
    )
}

fn caption_identifier(package: &[u8], movie: u64) -> TestResult<Option<u64>> {
    Ok(
        tsd::MovieArchive::decode(message_payload(package, movie, MOVIE_MESSAGE_TYPE)?.as_slice())?
            .super_
            .caption
            .map(|reference| reference.identifier),
    )
}

fn storage_text(package: &[u8], storage: u64) -> TestResult<Option<String>> {
    Ok(tswp::StorageArchive::decode(
        message_payload(package, storage, STORAGE_MESSAGE_TYPE)?.as_slice(),
    )?
    .text
    .into_iter()
    .next())
}

fn decoded_metadata(package: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let bytes = metadata_stream(package)?;
    Ok(tsp::PackageMetadata::decode(bytes.as_slice())?)
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn metadata_component_raw_payload(
    metadata: &[u8],
    identifier: u64,
    versioned: bool,
) -> TestResult<Vec<u8>> {
    let field_number = if versioned { 11 } else { 3 };
    for field in WireView::parse(metadata)?.fields() {
        if field.number() != field_number || field.wire_type() != 2 {
            continue;
        }
        if tsp::ComponentInfo::decode(field.payload())?.identifier == identifier {
            return Ok(field.payload().to_vec());
        }
    }
    Err(io::Error::other("missing metadata component").into())
}

fn metadata_component(
    metadata: &tsp::PackageMetadata,
    identifier: u64,
    versioned: bool,
) -> TestResult<&tsp::ComponentInfo> {
    let components = if versioned {
        &metadata.versioned_components
    } else {
        &metadata.components
    };
    components
        .iter()
        .find(|component| component.identifier == identifier)
        .ok_or_else(|| io::Error::other("missing metadata component").into())
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
    let object = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(movie))
        .ok_or_else(|| io::Error::other("missing movie"))?;
    object
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing movie message"))?
        .data = payload;
    replace_document_archive(source, archive)
}

fn with_metadata_payload(source: &[u8], payload: Vec<u8>) -> TestResult<Vec<u8>> {
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

fn append_nested_duplicate(payload: &[u8], path: &[u32]) -> TestResult<Vec<u8>> {
    let number = *path.first().ok_or_else(|| io::Error::other("empty path"))?;
    let mut output = Vec::with_capacity(payload.len() + 8);
    let mut found = false;
    for field in WireView::parse(payload)?.fields() {
        if !found && field.number() == number {
            if path.len() == 1 {
                output.extend_from_slice(field.raw());
                output.extend_from_slice(field.raw());
                found = true;
            } else {
                if field.wire_type() != 2 {
                    return Err(io::Error::other("nested field is not length-delimited").into());
                }
                let nested = append_nested_duplicate(field.payload(), &path[1..])?;
                append_length_delimited_field(&mut output, number, &nested)?;
                found = true;
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if found {
        Ok(output)
    } else {
        Err(io::Error::other("nested field not found").into())
    }
}

fn with_title_info_payload(source: &[u8], title: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(title)
        .ok_or_else(|| io::Error::other("missing title info"))?
        .messages[0]
        .data = payload;
    replace_document_archive(source, archive)
}

fn field_info_owner(source: &[u8], target: u64) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let theme = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(THEME))
        .ok_or_else(|| io::Error::other("missing theme"))?;
    theme.archive_info.message_infos[0]
        .field_infos
        .push(FieldInfo {
            path: FieldPath::new(vec![77, 1]),
            object_references: vec![target],
            ..Default::default()
        });
    replace_document_archive(source, archive)
}

#[test]
fn reads_active_and_standin_titles_without_touching_caption_or_media() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let package = Package::from_bytes(&source)?;
    let active_title = package.slide_movie_title(0usize, 0usize)?;
    assert_eq!(active_title, Some("North title".to_owned()));
    let standin_title = package.slide_movie_title(0usize, 1usize)?;
    assert_eq!(standin_title, None);
    let caption = package.slide_movie_caption(0usize, 0usize)?;
    assert_eq!(caption, None);
    assert_eq!(title_identifier(&source, MOVIES[0])?, Some(TITLES[0]));
    assert_eq!(caption_identifier(&source, MOVIES[0])?, Some(CAPTIONS[0]));
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
fn title_no_op_is_exact_and_debug_redacts_text() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_movie_title(0usize, 0usize)?
        .set("North title")?
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(exact_bytes(commit.package())?, source);
    let changed = package
        .edit_slide_movie_title(0usize, 0usize)?
        .set("secret title")?
        .commit()?;
    let debug = format!("{:?}", changed.patch());
    assert!(!debug.contains("secret title"));
    Ok(())
}

#[test]
fn active_title_replacement_preserves_caption_media_metadata_unknowns_and_inverse() -> TestResult<()>
{
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let before = decoded_metadata(&source)?;
    let before_metadata = metadata_stream(&source)?;
    let before_root_unknown = raw_fields(&before_metadata, ROOT_UNKNOWN)?;
    let before_selected_component =
        metadata_component_raw_payload(&before_metadata, DOCUMENT_COMPONENT, false)?;
    let before_selected_unknown = raw_fields(&before_selected_component, COMPONENT_UNKNOWN)?;
    let before_unrelated_component =
        metadata_component_raw_payload(&before_metadata, UNRELATED_COMPONENT, false)?;
    let before_versioned_component =
        metadata_component_raw_payload(&before_metadata, DOCUMENT_COMPONENT, true)?;
    let before_title_info = message_payload(&source, TITLES[0], CAPTION_INFO_MESSAGE_TYPE)?;
    let before_title_unknown = raw_fields(&before_title_info, COMPONENT_UNKNOWN)?;
    let before_title_storage = message_payload(&source, TITLE_STORAGES[0], STORAGE_MESSAGE_TYPE)?;
    let before_storage_unknown = raw_fields(&before_title_storage, STORAGE_UNKNOWN)?;
    let before_title = title_identifier(&source, MOVIES[0])?;
    let before_caption = caption_identifier(&source, MOVIES[0])?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_title(0usize, 0usize)?
        .set("East title")?
        .commit()?;
    assert_eq!(
        commit.package().slide_movie_title(0usize, 0usize)?,
        Some("East title".to_owned())
    );
    assert_eq!(commit.package().slide_movie_caption(0usize, 0usize)?, None);
    assert_eq!(
        title_identifier(&exact_bytes(commit.package())?, MOVIES[0])?,
        before_title
    );
    assert_eq!(
        caption_identifier(&exact_bytes(commit.package())?, MOVIES[0])?,
        before_caption
    );
    assert_eq!(
        storage_text(&exact_bytes(commit.package())?, TITLE_STORAGES[0])?,
        Some("East title".to_owned())
    );
    assert!(
        Archive::parse(&document_stream(&source)?)?
            .object(CAPTION_STORAGES[0])
            .is_none()
    );
    let after = decoded_metadata(&exact_bytes(commit.package())?)?;
    let target = exact_bytes(commit.package())?;
    let after_metadata = metadata_stream(&target)?;
    assert_eq!(
        raw_fields(&after_metadata, ROOT_UNKNOWN)?,
        before_root_unknown
    );
    assert_eq!(
        raw_fields(
            &metadata_component_raw_payload(&after_metadata, DOCUMENT_COMPONENT, false)?,
            COMPONENT_UNKNOWN,
        )?,
        before_selected_unknown
    );
    assert_eq!(
        metadata_component_raw_payload(&after_metadata, UNRELATED_COMPONENT, false)?,
        before_unrelated_component
    );
    assert_eq!(
        metadata_component_raw_payload(&after_metadata, DOCUMENT_COMPONENT, true)?,
        before_versioned_component
    );
    assert_eq!(
        raw_fields(
            &message_payload(&target, TITLES[0], CAPTION_INFO_MESSAGE_TYPE)?,
            COMPONENT_UNKNOWN,
        )?,
        before_title_unknown
    );
    assert_eq!(
        raw_fields(
            &message_payload(&target, TITLE_STORAGES[0], STORAGE_MESSAGE_TYPE)?,
            STORAGE_UNKNOWN,
        )?,
        before_storage_unknown
    );
    assert_eq!(after.last_object_identifier, before.last_object_identifier);
    assert_eq!(after.save_token, Some(11));
    assert_eq!(
        metadata_component(&after, DOCUMENT_COMPONENT, false)?.save_token,
        Some(11)
    );
    assert_eq!(
        metadata_component(&after, UNRELATED_COMPONENT, false)?.save_token,
        Some(7)
    );
    assert_eq!(
        metadata_component(&after, DOCUMENT_COMPONENT, true)?.save_token,
        Some(3)
    );
    assert_eq!(
        after.components[0].object_uuid_map_entries.len(),
        before.components[0].object_uuid_map_entries.len()
    );
    let restored = commit
        .package()
        .apply_slide_movie_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn standin_creation_allocates_four_title_objects_and_preserves_caption() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_title(0usize, 1usize)?
        .set("South title")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(
        commit.package().slide_movie_title(0usize, 1usize)?,
        Some("South title".to_owned())
    );
    assert_eq!(commit.package().slide_movie_caption(0usize, 1usize)?, None);
    assert_eq!(title_identifier(&target, MOVIES[1])?, Some(1_002));
    let archive = Archive::parse(&document_stream(&target)?)?;
    for (identifier, type_) in [
        (1_001, SHAPE_STYLE_MESSAGE_TYPE),
        (1_002, CAPTION_INFO_MESSAGE_TYPE),
        (1_003, STORAGE_MESSAGE_TYPE),
        (1_004, CAPTION_PLACEMENT_MESSAGE_TYPE),
    ] {
        assert_eq!(
            archive
                .object(identifier)
                .and_then(|object| object.messages.first())
                .map(|message| message.type_),
            Some(type_)
        );
    }
    let metadata = decoded_metadata(&target)?;
    assert_eq!(metadata.last_object_identifier, 1_004);
    assert_eq!(metadata.save_token, Some(11));
    let selected = metadata_component(&metadata, DOCUMENT_COMPONENT, false)?;
    for identifier in 1_001..=1_004 {
        assert!(
            selected
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == identifier)
        );
    }
    assert_eq!(caption_identifier(&target, MOVIES[1])?, Some(CAPTIONS[1]));
    assert!(
        Catalog::from_bytes(&target)?
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );
    let restored = commit
        .package()
        .apply_slide_movie_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn active_title_removal_uses_fresh_standin_and_retains_old_graph() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_title(0usize, 0usize)?
        .clear()?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(commit.package().slide_movie_title(0usize, 0usize)?, None);
    assert_eq!(title_identifier(&target, MOVIES[0])?, Some(1_001));
    let archive = Archive::parse(&document_stream(&target)?)?;
    for (identifier, type_) in [
        (TITLES[0], CAPTION_INFO_MESSAGE_TYPE),
        (TITLE_STORAGES[0], STORAGE_MESSAGE_TYPE),
        (TITLE_PLACEMENTS[0], CAPTION_PLACEMENT_MESSAGE_TYPE),
        (TITLE_STYLES[0], SHAPE_STYLE_MESSAGE_TYPE),
        (1_001, STANDIN_MESSAGE_TYPE),
    ] {
        assert_eq!(
            archive
                .object(identifier)
                .and_then(|object| object.messages.first())
                .map(|message| message.type_),
            Some(type_)
        );
    }
    let metadata = decoded_metadata(&target)?;
    assert_eq!(metadata.last_object_identifier, 1_001);
    assert_eq!(metadata.save_token, Some(11));
    let selected = metadata_component(&metadata, DOCUMENT_COMPONENT, false)?;
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_001)
    );
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == TITLES[0])
    );
    assert_eq!(caption_identifier(&target, MOVIES[0])?, Some(CAPTIONS[0]));
    let restored = commit
        .package()
        .apply_slide_movie_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selectors_non_file_movies_and_conflicts_are_typed_and_atomic() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_movie_title(SlideSelector::index(9), 0usize),
        Err(SlideMovieTitleError::SlidePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_movie_title(0usize, MovieSelector::index(9)),
        Err(SlideMovieTitleError::MoviePositionNotFound { .. })
    ));
    let mut movie = tsd::MovieArchive::decode(
        message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    movie.audio_only = Some(true);
    let hostile = with_movie_payload(&source, MOVIES[0], movie.encode_to_vec())?;
    let hostile_package = Package::from_bytes(&hostile)?;
    assert!(hostile_package.slide_movie_title(0usize, 0usize).is_err());
    assert_eq!(exact_bytes(&hostile_package)?, hostile);
    let commit = package
        .edit_slide_movie_title(0usize, 0usize)?
        .set("new title")?
        .commit()?;
    let other = Package::from_bytes(&synthetic_package_with_titles(
        [Some("Other"), None],
        METADATA_LAST_IDENTIFIER,
    )?)?;
    assert!(matches!(
        other.apply_slide_movie_title(commit.patch()),
        Err(SlideMovieTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn malformed_title_edge_info_storage_parent_and_duplicate_fields_fail_closed() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let movie = tsd::MovieArchive::decode(
        message_payload(&source, MOVIES[0], MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let malformed_movie = with_movie_payload(&source, MOVIES[0], vec![0x0a, 0x01, 0x5a])?;
    let duplicate_title = with_movie_payload(
        &source,
        MOVIES[0],
        append_nested_duplicate(&movie.encode_to_vec(), &[1, 10])?,
    )?;
    let wrong_parent = with_title_info_payload(&source, TITLES[0], title_info_payload(0, 999))?;
    let mut title = title_info_payload(0, MOVIES[0]);
    title.truncate(title.len().saturating_sub(1));
    let malformed_info = {
        let mut archive = Archive::parse(&document_stream(&source)?)?;
        archive
            .object_mut(TITLES[0])
            .ok_or_else(|| io::Error::other("missing title"))?
            .messages[0]
            .data = title;
        replace_document_archive(&source, archive)?
    };
    let malformed_storage = {
        let mut archive = Archive::parse(&document_stream(&source)?)?;
        archive
            .object_mut(TITLE_STORAGES[0])
            .ok_or_else(|| io::Error::other("missing storage"))?
            .messages[0]
            .data = vec![0x18, 0x01];
        replace_document_archive(&source, archive)?
    };
    for hostile in [
        malformed_movie,
        duplicate_title,
        wrong_parent,
        malformed_info,
        malformed_storage,
    ] {
        let package = Package::from_bytes(&hostile)?;
        let result = package
            .edit_slide_movie_title(0usize, 0usize)
            .and_then(|edit| edit.set("changed"))
            .and_then(|edit| edit.commit());
        assert!(matches!(
            result,
            Err(SlideMovieTitleError::InvalidSource)
                | Err(SlideMovieTitleError::UnsupportedDependency)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
    }
    Ok(())
}

#[test]
fn shared_aggregate_and_field_info_title_owners_are_rejected() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(THEME)
        .ok_or_else(|| io::Error::other("missing theme"))?
        .archive_info
        .message_infos[0]
        .object_references
        .push(TITLES[0]);
    let aggregate = replace_document_archive(&source, archive)?;
    let field = field_info_owner(&source, TITLES[0])?;
    for hostile in [aggregate, field] {
        let package = Package::from_bytes(&hostile)?;
        let result = package
            .edit_slide_movie_title(0usize, 0usize)
            .and_then(|edit| edit.set("changed"))
            .and_then(|edit| edit.commit());
        assert!(matches!(
            result,
            Err(SlideMovieTitleError::UnsupportedDependency)
                | Err(SlideMovieTitleError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
    }
    Ok(())
}

#[test]
fn aliases_and_reserved_metadata_ids_fail_closed_or_skip_safely() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(MOVIES[1])
        .ok_or_else(|| io::Error::other("missing movie"))?
        .messages[0]
        .data = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            title: Some(reference(TITLES[0])),
            ..tsd::MovieArchive::decode(movie_payload(1).as_slice())?.super_
        },
        ..tsd::MovieArchive::decode(movie_payload(1).as_slice())?
    }
    .encode_to_vec();
    let alias = replace_document_archive(&source, archive)?;
    let package = Package::from_bytes(&alias)?;
    let result = package
        .edit_slide_movie_title(0usize, 1usize)
        .and_then(|edit| edit.set("created"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(SlideMovieTitleError::UnsupportedDependency) | Err(SlideMovieTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, alias);
    let mut metadata = metadata_stream(&source)?;
    let field = WireView::parse(&metadata)?
        .fields()
        .find(|field| field.number() == 3)
        .ok_or_else(|| io::Error::other("missing metadata component"))?
        .raw()
        .to_vec();
    metadata.extend_from_slice(&field);
    let duplicate = with_metadata_payload(&source, metadata)?;
    let package = Package::from_bytes(&duplicate)?;
    let result = package
        .edit_slide_movie_title(0usize, 1usize)
        .and_then(|edit| edit.set("created"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(SlideMovieTitleError::InvalidSource) | Err(SlideMovieTitleError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, duplicate);
    let overflow = synthetic_package_with_titles([Some("North title"), None], u64::MAX)?;
    let package = Package::from_bytes(&overflow)?;
    let result = package
        .edit_slide_movie_title(0usize, 1usize)
        .and_then(|edit| edit.set("created"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(SlideMovieTitleError::InvalidSource)
            | Err(SlideMovieTitleError::LimitExceeded { .. })
            | Err(SlideMovieTitleError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, overflow);
    Ok(())
}

#[test]
fn missing_metadata_and_tight_ingress_are_atomic() -> TestResult<()> {
    let source =
        synthetic_package_with_titles([Some("North title"), None], METADATA_LAST_IDENTIFIER)?;
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
    let result = package
        .edit_slide_movie_title(0usize, 0usize)
        .and_then(|edit| edit.set("changed"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(SlideMovieTitleError::UnsupportedSource)
            | Err(SlideMovieTitleError::InvalidSource)
            | Err(SlideMovieTitleError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, no_metadata);
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let result = Package::from_bytes_with_options(
        &source,
        litchi_keynote::ReadOptions::new(limits, litchi_keynote::SemanticLimits::default()),
    );
    assert!(result.is_err());
    Ok(())
}
