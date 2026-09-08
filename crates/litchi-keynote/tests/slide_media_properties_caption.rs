//! Focused properties-read coverage for a semantically captioned movie.
//!
//! This synthetic package deliberately keeps the movie data and metadata
//! small.  It exercises the transitive caption-style witness without relying
//! on the materialized-asset closure required by duplication or mutation.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use litchi_keynote::{MovieSelector, Package, SlideMediaPropertiesError, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const MOVIES: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const CAPTIONS: [u64; 2] = [130, 131];
const STYLES: [u64; 2] = [160, 161];
const THEME: u64 = 80;
const STYLESHEET: u64 = 81;
const PARAGRAPH_STYLE: u64 = 82;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;

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

fn movie_object(movie: usize) -> TestResult<ArchiveObject> {
    let mut object = object_with_references(
        MOVIES[movie],
        MOVIE_MESSAGE_TYPE,
        movie_payload(movie),
        vec![TITLES[movie], CAPTIONS[movie], STYLES[movie]],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.data_references = vec![2_002, 2_001];
    info.field_infos = vec![
        FieldInfo {
            path: FieldPath::new(vec![14]),
            r#type: Some(FieldType::DataReference),
            data_references: vec![2_002],
            ..FieldInfo::default()
        },
        FieldInfo {
            path: FieldPath::new(vec![15]),
            r#type: Some(FieldType::DataReference),
            data_references: vec![2_001],
            ..FieldInfo::default()
        },
    ];
    Ok(object)
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
                ..tsd::GeometryArchive::default()
            }),
            parent: Some(reference(SLIDE)),
            title: Some(reference(TITLES[movie])),
            caption: Some(reference(CAPTIONS[movie])),
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

fn metadata_component_payload(
    identifier: u64,
    locator: &str,
    token: u64,
    object_identifiers: &[u64],
) -> Vec<u8> {
    tsp::ComponentInfo {
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
    .encode_to_vec()
}

fn metadata_payload() -> Vec<u8> {
    let ids = [
        1, 2, 3, 4, 80, 81, 82, 90, 100, 101, 110, 111, 130, 131, 140, 141, 150, 151, 160, 161,
    ];
    let document = metadata_component_payload(DOCUMENT_COMPONENT, "Document", 10, &ids);
    let unrelated = metadata_component_payload(UNRELATED_COMPONENT, "Unrelated", 7, &[900]);
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
    litchi_iwa_common::wire::append_varint_field(&mut payload, 1, 1_000)
        .expect("metadata last identifier must encode");
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 3, &document)
        .expect("document component must encode");
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 3, &unrelated)
        .expect("unrelated component must encode");
    litchi_iwa_common::wire::append_varint_field(&mut payload, 8, 10)
        .expect("metadata token must encode");
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 11, &versioned)
        .expect("versioned component must encode");
    payload
}

fn caption_theme_payload() -> Vec<u8> {
    let mut presets = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut presets,
        1,
        &reference(PARAGRAPH_STYLE).encode_to_vec(),
    )
    .expect("theme preset must encode");
    let mut theme_super = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(&mut theme_super, 210, &presets)
        .expect("theme super must encode");
    let mut theme = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(&mut theme, 1, &theme_super)
        .expect("theme must encode");
    theme
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(THEME),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(STYLESHEET),
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
        object(THEME, 10, caption_theme_payload())?,
        object(STYLESHEET, 9_002, Vec::new())?,
        object(PARAGRAPH_STYLE, 9_003, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for movie in 0..MOVIES.len() {
        objects.push(movie_object(movie)?);
        objects.push(object(TITLES[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(CAPTIONS[movie], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(STYLES[movie], SHAPE_STYLE_MESSAGE_TYPE, Vec::new())?);
    }
    let document_component = component(objects)?;
    let metadata = component(vec![object(METADATA_OBJECT, 11_006, metadata_payload())?])?;
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

fn caption_info_location(source: &[u8], movie_identifier: u64) -> TestResult<u64> {
    let archive = Archive::parse(&document_stream(source)?)?;
    archive
        .objects
        .iter()
        .find_map(|object| {
            let message = object
                .messages
                .iter()
                .find(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)?;
            let info = tsa::CaptionInfoArchive::decode(message.data.as_slice()).ok()?;
            (info.child_info_kind == Some(1)
                && info.super_.super_.super_.parent.as_ref()?.identifier == movie_identifier)
                .then_some(object.archive_info.identifier?)
        })
        .ok_or_else(|| io::Error::other("active caption info is missing").into())
}

fn mutate_caption_info(
    source: &[u8],
    movie_identifier: u64,
    edit: impl FnOnce(&mut Archive, u64, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let caption_identifier = caption_info_location(source, movie_identifier)?;
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let index = archive
        .object(caption_identifier)
        .ok_or_else(|| io::Error::other("caption info object is missing"))?
        .messages
        .iter()
        .position(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("caption info message is missing"))?;
    edit(&mut archive, caption_identifier, index)?;
    replace_document_stream(source, archive)
}

fn with_extra_caption_style_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_caption_info(source, MOVIES[0], |archive, identifier, index| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("caption info object is missing"))?;
        let info = tsa::CaptionInfoArchive::decode(object.messages[index].data.as_slice())?;
        let style = info
            .super_
            .super_
            .style
            .as_ref()
            .ok_or_else(|| io::Error::other("caption style edge is missing"))?
            .identifier;
        object.archive_info.message_infos[index]
            .object_references
            .push(style);
        Ok(())
    })
}

fn with_wrong_caption_parent(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_caption_info(source, MOVIES[0], |archive, identifier, index| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("caption info object is missing"))?;
        let mut info = tsa::CaptionInfoArchive::decode(object.messages[index].data.as_slice())?;
        info.super_.super_.super_.parent = Some(reference(MOVIES[1]));
        object.messages[index].data = info.encode_to_vec();
        object.archive_info.message_infos[index].length =
            u32::try_from(object.messages[index].data.len())?;
        Ok(())
    })
}

fn with_wrong_caption_style_type(source: &[u8]) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(source)?)?;
    let caption_identifier = caption_info_location(source, MOVIES[0])?;
    let caption = archive
        .object(caption_identifier)
        .ok_or_else(|| io::Error::other("caption info object is missing"))?;
    let message = caption
        .messages
        .iter()
        .find(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("caption info message is missing"))?;
    let info = tsa::CaptionInfoArchive::decode(message.data.as_slice())?;
    let style_identifier = info
        .super_
        .super_
        .style
        .as_ref()
        .ok_or_else(|| io::Error::other("caption style edge is missing"))?
        .identifier;
    mutate_style_type(source, style_identifier)
}

fn mutate_style_type(source: &[u8], style_identifier: u64) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let style = archive
        .object_mut(style_identifier)
        .ok_or_else(|| io::Error::other("caption style object is missing"))?;
    let index = style
        .messages
        .iter()
        .position(|message| message.type_ == SHAPE_STYLE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("caption shape style message is missing"))?;
    style.messages[index].type_ = STANDIN_MESSAGE_TYPE;
    style.archive_info.message_infos[index].type_ = STANDIN_MESSAGE_TYPE;
    replace_document_stream(source, archive)
}

fn with_wrong_caption_info_type(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_caption_info(source, MOVIES[0], |archive, identifier, index| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("caption info object is missing"))?;
        object.messages[index].type_ = SHAPE_STYLE_MESSAGE_TYPE;
        object.archive_info.message_infos[index].type_ = SHAPE_STYLE_MESSAGE_TYPE;
        Ok(())
    })
}

/// Reproduce the native source-built aggregate style edges on the sparse
/// fixture.  Caption/title creation owns these style objects, while the
/// MovieArchive header records them as private witnesses in addition to its
/// public `style` edge.  Native duplicate acceptance relies on that metadata.
fn add_private_style_witnesses(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let movie_payload = archive
        .object(MOVIES[0])
        .ok_or_else(|| io::Error::other("movie object is missing"))?
        .messages
        .iter()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("movie message is missing"))?
        .data
        .clone();
    let movie = tsd::MovieArchive::decode(movie_payload.as_slice())?;
    let mut styles = Vec::new();
    for reference in [movie.super_.title, movie.super_.caption]
        .into_iter()
        .flatten()
    {
        let info_object = archive
            .object(reference.identifier)
            .ok_or_else(|| io::Error::other("movie text object is missing"))?;
        let info_message = info_object
            .messages
            .iter()
            .find(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("movie text info message is missing"))?;
        let info = tsa::CaptionInfoArchive::decode(info_message.data.as_slice())?;
        let style = info
            .super_
            .super_
            .style
            .as_ref()
            .ok_or_else(|| io::Error::other("movie text style is missing"))?
            .identifier;
        if !styles.contains(&style) {
            styles.push(style);
        }
    }
    let movie = archive
        .object_mut(MOVIES[0])
        .ok_or_else(|| io::Error::other("movie object is missing"))?;
    let info = movie
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("movie message metadata is missing"))?;
    for style in styles {
        if !info.object_references.contains(&style) {
            info.object_references.push(style);
        }
    }
    replace_document_stream(source, archive)
}

fn movie_style_witnesses(source: &[u8]) -> TestResult<(Vec<u64>, Vec<u64>)> {
    let archive = Archive::parse(&document_stream(source)?)?;
    let movie = archive
        .object(MOVIES[0])
        .ok_or_else(|| io::Error::other("movie object is missing"))?;
    let message = movie
        .messages
        .iter()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("movie message is missing"))?;
    let movie_archive = tsd::MovieArchive::decode(message.data.as_slice())?;
    let mut private_styles = Vec::new();
    for reference in [movie_archive.super_.title, movie_archive.super_.caption]
        .into_iter()
        .flatten()
    {
        let info_object = archive
            .object(reference.identifier)
            .ok_or_else(|| io::Error::other("movie text object is missing"))?;
        let info_message = info_object
            .messages
            .iter()
            .find(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("movie text info message is missing"))?;
        let info = tsa::CaptionInfoArchive::decode(info_message.data.as_slice())?;
        let style = info
            .super_
            .super_
            .style
            .as_ref()
            .ok_or_else(|| io::Error::other("movie text style is missing"))?
            .identifier;
        if !private_styles.contains(&style) {
            private_styles.push(style);
        }
    }
    let header = &movie.archive_info.message_infos[0].object_references;
    for style in &private_styles {
        assert_eq!(header.iter().filter(|known| *known == style).count(), 1);
    }
    assert!(private_styles.iter().any(|style| {
        movie_archive
            .style
            .as_ref()
            .is_none_or(|movie_style| movie_style.identifier != *style)
    }));
    Ok((private_styles, header.clone()))
}

fn captioned_movie_package() -> TestResult<Vec<u8>> {
    let source = synthetic_package()?;
    let title = Package::from_bytes(&source)?
        .edit_slide_movie_title(SlideSelector::index(0), MovieSelector::index(0))?
        .set("Movie title")?
        .commit()?;
    let caption = title
        .package()
        .edit_slide_movie_caption(SlideSelector::index(0), MovieSelector::index(0))?
        .set("Movie caption")?
        .commit()?;
    let source = exact_bytes(caption.package())?;
    add_private_style_witnesses(&source)
}

#[test]
fn created_captioned_movie_reads_properties_and_retains_private_style_witnesses() -> TestResult {
    let source = captioned_movie_package()?;
    let (private_styles, header) = movie_style_witnesses(&source)?;
    assert!(!private_styles.is_empty());
    assert!(header.len() > private_styles.len());
    let package = Package::from_bytes(&source)?;
    let properties =
        package.slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))?;
    assert_eq!(properties.accessibility_description(), Some("Test movie"));
    assert_eq!(exact_bytes(&package)?, source);
    assert_eq!(
        package.slide_movie_title(0usize, 0usize)?,
        Some("Movie title".to_owned())
    );
    assert_eq!(
        package.slide_movie_caption(0usize, 0usize)?,
        Some("Movie caption".to_owned())
    );
    Ok(())
}

#[test]
fn caption_style_witness_mutations_reject_without_source_rewrite() -> TestResult<()> {
    let source = captioned_movie_package()?;
    let hostile = [
        with_extra_caption_style_reference(&source)?,
        with_wrong_caption_parent(&source)?,
        with_wrong_caption_style_type(&source)?,
        with_wrong_caption_info_type(&source)?,
    ];
    for hostile in hostile {
        let package = Package::from_bytes(&hostile)?;
        let result =
            package.slide_media_properties(SlideSelector::index(0), MovieSelector::index(0));
        assert!(matches!(
            result,
            Err(SlideMediaPropertiesError::InvalidSource)
                | Err(SlideMediaPropertiesError::UnsupportedDependency)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
    }
    Ok(())
}
