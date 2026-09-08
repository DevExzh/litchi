//! Selector-first Keynote movie geometry integration tests.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, MessageInfo, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use litchi_keynote::slide::media::{
    Point, Size,
    geometry::{MovieFlipAxis, MovieGeometry, MovieTransform},
};
use litchi_keynote::{Package, ReadOptions, SlideMovieGeometryError};
use prost::Message as _;

const DOCUMENT: &str = "Index/Document.iwa";
const METADATA: &str = "Index/Metadata.iwa";
const MOVIE_TYPE: u32 = 3_007;
const SLIDE_TYPE: u32 = 5;
const TABLE_INFO_TYPE: u32 = 6_000;
const TABLE_MODEL_TYPE: u32 = 6_001;
const HEADER_BUCKET_TYPE: u32 = 6_006;
const TABLE_STYLE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_TYPE: u32 = 6_247;
const STYLESHEET_TYPE: u32 = 401;
const MOVIES: [u64; 2] = [100, 101];
const AUDIO: u64 = 102;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TITLE: [u64; 2] = [110, 111];
const CAPTION: [u64; 2] = [130, 131];
const STYLE: [u64; 2] = [160, 161];
const UNKNOWN: u32 = 4_091;
const UNKNOWN_BYTES: &[u8] = b"movie-geometry-unknown-extension";
const GROUP_BYTES: &[u8] = b"movie-geometry-balanced-group";
type R<T> = Result<T, Box<dyn std::error::Error>>;

fn r(id: u64) -> tsp::Reference {
    tsp::Reference {
        identifier: id,
        ..Default::default()
    }
}

fn g(x: f32, y: f32, w: f32, h: f32) -> R<MovieGeometry> {
    Ok(MovieGeometry::new(
        Point { x, y },
        Size {
            width: w,
            height: h,
        },
    )?)
}

fn object(id: u64, ty: u32, data: Vec<u8>) -> R<ArchiveObject> {
    Ok(ArchiveObject::new(
        id,
        vec![RawMessage { type_: ty, data }],
    )?)
}

fn object_refs(id: u64, ty: u32, data: Vec<u8>, refs: Vec<u64>) -> R<ArchiveObject> {
    let mut value = object(id, ty, data)?;
    value.archive_info.message_infos[0].object_references = refs;
    Ok(value)
}

fn movie_payload(i: usize, file: bool) -> Vec<u8> {
    tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                size: Some(tsp::Size {
                    width: 800.0,
                    height: 300.0,
                }),
                flags: Some(0x20),
                angle: Some(17.5),
            }),
            parent: Some(r(SLIDE)),
            title: Some(r(TITLE[i.min(1)])),
            caption: Some(r(CAPTION[i.min(1)])),
            accessibility_description: Some("geometry locality sentinel".into()),
            ..Default::default()
        },
        movie_data: file.then_some(tsp::DataReference { identifier: 2_002 }),
        poster_image_data: file.then_some(tsp::DataReference { identifier: 2_001 }),
        style: Some(r(STYLE[i.min(1)])),
        original_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        natural_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        flags: Some(u32::from(!file)),
        ..Default::default()
    }
    .encode_to_vec()
}

fn uuid(id: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier: id,
        uuid: tsp::Uuid {
            lower: id + 10_000,
            upper: id + 20_000,
        },
    }
}

fn component_info(id: u64, locator: &str, token: u64, ids: &[u64]) -> R<Vec<u8>> {
    let mut value = tsp::ComponentInfo {
        identifier: id,
        preferred_locator: locator.into(),
        locator: Some(locator.into()),
        save_token: Some(token),
        object_uuid_map_entries: ids.iter().copied().map(uuid).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut value, 4_002, b"component geometry extension")?;
    Ok(value)
}

fn metadata() -> R<Vec<u8>> {
    let ids = [
        1, 2, 3, 4, 80, 81, 90, 100, 101, AUDIO, 110, 111, 130, 131, 160, 161,
    ];
    let current = component_info(1, "Document", 10, &ids)?;
    let other = component_info(2, "Unrelated", 7, &[900])?;
    let versioned = tsp::ComponentInfo {
        identifier: 1,
        preferred_locator: "Document".into(),
        locator: Some("Document".into()),
        save_token: Some(3),
        object_uuid_map_entries: vec![uuid(901)],
        ..Default::default()
    }
    .encode_to_vec();
    let mut value = Vec::new();
    append_varint_field(&mut value, 1, 1_000)?;
    append_length_delimited_field(&mut value, 3, &current)?;
    append_length_delimited_field(&mut value, 3, &other)?;
    append_varint_field(&mut value, 8, 10)?;
    append_length_delimited_field(&mut value, 11, &versioned)?;
    append_length_delimited_field(&mut value, 4_001, b"metadata geometry extension")?;
    Ok(value)
}

fn component(objects: Vec<ArchiveObject>) -> R<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn source() -> R<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: r(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: r(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![r(SLIDE_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: r(81),
        ..Default::default()
    };
    #[allow(deprecated)]
    let node = kn::SlideNodeArchive {
        slide: Some(r(SLIDE)),
        ..Default::default()
    };
    let slide = kn::SlideArchive {
        style: r(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: [MOVIES[0], AUDIO, MOVIES[1]].into_iter().map(r).collect(),
        drawables_z_order: [MOVIES[0], AUDIO, MOVIES[1]].into_iter().map(r).collect(),
        name: Some("Geometry".into()),
        in_document: true,
        ..Default::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_refs(
            SLIDE,
            5,
            slide.encode_to_vec(),
            vec![MOVIES[0], AUDIO, MOVIES[1]],
        )?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for i in 0..MOVIES.len() {
        objects.push(object_refs(
            MOVIES[i],
            MOVIE_TYPE,
            movie_payload(i, true),
            vec![TITLE[i], CAPTION[i], STYLE[i]],
        )?);
        objects.push(object(TITLE[i], 3_097, Vec::new())?);
        objects.push(object(CAPTION[i], 3_097, Vec::new())?);
        objects.push(object(STYLE[i], 2_025, Vec::new())?);
    }
    objects.push(object_refs(
        AUDIO,
        MOVIE_TYPE,
        movie_payload(0, false),
        vec![TITLE[0], CAPTION[0], STYLE[0]],
    )?);
    let metadata = component(vec![object(300, 11_006, metadata()?)?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"geometry sentinel".as_slice()),
            ("Data/movie.mov", b"synthetic movie bytes".as_slice()),
            ("Data/poster.png", b"synthetic poster bytes".as_slice()),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT, component(objects)?.as_slice()),
            (METADATA, metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn bytes(package: &Package) -> R<Vec<u8>> {
    let mut value = Vec::new();
    package.write_to(&mut value)?;
    Ok(value)
}
fn document_stream(package: &[u8]) -> R<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|e| e.name() == DOCUMENT)
        .ok_or_else(|| io::Error::other("document"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}
fn movie_payload_at(package: &[u8], id: u64) -> R<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .object(id)
        .and_then(|o| o.messages.iter().find(|m| m.type_ == MOVIE_TYPE))
        .map(|m| m.data.clone())
        .ok_or_else(|| io::Error::other("movie").into())
}
fn replace_document(source: &[u8], archive: Archive) -> R<Vec<u8>> {
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .map(|e| {
                if e.name() == DOCUMENT {
                    (e.name(), replacement.as_slice())
                } else {
                    (e.name(), e.data())
                }
            })
            .collect::<Vec<_>>(),
        Limits::default(),
    )?)
}

fn metadata_stream(package: &[u8]) -> R<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA)
        .ok_or_else(|| io::Error::other("metadata"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn replace_metadata(source: &[u8], archive: Archive) -> R<Vec<u8>> {
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .map(|entry| {
                if entry.name() == METADATA {
                    (entry.name(), replacement.as_slice())
                } else {
                    (entry.name(), entry.data())
                }
            })
            .collect::<Vec<_>>(),
        Limits::default(),
    )?)
}

fn with_metadata(source: &[u8], edit: impl FnOnce(&mut tsp::PackageMetadata)) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&metadata_stream(source)?)?;
    let object = archive
        .object_mut(300)
        .ok_or_else(|| io::Error::other("metadata object"))?;
    let mut metadata = tsp::PackageMetadata::decode(object.messages[0].data.as_slice())?;
    edit(&mut metadata);
    object.messages[0].data = metadata.encode_to_vec();
    replace_metadata(source, archive)
}

fn with_message_alias(source: &[u8], id: u64, alias_type: u32) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let object = archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("alias object"))?;
    let data = object
        .messages
        .first()
        .ok_or_else(|| io::Error::other("alias message"))?
        .data
        .clone();
    object.messages.push(RawMessage {
        type_: alias_type,
        data: data.clone(),
    });
    object
        .archive_info
        .message_infos
        .push(MessageInfo::new(alias_type, u32::try_from(data.len())?));
    replace_document(source, archive)
}

fn with_foreign_inbound(source: &[u8], target: u64, mode: ForeignInbound) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let mut foreign = object_refs(
        900,
        MOVIE_TYPE,
        movie_payload(0, true),
        vec![TITLE[0], CAPTION[0], STYLE[0]],
    )?;
    let info = &mut foreign.archive_info.message_infos[0];
    match mode {
        ForeignInbound::Object => info.object_references.push(target),
        ForeignInbound::Data => info.data_references.push(target),
        ForeignInbound::Field => info.field_infos.push(FieldInfo {
            path: FieldPath::new(vec![99]),
            object_references: vec![target],
            ..FieldInfo::default()
        }),
        ForeignInbound::FieldData => info.field_infos.push(FieldInfo {
            path: FieldPath::new(vec![100]),
            data_references: vec![target],
            ..FieldInfo::default()
        }),
    }
    archive.objects.push(foreign);
    replace_document(source, archive)
}

fn with_foreign_z_order(source: &[u8]) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let foreign = object_refs(
        900,
        MOVIE_TYPE,
        movie_payload(0, true),
        vec![TITLE[0], CAPTION[0], STYLE[0]],
    )?;
    let slide = archive
        .object_mut(SLIDE)
        .ok_or_else(|| io::Error::other("slide"))?;
    let message = slide
        .messages
        .iter_mut()
        .find(|message| message.type_ == 5)
        .ok_or_else(|| io::Error::other("slide message"))?;
    let mut payload = kn::SlideArchive::decode(message.data.as_slice())?;
    payload.owned_drawables.push(r(900));
    payload.drawables_z_order.push(r(900));
    message.data = payload.encode_to_vec();
    slide.archive_info.message_infos[0]
        .object_references
        .push(900);
    archive.objects.push(foreign);
    replace_document(source, archive)
}

fn with_movie_data_id(source: &[u8], id: u64, data_identifier: u64) -> R<Vec<u8>> {
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(source, id)?.as_slice())?;
    movie.movie_data = Some(tsp::DataReference {
        identifier: data_identifier,
    });
    with_movie(source, id, movie.encode_to_vec())
}

fn with_duplicate_movie_data(source: &[u8], id: u64) -> R<Vec<u8>> {
    let mut payload = movie_payload_at(source, id)?;
    let data = tsp::DataReference { identifier: 2_002 }.encode_to_vec();
    append_length_delimited_field(&mut payload, 14, &data)?;
    with_movie(source, id, payload)
}

fn with_movie_archive_data_refs(source: &[u8], id: u64, refs: Vec<u64>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .data_references = refs;
    replace_document(source, archive)
}

fn with_style_message_type(source: &[u8], id: u64, message_type: u32) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let style = archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("style"))?;
    let message = style
        .messages
        .iter_mut()
        .find(|message| message.type_ == 2_025)
        .ok_or_else(|| io::Error::other("style message"))?;
    message.type_ = message_type;
    style.archive_info.message_infos[0].type_ = message_type;
    replace_document(source, archive)
}

fn with_duplicate_style_message(source: &[u8], id: u64, message_type: u32) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let style = archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("style"))?;
    style.messages.push(RawMessage {
        type_: message_type,
        data: Vec::new(),
    });
    style
        .archive_info
        .message_infos
        .push(MessageInfo::new(message_type, 0));
    replace_document(source, archive)
}

#[derive(Clone, Copy)]
enum ForeignInbound {
    Object,
    Data,
    Field,
    FieldData,
}
fn with_alias_member(source: &[u8], id: u64) -> R<Vec<u8>> {
    let archive = Archive::parse(&document_stream(source)?)?;
    let alias = archive
        .object(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .clone();
    let alias_member = component(vec![alias])?;
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push(("Index/Alias.iwa", alias_member.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}
fn with_movie(source: &[u8], id: u64, payload: Vec<u8>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .messages
        .iter_mut()
        .find(|m| m.type_ == MOVIE_TYPE)
        .ok_or_else(|| io::Error::other("message"))?
        .data = payload;
    replace_document(source, archive)
}

fn with_transform_fields(
    source: &[u8],
    id: u64,
    flags: Option<u32>,
    angle: Option<f32>,
) -> R<Vec<u8>> {
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(source, id)?.as_slice())?;
    let geometry = movie
        .super_
        .geometry
        .as_mut()
        .ok_or_else(|| io::Error::other("geometry"))?;
    geometry.flags = flags;
    geometry.angle = angle;
    with_movie(source, id, movie.encode_to_vec())
}

fn with_original_size(source: &[u8], id: u64, size: Option<tsp::Size>) -> R<Vec<u8>> {
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(source, id)?.as_slice())?;
    movie.original_size = size;
    with_movie(source, id, movie.encode_to_vec())
}

fn with_locked_movie(source: &[u8], id: u64) -> R<Vec<u8>> {
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(source, id)?.as_slice())?;
    movie.super_.locked = Some(true);
    with_movie(source, id, movie.encode_to_vec())
}

fn with_geometry_extra(source: &[u8], id: u64, extra: &[u8]) -> R<Vec<u8>> {
    let payload = movie_payload_at(source, id)?;
    let root = litchi_iwa_common::wire::WireView::parse(&payload)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("movie super"))?;
    let drawable = litchi_iwa_common::wire::WireView::parse(super_field.payload())?;
    let geometry_field = drawable
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("movie geometry"))?;
    let mut geometry = geometry_field.payload().to_vec();
    geometry.extend_from_slice(extra);

    let mut drawable_output = Vec::new();
    for field in drawable.fields() {
        if field.number() == 1 {
            append_length_delimited_field(&mut drawable_output, 1, &geometry)?;
        } else {
            drawable_output.extend_from_slice(field.raw());
        }
    }
    let mut output = Vec::new();
    for field in root.fields() {
        if field.number() == 1 {
            append_length_delimited_field(&mut output, 1, &drawable_output)?;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    with_movie(source, id, output)
}

fn append_fixed32(payload: &mut Vec<u8>, field: u32, value: f32) {
    let mut key = u64::from(field) << 3 | 5;
    while key >= 0x80 {
        payload.push((key as u8 & 0x7f) | 0x80);
        key >>= 7;
    }
    payload.push(key as u8);
    payload.extend_from_slice(&value.to_le_bytes());
}

fn with_noncanonical_super_key(source: &[u8], id: u64) -> R<Vec<u8>> {
    let mut payload = movie_payload_at(source, id)?;
    if payload.first() != Some(&0x0a) {
        return Err(io::Error::other("unexpected movie super key").into());
    }
    payload[0] = 0x8a;
    payload.insert(1, 0);
    with_movie(source, id, payload)
}

fn with_noncanonical_parent_key(source: &[u8], id: u64) -> R<Vec<u8>> {
    let payload = movie_payload_at(source, id)?;
    let root = litchi_iwa_common::wire::WireView::parse(&payload)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("movie super"))?;
    let drawable = litchi_iwa_common::wire::WireView::parse(super_field.payload())?;
    let parent_field = drawable
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("movie parent"))?;
    let mut reference = parent_field.payload().to_vec();
    if reference.first() != Some(&0x08) {
        return Err(io::Error::other("unexpected reference key").into());
    }
    reference[0] = 0x88;
    reference.insert(1, 0);

    let mut drawable_output = Vec::new();
    for field in drawable.fields() {
        if field.number() == 2 {
            append_length_delimited_field(&mut drawable_output, 2, &reference)?;
        } else {
            drawable_output.extend_from_slice(field.raw());
        }
    }
    let mut output = Vec::new();
    for field in root.fields() {
        if field.number() == 1 {
            append_length_delimited_field(&mut output, 1, &drawable_output)?;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    with_movie(source, id, output)
}

fn append_wire_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn append_noncanonical_length_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
) -> R<()> {
    append_wire_varint(output, (u64::from(field_number) << 3) | 2);
    let mut length = Vec::new();
    append_wire_varint(&mut length, u64::try_from(payload.len())?);
    let last = length
        .last_mut()
        .ok_or_else(|| io::Error::other("length encoding"))?;
    *last |= 0x80;
    length.push(0);
    output.extend_from_slice(&length);
    output.extend_from_slice(payload);
    Ok(())
}

fn rewrite_noncanonical_length(payload: &[u8], path: &[u32]) -> R<Vec<u8>> {
    let field_number = *path
        .first()
        .ok_or_else(|| io::Error::other("empty wire path"))?;
    let view = litchi_iwa_common::wire::WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len().saturating_add(1));
    let mut matched = 0usize;
    for field in view.fields() {
        if field.number() != field_number {
            output.extend_from_slice(field.raw());
            continue;
        }
        matched = matched.saturating_add(1);
        if matched != 1 || field.wire_type() != 2 {
            return Err(io::Error::other("wire path is not unique and nested").into());
        }
        if path.len() == 1 {
            append_noncanonical_length_field(&mut output, field_number, field.payload())?;
        } else {
            let nested = rewrite_noncanonical_length(field.payload(), &path[1..])?;
            append_length_delimited_field(&mut output, field_number, &nested)?;
        }
    }
    if matched != 1 {
        return Err(io::Error::other("wire path not found").into());
    }
    Ok(output)
}

fn with_noncanonical_length_path(source: &[u8], id: u64, path: &[u32]) -> R<Vec<u8>> {
    let payload = movie_payload_at(source, id)?;
    with_movie(source, id, rewrite_noncanonical_length(&payload, path)?)
}

fn with_movie_payload_reference(
    source: &[u8],
    id: u64,
    field_number: u32,
    identifier: u64,
) -> R<Vec<u8>> {
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(source, id)?.as_slice())?;
    let reference = Some(r(identifier));
    match field_number {
        10 => movie.super_.title = reference,
        11 => movie.super_.caption = reference,
        19 => movie.style = reference,
        _ => return Err(io::Error::other("unsupported movie reference field").into()),
    }
    with_movie(source, id, movie.encode_to_vec())
}

fn with_field_info_path(
    source: &[u8],
    id: u64,
    path: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .field_infos
        .push(FieldInfo {
            path: FieldPath::new(path.to_vec()),
            object_references: object_references.to_vec(),
            data_references: data_references.to_vec(),
            ..FieldInfo::default()
        });
    replace_document(source, archive)
}

fn with_invalid_slide_drawable_role(source: &[u8]) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let slide = archive
        .object_mut(SLIDE)
        .ok_or_else(|| io::Error::other("slide"))?;
    let message = slide
        .messages
        .iter_mut()
        .find(|message| message.type_ == SLIDE_TYPE)
        .ok_or_else(|| io::Error::other("slide message"))?;
    let mut payload = kn::SlideArchive::decode(message.data.as_slice())?;
    payload.owned_drawables.push(r(SLIDE_NODE));
    payload.drawables_z_order.push(r(SLIDE_NODE));
    message.data = payload.encode_to_vec();
    slide.archive_info.message_infos[0]
        .object_references
        .push(SLIDE_NODE);
    replace_document(source, archive)
}

fn with_z_order_only_reference(source: &[u8], target: u64) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let slide = archive
        .object_mut(SLIDE)
        .ok_or_else(|| io::Error::other("slide"))?;
    let message = slide
        .messages
        .iter_mut()
        .find(|message| message.type_ == SLIDE_TYPE)
        .ok_or_else(|| io::Error::other("slide message"))?;
    let mut payload = kn::SlideArchive::decode(message.data.as_slice())?;
    payload.drawables_z_order.push(r(target));
    message.data = payload.encode_to_vec();
    slide.archive_info.message_infos[0]
        .object_references
        .push(target);
    replace_document(source, archive)
}

fn with_refs(source: &[u8], id: u64, refs: Vec<u64>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .object_references = refs;
    replace_document(source, archive)
}
fn with_field_info(source: &[u8], id: u64, refs: Vec<u64>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .field_infos
        .push(FieldInfo {
            path: FieldPath::new(vec![99, 1]),
            object_references: refs,
            ..FieldInfo::default()
        });
    replace_document(source, archive)
}
fn without_entry(source: &[u8], name: &str) -> R<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .filter(|e| e.name() != name)
            .map(|e| (e.name(), e.data()))
            .collect::<Vec<_>>(),
        Limits::default(),
    )?)
}
fn key(payload: &mut Vec<u8>, field: u32, wire: u8) {
    let mut value = (u64::from(field) << 3) | u64::from(wire);
    while value >= 0x80 {
        payload.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    payload.push(value as u8);
}
fn unknowns(payload: &mut Vec<u8>) -> R<()> {
    append_length_delimited_field(payload, UNKNOWN, UNKNOWN_BYTES)?;
    payload.extend_from_slice(&[0xd0, 0x05, 0x96, 0x81, 0x00]);
    key(payload, 4_092, 3);
    key(payload, 1, 2);
    payload.push(GROUP_BYTES.len() as u8);
    payload.extend_from_slice(GROUP_BYTES);
    key(payload, 4_092, 4);
    Ok(())
}
fn reject(source: &[u8]) -> R<()> {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => {
            assert!(package.slide_movie_geometry(0usize, 0usize).is_err());
            Ok(())
        },
    }
}

fn reject_transform(source: &[u8], label: &str) -> R<()> {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    assert!(
        package.slide_movie_transform(0usize, 0usize).is_err(),
        "accepted hostile movie-transform source: {label}"
    );
    Ok(())
}

#[test]
fn reads_file_movies_with_audio_sibling_and_media_assets() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let expected = g(100.0, 200.0, 800.0, 300.0)?;
    assert_eq!(
        package.slide_movie_geometry(0usize, 0usize)?,
        Some(expected)
    );
    assert!(package.slide_movie_geometry(0usize, 1usize).is_err());
    assert!(package.edit_slide_movie_geometry(0usize, 1usize).is_err());
    assert_eq!(
        package.slide_movie_geometry(0usize, 2usize)?,
        Some(expected)
    );
    assert_eq!(
        movie_payload_at(&source, MOVIES[1])?,
        movie_payload(1, true)
    );
    assert_eq!(
        Catalog::from_bytes(&source)?
            .iter()
            .find(|e| e.name() == "Data/movie.mov")
            .map(|e| e.data()),
        Some(b"synthetic movie bytes".as_slice())
    );
    Ok(())
}

#[test]
fn accepts_native_style_role_and_rejects_unknown_or_duplicate_styles() -> R<()> {
    let source = source()?;
    let native_style = with_style_message_type(&source, STYLE[0], 3_016)?;
    assert_eq!(
        Package::from_bytes(&native_style)?.slide_movie_geometry(0usize, 0usize)?,
        Some(g(100.0, 200.0, 800.0, 300.0)?)
    );

    reject(&with_style_message_type(&source, STYLE[0], UNKNOWN)?)?;
    reject(&with_duplicate_style_message(&source, STYLE[0], 2_025)?)?;
    Ok(())
}

#[test]
fn accepts_order_independent_unique_archive_info_references() -> R<()> {
    let source = source()?;
    let object_reordered = with_refs(&source, MOVIES[0], vec![STYLE[0], TITLE[0], CAPTION[0]])?;
    let data_reordered =
        with_movie_archive_data_refs(&object_reordered, MOVIES[0], vec![2_001, 2_002])?;
    assert_eq!(
        Package::from_bytes(&data_reordered)?.slide_movie_geometry(0usize, 0usize)?,
        Some(g(100.0, 200.0, 800.0, 300.0)?)
    );
    Ok(())
}

#[test]
fn no_op_is_exact_and_debug_redacts_geometry() -> R<()> {
    let source = source()?;
    let before = g(100.0, 200.0, 800.0, 300.0)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(before)?
        .commit()?;
    assert_eq!(bytes(commit.package())?, source);
    assert!(!commit.diagnostics().changed());
    assert!(!format!("{:?}", commit.patch()).contains("800.0"));
    Ok(())
}

#[test]
fn replacement_invalidates_previews_and_preserves_unselected_graph() -> R<()> {
    let source = source()?;
    let sibling = movie_payload_at(&source, MOVIES[1])?;
    let replacement = g(240.5, 315.25, 1_280.0, 720.0)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(replacement)?
        .commit()?;
    assert_eq!(
        commit.package().slide_movie_geometry(0usize, 0usize)?,
        Some(replacement)
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().deleted_previews() > 0);
    let candidate = bytes(commit.package())?;
    assert_eq!(movie_payload_at(&candidate, MOVIES[1])?, sibling);
    let catalog = Catalog::from_bytes(&candidate)?;
    assert!(catalog.iter().all(|e| !e.name().starts_with("preview")));
    assert_eq!(
        catalog
            .iter()
            .find(|e| e.name() == "Data/poster.png")
            .map(|e| e.data()),
        Some(b"synthetic poster bytes".as_slice())
    );
    Ok(())
}

#[test]
fn inverse_apply_and_stale_patch_conflict_are_exact() -> R<()> {
    let source = source()?;
    let patch = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(12.0, 34.0, 400.0, 250.0)?)?
        .commit()?;
    let restored = patch
        .package()
        .apply_slide_movie_geometry(&patch.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    let applied = Package::from_bytes(&source)?.apply_slide_movie_geometry(patch.patch())?;
    assert_eq!(bytes(applied.package())?, bytes(patch.package())?);
    let other = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(22.0, 44.0, 500.0, 260.0)?)?
        .commit()?;
    assert!(
        other
            .package()
            .apply_slide_movie_geometry(patch.patch())
            .is_err()
    );
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_partial_and_nonfinite_geometry_reject() -> R<()> {
    let source = source()?;
    let original = movie_payload_at(&source, MOVIES[0])?;
    let nested = tsd::GeometryArchive {
        position: Some(tsp::Point { x: 1.0, y: 2.0 }),
        size: Some(tsp::Size {
            width: 3.0,
            height: 4.0,
        }),
        ..Default::default()
    }
    .encode_to_vec();
    let mut duplicate = original.clone();
    append_length_delimited_field(&mut duplicate, 1, &nested)?;
    reject(&with_movie(&source, MOVIES[0], duplicate)?)?;
    let mut wrong = original.clone();
    append_varint_field(&mut wrong, 1, 7)?;
    reject(&with_movie(&source, MOVIES[0], wrong)?)?;
    let mut partial = tsd::MovieArchive::decode(original.as_slice())?;
    partial.super_.geometry.as_mut().unwrap().size = None;
    reject(&with_movie(&source, MOVIES[0], partial.encode_to_vec())?)?;
    let mut missing = tsd::MovieArchive::decode(original.as_slice())?;
    missing.super_.geometry = None;
    reject(&with_movie(&source, MOVIES[0], missing.encode_to_vec())?)?;
    let mut nonfinite = tsd::MovieArchive::decode(original.as_slice())?;
    nonfinite
        .super_
        .geometry
        .as_mut()
        .unwrap()
        .position
        .as_mut()
        .unwrap()
        .x = f32::NAN;
    reject(&with_movie(&source, MOVIES[0], nonfinite.encode_to_vec())?)?;
    Ok(())
}

#[test]
fn constructor_rejects_nonfinite_and_nonpositive_values() -> R<()> {
    for (p, s) in [
        (
            Point {
                x: f32::NAN,
                y: 0.0,
            },
            Size {
                width: 1.0,
                height: 1.0,
            },
        ),
        (
            Point {
                x: 0.0,
                y: f32::INFINITY,
            },
            Size {
                width: 1.0,
                height: 1.0,
            },
        ),
        (
            Point { x: 0.0, y: 0.0 },
            Size {
                width: 0.0,
                height: 1.0,
            },
        ),
        (
            Point { x: 0.0, y: 0.0 },
            Size {
                width: 1.0,
                height: -1.0,
            },
        ),
    ] {
        assert!(MovieGeometry::new(p, s).is_err());
    }
    Ok(())
}

#[test]
fn unknown_overlong_scalar_and_balanced_group_are_retained_when_ingress_admits() -> R<()> {
    let source = source()?;
    let mut payload = movie_payload_at(&source, MOVIES[0])?;
    unknowns(&mut payload)?;
    let Ok(hostile) = with_movie(&source, MOVIES[0], payload) else {
        return Ok(());
    };
    let Ok(package) = Package::from_bytes(&hostile) else {
        return Ok(());
    };
    let Ok(edit) = package.edit_slide_movie_geometry(0usize, 0usize) else {
        return Ok(());
    };
    let Ok(commit) = edit.set(g(1.0, 2.0, 3.0, 4.0)?)?.commit() else {
        return Ok(());
    };
    let Ok(candidate) = bytes(commit.package()) else {
        return Ok(());
    };
    let Ok(rewritten) = movie_payload_at(&candidate, MOVIES[0]) else {
        return Ok(());
    };
    assert!(
        rewritten
            .windows(UNKNOWN_BYTES.len())
            .any(|w| w == UNKNOWN_BYTES)
    );
    assert!(
        rewritten
            .windows(GROUP_BYTES.len())
            .any(|w| w == GROUP_BYTES)
    );
    assert!(
        rewritten
            .windows(5)
            .any(|w| w == [0xd0, 0x05, 0x96, 0x81, 0x00])
    );
    Ok(())
}

#[test]
fn poster_only_missing_data_and_non_file_sibling_are_refused() -> R<()> {
    let source = source()?;
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(&source, MOVIES[0])?.as_slice())?;
    movie.movie_data = None;
    reject(&with_movie(&source, MOVIES[0], movie.encode_to_vec())?)?;
    reject(&without_entry(&source, "Data/movie.mov")?)?;
    assert!(
        Package::from_bytes(&source)?
            .slide_movie_geometry(0usize, 1usize)
            .is_err()
    );
    Ok(())
}

#[test]
fn archive_info_refs_and_duplicate_physical_aliases_fail_atomically() -> R<()> {
    let source = source()?;
    for refs in [
        vec![],
        vec![TITLE[0]],
        vec![TITLE[0], TITLE[0], CAPTION[0]],
        vec![TITLE[0], CAPTION[0], 9_999],
    ] {
        reject(&with_refs(&source, MOVIES[0], refs)?)?;
    }
    reject(&with_field_info(&source, MOVIES[0], vec![CAPTION[0]])?)?;
    reject(&with_field_info(
        &source,
        MOVIES[0],
        vec![CAPTION[0], CAPTION[0]],
    )?)?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let movie = archive
        .object(MOVIES[0])
        .ok_or_else(|| io::Error::other("movie"))?
        .clone();
    archive.objects.push(movie);
    match replace_document(&source, archive) {
        Err(_) => {},
        Ok(value) => reject(&value)?,
    }
    reject(&with_alias_member(&source, MOVIES[0])?)?;
    Ok(())
}

#[test]
fn selectors_and_input_limit_fail_before_publication() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    assert!(package.slide_movie_geometry(9usize, 0usize).is_err());
    assert!(package.slide_movie_geometry(0usize, 9usize).is_err());
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
            ReadOptions::new(tight, litchi_keynote::SemanticLimits::default())
        )
        .is_err()
    );
    let before = bytes(&package)?;
    let result = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(1.0, 2.0, 3.0, 4.0)?)?
        .commit();
    if result.is_err() {
        assert_eq!(bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn parent_title_caption_playback_and_media_edges_are_unchanged() -> R<()> {
    let source = source()?;
    let before = tsd::MovieArchive::decode(movie_payload_at(&source, MOVIES[0])?.as_slice())?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(40.0, 50.0, 600.0, 320.0)?)?
        .commit()?;
    let after = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(commit.package())?, MOVIES[0])?.as_slice(),
    )?;
    assert_eq!(after.super_.parent, before.super_.parent);
    assert_eq!(after.super_.title, before.super_.title);
    assert_eq!(after.super_.caption, before.super_.caption);
    assert_eq!(after.movie_data, before.movie_data);
    assert_eq!(after.poster_image_data, before.poster_image_data);
    assert_eq!(after.start_time, before.start_time);
    assert_eq!(after.end_time, before.end_time);
    assert_eq!(
        after
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.flags),
        before
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.flags)
    );
    assert_eq!(
        after
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.angle),
        before
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.angle)
    );
    Ok(())
}

#[test]
fn transform_reads_defaults_and_flip_axes() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let before = MovieTransform::new(17.5, false)?;
    assert_eq!(package.slide_movie_transform(0usize, 0usize)?, Some(before));
    assert!(!before.is_reflected());
    assert!(!before.reflected());
    assert_eq!(
        before.flipped(MovieFlipAxis::Horizontal),
        MovieTransform::new(17.5, true)?
    );
    assert_eq!(
        before.flipped(MovieFlipAxis::Vertical),
        MovieTransform::new(197.5, true)?
    );
    assert_eq!(
        MovieTransform::new(180.0, true)?.flipped(MovieFlipAxis::Vertical),
        MovieTransform::identity()
    );
    Ok(())
}

#[test]
fn absent_flags_read_unreflected_but_horizontal_flip_is_rejected_atomically() -> R<()> {
    let source = with_transform_fields(&source()?, MOVIES[0], None, Some(17.5))?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_movie_transform(0usize, 0usize)?,
        Some(MovieTransform::new(17.5, false)?)
    );
    let before = bytes(&package)?;
    assert!(
        package
            .edit_slide_movie_geometry(0usize, 0usize)?
            .flip(MovieFlipAxis::Horizontal)
            .is_err()
    );
    assert_eq!(bytes(&package)?, before);
    Ok(())
}

#[test]
fn absent_flags_and_angle_read_identity_and_noop_preserves_presence() -> R<()> {
    let source = with_transform_fields(&source()?, MOVIES[0], None, None)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_movie_transform(0usize, 0usize)?,
        Some(MovieTransform::identity())
    );

    let no_op = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set_transform(MovieTransform::identity())?
        .commit()?;
    assert_eq!(bytes(no_op.package())?, source);
    let archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(no_op.package())?, MOVIES[0])?.as_slice(),
    )?;
    let geometry = archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    assert_eq!(geometry.flags, None);
    assert_eq!(geometry.angle, None);
    Ok(())
}

#[test]
fn absent_angle_horizontal_preserves_absence_and_vertical_inserts_half_turn() -> R<()> {
    let source = with_transform_fields(&source()?, MOVIES[0], Some(0x20), None)?;
    let package = Package::from_bytes(&source)?;

    let horizontal = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .flip(MovieFlipAxis::Horizontal)?
        .commit()?;
    let horizontal_archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(horizontal.package())?, MOVIES[0])?.as_slice(),
    )?;
    let horizontal_geometry = horizontal_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    assert_eq!(horizontal_geometry.flags, Some(0x24));
    assert_eq!(horizontal_geometry.angle, None);
    assert_eq!(
        horizontal.package().slide_movie_transform(0usize, 0usize)?,
        Some(MovieTransform::new(0.0, true)?)
    );

    let vertical = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .flip(MovieFlipAxis::Vertical)?
        .commit()?;
    let vertical_archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(vertical.package())?, MOVIES[0])?.as_slice(),
    )?;
    let vertical_geometry = vertical_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    assert_eq!(vertical_geometry.flags, Some(0x24));
    assert_eq!(vertical_geometry.angle, Some(180.0));
    assert_eq!(
        vertical.package().slide_movie_transform(0usize, 0usize)?,
        Some(MovieTransform::new(180.0, true)?)
    );
    Ok(())
}

#[test]
fn explicit_zero_presence_survives_noop_and_geometry_only_change() -> R<()> {
    let source = with_transform_fields(&source()?, MOVIES[0], Some(0), Some(0.0))?;
    let package = Package::from_bytes(&source)?;
    let identity = MovieTransform::identity();
    assert_eq!(
        package.slide_movie_transform(0usize, 0usize)?,
        Some(identity)
    );
    let no_op = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set_transform(identity)?
        .commit()?;
    assert_eq!(bytes(no_op.package())?, source);
    assert!(no_op.patch().is_noop());
    let no_op_archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(no_op.package())?, MOVIES[0])?.as_slice(),
    )?;
    let no_op_geometry = no_op_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    assert_eq!(no_op_geometry.flags, Some(0));
    assert_eq!(no_op_geometry.angle, Some(0.0));

    let geometry_only = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(12.0, 34.0, 400.0, 250.0)?)?
        .commit()?;
    let geometry_only_archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(geometry_only.package())?, MOVIES[0])?.as_slice(),
    )?;
    let geometry_only_geometry = geometry_only_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    assert_eq!(geometry_only_geometry.flags, Some(0));
    assert_eq!(geometry_only_geometry.angle, Some(0.0));
    Ok(())
}

#[test]
fn unknown_flag_bits_and_top_level_flags_survive_reflection_toggle() -> R<()> {
    let source = with_transform_fields(&source()?, MOVIES[0], Some(0x8000_0024), Some(17.5))?;
    let before_archive =
        tsd::MovieArchive::decode(movie_payload_at(&source, MOVIES[0])?.as_slice())?;
    let before_geometry = before_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    let before_flags = before_geometry
        .flags
        .ok_or_else(|| io::Error::other("flags"))?;
    let before_top_level_flags = before_archive.flags;

    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .flip(MovieFlipAxis::Horizontal)?
        .commit()?;
    let after_archive = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(commit.package())?, MOVIES[0])?.as_slice(),
    )?;
    let after_geometry = after_archive
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("geometry"))?;
    let after_flags = after_geometry
        .flags
        .ok_or_else(|| io::Error::other("flags"))?;
    assert_eq!(after_flags & !0x04, before_flags & !0x04);
    assert_eq!(after_flags & 0x04, 0);
    assert_eq!(after_archive.flags, before_top_level_flags);
    Ok(())
}

#[test]
fn combined_geometry_and_transform_round_trip_inverse_and_conflict() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let replacement = g(240.5, 315.25, 1_280.0, 720.0)?;
    let transform = MovieTransform::new(270.0, true)?;
    let commit = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(replacement)?
        .set_transform(transform)?
        .commit()?;
    assert_eq!(commit.patch().before(), g(100.0, 200.0, 800.0, 300.0)?);
    assert_eq!(commit.patch().after(), replacement);
    assert_eq!(
        commit.patch().before_transform(),
        MovieTransform::new(17.5, false)?
    );
    assert_eq!(commit.patch().after_transform(), transform);
    assert_eq!(
        commit.package().slide_movie_geometry(0usize, 0usize)?,
        Some(replacement)
    );
    assert_eq!(
        commit.package().slide_movie_transform(0usize, 0usize)?,
        Some(transform)
    );

    let candidate = bytes(commit.package())?;
    let inverse = commit
        .package()
        .apply_slide_movie_geometry(&commit.patch().inverse())?;
    assert_eq!(bytes(inverse.package())?, source);
    let applied = Package::from_bytes(&source)?.apply_slide_movie_geometry(commit.patch())?;
    assert_eq!(bytes(applied.package())?, candidate);

    let stale = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(22.0, 44.0, 500.0, 260.0)?)?
        .set_transform(MovieTransform::new(45.0, true)?)?
        .commit()?;
    assert!(
        stale
            .package()
            .apply_slide_movie_geometry(commit.patch())
            .is_err()
    );
    Ok(())
}

#[test]
fn restore_original_size_preserves_staged_position_transform_and_inverse() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let original_size = Size {
        width: 800.0,
        height: 300.0,
    };
    let staged_position = Point {
        x: 240.5,
        y: 315.25,
    };
    let staged_transform = MovieTransform::new(17.5, true)?;
    let edit = package.edit_slide_movie_geometry(0usize, 0usize)?;
    assert_eq!(edit.original_size(), Some(original_size));

    let edit = edit
        .set(MovieGeometry::new(
            staged_position,
            Size {
                width: 1_280.0,
                height: 720.0,
            },
        )?)?
        .flip(MovieFlipAxis::Horizontal)?
        .restore_original_size()?;
    assert_eq!(edit.after().position(), staged_position);
    assert_eq!(edit.after().size(), original_size);
    assert_eq!(edit.after_transform(), staged_transform);

    let commit = edit.commit()?;
    assert_eq!(
        commit.package().slide_movie_geometry(0usize, 0usize)?,
        Some(MovieGeometry::new(staged_position, original_size)?)
    );
    assert_eq!(
        commit.package().slide_movie_transform(0usize, 0usize)?,
        Some(staged_transform)
    );
    let inverse = commit
        .package()
        .apply_slide_movie_geometry(&commit.patch().inverse())?;
    assert_eq!(bytes(inverse.package())?, source);
    Ok(())
}

#[test]
fn restore_original_size_rejects_missing_and_invalid_metadata_atomically() -> R<()> {
    let source = source()?;
    let missing = with_original_size(&source, MOVIES[0], None)?;
    let missing_package = Package::from_bytes(&missing)?;
    let missing_before = bytes(&missing_package)?;
    let missing_error = missing_package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .restore_original_size()
        .unwrap_err();
    assert_eq!(
        missing_error,
        SlideMovieGeometryError::UnsupportedDependency
    );
    assert_eq!(bytes(&missing_package)?, missing_before);

    let invalid = with_original_size(
        &source,
        MOVIES[0],
        Some(tsp::Size {
            width: -1.0,
            height: 300.0,
        }),
    )?;
    let invalid_package = Package::from_bytes(&invalid)?;
    let invalid_before = bytes(&invalid_package)?;
    let invalid_error = invalid_package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .restore_original_size()
        .unwrap_err();
    assert_eq!(invalid_error, SlideMovieGeometryError::InvalidSource);
    assert_eq!(bytes(&invalid_package)?, invalid_before);
    Ok(())
}

#[test]
fn locked_movie_allows_exact_transform_noop_but_rejects_changes() -> R<()> {
    let source = with_locked_movie(&source()?, MOVIES[0])?;
    let package = Package::from_bytes(&source)?;
    let before = MovieTransform::new(17.5, false)?;
    let no_op = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set_transform(before)?
        .commit()?;
    assert_eq!(bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());

    let before_bytes = bytes(&package)?;
    assert!(
        package
            .edit_slide_movie_geometry(0usize, 0usize)?
            .flip(MovieFlipAxis::Horizontal)
            .and_then(|edit| edit.commit())
            .is_err()
    );
    assert!(
        package
            .edit_slide_movie_geometry(0usize, 0usize)?
            .set(g(1.0, 2.0, 3.0, 4.0)?)?
            .commit()
            .is_err()
    );
    assert_eq!(bytes(&package)?, before_bytes);
    Ok(())
}

#[test]
fn malformed_transform_duplicate_wrong_wire_and_nonfinite_angle_reject() -> R<()> {
    let source = source()?;
    let original = movie_payload_at(&source, MOVIES[0])?;

    let mut duplicate = Vec::new();
    append_fixed32(&mut duplicate, 4, 9.0);
    reject(&with_geometry_extra(&source, MOVIES[0], &duplicate)?)?;

    let mut wrong_wire = Vec::new();
    append_varint_field(&mut wrong_wire, 4, 9)?;
    reject(&with_geometry_extra(&source, MOVIES[0], &wrong_wire)?)?;

    let mut nonfinite = tsd::MovieArchive::decode(original.as_slice())?;
    nonfinite
        .super_
        .geometry
        .as_mut()
        .ok_or_else(|| io::Error::other("geometry"))?
        .angle = Some(f32::NAN);
    reject(&with_movie(&source, MOVIES[0], nonfinite.encode_to_vec())?)?;
    Ok(())
}

#[test]
fn hostile_movie_roles_and_global_inbound_routes_fail_closed() -> R<()> {
    let source = source()?;
    let cases = [
        (
            "movie carries slide role alias",
            with_message_alias(&source, MOVIES[0], 5)?,
        ),
        (
            "slide carries movie role alias",
            with_message_alias(&source, SLIDE, MOVIE_TYPE)?,
        ),
        (
            "foreign object inbound",
            with_foreign_inbound(&source, MOVIES[0], ForeignInbound::Object)?,
        ),
        (
            "foreign data inbound",
            with_foreign_inbound(&source, 2_002, ForeignInbound::Data)?,
        ),
        (
            "foreign data inbound to movie",
            with_foreign_inbound(&source, MOVIES[0], ForeignInbound::Data)?,
        ),
        (
            "foreign field-info inbound",
            with_foreign_inbound(&source, MOVIES[0], ForeignInbound::Field)?,
        ),
        (
            "foreign field-info data inbound",
            with_foreign_inbound(&source, 2_002, ForeignInbound::FieldData)?,
        ),
        (
            "foreign slide z-order object",
            with_foreign_z_order(&source)?,
        ),
    ];
    for (label, hostile) in cases {
        reject_transform(&hostile, label)?;
    }
    Ok(())
}

#[test]
fn hostile_metadata_identity_and_authority_routes_fail_closed() -> R<()> {
    let source = source()?;
    let cases = [
        (
            "missing current movie uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != MOVIES[0]);
            })?,
        ),
        (
            "movie uuid in wrong component",
            with_metadata(&source, |metadata| {
                let entry = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .iter()
                    .find(|entry| entry.identifier == MOVIES[0])
                    .cloned()
                    .expect("movie uuid");
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .retain(|candidate| candidate.identifier != MOVIES[0]);
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 2)
                    .expect("unrelated component")
                    .object_uuid_map_entries
                    .push(entry);
            })?,
        ),
        (
            "duplicate current movie uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .push(uuid(MOVIES[0]));
            })?,
        ),
        (
            "movie uuid collides with versioned registry",
            with_metadata(&source, |metadata| {
                metadata.versioned_components.push(tsp::ComponentInfo {
                    identifier: 1,
                    preferred_locator: "Document".into(),
                    locator: Some("Document".into()),
                    object_uuid_map_entries: vec![uuid(MOVIES[0])],
                    ..Default::default()
                });
            })?,
        ),
        (
            "ambiguous movie uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .ambiguous_object_identifiers
                    .push(MOVIES[0]);
            })?,
        ),
        (
            "movie used as metadata-map owner",
            with_metadata(&source, |metadata| {
                metadata.data_metadata_map = Some(r(MOVIES[0]));
            })?,
        ),
        (
            "movie used as data-reference owner",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .data_references
                    .push(tsp::ComponentDataReference {
                        data_identifier: 2_002,
                        object_reference_list: vec![
                            tsp::component_data_reference::ObjectReference {
                                object_identifier: MOVIES[0],
                                count: 1,
                            },
                        ],
                    });
            })?,
        ),
        (
            "missing current slide uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != SLIDE);
            })?,
        ),
        (
            "slide uuid in wrong component",
            with_metadata(&source, |metadata| {
                let entry = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .iter()
                    .find(|entry| entry.identifier == SLIDE)
                    .cloned()
                    .expect("slide uuid");
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .retain(|candidate| candidate.identifier != SLIDE);
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 2)
                    .expect("unrelated component")
                    .object_uuid_map_entries
                    .push(entry);
            })?,
        ),
        (
            "duplicate current slide uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .object_uuid_map_entries
                    .push(uuid(SLIDE));
            })?,
        ),
        (
            "slide uuid collides with versioned registry",
            with_metadata(&source, |metadata| {
                metadata.versioned_components.push(tsp::ComponentInfo {
                    identifier: 1,
                    preferred_locator: "Document".into(),
                    locator: Some("Document".into()),
                    object_uuid_map_entries: vec![uuid(SLIDE)],
                    ..Default::default()
                });
            })?,
        ),
        (
            "ambiguous slide uuid",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .ambiguous_object_identifiers
                    .push(SLIDE);
            })?,
        ),
        (
            "slide used as metadata-map owner",
            with_metadata(&source, |metadata| {
                metadata.data_metadata_map = Some(r(SLIDE));
            })?,
        ),
        (
            "slide used as data-reference owner",
            with_metadata(&source, |metadata| {
                metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 1)
                    .expect("document component")
                    .data_references
                    .push(tsp::ComponentDataReference {
                        data_identifier: 2_002,
                        object_reference_list: vec![
                            tsp::component_data_reference::ObjectReference {
                                object_identifier: SLIDE,
                                count: 1,
                            },
                        ],
                    });
            })?,
        ),
    ];
    for (label, hostile) in cases {
        reject_transform(&hostile, label)?;
    }
    Ok(())
}

#[test]
fn hostile_movie_data_and_noncanonical_reference_routes_fail_closed() -> R<()> {
    let source = source()?;
    let cases = [
        (
            "dangling movie data",
            with_movie_data_id(&source, MOVIES[0], 9_999)?,
        ),
        (
            "poster used as movie data",
            with_movie_data_id(&source, MOVIES[0], 2_001)?,
        ),
        (
            "duplicate archive data reference",
            with_movie_archive_data_refs(&source, MOVIES[0], vec![2_002, 2_002])?,
        ),
        (
            "duplicate movie data field",
            with_duplicate_movie_data(&source, MOVIES[0])?,
        ),
        (
            "wrong archive data reference",
            with_movie_archive_data_refs(&source, MOVIES[0], vec![2_001])?,
        ),
        (
            "unexpected archive data reference",
            with_movie_archive_data_refs(&source, MOVIES[0], vec![2_002, 9_999])?,
        ),
        (
            "noncanonical movie super key",
            with_noncanonical_super_key(&source, MOVIES[0])?,
        ),
        (
            "noncanonical movie parent key",
            with_noncanonical_parent_key(&source, MOVIES[0])?,
        ),
    ];
    for (label, hostile) in cases {
        reject_transform(&hostile, label)?;
    }
    Ok(())
}

#[test]
fn noncanonical_length_prefixes_on_movie_reference_and_size_routes_fail_closed() -> R<()> {
    let source = source()?;
    for (path, label) in [
        (&[1, 2][..], "movie parent reference length"),
        (&[1, 10][..], "movie title reference length"),
        (&[1, 11][..], "movie caption reference length"),
        (&[19][..], "movie style reference length"),
        (&[20][..], "movie original-size length"),
        (&[21][..], "movie natural-size length"),
    ] {
        reject_transform(
            &with_noncanonical_length_path(&source, MOVIES[0], path)?,
            label,
        )?;
    }
    Ok(())
}

#[test]
fn movie_payload_and_archive_info_routes_must_match_exactly() -> R<()> {
    let source = source()?;
    reject_transform(
        &with_movie_payload_reference(&source, MOVIES[0], 10, TITLE[1])?,
        "movie payload title differs from ArchiveInfo",
    )?;
    reject_transform(
        &with_refs(
            &source,
            MOVIES[0],
            vec![TITLE[0], CAPTION[0], STYLE[0], AUDIO],
        )?,
        "movie ArchiveInfo has an extra object reference",
    )?;
    reject_transform(
        &with_field_info_path(&source, MOVIES[0], &[10], &[TITLE[0]], &[])?,
        "movie ArchiveInfo has a canonical field route",
    )?;
    let duplicate = with_field_info_path(&source, MOVIES[0], &[10], &[TITLE[0]], &[])?;
    reject_transform(
        &with_field_info_path(&duplicate, MOVIES[0], &[10], &[TITLE[0]], &[])?,
        "movie ArchiveInfo has duplicate canonical field routes",
    )?;
    reject_transform(
        &with_field_info_path(&source, MOVIES[0], &[10, 1], &[TITLE[0]], &[])?,
        "movie ArchiveInfo has a noncanonical field route",
    )?;
    Ok(())
}

#[test]
fn invalid_owned_and_z_order_drawable_roles_fail_closed() -> R<()> {
    reject_transform(
        &with_invalid_slide_drawable_role(&source()?)?,
        "slide owns a non-drawable known-role object",
    )?;
    Ok(())
}

#[test]
fn z_order_only_dangling_and_wrong_role_ids_fail_closed() -> R<()> {
    let source = source()?;
    for (target, label) in [
        (9_999, "z-order-only dangling drawable"),
        (TITLE[0], "z-order-only title object"),
        (STYLE[0], "z-order-only style object"),
        (SLIDE, "z-order-only slide object"),
        (SLIDE_NODE, "z-order-only slide-node object"),
    ] {
        reject_transform(&with_z_order_only_reference(&source, target)?, label)?;
    }
    Ok(())
}

#[test]
fn every_known_movie_role_alias_fails_closed() -> R<()> {
    let source = source()?;
    for (role, label) in [
        (SLIDE_TYPE, "slide role alias"),
        (MOVIE_TYPE, "movie role duplicate"),
        (TABLE_INFO_TYPE, "table-info role alias"),
        (TABLE_MODEL_TYPE, "table-model role alias"),
        (HEADER_BUCKET_TYPE, "header-bucket role alias"),
        (TABLE_STYLE_TYPE, "table-style role alias"),
        (TABLE_STYLE_PRESET_TYPE, "table-style-preset role alias"),
        (TABLE_STYLE_NETWORK_TYPE, "table-style-network role alias"),
        (STYLESHEET_TYPE, "stylesheet role alias"),
    ] {
        reject_transform(&with_message_alias(&source, MOVIES[0], role)?, label)?;
    }
    Ok(())
}
