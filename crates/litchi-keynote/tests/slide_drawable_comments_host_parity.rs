//! Host-parity regressions for the focused Keynote direct-drawable comment
//! owner.
//!
//! These cases deliberately build the small graph shapes which the old
//! `KeynoteEditor` comment methods accepted but which a native fixture cannot
//! express on its own: two drawables sharing one root, two roots sharing one
//! reply, and a package with no annotation-author storage.  The physical
//! oracle is shared with the cross-component parity tests so the public
//! assertions remain selector and value based.

use std::{
    collections::{BTreeMap, BTreeSet},
    io,
};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, patch_length_delimited_field,
    transform_length_delimited_field,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsd, tsp};
use litchi_keynote::{DrawableKind, DrawableSelector, Package, SlideSelector};
use prost::Message as _;

#[path = "support/drawable_comment_fixtures.rs"]
mod cross_component_fixtures;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/drawable-comments-source-native.key");
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

const SUPPORTED_ROUTE_VARIANTS: &[(u32, &[u32])] = &[
    (3_002, &[6]),
    (3_004, &[1, 6]),
    (3_005, &[1, 6]),
    (3_006, &[1, 6]),
    (3_007, &[1, 6]),
    (3_008, &[1, 6]),
    (5_021, &[1, 6]),
    (6_000, &[1, 6]),
    (3_009, &[1, 1, 6]),
    (2_011, &[1, 1, 6]),
    (6_007, &[1, 1, 6]),
    (2_014, &[1, 1, 1, 6]),
    (7, &[1, 1, 1, 6]),
    (12, &[1, 1, 1, 6]),
];

#[derive(Debug, Clone)]
struct DrawableEntry {
    identifier: u64,
    message_type: u32,
    message_index: usize,
    comment_path: &'static [u32],
}

#[derive(Debug, Clone)]
struct FixtureProfile {
    drawables: Vec<DrawableEntry>,
    commented_position: usize,
    empty_positions: Vec<usize>,
}

#[derive(Debug, Clone, PartialEq)]
struct CommentSignature {
    text: Option<String>,
    date_bits: Option<u64>,
    author: Option<u64>,
    uuid: Option<(u64, u64)>,
    replies: Vec<u64>,
    payload: Vec<u8>,
}

#[derive(Debug, Clone)]
struct SharedGraphFixture {
    bytes: Vec<u8>,
    original_position: usize,
    sibling_position: usize,
    root_identifier: u64,
    reply_identifier: u64,
    second_root_identifier: Option<u64>,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn route_path(message_type: u32) -> Option<&'static [u32]> {
    Some(match message_type {
        3_002 => &[6],
        3_004..=3_008 | 5_021 | 6_000 => &[1, 6],
        3_009 | 2_011 | 6_007 => &[1, 1, 6],
        2_014 | 7 | 12 => &[1, 1, 1, 6],
        _ => return None,
    })
}

fn component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        archives.push((entry.name().to_owned(), archive));
    }
    Ok(archives)
}

fn component_containing_object(source: &[u8], identifier: u64) -> TestResult<(String, Archive)> {
    component_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")).into())
}

fn replace_component(source: &[u8], component_name: &str, archive: Archive) -> TestResult<Vec<u8>> {
    let bytes = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&bytes)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
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

fn first_slide(source: &[u8]) -> TestResult<(String, u64, Vec<u64>)> {
    for (component_name, archive) in component_archives(source)? {
        if !component_name.starts_with("Index/Slide-") {
            continue;
        }
        for object in &archive.objects {
            let Some((_, message)) = object
                .messages
                .iter()
                .enumerate()
                .find(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
            else {
                continue;
            };
            let slide = kn::SlideArchive::decode(message.data.as_slice())?;
            if !slide.owned_drawables.is_empty() {
                return Ok((
                    component_name,
                    object.archive_info.identifier.ok_or_else(|| {
                        io::Error::other("rooted slide object has no archive identifier")
                    })?,
                    slide
                        .owned_drawables
                        .into_iter()
                        .map(|reference| reference.identifier)
                        .collect(),
                ));
            }
        }
    }
    Err(io::Error::other("native fixture has no rooted slide").into())
}

fn nested_payload(payload: &[u8], path: &[u32]) -> TestResult<Vec<u8>> {
    let mut current = payload;
    for field_number in path {
        let fields = WireView::parse(current)?
            .fields()
            .filter(|field| field.number() == *field_number)
            .collect::<Vec<_>>();
        if fields.len() != 1 || fields[0].wire_type() != 2 {
            return Err(io::Error::other("square payload has no unique nested envelope").into());
        }
        fields[0].validate_canonical_framing()?;
        current = fields[0].payload();
    }
    Ok(current.to_owned())
}

fn wrap_field_one(payload: &[u8], count: usize) -> TestResult<Vec<u8>> {
    let mut wrapped = payload.to_owned();
    for _ in 0..count {
        let mut outer = Vec::new();
        append_length_delimited_field(&mut outer, 1, &wrapped)?;
        wrapped = outer;
    }
    Ok(wrapped)
}

fn reorder_slide_owned_drawables(archive: &mut Archive, slide_identifier: u64) -> TestResult<()> {
    let slide = archive
        .object(slide_identifier)
        .ok_or_else(|| io::Error::other("rooted slide disappeared before reordering"))?;
    let message_index = slide
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("rooted slide payload disappeared"))?;
    let payload = slide.messages[message_index].data.clone();
    let fields = WireView::parse(&payload)?.fields().collect::<Vec<_>>();
    let mut owned = fields
        .iter()
        .copied()
        .filter(|field| field.number() == 7 && field.wire_type() == 2)
        .collect::<Vec<_>>();
    if owned.len() < 2 {
        return Err(io::Error::other("rooted slide has too few owned drawables").into());
    }
    let mut output = Vec::with_capacity(payload.len());
    for field in fields {
        if field.number() == 7 && field.wire_type() == 2 {
            output.extend_from_slice(owned.pop().expect("owned drawable count checked").raw());
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    let slide = archive
        .object_mut(slide_identifier)
        .ok_or_else(|| io::Error::other("rooted slide disappeared during reordering"))?;
    slide.replace_message(
        message_index,
        RawMessage {
            type_: SLIDE_MESSAGE_TYPE,
            data: output,
        },
    )?;
    Ok(())
}

fn morph_square_route(
    source: &[u8],
    message_type: u32,
    comment_path: &[u32],
    reverse_source_order: bool,
) -> TestResult<(Vec<u8>, u64, usize)> {
    let profile = fixture_profile(source)?;
    let (source_position, entry) = profile
        .drawables
        .iter()
        .enumerate()
        .find(|(_, entry)| {
            entry.message_type == 2_011
                && comment_identifier_for_entry(source, entry)
                    .ok()
                    .flatten()
                    .is_none()
        })
        .ok_or_else(|| io::Error::other("native fixture has no empty square drawable"))?;
    let (component_name, mut archive) = component_containing_object(source, entry.identifier)?;
    let object = archive
        .object(entry.identifier)
        .ok_or_else(|| io::Error::other("square drawable disappeared before morphing"))?;
    let payload = object
        .messages
        .get(entry.message_index)
        .ok_or_else(|| io::Error::other("square payload disappeared before morphing"))?
        .data
        .clone();
    let inner = nested_payload(&payload, &[1, 1])?;
    let data = wrap_field_one(&inner, comment_path.len().saturating_sub(1))?;
    let object = archive
        .object_mut(entry.identifier)
        .ok_or_else(|| io::Error::other("square drawable disappeared during morphing"))?;
    object.replace_message(
        entry.message_index,
        RawMessage {
            type_: message_type,
            data,
        },
    )?;
    let (_, slide_identifier, _) = first_slide(source)?;
    if reverse_source_order {
        reorder_slide_owned_drawables(&mut archive, slide_identifier)?;
    }
    let target_position = if reverse_source_order {
        profile.drawables.len() - source_position - 1
    } else {
        source_position
    };
    Ok((
        replace_component(source, &component_name, archive)?,
        entry.identifier,
        target_position,
    ))
}

fn expected_route_kind(message_type: u32) -> DrawableKind {
    match message_type {
        3_002 => DrawableKind::Drawable,
        3_004 => DrawableKind::Shape,
        3_005 => DrawableKind::Image,
        3_006 => DrawableKind::Mask,
        3_007 => DrawableKind::Movie,
        3_008 => DrawableKind::Group,
        3_009 => DrawableKind::ConnectionLine,
        5_021 => DrawableKind::Chart,
        6_000 => DrawableKind::Table,
        6_007 => DrawableKind::WordProcessingTable,
        2_011 | 2_014 => DrawableKind::Shape,
        7 | 12 => DrawableKind::Placeholder,
        _ => panic!("unsupported test route {message_type}"),
    }
}

fn unique_message(object: &ArchiveObject, message_type: u32) -> TestResult<(usize, &[u8])> {
    let identifier = object
        .archive_info
        .identifier
        .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
    let mut result = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if result.replace((index, message.data.as_slice())).is_some() {
            return Err(io::Error::other(format!(
                "object {identifier} repeats message type {message_type}"
            ))
            .into());
        }
    }
    result.ok_or_else(|| {
        io::Error::other(format!(
            "object {identifier} has no message type {message_type}"
        ))
        .into()
    })
}

fn comment_identifier(payload: &[u8], path: &[u32]) -> TestResult<Option<u64>> {
    let mut current = payload;
    for (depth, field_number) in path.iter().copied().enumerate() {
        let fields = WireView::parse(current)?
            .fields()
            .filter(|field| field.number() == field_number)
            .collect::<Vec<_>>();
        if fields.len() > 1 {
            return Err(io::Error::other(format!(
                "drawable route repeats field {field_number} at depth {depth}"
            ))
            .into());
        }
        let Some(field) = fields.first() else {
            return Ok(None);
        };
        if field.wire_type() != 2 {
            return Err(io::Error::other("drawable comment route is not length-delimited").into());
        }
        current = field.payload();
    }
    let fields = WireView::parse(current)?
        .fields()
        .filter(|field| field.number() == 1)
        .collect::<Vec<_>>();
    if fields.is_empty() {
        return Err(io::Error::other("comment reference has no identifier").into());
    }
    if fields.len() != 1 || fields[0].wire_type() != 0 {
        return Err(io::Error::other("comment reference is malformed").into());
    }
    let reference = tsp::Reference::decode(current)?;
    if reference.identifier == 0 {
        return Err(io::Error::other("comment reference has zero identifier").into());
    }
    Ok(Some(reference.identifier))
}

fn fixture_profile(source: &[u8]) -> TestResult<FixtureProfile> {
    let (_slide_component, _slide_identifier, identifiers) = first_slide(source)?;
    let mut drawables = Vec::new();
    for identifier in identifiers {
        let (_, archive) = component_containing_object(source, identifier)?;
        let object = archive
            .object(identifier)
            .ok_or_else(|| io::Error::other("drawable object disappeared"))?;
        let Some((message_index, message)) = object
            .messages
            .iter()
            .enumerate()
            .find(|(_, message)| route_path(message.type_).is_some())
        else {
            continue;
        };
        let comment_path = route_path(message.type_).expect("route checked above");
        let _ = comment_identifier(&message.data, comment_path)?;
        drawables.push(DrawableEntry {
            identifier,
            message_type: message.type_,
            message_index,
            comment_path,
        });
    }
    let package = Package::from_bytes(source)?;
    let summaries = package.slide_drawables(SlideSelector::index(0))?;
    if summaries.len() != drawables.len() {
        return Err(io::Error::other(format!(
            "native route inventory has {} entries but focused package has {}",
            drawables.len(),
            summaries.len()
        ))
        .into());
    }
    let commented_position = summaries
        .iter()
        .position(|summary| summary.has_comment())
        .ok_or_else(|| io::Error::other("native fixture has no commented drawable"))?;
    let empty_positions = summaries
        .iter()
        .enumerate()
        .filter_map(|(position, summary)| (!summary.has_comment()).then_some(position))
        .collect::<Vec<_>>();
    if empty_positions.len() < 2 {
        return Err(io::Error::other("native fixture has fewer than two empty drawables").into());
    }
    Ok(FixtureProfile {
        drawables,
        commented_position,
        empty_positions,
    })
}

fn append_comment_reference(payload: &[u8], path: &[u32], identifier: u64) -> TestResult<Vec<u8>> {
    let Some((&field_number, remainder)) = path.split_first() else {
        return Err(io::Error::other("empty drawable comment path").into());
    };
    if remainder.is_empty() {
        let reference = tsp::Reference {
            identifier,
            ..Default::default()
        }
        .encode_to_vec();
        let mut output = payload.to_vec();
        append_length_delimited_field(&mut output, field_number, &reference)?;
        return Ok(output);
    }
    Ok(transform_length_delimited_field(
        payload,
        field_number,
        |nested| append_comment_reference(nested, remainder, identifier),
    )?)
}

fn attach_comment(source: &[u8], entry: &DrawableEntry, comment_id: u64) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) = component_containing_object(source, entry.identifier)?;
    let object = archive
        .object_mut(entry.identifier)
        .ok_or_else(|| io::Error::other("drawable disappeared before attachment"))?;
    let message = object
        .messages
        .get(entry.message_index)
        .ok_or_else(|| io::Error::other("drawable message disappeared before attachment"))?;
    if comment_identifier(&message.data, entry.comment_path)?.is_some() {
        return Err(io::Error::other("drawable already owns a comment").into());
    }
    let data = append_comment_reference(&message.data, entry.comment_path, comment_id)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(entry.message_index)
        .ok_or_else(|| io::Error::other("drawable message metadata disappeared"))?;
    if info.object_references.contains(&comment_id) {
        return Err(io::Error::other("drawable already references comment").into());
    }
    info.object_references.push(comment_id);
    object.replace_message(
        entry.message_index,
        RawMessage {
            type_: entry.message_type,
            data,
        },
    )?;
    replace_component(source, &component_name, archive)
}

fn comment_object(
    source: &[u8],
    identifier: u64,
) -> TestResult<(String, Archive, usize, tsd::CommentStorageArchive)> {
    let (component_name, archive) = component_containing_object(source, identifier)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("comment object disappeared"))?;
    let (message_index, message) = unique_message(object, COMMENT_STORAGE_MESSAGE_TYPE)?;
    let comment = tsd::CommentStorageArchive::decode(message)?;
    Ok((component_name, archive, message_index, comment))
}

fn max_identifier(source: &[u8]) -> TestResult<u64> {
    component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| archive.objects.into_iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .ok_or_else(|| io::Error::other("native fixture has no objects").into())
}

fn append_reply_and_second_root(
    source: &[u8],
    root_identifier: u64,
) -> TestResult<(Vec<u8>, u64, u64)> {
    let (component_name, mut archive, message_index, root) =
        comment_object(source, root_identifier)?;
    let reply_identifier = max_identifier(source)?
        .checked_add(1)
        .ok_or_else(|| io::Error::other("reply identifier overflow"))?;
    let second_root_identifier = reply_identifier
        .checked_add(1)
        .ok_or_else(|| io::Error::other("second root identifier overflow"))?;
    let author = root.author.clone();
    let reply_payload = tsd::CommentStorageArchive {
        text: Some("shared reply branch".to_owned()),
        creation_date: Some(tsp::Date { seconds: 42.5 }),
        author: author.clone(),
        storage_uuid: Some(tsp::Uuid {
            lower: 0x1111_2222_3333_4444,
            upper: 0x5555_6666_7777_8888,
        }),
        ..Default::default()
    }
    .encode_to_vec();
    let mut reply = ArchiveObject::new(
        reply_identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: reply_payload,
        }],
    )?;
    if let Some(author) = author.as_ref() {
        reply.archive_info.message_infos[0]
            .object_references
            .push(author.identifier);
    }
    archive.insert_object(reply)?;

    let mut root_payload = archive
        .object(root_identifier)
        .and_then(|object| object.messages.get(message_index))
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("root payload disappeared"))?;
    append_length_delimited_field(
        &mut root_payload,
        4,
        &tsp::Reference {
            identifier: reply_identifier,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let root_object = archive
        .object_mut(root_identifier)
        .ok_or_else(|| io::Error::other("root object disappeared"))?;
    root_object.archive_info.message_infos[message_index]
        .object_references
        .push(reply_identifier);
    root_object.replace_message(
        message_index,
        RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: root_payload,
        },
    )?;

    let second_root_payload = tsd::CommentStorageArchive {
        text: Some("second shared root".to_owned()),
        creation_date: Some(tsp::Date { seconds: 43.5 }),
        author,
        replies: vec![tsp::Reference {
            identifier: reply_identifier,
            ..Default::default()
        }],
        storage_uuid: Some(tsp::Uuid {
            lower: 0x9999_aaaa_bbbb_cccc,
            upper: 0xdddd_eeee_ffff_0001,
        }),
        ..Default::default()
    }
    .encode_to_vec();
    let mut second_root = ArchiveObject::new(
        second_root_identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: second_root_payload,
        }],
    )?;
    if let Some(author) = root.author {
        second_root.archive_info.message_infos[0]
            .object_references
            .push(author.identifier);
    }
    second_root.archive_info.message_infos[0]
        .object_references
        .push(reply_identifier);
    archive.insert_object(second_root)?;
    Ok((
        replace_component(source, &component_name, archive)?,
        reply_identifier,
        second_root_identifier,
    ))
}

fn comment_signatures(source: &[u8]) -> TestResult<BTreeMap<u64, CommentSignature>> {
    let mut signatures = BTreeMap::new();
    for (_, archive) in component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("comment object has no identifier"))?;
            let Some((_, message)) = object
                .messages
                .iter()
                .enumerate()
                .find(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            signatures.insert(
                identifier,
                CommentSignature {
                    text: comment.text,
                    date_bits: comment.creation_date.map(|date| date.seconds.to_bits()),
                    author: comment.author.map(|reference| reference.identifier),
                    uuid: comment.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
                    replies: comment
                        .replies
                        .into_iter()
                        .map(|reference| reference.identifier)
                        .collect(),
                    payload: message.data.clone(),
                },
            );
        }
    }
    Ok(signatures)
}

fn native_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| archive.objects.into_iter())
        .filter_map(|object| object.archive_info.identifier)
        .collect())
}

fn assert_archive_references_are_live(source: &[u8]) -> TestResult<()> {
    let identifiers = native_object_ids(source)?;
    for (_, archive) in component_archives(source)? {
        for object in archive.objects {
            let owner = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            for message in &object.archive_info.message_infos {
                for identifier in &message.object_references {
                    assert!(
                        identifiers.contains(identifier),
                        "object {owner} has dangling message reference {identifier}"
                    );
                }
                for field in &message.field_infos {
                    for identifier in &field.object_references {
                        assert!(
                            identifiers.contains(identifier),
                            "object {owner} has dangling field reference {identifier}"
                        );
                    }
                }
            }
        }
    }
    Ok(())
}

fn annotation_author_state(
    source: &[u8],
) -> TestResult<(BTreeSet<u64>, BTreeSet<u64>, BTreeSet<u64>)> {
    let mut author_ids = BTreeSet::new();
    let mut registered_ids = BTreeSet::new();
    let mut aggregate_ids = BTreeSet::new();
    for (_, archive) in component_archives(source)? {
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(io::Error::other("author object has no identifier").into());
            };
            let mut is_author = false;
            for (message_index, message) in object.messages.iter().enumerate() {
                match message.type_ {
                    ANNOTATION_AUTHOR_MESSAGE_TYPE => is_author = true,
                    ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE => {
                        let info = object
                            .archive_info
                            .message_infos
                            .get(message_index)
                            .ok_or_else(|| {
                                io::Error::other("author storage header is missing message info")
                            })?;
                        for &reference in &info.object_references {
                            if reference == 0 || !aggregate_ids.insert(reference) {
                                return Err(io::Error::other(
                                    "author storage header has an invalid or duplicate reference",
                                )
                                .into());
                            }
                        }
                        for field in WireView::parse(&message.data)?.fields() {
                            if field.number() != 1 || field.wire_type() != 2 {
                                continue;
                            }
                            let reference = tsp::Reference::decode(field.payload())?;
                            if reference.identifier == 0 {
                                return Err(io::Error::other(
                                    "author storage contains a zero identifier",
                                )
                                .into());
                            }
                            if !registered_ids.insert(reference.identifier) {
                                return Err(io::Error::other(
                                    "author storage payload has a duplicate reference",
                                )
                                .into());
                            }
                        }
                    },
                    _ => {},
                }
            }
            if is_author {
                author_ids.insert(identifier);
            }
        }
    }
    Ok((author_ids, registered_ids, aggregate_ids))
}

fn native_watermark(source: &[u8]) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing native metadata component"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    let message = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing native package metadata"))?;
    Ok(tsp::PackageMetadata::decode(message.data.as_slice())?.last_object_identifier)
}

fn shared_root_fixture() -> TestResult<SharedGraphFixture> {
    let profile = fixture_profile(SOURCE)?;
    let root_identifier =
        comment_identifier_for_entry(SOURCE, &profile.drawables[profile.commented_position])?
            .ok_or_else(|| io::Error::other("commented drawable has no root"))?;
    let (with_reply, reply_identifier, _) = append_reply_and_second_root(SOURCE, root_identifier)?;
    let sibling_position = profile.empty_positions[0];
    let with_shared_root = attach_comment(
        &with_reply,
        &profile.drawables[sibling_position],
        root_identifier,
    )?;
    Ok(SharedGraphFixture {
        bytes: with_shared_root,
        original_position: profile.commented_position,
        sibling_position,
        root_identifier,
        reply_identifier,
        second_root_identifier: None,
    })
}

fn shared_reply_fixture() -> TestResult<SharedGraphFixture> {
    let profile = fixture_profile(SOURCE)?;
    let root_identifier =
        comment_identifier_for_entry(SOURCE, &profile.drawables[profile.commented_position])?
            .ok_or_else(|| io::Error::other("commented drawable has no root"))?;
    let (with_reply, reply_identifier, second_root_identifier) =
        append_reply_and_second_root(SOURCE, root_identifier)?;
    let second_root_position = profile.empty_positions[0];
    let with_second_root = attach_comment(
        &with_reply,
        &profile.drawables[second_root_position],
        second_root_identifier,
    )?;
    Ok(SharedGraphFixture {
        bytes: with_second_root,
        original_position: profile.commented_position,
        sibling_position: second_root_position,
        root_identifier,
        reply_identifier,
        second_root_identifier: Some(second_root_identifier),
    })
}

fn comment_identifier_for_entry(source: &[u8], entry: &DrawableEntry) -> TestResult<Option<u64>> {
    let (_, archive) = component_containing_object(source, entry.identifier)?;
    let object = archive
        .object(entry.identifier)
        .ok_or_else(|| io::Error::other("drawable object disappeared"))?;
    let message = object
        .messages
        .get(entry.message_index)
        .ok_or_else(|| io::Error::other("drawable message disappeared"))?;
    comment_identifier(&message.data, entry.comment_path)
}

#[test]
fn every_supported_wrapper_route_round_trips_from_a_native_square() -> TestResult {
    for (index, &(message_type, comment_path)) in SUPPORTED_ROUTE_VARIANTS.iter().enumerate() {
        let reverse_source_order = index == SUPPORTED_ROUTE_VARIANTS.len() - 1;
        let (source, target_identifier, target_position) =
            morph_square_route(SOURCE, message_type, comment_path, reverse_source_order)?;
        let package = Package::from_bytes(&source)?;
        let slide = SlideSelector::index(0);
        let target = DrawableSelector::index(target_position);
        let summaries = package.slide_drawables(slide)?;
        assert_eq!(
            summaries[target_position].kind(),
            expected_route_kind(message_type)
        );
        assert!(!summaries[target_position].has_comment());
        if reverse_source_order {
            assert_eq!(target_position, summaries.len() - 1);
        }

        let before = exact_bytes(&package)?;
        let created_text = format!("route {message_type} created");
        let created = package
            .edit_slide_drawable_comment(slide, target)?
            .set(&created_text)?
            .commit()?;
        let comment = created
            .package()
            .slide_drawable_comment(slide, target)?
            .ok_or_else(|| io::Error::other("route comment was not created"))?;
        assert_eq!(comment.text(), created_text.as_str());

        let cleared = created
            .package()
            .edit_slide_drawable_comment(slide, target)?
            .clear()?
            .commit()?;
        assert!(
            cleared
                .package()
                .slide_drawable_comment(slide, target)?
                .is_none()
        );
        let restored = cleared
            .package()
            .apply_slide_drawable_comment(&cleared.patch().inverse())?;
        assert_eq!(
            restored
                .package()
                .slide_drawable_comment(slide, target)?
                .as_ref()
                .map(|comment| comment.text()),
            Some(created_text.as_str())
        );
        let restored = restored
            .package()
            .apply_slide_drawable_comment(&created.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, before);

        let (_, archive) = component_containing_object(&source, target_identifier)?;
        let object = archive
            .object(target_identifier)
            .ok_or_else(|| io::Error::other("morphed route object disappeared"))?;
        assert!(
            object
                .messages
                .iter()
                .any(|message| message.type_ == message_type)
        );
    }
    Ok(())
}

#[test]
fn focused_reachability_guard_preserves_slide_title_and_noop_bytes() -> TestResult {
    let package = Package::from_bytes(SOURCE)?;
    let slide = SlideSelector::index(0);
    let summaries = package.slide_drawables(slide)?;
    let target = DrawableSelector::index(
        summaries
            .iter()
            .position(|summary| !summary.has_comment())
            .ok_or_else(|| io::Error::other("fixture has no empty drawable"))?,
    );
    let title = package
        .show()?
        .slides()
        .first()
        .and_then(|slide| slide.title())
        .map(str::to_owned);
    assert!(
        package
            .slide_drawable_comment(slide, DrawableSelector::index(summaries.len()))
            .is_err()
    );
    assert!(
        package
            .edit_slide_drawable_comment(slide, DrawableSelector::index(summaries.len()))
            .is_err()
    );

    let created = package
        .edit_slide_drawable_comment(slide, target)?
        .set("Title annotation")?
        .commit()?;
    assert_eq!(
        created
            .package()
            .show()?
            .slides()
            .first()
            .and_then(|slide| slide.title()),
        title.as_deref()
    );
    let bytes = exact_bytes(created.package())?;
    let noop = created
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("Title annotation")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, bytes);

    let reparsed = Package::from_bytes(&bytes)?;
    let cleared = reparsed
        .edit_slide_drawable_comment(slide, target)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(slide, target)?
            .is_none()
    );
    assert_eq!(
        cleared
            .package()
            .show()?
            .slides()
            .first()
            .and_then(|slide| slide.title()),
        title.as_deref()
    );
    Ok(())
}

#[test]
fn generated_author_create_and_clear_restores_registry_membership_and_suffix() -> TestResult {
    let profile = fixture_profile(SOURCE)?;
    let target = DrawableSelector::index(profile.empty_positions[0]);
    let slide = SlideSelector::index(0);
    let package = Package::from_bytes(SOURCE)?;
    let before = exact_bytes(&package)?;
    let before_authors = annotation_author_state(&before)?;
    let before_watermark = native_watermark(&before)?;
    assert!(!before_authors.0.is_empty());
    assert!(before_authors.1.is_superset(&before_authors.0));
    assert!(before_authors.2.is_subset(&before_authors.1));
    assert_archive_references_are_live(&before)?;

    let created = package
        .edit_slide_drawable_comment(slide, target)?
        .set("generated author cleanup")?
        .commit()?;
    let created_bytes = exact_bytes(created.package())?;
    let created_authors = annotation_author_state(&created_bytes)?;
    let generated_author = created_authors
        .0
        .difference(&before_authors.0)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(generated_author.len(), 1);
    assert_eq!(
        created_authors.1,
        before_authors
            .1
            .iter()
            .copied()
            .chain(generated_author.iter().copied())
            .collect()
    );
    assert!(created_authors.2.is_subset(&created_authors.1));
    assert!(created_authors.2.contains(&generated_author[0]));
    assert!(native_watermark(&created_bytes)? > before_watermark);
    assert_archive_references_are_live(&created_bytes)?;

    let cleared = created
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .clear()?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package())?;
    let cleared_authors = annotation_author_state(&cleared_bytes)?;
    assert_eq!(cleared_authors, before_authors);
    assert!(cleared_authors.2.is_subset(&cleared_authors.1));
    assert_eq!(native_watermark(&cleared_bytes)?, before_watermark);
    assert_archive_references_are_live(&cleared_bytes)?;
    assert_eq!(
        native_object_ids(&cleared_bytes)?,
        native_object_ids(&before)?
    );
    Ok(())
}

#[test]
fn clearing_non_tail_comment_preserves_watermark_until_tail_release() -> TestResult {
    let profile = fixture_profile(SOURCE)?;
    let slide = SlideSelector::index(0);
    let first_target = DrawableSelector::index(profile.empty_positions[0]);
    let second_target = DrawableSelector::index(profile.empty_positions[1]);
    let package = Package::from_bytes(SOURCE)?;
    let before = exact_bytes(&package)?;
    let before_authors = annotation_author_state(&before)?;
    let before_watermark = native_watermark(&before)?;

    let first = package
        .edit_slide_drawable_comment(slide, first_target)?
        .set("non-tail comment")?
        .commit()?;
    let first_watermark = native_watermark(&exact_bytes(first.package())?)?;
    assert!(first_watermark > before_watermark);

    let second = first
        .package()
        .edit_slide_drawable_comment(slide, second_target)?
        .set("tail comment")?
        .commit()?;
    let second_bytes = exact_bytes(second.package())?;
    let second_watermark = native_watermark(&second_bytes)?;
    assert!(second_watermark > first_watermark);
    assert_archive_references_are_live(&second_bytes)?;

    let cleared_non_tail = second
        .package()
        .edit_slide_drawable_comment(slide, first_target)?
        .clear()?
        .commit()?;
    let cleared_non_tail_bytes = exact_bytes(cleared_non_tail.package())?;
    assert_eq!(native_watermark(&cleared_non_tail_bytes)?, second_watermark);
    assert!(
        cleared_non_tail
            .package()
            .slide_drawable_comment(slide, first_target)?
            .is_none()
    );
    assert_eq!(
        cleared_non_tail
            .package()
            .slide_drawable_comment(slide, second_target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("tail comment")
    );
    let non_tail_authors = annotation_author_state(&cleared_non_tail_bytes)?;
    assert!(non_tail_authors.2.is_subset(&non_tail_authors.1));
    assert_archive_references_are_live(&cleared_non_tail_bytes)?;

    let cleared_tail = cleared_non_tail
        .package()
        .edit_slide_drawable_comment(slide, second_target)?
        .clear()?
        .commit()?;
    let cleared_tail_bytes = exact_bytes(cleared_tail.package())?;
    assert_eq!(native_watermark(&cleared_tail_bytes)?, before_watermark);
    assert_eq!(
        annotation_author_state(&cleared_tail_bytes)?,
        before_authors
    );
    assert_eq!(
        native_object_ids(&cleared_tail_bytes)?,
        native_object_ids(&before)?
    );
    assert_archive_references_are_live(&cleared_tail_bytes)?;
    Ok(())
}

#[test]
fn cross_component_root_read_and_edit_commit_without_source_mutation() -> TestResult {
    let fixture = cross_component_fixtures::root_foreign_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = exact_bytes(&package)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(fixture.target.position);
    assert!(package.slide_drawable_comment(slide, target)?.is_some());
    let commit = package
        .edit_slide_drawable_comment(slide, target)?
        .set("cross-component mutation")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_drawable_comment(slide, target)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("cross-component mutation")
    );
    assert_eq!(exact_bytes(&package)?, before);
    let restored = commit
        .package()
        .apply_slide_drawable_comment(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before);

    let changed = commit
        .package()
        .edit_slide_drawable_comment(slide, target)?
        .set("different source")?
        .commit()?;
    let changed_before = exact_bytes(changed.package())?;
    assert!(
        changed
            .package()
            .apply_slide_drawable_comment(commit.patch())
            .is_err()
    );
    assert_eq!(exact_bytes(changed.package())?, changed_before);
    Ok(())
}

#[test]
fn all_cross_component_fixture_builders_preserve_native_metadata() -> TestResult<()> {
    let (primary, sibling) =
        cross_component_fixtures::primary_targets(cross_component_fixtures::SOURCE)?;
    assert_ne!(primary.position, sibling.position);

    let builders: [(
        &str,
        fn() -> cross_component_fixtures::TestResult<cross_component_fixtures::RelocatedFixture>,
    ); 5] = [
        ("root", cross_component_fixtures::root_foreign_fixture),
        ("reply", cross_component_fixtures::reply_foreign_fixture),
        (
            "shared-root",
            cross_component_fixtures::shared_foreign_root_fixture,
        ),
        (
            "shared-reply",
            cross_component_fixtures::shared_foreign_reply_fixture,
        ),
        (
            "empty-drawable",
            cross_component_fixtures::foreign_drawable_fixture,
        ),
    ];

    for (name, build) in builders {
        let fixture = build()?;
        let package = Package::from_bytes(&fixture.bytes)?;
        let slide = SlideSelector::index(0);
        let target = DrawableSelector::index(fixture.target.position);
        let summaries = package.slide_drawables(slide)?;
        assert!(
            fixture.target.position < summaries.len(),
            "{name} fixture target is outside the slide inventory"
        );
        let selected = package.slide_drawable_comment(slide, target)?;
        if fixture.root_identifier.is_some() {
            assert!(selected.is_some(), "{name} fixture lost its root comment");
        } else {
            assert!(
                selected.is_none(),
                "{name} fixture unexpectedly has a root comment"
            );
        }
        if fixture.reply_identifier.is_some() {
            assert!(
                !package
                    .slide_drawable_comment_replies(slide, target)?
                    .is_empty(),
                "{name} fixture lost its relocated reply"
            );
        }

        let moved = fixture
            .reply_identifier
            .or(fixture.root_identifier)
            .unwrap_or(fixture.target.identifier);
        cross_component_fixtures::assert_metadata_relocated(
            &fixture.metadata_source,
            &fixture.bytes,
            moved,
        )?;

        let archives = cross_component_fixtures::component_archives(&fixture.bytes)?;
        let destination = if fixture.root_identifier.is_some() || fixture.reply_identifier.is_some()
        {
            cross_component_fixtures::AUTHOR_STORAGE_COMPONENT
        } else {
            cross_component_fixtures::STYLESHEET_COMPONENT
        };
        let moved_object = archives
            .iter()
            .find(|(component, archive)| {
                component == destination && archive.object(moved).is_some()
            })
            .and_then(|(_, archive)| archive.object(moved))
            .ok_or_else(|| io::Error::other(format!("{name} object was not relocated")))?;
        if fixture.root_identifier.is_some() || fixture.reply_identifier.is_some() {
            assert!(
                moved_object.messages.iter().any(|message| message.type_
                    == cross_component_fixtures::COMMENT_STORAGE_MESSAGE_TYPE),
                "{name} relocation did not retain comment storage payload"
            );
        }
        assert_archive_references_are_live(&fixture.bytes)?;
    }
    Ok(())
}

fn without_annotation_author_storage(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut current = source.to_vec();
    let mut removed_storage = false;
    for (component_name, mut archive) in component_archives(&current)? {
        let identifiers = archive
            .objects
            .iter()
            .filter(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
            })
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        if identifiers.is_empty() {
            continue;
        }
        for identifier in identifiers {
            archive.remove_object(identifier);
        }
        current = replace_component(&current, &component_name, archive)?;
        removed_storage = true;
    }
    if !removed_storage {
        return Err(io::Error::other("native fixture has no author storage").into());
    }

    // The root document is the ownership witness for annotation-author
    // storage.  Remove that reference as well as the object itself; leaving a
    // field-7 edge behind would make this an invalid package rather than the
    // legitimate host case where author registration is unavailable.
    for (component_name, mut archive) in component_archives(&current)? {
        let mut changed = false;
        for object in &mut archive.objects {
            let Some((message_index, message)) = object
                .messages
                .iter()
                .enumerate()
                .find(|(_, message)| message.type_ == 1)
            else {
                continue;
            };
            let storage_identifier = reference_at_path(&message.data, &[3, 1, 7])?;
            let Some(storage_identifier) = storage_identifier else {
                continue;
            };
            let data = transform_length_delimited_field(&message.data, 3, |tsa| {
                transform_length_delimited_field(tsa, 1, |tsk| {
                    patch_length_delimited_field(tsk, 7, true, None)
                })
            })?;
            object.archive_info.message_infos[message_index]
                .object_references
                .retain(|identifier| *identifier != storage_identifier);
            object.replace_message(message_index, RawMessage { type_: 1, data })?;
            changed = true;
            break;
        }
        if changed {
            current = replace_component(&current, &component_name, archive)?;
            return Ok(current);
        }
    }
    Err(io::Error::other("native document has no annotation-author-storage edge").into())
}

fn reference_at_path(payload: &[u8], path: &[u32]) -> TestResult<Option<u64>> {
    let mut current = payload;
    for field_number in path {
        let fields = WireView::parse(current)?
            .fields()
            .filter(|field| field.number() == *field_number)
            .collect::<Vec<_>>();
        if fields.is_empty() {
            return Ok(None);
        }
        if fields.len() != 1 || fields[0].wire_type() != 2 {
            return Err(io::Error::other("document reference path is malformed").into());
        }
        current = fields[0].payload();
    }
    let reference = tsp::Reference::decode(current)?;
    (reference.identifier != 0)
        .then_some(reference.identifier)
        .map(Some)
        .ok_or_else(|| io::Error::other("document author-storage reference is zero").into())
}

#[test]
fn shared_root_cow_and_clear_keep_every_retained_descendant() -> TestResult {
    let fixture = shared_root_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let original = DrawableSelector::index(fixture.original_position);
    let sibling = DrawableSelector::index(fixture.sibling_position);
    let before = exact_bytes(&package)?;
    let source_signatures = comment_signatures(&before)?;
    let source_root = package
        .slide_drawable_comment(slide, original)?
        .ok_or_else(|| io::Error::other("shared root is unreadable"))?;
    let source_replies = package.slide_drawable_comment_replies(slide, original)?;
    assert_eq!(source_replies.len(), 1);
    assert_eq!(source_replies[0].text(), "shared reply branch");
    assert_eq!(
        package
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some(source_root.text())
    );

    let cow = package
        .edit_slide_drawable_comment(slide, original)?
        .set("selected shared-root replacement")?
        .commit()?;
    assert_eq!(
        cow.package()
            .slide_drawable_comment(slide, original)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("selected shared-root replacement")
    );
    assert_eq!(
        cow.package()
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some(source_root.text())
    );
    assert_eq!(
        cow.package()
            .slide_drawable_comment_replies(slide, sibling)?[0]
            .text(),
        "shared reply branch"
    );
    let cow_signatures = comment_signatures(&exact_bytes(cow.package())?)?;
    assert_eq!(
        cow_signatures.get(&fixture.root_identifier),
        source_signatures.get(&fixture.root_identifier)
    );
    assert_eq!(
        cow_signatures.get(&fixture.reply_identifier),
        source_signatures.get(&fixture.reply_identifier)
    );

    let cleared = cow
        .package()
        .edit_slide_drawable_comment(slide, original)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(slide, original)?
            .is_none()
    );
    assert_eq!(
        cleared
            .package()
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some(source_root.text())
    );
    assert_eq!(
        cleared
            .package()
            .slide_drawable_comment_replies(slide, sibling)?[0]
            .text(),
        "shared reply branch"
    );
    let cleared_signatures = comment_signatures(&exact_bytes(cleared.package())?)?;
    assert_eq!(
        cleared_signatures.get(&fixture.root_identifier),
        source_signatures.get(&fixture.root_identifier)
    );
    assert_eq!(
        cleared_signatures.get(&fixture.reply_identifier),
        source_signatures.get(&fixture.reply_identifier)
    );
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    let restored = restored
        .package()
        .apply_slide_drawable_comment(&cow.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before);
    Ok(())
}

#[test]
fn shared_reply_clear_keeps_the_reply_for_the_unselected_root() -> TestResult {
    let fixture = shared_reply_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let slide = SlideSelector::index(0);
    let selected = DrawableSelector::index(fixture.original_position);
    let sibling = DrawableSelector::index(fixture.sibling_position);
    let before = exact_bytes(&package)?;
    let signatures = comment_signatures(&before)?;
    let second_root_identifier = fixture
        .second_root_identifier
        .ok_or_else(|| io::Error::other("shared-reply fixture has no second root"))?;
    assert_eq!(
        package
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("second shared root")
    );
    assert_eq!(
        package.slide_drawable_comment_replies(slide, sibling)?[0].text(),
        "shared reply branch"
    );

    let cleared = package
        .edit_slide_drawable_comment(slide, selected)?
        .clear()?
        .commit()?;
    assert!(
        cleared
            .package()
            .slide_drawable_comment(slide, selected)?
            .is_none()
    );
    assert_eq!(
        cleared
            .package()
            .slide_drawable_comment(slide, sibling)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("second shared root")
    );
    assert_eq!(
        cleared
            .package()
            .slide_drawable_comment_replies(slide, sibling)?[0]
            .text(),
        "shared reply branch"
    );
    let after_signatures = comment_signatures(&exact_bytes(cleared.package())?)?;
    assert_eq!(
        after_signatures.get(&fixture.reply_identifier),
        signatures.get(&fixture.reply_identifier)
    );
    assert_eq!(
        after_signatures.get(&second_root_identifier),
        signatures.get(&second_root_identifier)
    );
    let restored = cleared
        .package()
        .apply_slide_drawable_comment(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before);
    Ok(())
}

#[test]
fn creation_without_author_storage_keeps_an_authorless_comment() -> TestResult {
    let source = without_annotation_author_storage(SOURCE)?;
    let profile = fixture_profile(&source)?;
    let package = Package::from_bytes(&source)?;
    let slide = SlideSelector::index(0);
    let target = DrawableSelector::index(profile.empty_positions[0]);
    let created = package
        .edit_slide_drawable_comment(slide, target)?
        .set("authorless focused comment")?
        .commit()?;
    let comment = created
        .package()
        .slide_drawable_comment(slide, target)?
        .ok_or_else(|| io::Error::other("authorless comment was not created"))?;
    assert_eq!(comment.text(), "authorless focused comment");
    assert!(comment.author().is_none());
    assert!(comment.timestamp().is_some());
    let restored = created
        .package()
        .apply_slide_drawable_comment(&created.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}
