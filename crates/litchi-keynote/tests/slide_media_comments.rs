//! Focused Keynote media lifecycle coverage for direct drawable comments.
//!
//! The native fixtures in this test deliberately keep comment storage in the
//! slide component while the annotation author lives in the shared author
//! component.  The lifecycle owner must clone every comment-storage node,
//! preserve each storage UUID, and continue sharing the author object.  Final
//! removal must cull the comment-storage closure and its last component-level
//! author edge while retaining the shared author records.

use std::{
    collections::{BTreeMap, BTreeSet},
    env, fs, io,
    path::Path,
};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    decode_varint_from_bytes, encode_varint_into,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsd, tsp};
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideMediaLifecycleError, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_COMMENT_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_COMMENT_DUPLICATE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-duplicate-native.key");
const NATIVE_AUDIO_COMMENT_DUPLICATE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-comment-duplicate-native.key");
const NATIVE_AUDIO_COMMENT_REMOVAL: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-comment-removal-native.key");
const NATIVE_COMMENT_SHARED_REMOVAL: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-shared-removal-native.key");
const NATIVE_COMMENT_FINAL_REMOVAL: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-final-removal-native.key");

const NATIVE_SLIDE: u64 = 2_652_150;
const NATIVE_AUDIO_A: u64 = 2_652_595;
const NATIVE_AUDIO_B: u64 = 2_652_622;
const NATIVE_MOVIE_B: u64 = 2_653_610;
const NATIVE_COMMENT_MOVIE: u64 = 2_653_286;
const NATIVE_COMMENT_MOVIE_CLONE: u64 = 2_653_814;
const NATIVE_COMMENT_ROOT: u64 = 2_653_723;
const NATIVE_COMMENT_ROOT_CLONE: u64 = 2_653_826;
const NATIVE_COMMENT_AUTHOR: u64 = 2_653_721;
const NATIVE_COMMENT_AUTHOR_STORAGE: u64 = 2_652_381;
const AUDIO_COMMENT_TEXT: &str = "programmatic native audio comment";
const AUDIO_REPLY_TEXT: &str = "programmatic native audio reply";
const SAVED_AUDIO_DUPLICATE_ENV: &str = "LITCHI_KEYNOTE_AUDIO_COMMENT_NATIVE_SAVED_DUPLICATE_PATH";
const SAVED_AUDIO_REMOVAL_ENV: &str = "LITCHI_KEYNOTE_AUDIO_COMMENT_NATIVE_SAVED_REMOVAL_PATH";
const NATIVE_SLIDE_MESSAGE_TYPE: u32 = 5;
const NATIVE_MOVIE_MESSAGE_TYPE: u32 = 3_007;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const NATIVE_ROTATED_COMPONENT_BEFORE: u64 = 2_653_653;
const NATIVE_ROTATED_COMPONENT_AFTER: u64 = 2_654_075;
const NATIVE_ROTATED_DOCUMENT_OBJECT: u64 = 2_652_149;

fn native_comment_author_edge() -> (u64, u64, Option<u64>) {
    (
        NATIVE_SLIDE,
        NATIVE_COMMENT_AUTHOR_STORAGE,
        Some(NATIVE_COMMENT_AUTHOR),
    )
}

const SYNTHETIC_REPLY_ID: u64 = 2_653_900;
const SYNTHETIC_SURVIVING_COMMENT_ID: u64 = 2_654_500;
const SYNTHETIC_MISSING_MOVIE_COMMENT_ID: u64 = 2_654_501;
const SYNTHETIC_MISSING_REPLY_ID: u64 = 2_654_502;
const SYNTHETIC_REPLY_UUID: tsp::Uuid = tsp::Uuid {
    lower: 0x0bad_cafe_dead_beef,
    upper: 0x0123_4567_89ab_cdef,
};
const SYNTHETIC_SURVIVING_COMMENT_UUID: tsp::Uuid = tsp::Uuid {
    lower: 0x1357_9bdf_2468_ace0,
    upper: 0x0eca_8642_fdb9_7531,
};
const UNKNOWN_COMMENT_MARKER: &[u8] = b"litchi-keynote-comment-unknown-extension";
const UNKNOWN_COMMENT_REFERENCE_MARKER: &[u8] =
    b"litchi-keynote-comment-reference-unknown-extension";

#[derive(Debug, Clone, Copy)]
enum MovieCommentReferenceMutation {
    UnknownField,
    DeprecatedType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentStorageNode {
    identifier: u64,
    text: String,
    author: Option<u64>,
    replies: Vec<u64>,
    storage_uuid: Option<(u64, u64)>,
    payload: Vec<u8>,
    object_references: Vec<u64>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentGraph {
    movie_identifier: u64,
    root_identifier: u64,
    nodes: Vec<CommentStorageNode>,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn sorted_media_payloads(
    package: &Package,
    positions: impl IntoIterator<Item = usize>,
    part: MediaPart,
) -> TestResult<Vec<Vec<u8>>> {
    let mut payloads = positions
        .into_iter()
        .map(|position| {
            package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(position),
                    part,
                )
                .map(|payload| payload.to_vec())
                .map_err(Into::into)
        })
        .collect::<TestResult<Vec<_>>>()?;
    payloads.sort();
    Ok(payloads)
}

fn sorted_movie_payloads(
    package: &Package,
    positions: impl IntoIterator<Item = usize>,
) -> TestResult<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut payloads = positions
        .into_iter()
        .map(|position| {
            Ok((
                package
                    .slide_media_data(
                        SlideSelector::index(0),
                        MovieSelector::index(position),
                        MediaPart::Content,
                    )?
                    .to_vec(),
                package
                    .slide_media_data(
                        SlideSelector::index(0),
                        MovieSelector::index(position),
                        MediaPart::Poster,
                    )?
                    .to_vec(),
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    payloads.sort();
    Ok(payloads)
}

fn export_if_requested(package: &Package, name: &str) -> TestResult<()> {
    let Some(directory) = env::var_os("LITCHI_KEYNOTE_MEDIA_COMMENT_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    let path = directory.join(name);
    fs::write(&path, exact_bytes(package)?)?;
    eprintln!(
        "exported Keynote media-comment candidate to {}",
        path.display()
    );
    Ok(())
}

fn read_fixture_override(environment: &str, fixture: &[u8]) -> TestResult<Vec<u8>> {
    if let Some(path) = env::var_os(environment) {
        return Ok(fs::read(path)?);
    }
    Ok(fixture.to_vec())
}

fn native_component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
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

fn native_component_containing_object(
    source: &[u8],
    identifier: u64,
) -> TestResult<(String, Archive)> {
    native_component_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")))
        .map_err(Into::into)
}

fn native_object_exists(source: &[u8], identifier: u64) -> TestResult<bool> {
    Ok(native_component_archives(source)?
        .into_iter()
        .any(|(_, archive)| archive.object(identifier).is_some()))
}

fn native_metadata(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let (_, archive) = native_component_archives(source)?
        .into_iter()
        .find(|(name, _)| name == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing native PackageMetadata component"))?;
    let payload = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or_else(|| io::Error::other("missing native PackageMetadata payload"))?;
    Ok(tsp::PackageMetadata::decode(payload)?)
}

fn native_external_edges(source: &[u8]) -> TestResult<BTreeSet<(u64, u64, Option<u64>)>> {
    Ok(native_metadata(source)?
        .components
        .into_iter()
        .flat_map(|component| {
            component
                .external_references
                .into_iter()
                .map(move |reference| {
                    (
                        component.identifier,
                        reference.component_identifier,
                        reference.object_identifier,
                    )
                })
        })
        .collect())
}

fn native_component_effective_locators(source: &[u8]) -> TestResult<BTreeMap<u64, String>> {
    let mut locators = BTreeMap::new();
    for component in native_metadata(source)?.components {
        let locator = component
            .locator
            .unwrap_or(component.preferred_locator)
            .to_owned();
        if locators.insert(component.identifier, locator).is_some() {
            return Err(io::Error::other("native metadata repeats a current component").into());
        }
    }
    Ok(locators)
}

fn native_semantic_component_key(locators: &BTreeMap<u64, String>, identifier: u64) -> String {
    locators
        .get(&identifier)
        .cloned()
        .unwrap_or_else(|| format!("#component-{identifier}"))
}

fn native_semantic_external_edges(
    source: &[u8],
) -> TestResult<BTreeSet<(String, String, Option<u64>)>> {
    let locators = native_component_effective_locators(source)?;
    Ok(native_metadata(source)?
        .components
        .into_iter()
        .flat_map(|component| {
            let source_key = native_semantic_component_key(&locators, component.identifier);
            let locators = &locators;
            component
                .external_references
                .into_iter()
                .map(move |reference| {
                    let target_key =
                        native_semantic_component_key(locators, reference.component_identifier);
                    (source_key.clone(), target_key, reference.object_identifier)
                })
        })
        .collect())
}

fn native_semantic_comment_author_edge(source: &[u8]) -> TestResult<(String, String, Option<u64>)> {
    let locators = native_component_effective_locators(source)?;
    let (source_identifier, target_identifier, object_identifier) = native_comment_author_edge();
    Ok((
        native_semantic_component_key(&locators, source_identifier),
        native_semantic_component_key(&locators, target_identifier),
        object_identifier,
    ))
}

fn assert_compact_edge_sets_equal<T: Ord + std::fmt::Debug>(
    actual: &BTreeSet<T>,
    expected: &BTreeSet<T>,
    context: &str,
) {
    let missing = expected.difference(actual).take(4).collect::<Vec<_>>();
    let unexpected = actual.difference(expected).take(4).collect::<Vec<_>>();
    assert!(
        missing.is_empty() && unexpected.is_empty(),
        "{context}: edge sets differ (actual {}, expected {}), first missing {missing:?}, first unexpected {unexpected:?}",
        actual.len(),
        expected.len()
    );
}

fn native_component_stream_containing_object(
    source: &[u8],
    identifier: u64,
) -> TestResult<(String, Vec<u8>)> {
    for entry in Catalog::from_bytes(source)?.iter() {
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
        if archive.object(identifier).is_some() {
            return Ok((entry.name().to_owned(), stream));
        }
    }
    Err(io::Error::other(format!(
        "missing native object stream containing {identifier}"
    ))
    .into())
}

fn native_object_message<'a>(
    archive: &'a Archive,
    identifier: u64,
    type_: u32,
) -> TestResult<&'a [u8]> {
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .map(|message| message.data.as_slice())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier} type {type_}")))
        .map_err(Into::into)
}

fn replace_native_component_archive(
    source: &[u8],
    component_name: &str,
    archive: Archive,
) -> TestResult<Vec<u8>> {
    let bytes = archive.to_bytes()?;
    replace_native_component_stream(source, component_name, &bytes)
}

fn replace_native_comment_payload(
    source: &[u8],
    identifier: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    replace_native_object_payload(source, identifier, COMMENT_STORAGE_MESSAGE_TYPE, payload)
}

fn replace_native_movie_payload(
    source: &[u8],
    identifier: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    replace_native_object_payload(source, identifier, NATIVE_MOVIE_MESSAGE_TYPE, payload)
}

fn replace_native_object_payload(
    source: &[u8],
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) = native_component_containing_object(source, identifier)?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other(format!("missing comment object {identifier}")))?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or_else(|| {
            io::Error::other(format!(
                "missing object {identifier} message type {message_type}"
            ))
        })?;
    object.replace_message(
        index,
        RawMessage {
            type_: message_type,
            data: payload,
        },
    )?;
    replace_native_component_archive(source, &component_name, archive)
}

fn with_unknown_archive_info_field(
    source: &[u8],
    identifier: u64,
    field_number: u32,
    payload: &[u8],
) -> TestResult<Vec<u8>> {
    let (component_name, bytes) = native_component_stream_containing_object(source, identifier)?;
    let archive = Archive::parse(&bytes)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other(format!("missing archive object {identifier}")))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(
        bytes
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("archive object header offset is invalid"))?,
    )?;
    let header_length = usize::try_from(header_length)?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("archive object header start overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("archive object header end overflow"))?;
    if header_end != data_offset {
        return Err(io::Error::other("archive object offsets disagree").into());
    }
    let mut header = bytes
        .get(header_start..header_end)
        .ok_or_else(|| io::Error::other("archive object header range is invalid"))?
        .to_vec();
    append_length_delimited_field(&mut header, field_number, payload)?;

    let mut modified = Vec::with_capacity(bytes.len().saturating_add(header.len()));
    modified.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut modified, u64::try_from(header.len())?);
    modified.extend_from_slice(&header);
    modified.extend_from_slice(&bytes[data_offset..]);
    let reparsed = Archive::parse(&modified)?;
    assert_eq!(reparsed.to_bytes()?, modified);
    replace_native_component_stream(source, &component_name, &modified)
}

fn replace_native_component_stream(
    source: &[u8],
    component_name: &str,
    bytes: &[u8],
) -> TestResult<Vec<u8>> {
    let replacement = SnappyStream::compress(bytes)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
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

fn with_native_movie_comment_reference_mutation(
    source: &[u8],
    mutation: MovieCommentReferenceMutation,
) -> TestResult<Vec<u8>> {
    let (_, archive) = native_component_containing_object(source, NATIVE_COMMENT_MOVIE)?;
    let movie_payload =
        native_object_message(&archive, NATIVE_COMMENT_MOVIE, NATIVE_MOVIE_MESSAGE_TYPE)?.to_vec();
    let root = WireView::parse(&movie_payload)?;
    let mut rewritten_root = Vec::with_capacity(movie_payload.len().saturating_add(8));
    let mut changed = false;
    for root_field in root.fields() {
        if root_field.number() != 1 {
            rewritten_root.extend_from_slice(root_field.raw());
            continue;
        }
        let drawable = WireView::parse(root_field.payload())?;
        let mut rewritten_drawable =
            Vec::with_capacity(root_field.payload().len().saturating_add(8));
        for drawable_field in drawable.fields() {
            if drawable_field.number() != 6 {
                rewritten_drawable.extend_from_slice(drawable_field.raw());
                continue;
            }
            if changed {
                return Err(io::Error::other("native movie has duplicate comment fields").into());
            }
            let mut reference = drawable_field.payload().to_vec();
            match mutation {
                MovieCommentReferenceMutation::UnknownField => {
                    append_length_delimited_field(
                        &mut reference,
                        99,
                        UNKNOWN_COMMENT_REFERENCE_MARKER,
                    )?;
                },
                MovieCommentReferenceMutation::DeprecatedType => {
                    append_varint_field(&mut reference, 2, 1)?;
                },
            }
            append_length_delimited_field(&mut rewritten_drawable, 6, &reference)?;
            changed = true;
        }
        append_length_delimited_field(&mut rewritten_root, 1, &rewritten_drawable)?;
    }
    if !changed {
        return Err(io::Error::other("native movie has no comment field").into());
    }
    replace_native_movie_payload(source, NATIVE_COMMENT_MOVIE, rewritten_root)
}

fn with_native_shared_movie_comment(
    source: &[u8],
    movie_identifier: u64,
    root_identifier: u64,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) =
        native_component_containing_object(source, movie_identifier)?;
    let movie_object = archive
        .object_mut(movie_identifier)
        .ok_or_else(|| io::Error::other(format!("missing native movie {movie_identifier}")))?;
    let message_index = movie_object
        .messages
        .iter()
        .position(|message| message.type_ == NATIVE_MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native movie has no movie payload"))?;
    let payload = movie_object.messages[message_index].data.clone();
    let root = WireView::parse(&payload)?;
    let mut rewritten_root = Vec::with_capacity(payload.len().saturating_add(16));
    let mut drawable_count = 0usize;
    for root_field in root.fields() {
        if root_field.number() != 1 {
            rewritten_root.extend_from_slice(root_field.raw());
            continue;
        }
        drawable_count = drawable_count
            .checked_add(1)
            .ok_or_else(|| io::Error::other("native movie drawable count overflow"))?;
        let drawable = WireView::parse(root_field.payload())?;
        let mut rewritten_drawable =
            Vec::with_capacity(root_field.payload().len().saturating_add(8));
        let mut comment_count = 0usize;
        for drawable_field in drawable.fields() {
            if drawable_field.number() == 6 {
                comment_count = comment_count
                    .checked_add(1)
                    .ok_or_else(|| io::Error::other("native movie comment count overflow"))?;
            }
            rewritten_drawable.extend_from_slice(drawable_field.raw());
        }
        if comment_count != 0 {
            return Err(io::Error::other("native movie already has a direct comment").into());
        }
        let reference = tsp::Reference {
            identifier: root_identifier,
            ..Default::default()
        }
        .encode_to_vec();
        append_length_delimited_field(&mut rewritten_drawable, 6, &reference)?;
        append_length_delimited_field(&mut rewritten_root, 1, &rewritten_drawable)?;
    }
    if drawable_count != 1 {
        return Err(io::Error::other("native movie must have one drawable envelope").into());
    }
    let info = movie_object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or_else(|| io::Error::other("native movie message metadata is missing"))?;
    if info.object_references.contains(&root_identifier) {
        return Err(io::Error::other("native movie already references comment root").into());
    }
    info.object_references.push(root_identifier);
    movie_object.replace_message(
        message_index,
        RawMessage {
            type_: NATIVE_MOVIE_MESSAGE_TYPE,
            data: rewritten_root,
        },
    )?;
    replace_native_component_archive(source, &component_name, archive)
}

fn without_native_movie_comment_header_reference(
    source: &[u8],
    movie_identifier: u64,
    root_identifier: u64,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) =
        native_component_containing_object(source, movie_identifier)?;
    let movie_object = archive
        .object_mut(movie_identifier)
        .ok_or_else(|| io::Error::other(format!("missing native movie {movie_identifier}")))?;
    let message_index = movie_object
        .messages
        .iter()
        .position(|message| message.type_ == NATIVE_MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native movie has no movie payload"))?;
    let info = movie_object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or_else(|| io::Error::other("native movie message metadata is missing"))?;
    let before = info.object_references.len();
    info.object_references
        .retain(|identifier| *identifier != root_identifier);
    if info.object_references.len() + 1 != before {
        return Err(io::Error::other("native movie comment header reference is not unique").into());
    }
    replace_native_component_archive(source, &component_name, archive)
}

fn with_native_unselected_comment_reply(
    source: &[u8],
    reply_identifier: u64,
    include_reply_header_reference: bool,
) -> TestResult<Vec<u8>> {
    let graph = native_comment_graph(source, NATIVE_COMMENT_MOVIE)?;
    let root = native_comment_node(source, graph.root_identifier)?;
    let author = root
        .author
        .ok_or_else(|| io::Error::other("native comment has no author"))?;
    let (component_name, mut archive) =
        native_component_containing_object(source, graph.root_identifier)?;
    let payload = tsd::CommentStorageArchive {
        text: Some("Unselected hidden reply reference".to_owned()),
        creation_date: Some(tsp::Date { seconds: 43.0 }),
        author: Some(tsp::Reference {
            identifier: author,
            ..Default::default()
        }),
        replies: vec![tsp::Reference {
            identifier: reply_identifier,
            ..Default::default()
        }],
        storage_uuid: Some(SYNTHETIC_SURVIVING_COMMENT_UUID),
        ..Default::default()
    }
    .encode_to_vec();
    let mut object = ArchiveObject::new(
        SYNTHETIC_SURVIVING_COMMENT_ID,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
    )?;
    object.archive_info.message_infos[0]
        .object_references
        .push(author);
    if include_reply_header_reference {
        object.archive_info.message_infos[0]
            .object_references
            .push(reply_identifier);
    }
    archive.insert_object(object)?;
    replace_native_component_archive(source, &component_name, archive)
}

fn with_native_unselected_comment_reply_missing_header(source: &[u8]) -> TestResult<Vec<u8>> {
    with_native_unselected_comment_reply(source, SYNTHETIC_REPLY_ID, false)
}

fn with_native_unselected_comment_reply_missing_target(source: &[u8]) -> TestResult<Vec<u8>> {
    with_native_unselected_comment_reply(source, SYNTHETIC_MISSING_REPLY_ID, true)
}

fn native_movie_archive(source: &[u8], identifier: u64) -> TestResult<tsd::MovieArchive> {
    let (_, archive) = native_component_containing_object(source, identifier)?;
    Ok(tsd::MovieArchive::decode(native_object_message(
        &archive,
        identifier,
        NATIVE_MOVIE_MESSAGE_TYPE,
    )?)?)
}

fn native_slide_archive(source: &[u8]) -> TestResult<kn::SlideArchive> {
    let (_, archive) = native_component_containing_object(source, NATIVE_SLIDE)?;
    Ok(kn::SlideArchive::decode(native_object_message(
        &archive,
        NATIVE_SLIDE,
        NATIVE_SLIDE_MESSAGE_TYPE,
    )?)?)
}

fn native_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    Ok(native_slide_archive(source)?
        .owned_drawables
        .into_iter()
        .filter_map(|reference| {
            let identifier = reference.identifier;
            native_movie_archive(source, identifier)
                .ok()
                .map(|_| identifier)
        })
        .collect())
}

fn native_audio_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    Ok(native_media_ids(source)?
        .into_iter()
        .filter(|identifier| {
            native_movie_archive(source, *identifier)
                .ok()
                .is_some_and(|movie| movie.audio_only == Some(true))
        })
        .collect())
}

fn native_commented_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    Ok(native_media_ids(source)?
        .into_iter()
        .filter(|identifier| {
            native_movie_archive(source, *identifier)
                .ok()
                .and_then(|movie| movie.super_.comment)
                .is_some()
        })
        .collect())
}

fn native_commented_audio_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    Ok(native_audio_ids(source)?
        .into_iter()
        .filter(|identifier| {
            native_movie_archive(source, *identifier)
                .ok()
                .and_then(|movie| movie.super_.comment)
                .is_some()
        })
        .collect())
}

fn native_comment_node(source: &[u8], identifier: u64) -> TestResult<CommentStorageNode> {
    let (_, archive) = native_component_containing_object(source, identifier)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other(format!("missing comment object {identifier}")))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other(format!("missing comment payload {identifier}")))?;
    let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
    let object_references = object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        })
        .copied()
        .collect();
    Ok(CommentStorageNode {
        identifier,
        text: comment
            .text
            .ok_or_else(|| io::Error::other(format!("comment {identifier} has no text")))?,
        author: comment.author.map(|reference| reference.identifier),
        replies: comment
            .replies
            .into_iter()
            .map(|reference| reference.identifier)
            .collect(),
        storage_uuid: comment.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
        payload: message.data.clone(),
        object_references,
    })
}

fn native_annotation_author_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for (_, archive) in native_component_archives(source)? {
        identifiers.extend(archive.objects.into_iter().filter_map(|object| {
            object.archive_info.identifier.filter(|_| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
            })
        }));
    }
    Ok(identifiers)
}

fn native_annotation_author_storage_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for (_, archive) in native_component_archives(source)? {
        identifiers.extend(archive.objects.into_iter().filter_map(|object| {
            object.archive_info.identifier.filter(|_| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
            })
        }));
    }
    Ok(identifiers)
}

fn native_comment_graph(source: &[u8], movie_identifier: u64) -> TestResult<CommentGraph> {
    let movie = native_movie_archive(source, movie_identifier)?;
    let root_identifier = movie
        .super_
        .comment
        .ok_or_else(|| io::Error::other(format!("movie {movie_identifier} has no comment")))?
        .identifier;
    let mut nodes = Vec::new();
    let mut pending = vec![root_identifier];
    let mut seen = BTreeSet::new();
    while let Some(identifier) = pending.pop() {
        if !seen.insert(identifier) {
            continue;
        }
        let node = native_comment_node(source, identifier)?;
        pending.extend(node.replies.iter().rev().copied());
        nodes.push(node);
    }
    Ok(CommentGraph {
        movie_identifier,
        root_identifier,
        nodes,
    })
}

fn assert_comment_graph_clone(
    source: &[u8],
    candidate: &[u8],
    source_movie: u64,
    clone_movie: u64,
) -> TestResult<()> {
    let source_graph = native_comment_graph(source, source_movie)?;
    let candidate_graph = native_comment_graph(candidate, clone_movie)?;
    assert_eq!(source_graph.movie_identifier, source_movie);
    assert_eq!(candidate_graph.movie_identifier, clone_movie);
    assert_ne!(
        source_graph.root_identifier,
        candidate_graph.root_identifier
    );
    assert_eq!(source_graph.nodes.len(), candidate_graph.nodes.len());

    let source_ids = source_graph
        .nodes
        .iter()
        .map(|node| node.identifier)
        .collect::<BTreeSet<_>>();
    let candidate_ids = candidate_graph
        .nodes
        .iter()
        .map(|node| node.identifier)
        .collect::<BTreeSet<_>>();
    assert!(source_ids.is_disjoint(&candidate_ids));

    let remap = source_graph
        .nodes
        .iter()
        .zip(&candidate_graph.nodes)
        .map(|(source, candidate)| (source.identifier, candidate.identifier))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(
        remap.get(&source_graph.root_identifier),
        Some(&candidate_graph.root_identifier)
    );
    for (source_node, candidate_node) in source_graph.nodes.iter().zip(&candidate_graph.nodes) {
        assert_eq!(candidate_node.text, source_node.text);
        assert_eq!(candidate_node.author, source_node.author);
        assert_eq!(candidate_node.storage_uuid, source_node.storage_uuid);
        let expected_replies = source_node
            .replies
            .iter()
            .map(|identifier| {
                remap
                    .get(identifier)
                    .copied()
                    .ok_or_else(|| io::Error::other("comment reply escaped selected graph"))
            })
            .collect::<Result<Vec<_>, _>>()?;
        assert_eq!(candidate_node.replies, expected_replies);
        assert!(
            candidate_node
                .object_references
                .iter()
                .any(|identifier| *identifier == candidate_node.author.unwrap_or_default())
                || candidate_node.author.is_none(),
            "cloned comment lost its author reference"
        );
    }

    let source_author_ids = source_graph
        .nodes
        .iter()
        .filter_map(|node| node.author)
        .collect::<BTreeSet<_>>();
    let candidate_author_ids = candidate_graph
        .nodes
        .iter()
        .filter_map(|node| node.author)
        .collect::<BTreeSet<_>>();
    assert_eq!(candidate_author_ids, source_author_ids);
    assert_eq!(
        native_annotation_author_ids(candidate)?,
        native_annotation_author_ids(source)?,
        "comment duplication must reuse annotation-author objects"
    );
    Ok(())
}

fn assert_audio_comment_shape(source: &[u8], audio_identifier: u64) -> TestResult<()> {
    let graph = native_comment_graph(source, audio_identifier)?;
    assert_eq!(graph.nodes.len(), 2);
    let root = &graph.nodes[0];
    let reply = &graph.nodes[1];
    assert_eq!(root.text, AUDIO_COMMENT_TEXT);
    assert_eq!(root.replies, vec![reply.identifier]);
    assert_eq!(reply.text, AUDIO_REPLY_TEXT);
    assert!(root.author.is_some());
    assert_eq!(root.author, reply.author);
    assert!(root.storage_uuid.is_some());
    assert!(reply.storage_uuid.is_some());
    assert_ne!(root.storage_uuid, reply.storage_uuid);
    Ok(())
}

fn assert_comment_graph_semantics(expected: &CommentGraph, actual: &CommentGraph) {
    assert_eq!(actual.nodes.len(), expected.nodes.len());
    for (expected_node, actual_node) in expected.nodes.iter().zip(&actual.nodes) {
        assert_eq!(actual_node.text, expected_node.text);
        assert_eq!(actual_node.author, expected_node.author);
        assert_eq!(actual_node.storage_uuid, expected_node.storage_uuid);
        assert_eq!(actual_node.replies.len(), expected_node.replies.len());
    }
}

fn with_native_reply_graph(source: &[u8], unknown_fields: bool) -> TestResult<Vec<u8>> {
    let root_identifier = native_comment_graph(source, NATIVE_COMMENT_MOVIE)?.root_identifier;
    let root = native_comment_node(source, root_identifier)?;
    let author = root
        .author
        .ok_or_else(|| io::Error::other("native comment has no author"))?;
    let (component_name, mut archive) =
        native_component_containing_object(source, root_identifier)?;
    let root_object = archive
        .object_mut(root_identifier)
        .ok_or_else(|| io::Error::other("missing native comment root"))?;
    let root_index = root_object
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing native comment root payload"))?;
    let mut root_payload = root_object.messages[root_index].data.clone();
    append_length_delimited_field(
        &mut root_payload,
        4,
        &tsp::Reference {
            identifier: SYNTHETIC_REPLY_ID,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    if unknown_fields {
        append_length_delimited_field(&mut root_payload, 199, UNKNOWN_COMMENT_MARKER)?;
    }
    root_object.replace_message(
        root_index,
        RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: root_payload,
        },
    )?;
    root_object.archive_info.message_infos[root_index]
        .object_references
        .push(SYNTHETIC_REPLY_ID);

    let mut reply_payload = tsd::CommentStorageArchive {
        text: Some("Synthetic lifecycle reply".to_owned()),
        creation_date: Some(tsp::Date { seconds: 42.0 }),
        author: Some(tsp::Reference {
            identifier: author,
            ..Default::default()
        }),
        storage_uuid: Some(SYNTHETIC_REPLY_UUID),
        ..Default::default()
    }
    .encode_to_vec();
    if unknown_fields {
        append_length_delimited_field(&mut reply_payload, 199, UNKNOWN_COMMENT_MARKER)?;
    }
    let mut reply = ArchiveObject::new(
        SYNTHETIC_REPLY_ID,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: reply_payload,
        }],
    )?;
    reply.archive_info.message_infos[0]
        .object_references
        .push(author);
    archive.insert_object(reply)?;
    replace_native_component_archive(source, &component_name, archive)
}

fn with_native_comment_cycle(source: &[u8]) -> TestResult<Vec<u8>> {
    let source = with_native_reply_graph(source, false)?;
    let graph = native_comment_graph(&source, NATIVE_COMMENT_MOVIE)?;
    let root = graph.root_identifier;
    let node = graph
        .nodes
        .first()
        .ok_or_else(|| io::Error::other("reply fixture has no root"))?;
    let mut payload = node.payload.clone();
    append_length_delimited_field(
        &mut payload,
        4,
        &tsp::Reference {
            identifier: root,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    replace_native_comment_payload(&source, root, payload)
}

fn with_native_malformed_reply(source: &[u8]) -> TestResult<Vec<u8>> {
    let source = with_native_reply_graph(source, false)?;
    let graph = native_comment_graph(&source, NATIVE_COMMENT_MOVIE)?;
    let reply = graph
        .nodes
        .get(1)
        .ok_or_else(|| io::Error::other("reply fixture has no reply"))?;
    replace_native_comment_payload(&source, reply.identifier, vec![0xff])
}

fn with_native_unknown_root_reply_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let source = with_native_reply_graph(source, true)?;
    let graph = native_comment_graph(&source, NATIVE_COMMENT_MOVIE)?;
    let root = graph.root_identifier;
    let node = graph
        .nodes
        .first()
        .ok_or_else(|| io::Error::other("reply fixture has no root"))?;
    let mut payload = node.payload.clone();
    append_length_delimited_field(
        &mut payload,
        199,
        &tsp::Reference {
            identifier: SYNTHETIC_REPLY_ID,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    replace_native_comment_payload(&source, root, payload)
}

#[test]
fn native_comment_duplicate_oracle_matches_storage_identity_policy() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    let duplicate = Package::from_bytes(NATIVE_COMMENT_DUPLICATE)?;
    baseline.validate()?;
    duplicate.validate()?;
    assert_eq!(
        native_media_ids(NATIVE_COMMENT_BASELINE)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_COMMENT_MOVIE,
            NATIVE_MOVIE_B,
        ]
    );
    assert_eq!(
        native_media_ids(NATIVE_COMMENT_DUPLICATE)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_COMMENT_MOVIE,
            NATIVE_MOVIE_B,
            NATIVE_COMMENT_MOVIE_CLONE,
        ]
    );
    assert_eq!(
        native_commented_media_ids(NATIVE_COMMENT_BASELINE)?,
        vec![NATIVE_COMMENT_MOVIE]
    );
    assert_eq!(
        native_commented_media_ids(NATIVE_COMMENT_DUPLICATE)?,
        vec![NATIVE_COMMENT_MOVIE, NATIVE_COMMENT_MOVIE_CLONE]
    );

    let source_graph = native_comment_graph(NATIVE_COMMENT_BASELINE, NATIVE_COMMENT_MOVIE)?;
    let clone_graph = native_comment_graph(NATIVE_COMMENT_DUPLICATE, NATIVE_COMMENT_MOVIE_CLONE)?;
    assert_eq!(source_graph.root_identifier, NATIVE_COMMENT_ROOT);
    assert_eq!(clone_graph.root_identifier, NATIVE_COMMENT_ROOT_CLONE);
    assert_comment_graph_clone(
        NATIVE_COMMENT_BASELINE,
        NATIVE_COMMENT_DUPLICATE,
        NATIVE_COMMENT_MOVIE,
        NATIVE_COMMENT_MOVIE_CLONE,
    )?;
    assert_eq!(
        native_comment_node(NATIVE_COMMENT_DUPLICATE, NATIVE_COMMENT_ROOT)?.payload,
        native_comment_node(NATIVE_COMMENT_BASELINE, NATIVE_COMMENT_ROOT)?.payload
    );
    Ok(())
}

#[test]
fn native_comment_removal_oracles_retain_authors_and_cull_only_final_edge() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    let shared = Package::from_bytes(NATIVE_COMMENT_SHARED_REMOVAL)?;
    let final_removal = Package::from_bytes(NATIVE_COMMENT_FINAL_REMOVAL)?;
    baseline.validate()?;
    shared.validate()?;
    final_removal.validate()?;

    let baseline_edges = native_external_edges(NATIVE_COMMENT_BASELINE)?;
    let shared_edges = native_external_edges(NATIVE_COMMENT_SHARED_REMOVAL)?;
    let final_edges = native_external_edges(NATIVE_COMMENT_FINAL_REMOVAL)?;
    let author_edge = native_comment_author_edge();
    assert_eq!(baseline_edges.len(), 700);
    assert_eq!(shared_edges.len(), 700);
    assert_eq!(final_edges.len(), 699);
    assert!(baseline_edges.contains(&author_edge));
    assert_compact_edge_sets_equal(
        &shared_edges,
        &baseline_edges,
        "native shared-removal raw metadata",
    );
    assert!(!final_edges.contains(&author_edge));
    let baseline_rotated_locators = native_component_effective_locators(NATIVE_COMMENT_BASELINE)?;
    let final_rotated_locators = native_component_effective_locators(NATIVE_COMMENT_FINAL_REMOVAL)?;
    let baseline_rotated = baseline_rotated_locators
        .get(&NATIVE_ROTATED_COMPONENT_BEFORE)
        .ok_or_else(|| io::Error::other("baseline rotated component locator is missing"))?;
    let final_rotated = final_rotated_locators
        .get(&NATIVE_ROTATED_COMPONENT_AFTER)
        .ok_or_else(|| io::Error::other("final rotated component locator is missing"))?;
    assert_eq!(
        baseline_rotated, "ViewState",
        "native baseline rotated component {} for object {}",
        NATIVE_ROTATED_COMPONENT_BEFORE, NATIVE_ROTATED_DOCUMENT_OBJECT
    );
    assert_eq!(
        final_rotated, "ViewState-2654075",
        "native final rotated component {} for object {}",
        NATIVE_ROTATED_COMPONENT_AFTER, NATIVE_ROTATED_DOCUMENT_OBJECT
    );
    assert_eq!(
        baseline_rotated_locators.get(&NATIVE_ROTATED_COMPONENT_AFTER),
        None,
        "baseline unexpectedly contains final rotated component identity"
    );
    assert_eq!(
        final_rotated_locators.get(&NATIVE_ROTATED_COMPONENT_BEFORE),
        None,
        "final unexpectedly retains baseline rotated component identity"
    );
    let removed_rotation_edges = [
        (1, NATIVE_ROTATED_COMPONENT_BEFORE, None),
        (
            NATIVE_ROTATED_COMPONENT_BEFORE,
            1,
            Some(NATIVE_ROTATED_DOCUMENT_OBJECT),
        ),
    ];
    let inserted_rotation_edges = [
        (1, NATIVE_ROTATED_COMPONENT_AFTER, None),
        (
            NATIVE_ROTATED_COMPONENT_AFTER,
            1,
            Some(NATIVE_ROTATED_DOCUMENT_OBJECT),
        ),
    ];
    let mut expected_final_edges = baseline_edges.clone();
    assert!(expected_final_edges.remove(&author_edge));
    for edge in removed_rotation_edges {
        assert!(expected_final_edges.remove(&edge));
    }
    for edge in inserted_rotation_edges {
        assert!(expected_final_edges.insert(edge));
    }
    assert_eq!(final_edges, expected_final_edges);

    assert_eq!(
        native_media_ids(NATIVE_COMMENT_SHARED_REMOVAL)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_COMMENT_MOVIE,
            NATIVE_MOVIE_B,
        ]
    );
    assert_eq!(
        native_commented_media_ids(NATIVE_COMMENT_SHARED_REMOVAL)?,
        vec![NATIVE_COMMENT_MOVIE]
    );
    assert_eq!(
        native_comment_node(NATIVE_COMMENT_SHARED_REMOVAL, NATIVE_COMMENT_ROOT)?.storage_uuid,
        native_comment_node(NATIVE_COMMENT_BASELINE, NATIVE_COMMENT_ROOT)?.storage_uuid
    );
    assert_eq!(
        native_media_ids(NATIVE_COMMENT_FINAL_REMOVAL)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert!(native_commented_media_ids(NATIVE_COMMENT_FINAL_REMOVAL)?.is_empty());
    assert!(native_object_exists(
        NATIVE_COMMENT_SHARED_REMOVAL,
        NATIVE_COMMENT_MOVIE
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_SHARED_REMOVAL,
        NATIVE_COMMENT_MOVIE_CLONE
    )?);
    assert!(native_object_exists(
        NATIVE_COMMENT_SHARED_REMOVAL,
        NATIVE_COMMENT_ROOT
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_SHARED_REMOVAL,
        NATIVE_COMMENT_ROOT_CLONE
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_FINAL_REMOVAL,
        NATIVE_COMMENT_MOVIE
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_FINAL_REMOVAL,
        NATIVE_COMMENT_MOVIE_CLONE
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_FINAL_REMOVAL,
        NATIVE_COMMENT_ROOT
    )?);
    assert!(!native_object_exists(
        NATIVE_COMMENT_FINAL_REMOVAL,
        NATIVE_COMMENT_ROOT_CLONE
    )?);
    let baseline_author_ids = native_annotation_author_ids(NATIVE_COMMENT_BASELINE)?;
    let baseline_author_storage_ids =
        native_annotation_author_storage_ids(NATIVE_COMMENT_BASELINE)?;
    for source in [
        NATIVE_COMMENT_BASELINE,
        NATIVE_COMMENT_SHARED_REMOVAL,
        NATIVE_COMMENT_FINAL_REMOVAL,
    ] {
        assert!(native_object_exists(source, NATIVE_COMMENT_AUTHOR)?);
        assert!(native_object_exists(source, NATIVE_COMMENT_AUTHOR_STORAGE)?);
        assert_eq!(native_annotation_author_ids(source)?, baseline_author_ids);
        assert_eq!(
            native_annotation_author_storage_ids(source)?,
            baseline_author_storage_ids
        );
    }
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_preserves_comment_graph() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    let source = exact_bytes(&source_package)?;
    let commit =
        source_package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(exact_bytes(&source_package)?, source);
    let candidates = native_commented_media_ids(&candidate)?;
    assert_eq!(candidates.len(), 2);
    let clone_movie = candidates
        .into_iter()
        .find(|identifier| *identifier != NATIVE_COMMENT_MOVIE)
        .ok_or_else(|| io::Error::other("focused clone has no direct comment"))?;
    assert_eq!(
        native_media_ids(&candidate)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_COMMENT_MOVIE,
            NATIVE_MOVIE_B,
            clone_movie,
        ]
    );
    assert_comment_graph_clone(&source, &candidate, NATIVE_COMMENT_MOVIE, clone_movie)?;
    assert_eq!(
        native_comment_node(&candidate, NATIVE_COMMENT_ROOT)?.payload,
        native_comment_node(&source, NATIVE_COMMENT_ROOT)?.payload,
        "source comment bytes changed while duplicating media"
    );
    export_if_requested(
        commit.package(),
        "focused-media-comment-duplicate-movie.key",
    )?;

    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn duplicate_and_remove_selected_native_commented_audio_preserves_comment_graph() -> TestResult {
    // This permanent source was produced by the programmatic native probe
    // recorded in ADR 0028.  It already contains Audio A's comment/reply, so
    // this focused test does not need the deprecated host comment mutator to
    // manufacture a source package.
    let source_package = Package::from_bytes(NATIVE_AUDIO_COMMENT_REMOVAL)?;
    let source = exact_bytes(&source_package)?;
    source_package.validate()?;

    let source_audio_ids = native_audio_ids(&source)?;
    assert_eq!(source_audio_ids, vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B]);
    let source_audio_comment_ids = native_commented_audio_ids(&source)?;
    assert_eq!(source_audio_comment_ids, vec![NATIVE_AUDIO_A]);
    assert_audio_comment_shape(&source, NATIVE_AUDIO_A)?;
    let source_movie = native_commented_media_ids(&source)?
        .into_iter()
        .find(|identifier| *identifier != NATIVE_AUDIO_A)
        .ok_or_else(|| io::Error::other("native commented movie is missing"))?;
    assert_eq!(source_movie, NATIVE_COMMENT_MOVIE);
    let source_movie_graph = native_comment_graph(&source, source_movie)?;
    let source_audio_payloads = sorted_media_payloads(&source_package, [0, 1], MediaPart::Content)?;
    let source_audio_a_payload = sorted_media_payloads(&source_package, [0], MediaPart::Content)?;
    let source_movie_payloads = sorted_movie_payloads(&source_package, [2, 3])?;
    let source_edges = native_external_edges(&source)?;
    let source_author_ids = native_annotation_author_ids(&source)?;
    let source_author_storage_ids = native_annotation_author_storage_ids(&source)?;

    let duplicate =
        source_package.duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))?;
    assert!(!duplicate.patch().is_noop());
    assert_eq!(exact_bytes(&source_package)?, source);
    let duplicate_bytes = exact_bytes(duplicate.package())?;
    duplicate.package().validate()?;

    let duplicate_audio_ids = native_audio_ids(&duplicate_bytes)?;
    assert_eq!(duplicate_audio_ids.len(), 3);
    assert_eq!(&duplicate_audio_ids[..2], &source_audio_ids);
    let clone_audio = *duplicate_audio_ids
        .last()
        .ok_or_else(|| io::Error::other("focused audio clone is missing"))?;
    assert_ne!(clone_audio, NATIVE_AUDIO_A);
    assert_eq!(native_commented_audio_ids(&duplicate_bytes)?.len(), 2);
    assert_audio_comment_shape(&duplicate_bytes, NATIVE_AUDIO_A)?;
    assert_audio_comment_shape(&duplicate_bytes, clone_audio)?;
    assert_comment_graph_clone(&source, &duplicate_bytes, NATIVE_AUDIO_A, clone_audio)?;
    assert_comment_graph_semantics(
        &source_movie_graph,
        &native_comment_graph(&duplicate_bytes, source_movie)?,
    );
    assert_eq!(native_external_edges(&duplicate_bytes)?, source_edges);
    assert_eq!(
        native_annotation_author_ids(&duplicate_bytes)?,
        source_author_ids
    );
    assert_eq!(
        native_annotation_author_storage_ids(&duplicate_bytes)?,
        source_author_storage_ids
    );

    let mut expected_duplicate_audio_payloads = source_audio_payloads.clone();
    expected_duplicate_audio_payloads.extend(source_audio_a_payload);
    expected_duplicate_audio_payloads.sort();
    assert_eq!(
        sorted_media_payloads(duplicate.package(), [0, 1, 4], MediaPart::Content)?,
        expected_duplicate_audio_payloads
    );
    assert_eq!(
        sorted_movie_payloads(duplicate.package(), [2, 3])?,
        source_movie_payloads
    );
    export_if_requested(
        duplicate.package(),
        "focused-media-comment-duplicate-audio.key",
    )?;

    // The focused selector addresses the complete source-order media list;
    // duplication appends the clone after the two existing movies.
    let removed = duplicate
        .package()
        .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(4))?;
    assert!(!removed.patch().is_noop());
    assert_eq!(exact_bytes(duplicate.package())?, duplicate_bytes);
    let removed_bytes = exact_bytes(removed.package())?;
    removed.package().validate()?;
    assert_eq!(native_audio_ids(&removed_bytes)?, source_audio_ids);
    assert_eq!(
        native_commented_audio_ids(&removed_bytes)?,
        vec![NATIVE_AUDIO_A]
    );
    assert_audio_comment_shape(&removed_bytes, NATIVE_AUDIO_A)?;
    assert!(!native_object_exists(&removed_bytes, clone_audio)?);
    assert_comment_graph_semantics(
        &source_movie_graph,
        &native_comment_graph(&removed_bytes, source_movie)?,
    );
    assert_eq!(native_external_edges(&removed_bytes)?, source_edges);
    assert_eq!(
        native_annotation_author_ids(&removed_bytes)?,
        source_author_ids
    );
    assert_eq!(
        native_annotation_author_storage_ids(&removed_bytes)?,
        source_author_storage_ids
    );
    assert_eq!(
        sorted_media_payloads(removed.package(), [0, 1], MediaPart::Content)?,
        source_audio_payloads
    );
    assert_eq!(
        sorted_movie_payloads(removed.package(), [2, 3])?,
        source_movie_payloads
    );
    export_if_requested(removed.package(), "focused-media-comment-remove-audio.key")?;

    let restored_duplicate = removed
        .package()
        .apply_slide_media_lifecycle(&removed.patch().inverse())?;
    assert_eq!(exact_bytes(restored_duplicate.package())?, duplicate_bytes);
    let restored_source = restored_duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn duplicate_selected_commented_movie_clones_synthetic_replies() -> TestResult {
    let source = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let source_package = Package::from_bytes(&source)?;
    let before = exact_bytes(&source_package)?;
    let commit =
        source_package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(exact_bytes(&source_package)?, before);
    let clone_movie = native_commented_media_ids(&candidate)?
        .into_iter()
        .find(|identifier| *identifier != NATIVE_COMMENT_MOVIE)
        .ok_or_else(|| io::Error::other("reply clone has no direct comment"))?;
    assert_eq!(
        native_media_ids(&candidate)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_COMMENT_MOVIE,
            NATIVE_MOVIE_B,
            clone_movie,
        ]
    );
    assert_comment_graph_clone(&source, &candidate, NATIVE_COMMENT_MOVIE, clone_movie)?;
    assert_eq!(
        native_comment_node(&candidate, NATIVE_COMMENT_ROOT)?.payload,
        native_comment_node(&source, NATIVE_COMMENT_ROOT)?.payload,
        "source comment bytes changed while duplicating media"
    );
    let source_graph = native_comment_graph(&source, NATIVE_COMMENT_MOVIE)?;
    let candidate_graph = native_comment_graph(&candidate, clone_movie)?;
    assert_eq!(source_graph.nodes.len(), 2);
    assert_eq!(candidate_graph.nodes.len(), 2);
    assert_ne!(
        source_graph.nodes[1].identifier,
        candidate_graph.nodes[1].identifier
    );
    assert_eq!(
        source_graph.nodes[1].storage_uuid,
        candidate_graph.nodes[1].storage_uuid
    );
    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before);
    Ok(())
}

#[test]
fn selected_native_commented_movie_removal_culls_storage_and_reuses_author_until_final_owner()
-> TestResult {
    let source_package = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    let source = exact_bytes(&source_package)?;
    let source_edges = native_external_edges(&source)?;
    let author_edge = native_comment_author_edge();
    assert!(source_edges.contains(&author_edge));
    let source_author_ids = native_annotation_author_ids(&source)?;
    let source_author_storage_ids = native_annotation_author_storage_ids(&source)?;
    assert!(source_author_ids.contains(&NATIVE_COMMENT_AUTHOR));
    assert!(source_author_storage_ids.contains(&NATIVE_COMMENT_AUTHOR_STORAGE));

    let duplicate =
        source_package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let duplicate_bytes = exact_bytes(duplicate.package())?;
    let clone_movie = native_commented_media_ids(&duplicate_bytes)?
        .into_iter()
        .find(|identifier| *identifier != NATIVE_COMMENT_MOVIE)
        .ok_or_else(|| io::Error::other("duplicate has no selected comment clone"))?;
    assert_eq!(native_external_edges(&duplicate_bytes)?, source_edges);
    assert_comment_graph_clone(&source, &duplicate_bytes, NATIVE_COMMENT_MOVIE, clone_movie)?;

    let shared = duplicate
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let shared_bytes = exact_bytes(shared.package())?;
    assert_eq!(
        native_media_ids(&shared_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B, clone_movie]
    );
    assert_eq!(
        native_commented_media_ids(&shared_bytes)?,
        vec![clone_movie]
    );
    assert!(!native_object_exists(&shared_bytes, NATIVE_COMMENT_MOVIE)?);
    assert!(!native_object_exists(&shared_bytes, NATIVE_COMMENT_ROOT)?);
    assert!(native_object_exists(&shared_bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(
        &shared_bytes,
        NATIVE_COMMENT_AUTHOR_STORAGE
    )?);
    assert_eq!(
        native_annotation_author_ids(&shared_bytes)?,
        source_author_ids
    );
    assert_eq!(
        native_annotation_author_storage_ids(&shared_bytes)?,
        source_author_storage_ids
    );
    assert_eq!(native_external_edges(&shared_bytes)?, source_edges);
    assert_comment_graph_clone(&source, &shared_bytes, NATIVE_COMMENT_MOVIE, clone_movie)?;
    export_if_requested(
        shared.package(),
        "focused-media-comment-remove-shared-movie.key",
    )?;

    let final_removal = shared
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(3))?;
    let final_bytes = exact_bytes(final_removal.package())?;
    assert_eq!(
        native_media_ids(&final_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert!(native_commented_media_ids(&final_bytes)?.is_empty());
    assert!(!native_object_exists(&final_bytes, NATIVE_COMMENT_MOVIE)?);
    assert!(!native_object_exists(&final_bytes, clone_movie)?);
    assert!(!native_object_exists(&final_bytes, NATIVE_COMMENT_ROOT)?);
    assert!(!native_object_exists(
        &final_bytes,
        NATIVE_COMMENT_ROOT_CLONE
    )?);
    assert!(native_object_exists(&final_bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(
        &final_bytes,
        NATIVE_COMMENT_AUTHOR_STORAGE
    )?);
    assert_eq!(
        native_annotation_author_ids(&final_bytes)?,
        source_author_ids
    );
    assert_eq!(
        native_annotation_author_storage_ids(&final_bytes)?,
        source_author_storage_ids
    );
    let mut expected_final_edges = source_edges.clone();
    assert!(expected_final_edges.remove(&author_edge));
    assert_eq!(native_external_edges(&final_bytes)?, expected_final_edges);
    export_if_requested(
        final_removal.package(),
        "focused-media-comment-remove-final-movie.key",
    )?;

    let restored_shared = final_removal
        .package()
        .apply_slide_media_lifecycle(&final_removal.patch().inverse())?;
    let restored_duplicate = restored_shared
        .package()
        .apply_slide_media_lifecycle(&shared.patch().inverse())?;
    let restored_source = restored_duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn remove_selected_commented_movie_with_synthetic_reply_graph_is_atomic_and_reversible()
-> TestResult {
    let source = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let source_package = Package::from_bytes(&source)?;
    let before = exact_bytes(&source_package)?;
    let source_graph = native_comment_graph(&source, NATIVE_COMMENT_MOVIE)?;
    let reply = source_graph
        .nodes
        .iter()
        .find(|node| node.identifier == SYNTHETIC_REPLY_ID)
        .ok_or_else(|| io::Error::other("reply fixture has no synthetic reply"))?;

    let removal =
        source_package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let candidate = exact_bytes(removal.package())?;
    assert_eq!(native_commented_media_ids(&candidate)?, Vec::<u64>::new());
    assert_eq!(
        native_media_ids(&candidate)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert!(!native_object_exists(
        &candidate,
        source_graph.root_identifier
    )?);
    assert!(!native_object_exists(&candidate, reply.identifier)?);
    assert!(native_object_exists(&candidate, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(
        &candidate,
        NATIVE_COMMENT_AUTHOR_STORAGE
    )?);
    let mut expected_edges = native_external_edges(&source)?;
    assert!(expected_edges.remove(&native_comment_author_edge()));
    assert_eq!(native_external_edges(&candidate)?, expected_edges);
    let restored = removal
        .package()
        .apply_slide_media_lifecycle(&removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, before);
    Ok(())
}

#[test]
fn remove_shared_comment_root_preserves_replies_until_final_media_owner() -> TestResult {
    let with_reply = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let root_identifier = native_comment_graph(&with_reply, NATIVE_COMMENT_MOVIE)?.root_identifier;
    let source = with_native_shared_movie_comment(&with_reply, NATIVE_MOVIE_B, root_identifier)?;
    let source_package = Package::from_bytes(&source)?;
    let before = exact_bytes(&source_package)?;
    let author_edge = native_comment_author_edge();
    let source_edges = native_external_edges(&source)?;
    assert!(source_edges.contains(&author_edge));
    assert_eq!(
        native_commented_media_ids(&source)?,
        vec![NATIVE_COMMENT_MOVIE, NATIVE_MOVIE_B]
    );
    assert_eq!(
        native_comment_graph(&source, NATIVE_MOVIE_B)?.root_identifier,
        root_identifier
    );

    let first_removal =
        source_package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let first_bytes = exact_bytes(first_removal.package())?;
    assert_eq!(
        native_media_ids(&first_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert_eq!(
        native_commented_media_ids(&first_bytes)?,
        vec![NATIVE_MOVIE_B]
    );
    assert_eq!(
        native_comment_graph(&first_bytes, NATIVE_MOVIE_B)?.root_identifier,
        root_identifier
    );
    assert_eq!(
        native_comment_node(&first_bytes, root_identifier)?.storage_uuid,
        native_comment_node(&source, root_identifier)?.storage_uuid
    );
    assert_eq!(
        native_comment_node(&first_bytes, SYNTHETIC_REPLY_ID)?.storage_uuid,
        native_comment_node(&source, SYNTHETIC_REPLY_ID)?.storage_uuid
    );
    assert!(native_object_exists(&first_bytes, root_identifier)?);
    assert!(native_object_exists(&first_bytes, SYNTHETIC_REPLY_ID)?);
    assert!(native_object_exists(&first_bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(
        &first_bytes,
        NATIVE_COMMENT_AUTHOR_STORAGE
    )?);
    assert_eq!(native_external_edges(&first_bytes)?, source_edges);
    export_if_requested(
        first_removal.package(),
        "focused-media-comment-shared-root-remove-first.key",
    )?;

    let final_removal = first_removal
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let final_bytes = exact_bytes(final_removal.package())?;
    assert_eq!(
        native_media_ids(&final_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B]
    );
    assert!(native_commented_media_ids(&final_bytes)?.is_empty());
    assert!(!native_object_exists(&final_bytes, NATIVE_MOVIE_B)?);
    assert!(!native_object_exists(&final_bytes, root_identifier)?);
    assert!(!native_object_exists(&final_bytes, SYNTHETIC_REPLY_ID)?);
    assert!(native_object_exists(&final_bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(
        &final_bytes,
        NATIVE_COMMENT_AUTHOR_STORAGE
    )?);
    let mut expected_final_edges = source_edges.clone();
    assert!(expected_final_edges.remove(&author_edge));
    assert_eq!(native_external_edges(&final_bytes)?, expected_final_edges);
    export_if_requested(
        final_removal.package(),
        "focused-media-comment-shared-root-remove-final.key",
    )?;

    let restored_first = final_removal
        .package()
        .apply_slide_media_lifecycle(&final_removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored_first.package())?, first_bytes);
    let restored_source = restored_first
        .package()
        .apply_slide_media_lifecycle(&first_removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, before);
    Ok(())
}

#[test]
fn remove_selected_comment_rejects_surviving_movie_payload_header_mismatch_atomically() -> TestResult
{
    let with_reply = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let root_identifier = native_comment_graph(&with_reply, NATIVE_COMMENT_MOVIE)?.root_identifier;
    let attached = with_native_shared_movie_comment(&with_reply, NATIVE_MOVIE_B, root_identifier)?;
    let source =
        without_native_movie_comment_header_reference(&attached, NATIVE_MOVIE_B, root_identifier)?;
    let movie = native_movie_archive(&source, NATIVE_MOVIE_B)?;
    assert_eq!(
        movie
            .super_
            .comment
            .ok_or_else(|| io::Error::other("surviving movie comment payload is missing"))?
            .identifier,
        root_identifier
    );
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn remove_selected_comment_rejects_surviving_movie_payload_missing_target_atomically() -> TestResult
{
    let source = with_native_shared_movie_comment(
        NATIVE_COMMENT_BASELINE,
        NATIVE_MOVIE_B,
        SYNTHETIC_MISSING_MOVIE_COMMENT_ID,
    )?;
    let movie = native_movie_archive(&source, NATIVE_MOVIE_B)?;
    assert_eq!(
        movie
            .super_
            .comment
            .ok_or_else(|| io::Error::other("surviving movie comment payload is missing"))?
            .identifier,
        SYNTHETIC_MISSING_MOVIE_COMMENT_ID
    );
    assert!(!native_object_exists(
        &source,
        SYNTHETIC_MISSING_MOVIE_COMMENT_ID
    )?);
    let (_, archive) = native_component_containing_object(&source, NATIVE_MOVIE_B)?;
    let movie_object = archive
        .object(NATIVE_MOVIE_B)
        .ok_or_else(|| io::Error::other("missing surviving movie object"))?;
    let message_index = movie_object
        .messages
        .iter()
        .position(|message| message.type_ == NATIVE_MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("surviving movie has no movie payload"))?;
    assert!(
        movie_object.archive_info.message_infos[message_index]
            .object_references
            .contains(&SYNTHETIC_MISSING_MOVIE_COMMENT_ID)
    );
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn remove_selected_comment_rejects_unselected_reply_payload_header_mismatch_atomically()
-> TestResult {
    let with_reply = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let source = with_native_unselected_comment_reply_missing_header(&with_reply)?;
    let hidden = native_comment_node(&source, SYNTHETIC_SURVIVING_COMMENT_ID)?;
    assert_eq!(hidden.replies, vec![SYNTHETIC_REPLY_ID]);
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn remove_selected_comment_rejects_unselected_reply_payload_missing_target_atomically() -> TestResult
{
    let with_reply = with_native_reply_graph(NATIVE_COMMENT_BASELINE, false)?;
    let source = with_native_unselected_comment_reply_missing_target(&with_reply)?;
    let hidden = native_comment_node(&source, SYNTHETIC_SURVIVING_COMMENT_ID)?;
    assert_eq!(hidden.replies, vec![SYNTHETIC_MISSING_REPLY_ID]);
    assert!(!native_object_exists(&source, SYNTHETIC_MISSING_REPLY_ID)?);
    let (_, archive) = native_component_containing_object(&source, SYNTHETIC_SURVIVING_COMMENT_ID)?;
    let hidden_object = archive
        .object(SYNTHETIC_SURVIVING_COMMENT_ID)
        .ok_or_else(|| io::Error::other("missing surviving comment storage"))?;
    let message_index = hidden_object
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("surviving comment has no storage payload"))?;
    assert!(
        hidden_object.archive_info.message_infos[message_index]
            .object_references
            .contains(&SYNTHETIC_MISSING_REPLY_ID)
    );
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_reply_cycles_atomically() -> TestResult {
    let source = with_native_comment_cycle(NATIVE_COMMENT_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_malformed_reply_atomically() -> TestResult {
    let source = with_native_malformed_reply(NATIVE_COMMENT_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_unknown_root_reply_reference_atomically()
-> TestResult {
    let source = with_native_unknown_root_reply_reference(NATIVE_COMMENT_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_unknown_archive_info_atomically() -> TestResult
{
    let hidden_reference = tsp::Reference {
        identifier: NATIVE_COMMENT_ROOT,
        ..Default::default()
    }
    .encode_to_vec();
    let source = with_unknown_archive_info_field(
        NATIVE_COMMENT_BASELINE,
        NATIVE_COMMENT_MOVIE,
        99,
        &hidden_reference,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn remove_native_movie_rejects_unknown_surviving_archive_info_atomically() -> TestResult {
    let hidden_reference = tsp::Reference {
        identifier: NATIVE_MOVIE_B,
        ..Default::default()
    }
    .encode_to_vec();
    let source = with_unknown_archive_info_field(
        NATIVE_COMMENT_BASELINE,
        NATIVE_COMMENT_MOVIE,
        99,
        &hidden_reference,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(3)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_unknown_comment_reference_atomically()
-> TestResult {
    let source = with_native_movie_comment_reference_mutation(
        NATIVE_COMMENT_BASELINE,
        MovieCommentReferenceMutation::UnknownField,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_selected_native_commented_movie_rejects_deprecated_comment_reference_atomically()
-> TestResult {
    let source = with_native_movie_comment_reference_mutation(
        NATIVE_COMMENT_BASELINE,
        MovieCommentReferenceMutation::DeprecatedType,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_saved_audio_comment_duplicate_fixture_has_strict_readback() -> TestResult {
    let bytes = read_fixture_override(SAVED_AUDIO_DUPLICATE_ENV, NATIVE_AUDIO_COMMENT_DUPLICATE)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(exact_bytes(&package)?, bytes);

    let source = Package::from_bytes(NATIVE_AUDIO_COMMENT_REMOVAL)?;
    let source_audio_payloads = sorted_media_payloads(&source, [0, 1], MediaPart::Content)?;
    let source_audio_a_payload = sorted_media_payloads(&source, [0], MediaPart::Content)?;
    let source_movie_payloads = sorted_movie_payloads(&source, [2, 3])?;
    assert_eq!(native_audio_ids(&bytes)?.len(), 3);
    let audio_comments = native_commented_audio_ids(&bytes)?;
    assert_eq!(audio_comments.len(), 2);
    assert_audio_comment_shape(&bytes, NATIVE_AUDIO_A)?;
    let clone_audio = audio_comments
        .into_iter()
        .find(|identifier| *identifier != NATIVE_AUDIO_A)
        .ok_or_else(|| io::Error::other("native-saved audio clone has no comment"))?;
    assert_audio_comment_shape(&bytes, clone_audio)?;
    assert_comment_graph_clone(
        NATIVE_AUDIO_COMMENT_REMOVAL,
        &bytes,
        NATIVE_AUDIO_A,
        clone_audio,
    )?;
    assert_eq!(native_commented_media_ids(&bytes)?.len(), 3);
    let movie_graph = native_comment_graph(NATIVE_COMMENT_BASELINE, NATIVE_COMMENT_MOVIE)?;
    assert_comment_graph_semantics(
        &movie_graph,
        &native_comment_graph(&bytes, NATIVE_COMMENT_MOVIE)?,
    );

    let mut expected_audio_payloads = source_audio_payloads;
    expected_audio_payloads.extend(source_audio_a_payload);
    expected_audio_payloads.sort();
    assert_eq!(
        sorted_media_payloads(&package, [0, 1, 4], MediaPart::Content)?,
        expected_audio_payloads
    );
    assert_eq!(
        sorted_movie_payloads(&package, [2, 3])?,
        source_movie_payloads
    );
    Ok(())
}

#[test]
fn native_saved_audio_comment_removal_fixture_has_strict_readback() -> TestResult {
    let bytes = read_fixture_override(SAVED_AUDIO_REMOVAL_ENV, NATIVE_AUDIO_COMMENT_REMOVAL)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(exact_bytes(&package)?, bytes);
    assert_eq!(
        native_audio_ids(&bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B]
    );
    assert_eq!(native_commented_audio_ids(&bytes)?, vec![NATIVE_AUDIO_A]);
    assert_audio_comment_shape(&bytes, NATIVE_AUDIO_A)?;
    assert_eq!(native_commented_media_ids(&bytes)?.len(), 2);
    let source_audio_graph = native_comment_graph(NATIVE_AUDIO_COMMENT_REMOVAL, NATIVE_AUDIO_A)?;
    assert_comment_graph_semantics(
        &source_audio_graph,
        &native_comment_graph(&bytes, NATIVE_AUDIO_A)?,
    );
    let baseline_movie_graph = native_comment_graph(NATIVE_COMMENT_BASELINE, NATIVE_COMMENT_MOVIE)?;
    assert_comment_graph_semantics(
        &baseline_movie_graph,
        &native_comment_graph(&bytes, NATIVE_COMMENT_MOVIE)?,
    );

    let baseline = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    assert_eq!(
        sorted_media_payloads(&package, [0, 1], MediaPart::Content)?,
        sorted_media_payloads(&baseline, [0, 1], MediaPart::Content)?
    );
    assert_eq!(
        sorted_movie_payloads(&package, [2, 3])?,
        sorted_movie_payloads(&baseline, [2, 3])?
    );
    Ok(())
}

#[test]
fn native_saved_media_comment_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_MEDIA_COMMENT_NATIVE_SAVED_PATH") else {
        return Ok(());
    };
    let bytes = fs::read(path)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(native_media_ids(&bytes)?.len(), 5);
    let candidates = native_commented_media_ids(&bytes)?;
    assert_eq!(candidates.len(), 2);
    let clone_movie = candidates
        .into_iter()
        .find(|identifier| *identifier != NATIVE_COMMENT_MOVIE)
        .ok_or_else(|| io::Error::other("native-saved clone has no direct comment"))?;
    assert_comment_graph_clone(
        NATIVE_COMMENT_BASELINE,
        &bytes,
        NATIVE_COMMENT_MOVIE,
        clone_movie,
    )?;
    assert_eq!(exact_bytes(&package)?, bytes);
    Ok(())
}

#[test]
fn native_saved_media_comment_shared_removal_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_MEDIA_COMMENT_NATIVE_SAVED_SHARED_REMOVAL_PATH")
    else {
        return Ok(());
    };
    let bytes = fs::read(path)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    let media_ids = native_media_ids(&bytes)?;
    assert_eq!(media_ids.len(), 4);
    assert_eq!(
        &media_ids[..3],
        &[NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    let commented = native_commented_media_ids(&bytes)?;
    assert_eq!(commented.len(), 1);
    let clone_movie = commented[0];
    assert_ne!(clone_movie, NATIVE_COMMENT_MOVIE);
    assert_eq!(media_ids[3], clone_movie);
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_MOVIE)?);
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_ROOT)?);
    assert_comment_graph_clone(
        NATIVE_COMMENT_BASELINE,
        &bytes,
        NATIVE_COMMENT_MOVIE,
        clone_movie,
    )?;
    assert!(native_object_exists(&bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(&bytes, NATIVE_COMMENT_AUTHOR_STORAGE)?);
    assert_eq!(
        native_annotation_author_ids(&bytes)?,
        native_annotation_author_ids(NATIVE_COMMENT_BASELINE)?
    );
    assert_eq!(
        native_annotation_author_storage_ids(&bytes)?,
        native_annotation_author_storage_ids(NATIVE_COMMENT_BASELINE)?
    );
    let baseline_edges = native_external_edges(NATIVE_COMMENT_BASELINE)?;
    let saved_edges = native_external_edges(&bytes)?;
    assert_eq!(saved_edges.len(), 700);
    assert!(saved_edges.contains(&native_comment_author_edge()));
    assert_compact_edge_sets_equal(
        &native_semantic_external_edges(&bytes)?,
        &native_semantic_external_edges(NATIVE_COMMENT_BASELINE)?,
        "native saved shared-removal semantic metadata",
    );
    assert_eq!(baseline_edges.len(), 700);
    assert_eq!(exact_bytes(&package)?, bytes);
    Ok(())
}

#[test]
fn native_saved_media_comment_removal_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_MEDIA_COMMENT_NATIVE_SAVED_REMOVAL_PATH") else {
        return Ok(());
    };
    let bytes = fs::read(path)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(
        native_media_ids(&bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert!(native_commented_media_ids(&bytes)?.is_empty());
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_MOVIE)?);
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_MOVIE_CLONE)?);
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_ROOT)?);
    assert!(!native_object_exists(&bytes, NATIVE_COMMENT_ROOT_CLONE)?);
    assert!(native_object_exists(&bytes, NATIVE_COMMENT_AUTHOR)?);
    assert!(native_object_exists(&bytes, NATIVE_COMMENT_AUTHOR_STORAGE)?);
    assert_eq!(
        native_annotation_author_ids(&bytes)?,
        native_annotation_author_ids(NATIVE_COMMENT_BASELINE)?
    );
    assert_eq!(
        native_annotation_author_storage_ids(&bytes)?,
        native_annotation_author_storage_ids(NATIVE_COMMENT_BASELINE)?
    );
    let saved_edges = native_external_edges(&bytes)?;
    assert_eq!(saved_edges.len(), 699);
    assert!(!saved_edges.contains(&native_comment_author_edge()));
    let baseline_semantic_edges = native_semantic_external_edges(NATIVE_COMMENT_BASELINE)?;
    let mut expected_semantic_edges = baseline_semantic_edges.clone();
    let author_semantic_edge = native_semantic_comment_author_edge(NATIVE_COMMENT_BASELINE)?;
    assert!(expected_semantic_edges.remove(&author_semantic_edge));
    let saved_semantic_edges = native_semantic_external_edges(&bytes)?;
    assert_compact_edge_sets_equal(
        &saved_semantic_edges,
        &expected_semantic_edges,
        "native saved final-removal semantic metadata",
    );
    assert_eq!(exact_bytes(&package)?, bytes);
    Ok(())
}
