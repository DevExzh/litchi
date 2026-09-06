//! Focused Keynote media lifecycle coverage for direct drawable comments.
//!
//! The native fixtures in this test deliberately keep comment storage in the
//! slide component while the annotation author lives in the shared author
//! component.  The lifecycle owner must clone every comment-storage node,
//! preserve each storage UUID, and continue sharing the author object.
//! Removal remains a compatibility guard until comment metadata ownership is
//! implemented by the focused owner.

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
use litchi_keynote::{MovieSelector, Package, SlideMediaLifecycleError, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_COMMENT_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_COMMENT_DUPLICATE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-duplicate-native.key");

const NATIVE_SLIDE: u64 = 2_652_150;
const NATIVE_AUDIO_A: u64 = 2_652_595;
const NATIVE_AUDIO_B: u64 = 2_652_622;
const NATIVE_MOVIE_B: u64 = 2_653_610;
const NATIVE_COMMENT_MOVIE: u64 = 2_653_286;
const NATIVE_COMMENT_MOVIE_CLONE: u64 = 2_653_814;
const NATIVE_COMMENT_ROOT: u64 = 2_653_723;
const NATIVE_COMMENT_ROOT_CLONE: u64 = 2_653_826;
const NATIVE_COMMENT_AUTHOR: u64 = 2_653_721;
const NATIVE_SLIDE_MESSAGE_TYPE: u32 = 5;
const NATIVE_MOVIE_MESSAGE_TYPE: u32 = 3_007;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const SYNTHETIC_REPLY_ID: u64 = 2_653_900;
const SYNTHETIC_REPLY_UUID: tsp::Uuid = tsp::Uuid {
    lower: 0x0bad_cafe_dead_beef,
    upper: 0x0123_4567_89ab_cdef,
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
    assert!(source_author_ids.contains(&NATIVE_COMMENT_AUTHOR));
    assert_eq!(
        native_annotation_author_ids(candidate)?,
        native_annotation_author_ids(source)?,
        "comment duplication must reuse annotation-author objects"
    );
    Ok(())
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
fn duplicate_selected_native_commented_movie_clones_replies() -> TestResult {
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
fn selected_native_commented_movie_removal_remains_unsupported_atomically() -> TestResult {
    let package = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2)),
        Err(SlideMediaLifecycleError::UnsupportedComment)
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
