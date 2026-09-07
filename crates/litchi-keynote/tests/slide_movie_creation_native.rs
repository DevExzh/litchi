//! Native Keynote acceptance coverage for a freshly inserted file movie.
//!
//! The source fixture was appended, saved, actually closed, and reopened in
//! Keynote 14.4.  The focused package assertions remain selector-first; the
//! generated protobufs below are a fixture oracle for the five-object movie
//! closure and its movie-start build.

use std::{
    collections::{BTreeMap, BTreeSet},
    fmt, io,
    time::Duration,
};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd};
use litchi_keynote::{MediaPart, MovieKind, MovieSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/slide-movie-creation-source-native.key");
const NATIVE_FOCUSED_REUSE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/slide-movie-creation-focused-native.key");
const NATIVE_FOCUSED_FRESH: &[u8] = include_bytes!(
    "../../../test-data/iwork/keynote/slide-movie-creation-fresh-focused-native.key"
);
const SOURCE_MOVIE_POSITION: usize = 2;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn with_context<T, E>(result: Result<T, E>, context: impl fmt::Display) -> TestResult<T>
where
    E: fmt::Display,
{
    Ok(result.map_err(|error| io::Error::other(format!("{context}: {error}")))?)
}

fn source_assets(source: &Package) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let movie = source
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Content,
        )?
        .to_vec();
    let poster = source
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Poster,
        )?
        .to_vec();
    if movie.is_empty() || poster.is_empty() {
        return Err(io::Error::other("native source movie assets are empty").into());
    }
    Ok((movie, poster))
}

fn fresh_movie(source: &[u8]) -> Vec<u8> {
    let mut movie = source.to_vec();
    movie.extend_from_slice(&16u32.to_be_bytes());
    movie.extend_from_slice(b"free");
    movie.extend_from_slice(b"litchi!!");
    movie
}

fn png_chunk(kind: &[u8; 4], payload: &[u8]) -> Vec<u8> {
    let length = u32::try_from(payload.len()).expect("test PNG chunk fits u32");
    let mut crc_input = Vec::with_capacity(kind.len() + payload.len());
    crc_input.extend_from_slice(kind);
    crc_input.extend_from_slice(payload);
    let mut crc = 0xffff_ffffu32;
    for byte in crc_input {
        crc ^= u32::from(byte);
        for _ in 0..8 {
            let mask = 0u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    let mut output = Vec::with_capacity(12 + payload.len());
    output.extend_from_slice(&length.to_be_bytes());
    output.extend_from_slice(kind);
    output.extend_from_slice(payload);
    output.extend_from_slice(&(!crc).to_be_bytes());
    output
}

fn fresh_poster(source: &[u8]) -> TestResult<Vec<u8>> {
    let type_offset = source
        .windows(4)
        .position(|window| window == b"IEND")
        .ok_or_else(|| io::Error::other("native source poster has no IEND chunk"))?;
    let chunk_start = type_offset
        .checked_sub(4)
        .ok_or_else(|| io::Error::other("native source poster has a truncated IEND chunk"))?;
    let mut poster = Vec::with_capacity(source.len() + 32);
    poster.extend_from_slice(&source[..chunk_start]);
    poster.extend_from_slice(&png_chunk(b"tEXt", b"Comment\0litchi fresh poster"));
    poster.extend_from_slice(&source[chunk_start..]);
    Ok(poster)
}

fn native_archives(source: &[u8]) -> TestResult<Vec<Archive>> {
    Catalog::from_bytes(source)?
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
        .map(|entry| {
            Ok(Archive::parse(
                &SnappyStream::decompress(entry.data())?.into_bytes(),
            )?)
        })
        .collect()
}

fn native_object_message<T>(source: &[u8], identifier: u64, message_type: u32) -> TestResult<T>
where
    T: prost::Message + Default,
{
    for archive in native_archives(source)? {
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let Some(message) = object
            .messages
            .iter()
            .find(|message| message.type_ == message_type)
        else {
            continue;
        };
        return Ok(T::decode(message.data.as_slice())?);
    }
    Err(io::Error::other(format!(
        "native object {identifier} message {message_type} is missing"
    ))
    .into())
}

fn native_slide(source: &[u8]) -> TestResult<(u64, kn::SlideArchive)> {
    // Document -> Show -> first SlideTree entry -> SlideNode -> Slide is the
    // presentation topology.  Searching for any media-bearing archive also
    // finds layout/master slide records in real Keynote packages.
    let document: kn::DocumentArchive = native_object_message(source, 1, 1)?;
    let show: kn::ShowArchive = native_object_message(source, document.show.identifier, 2)?;
    let node_reference = show
        .slide_tree
        .slides
        .first()
        .ok_or_else(|| io::Error::other("native presentation slide tree is empty"))?;
    let node: kn::SlideNodeArchive =
        native_object_message(source, node_reference.identifier, SLIDE_NODE_MESSAGE_TYPE)?;
    let slide_reference = node
        .slide
        .ok_or_else(|| io::Error::other("native presentation slide node has no slide"))?;
    let slide: kn::SlideArchive =
        native_object_message(source, slide_reference.identifier, SLIDE_MESSAGE_TYPE)?;
    Ok((slide_reference.identifier, slide))
}

fn native_slide_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let (slide_identifier, _) = native_slide(source)?;
    let archive = native_archives(source)?
        .into_iter()
        .find(|archive| archive.object(slide_identifier).is_some())
        .ok_or_else(|| io::Error::other("native slide component is missing"))?;
    Ok(archive
        .objects
        .into_iter()
        .filter_map(|object| object.archive_info.identifier)
        .collect())
}

fn native_movie(source: &[u8], identifier: u64) -> TestResult<tsd::MovieArchive> {
    for archive in native_archives(source)? {
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let Some(message) = object
            .messages
            .iter()
            .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        else {
            continue;
        };
        return Ok(tsd::MovieArchive::decode(message.data.as_slice())?);
    }
    Err(io::Error::other("native movie object is missing").into())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NativeObjectGraph {
    reference_identifier: u64,
    messages: Vec<(u32, Vec<u8>)>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NativeStandinKind {
    EmptyStandin,
    CaptionInfo,
    Other,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NativeStandinGraph {
    reference_identifier: u64,
    kind: NativeStandinKind,
    messages: Vec<(u32, Vec<u8>)>,
    storage: Option<NativeObjectGraph>,
}

fn native_object_graph(source: &[u8], reference_identifier: u64) -> TestResult<NativeObjectGraph> {
    let object = native_archives(source)?
        .into_iter()
        .find_map(|archive| {
            archive
                .objects
                .into_iter()
                .find(|object| object.archive_info.identifier == Some(reference_identifier))
        })
        .ok_or_else(|| {
            io::Error::other(format!("native object {reference_identifier} is missing"))
        })?;
    Ok(NativeObjectGraph {
        reference_identifier,
        messages: object
            .messages
            .into_iter()
            .map(|message| (message.type_, message.data))
            .collect(),
    })
}

#[allow(
    deprecated,
    reason = "The native oracle preserves both current and legacy caption storage edges."
)]
fn native_caption_storage_identifier(caption: &tsa::CaptionInfoArchive) -> Option<u64> {
    caption
        .super_
        .owned_storage
        .as_ref()
        .or(caption.super_.deprecated_storage.as_ref())
        .map(|reference| reference.identifier)
}

fn native_standin_graph(
    source: &[u8],
    reference_identifier: Option<u64>,
) -> TestResult<Option<NativeStandinGraph>> {
    let Some(reference_identifier) = reference_identifier else {
        return Ok(None);
    };
    let object = native_object_graph(source, reference_identifier)?;

    // Native audio title/caption edges use an empty type-3097 stand-in,
    // while a populated file-movie caption can point at type 633 and its
    // type-2001 text storage. Preserve both forms as exact object messages so
    // this oracle checks the complete native graph, including caption text,
    // without assuming every edge has a semantic payload.
    let storage_identifier = object
        .messages
        .iter()
        .find(|(message_type, _)| *message_type == CAPTION_INFO_MESSAGE_TYPE)
        .map(|(_, data)| tsa::CaptionInfoArchive::decode(data.as_slice()))
        .transpose()?
        .and_then(|caption| native_caption_storage_identifier(&caption));
    let storage = storage_identifier
        .map(|identifier| native_object_graph(source, identifier))
        .transpose()?;
    if let Some(storage) = storage.as_ref()
        && !storage
            .messages
            .iter()
            .any(|(message_type, _)| *message_type == STORAGE_MESSAGE_TYPE)
    {
        return Err(io::Error::other(format!(
            "native caption storage {storage_identifier:?} has no type-{STORAGE_MESSAGE_TYPE} message"
        ))
        .into());
    }
    let kind = if object.messages.len() == 1
        && object.messages[0].0 == STANDIN_MESSAGE_TYPE
        && object.messages[0].1.is_empty()
    {
        NativeStandinKind::EmptyStandin
    } else if object
        .messages
        .iter()
        .any(|(message_type, _)| *message_type == CAPTION_INFO_MESSAGE_TYPE)
    {
        NativeStandinKind::CaptionInfo
    } else {
        NativeStandinKind::Other
    };

    Ok(Some(NativeStandinGraph {
        reference_identifier,
        kind,
        messages: object.messages,
        storage,
    }))
}

fn native_movie_standin_graphs(
    source: &[u8],
    movie_position: usize,
) -> TestResult<(Option<NativeStandinGraph>, Option<NativeStandinGraph>)> {
    let movie_identifier = *native_media_ids(source)?
        .get(movie_position)
        .ok_or_else(|| io::Error::other("native package has no selected media"))?;
    let movie = native_movie(source, movie_identifier)?;
    Ok((
        native_standin_graph(
            source,
            movie
                .super_
                .title
                .as_ref()
                .map(|reference| reference.identifier),
        )?,
        native_standin_graph(
            source,
            movie
                .super_
                .caption
                .as_ref()
                .map(|reference| reference.identifier),
        )?,
    ))
}

fn native_file_movie_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    let (slide_identifier, slide) = native_slide(source)?;
    let mut identifiers = Vec::new();
    for reference in slide.owned_drawables {
        let Some(movie) = archives.iter().find_map(|archive| {
            archive.object(reference.identifier).and_then(|object| {
                object.messages.iter().find_map(|message| {
                    (message.type_ == MOVIE_MESSAGE_TYPE)
                        .then(|| tsd::MovieArchive::decode(message.data.as_slice()).ok())
                        .flatten()
                })
            })
        }) else {
            continue;
        };
        if movie
            .super_
            .parent
            .as_ref()
            .is_some_and(|parent| parent.identifier == slide_identifier)
            && movie.audio_only != Some(true)
        {
            identifiers.push(reference.identifier);
        }
    }
    Ok(identifiers)
}

fn native_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    let (slide_identifier, slide) = native_slide(source)?;
    let mut identifiers = Vec::new();
    for reference in slide.owned_drawables {
        let Some(movie) = archives.iter().find_map(|archive| {
            archive.object(reference.identifier).and_then(|object| {
                object.messages.iter().find_map(|message| {
                    (message.type_ == MOVIE_MESSAGE_TYPE)
                        .then(|| tsd::MovieArchive::decode(message.data.as_slice()).ok())
                        .flatten()
                })
            })
        }) else {
            continue;
        };
        if movie
            .super_
            .parent
            .as_ref()
            .is_some_and(|parent| parent.identifier == slide_identifier)
        {
            identifiers.push(reference.identifier);
        }
    }
    Ok(identifiers)
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
enum CommentIdentity {
    StorageUuid { lower: u64, upper: u64 },
    Identifier(u64),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentSignature {
    identity: CommentIdentity,
    text: String,
    has_author: bool,
    replies: Vec<CommentIdentity>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentGraphSignature {
    root: CommentIdentity,
    nodes: Vec<CommentSignature>,
}

fn native_comment_graph_signature(
    source: &[u8],
    movie_position: usize,
) -> TestResult<Option<CommentGraphSignature>> {
    const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
    let archives = native_archives(source)?;
    let media_identifier = *native_media_ids(source)?
        .get(movie_position)
        .ok_or_else(|| io::Error::other("native package has no selected media"))?;
    let movie = native_movie(source, media_identifier)?;
    let Some(root_identifier) = movie.super_.comment.map(|reference| reference.identifier) else {
        return Ok(None);
    };

    let mut pending = vec![root_identifier];
    let mut nodes = BTreeMap::new();
    while let Some(identifier) = pending.pop() {
        if nodes.contains_key(&identifier) {
            continue;
        }
        let object = archives
            .iter()
            .find_map(|archive| archive.object(identifier))
            .ok_or_else(|| io::Error::other("native comment object is missing"))?;
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("native comment storage payload is missing"))?;
        let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
        let identity = comment
            .storage_uuid
            .as_ref()
            .map(|uuid| CommentIdentity::StorageUuid {
                lower: uuid.lower,
                upper: uuid.upper,
            })
            .unwrap_or(CommentIdentity::Identifier(identifier));
        let text = comment
            .text
            .ok_or_else(|| io::Error::other("native comment text is missing"))?;
        let replies = comment
            .replies
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        pending.extend(replies.iter().rev().copied());
        nodes.insert(
            identifier,
            (identity, text, comment.author.is_some(), replies),
        );
    }

    let root = nodes
        .get(&root_identifier)
        .map(|node| node.0.clone())
        .ok_or_else(|| io::Error::other("native comment root is missing"))?;
    let identities = nodes
        .iter()
        .map(|(identifier, (identity, _, _, _))| (*identifier, identity.clone()))
        .collect::<BTreeMap<_, _>>();
    let mut signatures = nodes
        .into_iter()
        .map(|(_identifier, (identity, text, has_author, replies))| {
            let replies = replies
                .into_iter()
                .map(|reply| {
                    identities
                        .get(&reply)
                        .cloned()
                        .ok_or_else(|| io::Error::other("native comment reply is missing"))
                })
                .collect::<Result<Vec<_>, _>>()?;
            Ok(CommentSignature {
                identity,
                text,
                has_author,
                replies,
            })
        })
        .collect::<TestResult<Vec<_>>>()?;
    signatures.sort_by(|left, right| left.identity.cmp(&right.identity));
    Ok(Some(CommentGraphSignature {
        root,
        nodes: signatures,
    }))
}

fn assert_existing_media_unchanged(
    baseline: &Package,
    focused: &Package,
    baseline_movies: &[litchi_keynote::slide::media::MovieInfo],
) -> TestResult {
    let baseline_bytes = exact_bytes(baseline)?;
    let focused_bytes = exact_bytes(focused)?;
    for (index, baseline_movie) in baseline_movies.iter().enumerate() {
        let focused_movie = focused
            .slides()?
            .first()
            .ok_or_else(|| io::Error::other("focused slide is missing"))?
            .movies()
            .get(index)
            .ok_or_else(|| io::Error::other("focused sibling media is missing"))?;
        assert_eq!(
            focused_movie.kind(),
            baseline_movie.kind(),
            "media kind {index}"
        );
        assert_eq!(
            focused_movie.position(),
            baseline_movie.position(),
            "media position {index}"
        );
        assert_eq!(
            focused_movie.size(),
            baseline_movie.size(),
            "media size {index}"
        );
        assert_eq!(
            focused_movie.original_size(),
            baseline_movie.original_size(),
            "original media size {index}"
        );
        assert_eq!(
            focused_movie.natural_size(),
            baseline_movie.natural_size(),
            "natural media size {index}"
        );
        assert_eq!(
            focused_movie.playback(),
            baseline_movie.playback(),
            "playback controls {index}"
        );
        assert_eq!(
            with_context(
                focused
                    .slide_media_properties(SlideSelector::index(0), MovieSelector::index(index),),
                format!("focused media properties at source position {index}"),
            )?,
            with_context(
                baseline
                    .slide_media_properties(SlideSelector::index(0), MovieSelector::index(index),),
                format!("baseline media properties at source position {index}"),
            )?,
            "media properties {index}"
        );
        // Compare native title/caption stand-ins independently of the public
        // editor's graph admission. Audio uses empty type-3097 stand-ins;
        // file-movie captions can carry a populated type-633 graph. Exact
        // message payloads preserve both representations, including text.
        assert_eq!(
            with_context(
                native_movie_standin_graphs(&focused_bytes, index),
                format!("focused title/caption graph at source position {index}"),
            )?,
            with_context(
                native_movie_standin_graphs(&baseline_bytes, index),
                format!("baseline title/caption graph at source position {index}"),
            )?,
            "title/caption graph {index}"
        );
        assert_eq!(
            with_context(
                native_comment_graph_signature(&focused_bytes, index),
                format!("focused comment graph at source position {index}"),
            )?,
            with_context(
                native_comment_graph_signature(&baseline_bytes, index),
                format!("baseline comment graph at source position {index}"),
            )?,
            "comment/reply graph {index}"
        );
        assert_eq!(
            with_context(
                focused.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    MediaPart::Content,
                ),
                format!("focused media content at source position {index}"),
            )?,
            with_context(
                baseline.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    MediaPart::Content,
                ),
                format!("baseline media content at source position {index}"),
            )?,
            "media content {index}"
        );
        if !baseline_movie.is_audio() {
            assert_eq!(
                with_context(
                    focused.slide_media_data(
                        SlideSelector::index(0),
                        MovieSelector::index(index),
                        MediaPart::Poster,
                    ),
                    format!("focused media poster at source position {index}"),
                )?,
                with_context(
                    baseline.slide_media_data(
                        SlideSelector::index(0),
                        MovieSelector::index(index),
                        MediaPart::Poster,
                    ),
                    format!("baseline media poster at source position {index}"),
                )?,
                "media poster {index}"
            );
        }
    }
    Ok(())
}

fn assert_native_candidate(
    focused_source: &[u8],
    movie_data: &[u8],
    poster_data: &[u8],
    position: (f32, f32),
    size: (f32, f32),
    natural_size: (f32, f32),
    duration: Duration,
) -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let focused = Package::from_bytes(focused_source)?;
    baseline.validate()?;
    focused.validate()?;
    let baseline_movies = baseline
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("baseline slide is missing"))?
        .movies()
        .to_vec();
    let focused_movies = focused
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("focused slide is missing"))?
        .movies();
    assert_eq!(baseline_movies.len(), 4);
    assert_eq!(focused_movies.len(), baseline_movies.len() + 1);
    assert_existing_media_unchanged(&baseline, &focused, &baseline_movies)?;

    let created = focused_movies
        .last()
        .ok_or_else(|| io::Error::other("focused movie is missing"))?;
    assert_eq!(created.kind(), MovieKind::File);
    assert_eq!(
        created.position().map(|point| (point.x, point.y)),
        Some(position)
    );
    assert_eq!(
        created.size().map(|value| (value.width, value.height)),
        Some(size)
    );
    assert_eq!(
        created
            .natural_size()
            .map(|value| (value.width, value.height)),
        Some(natural_size)
    );
    assert_eq!(
        created
            .original_size()
            .map(|value| (value.width, value.height)),
        Some(natural_size)
    );
    assert_eq!(created.duration(), Some(duration));
    let created_position = baseline_movies.len();
    assert_eq!(
        focused.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(created_position),
            MediaPart::Content,
        )?,
        movie_data
    );
    assert_eq!(
        focused.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(created_position),
            MediaPart::Poster,
        )?,
        poster_data
    );

    let baseline_movie_ids = native_file_movie_ids(NATIVE_BASELINE)?;
    let focused_movie_ids = native_file_movie_ids(focused_source)?;
    assert_eq!(baseline_movie_ids.len(), 2);
    assert_eq!(focused_movie_ids.len(), baseline_movie_ids.len() + 1);
    let new_movie_ids = focused_movie_ids
        .iter()
        .copied()
        .filter(|identifier| !baseline_movie_ids.contains(identifier))
        .collect::<Vec<_>>();
    assert_eq!(new_movie_ids.len(), 1);
    let new_movie_id = new_movie_ids[0];
    let movie = native_movie(focused_source, new_movie_id)?;
    assert!(movie.movie_data.is_some());
    assert!(movie.poster_image_data.is_some());
    assert_ne!(movie.audio_only, Some(true));
    let slide_identifier = native_slide(focused_source)?.0;
    assert_eq!(
        movie.super_.parent.as_ref().map(|parent| parent.identifier),
        Some(slide_identifier)
    );
    let title_id = movie.super_.title.as_ref().map(|title| title.identifier);
    let caption_id = movie
        .super_
        .caption
        .as_ref()
        .map(|caption| caption.identifier);
    assert!(title_id.is_some(), "native movie title stand-in is missing");
    assert!(
        caption_id.is_some(),
        "native movie caption stand-in is missing"
    );

    let (_, baseline_slide) = native_slide(NATIVE_BASELINE)?;
    let (_, focused_slide) = native_slide(focused_source)?;
    let baseline_builds = baseline_slide
        .builds
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let focused_builds = focused_slide
        .builds
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let new_builds = focused_builds
        .difference(&baseline_builds)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(new_builds.len(), 1);
    let new_build_id = new_builds[0];
    let build: kn::BuildArchive =
        native_object_message(focused_source, new_build_id, BUILD_MESSAGE_TYPE)?;
    assert_eq!(
        build.drawable.as_ref().map(|drawable| drawable.identifier),
        Some(new_movie_id)
    );
    assert_eq!(
        build
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|animation| animation.effect.as_deref()),
        Some("apple:movie-start")
    );

    let baseline_chunks = baseline_slide
        .build_chunks
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let focused_chunks = focused_slide
        .build_chunks
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let new_chunks = focused_chunks
        .difference(&baseline_chunks)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(new_chunks.len(), 1);
    let new_chunk_id = new_chunks[0];
    let chunk: kn::BuildChunkArchive =
        native_object_message(focused_source, new_chunk_id, BUILD_CHUNK_MESSAGE_TYPE)?;
    assert_eq!(
        chunk.build.as_ref().map(|reference| reference.identifier),
        Some(new_build_id)
    );

    let baseline_ids = native_slide_object_ids(NATIVE_BASELINE)?;
    let focused_ids = native_slide_object_ids(focused_source)?;
    let new_ids = focused_ids
        .difference(&baseline_ids)
        .copied()
        .collect::<BTreeSet<_>>();
    assert_eq!(new_ids.len(), 5);
    let closure = [
        new_movie_id,
        title_id.ok_or_else(|| io::Error::other("native title ref disappeared"))?,
        caption_id.ok_or_else(|| io::Error::other("native caption ref disappeared"))?,
        new_build_id,
        new_chunk_id,
    ]
    .into_iter()
    .collect::<BTreeSet<_>>();
    assert_eq!(closure, new_ids);
    Ok(())
}

fn assert_replay_and_inverse(
    source_bytes: &[u8],
    movie_data: &[u8],
    poster_data: &[u8],
    options: litchi_keynote::slide::movie::Options,
) -> TestResult {
    let source = Package::from_bytes(source_bytes)?;
    let source_snapshot = exact_bytes(&source)?;
    let commit = with_context(
        source.add_slide_movie(
            SlideSelector::index(0),
            "native-replay-movie.mov",
            movie_data,
            "native-replay-poster.png",
            poster_data,
            options,
        ),
        "create another movie from native saved fixture",
    )?;
    let candidate = exact_bytes(commit.package())?;
    let replay = with_context(
        source.apply_slide_movie_creation(commit.patch()),
        "replay creation from native saved fixture",
    )?;
    assert_eq!(exact_bytes(replay.package())?, candidate);
    let restored = with_context(
        commit
            .package()
            .apply_slide_movie_creation(&commit.patch().inverse()),
        "inverse creation from native saved fixture",
    )?;
    assert_eq!(exact_bytes(restored.package())?, source_snapshot);
    Ok(())
}

#[test]
fn saved_native_source_contains_one_reused_file_movie_and_complete_build_closure() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let focused = Package::from_bytes(NATIVE_SOURCE)?;
    baseline.validate()?;
    focused.validate()?;

    let baseline_movies = baseline
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("baseline slide is missing"))?
        .movies()
        .to_vec();
    let focused_movies = focused
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("focused slide is missing"))?
        .movies();
    assert_eq!(baseline_movies.len(), 4);
    assert_eq!(focused_movies.len(), baseline_movies.len() + 1);
    assert_existing_media_unchanged(&baseline, &focused, &baseline_movies)?;

    let created = focused_movies
        .last()
        .ok_or_else(|| io::Error::other("focused movie is missing"))?;
    assert_eq!(created.kind(), MovieKind::File);
    assert_eq!(
        created.size().map(|size| (size.width, size.height)),
        Some((320.0, 180.0))
    );
    assert_eq!(
        created.natural_size().map(|size| (size.width, size.height)),
        Some((320.0, 180.0))
    );
    assert_eq!(created.duration(), Some(Duration::from_secs(2)));
    assert_eq!(
        focused.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(baseline_movies.len()),
            MediaPart::Content,
        )?,
        baseline.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Content,
        )?
    );
    assert_eq!(
        focused.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(baseline_movies.len()),
            MediaPart::Poster,
        )?,
        baseline.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Poster,
        )?
    );

    let baseline_movie_ids = native_file_movie_ids(NATIVE_BASELINE)?;
    let focused_movie_ids = native_file_movie_ids(NATIVE_SOURCE)?;
    assert_eq!(baseline_movie_ids.len(), 2);
    assert_eq!(focused_movie_ids.len(), baseline_movie_ids.len() + 1);
    let new_movie_ids = focused_movie_ids
        .iter()
        .copied()
        .filter(|identifier| !baseline_movie_ids.contains(identifier))
        .collect::<Vec<_>>();
    assert_eq!(new_movie_ids.len(), 1);
    let new_movie_id = new_movie_ids[0];
    let movie = native_movie(NATIVE_SOURCE, new_movie_id)?;
    assert!(movie.movie_data.is_some());
    assert!(movie.poster_image_data.is_some());
    assert_ne!(movie.audio_only, Some(true));
    let slide_identifier = native_slide(NATIVE_SOURCE)?.0;
    assert_eq!(
        movie.super_.parent.as_ref().map(|parent| parent.identifier),
        Some(slide_identifier)
    );
    let title_id = movie.super_.title.as_ref().map(|title| title.identifier);
    let caption_id = movie
        .super_
        .caption
        .as_ref()
        .map(|caption| caption.identifier);
    assert!(title_id.is_some(), "native movie title stand-in is missing");
    assert!(
        caption_id.is_some(),
        "native movie caption stand-in is missing"
    );

    let (_, baseline_slide) = native_slide(NATIVE_BASELINE)?;
    let (_, focused_slide) = native_slide(NATIVE_SOURCE)?;
    let baseline_builds = baseline_slide
        .builds
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let focused_builds = focused_slide
        .builds
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let new_builds = focused_builds
        .difference(&baseline_builds)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(new_builds.len(), 1);
    let new_build_id = new_builds[0];
    let build: kn::BuildArchive =
        native_object_message(NATIVE_SOURCE, new_build_id, BUILD_MESSAGE_TYPE)?;
    assert_eq!(
        build.drawable.as_ref().map(|drawable| drawable.identifier),
        Some(new_movie_id)
    );
    assert_eq!(
        build
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|animation| animation.effect.as_deref()),
        Some("apple:movie-start")
    );

    let baseline_chunks = baseline_slide
        .build_chunks
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let focused_chunks = focused_slide
        .build_chunks
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();
    let new_chunks = focused_chunks
        .difference(&baseline_chunks)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(new_chunks.len(), 1);
    let new_chunk_id = new_chunks[0];
    let chunk: kn::BuildChunkArchive =
        native_object_message(NATIVE_SOURCE, new_chunk_id, BUILD_CHUNK_MESSAGE_TYPE)?;
    assert_eq!(
        chunk.build.as_ref().map(|reference| reference.identifier),
        Some(new_build_id)
    );

    // Native saves also author separate ViewState and related metadata objects.
    // The creation closure is exactly the objects appended to this slide's
    // component, whose identity is resolved from the presentation topology.
    let baseline_ids = native_slide_object_ids(NATIVE_BASELINE)?;
    let focused_ids = native_slide_object_ids(NATIVE_SOURCE)?;
    let new_ids = focused_ids
        .difference(&baseline_ids)
        .copied()
        .collect::<BTreeSet<_>>();
    assert_eq!(new_ids.len(), 5);
    let closure = [
        new_movie_id,
        title_id.ok_or_else(|| io::Error::other("native title ref disappeared"))?,
        caption_id.ok_or_else(|| io::Error::other("native caption ref disappeared"))?,
        new_build_id,
        new_chunk_id,
    ]
    .into_iter()
    .collect::<BTreeSet<_>>();
    assert_eq!(closure, new_ids);
    Ok(())
}

#[test]
fn native_saved_file_movie_creation_reuse_candidate_is_stable() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let (movie, poster) = source_assets(&baseline)?;
    assert_native_candidate(
        NATIVE_FOCUSED_REUSE,
        &movie,
        &poster,
        (120.5, 240.25),
        (320.0, 180.0),
        (320.0, 180.0),
        Duration::from_secs(2),
    )?;
    let options = litchi_keynote::slide::movie::Options::new(
        Point {
            x: 120.5,
            y: 240.25,
        },
        Size {
            width: 320.0,
            height: 180.0,
        },
        Duration::from_secs(2),
    )?
    .with_natural_size(Size {
        width: 320.0,
        height: 180.0,
    })?;
    assert_replay_and_inverse(NATIVE_FOCUSED_REUSE, &movie, &poster, options)
}

#[test]
fn native_saved_file_movie_creation_fresh_candidate_is_stable() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let (source_movie, source_poster) = source_assets(&baseline)?;
    let movie = fresh_movie(&source_movie);
    let poster = fresh_poster(&source_poster)?;
    assert_native_candidate(
        NATIVE_FOCUSED_FRESH,
        &movie,
        &poster,
        (321.0, 42.0),
        (640.0, 360.0),
        (320.0, 180.0),
        Duration::from_millis(1_250),
    )?;
    let options = litchi_keynote::slide::movie::Options::new(
        Point { x: 321.0, y: 42.0 },
        Size {
            width: 640.0,
            height: 360.0,
        },
        Duration::from_millis(1_250),
    )?
    .with_natural_size(Size {
        width: 320.0,
        height: 180.0,
    })?;
    assert_replay_and_inverse(NATIVE_FOCUSED_FRESH, &movie, &poster, options)
}
