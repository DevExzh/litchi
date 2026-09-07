//! Native Keynote acceptance coverage for a freshly created slide audio clip.
//!
//! The focused fixture was opened, saved, closed, and reopened in Keynote
//! 14.4.  These assertions keep the public package API selector-first while
//! using a small native oracle for the one appended audio build.

use std::{collections::BTreeSet, io, time::Duration};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::shape::geometry::Point;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{kn, tsd};
use litchi_keynote::{
    MediaPart, MovieKind, MovieSelector, Package, SlideSelector, slide::audio::Options,
};
use prost::Message;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_FOCUSED: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/slide-audio-creation-focused-native.key");
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;

fn audio() -> Vec<u8> {
    // A complete 100 ms, mono, signed 16-bit PCM WAV.  This is the exact
    // payload used to create the focused Keynote fixture.
    let mut data = Vec::new();
    data.extend_from_slice(b"RIFF");
    data.extend_from_slice(&1_636u32.to_le_bytes());
    data.extend_from_slice(b"WAVEfmt ");
    data.extend_from_slice(&16u32.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&8_000u32.to_le_bytes());
    data.extend_from_slice(&16_000u32.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    data.extend_from_slice(&16u16.to_le_bytes());
    data.extend_from_slice(b"data");
    data.extend_from_slice(&1_600u32.to_le_bytes());
    for sample in 0..800i16 {
        data.extend_from_slice(&((sample % 50 - 25) * 400).to_le_bytes());
    }
    data
}

fn creation_options() -> TestResult<Options> {
    Ok(Options::new(
        Point {
            x: 120.5,
            y: 240.25,
        },
        Duration::from_millis(100),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
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

fn native_slide(source: &[u8]) -> TestResult<(u64, kn::SlideArchive)> {
    // Follow the same presentation-root topology as the package reader:
    // Document -> Show -> first SlideTree entry -> SlideNode -> Slide.  A
    // media-child heuristic is insufficient because Keynote also stores
    // layout/master slide archives with movie-like drawable references.
    let document: kn::DocumentArchive = native_object_message(source, 1, 1)?;
    let show: kn::ShowArchive = native_object_message(source, document.show.identifier, 2)?;
    let node_reference = show
        .slide_tree
        .slides
        .first()
        .ok_or_else(|| io::Error::other("native presentation slide tree is empty"))?;
    let node: kn::SlideNodeArchive = native_object_message(source, node_reference.identifier, 4)?;
    let slide_reference = node
        .slide
        .ok_or_else(|| io::Error::other("native presentation slide node has no slide"))?;
    let slide: kn::SlideArchive =
        native_object_message(source, slide_reference.identifier, SLIDE_MESSAGE_TYPE)?;
    Ok((slide_reference.identifier, slide))
}

fn native_object_message<T>(source: &[u8], identifier: u64, message_type: u32) -> TestResult<T>
where
    T: Message + Default,
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

fn native_audio_ids(source: &[u8]) -> TestResult<Vec<u64>> {
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
            && movie.audio_only == Some(true)
        {
            identifiers.push(reference.identifier);
        }
    }
    Ok(identifiers)
}

fn assert_existing_media_unchanged(
    baseline: &Package,
    focused: &Package,
    baseline_movies: &[litchi_keynote::slide::media::MovieInfo],
) -> TestResult {
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
            focused.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Content,
            )?,
            baseline.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Content,
            )?,
            "media content {index}"
        );
        if !baseline_movie.is_audio() {
            assert_eq!(
                focused.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    MediaPart::Poster,
                )?,
                baseline.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    MediaPart::Poster,
                )?,
                "media poster {index}"
            );
        }
    }
    Ok(())
}

#[test]
fn saved_keynote_audio_creation_appends_native_audio_and_start_build() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let focused = Package::from_bytes(NATIVE_FOCUSED)?;
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

    let requested = creation_options()?;
    let created = focused_movies
        .last()
        .ok_or_else(|| io::Error::other("focused audio is missing"))?;
    assert_eq!(created.kind(), MovieKind::Audio);
    assert_eq!(
        created.position().map(|point| (point.x, point.y)),
        Some((requested.position().x, requested.position().y))
    );
    assert_eq!(created.duration(), Some(requested.duration()));
    assert_eq!(
        focused.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(baseline_movies.len()),
            MediaPart::Content,
        )?,
        audio()
    );

    let baseline_audio = native_audio_ids(NATIVE_BASELINE)?;
    let focused_audio = native_audio_ids(NATIVE_FOCUSED)?;
    assert_eq!(focused_audio.len(), baseline_audio.len() + 1);
    let new_audio = focused_audio
        .iter()
        .copied()
        .filter(|identifier| !baseline_audio.contains(identifier))
        .collect::<Vec<_>>();
    assert_eq!(new_audio.len(), 1);
    let new_audio = new_audio[0];

    let (_, baseline_slide) = native_slide(NATIVE_BASELINE)?;
    let (_, focused_slide) = native_slide(NATIVE_FOCUSED)?;
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
    let new_build = native_object_message::<kn::BuildArchive>(
        NATIVE_FOCUSED,
        new_builds[0],
        BUILD_MESSAGE_TYPE,
    )?;
    assert_eq!(
        new_build
            .drawable
            .as_ref()
            .map(|reference| reference.identifier),
        Some(new_audio)
    );
    assert_eq!(
        new_build
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|animation| animation.effect.as_deref()),
        Some("apple:audio-start")
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
    let new_chunk = native_object_message::<kn::BuildChunkArchive>(
        NATIVE_FOCUSED,
        new_chunks[0],
        BUILD_CHUNK_MESSAGE_TYPE,
    )?;
    assert_eq!(
        new_chunk
            .build
            .as_ref()
            .map(|reference| reference.identifier),
        Some(new_builds[0])
    );
    Ok(())
}

#[test]
fn audio_creation_from_saved_keynote_fixture_round_trips_exactly() -> TestResult {
    let source = Package::from_bytes(NATIVE_FOCUSED)?;
    let data = audio();
    let requested = Options::new(Point { x: 321.0, y: 42.0 }, Duration::from_millis(250))?;
    let commit =
        source.add_slide_audio(SlideSelector::index(0), "second-pcm.wav", &data, requested)?;
    let candidate_bytes = exact_bytes(commit.package())?;
    assert_ne!(candidate_bytes, exact_bytes(&source)?);
    assert_eq!(commit.patch().created_objects(), 5);
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(commit.patch().movie_position()),
            MediaPart::Content,
        )?,
        data
    );

    let replay = source.apply_slide_audio_creation(commit.patch())?;
    assert_eq!(exact_bytes(replay.package())?, candidate_bytes);
    let restored = commit
        .package()
        .apply_slide_audio_creation(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, exact_bytes(&source)?);
    Ok(())
}
