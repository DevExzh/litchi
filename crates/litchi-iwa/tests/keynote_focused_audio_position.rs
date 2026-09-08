//! Compatibility coverage for the focused Keynote slide-audio position owner.
//!
//! Focused package creation plus host observations verify the selector
//! transaction through mixed source-order media after raw creation and
//! property APIs were retired.

use std::{env, error::Error, fs, io, path::PathBuf, time::Duration};

use litchi_iwa::keynote::{BuildStart, KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa_archive::{iwa::Archive, package::Catalog};
use litchi_iwa_common::shape::geometry::{Point as HostPoint, Size};
use litchi_iwa_protos::{kn, tsd};
use litchi_keynote::slide::audio::Options as SlideAudioOptions;
use litchi_keynote::slide::media::{MovieKind, Point as FocusedPoint};
use litchi_keynote::slide::movie::Options as SlideMovieOptions;
use litchi_keynote::{DrawableSelector, MediaPart, MovieSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn Error>>;

const MOVIE_A: &[u8] = b"\0\0\0\x18ftypqt  focused-position-movie-a";
const MOVIE_B: &[u8] = b"\0\0\0\x18ftypqt  focused-position-movie-b";
const AUDIO_A: &[u8] = b"FORM\0\0\0\x10AIFCfocused-position-audio-a";
const AUDIO_B: &[u8] = b"FORM\0\0\0\x10AIFCfocused-position-audio-b";
const POSTER_A: &[u8] = b"\x89PNG\r\n\x1a\nfocused-position-poster-a";
const POSTER_B: &[u8] = b"\x89PNG\r\n\x1a\nfocused-position-poster-b";

#[derive(Debug, Clone)]
struct SourceFixture {
    bytes: Vec<u8>,
    movie_b_id: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MovieSnapshot {
    content: Vec<u8>,
    poster: Vec<u8>,
    geometry: String,
    transform: String,
    properties: String,
    playback: String,
    original_size: String,
    natural_size: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AudioSnapshot {
    content: Vec<u8>,
    position: (u32, u32),
    properties: String,
    playback: String,
    duration: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct HostBuild {
    target_kind: MovieKind,
    effect: String,
    start: String,
    duration: u64,
    delay: u64,
    chunks: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MediaSnapshot {
    movies: Vec<MovieSnapshot>,
    audio: Vec<AudioSnapshot>,
    builds: Vec<HostBuild>,
    assets: Vec<Vec<u8>>,
    movie_b_comment: Option<String>,
}

fn movie_options(position: HostPoint, size: Size, duration: Duration) -> SlideMovieOptions {
    SlideMovieOptions::new(position, size, duration).expect("finite source-built movie options")
}

fn audio_options(position: HostPoint, duration: Duration) -> SlideAudioOptions {
    SlideAudioOptions::new(position, duration).expect("finite source-built audio options")
}

fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn focused_drawable_type(message_type: u32) -> bool {
    matches!(
        message_type,
        3_002 | 3_004..=3_009 | 5_021 | 6_000 | 6_007 | 2_011 | 2_014 | 7 | 12
    )
}

fn focused_drawable_selector(source: &[u8], drawable_id: u64) -> TestResult<DrawableSelector> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = match litchi_iwa_archive::iwa::SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            let Some(message) = object.messages.iter().find(|message| message.type_ == 5) else {
                continue;
            };
            let Ok(slide) = kn::SlideArchive::decode(message.data.as_slice()) else {
                continue;
            };
            let mut position = 0usize;
            for reference in slide.owned_drawables {
                let Some(drawable) = archive.object(reference.identifier) else {
                    continue;
                };
                if !drawable
                    .messages
                    .iter()
                    .any(|message| focused_drawable_type(message.type_))
                {
                    continue;
                }
                if reference.identifier == drawable_id {
                    return Ok(DrawableSelector::index(position));
                }
                position = position.saturating_add(1);
            }
        }
    }
    Err(io::Error::other("focused drawable selector target is missing").into())
}

/// Read only the native media identities needed to correlate host build and
/// comment observations. This is a private test oracle; production callers
/// continue to select media through the focused semantic package API.
fn native_media_ids(source: &[u8]) -> TestResult<Vec<(u64, MovieKind)>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut movies = Vec::new();
    let mut slides = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = match litchi_iwa_archive::iwa::SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            for message in &object.messages {
                match message.type_ {
                    5 => {
                        if let Ok(slide) = kn::SlideArchive::decode(message.data.as_slice()) {
                            slides.push((identifier, slide));
                        }
                    },
                    3_007 => {
                        if let Ok(movie) = tsd::MovieArchive::decode(message.data.as_slice()) {
                            let kind = if movie.is_live_video == Some(true) {
                                MovieKind::LiveVideo
                            } else if movie.audio_only == Some(true) {
                                MovieKind::Audio
                            } else if movie.flags.is_some_and(|flags| flags & 1 != 0) {
                                MovieKind::Placeholder
                            } else {
                                MovieKind::File
                            };
                            movies.push((identifier, movie, kind));
                        }
                    },
                    _ => {},
                }
            }
        }
    }

    let mut fallback = None;
    for (slide_identifier, slide) in slides {
        let candidate = slide
            .owned_drawables
            .into_iter()
            .filter_map(|reference| {
                movies
                    .iter()
                    .find(|(identifier, movie, _)| {
                        *identifier == reference.identifier
                            && movie
                                .super_
                                .parent
                                .as_ref()
                                .is_some_and(|parent| parent.identifier == slide_identifier)
                    })
                    .map(|(identifier, _, kind)| (*identifier, *kind))
            })
            .collect::<Vec<_>>();
        if candidate.is_empty() {
            continue;
        }
        let has_audio = candidate.iter().any(|(_, kind)| *kind == MovieKind::Audio);
        let has_file = candidate.iter().any(|(_, kind)| *kind == MovieKind::File);
        if has_audio && has_file {
            return Ok(candidate);
        }
        fallback.get_or_insert(candidate);
    }
    fallback.ok_or_else(|| io::Error::other("native package has no slide-owned media").into())
}

fn last_native_media_id(source: &[u8], kind: MovieKind) -> TestResult<u64> {
    native_media_ids(source)?
        .into_iter()
        .rev()
        .find_map(|(identifier, actual)| (actual == kind).then_some(identifier))
        .ok_or_else(|| io::Error::other("focused media creation produced no native media").into())
}

fn add_audio(
    editor: &mut KeynoteEditor,
    preferred_filename: &str,
    data: &[u8],
    options: SlideAudioOptions,
) -> TestResult {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit =
        package.add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)?;
    let bytes = package_bytes(commit.package())?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn add_movie(
    editor: &mut KeynoteEditor,
    preferred_movie_filename: &str,
    movie_data: &[u8],
    preferred_poster_filename: &str,
    poster_data: &[u8],
    options: SlideMovieOptions,
) -> TestResult<u64> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package.add_slide_movie(
        SlideSelector::index(0),
        preferred_movie_filename,
        movie_data,
        preferred_poster_filename,
        poster_data,
        options,
    )?;
    let bytes = package_bytes(commit.package())?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    last_native_media_id(&bytes, MovieKind::File)
}

fn source_fixture() -> TestResult<SourceFixture> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Focused audio position")
        .subtitle("Source-built selector compatibility")
        .build()?;
    let movie_a_id = add_movie(
        &mut editor,
        "position-movie-a.mov",
        MOVIE_A,
        "position-poster-a.png",
        POSTER_A,
        movie_options(
            HostPoint { x: 120.0, y: 80.0 },
            Size {
                width: 640.0,
                height: 360.0,
            },
            Duration::from_secs(8),
        ),
    )?;
    add_audio(
        &mut editor,
        "position-audio-a.aiff",
        AUDIO_A,
        audio_options(HostPoint { x: 960.0, y: 540.0 }, Duration::from_secs(12)),
    )?;
    add_movie(
        &mut editor,
        "position-movie-b.mov",
        MOVIE_B,
        "position-poster-b.png",
        POSTER_B,
        movie_options(
            HostPoint { x: 500.0, y: 200.0 },
            Size {
                width: 480.0,
                height: 270.0,
            },
            Duration::from_secs(5),
        ),
    )?;
    add_audio(
        &mut editor,
        "position-audio-b.aiff",
        AUDIO_B,
        audio_options(HostPoint { x: 240.0, y: 300.0 }, Duration::from_secs(7)),
    )?;

    let source_build = editor
        .slide_builds(0)?
        .into_iter()
        .find(|build| build.drawable_object_id == movie_a_id)
        .ok_or_else(|| io::Error::other("created movie has no automatic build"))?;
    let mut second_build = source_build.settings;
    second_build.set_start(BuildStart::AfterPrevious)?;
    editor.add_slide_build(0, movie_a_id, second_build)?;

    let bytes = editor.to_bytes()?;
    let movie_b_id = native_media_ids(&bytes)?
        .into_iter()
        .filter(|(_, kind)| *kind == MovieKind::File)
        .nth(1)
        .map(|(identifier, _)| identifier)
        .ok_or_else(|| io::Error::other("source-built movie B is missing"))?;
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let selector = focused_drawable_selector(&editor.to_bytes()?, movie_b_id)?;
    let bytes = package
        .edit_slide_drawable_comment(SlideSelector::index(0), selector)?
        .set("unselected movie comment survives audio position edits")?
        .commit()?
        .into_package();
    let bytes = package_bytes(&bytes)?;
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let movie_b_id = native_media_ids(&reopened.to_bytes()?)?
        .into_iter()
        .filter(|(_, kind)| *kind == MovieKind::File)
        .nth(1)
        .map(|(identifier, _)| identifier)
        .ok_or_else(|| io::Error::other("reopened movie B is missing"))?;
    Ok(SourceFixture { bytes, movie_b_id })
}

fn semantic_snapshot(
    package: &Package,
    editor: &KeynoteEditor,
    movie_b_id: u64,
) -> TestResult<MediaSnapshot> {
    let slide = package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("focused package has no first slide"))?;
    let mut movies = Vec::new();
    let mut audio = Vec::new();
    let mut assets = Vec::new();
    for (movie_position, movie) in slide.movies().iter().copied().enumerate() {
        let content = package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
                MediaPart::Content,
            )?
            .to_vec();
        assets.push(content.clone());
        let properties = package.slide_media_properties(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
        )?;
        if movie.is_audio() {
            let position = movie
                .position()
                .ok_or_else(|| io::Error::other("audio summary has no position"))?;
            audio.push(AudioSnapshot {
                content,
                position: (position.x.to_bits(), position.y.to_bits()),
                properties: format!("{:?}", properties),
                playback: format!("{:?}", movie.playback()),
                duration: u64::try_from(
                    movie
                        .duration()
                        .ok_or_else(|| io::Error::other("audio summary has no duration"))?
                        .as_nanos(),
                )?,
            });
        } else {
            let poster = package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(movie_position),
                    MediaPart::Poster,
                )?
                .to_vec();
            assets.push(poster.clone());
            movies.push(MovieSnapshot {
                content,
                poster,
                geometry: format!("{:?}", (movie.position(), movie.size())),
                transform: format!("{:?}", movie.transform()),
                properties: format!("{:?}", properties),
                playback: format!("{:?}", movie.playback()),
                original_size: format!("{:?}", movie.original_size()),
                natural_size: format!("{:?}", movie.natural_size()),
            });
        }
    }
    assets.sort();
    assets.dedup();

    let media_ids = native_media_ids(&editor.to_bytes()?)?;
    let mut builds = editor
        .slide_builds(0)?
        .into_iter()
        .map(|build| -> TestResult<HostBuild> {
            let (media_index, target_kind) = media_ids
                .iter()
                .position(|(identifier, _)| *identifier == build.drawable_object_id)
                .and_then(|media_index| {
                    let kind = media_ids[media_index].1;
                    let slot = media_ids[..media_index]
                        .iter()
                        .filter(|(_, candidate)| *candidate == kind)
                        .count();
                    match kind {
                        MovieKind::File => Some((slot, MovieKind::File)),
                        MovieKind::Audio => Some((slot, MovieKind::Audio)),
                        _ => None,
                    }
                })
                .ok_or_else(|| io::Error::other("build targets unknown media"))?;
            let semantic = build.settings.semantic()?;
            let _ = media_index;
            Ok(HostBuild {
                target_kind,
                effect: format!("{:?}", semantic.effect()),
                start: format!("{:?}", semantic.start()),
                duration: semantic.duration().as_f64().to_bits(),
                delay: semantic.delay().as_f64().to_bits(),
                chunks: build.chunks.len(),
            })
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;

    builds.sort_by(|left, right| {
        format!("{:?}", left.target_kind)
            .cmp(&format!("{:?}", right.target_kind))
            .then(left.effect.cmp(&right.effect))
            .then(left.start.cmp(&right.start))
            .then(left.duration.cmp(&right.duration))
            .then(left.delay.cmp(&right.delay))
            .then(left.chunks.cmp(&right.chunks))
    });
    Ok(MediaSnapshot {
        movies,
        audio,
        builds,
        assets,
        movie_b_comment: {
            let selector = focused_drawable_selector(&editor.to_bytes()?, movie_b_id)?;
            package
                .slide_drawable_comment(SlideSelector::index(0), selector)?
                .map(|comment| comment.text().to_owned())
        },
    })
}

fn export_candidate(name: &str, bytes: &[u8]) -> TestResult {
    let Ok(directory) = env::var("LITCHI_KEYNOTE_AUDIO_POSITION_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    fs::create_dir_all(&directory)?;
    fs::write(directory.join(name), bytes)?;
    Ok(())
}

#[test]
fn source_built_audio_position_preserves_host_observations_and_opaque_state() -> TestResult {
    let fixture = source_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let source_bytes = package_bytes(&package)?;
    let baseline_editor = KeynoteEditor::from_bytes(&source_bytes)?;
    let baseline = semantic_snapshot(&package, &baseline_editor, fixture.movie_b_id)?;
    assert_eq!(baseline.movies.len(), 2);
    assert_eq!(baseline.audio.len(), 2);
    assert_eq!(baseline.builds.len(), 5);
    assert_eq!(baseline.assets.len(), 6);
    assert_eq!(
        baseline.movie_b_comment.as_deref(),
        Some("unselected movie comment survives audio position edits")
    );

    let show = package.show()?;
    let aggregate = show.slides()[0].movies();
    assert_eq!(aggregate.len(), 4);
    assert!(aggregate[0].kind() == MovieKind::File);
    assert!(aggregate[1].is_audio());
    assert!(aggregate[2].kind() == MovieKind::File);
    assert!(aggregate[3].is_audio());
    for movie_position in [0, 2] {
        assert!(
            package
                .edit_slide_audio_position(
                    SlideSelector::index(0),
                    MovieSelector::index(movie_position)
                )
                .is_err()
        );
    }

    let audio_a_position = FocusedPoint {
        x: 1_120.0,
        y: 420.0,
    };
    let audio_b_position = FocusedPoint { x: 77.0, y: 88.0 };
    let focused = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(1))?
        .set(audio_a_position)?
        .commit()?
        .into_package();
    let focused = focused
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(3))?
        .set(audio_b_position)?
        .commit()?;
    let focused_bytes = package_bytes(focused.package())?;
    let focused_editor = KeynoteEditor::from_bytes(&focused_bytes)?;
    let mut expected = baseline.clone();
    expected.audio[0].position = (audio_a_position.x.to_bits(), audio_a_position.y.to_bits());
    expected.audio[1].position = (audio_b_position.x.to_bits(), audio_b_position.y.to_bits());
    assert_eq!(
        semantic_snapshot(focused.package(), &focused_editor, fixture.movie_b_id)?,
        expected
    );

    export_candidate("audio-position-source-built.key", &source_bytes)?;
    export_candidate("audio-position-source-built-focused.key", &focused_bytes)?;
    Ok(())
}

#[test]
fn source_built_audio_position_validation_is_atomic() -> TestResult {
    let fixture = source_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = package_bytes(&package)?;
    for position in [
        FocusedPoint {
            x: f32::NAN,
            y: 0.0,
        },
        FocusedPoint {
            x: f32::INFINITY,
            y: 0.0,
        },
    ] {
        assert!(
            package
                .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(1))?
                .set(position)
                .is_err()
        );
        assert_eq!(package_bytes(&package)?, before);
    }
    assert!(
        package
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(9))
            .is_err()
    );
    assert_eq!(package_bytes(&package)?, before);
    Ok(())
}
