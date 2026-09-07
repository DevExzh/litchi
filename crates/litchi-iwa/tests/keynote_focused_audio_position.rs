//! Compatibility coverage for the focused Keynote slide-audio position owner.
//!
//! Focused package creation plus host observations verify the selector
//! transaction through mixed source-order media after raw creation and
//! property APIs were retired.

use std::{env, error::Error, fs, io, path::PathBuf, time::Duration};

use litchi_iwa::keynote::{BuildStart, KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa_common::shape::geometry::{Point as HostPoint, Size};
use litchi_keynote::slide::audio::Options as SlideAudioOptions;
use litchi_keynote::slide::media::{MovieKind, Point as FocusedPoint};
use litchi_keynote::slide::movie::Options as SlideMovieOptions;
use litchi_keynote::{MovieSelector, Package, SlideSelector};

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
struct HostMovie {
    content: Vec<u8>,
    poster: Vec<u8>,
    geometry: String,
    properties: String,
    playback: String,
    original_size: String,
    natural_size: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct HostAudio {
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
struct HostSnapshot {
    movies: Vec<HostMovie>,
    audio: Vec<HostAudio>,
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
    editor
        .slide_movies(0)?
        .into_iter()
        .last()
        .map(|movie| movie.drawable_object_id)
        .ok_or_else(|| io::Error::other("focused movie creation produced no host movie").into())
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

    let movie_b_id = editor
        .slide_movies(0)?
        .into_iter()
        .nth(1)
        .ok_or_else(|| io::Error::other("source-built movie B is missing"))?
        .drawable_object_id;
    #[allow(deprecated)]
    editor.set_slide_drawable_comment(
        0,
        movie_b_id,
        "unselected movie comment survives audio position edits",
    )?;
    let bytes = editor.to_bytes()?;
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let movie_b_id = reopened
        .slide_movies(0)?
        .into_iter()
        .nth(1)
        .ok_or_else(|| io::Error::other("reopened movie B is missing"))?
        .drawable_object_id;
    Ok(SourceFixture { bytes, movie_b_id })
}

fn host_snapshot(editor: &KeynoteEditor, movie_b_id: u64) -> TestResult<HostSnapshot> {
    let movies = editor.slide_movies(0)?;
    let movie_ids = movies
        .iter()
        .map(|movie| movie.drawable_object_id)
        .collect::<Vec<_>>();
    let audio = editor.slide_audio(0)?;
    let audio_ids = audio
        .iter()
        .map(|audio| audio.drawable_object_id)
        .collect::<Vec<_>>();
    let movies = movies
        .into_iter()
        .map(|movie| {
            let content = movie
                .movie_data_identifier
                .ok_or_else(|| io::Error::other("movie has no content data"))?;
            let poster = movie
                .poster_image_data_identifier
                .ok_or_else(|| io::Error::other("movie has no poster data"))?;
            Ok(HostMovie {
                content: editor.extract_media(content)?,
                poster: editor.extract_media(poster)?,
                geometry: format!("{:?}", movie.geometry),
                properties: format!("{:?}", movie.properties),
                playback: format!("{:?}", movie.playback),
                original_size: format!("{:?}", movie.original_size),
                natural_size: format!("{:?}", movie.natural_size),
            })
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    let audio = audio
        .into_iter()
        .map(|audio| {
            Ok(HostAudio {
                content: editor.extract_media(audio.audio_data_identifier)?,
                position: (audio.position.x.to_bits(), audio.position.y.to_bits()),
                properties: format!("{:?}", audio.properties),
                playback: format!("{:?}", audio.playback),
                duration: u64::try_from(audio.duration.as_nanos())?,
            })
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;

    let mut builds = editor
        .slide_builds(0)?
        .into_iter()
        .map(|build| -> TestResult<HostBuild> {
            let target_kind = if movie_ids.contains(&build.drawable_object_id) {
                MovieKind::File
            } else if audio_ids.contains(&build.drawable_object_id) {
                MovieKind::Audio
            } else {
                return Err(io::Error::other("build targets unknown media").into());
            };
            let semantic = build.settings.semantic()?;
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
    let mut assets = editor
        .media_assets()?
        .into_iter()
        .map(|asset| editor.extract_media(asset.data_identifier))
        .collect::<Result<Vec<_>, _>>()?;
    assets.sort();
    Ok(HostSnapshot {
        movies,
        audio,
        builds,
        assets,
        movie_b_comment: {
            #[allow(deprecated)]
            editor
                .slide_drawable_comment(0, movie_b_id)?
                .map(|comment| comment.comment.text)
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
    let baseline = host_snapshot(&baseline_editor, fixture.movie_b_id)?;
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
        host_snapshot(&focused_editor, fixture.movie_b_id)?,
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
