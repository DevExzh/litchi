//! Compatibility coverage for selector-first Keynote media lifecycle edits.
//!
//! The fixture uses focused package creation for both movies and audio. The
//! focused package then performs the lifecycle transaction using source-order
//! selectors, and each candidate is reopened by the host editor before its
//! semantic projection is compared with the expected state.

use std::collections::HashSet;
use std::error::Error;
use std::io;
use std::time::Duration;

use litchi_iwa::keynote::{BuildStart, KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa_archive::{iwa::Archive, package::Catalog};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_iwa_protos::{kn, tsd};
use litchi_keynote::slide::audio::Options as SlideAudioOptions;
use litchi_keynote::slide::media::MovieKind;
use litchi_keynote::slide::movie::Options as SlideMovieOptions;
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult = Result<(), Box<dyn Error>>;

const MOVIE_A: &[u8] = b"\0\0\0\x18ftypqt  source-built-movie-a";
const MOVIE_B: &[u8] = b"\0\0\0\x18ftypqt  source-built-movie-b";
const AUDIO_A: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio-a";
const AUDIO_B: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio-b";
const POSTER_A: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-poster-a";
const POSTER_B: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-poster-b";

#[derive(Debug, Clone)]
struct SourceFixture {
    bytes: Vec<u8>,
    movie_a_id: u64,
    movie_b_id: u64,
    audio_a_id: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MovieSnapshot {
    content: Vec<u8>,
    poster: Vec<u8>,
    position: Option<(u32, u32)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AudioSnapshot {
    content: Vec<u8>,
    position: (u32, u32),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum MediaSlot {
    Movie(usize),
    Audio(usize),
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct BuildSnapshot {
    target: MediaSlot,
    effect: String,
    start: String,
    duration_bits: u64,
    delay_bits: u64,
    chunks: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MediaSnapshot {
    movies: Vec<MovieSnapshot>,
    audio: Vec<AudioSnapshot>,
    builds: Vec<BuildSnapshot>,
}

fn movie_options(position: Point, size: Size, duration: Duration) -> SlideMovieOptions {
    SlideMovieOptions::new(position, size, duration)
        .expect("source-built fixture uses finite movie values")
}

fn audio_options(position: Point, duration: Duration) -> SlideAudioOptions {
    SlideAudioOptions::new(position, duration)
        .expect("source-built fixture uses finite audio values")
}

fn package_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// Read only native media identities needed by the host build/comment oracle.
/// The production test assertions select media through the focused semantic
/// API; this generated-Prost projection remains private to the integration
/// test so host graph correlations do not become a public raw-ID API.
fn native_media_ids(source: &[u8]) -> Result<Vec<(u64, MovieKind)>, Box<dyn Error>> {
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

fn last_native_media_id(source: &[u8], kind: MovieKind) -> Result<u64, Box<dyn Error>> {
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
) -> Result<u64, Box<dyn Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit =
        package.add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)?;
    let bytes = package_bytes(commit.package())?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    last_native_media_id(&bytes, MovieKind::Audio)
}

fn add_movie(
    editor: &mut KeynoteEditor,
    preferred_movie_filename: &str,
    movie_data: &[u8],
    preferred_poster_filename: &str,
    poster_data: &[u8],
    options: SlideMovieOptions,
) -> Result<u64, Box<dyn Error>> {
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

fn package_watermark(bytes: &[u8]) -> Result<u64, Box<dyn Error>> {
    use prost::Message as _;

    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing metadata component"))?;
    let stream = litchi_iwa_archive::iwa::SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = litchi_iwa_archive::iwa::Archive::parse(&stream)?;
    let message = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == 11_006)
        .ok_or_else(|| io::Error::other("missing package metadata"))?;
    Ok(
        litchi_iwa_protos::tsp::PackageMetadata::decode(message.data.as_slice())?
            .last_object_identifier,
    )
}

#[test]
fn focused_suffix_release_and_next_allocation_reopen_in_host() -> TestResult {
    let fixture = source_fixture()?;
    for (source_index, is_movie) in [(0, true), (1, false)] {
        let package = Package::from_bytes(&fixture.bytes)?;
        let duplicated = package
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(source_index))?;
        let duplicate_bytes = package_bytes(duplicated.package())?;
        let old_watermark = package_watermark(&duplicate_bytes)?;
        let removed = duplicated
            .package()
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(4))?;
        let removed_bytes = package_bytes(removed.package())?;
        let released = package_watermark(&removed_bytes)?;
        assert!(released < old_watermark);
        let removed_host = KeynoteEditor::from_bytes(&removed_bytes)?;
        assert_eq!(released, package_watermark(&removed_host.to_bytes()?)?);
        let removed_slide = removed
            .package()
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("removed package has no first slide"))?;
        assert_eq!(
            removed_slide
                .movies()
                .iter()
                .filter(|movie| movie.kind() == MovieKind::File)
                .count(),
            2
        );
        assert_eq!(removed_slide.audio().count(), 2,);

        let repeated = removed
            .package()
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(source_index))?;
        let repeated_bytes = package_bytes(repeated.package())?;
        let _repeated_host = KeynoteEditor::from_bytes(&repeated_bytes)?;
        let desired_kind = if is_movie {
            MovieKind::File
        } else {
            MovieKind::Audio
        };
        let next_focused_id = last_native_media_id(&repeated_bytes, desired_kind)?;
        assert!(next_focused_id > released);
        let repeated_slide = repeated
            .package()
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("repeated package has no first slide"))?;
        assert_eq!(
            repeated_slide
                .movies()
                .iter()
                .filter(|movie| movie.kind() == MovieKind::File)
                .count(),
            if is_movie { 3 } else { 2 }
        );
        assert_eq!(repeated_slide.audio().count(), if is_movie { 2 } else { 3 });
        assert_eq!(package_bytes(&package)?, fixture.bytes);
        let restored = removed
            .package()
            .apply_slide_media_lifecycle(&removed.patch().inverse())?;
        assert_eq!(package_bytes(restored.package())?, duplicate_bytes);
    }
    Ok(())
}

fn movie_labels(
    package: &Package,
    movie_position: usize,
) -> Result<(Option<String>, Option<String>), Box<dyn Error>> {
    let slide = SlideSelector::index(0);
    let movie = MovieSelector::index(movie_position);
    Ok((
        package.slide_movie_title(slide, movie)?,
        package.slide_movie_caption(slide, movie)?,
    ))
}

fn set_movie_labels(
    package: Package,
    movie_position: usize,
    title: &str,
    caption: &str,
) -> Result<Package, Box<dyn Error>> {
    let package = package
        .edit_slide_movie_title(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
        )?
        .set(title)?
        .commit()?
        .into_package();
    Ok(package
        .edit_slide_movie_caption(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
        )?
        .set(caption)?
        .commit()?
        .into_package())
}

#[allow(deprecated)]
fn source_fixture() -> Result<SourceFixture, Box<dyn Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Focused media lifecycle")
        .subtitle("Source-built compatibility fixture")
        .build()?;

    // Keep the source order mixed so MovieSelector exercises one stable
    // sequence for both movies and audio: movie A, audio A, movie B, audio B.
    let movie_a_id = add_movie(
        &mut editor,
        "movie-a.mov",
        MOVIE_A,
        "poster-a.png",
        POSTER_A,
        movie_options(
            Point { x: 120.0, y: 80.0 },
            Size {
                width: 640.0,
                height: 360.0,
            },
            Duration::from_secs(8),
        ),
    )?;
    let audio_a_id = add_audio(
        &mut editor,
        "audio-a.aiff",
        AUDIO_A,
        audio_options(Point { x: 960.0, y: 540.0 }, Duration::from_secs(12)),
    )?;
    let movie_b_id = add_movie(
        &mut editor,
        "movie-b.mov",
        MOVIE_B,
        "poster-b.png",
        POSTER_B,
        movie_options(
            Point { x: 500.0, y: 200.0 },
            Size {
                width: 480.0,
                height: 270.0,
            },
            Duration::from_secs(5),
        ),
    )?;
    add_audio(
        &mut editor,
        "audio-b.aiff",
        AUDIO_B,
        audio_options(Point { x: 240.0, y: 300.0 }, Duration::from_secs(7)),
    )?;

    // Focused creation creates one native build/chunk per media. Add
    // a second build to movie A so the focused clone must preserve more than
    // the common one-build case.
    let source_build = editor
        .slide_builds(0)?
        .into_iter()
        .find(|build| build.drawable_object_id == movie_a_id)
        .ok_or_else(|| io::Error::other("created movie A has no automatic build"))?;
    let mut second_build = source_build.settings;
    second_build.set_start(BuildStart::AfterPrevious)?;
    editor.add_slide_build(0, movie_a_id, second_build)?;

    #[allow(deprecated)]
    editor.set_slide_drawable_comment(0, movie_b_id, "unselected movie B lifecycle comment")?;

    let package = set_movie_labels(
        Package::from_bytes(&editor.to_bytes()?)?,
        0,
        "Movie A title",
        "Movie A caption",
    )?;
    let bytes = package_bytes(&package)?;

    // Reopen once before returning so this fixture itself proves that labels,
    // comments, and the extra build are accepted by the host reader.
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let reopened_package = Package::from_bytes(&bytes)?;
    let reopened_slide = reopened_package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("reopened package has no first slide"))?;
    assert_eq!(
        reopened_slide
            .movies()
            .iter()
            .filter(|movie| movie.kind() == MovieKind::File)
            .count(),
        2
    );
    assert_eq!(reopened_slide.audio().count(), 2);
    assert_eq!(reopened.slide_builds(0)?.len(), 5);
    assert_eq!(
        reopened
            .slide_drawable_comment(0, movie_b_id)?
            .map(|comment| comment.comment.text),
        Some("unselected movie B lifecycle comment".to_owned())
    );

    Ok(SourceFixture {
        bytes,
        movie_a_id,
        movie_b_id,
        audio_a_id,
    })
}

fn semantic_snapshot(
    package: &Package,
    editor: &KeynoteEditor,
) -> Result<MediaSnapshot, Box<dyn Error>> {
    let slide = package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("focused package has no first slide"))?;
    let mut movies = Vec::new();
    let mut audio = Vec::new();
    for (movie_position, movie) in slide.movies().iter().copied().enumerate() {
        let content = package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
                MediaPart::Content,
            )?
            .to_vec();
        if movie.is_audio() {
            let position = movie
                .position()
                .ok_or_else(|| io::Error::other("audio summary has no position"))?;
            audio.push(AudioSnapshot {
                content,
                position: (position.x.to_bits(), position.y.to_bits()),
            });
        } else {
            let poster = package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(movie_position),
                    MediaPart::Poster,
                )?
                .to_vec();
            movies.push(MovieSnapshot {
                content,
                poster,
                position: movie
                    .position()
                    .map(|point| (point.x.to_bits(), point.y.to_bits())),
            });
        }
    }

    let media_ids = native_media_ids(&editor.to_bytes()?)?;
    let mut builds = editor
        .slide_builds(0)?
        .into_iter()
        .map(|build| {
            let target = media_ids
                .iter()
                .position(|(identifier, _)| *identifier == build.drawable_object_id)
                .and_then(|media_index| {
                    let kind = media_ids[media_index].1;
                    let slot = media_ids[..media_index]
                        .iter()
                        .filter(|(_, candidate)| *candidate == kind)
                        .count();
                    match kind {
                        MovieKind::File => Some(MediaSlot::Movie(slot)),
                        MovieKind::Audio => Some(MediaSlot::Audio(slot)),
                        _ => None,
                    }
                })
                .ok_or_else(|| io::Error::other("host build targets unknown media"))?;
            let semantic = build.settings.semantic()?;
            Ok(BuildSnapshot {
                target,
                effect: format!("{:?}", semantic.effect()),
                start: format!("{:?}", semantic.start()),
                duration_bits: semantic.duration().as_f64().to_bits(),
                delay_bits: semantic.delay().as_f64().to_bits(),
                chunks: build.chunks.len(),
            })
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    builds.sort();

    Ok(MediaSnapshot {
        movies,
        audio,
        builds,
    })
}

fn host_from_package(package: &Package) -> Result<KeynoteEditor, Box<dyn Error>> {
    Ok(KeynoteEditor::from_bytes(&package_bytes(package)?)?)
}

fn offset_position(position: Option<(u32, u32)>) -> Option<(u32, u32)> {
    position.map(|(x, y)| {
        (
            (f32::from_bits(x) + 10.0).to_bits(),
            (f32::from_bits(y) + 10.0).to_bits(),
        )
    })
}

fn offset_audio_position(position: (u32, u32)) -> (u32, u32) {
    (
        (f32::from_bits(position.0) + 10.0).to_bits(),
        (f32::from_bits(position.1) + 10.0).to_bits(),
    )
}

fn expected_builds_with_clone(
    baseline: &[BuildSnapshot],
    source: MediaSlot,
    clone: MediaSlot,
) -> Vec<BuildSnapshot> {
    let mut expected = baseline.to_vec();
    expected.extend(
        baseline
            .iter()
            .filter(|build| build.target == source)
            .map(|build| {
                let mut clone_build = build.clone();
                clone_build.target = clone;
                clone_build
            }),
    );
    expected.sort();
    expected
}

fn has_media_bytes(package: &Package, expected: &[u8]) -> Result<bool, Box<dyn Error>> {
    let slide = package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("focused package has no first slide"))?;
    for movie_position in 0..slide.movies().len() {
        if package.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
            MediaPart::Content,
        )? == expected
        {
            return Ok(true);
        }
        if !slide.movies()[movie_position].is_audio()
            && package.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
                MediaPart::Poster,
            )? == expected
        {
            return Ok(true);
        }
    }
    Ok(false)
}

#[allow(deprecated)]
fn comment_text(
    editor: &KeynoteEditor,
    drawable_object_id: u64,
) -> Result<Option<String>, Box<dyn Error>> {
    Ok(editor
        .slide_drawable_comment(0, drawable_object_id)?
        .map(|comment| comment.comment.text))
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentStorageSnapshot {
    storage_id: u64,
    text: String,
    author_id: Option<u64>,
    storage_uuid: Option<(u64, u64)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawableCommentSnapshot {
    root: CommentStorageSnapshot,
    replies: Vec<CommentStorageSnapshot>,
}

fn storage_snapshot(
    storage_id: u64,
    comment: &litchi_iwa_common::comment::Comment,
) -> CommentStorageSnapshot {
    CommentStorageSnapshot {
        storage_id,
        text: comment.text.clone(),
        author_id: comment.author_id.map(|author| author.get()),
        storage_uuid: comment
            .storage_uuid
            .map(|uuid| (uuid.lower(), uuid.upper())),
    }
}

#[allow(deprecated)]
fn drawable_comment_snapshot(
    editor: &KeynoteEditor,
    drawable_object_id: u64,
) -> Result<Option<DrawableCommentSnapshot>, Box<dyn Error>> {
    let Some(root) = editor.slide_drawable_comment(0, drawable_object_id)? else {
        return Ok(None);
    };
    let replies = editor.slide_drawable_comment_replies(0, drawable_object_id)?;
    Ok(Some(DrawableCommentSnapshot {
        root: storage_snapshot(root.storage_id.get(), &root.comment),
        replies: replies
            .into_iter()
            .map(|reply| storage_snapshot(reply.storage_id.get(), &reply.comment))
            .collect(),
    }))
}

fn assert_cloned_comment(source: &DrawableCommentSnapshot, cloned: &DrawableCommentSnapshot) {
    assert_eq!(source.replies.len(), cloned.replies.len());
    let source_nodes = std::iter::once(&source.root).chain(&source.replies);
    let cloned_nodes = std::iter::once(&cloned.root).chain(&cloned.replies);
    for (before, after) in source_nodes.zip(cloned_nodes) {
        assert_ne!(before.storage_id, after.storage_id);
        assert_eq!(before.text, after.text);
        assert_eq!(before.author_id, after.author_id);
        assert_eq!(before.storage_uuid, after.storage_uuid);
        assert!(
            std::iter::once(&source.root)
                .chain(&source.replies)
                .all(|node| node.storage_id != after.storage_id)
        );
    }
}

#[test]
fn source_built_media_lifecycle_matches_reopened_host_semantics() -> TestResult {
    let fixture = source_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let source_bytes = package_bytes(&package)?;
    let baseline_editor = host_from_package(&package)?;
    let baseline = semantic_snapshot(&package, &baseline_editor)?;

    assert_eq!(baseline.movies.len(), 2);
    assert_eq!(baseline.audio.len(), 2);
    assert_eq!(baseline.builds.len(), 5);
    assert_eq!(
        baseline
            .builds
            .iter()
            .map(|build| build.chunks)
            .sum::<usize>(),
        5
    );
    assert_eq!(
        movie_labels(&package, 0)?,
        (
            Some("Movie A title".to_owned()),
            Some("Movie A caption".to_owned())
        )
    );
    assert_eq!(
        comment_text(&baseline_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );

    let duplicate_movie = package
        .duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(0))
        .map_err(|error| io::Error::other(format!("duplicate source-built movie: {error}")))?;
    assert_eq!(
        package_bytes(&package)?,
        source_bytes,
        "source snapshot was mutated"
    );
    let duplicate_movie_editor = host_from_package(duplicate_movie.package())?;
    let duplicate_movie_state =
        semantic_snapshot(duplicate_movie.package(), &duplicate_movie_editor)?;
    let mut expected_movies = baseline.movies.clone();
    let mut cloned_movie = baseline.movies[0].clone();
    cloned_movie.position = offset_position(cloned_movie.position);
    expected_movies.push(cloned_movie);
    assert_eq!(duplicate_movie_state.movies, expected_movies);
    assert_eq!(duplicate_movie_state.audio, baseline.audio);
    assert_eq!(
        duplicate_movie_state.builds,
        expected_builds_with_clone(&baseline.builds, MediaSlot::Movie(0), MediaSlot::Movie(2))
    );
    assert_eq!(
        movie_labels(duplicate_movie.package(), 4)?,
        (
            Some("Movie A title".to_owned()),
            Some("Movie A caption".to_owned())
        )
    );
    assert_eq!(
        comment_text(&duplicate_movie_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    assert_eq!(duplicate_movie.diagnostics().source_media_count(), 4);
    assert_eq!(duplicate_movie.diagnostics().target_media_count(), 5);
    assert!(duplicate_movie.diagnostics().created_objects() > 0);
    assert_eq!(duplicate_movie.diagnostics().removed_objects(), 0);
    assert!(has_media_bytes(duplicate_movie.package(), MOVIE_A)?);
    assert!(has_media_bytes(duplicate_movie.package(), POSTER_A)?);

    let duplicate_movie_inverse = duplicate_movie.patch().inverse();
    let restored_movie = duplicate_movie
        .package()
        .apply_slide_media_lifecycle(&duplicate_movie_inverse)?;
    assert_eq!(package_bytes(restored_movie.package())?, source_bytes);

    let foreign = package
        .edit_slide_movie_title(SlideSelector::index(0), MovieSelector::index(0))?
        .set("foreign source title")?
        .commit()?
        .into_package();
    let foreign_before = package_bytes(&foreign)?;
    assert!(
        foreign
            .apply_slide_media_lifecycle(duplicate_movie.patch())
            .is_err()
    );
    assert_eq!(package_bytes(&foreign)?, foreign_before);

    let duplicate_audio =
        package.duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(1))?;
    let duplicate_audio_editor = host_from_package(duplicate_audio.package())?;
    let duplicate_audio_state =
        semantic_snapshot(duplicate_audio.package(), &duplicate_audio_editor)?;
    assert_eq!(duplicate_audio_state.movies, baseline.movies);
    let mut expected_audio = baseline.audio.clone();
    let mut cloned_audio = baseline.audio[0].clone();
    cloned_audio.position = offset_audio_position(cloned_audio.position);
    expected_audio.push(cloned_audio);
    assert_eq!(duplicate_audio_state.audio, expected_audio);
    assert_eq!(
        duplicate_audio_state.builds,
        expected_builds_with_clone(&baseline.builds, MediaSlot::Audio(0), MediaSlot::Audio(2))
    );
    assert!(has_media_bytes(duplicate_audio.package(), AUDIO_A)?);
    assert_eq!(duplicate_audio.diagnostics().source_media_count(), 4);
    assert_eq!(duplicate_audio.diagnostics().target_media_count(), 5);

    let duplicate_audio_inverse = duplicate_audio.patch().inverse();
    let restored_audio = duplicate_audio
        .package()
        .apply_slide_media_lifecycle(&duplicate_audio_inverse)?;
    assert_eq!(package_bytes(restored_audio.package())?, source_bytes);

    // Duplicate first, then remove each owner in source order.  This checks
    // shared DataInfo ownership first and final physical-data reclamation
    // second, while each intermediate package is independently host-openable.
    let movie_clone =
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(0))?;
    let movie_shared = movie_clone
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(0))?;
    let movie_shared_editor = host_from_package(movie_shared.package())?;
    let movie_shared_state = semantic_snapshot(movie_shared.package(), &movie_shared_editor)?;
    assert_eq!(movie_shared_state.movies.len(), 2);
    assert_eq!(movie_shared_state.audio.len(), 2);
    assert_eq!(movie_shared_state.movies[0].content, MOVIE_B);
    assert_eq!(movie_shared_state.movies[1].content, MOVIE_A);
    assert_eq!(movie_shared_state.builds.len(), 5);
    assert_eq!(
        movie_shared_state
            .builds
            .iter()
            .map(|build| build.chunks)
            .sum::<usize>(),
        5
    );
    assert!(has_media_bytes(movie_shared.package(), MOVIE_A)?);
    assert!(has_media_bytes(movie_shared.package(), POSTER_A)?);
    assert_eq!(
        movie_labels(movie_shared.package(), 3)?,
        (
            Some("Movie A title".to_owned()),
            Some("Movie A caption".to_owned())
        )
    );
    assert_eq!(
        comment_text(&movie_shared_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );

    let movie_final = movie_shared
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(3))?;
    let movie_final_editor = host_from_package(movie_final.package())?;
    let movie_final_state = semantic_snapshot(movie_final.package(), &movie_final_editor)?;
    assert_eq!(movie_final_state.movies.len(), 1);
    assert_eq!(movie_final_state.movies[0].content, MOVIE_B);
    assert_eq!(movie_final_state.audio, baseline.audio);
    assert_eq!(
        comment_text(&movie_final_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    assert_eq!(movie_final_state.builds.len(), 3);
    assert_eq!(
        movie_final_state
            .builds
            .iter()
            .map(|build| build.chunks)
            .sum::<usize>(),
        3
    );
    assert!(!has_media_bytes(movie_final.package(), MOVIE_A)?);
    assert!(!has_media_bytes(movie_final.package(), POSTER_A)?);
    assert!(has_media_bytes(movie_final.package(), MOVIE_B)?);
    assert!(has_media_bytes(movie_final.package(), POSTER_B)?);
    assert!(has_media_bytes(movie_final.package(), AUDIO_A)?);
    assert!(has_media_bytes(movie_final.package(), AUDIO_B)?);

    let movie_final_inverse = movie_final.patch().inverse();
    let movie_shared_inverse = movie_shared.patch().inverse();
    let movie_clone_inverse = movie_clone.patch().inverse();
    let restored_movie_shared = movie_final
        .package()
        .apply_slide_media_lifecycle(&movie_final_inverse)?;
    let restored_movie_clone = restored_movie_shared
        .package()
        .apply_slide_media_lifecycle(&movie_shared_inverse)?;
    let restored_movie_source = restored_movie_clone
        .package()
        .apply_slide_media_lifecycle(&movie_clone_inverse)?;
    assert_eq!(
        package_bytes(restored_movie_source.package())?,
        source_bytes
    );

    let audio_clone =
        package.duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(1))?;
    let audio_shared = audio_clone
        .package()
        .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(1))?;
    let audio_shared_editor = host_from_package(audio_shared.package())?;
    let audio_shared_state = semantic_snapshot(audio_shared.package(), &audio_shared_editor)?;
    assert_eq!(audio_shared_state.movies, baseline.movies);
    assert_eq!(audio_shared_state.audio[0].content, AUDIO_B);
    assert_eq!(audio_shared_state.audio[1].content, AUDIO_A);
    assert_eq!(audio_shared_state.builds.len(), 5);
    assert_eq!(
        comment_text(&audio_shared_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    assert!(has_media_bytes(audio_shared.package(), AUDIO_A)?);

    let audio_final = audio_shared
        .package()
        .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(3))?;
    let audio_final_editor = host_from_package(audio_final.package())?;
    let audio_final_state = semantic_snapshot(audio_final.package(), &audio_final_editor)?;
    assert_eq!(audio_final_state.movies, baseline.movies);
    assert_eq!(audio_final_state.audio.len(), 1);
    assert_eq!(audio_final_state.audio[0].content, AUDIO_B);
    assert_eq!(
        comment_text(&audio_final_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    assert_eq!(audio_final_state.builds.len(), 4);
    assert_eq!(
        audio_final_state
            .builds
            .iter()
            .map(|build| build.chunks)
            .sum::<usize>(),
        4
    );
    assert!(!has_media_bytes(audio_final.package(), AUDIO_A)?);
    assert!(has_media_bytes(audio_final.package(), AUDIO_B)?);
    assert!(has_media_bytes(audio_final.package(), MOVIE_A)?);
    assert!(has_media_bytes(audio_final.package(), POSTER_A)?);
    assert!(has_media_bytes(audio_final.package(), MOVIE_B)?);
    assert!(has_media_bytes(audio_final.package(), POSTER_B)?);

    let audio_final_inverse = audio_final.patch().inverse();
    let audio_shared_inverse = audio_shared.patch().inverse();
    let audio_clone_inverse = audio_clone.patch().inverse();
    let restored_audio_shared = audio_final
        .package()
        .apply_slide_media_lifecycle(&audio_final_inverse)?;
    let restored_audio_clone = restored_audio_shared
        .package()
        .apply_slide_media_lifecycle(&audio_shared_inverse)?;
    let restored_audio_source = restored_audio_clone
        .package()
        .apply_slide_media_lifecycle(&audio_clone_inverse)?;
    assert_eq!(
        package_bytes(restored_audio_source.package())?,
        source_bytes
    );
    assert_eq!(
        package_bytes(&package)?,
        source_bytes,
        "source snapshot changed"
    );
    Ok(())
}

#[test]
fn source_built_media_lifecycle_type_guards_are_atomic() -> TestResult {
    let fixture = source_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = package_bytes(&package)?;

    // The source order is movie A, audio A, movie B, audio B.  Typed wrappers
    // must reject a selector of the other kind without publishing anything.
    assert!(
        package
            .duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(1))
            .is_err()
    );
    assert!(
        package
            .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(1))
            .is_err()
    );
    assert!(
        package
            .duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(package_bytes(&package)?, before);
    Ok(())
}

#[test]
#[allow(deprecated)]
fn source_built_selected_commented_movie_duplicates_and_removal_is_atomic() -> TestResult {
    let fixture = source_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = package_bytes(&package)?;

    // The mixed source order is movie A, audio A, movie B, audio B.  Movie B
    // owns an existing direct comment. Duplication preserves its semantic
    // identity while allocating an independent storage graph.
    let editor = host_from_package(&package)?;
    assert_eq!(
        comment_text(&editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    let source_comment = drawable_comment_snapshot(&editor, fixture.movie_b_id)?
        .ok_or_else(|| io::Error::other("source movie comment missing"))?;
    let duplicate =
        package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let duplicate_editor = host_from_package(duplicate.package())?;
    let cloned_movie_id =
        last_native_media_id(&package_bytes(duplicate.package())?, MovieKind::File)?;
    let cloned_comment = drawable_comment_snapshot(&duplicate_editor, cloned_movie_id)?
        .ok_or_else(|| io::Error::other("cloned movie comment missing"))?;
    assert_cloned_comment(&source_comment, &cloned_comment);
    let restored = duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, before);
    let removed = package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let removed_editor = host_from_package(removed.package())?;
    let remaining = semantic_snapshot(removed.package(), &removed_editor)?;
    assert_eq!(remaining.movies.len(), 1);
    assert_eq!(remaining.movies[0].content, MOVIE_A);
    assert_eq!(remaining.audio, semantic_snapshot(&package, &editor)?.audio);
    assert!(
        native_media_ids(&package_bytes(removed.package())?)?
            .iter()
            .all(|(identifier, _)| *identifier != fixture.movie_b_id)
    );
    let restored_removed = removed
        .package()
        .apply_slide_media_lifecycle(&removed.patch().inverse())?;
    assert_eq!(package_bytes(restored_removed.package())?, before);
    assert_eq!(package_bytes(&package)?, before);
    Ok(())
}

#[test]
#[allow(deprecated)]
fn source_built_selected_commented_audio_with_reply_duplicates_and_removal_is_atomic() -> TestResult
{
    let fixture = source_fixture()?;
    let mut editor = KeynoteEditor::from_bytes(&fixture.bytes)?;
    #[allow(deprecated)]
    editor.set_slide_drawable_comment(0, fixture.audio_a_id, "selected audio lifecycle comment")?;
    #[allow(deprecated)]
    let reply_id = editor.add_slide_drawable_comment_reply(
        0,
        fixture.audio_a_id,
        "selected audio lifecycle reply",
    )?;
    let bytes = editor.to_bytes()?;
    let reopened = KeynoteEditor::from_bytes(&bytes)?;
    let comment = drawable_comment_snapshot(&reopened, fixture.audio_a_id)?
        .ok_or_else(|| io::Error::other("selected audio comment was not persisted"))?;
    assert_eq!(comment.root.text, "selected audio lifecycle comment");
    assert_eq!(comment.replies.len(), 1);
    assert_eq!(comment.replies[0].storage_id, reply_id);
    assert_eq!(comment.replies[0].text, "selected audio lifecycle reply");
    assert!(comment.root.storage_uuid.is_some());
    assert!(comment.replies[0].storage_uuid.is_some());
    assert_ne!(
        comment.root.storage_uuid, comment.replies[0].storage_uuid,
        "root and reply must have distinct storage UUIDs"
    );

    let package = Package::from_bytes(&bytes)?;
    let before = package_bytes(&package)?;
    let duplicate =
        package.duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(1))?;
    let duplicate_editor = host_from_package(duplicate.package())?;
    let cloned_audio_id =
        last_native_media_id(&package_bytes(duplicate.package())?, MovieKind::Audio)?;
    let cloned_comment = drawable_comment_snapshot(&duplicate_editor, cloned_audio_id)?
        .ok_or_else(|| io::Error::other("cloned audio comment missing"))?;
    assert_cloned_comment(&comment, &cloned_comment);
    assert_eq!(
        drawable_comment_snapshot(&duplicate_editor, fixture.audio_a_id)?,
        Some(comment)
    );
    let restored = duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(package_bytes(restored.package())?, before);
    let removed = package.remove_slide_audio(SlideSelector::index(0), MovieSelector::index(1))?;
    let removed_editor = host_from_package(removed.package())?;
    let remaining = semantic_snapshot(removed.package(), &removed_editor)?;
    assert_eq!(remaining.audio.len(), 1);
    assert_eq!(remaining.audio[0].content, AUDIO_B);
    assert_eq!(
        remaining.movies,
        semantic_snapshot(&package, &reopened)?.movies
    );
    assert_eq!(
        comment_text(&removed_editor, fixture.movie_b_id)?,
        Some("unselected movie B lifecycle comment".to_owned())
    );
    assert!(
        native_media_ids(&package_bytes(removed.package())?)?
            .iter()
            .all(|(identifier, _)| *identifier != fixture.audio_a_id)
    );
    let restored_removed = removed
        .package()
        .apply_slide_media_lifecycle(&removed.patch().inverse())?;
    assert_eq!(package_bytes(restored_removed.package())?, before);
    assert_eq!(package_bytes(&package)?, before);
    Ok(())
}

#[test]
fn source_built_multiple_reply_threads_survive_focused_media_lifecycle() -> TestResult {
    let fixture = source_fixture()?;
    let mut editor = KeynoteEditor::from_bytes(&fixture.bytes)?;

    #[allow(deprecated)]
    editor.set_slide_drawable_comment(0, fixture.movie_a_id, "selected movie lifecycle comment")?;
    #[allow(deprecated)]
    editor.set_slide_drawable_comment(0, fixture.audio_a_id, "selected audio lifecycle comment")?;

    let movie_before_reply = drawable_comment_snapshot(&editor, fixture.movie_a_id)?
        .ok_or_else(|| io::Error::other("selected movie comment was not created"))?;
    let audio_before_reply = drawable_comment_snapshot(&editor, fixture.audio_a_id)?
        .ok_or_else(|| io::Error::other("selected audio comment was not created"))?;
    assert!(movie_before_reply.replies.is_empty());
    assert!(audio_before_reply.replies.is_empty());

    #[allow(deprecated)]
    let movie_reply_id = editor.add_slide_drawable_comment_reply(
        0,
        fixture.movie_a_id,
        "selected movie lifecycle reply",
    )?;
    #[allow(deprecated)]
    let audio_reply_id = editor.add_slide_drawable_comment_reply(
        0,
        fixture.audio_a_id,
        "selected audio lifecycle reply",
    )?;

    let movie = drawable_comment_snapshot(&editor, fixture.movie_a_id)?
        .ok_or_else(|| io::Error::other("selected movie comment disappeared"))?;
    let audio = drawable_comment_snapshot(&editor, fixture.audio_a_id)?
        .ok_or_else(|| io::Error::other("selected audio comment disappeared"))?;
    assert_eq!(movie.root.text, "selected movie lifecycle comment");
    assert_eq!(audio.root.text, "selected audio lifecycle comment");
    assert_eq!(movie.replies.len(), 1);
    assert_eq!(audio.replies.len(), 1);
    assert_eq!(movie.replies[0].storage_id, movie_reply_id);
    assert_eq!(audio.replies[0].storage_id, audio_reply_id);
    assert_eq!(movie.replies[0].text, "selected movie lifecycle reply");
    assert_eq!(audio.replies[0].text, "selected audio lifecycle reply");

    // Legacy reply insertion copy-on-writes the root while preserving its
    // storage UUID.  Reply storage gets a fresh object ID and UUID.  The
    // generated author is shared across the two selected threads.
    assert_ne!(movie.root.storage_id, movie_before_reply.root.storage_id);
    assert_eq!(
        movie.root.storage_uuid,
        movie_before_reply.root.storage_uuid
    );
    assert_ne!(audio.root.storage_id, audio_before_reply.root.storage_id);
    assert_eq!(
        audio.root.storage_uuid,
        audio_before_reply.root.storage_uuid
    );
    assert_eq!(movie.root.author_id, audio.root.author_id);
    assert_eq!(movie.root.author_id, movie.replies[0].author_id);
    assert_eq!(audio.root.author_id, audio.replies[0].author_id);
    let uuids = [
        movie.root.storage_uuid,
        movie.replies[0].storage_uuid,
        audio.root.storage_uuid,
        audio.replies[0].storage_uuid,
    ];
    assert!(uuids.iter().all(Option::is_some));
    let unique_uuids = uuids.into_iter().flatten().collect::<HashSet<_>>();
    assert_eq!(unique_uuids.len(), 4);

    let reopened = KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
    assert_eq!(
        drawable_comment_snapshot(&reopened, fixture.movie_a_id)?,
        Some(movie.clone())
    );
    assert_eq!(
        drawable_comment_snapshot(&reopened, fixture.audio_a_id)?,
        Some(audio.clone())
    );

    let before_operations = editor.to_bytes()?;
    let package = Package::from_bytes(&before_operations)?;
    for (source_index, selected_id, source_thread, other_id, other_thread) in [
        (0, fixture.movie_a_id, &movie, fixture.audio_a_id, &audio),
        (1, fixture.audio_a_id, &audio, fixture.movie_a_id, &movie),
    ] {
        let duplicate = package
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(source_index))?;
        let duplicate_host = host_from_package(duplicate.package())?;
        let clone_id = last_native_media_id(
            &package_bytes(duplicate.package())?,
            if source_index == 0 {
                MovieKind::File
            } else {
                MovieKind::Audio
            },
        )?;
        let cloned_thread = drawable_comment_snapshot(&duplicate_host, clone_id)?
            .ok_or_else(|| io::Error::other("focused clone lost its comment thread"))?;
        assert_cloned_comment(source_thread, &cloned_thread);
        assert_eq!(
            drawable_comment_snapshot(&duplicate_host, other_id)?,
            Some(other_thread.clone())
        );

        let removed = package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(source_index))?;
        let removed_host = host_from_package(removed.package())?;
        assert!(
            native_media_ids(&package_bytes(removed.package())?)?
                .iter()
                .all(|(identifier, _)| *identifier != selected_id)
        );
        assert_eq!(
            drawable_comment_snapshot(&removed_host, other_id)?,
            Some(other_thread.clone())
        );
        let restored = removed
            .package()
            .apply_slide_media_lifecycle(&removed.patch().inverse())?;
        assert_eq!(package_bytes(restored.package())?, before_operations);
        let restored = duplicate
            .package()
            .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
        assert_eq!(package_bytes(restored.package())?, before_operations);
    }
    assert_eq!(package_bytes(&package)?, before_operations);
    Ok(())
}
