//! Fresh file-movie creation through semantic selectors and exact-source patches.
//!
//! The movie and poster bytes used by the creation tests are borrowed from a
//! complete native fixture.  This keeps the package transaction responsible
//! for creating a new graph while the test oracle remains a real Keynote media
//! pair.  Set `LITCHI_KEYNOTE_MOVIE_CREATION_OUTPUT_DIR` to export a candidate
//! package and its two payloads for a native Keynote open/save/close/reopen
//! acceptance run.

use std::{collections::BTreeMap, env, fs, io, path::PathBuf, time::Duration};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::{
    MediaPart, MovieKind, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::movie::Options,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const SOURCE_MOVIE_POSITION: usize = 2;
const NATIVE_MOVIE_FILENAME: &str = "native-source-movie.mov";
const NATIVE_POSTER_FILENAME: &str = "native-source-poster.png";
const REUSED_CANDIDATE_FILENAME: &str = "movie-creation-candidate.key";
const FRESH_CANDIDATE_FILENAME: &str = "movie-creation-fresh-candidate.key";

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn source_assets(source: &Package) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let movie = source.slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(SOURCE_MOVIE_POSITION),
        MediaPart::Content,
    )?;
    let poster = source.slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(SOURCE_MOVIE_POSITION),
        MediaPart::Poster,
    )?;
    if movie.is_empty() || poster.is_empty() {
        return Err(io::Error::other("native source movie assets are empty").into());
    }
    Ok((movie.to_vec(), poster.to_vec()))
}

fn options() -> TestResult<Options> {
    Ok(Options::new(
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
    })?)
}

fn fresh_options() -> TestResult<Options> {
    Ok(Options::new(
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
    })?)
}

fn fresh_movie(source: &[u8]) -> Vec<u8> {
    // ISO-BMFF permits a `free` atom at the end of a movie.  Appending one
    // keeps the borrowed QuickTime payload structurally valid while producing
    // a distinct digest for the fresh-data path.
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

fn members(source: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

fn export_candidate(
    candidate: &Package,
    movie: &[u8],
    poster: &[u8],
    candidate_filename: &str,
    movie_filename: &str,
    poster_filename: &str,
) -> TestResult {
    let Ok(directory) = env::var("LITCHI_KEYNOTE_MOVIE_CREATION_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    fs::create_dir_all(&directory)?;
    fs::write(directory.join(candidate_filename), exact_bytes(candidate)?)?;
    fs::write(directory.join(movie_filename), movie)?;
    fs::write(directory.join(poster_filename), poster)?;
    Ok(())
}

fn create_movie(
    source: &Package,
    movie_filename: &str,
    movie: &[u8],
    poster_filename: &str,
    poster: &[u8],
    requested: Options,
) -> TestResult<litchi_keynote::SlideMovieCreationCommit> {
    Ok(source.add_slide_movie(
        SlideSelector::index(0),
        movie_filename,
        movie,
        poster_filename,
        poster,
        requested,
    )?)
}

#[test]
fn creates_file_movie_with_borrowed_native_assets_and_restores_exact_source() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let before_movies = source.slides()?[0].movies().to_vec();
    let (movie, poster) = source_assets(&source)?;
    let requested = options()?;
    let commit = create_movie(
        &source,
        "fresh-source-movie.mov",
        &movie,
        "fresh-source-poster.png",
        &poster,
        requested,
    )?;

    let patch = commit.patch();
    assert!(commit.diagnostics().changed());
    assert_eq!(patch.source_media_count(), before_movies.len());
    assert_eq!(patch.target_media_count(), before_movies.len() + 1);
    assert_eq!(patch.movie_position().get(), before_movies.len());
    assert_eq!(patch.created_objects(), 5);
    // Both payloads are borrowed byte-for-byte from an existing native movie,
    // so metadata planning reuses their materialized DataInfo records despite
    // the distinct preferred filenames.
    assert_eq!(patch.created_data(), 0);
    assert_eq!(patch.options(), requested);

    let after_movies = commit.package().slides()?[0].movies();
    assert_eq!(
        &after_movies[..before_movies.len()],
        before_movies.as_slice()
    );
    let created = after_movies
        .last()
        .ok_or_else(|| io::Error::other("missing created movie"))?;
    assert_eq!(created.kind(), MovieKind::File);
    assert_eq!(
        created.position().map(|point| (point.x, point.y)),
        Some((requested.position().x, requested.position().y))
    );
    assert_eq!(
        created.size().map(|size| (size.width, size.height)),
        Some((requested.size().width, requested.size().height))
    );
    assert_eq!(
        created.natural_size().map(|size| (size.width, size.height)),
        Some((
            requested.natural_size().width,
            requested.natural_size().height
        ))
    );
    assert_eq!(
        created
            .original_size()
            .map(|size| (size.width, size.height)),
        Some((
            requested.natural_size().width,
            requested.natural_size().height
        ))
    );
    assert_eq!(created.duration(), Some(requested.duration()));
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(patch.movie_position()),
            MediaPart::Content,
        )?,
        movie.as_slice()
    );
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(patch.movie_position()),
            MediaPart::Poster,
        )?,
        poster.as_slice()
    );

    // The new movie receives one native start build; source builds remain
    // intact and the semantic source snapshot contains no native IDs.
    assert_eq!(
        commit.package().slides()?[0].builds().len(),
        source.slides()?[0].builds().len() + 1
    );
    export_candidate(
        commit.package(),
        &movie,
        &poster,
        REUSED_CANDIDATE_FILENAME,
        NATIVE_MOVIE_FILENAME,
        NATIVE_POSTER_FILENAME,
    )?;

    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    let inverse = patch.inverse();
    assert_eq!(inverse.removed_objects(), 5);
    assert_eq!(inverse.removed_data(), 0);
    let restored = commit.package().apply_slide_movie_creation(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE_SOURCE);
    assert_eq!(
        restored.diagnostics().restored_previews(),
        patch.deleted_previews()
    );

    let replay = source.apply_slide_movie_creation(patch)?;
    assert_eq!(
        exact_bytes(replay.package())?,
        exact_bytes(commit.package())?
    );
    let twice = source.apply_slide_movie_creation(&inverse.inverse())?;
    assert_eq!(
        exact_bytes(twice.package())?,
        exact_bytes(commit.package())?
    );
    Ok(())
}

#[test]
fn fresh_movie_and_poster_content_create_two_new_data_records() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (source_movie, source_poster) = source_assets(&source)?;
    let movie = fresh_movie(&source_movie);
    let poster = fresh_poster(&source_poster)?;
    assert_ne!(movie, source_movie);
    assert_ne!(poster, source_poster);
    let requested = fresh_options()?;

    let commit = create_movie(
        &source,
        "fresh-content.mov",
        &movie,
        "fresh-content.png",
        &poster,
        requested,
    )?;
    assert_eq!(commit.patch().created_data(), 2);
    assert_eq!(commit.patch().created_objects(), 5);
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(commit.patch().movie_position()),
            MediaPart::Content,
        )?,
        movie.as_slice()
    );
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(commit.patch().movie_position()),
            MediaPart::Poster,
        )?,
        poster.as_slice()
    );
    let created = commit.package().slides()?[0]
        .movies()
        .last()
        .ok_or_else(|| io::Error::other("missing fresh movie"))?;
    assert_eq!(
        created
            .original_size()
            .map(|size| (size.width, size.height)),
        Some((
            requested.natural_size().width,
            requested.natural_size().height
        ))
    );
    export_candidate(
        commit.package(),
        &movie,
        &poster,
        FRESH_CANDIDATE_FILENAME,
        "fresh-content-movie.mov",
        "fresh-content-poster.png",
    )?;
    let restored = commit
        .package()
        .apply_slide_movie_creation(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn identical_movie_and_poster_content_reuses_data_records() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (movie, poster) = source_assets(&source)?;
    let first = create_movie(
        &source,
        "first-source-movie.mov",
        &movie,
        "first-source-poster.png",
        &poster,
        options()?,
    )?;
    let second = create_movie(
        first.package(),
        "second-source-movie.mov",
        &movie,
        "second-source-poster.png",
        &poster,
        fresh_options()?,
    )?;

    assert_eq!(first.patch().created_data(), 0);
    assert_eq!(second.patch().created_data(), 0);
    assert_eq!(second.patch().created_objects(), 5);
    let first_data = members(&exact_bytes(first.package())?)?
        .into_iter()
        .filter(|(name, _)| name.starts_with("Data/"))
        .collect::<BTreeMap<_, _>>();
    let second_data = members(&exact_bytes(second.package())?)?
        .into_iter()
        .filter(|(name, _)| name.starts_with("Data/"))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(first_data, second_data);

    let restored = second
        .package()
        .apply_slide_movie_creation(&second.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(first.package())?
    );
    Ok(())
}

#[test]
fn mixed_new_and_reused_assets_register_only_one_record() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (source_movie, source_poster) = source_assets(&source)?;
    let new_movie = fresh_movie(&source_movie);
    let new_poster = fresh_poster(&source_poster)?;
    let source_members = members(NATIVE_SOURCE)?;
    for new_part in [MediaPart::Content, MediaPart::Poster] {
        let movie = if new_part == MediaPart::Content {
            &new_movie
        } else {
            &source_movie
        };
        let poster = if new_part == MediaPart::Poster {
            &new_poster
        } else {
            &source_poster
        };
        let commit = create_movie(&source, "mixed.mov", movie, "mixed.png", poster, options()?)?;
        assert_eq!(commit.patch().created_data(), 1, "{new_part:?}");
        for (part, expected) in [(MediaPart::Content, movie), (MediaPart::Poster, poster)] {
            assert_eq!(
                commit.package().slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::position(commit.patch().movie_position()),
                    part,
                )?,
                expected.as_slice(),
            );
        }
        let candidate_members = members(&exact_bytes(commit.package())?)?;
        for (name, bytes) in source_members
            .iter()
            .filter(|(name, _)| name.starts_with("Data/"))
        {
            assert_eq!(candidate_members.get(name), Some(bytes), "{name}");
        }
        assert_eq!(
            candidate_members
                .keys()
                .filter(|name| name.starts_with("Data/"))
                .count(),
            source_members
                .keys()
                .filter(|name| name.starts_with("Data/"))
                .count()
                + 1,
        );
        let restored = commit
            .package()
            .apply_slide_movie_creation(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, NATIVE_SOURCE);
        assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    }
    Ok(())
}

#[test]
fn oversized_filenames_retain_movie_and_poster_error_provenance() -> TestResult {
    use litchi_keynote::SlideMovieCreationError;

    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (movie, poster) = source_assets(&source)?;
    let movie_name = format!("{}.mov", "m".repeat(4_093));
    let poster_name = format!("{}.png", "p".repeat(4_093));
    assert_eq!(
        source
            .add_slide_movie(
                SlideSelector::index(0),
                &movie_name,
                &movie,
                "poster.png",
                &poster,
                options()?
            )
            .err(),
        Some(SlideMovieCreationError::InvalidMovieFilename),
    );
    assert_eq!(
        source
            .add_slide_movie(
                SlideSelector::index(0),
                "movie.mov",
                &movie,
                &poster_name,
                &poster,
                options()?
            )
            .err(),
        Some(SlideMovieCreationError::InvalidPosterFilename),
    );
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn patch_requires_the_exact_source_snapshot() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (movie, poster) = source_assets(&source)?;
    let first = create_movie(
        &source,
        "stale-source-movie.mov",
        &movie,
        "stale-source-poster.png",
        &poster,
        options()?,
    )?;
    assert!(
        first
            .package()
            .apply_slide_movie_creation(first.patch())
            .is_err()
    );
    assert!(
        source
            .apply_slide_movie_creation(&first.patch().inverse())
            .is_err()
    );
    Ok(())
}

#[test]
fn semantic_selector_and_invalid_inputs_are_atomic() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (movie, poster) = source_assets(&source)?;
    let before = exact_bytes(&source)?;
    let requested = options()?;

    assert!(
        source
            .add_slide_movie(
                SlideSelector::index(usize::MAX),
                "movie.mov",
                &movie,
                "poster.png",
                &poster,
                requested,
            )
            .is_err()
    );
    assert!(
        source
            .add_slide_movie(
                SlideSelector::name(""),
                "movie.mov",
                &movie,
                "poster.png",
                &poster,
                requested,
            )
            .is_err()
    );

    for name in [
        "",
        "../movie.mov",
        "folder/movie.mov",
        "movie\\track.mov",
        "movie.wav",
        "movie\0.mov",
    ] {
        assert!(
            source
                .add_slide_movie(
                    SlideSelector::index(0),
                    name,
                    &movie,
                    "poster.png",
                    &poster,
                    requested,
                )
                .is_err()
        );
    }
    for name in [
        "",
        "../poster.png",
        "folder/poster.png",
        "poster\\image.png",
        "poster.wav",
        "poster\0.png",
    ] {
        assert!(
            source
                .add_slide_movie(
                    SlideSelector::index(0),
                    "movie.mov",
                    &movie,
                    name,
                    &poster,
                    requested,
                )
                .is_err()
        );
    }
    for (invalid_movie, invalid_poster) in [
        (b"".as_slice(), poster.as_slice()),
        (b"not a video payload".as_slice(), poster.as_slice()),
        (movie.as_slice(), b"not an image payload".as_slice()),
    ] {
        assert!(
            source
                .add_slide_movie(
                    SlideSelector::index(0),
                    "movie.mov",
                    invalid_movie,
                    "poster.png",
                    invalid_poster,
                    requested,
                )
                .is_err()
        );
    }
    assert_eq!(exact_bytes(&source)?, before);
    Ok(())
}

#[test]
fn entry_limit_refusal_does_not_mutate_the_source() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (movie, poster) = source_assets(&source)?;
    let largest = Catalog::from_bytes(NATIVE_SOURCE)?
        .iter()
        .map(|entry| entry.data().len())
        .max()
        .ok_or_else(|| io::Error::other("empty native fixture"))?;
    let defaults = Limits::default();
    let limits = Limits::new(
        defaults.max_input_bytes(),
        defaults.max_entries(),
        largest as u64,
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let limited = Package::from_bytes_with_options(
        NATIVE_SOURCE,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let mut oversized_movie = movie;
    oversized_movie.resize(largest + 1, 0);
    let before = exact_bytes(&limited)?;
    assert!(
        limited
            .add_slide_movie(
                SlideSelector::index(0),
                "oversized.mov",
                &oversized_movie,
                "poster.png",
                &poster,
                options()?,
            )
            .is_err()
    );
    assert_eq!(exact_bytes(&limited)?, before);
    Ok(())
}
