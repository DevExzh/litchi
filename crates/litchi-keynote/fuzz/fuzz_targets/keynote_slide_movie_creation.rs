#![no_main]

//! Bounded selector-first fuzzing for fresh Keynote file movies.
//!
//! The target uses the checked-in native package as both the graph oracle and
//! the source of a valid movie/poster pair. It mutates semantic selectors,
//! filenames, payloads, and finite movie options while requiring failed
//! admission to preserve the source and successful transactions to reopen,
//! validate, replay, invert, and double-invert exactly.

mod support;

use std::{sync::OnceLock, time::Duration};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::{
    Limits, MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::movie::Options,
};

use self::support::{
    MAX_INPUT_BYTES, MAX_OBJECTS, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS,
    MAX_TEXT_STORAGES, control, finite_axis, observe_error, package_bytes, read_options,
};

const TARGET_INPUT: &[u8] = b"target-file-movie";
const NATIVE_KEYNOTE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/media-comments-baseline-native.key");

fuzz_target!(|data: &[u8]| {
    let package = native_package();
    let (source_movie, source_poster) = native_assets(package);
    exercise_arbitrary_input(data, source_movie, source_poster);
    exercise_package(package, data, source_movie, source_poster);
});

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, read_options())
            .unwrap_or_else(|error| panic!("native Keynote movie seed must open: {error}"));
        verify_seed_success_contract(&package);
        package
    })
}

fn native_assets(package: &Package) -> (&'static [u8], &'static [u8]) {
    static ASSETS: OnceLock<(Vec<u8>, Vec<u8>)> = OnceLock::new();
    let assets = ASSETS.get_or_init(|| {
        let movie = package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(2),
                MediaPart::Content,
            )
            .unwrap_or_else(|error| panic!("native movie seed must expose movie bytes: {error}"));
        let poster = package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(2),
                MediaPart::Poster,
            )
            .unwrap_or_else(|error| panic!("native movie seed must expose poster bytes: {error}"));
        if movie.is_empty() || poster.is_empty() {
            panic!("native movie seed assets must be nonempty");
        }
        (movie.to_vec(), poster.to_vec())
    });
    (&assets.0, &assets.1)
}

fn exercise_arbitrary_input(data: &[u8], source_movie: &[u8], source_poster: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }
    match Package::from_bytes_with_options(data, read_options()) {
        Ok(package) => exercise_package(&package, data, source_movie, source_poster),
        Err(error) => observe_error(error),
    }
}

fn exercise_package(package: &Package, data: &[u8], source_movie: &[u8], source_poster: &[u8]) {
    let source_bytes = package_bytes(package);
    let movie = requested_movie(data, source_movie);
    let poster = requested_poster(data, source_poster);
    let movie_filename = requested_movie_filename(data);
    let poster_filename = requested_poster_filename(data);
    let options = requested_options(data);
    let commit = match package.add_slide_movie(
        requested_slide(data),
        &movie_filename,
        &movie,
        &poster_filename,
        &poster,
        options,
    ) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    let patch = commit.patch();
    assert!(
        !patch.is_noop(),
        "fresh movie creation must change the source"
    );
    assert_eq!(patch.options(), options);
    assert_eq!(patch.created_objects(), 5);
    assert!(patch.created_data() <= 2);
    let candidate_bytes = package_bytes(commit.package());
    assert_ne!(candidate_bytes, source_bytes);
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("created movie package must validate: {error}"));
    assert_media_payloads(commit.package(), patch, &movie, &poster);
    assert_eq!(package_bytes(package), source_bytes);

    let replay = package
        .apply_slide_movie_creation(patch)
        .unwrap_or_else(|error| panic!("movie-creation forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    replay
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("replayed movie package must validate: {error}"));

    let inverse = patch.inverse();
    let restored = commit
        .package()
        .apply_slide_movie_creation(&inverse)
        .unwrap_or_else(|error| panic!("movie-creation inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    restored
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("inverse-restored movie package must validate: {error}"));

    let double_inverse = inverse.inverse();
    let twice = package
        .apply_slide_movie_creation(&double_inverse)
        .unwrap_or_else(|error| panic!("movie-creation double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    assert_eq!(package_bytes(package), source_bytes);

    assert!(
        commit.package().apply_slide_movie_creation(patch).is_err(),
        "a movie-creation patch must reject its post-state as a source"
    );
    assert!(
        package.apply_slide_movie_creation(&inverse).is_err(),
        "an inverse movie-creation patch must reject its original source"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn verify_seed_success_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let (movie, poster) = native_assets_without_cache(package);
    let options = seed_options();
    let commit = package
        .add_slide_movie(
            SlideSelector::index(0),
            "fuzz-seed.mov",
            &movie,
            "fuzz-seed.png",
            &poster,
            options,
        )
        .unwrap_or_else(|error| panic!("native movie seed must commit: {error}"));
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().options(), options);
    assert_eq!(commit.patch().created_objects(), 5);
    assert_ne!(package_bytes(commit.package()), source_bytes);
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("native movie candidate must validate: {error}"));
    assert_media_payloads(commit.package(), commit.patch(), &movie, &poster);

    let candidate_bytes = package_bytes(commit.package());
    let replay = package
        .apply_slide_movie_creation(commit.patch())
        .unwrap_or_else(|error| panic!("native movie replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    let inverse = commit.patch().inverse();
    let restored = commit
        .package()
        .apply_slide_movie_creation(&inverse)
        .unwrap_or_else(|error| panic!("native movie inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    let twice = package
        .apply_slide_movie_creation(&inverse.inverse())
        .unwrap_or_else(|error| panic!("native movie double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    assert_eq!(package_bytes(package), source_bytes);

    exercise_limit_budget(&movie, &poster);
}

fn native_assets_without_cache(package: &Package) -> (Vec<u8>, Vec<u8>) {
    let movie = package
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Content,
        )
        .unwrap_or_else(|error| panic!("native movie seed must expose movie bytes: {error}"));
    let poster = package
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Poster,
        )
        .unwrap_or_else(|error| panic!("native movie seed must expose poster bytes: {error}"));
    (movie.to_vec(), poster.to_vec())
}

fn assert_media_payloads(
    package: &Package,
    patch: &litchi_keynote::SlideMovieCreationPatch,
    movie: &[u8],
    poster: &[u8],
) {
    assert_eq!(
        package
            .slide_media_data(
                SlideSelector::position(patch.slide_position()),
                MovieSelector::position(patch.movie_position()),
                MediaPart::Content,
            )
            .unwrap_or_else(|error| panic!("created movie data must read back: {error}")),
        movie,
    );
    assert_eq!(
        package
            .slide_media_data(
                SlideSelector::position(patch.slide_position()),
                MovieSelector::position(patch.movie_position()),
                MediaPart::Poster,
            )
            .unwrap_or_else(|error| panic!("created poster data must read back: {error}")),
        poster,
    );
}

fn exercise_limit_budget(movie: &[u8], poster: &[u8]) {
    let semantic = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        1,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| panic!("finite movie semantic limit must build: {error}"));
    let package = Package::from_bytes_with_options(
        NATIVE_KEYNOTE,
        ReadOptions::new(Limits::default(), semantic),
    )
    .unwrap_or_else(|error| panic!("native movie limit seed must open: {error}"));
    let source_bytes = package_bytes(&package);
    assert!(
        package
            .add_slide_movie(
                SlideSelector::index(0),
                "limit.mov",
                movie,
                "limit.png",
                poster,
                seed_options(),
            )
            .is_err(),
        "the tight reference budget must reject movie creation"
    );
    assert_eq!(package_bytes(&package), source_bytes);
}

fn requested_slide(data: &[u8]) -> SlideSelector<'static> {
    if data.starts_with(TARGET_INPUT) {
        return SlideSelector::index(0);
    }
    match control(data, 0) % 6 {
        0..=3 => SlideSelector::index(0),
        4 => SlideSelector::index(usize::MAX),
        _ => SlideSelector::name(""),
    }
}

fn requested_movie_filename(data: &[u8]) -> String {
    if data.starts_with(TARGET_INPUT) {
        return "fuzz-seed.mov".to_owned();
    }
    match control(data, 1) % 8 {
        0 => String::new(),
        1 => "../movie.mov".to_owned(),
        2 => "folder/movie.mov".to_owned(),
        3 => "movie.wav".to_owned(),
        4 => "movie\0.mov".to_owned(),
        _ => format!("fuzz-{:02x}.mov", control(data, 2)),
    }
}

fn requested_poster_filename(data: &[u8]) -> String {
    if data.starts_with(TARGET_INPUT) {
        return "fuzz-seed.png".to_owned();
    }
    match control(data, 3) % 8 {
        0 => String::new(),
        1 => "../poster.png".to_owned(),
        2 => "folder/poster.png".to_owned(),
        3 => "poster.wav".to_owned(),
        4 => "poster\0.png".to_owned(),
        _ => format!("fuzz-{:02x}.png", control(data, 4)),
    }
}

fn requested_options(data: &[u8]) -> Options {
    if data.starts_with(TARGET_INPUT) {
        return seed_options();
    }
    let position = Point {
        x: finite_axis(data, 5, 120.5),
        y: finite_axis(data, 9, 240.25),
    };
    let size = Size {
        width: positive_axis(data, 13, 640.0),
        height: positive_axis(data, 17, 360.0),
    };
    let natural_size = Size {
        width: positive_axis(data, 21, 1_280.0),
        height: positive_axis(data, 25, 720.0),
    };
    Options::new(
        position,
        size,
        Duration::from_millis(1 + (u64::from(control(data, 29)) % 4_000)),
    )
    .and_then(|options| options.with_natural_size(natural_size))
    .unwrap_or_else(|error| panic!("bounded movie options must be valid: {error}"))
}

fn seed_options() -> Options {
    Options::new(
        Point {
            x: 120.5,
            y: 240.25,
        },
        Size {
            width: 640.0,
            height: 360.0,
        },
        Duration::from_millis(1_250),
    )
    .and_then(|options| {
        options.with_natural_size(Size {
            width: 1_280.0,
            height: 720.0,
        })
    })
    .unwrap_or_else(|error| panic!("fixed movie options must be valid: {error}"))
}

fn positive_axis(data: &[u8], offset: usize, fallback: f32) -> f32 {
    finite_axis(data, offset, fallback).abs().max(1.0)
}

fn requested_movie(data: &[u8], source: &[u8]) -> Vec<u8> {
    if data.starts_with(TARGET_INPUT) {
        return source.to_vec();
    }
    match control(data, 31) % 8 {
        0 => Vec::new(),
        1 => b"not a video payload".to_vec(),
        2 => source[..source.len().min(8)].to_vec(),
        3 | 4 | 5 => source.to_vec(),
        6 => mutate_payload(source, data, 32),
        _ => b"\0\0\0\0".to_vec(),
    }
}

fn requested_poster(data: &[u8], source: &[u8]) -> Vec<u8> {
    if data.starts_with(TARGET_INPUT) {
        return source.to_vec();
    }
    match control(data, 33) % 8 {
        0 => Vec::new(),
        1 => b"not an image payload".to_vec(),
        2 => source[..source.len().min(8)].to_vec(),
        3 | 4 | 5 => source.to_vec(),
        6 => mutate_payload(source, data, 34),
        _ => b"\0\0\0\0".to_vec(),
    }
}

fn mutate_payload(source: &[u8], data: &[u8], offset: usize) -> Vec<u8> {
    let mut payload = source.to_vec();
    let index = usize::from(control(data, offset)) % payload.len().max(1);
    if let Some(byte) = payload.get_mut(index) {
        *byte ^= control(data, offset.saturating_add(1)).max(1);
    }
    payload
}
