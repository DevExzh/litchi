#![no_main]

//! Bounded fuzzing for selector-first Keynote slide media replacement.
//!
//! Arbitrary input is always admitted through the normal bounded package
//! reader.  A tiny synthetic source then guarantees that the mutation path is
//! reached often enough to exercise shared content/poster records, audio
//! content, exact no-ops, inverse patches, stale conflicts, and source
//! preservation without embedding a large native Keynote document.

mod keynote_slide_media_data_seed;

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
};

use keynote_slide_media_data_seed as seed;

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const MAX_ENTRIES: usize = 256;
const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SLIDES: usize = 512;
const MAX_REFERENCES: usize = 32 * 1024;
const MAX_TEXT_STORAGES: usize = 8 * 1024;
const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;
const MAX_REPLACEMENT_BYTES: usize = 256;

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_seed(data);
    exercise_input_limit();
});

fn options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid media fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid media fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn seed_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(seed::bytes(), options())
            .unwrap_or_else(|error| panic!("tiny Keynote media fuzz seed must open: {error}"));
        verify_seed_mutation_paths(&package);
        package
    })
}

// A malformed synthetic seed must fail the harness instead of silently
// routing every generated input through refusal-only coverage.
fn verify_seed_mutation_paths(package: &Package) {
    for (position, part) in [
        (0, MediaPart::Content),
        (0, MediaPart::Poster),
        (2, MediaPart::Content),
    ] {
        let slide = SlideSelector::index(0);
        let movie = MovieSelector::index(position);
        let mut replacement = package
            .slide_media_data(slide, movie, part)
            .unwrap_or_else(|error| panic!("fuzz seed media must read: {error}"))
            .to_vec();
        *replacement.last_mut().expect("fuzz seed media is nonempty") ^= 1;
        let commit = package
            .edit_slide_media_data(slide, movie, part)
            .and_then(|edit| edit.set(&replacement))
            .and_then(|edit| edit.commit())
            .unwrap_or_else(|error| panic!("fuzz seed must exercise changed media: {error}"));
        assert!(!commit.patch().is_noop());
        let restored = commit
            .package()
            .apply_slide_media_data(&commit.patch().inverse())
            .unwrap_or_else(|error| panic!("fuzz seed inverse must apply: {error}"));
        assert_eq!(
            package_bytes(restored.package()).as_deref(),
            Some(seed::bytes())
        );
    }
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }
}

fn exercise_seed(data: &[u8]) {
    exercise_package(seed_package(), data);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let Some(source_bytes) = package_bytes(package) else {
        return;
    };
    let movie_position = usize::from(control(data, 0) % 3);
    let part = if control(data, 1) & 1 == 0 {
        MediaPart::Content
    } else {
        MediaPart::Poster
    };
    let slide = if control(data, 2) & 1 == 0 {
        SlideSelector::index(0)
    } else {
        SlideSelector::name("Fuzz media")
    };
    let movie = MovieSelector::index(movie_position);
    let before = match package.slide_media_data(slide, movie, part) {
        Ok(bytes) => bytes.to_vec(),
        Err(error) => {
            observe_error(error);
            assert_eq!(
                package_bytes(package).as_deref(),
                Some(source_bytes.as_slice())
            );
            return;
        },
    };
    let replacement = replacement(part, movie_position, &before, data);
    let edit = match package.edit_slide_media_data(slide, movie, part) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(
                package_bytes(package).as_deref(),
                Some(source_bytes.as_slice())
            );
            return;
        },
    };
    let use_noop = control(data, 3) & 1 == 0;
    let desired = if use_noop {
        before.clone()
    } else if control(data, 4) & 1 != 0 {
        incompatible_replacement(part, movie_position)
    } else {
        replacement.clone()
    };
    let edit = match edit.set(&desired) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(
                package_bytes(package).as_deref(),
                Some(source_bytes.as_slice())
            );
            return;
        },
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(
                package_bytes(package).as_deref(),
                Some(source_bytes.as_slice())
            );
            return;
        },
    };
    let Some(candidate_bytes) = package_bytes(commit.package()) else {
        return;
    };
    assert_eq!(commit.patch().is_noop(), source_bytes == candidate_bytes);
    assert_eq!(commit.patch().before_length(), before.len());
    assert_eq!(commit.patch().after_length(), desired.len());
    assert_eq!(commit.patch().part(), part);
    assert_eq!(commit.diagnostics().changed(), !commit.patch().is_noop());
    assert_eq!(
        commit
            .package()
            .slide_media_data(slide, movie, part)
            .unwrap_or_else(|error| panic!("committed media read failed: {error}")),
        desired.as_slice(),
    );

    if !commit.patch().is_noop() {
        assert!(
            commit
                .package()
                .apply_slide_media_data(commit.patch())
                .is_err()
        );
    }

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = commit
        .package()
        .apply_slide_media_data(&inverse)
        .unwrap_or_else(|error| panic!("media inverse must apply: {error}"));
    assert_eq!(
        package_bytes(restored.package()).as_deref(),
        Some(source_bytes.as_slice())
    );
    assert_eq!(
        restored
            .package()
            .slide_media_data(slide, movie, part)
            .unwrap_or_else(|error| panic!("restored media read failed: {error}")),
        before.as_slice(),
    );

    let reapplied = package
        .apply_slide_media_data(commit.patch())
        .unwrap_or_else(|error| panic!("media forward patch must apply: {error}"));
    assert_eq!(package_bytes(reapplied.package()), Some(candidate_bytes));

    // A shared data record is visible through the second file movie.  The
    // audio record is intentionally independent and is therefore unchanged.
    if movie_position < 2 {
        for shared_movie in 0..2 {
            assert_eq!(
                commit
                    .package()
                    .slide_media_data(
                        SlideSelector::index(0),
                        MovieSelector::index(shared_movie),
                        part
                    )
                    .unwrap_or_else(|error| panic!("shared media read failed: {error}")),
                desired.as_slice(),
            );
        }
        let other_part = match part {
            MediaPart::Content => MediaPart::Poster,
            MediaPart::Poster => MediaPart::Content,
            _ => MediaPart::Content,
        };
        assert_eq!(
            package
                .slide_media_data(SlideSelector::index(0), MovieSelector::index(0), other_part,)
                .ok(),
            commit
                .package()
                .slide_media_data(SlideSelector::index(0), MovieSelector::index(0), other_part,)
                .ok(),
        );
    }
}

fn replacement(part: MediaPart, movie_position: usize, before: &[u8], data: &[u8]) -> Vec<u8> {
    let prefix = match (part, movie_position) {
        (MediaPart::Content, 2) => seed::REPLACED_AUDIO_BYTES,
        (MediaPart::Content, _) => seed::REPLACED_MOVIE_BYTES,
        (MediaPart::Poster, _) => seed::REPLACED_POSTER_BYTES,
        _ => seed::REPLACED_MOVIE_BYTES,
    };
    let mut value = prefix.to_vec();
    let remaining = MAX_REPLACEMENT_BYTES.saturating_sub(value.len());
    value.extend(
        data.iter()
            .take(remaining.min(64))
            .map(|byte| byte.rotate_left(1)),
    );
    if value == before {
        value.push(0x5a);
    }
    value
}

fn incompatible_replacement(part: MediaPart, movie_position: usize) -> Vec<u8> {
    match (part, movie_position) {
        (MediaPart::Content, 2) => seed::REPLACED_MOVIE_BYTES.to_vec(),
        (MediaPart::Content, _) => seed::REPLACED_POSTER_BYTES.to_vec(),
        (MediaPart::Poster, _) => seed::REPLACED_MOVIE_BYTES.to_vec(),
        _ => seed::REPLACED_POSTER_BYTES.to_vec(),
    }
}

fn exercise_input_limit() {
    let source = seed::bytes();
    let defaults = Limits::default();
    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1)).unwrap_or(0),
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )
    .unwrap_or_else(|error| panic!("finite input limit must construct: {error}"));
    if let Err(error) = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(tight, SemanticLimits::default()),
    ) {
        observe_error(error);
    } else {
        panic!("tight media fuzz input limit unexpectedly accepted source");
    }
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Option<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).ok().map(|()| bytes)
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
