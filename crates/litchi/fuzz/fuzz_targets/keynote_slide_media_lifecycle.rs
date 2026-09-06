#![no_main]

//! Bounded fuzzing for selector-first Keynote media lifecycle transactions.
//!
//! The target replays arbitrary command bytes against a tiny synthetic
//! package.  The package has two file movies sharing their content and poster
//! records plus one independent audio drawable, so every operation reaches a
//! meaningful ownership case without embedding a native Keynote package in
//! the fuzz binary.  The transaction remains source-bound: a candidate can be
//! inverted only from its exact source snapshot, and a stale or foreign
//! source must be rejected without changing its bytes.

mod keynote_slide_media_data_seed;

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{Limits, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector};
use litchi_iwa_archive::package::Catalog;

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
        .unwrap_or_else(|error| unreachable!("valid lifecycle archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid lifecycle semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn seed_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(seed::bytes(), options())
            .unwrap_or_else(|error| panic!("tiny Keynote lifecycle fuzz seed must open: {error}"));
        verify_seed_lifecycle_paths(&package);
        package
    })
}

fn verify_seed_lifecycle_paths(package: &Package) {
    let source = package_bytes(package)
        .unwrap_or_else(|| panic!("tiny Keynote lifecycle fuzz seed must serialize"));
    for (duplicate, slide, movie) in [
        (true, SlideSelector::index(0), MovieSelector::index(0)),
        (
            true,
            SlideSelector::name("Fuzz media"),
            MovieSelector::index(2),
        ),
        (
            false,
            SlideSelector::name("Fuzz media"),
            MovieSelector::index(0),
        ),
        (false, SlideSelector::index(0), MovieSelector::index(2)),
    ] {
        let commit = if duplicate {
            package
                .duplicate_slide_media(slide, movie)
                .unwrap_or_else(|error| panic!("seed lifecycle duplication must succeed: {error}"))
        } else {
            package
                .remove_slide_media(slide, movie)
                .unwrap_or_else(|error| panic!("seed lifecycle removal must succeed: {error}"))
        };
        let candidate = package_bytes(commit.package())
            .unwrap_or_else(|| panic!("seed lifecycle candidate must serialize"));
        assert_ne!(candidate, source);
        let inverse = commit.patch().inverse();
        let restored = commit
            .package()
            .apply_slide_media_lifecycle(&inverse)
            .unwrap_or_else(|error| panic!("seed lifecycle inverse must succeed: {error}"));
        assert_eq!(
            package_bytes(restored.package()).as_deref(),
            Some(source.as_slice())
        );
        assert_eq!(package_bytes(package).as_deref(), Some(source.as_slice()));
    }
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => {
            let before = package_bytes(&package);
            let slide = if control(data, 0) & 1 == 0 {
                SlideSelector::index(0)
            } else {
                SlideSelector::name("Fuzz media")
            };
            let movie = MovieSelector::index(usize::from(control(data, 1) % 3));
            let result = if control(data, 2) & 1 == 0 {
                package.duplicate_slide_media(slide, movie)
            } else {
                package.remove_slide_media(slide, movie)
            };
            if let Err(error) = result {
                observe_error(error);
            }
            assert_eq!(package_bytes(&package), before);
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_seed(data: &[u8]) {
    let package = seed_package();
    exercise_lifecycle(package, data);
}

fn exercise_lifecycle(package: &Package, data: &[u8]) {
    let Some(source_bytes) = package_bytes(package) else {
        return;
    };
    let slide = if control(data, 0) & 1 == 0 {
        SlideSelector::index(0)
    } else {
        SlideSelector::name("Fuzz media")
    };
    let movie_position = usize::from(control(data, 1) % 3);
    let movie = MovieSelector::index(movie_position);
    let duplicate = control(data, 2) & 1 == 0;
    let result = if duplicate {
        package.duplicate_slide_media(slide, movie)
    } else {
        package.remove_slide_media(slide, movie)
    };

    let commit = match result {
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
    let candidate_bytes = package_bytes(commit.package());
    assert_ne!(candidate_bytes.as_deref(), Some(source_bytes.as_slice()));
    assert_eq!(
        package_bytes(package).as_deref(),
        Some(source_bytes.as_slice())
    );

    // The unrelated sentinel and all media records that are not the selected
    // final owner retain their exact bytes.  A file movie is shared by two
    // owners in the seed, so removing either one preserves its data records.
    assert_member_unchanged(
        &source_bytes,
        candidate_bytes.as_deref(),
        "Data/sentinel.bin",
    );
    assert_member_unchanged(&source_bytes, candidate_bytes.as_deref(), "Data/movie.mov");
    assert_member_unchanged(&source_bytes, candidate_bytes.as_deref(), "Data/poster.png");
    if duplicate || movie_position != 2 {
        assert_member_unchanged(&source_bytes, candidate_bytes.as_deref(), "Data/audio.m4a");
    }

    // The inverse is a typed, exact-source operation.  Applying it from the
    // candidate restores the complete source bytes; applying the forward
    // patch twice from its original source is deterministic, while applying
    // it to the candidate is a stale-source rejection.
    let inverse = commit.patch().inverse();
    let restored_forward = package
        .apply_slide_media_lifecycle(&inverse.inverse())
        .unwrap_or_else(|error| panic!("double inverse must replay forward: {error}"));
    assert_eq!(package_bytes(restored_forward.package()), candidate_bytes);
    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&inverse)
        .unwrap_or_else(|error| panic!("lifecycle inverse must apply: {error}"));
    assert_eq!(
        package_bytes(restored.package()).as_deref(),
        Some(source_bytes.as_slice())
    );
    assert!(
        commit
            .package()
            .apply_slide_media_lifecycle(commit.patch())
            .is_err(),
        "a lifecycle patch must reject its post-state as a source"
    );

    let reapplied = package
        .apply_slide_media_lifecycle(commit.patch())
        .unwrap_or_else(|error| panic!("lifecycle forward patch must apply: {error}"));
    assert_eq!(package_bytes(reapplied.package()), candidate_bytes);
    assert_eq!(
        package_bytes(package).as_deref(),
        Some(source_bytes.as_slice())
    );
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
    .unwrap_or_else(|error| panic!("finite lifecycle input limit must construct: {error}"));
    if let Err(error) =
        Package::from_bytes_with_options(source, ReadOptions::new(tight, SemanticLimits::default()))
    {
        observe_error(error);
    } else {
        panic!("tight lifecycle input limit unexpectedly accepted source");
    }
}

fn assert_member_unchanged(before: &[u8], after: Option<&[u8]>, name: &str) {
    assert_eq!(
        member(before, name),
        after.and_then(|bytes| member(bytes, name)),
        "member {name} changed"
    );
}

fn member(source: &[u8], name: &str) -> Option<Vec<u8>> {
    Catalog::from_bytes(source)
        .ok()?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
}

fn package_bytes(package: &Package) -> Option<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).ok().map(|()| bytes)
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
