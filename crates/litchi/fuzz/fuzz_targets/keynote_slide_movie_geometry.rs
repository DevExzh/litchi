#![no_main]

//! Bounded selector-first fuzzing for existing Keynote movie geometry.
//!
//! Geometry is intentionally a scalar leaf: position and displayed size are
//! public, while native flags/angle and all movie graph records stay opaque.
//! The target therefore exercises read/set/no-op, exact patch replay,
//! conflicts, inverse restoration, and source atomicity without manufacturing
//! native movie identifiers.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::media::{Point, Size, geometry::MovieGeometry},
};

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
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

fuzz_target!(|data: &[u8]| {
    if data.len() <= usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX)
        && let Ok(package) = Package::from_bytes_with_options(data, options())
    {
        exercise_package(&package, data);
    }

    // CRC-protected ZIP mutations seldom preserve enough native graph state
    // to reach a movie. Reuse the checked-in Keynote source for deterministic
    // selector, transaction, conflict, and limit coverage.
    exercise_package(native_package(), data);
    exercise_selector_errors(native_package());
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
        .unwrap_or_else(|error| unreachable!("valid Keynote geometry archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote geometry semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_KEYNOTE, options())
            .unwrap_or_else(|error| panic!("native Keynote geometry seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    let movie = MovieSelector::index(usize::from(control(data, 0) % 4));
    let before = match package.slide_movie_geometry(slide, movie) {
        Ok(Some(value)) => value,
        Ok(None) => return,
        Err(error) => {
            observe_error(error);
            return;
        },
    };

    let edit = match package.edit_slide_movie_geometry(slide, movie) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let replacement = replacement(before, data);
    let edit = match if control(data, 1) & 1 == 0 {
        edit.set(before)
    } else {
        edit.set(replacement)
    } {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let desired = if control(data, 1) & 1 == 0 {
        before
    } else {
        replacement
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch().clone();
    let candidate_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), desired);
    assert_eq!(patch.is_noop(), candidate_bytes == source_bytes);
    if !patch.is_noop() {
        assert!(
            commit.diagnostics().deleted_previews() > 0,
            "visual geometry changes must invalidate canonical previews"
        );
    }
    assert_eq!(
        commit
            .package()
            .slide_movie_geometry(slide, movie)
            .unwrap_or_else(|error| panic!("geometry candidate read failed: {error}")),
        Some(desired),
    );

    let applied = match package.apply_slide_movie_geometry(&patch) {
        Ok(applied) => applied,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), candidate_bytes);
    assert_eq!(patch.inverse().inverse(), patch);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_slide_movie_geometry(&patch)
                .is_err(),
            "geometry patch unexpectedly applied twice"
        );
        assert!(
            package
                .apply_slide_movie_geometry(&patch.inverse())
                .is_err(),
            "geometry inverse unexpectedly applied to source"
        );
    }
    let restored = applied
        .package()
        .apply_slide_movie_geometry(&patch.inverse())
        .unwrap_or_else(|error| panic!("geometry inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .slide_movie_geometry(slide, movie)
            .unwrap_or_else(|error| panic!("restored geometry read failed: {error}")),
        Some(before),
    );
    black_box(commit.diagnostics());
}

fn replacement(before: MovieGeometry, data: &[u8]) -> MovieGeometry {
    let position = Point {
        x: before.position().x + f32::from(control(data, 2) % 16) + 1.0,
        y: before.position().y + f32::from(control(data, 3) % 16) + 1.0,
    };
    let size = Size {
        width: before.size().width + f32::from(control(data, 4) % 32) + 1.0,
        height: before.size().height + f32::from(control(data, 5) % 32) + 1.0,
    };
    MovieGeometry::new(position, size)
        .unwrap_or_else(|error| panic!("bounded geometry replacement must validate: {error}"))
}

fn exercise_selector_errors(package: &Package) {
    let result =
        package.slide_movie_geometry(SlideSelector::index(usize::MAX), MovieSelector::index(0));
    assert!(result.is_err());
    let result = package
        .edit_slide_movie_geometry(SlideSelector::index(0), MovieSelector::index(usize::MAX));
    assert!(result.is_err());
}

fn exercise_input_limit() {
    let defaults = Limits::default();
    let tight = Limits::new(
        u64::try_from(NATIVE_KEYNOTE.len().saturating_sub(1)).unwrap_or(0),
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )
    .unwrap_or_else(|error| panic!("finite limit must construct: {error}"));
    assert!(
        Package::from_bytes_with_options(
            NATIVE_KEYNOTE,
            ReadOptions::new(tight, SemanticLimits::default()),
        )
        .is_err()
    );
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote geometry package must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
