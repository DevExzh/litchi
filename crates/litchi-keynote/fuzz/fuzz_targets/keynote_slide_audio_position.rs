#![no_main]

//! Bounded selector-first fuzzing for the Keynote slide-audio position API.
//!
//! The target keeps the native package hidden behind [`Package`] and exercises
//! the same invariants as the focused integration tests: finite semantic
//! points, exact no-ops, source-preserving failed edits, forward replay,
//! inverse restoration, stale-source rejection, wrong-media selectors, and a
//! deliberately tight semantic budget.  A checked-in native package keeps the
//! mutation path reachable without copying a large synthetic archive into the
//! fuzz crate.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_keynote::{
    Limits, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector, slide::media::Point,
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
// Keep arbitrary inputs small while allowing the native seed's complete
// transaction ledger (selection, validation, replay, and retained artifacts).
const MAX_PACKAGE_BYTES: u64 = 8 * 1024 * 1024;
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

const TARGET_INPUT: &[u8] = b"target-1120-420";
const NATIVE_KEYNOTE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/media-comments-baseline-native.key");

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_seed(data);
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_PACKAGE_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid audio-position archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid audio-position semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| {
                panic!("native Keynote audio-position seed must open: {error}")
            });
        verify_seed_success_contract(&package);
        package
    })
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }

    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }
}

fn exercise_seed(data: &[u8]) {
    exercise_package(native_package(), data);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    let audio = MovieSelector::index(usize::from(control(data, 0) % 4));
    let before = match package.slide_audio_position(slide, audio) {
        Ok(point) => point,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let desired = requested_point(before, data);

    let edit = match package.edit_slide_audio_position(slide, audio) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let edit = match edit.set(desired) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
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
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_eq!(
        commit
            .package()
            .slide_audio_position(slide, audio)
            .unwrap_or_else(|error| panic!("committed audio-position read failed: {error}")),
        desired,
    );
    assert_eq!(package_bytes(package), source_bytes);

    let replay = package
        .apply_slide_audio_position(&patch)
        .unwrap_or_else(|error| panic!("audio-position forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = commit
        .package()
        .apply_slide_audio_position(&inverse)
        .unwrap_or_else(|error| panic!("audio-position inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .slide_audio_position(slide, audio)
            .unwrap_or_else(|error| panic!("restored audio-position read failed: {error}")),
        before,
    );

    if !patch.is_noop() {
        assert!(
            commit.package().apply_slide_audio_position(&patch).is_err(),
            "a changed audio-position patch must reject its post-state as a source"
        );
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn verify_seed_success_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    let audio = MovieSelector::index(0);
    let before = package
        .slide_audio_position(slide, audio)
        .unwrap_or_else(|error| panic!("audio-position seed must be readable: {error}"));
    let target = Point {
        x: 1_120.0,
        y: 420.0,
    };
    let commit = package
        .edit_slide_audio_position(slide, audio)
        .unwrap_or_else(|error| panic!("audio-position seed edit must prepare: {error}"))
        .set(target)
        .unwrap_or_else(|error| panic!("audio-position seed target must be finite: {error}"))
        .commit()
        .unwrap_or_else(|error| panic!("audio-position seed target must commit: {error}"));
    assert_ne!(
        before, target,
        "the mandatory seed target must change the source"
    );
    assert!(!commit.patch().is_noop());
    let candidate_bytes = package_bytes(commit.package());
    assert_ne!(candidate_bytes, source_bytes);
    assert_eq!(
        commit
            .package()
            .slide_audio_position(slide, audio)
            .unwrap_or_else(|error| panic!("seed target must read back: {error}")),
        target,
    );

    let replay = package
        .apply_slide_audio_position(commit.patch())
        .unwrap_or_else(|error| panic!("seed target forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    let restored = commit
        .package()
        .apply_slide_audio_position(&commit.patch().inverse())
        .unwrap_or_else(|error| panic!("seed target inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);

    let noop = package
        .edit_slide_audio_position(slide, audio)
        .unwrap_or_else(|error| panic!("seed no-op edit must prepare: {error}"))
        .set(before)
        .unwrap_or_else(|error| panic!("seed no-op point must be finite: {error}"))
        .commit()
        .unwrap_or_else(|error| panic!("seed no-op must commit: {error}"));
    assert!(noop.patch().is_noop());
    assert_eq!(package_bytes(noop.package()), source_bytes);
    let noop_replay = package
        .apply_slide_audio_position(noop.patch())
        .unwrap_or_else(|error| panic!("seed no-op replay must apply: {error}"));
    assert_eq!(package_bytes(noop_replay.package()), source_bytes);

    for point in [
        Point {
            x: f32::NAN,
            y: 0.0,
        },
        Point {
            x: 0.0,
            y: f32::INFINITY,
        },
        Point {
            x: f32::NEG_INFINITY,
            y: 0.0,
        },
    ] {
        assert!(package.edit_slide_audio_position(slide, audio).is_ok());
        assert!(
            package
                .edit_slide_audio_position(slide, audio)
                .unwrap_or_else(|error| panic!("finite seed edit must prepare: {error}"))
                .set(point)
                .is_err()
        );
        assert_eq!(package_bytes(package), source_bytes);
    }

    exercise_wrong_kind(package, &source_bytes);
    exercise_limit_budget();
}

fn exercise_wrong_kind(package: &Package, source_bytes: &[u8]) {
    assert!(
        package
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(2),)
            .is_err(),
        "file movie unexpectedly admitted as slide audio position edit"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_limit_budget() {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| panic!("finite audio-position semantic limit must build: {error}"));
    let options = ReadOptions::new(Limits::default(), semantic);
    match Package::from_bytes_with_options(NATIVE_KEYNOTE, options) {
        Err(error) => observe_error(error),
        Ok(package) => {
            let source_bytes = package_bytes(&package);
            assert!(
                package
                    .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                    .is_err()
            );
            assert_eq!(package_bytes(&package), source_bytes);
        },
    }
}

fn requested_point(before: Point, data: &[u8]) -> Point {
    if data.starts_with(TARGET_INPUT) {
        return Point {
            x: 1_120.0,
            y: 420.0,
        };
    }

    match control(data, 0) % 8 {
        0 => before,
        1 => Point {
            x: finite_axis(data, 1, before.x),
            y: finite_axis(data, 5, before.y),
        },
        2 => Point {
            x: f32::NAN,
            y: before.y,
        },
        3 => Point {
            x: before.x,
            y: f32::INFINITY,
        },
        4 => Point {
            x: f32::NEG_INFINITY,
            y: before.y,
        },
        _ => Point {
            x: finite_axis(data, 1, 0.0),
            y: finite_axis(data, 5, 0.0),
        },
    }
}

fn finite_axis(data: &[u8], offset: usize, fallback: f32) -> f32 {
    let bits = u32::from(control(data, offset))
        | (u32::from(control(data, offset.saturating_add(1))) << 8)
        | (u32::from(control(data, offset.saturating_add(2))) << 16)
        | (u32::from(control(data, offset.saturating_add(3))) << 24);
    let value = f32::from_bits(bits);
    if value.is_finite() {
        let value = value.clamp(-1_000_000.0, 1_000_000.0);
        if value == 0.0 { 0.0 } else { value }
    } else {
        fallback
    }
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote fuzz package must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
