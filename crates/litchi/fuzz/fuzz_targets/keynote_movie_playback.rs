#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;
use std::time::Duration;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::media::{MediaLoopMode, MediaPlaybackSettings, MediaVolume},
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
    if data.len() <= usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        if let Ok(package) = Package::from_bytes_with_options(data, fuzz_options()) {
            exercise_package(&package, data);
        }
    }

    // The tracked basic.key seed is a valid package even when it has no
    // file-backed movie. If a future seed gains one, the same harness reaches
    // the full selector/commit/inverse route without changing the boundary.
    exercise_package(native_package(), data);
    exercise_input_limit();
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| panic!("native Keynote fuzz seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let movie_position = usize::from(control(data, 0) % 4);
    let selector = (
        SlideSelector::index(0),
        MovieSelector::index(movie_position),
    );
    let before = match package.slide_movie_playback_settings(selector.0, selector.1) {
        Ok(Some(settings)) => settings,
        Ok(None) => return,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let source_bytes = package_bytes(package);
    let replacement = changed_settings(before, data);
    let edit = match package.edit_slide_movie_playback_settings(selector.0, selector.1) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let edit = match edit.set(if control(data, 1) & 1 == 0 {
        before
    } else {
        replacement
    }) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
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
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == committed_bytes);
    assert_eq!(
        commit
            .package()
            .slide_movie_playback_settings(selector.0, selector.1)
            .unwrap_or_else(|error| panic!("committed playback read failed: {error}")),
        Some(patch.after()),
    );

    let applied = match package.apply_slide_movie_playback_settings(&patch) {
        Ok(applied) => applied,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_movie_playback_settings(&inverse)
        .unwrap_or_else(|error| panic!("playback inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .slide_movie_playback_settings(selector.0, selector.1)
            .unwrap_or_else(|error| panic!("restored playback read failed: {error}")),
        Some(before),
    );

    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_slide_movie_playback_settings(&patch)
                .is_err()
        );
        assert!(
            package
                .apply_slide_movie_playback_settings(&inverse)
                .is_err()
        );
    }
}

fn changed_settings(before: MediaPlaybackSettings, data: &[u8]) -> MediaPlaybackSettings {
    let end = 4.0 + f32::from(control(data, 2) % 80) / 10.0;
    let start = f32::from(control(data, 3) % 20) / 100.0;
    let end = end.max(start + 0.1);
    MediaPlaybackSettings::new(Duration::from_secs_f32(end))
        .with_start_time((control(data, 4) & 1 == 0).then(|| Duration::from_secs_f32(start)))
        .with_poster_time(
            (control(data, 5) & 1 == 0)
                .then(|| Duration::from_secs_f32(f32::from(control(data, 6) % 50) / 10.0)),
        )
        .with_loop_mode(Some(match before.loop_mode {
            Some(MediaLoopMode::Repeat) => MediaLoopMode::BackAndForth,
            _ => MediaLoopMode::Repeat,
        }))
        .with_volume(
            MediaVolume::new(f32::from(control(data, 7)) / 255.0)
                .ok()
                .or(before.volume),
        )
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
        .unwrap_or_else(|error| panic!("writing Keynote fuzz package must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
