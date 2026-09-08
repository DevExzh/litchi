#![no_main]

//! Bounded selector-first fuzzing for Keynote slide media properties.
//!
//! The target keeps native package details behind [`Package`] and exercises
//! the same transaction invariants as the focused integration tests: omitted
//! values remain distinct from explicit empty/false values, exact no-ops keep
//! the source artifact byte-for-byte unchanged, failed source checks publish
//! nothing, forward replay and inverse restoration are exact, and a reopened
//! candidate remains fully valid.  The native seeds deliberately use
//! source-order selectors for audio and file movies, and include a placeholder
//! fixture whose properties are readable but whose mutation is rejected
//! atomically.  This covers the complete movie sibling list rather than a
//! kind-specific renumbering.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_keynote::{
    Limits, MediaPart, MovieKind, MovieSelector, Package, ReadOptions, SemanticLimits,
    SlideMediaPropertiesError, SlideSelector, slide::media::MediaProperties,
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

const TARGET_INPUT: &[u8] = b"target-media-properties";
const NATIVE_KEYNOTE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_PLACEHOLDER: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/media-properties-placeholder-native.key");

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_seed(data);
    exercise_placeholder_seed();
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
        .unwrap_or_else(|error| unreachable!("valid media-properties archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid media-properties semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| {
                panic!("native Keynote media-properties seed must open: {error}")
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

fn placeholder_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_PLACEHOLDER, fuzz_options())
            .unwrap_or_else(|error| {
                panic!("native Keynote placeholder-properties seed must open: {error}")
            });
        verify_placeholder_contract(&package);
        package
    })
}

fn exercise_placeholder_seed() {
    let package = placeholder_package();
    let source_bytes = package_bytes(package);
    let kinds = read_movie_inventory(package, &source_bytes)
        .unwrap_or_else(|| panic!("native placeholder-properties inventory must be readable"));
    assert_eq!(
        kinds,
        vec![
            MovieKind::Audio,
            MovieKind::Audio,
            MovieKind::File,
            MovieKind::File,
            MovieKind::Placeholder,
        ]
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let Some(movie_kinds) = read_movie_inventory(package, &source_bytes) else {
        return;
    };
    // MovieSelector is source-order over every MovieArchive sibling.  The
    // baseline has AudioA at 0 and FileA at 2; exercising both catches any
    // accidental kind-specific renumbering in the focused adapter.
    for (selector, offset, sibling) in [
        (MovieSelector::index(0), 0_usize, MovieSelector::index(2)),
        (MovieSelector::index(2), 16_usize, MovieSelector::index(0)),
    ] {
        if !movie_kinds
            .get(selector.as_index())
            .is_some_and(|kind| matches!(kind, MovieKind::Audio | MovieKind::File))
        {
            continue;
        }
        exercise_selected(package, data, selector, sibling, offset, &source_bytes);
    }
}

fn verify_placeholder_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    package.validate().unwrap_or_else(|error| {
        panic!("native placeholder-properties seed must validate: {error}")
    });
    let show = package
        .show()
        .unwrap_or_else(|error| panic!("native placeholder-properties show must decode: {error}"));
    let slide = show
        .slides()
        .first()
        .unwrap_or_else(|| panic!("native placeholder-properties seed must contain a slide"));
    let movies = slide.movies();
    assert_eq!(movies.len(), 5);
    assert_eq!(
        movies.iter().map(|movie| movie.kind()).collect::<Vec<_>>(),
        vec![
            MovieKind::Audio,
            MovieKind::Audio,
            MovieKind::File,
            MovieKind::File,
            MovieKind::Placeholder,
        ]
    );
    assert_eq!(
        movies[4]
            .position()
            .map(|position| (position.x, position.y)),
        Some((321.0, 42.0))
    );
    assert_eq!(
        movies[4].size().map(|size| (size.width, size.height)),
        Some((640.0, 360.0))
    );
    assert_eq!(
        package
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(4))
            .unwrap_or_else(|error| panic!("native placeholder properties must read: {error}"))
            .accessibility_description(),
        Some("Native movie placeholder — accessible 北区")
    );
    let error =
        package.edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(4));
    assert!(matches!(
        error,
        Err(SlideMediaPropertiesError::WrongMediaKind)
    ));
    assert_eq!(package_bytes(package), source_bytes);
}

fn read_movie_inventory(package: &Package, source_bytes: &[u8]) -> Option<Vec<MovieKind>> {
    let show = match package.show() {
        Ok(show) => show,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return None;
        },
    };
    let Some(slide) = show.slides().first() else {
        observe_error("Keynote package has no slide");
        assert_eq!(package_bytes(package), source_bytes);
        return None;
    };
    let mut kinds = Vec::with_capacity(slide.movies().len());
    for (index, movie) in slide.movies().iter().enumerate() {
        kinds.push(movie.kind());
        if package
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(index))
            .is_err()
        {
            observe_error("movie properties read failed");
            assert_eq!(package_bytes(package), source_bytes);
            return None;
        }
        if matches!(movie.kind(), MovieKind::Placeholder | MovieKind::LiveVideo) {
            let error = package
                .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(index));
            assert!(
                matches!(error, Err(SlideMediaPropertiesError::WrongMediaKind)),
                "non-file/audio movie properties must remain read-only"
            );
            assert_eq!(package_bytes(package), source_bytes);
        }
    }
    Some(kinds)
}

fn exercise_selected(
    package: &Package,
    data: &[u8],
    selector: MovieSelector,
    untouched_selector: MovieSelector,
    offset: usize,
    source_bytes: &[u8],
) {
    let slide = SlideSelector::index(0);
    let before = match package.slide_media_properties(slide, selector) {
        Ok(properties) => properties,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let untouched_before = package
        .slide_media_properties(slide, untouched_selector)
        .ok();
    let desired = requested_properties(&before, data, offset);

    let edit = match package.edit_slide_media_properties(slide, selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let edit = match edit.set(desired.clone()) {
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
    assert_eq!(patch.before(), &before);
    assert_eq!(patch.after(), &desired);
    assert_eq!(patch.is_noop(), candidate_bytes == source_bytes);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_eq!(
        commit
            .package()
            .slide_media_properties(slide, selector)
            .unwrap_or_else(|error| panic!("committed media-properties read failed: {error}")),
        desired,
    );
    assert_untouched_media(
        commit.package(),
        slide,
        untouched_selector,
        untouched_before.as_ref(),
    );
    assert_media_bytes_unchanged(package, commit.package(), slide);
    commit.package().validate().unwrap_or_else(|error| {
        panic!("committed media-properties package must validate: {error}")
    });
    assert_eq!(package_bytes(package), source_bytes);

    let replay = package
        .apply_slide_media_properties(&patch)
        .unwrap_or_else(|error| panic!("media-properties forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    replay
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("replayed media-properties package must validate: {error}"));

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = commit
        .package()
        .apply_slide_media_properties(&inverse)
        .unwrap_or_else(|error| panic!("media-properties inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .slide_media_properties(slide, selector)
            .unwrap_or_else(|error| panic!("restored media-properties read failed: {error}")),
        before,
    );
    restored
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("restored media-properties package must validate: {error}"));

    if !patch.is_noop() {
        let candidate_before = package_bytes(commit.package());
        assert!(
            commit
                .package()
                .apply_slide_media_properties(&patch)
                .is_err(),
            "a changed media-properties patch must reject its post-state as a source"
        );
        assert_eq!(package_bytes(commit.package()), candidate_before);
    }

    exercise_noop(package, slide, selector, &before, source_bytes);
    assert_eq!(package_bytes(package), source_bytes);
}

fn verify_seed_success_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    for (selector, offset) in [
        (MovieSelector::index(0), 0_usize),
        (MovieSelector::index(2), 16_usize),
    ] {
        let before = package
            .slide_media_properties(slide, selector)
            .unwrap_or_else(|error| panic!("media-properties seed must be readable: {error}"));
        let target = seed_target_properties(offset);
        let commit = package
            .edit_slide_media_properties(slide, selector)
            .unwrap_or_else(|error| panic!("media-properties seed edit must prepare: {error}"))
            .set(target.clone())
            .unwrap_or_else(|error| panic!("media-properties seed target must stage: {error}"))
            .commit()
            .unwrap_or_else(|error| panic!("media-properties seed target must commit: {error}"));
        assert_ne!(
            before, target,
            "the mandatory media-properties seed target must change the source"
        );
        assert!(!commit.patch().is_noop());
        let candidate_bytes = package_bytes(commit.package());
        assert_ne!(candidate_bytes, source_bytes);
        assert_eq!(
            commit
                .package()
                .slide_media_properties(slide, selector)
                .unwrap_or_else(|error| panic!("seed target must read back: {error}")),
            target,
        );
        commit
            .package()
            .validate()
            .unwrap_or_else(|error| panic!("seed target package must validate: {error}"));

        let replay = package
            .apply_slide_media_properties(commit.patch())
            .unwrap_or_else(|error| panic!("seed target forward replay must apply: {error}"));
        assert_eq!(package_bytes(replay.package()), candidate_bytes);
        let restored = commit
            .package()
            .apply_slide_media_properties(&commit.patch().inverse())
            .unwrap_or_else(|error| panic!("seed target inverse must apply: {error}"));
        assert_eq!(package_bytes(restored.package()), source_bytes);

        exercise_noop(package, slide, selector, &before, &source_bytes);
    }
    assert_eq!(package_bytes(package), source_bytes);
    exercise_limit_budget();
}

fn exercise_noop(
    package: &Package,
    slide: SlideSelector,
    selector: MovieSelector,
    before: &MediaProperties,
    source_bytes: &[u8],
) {
    let noop = package
        .edit_slide_media_properties(slide, selector)
        .unwrap_or_else(|error| panic!("media-properties no-op edit must prepare: {error}"))
        .set(before.clone())
        .unwrap_or_else(|error| panic!("media-properties no-op target must stage: {error}"))
        .commit()
        .unwrap_or_else(|error| panic!("media-properties no-op must commit: {error}"));
    assert!(noop.patch().is_noop());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(package_bytes(noop.package()), source_bytes);
    noop.package()
        .validate()
        .unwrap_or_else(|error| panic!("media-properties no-op package must validate: {error}"));
    let noop_replay = package
        .apply_slide_media_properties(noop.patch())
        .unwrap_or_else(|error| panic!("media-properties no-op replay must apply: {error}"));
    assert_eq!(package_bytes(noop_replay.package()), source_bytes);
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
    .unwrap_or_else(|error| panic!("finite media-properties semantic limit must build: {error}"));
    let options = ReadOptions::new(Limits::default(), semantic);
    match Package::from_bytes_with_options(NATIVE_KEYNOTE, options) {
        Err(error) => observe_error(error),
        Ok(package) => {
            let source_bytes = package_bytes(&package);
            assert!(
                package
                    .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
                    .is_err()
            );
            assert_eq!(package_bytes(&package), source_bytes);
        },
    }
}

fn requested_properties(before: &MediaProperties, data: &[u8], offset: usize) -> MediaProperties {
    if data.starts_with(TARGET_INPUT) {
        return seed_target_properties(offset);
    }
    if control(data, offset) % 7 == 0 {
        return before.clone();
    }
    MediaProperties::new()
        .with_hyperlink_url(requested_string(data, offset.saturating_add(1)))
        .with_locked(requested_bool(data, offset.saturating_add(3)))
        .with_aspect_ratio_locked(requested_bool(data, offset.saturating_add(4)))
        .with_accessibility_description(requested_string(data, offset.saturating_add(5)))
}

fn seed_target_properties(offset: usize) -> MediaProperties {
    let suffix = if offset == 0 { "audio-a" } else { "file-a" };
    MediaProperties::new()
        .with_hyperlink_url(Some(format!("https://example.test/{suffix}")))
        .with_locked(Some(true))
        .with_aspect_ratio_locked(Some(true))
        .with_accessibility_description(Some(format!("{suffix} — accessible 北区 🎵")))
}

fn requested_string(data: &[u8], offset: usize) -> Option<String> {
    match control(data, offset) % 6 {
        0 => None,
        1 => Some(String::new()),
        2 => Some("https://example.test/media".to_owned()),
        3 => Some("説明 🎵 北区".to_owned()),
        _ => {
            let start = offset.saturating_add(1).min(data.len());
            let end = start.saturating_add(48).min(data.len());
            Some(String::from_utf8_lossy(&data[start..end]).into_owned())
        },
    }
}

fn requested_bool(data: &[u8], offset: usize) -> Option<bool> {
    match control(data, offset) % 3 {
        0 => None,
        1 => Some(false),
        _ => Some(true),
    }
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn assert_untouched_media(
    package: &Package,
    slide: SlideSelector,
    selector: MovieSelector,
    expected: Option<&MediaProperties>,
) {
    if let Some(expected) = expected {
        assert_eq!(
            package
                .slide_media_properties(slide, selector)
                .unwrap_or_else(|error| panic!("untouched media-properties read failed: {error}")),
            expected.clone(),
            "the sibling media properties must remain untouched"
        );
    }
}

fn assert_media_bytes_unchanged(before: &Package, after: &Package, slide: SlideSelector) {
    for movie in 0..4 {
        for part in [MediaPart::Content, MediaPart::Poster] {
            let expected = before
                .slide_media_data(slide, MovieSelector::index(movie), part)
                .ok()
                .map(ToOwned::to_owned);
            let Some(expected) = expected else {
                continue;
            };
            let actual = after
                .slide_media_data(slide, MovieSelector::index(movie), part)
                .ok()
                .map(ToOwned::to_owned);
            assert_eq!(
                actual,
                Some(expected),
                "media bytes changed for source-order media {movie} part {part:?}"
            );
        }
    }
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
