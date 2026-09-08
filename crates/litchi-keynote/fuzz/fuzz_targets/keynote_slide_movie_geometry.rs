#![no_main]

//! Bounded selector-first fuzzing for Keynote movie geometry.
//!
//! The target exercises the semantic geometry transaction against arbitrary
//! bounded package bytes and a native source-order fixture.  A successful
//! edit composes placement, rotation, reflection, and original-size restore,
//! then checks exact replay and inverse restoration.  Failed admission must
//! leave the immutable source package unchanged.

mod support;

use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_keynote::{
    MovieSelector, Package, SlideSelector,
    slide::media::{
        Point, Size,
        geometry::{MovieFlipAxis, MovieGeometry, MovieTransform},
    },
};

use self::support::{
    MAX_INPUT_BYTES, control, finite_axis, observe_error, package_bytes, read_options,
};

const TARGET_INPUT: &[u8] = b"target-movie-geometry";
const NATIVE_KEYNOTE: &[u8] = include_bytes!(
    "../../../../test-data/iwork/keynote/slide-movie-creation-fresh-focused-native.key"
);

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_native_seed(data);
});

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, read_options())
            .unwrap_or_else(|error| panic!("native Keynote geometry seed must open: {error}"));
        verify_native_contract(&package);
        package
    })
}

fn verify_native_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    let movie = MovieSelector::index(4);
    assert_eq!(
        package.slide_movie_geometry(slide, movie),
        Ok(Some(
            MovieGeometry::new(
                Point { x: 321.0, y: 42.0 },
                Size {
                    width: 640.0,
                    height: 360.0,
                },
            )
            .unwrap_or_else(|error| panic!("native geometry seed must be valid: {error}")),
        ))
    );
    let edit = package
        .edit_slide_movie_geometry(slide, movie)
        .unwrap_or_else(|error| panic!("native geometry seed edit must prepare: {error}"));
    let original_size = edit
        .original_size()
        .unwrap_or_else(|| panic!("native geometry seed must retain original dimensions"));
    let commit = edit
        .set(
            MovieGeometry::new(
                Point {
                    x: 450.5,
                    y: 120.25,
                },
                Size {
                    width: 480.0,
                    height: 270.0,
                },
            )
            .unwrap_or_else(|error| panic!("native geometry target must be valid: {error}")),
        )
        .and_then(|edit| {
            edit.set_transform(
                MovieTransform::new(27.5, false).unwrap_or_else(|error| {
                    panic!("native transform target must be valid: {error}")
                }),
            )
        })
        .and_then(|edit| edit.flip(MovieFlipAxis::Horizontal))
        .and_then(|edit| edit.restore_original_size())
        .and_then(|edit| edit.commit())
        .unwrap_or_else(|error| panic!("native geometry seed transaction must commit: {error}"));
    assert_eq!(
        commit.patch().after(),
        MovieGeometry::new(
            Point {
                x: 450.5,
                y: 120.25,
            },
            original_size,
        )
        .unwrap_or_else(|error| panic!("native restored geometry must be valid: {error}"))
    );
    assert_eq!(
        commit.patch().after_transform(),
        MovieTransform::new(27.5, true)
            .unwrap_or_else(|error| panic!("native flipped transform must be valid: {error}"))
    );
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("native geometry candidate must validate: {error}"));
    let candidate_bytes = package_bytes(commit.package());
    let replay = package
        .apply_slide_movie_geometry(commit.patch())
        .unwrap_or_else(|error| panic!("native geometry forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    let restored = commit
        .package()
        .apply_slide_movie_geometry(&commit.patch().inverse())
        .unwrap_or_else(|error| panic!("native geometry inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    let twice = package
        .apply_slide_movie_geometry(&commit.patch().inverse().inverse())
        .unwrap_or_else(|error| panic!("native geometry double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }
    match Package::from_bytes_with_options(data, read_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }
}

fn exercise_native_seed(data: &[u8]) {
    exercise_package(native_package(), data);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let slide = SlideSelector::index(0);
    let movie = requested_movie(data);

    let before = match package.slide_movie_geometry(slide, movie) {
        Ok(Some(geometry)) => geometry,
        Ok(None) => {
            // A source may expose an admitted movie without a displayed
            // geometry.  The edit path must reject it without publication.
            observe_error("movie geometry is absent");
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let before_transform = match package.slide_movie_transform(slide, movie) {
        Ok(Some(transform)) => transform,
        Ok(None) => {
            observe_error("movie transform is absent");
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
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
    let original_size = edit.original_size();
    let desired = requested_geometry(data);
    let should_restore = requested_restore(data) && original_size.is_some();
    let desired_transform = requested_transform(data, before_transform);

    let edit = match edit.set(desired) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let edit = match edit.set_transform(desired_transform) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let edit = if requested_flip(data) {
        match edit.flip(requested_flip_axis(data)) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_eq!(package_bytes(package), source_bytes);
                return;
            },
        }
    } else {
        edit
    };
    let expected_transform = edit.after_transform();
    let edit = if should_restore {
        match edit.restore_original_size() {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_eq!(package_bytes(package), source_bytes);
                return;
            },
        }
    } else {
        edit
    };
    let expected_geometry = edit.after();

    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch();
    let candidate_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), expected_geometry);
    assert_eq!(patch.before_transform(), before_transform);
    assert_eq!(patch.after_transform(), expected_transform);
    assert_eq!(patch.is_noop(), candidate_bytes == source_bytes);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_eq!(
        commit.package().slide_movie_geometry(slide, movie),
        Ok(Some(expected_geometry))
    );
    assert_eq!(
        commit.package().slide_movie_transform(slide, movie),
        Ok(Some(expected_transform))
    );
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("geometry candidate must validate: {error}"));
    assert_eq!(package_bytes(package), source_bytes);

    let replay = match package.apply_slide_movie_geometry(patch) {
        Ok(commit) => commit,
        Err(error) => panic!("geometry forward replay must apply: {error}"),
    };
    assert_eq!(package_bytes(replay.package()), candidate_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch.clone());
    let restored = match commit.package().apply_slide_movie_geometry(&inverse) {
        Ok(commit) => commit,
        Err(error) => panic!("geometry inverse must apply: {error}"),
    };
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored.package().slide_movie_geometry(slide, movie),
        Ok(Some(before))
    );
    assert_eq!(
        restored.package().slide_movie_transform(slide, movie),
        Ok(Some(before_transform))
    );

    let twice = package
        .apply_slide_movie_geometry(&inverse.inverse())
        .unwrap_or_else(|error| panic!("geometry double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    if !patch.is_noop() {
        assert!(
            commit.package().apply_slide_movie_geometry(patch).is_err(),
            "a changed geometry patch must reject its post-state as a source"
        );
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn requested_movie(data: &[u8]) -> MovieSelector {
    if data.starts_with(TARGET_INPUT) {
        // The fresh native fixture has two audio entries followed by the
        // existing movie entries and a newly authored movie at source index 4.
        // Keeping this selector in the all-media source order makes the seed
        // exercise the regression that previously filtered audio siblings.
        return MovieSelector::index(4);
    }
    MovieSelector::index(usize::from(control(data, 0) % 8))
}

fn requested_geometry(data: &[u8]) -> MovieGeometry {
    MovieGeometry::new(
        Point {
            x: finite_axis(data, 1, 450.5),
            y: finite_axis(data, 5, 120.25),
        },
        Size {
            width: positive_axis(data, 9, 480.0),
            height: positive_axis(data, 13, 270.0),
        },
    )
    .unwrap_or_else(|error| panic!("bounded geometry must be valid: {error}"))
}

fn requested_transform(data: &[u8], before: MovieTransform) -> MovieTransform {
    if data.starts_with(TARGET_INPUT) {
        return MovieTransform::new(27.5, false)
            .unwrap_or_else(|error| panic!("fixed transform must be valid: {error}"));
    }
    let angle = finite_axis(data, 17, before.angle_degrees());
    MovieTransform::new(angle, control(data, 21) & 1 != 0)
        .unwrap_or_else(|error| panic!("bounded transform must be valid: {error}"))
}

fn requested_restore(data: &[u8]) -> bool {
    data.starts_with(TARGET_INPUT) || control(data, 22) & 1 != 0
}

fn requested_flip(data: &[u8]) -> bool {
    data.starts_with(TARGET_INPUT) || control(data, 23) & 1 != 0
}

fn requested_flip_axis(data: &[u8]) -> MovieFlipAxis {
    if control(data, 24) & 1 == 0 {
        MovieFlipAxis::Horizontal
    } else {
        MovieFlipAxis::Vertical
    }
}

fn positive_axis(data: &[u8], offset: usize, fallback: f32) -> f32 {
    let value = finite_axis(data, offset, fallback).abs();
    if value > f32::EPSILON {
        value
    } else {
        fallback
    }
}
