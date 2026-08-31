#![no_main]

//! Bounded selector-first Keynote chart-Arrange transaction fuzzing.
//!
//! Arbitrary bytes stay on the bounded Keynote ingress path and are also
//! interpreted as command bytes against the small source-built chart graphs
//! used by the chart-title target.  Reusing those package fixtures keeps the
//! target independent of ZIP mutation while retaining a real slide/chart
//! graph.  All transaction calls use archive-free public selectors and values;
//! native identifiers and wire objects never cross this boundary.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    ChartArrangement, ChartSelector, Limits, Package, ReadError, ReadOptions, SemanticLimits,
    SlideSelector,
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
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
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_chart_arrangement_selector_118__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_chart_arrangement_malformed_118__";

// These are tiny source-built packages shared with the chart-title target.
// They contain a valid slide/chart graph and deliberately avoid embedding a
// native Keynote fixture in this target's corpus.
const VALID_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_chart_title/valid.hex");
const VISIBLE_EMPTY_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/visible_empty.hex");
const HIDDEN_STALE_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/hidden_stale.hex");
const MALFORMED_PARENT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/malformed_parent.hex");
const MALFORMED_WIRE_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/malformed_wire.hex");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Arbitrary ZIP/IWA bytes seldom survive far enough to reach chart graph
    // semantics.  Run the same bytes as bounded command input against every
    // fixed source package so every mutation can exercise the public read,
    // edit, commit, patch, conflict, and inverse paths.
    for package in source_packages() {
        exercise_package(package, data);
    }
    exercise_malformed_packages();
    exercise_selector_validation();
    exercise_limit_guards();
    exercise_patch_conflicts();
    exercise_redacted_malformed_ingress();
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
        .unwrap_or_else(|error| unreachable!("valid chart-arrangement archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid chart-arrangement semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES
        .get_or_init(|| {
            [VALID_PACKAGE, VISIBLE_EMPTY_PACKAGE, HIDDEN_STALE_PACKAGE]
                .into_iter()
                .map(|bytes| {
                    Package::from_bytes_with_options(bytes, fuzz_options()).unwrap_or_else(
                        |error| panic!("source-built chart-arrangement package must open: {error}"),
                    )
                })
                .collect::<Vec<_>>()
                .into_boxed_slice()
        })
        .as_ref()
}

fn exercise_untrusted_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(read_u16(data, 0));
    let chart = ChartSelector::index(read_u16(data, 2));
    observe_result(package.slide_chart_arrangements(slide));
    observe_result(package.slide_chart_arrangement(slide, chart));
    observe_result(package.slide_chart_arrangement(
        SlideSelector::name(PRIVATE_SELECTOR),
        ChartSelector::index(0),
    ));
    assert_eq!(package_bytes(package), source);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let catalog = match package.slide_chart_catalog(SlideSelector::index(0)) {
        Ok(catalog) => catalog,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    if catalog.is_empty() {
        let batch = package
            .slide_chart_arrangements(0usize)
            .unwrap_or_else(|error| panic!("empty chart arrangement batch failed: {error}"));
        assert!(batch.is_empty());
        assert_eq!(package_bytes(package), source);
        return;
    }
    match package.slide_chart_catalog(SlideSelector::name("Charts")) {
        Ok(named_catalog) => assert_eq!(named_catalog, catalog),
        Err(error) => observe_error(error),
    }

    // Read every chart through both semantic selector forms.  The name route
    // is intentionally attempted only for non-empty titles; positional
    // selection remains the safe path for absent/empty native titles.
    let batch = package
        .slide_chart_arrangements(0usize)
        .unwrap_or_else(|error| panic!("chart arrangement batch failed: {error}"));
    assert_eq!(batch.len(), catalog.len());
    for descriptor in catalog.charts() {
        let position = descriptor.position();
        let by_position = package
            .slide_chart_arrangement(0usize, ChartSelector::index(position))
            .unwrap_or_else(|error| panic!("chart arrangement read failed: {error}"));
        assert_eq!(batch[position], by_position);
        if let Some(name) = descriptor.title().filter(|title| !title.is_empty()) {
            match package.slide_chart_arrangement(0usize, ChartSelector::name(name)) {
                Ok(by_name) => assert_eq!(by_name, by_position),
                Err(error) => observe_error(error),
            }
        }
    }
    assert_eq!(package_bytes(package), source);

    let chart_position = usize::from(control(data, 0)) % catalog.len();
    let chart_selector = if control(data, 1) & 1 == 0 {
        ChartSelector::index(chart_position)
    } else {
        catalog.charts()[chart_position].selector()
    };
    exercise_arrangement(package, chart_position, chart_selector, data, &source);
    exercise_selector_failures(package, &source);
}

fn exercise_arrangement(
    package: &Package,
    chart_position: usize,
    chart_selector: ChartSelector<'_>,
    data: &[u8],
    source: &[u8],
) {
    let before = match package.slide_chart_arrangement(0usize, chart_selector) {
        Ok(before) => before,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let edit = match package.edit_slide_chart_arrangement(0usize, chart_selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.before(), before);

    // Cover both independent setter methods, complete replacements, and an
    // exact no-op.  `ChartArrangement` is Copy, so no input-derived heap
    // value can escape the bounded command envelope.
    let edit = match control(data, 2) % 5 {
        0 => edit.set(before),
        1 => edit.set_locked(!before.locked()),
        2 => edit.set_constrain_proportions(!before.constrain_proportions()),
        3 => edit.set(ChartArrangement::new(true, false)),
        _ => edit.set(ChartArrangement::new(true, true)),
    };
    let requested = edit.after();
    assert_eq!(edit.after(), requested);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), requested);
    assert_eq!(patch.is_noop(), before == requested);
    let diagnostics = *commit.diagnostics();
    assert_eq!(diagnostics.changed(), before != requested);
    assert_eq!(diagnostics.full_reparse_performed(), before != requested);
    assert_eq!(diagnostics.touched_components() > 0, before != requested);
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), committed_bytes == source);
    assert_eq!(
        commit
            .package()
            .slide_chart_arrangement(0usize, chart_position)
            .unwrap_or_else(|error| panic!("committed chart arrangement read failed: {error}")),
        requested,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_slide_chart_arrangement(&patch)
        .unwrap_or_else(|error| panic!("fresh chart-arrangement patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        match applied.package().apply_slide_chart_arrangement(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed chart-arrangement patch must conflict with its target"),
        }
        match package.apply_slide_chart_arrangement(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed chart-arrangement inverse must conflict with its source"),
        }
    }
    let restored = applied
        .package()
        .apply_slide_chart_arrangement(&inverse)
        .unwrap_or_else(|error| panic!("fresh chart-arrangement inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_chart_arrangement(0usize, chart_position)
            .unwrap_or_else(|error| panic!("restored chart arrangement read failed: {error}")),
        before,
    );
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    let empty_slide = package
        .slide_chart_arrangement(SlideSelector::name(""), ChartSelector::index(0))
        .expect_err("empty slide selector was accepted");
    observe_error(empty_slide);
    let empty_chart = package
        .slide_chart_arrangement(0usize, ChartSelector::name(""))
        .expect_err("empty chart selector was accepted");
    observe_error(empty_chart);
    let missing_chart = package
        .slide_chart_arrangement(0usize, ChartSelector::index(usize::MAX))
        .expect_err("out-of-range chart selector was accepted");
    observe_error(missing_chart);
    assert_eq!(package_bytes(package), source);
}

fn exercise_selector_validation() {
    let package = &source_packages()[0];
    let source = package_bytes(package);
    let missing_slide = package
        .slide_chart_arrangement(SlideSelector::name(PRIVATE_SELECTOR), 0usize)
        .expect_err("missing private slide selector was accepted");
    observe_redacted(missing_slide, PRIVATE_SELECTOR);
    let missing_batch = package
        .slide_chart_arrangements(SlideSelector::name(PRIVATE_SELECTOR))
        .expect_err("missing private slide batch selector was accepted");
    observe_redacted(missing_batch, PRIVATE_SELECTOR);
    let missing_chart = package
        .slide_chart_arrangement(0usize, ChartSelector::index(usize::MAX))
        .expect_err("out-of-range chart selector was accepted");
    observe_error(missing_chart);
    assert_eq!(package_bytes(package), source);
}

fn exercise_malformed_packages() {
    static MALFORMED: OnceLock<Box<[Package]>> = OnceLock::new();
    let packages = MALFORMED.get_or_init(|| {
        [MALFORMED_PARENT_PACKAGE, MALFORMED_WIRE_PACKAGE]
            .into_iter()
            .map(|bytes| {
                Package::from_bytes_with_options(bytes, fuzz_options())
                    .unwrap_or_else(|error| panic!("malformed chart package must open: {error}"))
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    });
    for package in packages {
        let source = package_bytes(package);
        let read = package.slide_chart_arrangement(0usize, 0usize);
        assert!(read.is_err(), "malformed chart arrangement was accepted");
        if let Err(error) = read {
            observe_error(error);
        }
        let batch = package.slide_chart_arrangements(0usize);
        assert!(
            batch.is_err(),
            "malformed chart arrangement batch was accepted"
        );
        if let Err(error) = batch {
            observe_error(error);
        }
        let edit = package.edit_slide_chart_arrangement(
            SlideSelector::name("Charts"),
            ChartSelector::name(PRIVATE_SELECTOR),
        );
        assert!(edit.is_err(), "malformed chart arrangement edit was staged");
        if let Err(error) = edit {
            observe_redacted(error, PRIVATE_SELECTOR);
        }
        assert_eq!(package_bytes(package), source);
    }
}

fn exercise_limit_guards() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => panic!("oversized Keynote chart input returned wrong error: {error}"),
        Ok(_) => panic!("oversized Keynote chart input was accepted"),
    }

    let low_archive = Limits::new(1, 1, 1, 1, 1)
        .unwrap_or_else(|error| unreachable!("valid low chart archive limits: {error}"));
    match Package::from_bytes_with_options(
        VALID_PACKAGE,
        ReadOptions::new(low_archive, SemanticLimits::default()),
    ) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("low physical chart package limits were not enforced"),
    }

    let low_semantic = SemanticLimits::new(
        1,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid low chart semantic limits: {error}"));
    match Package::from_bytes_with_options(
        VALID_PACKAGE,
        ReadOptions::new(fuzz_options().archive(), low_semantic),
    ) {
        Err(error) => observe_error(error),
        Ok(package) => {
            let result = package.slide_chart_arrangement(0usize, 0usize);
            assert!(result.is_err(), "low chart semantic limits were bypassed");
            if let Err(error) = result {
                observe_error(error);
            }
            let batch = package.slide_chart_arrangements(0usize);
            assert!(batch.is_err(), "low chart batch limits were bypassed");
            if let Err(error) = batch {
                observe_error(error);
            }
        },
    }
}

fn exercise_patch_conflicts() {
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let source = &source_packages()[0];
        let before = source
            .slide_chart_arrangement(0usize, 0usize)
            .unwrap_or_else(|error| panic!("patch source chart must resolve: {error}"));
        let requested = before.with_locked(!before.locked());
        let commit = source
            .edit_slide_chart_arrangement(0usize, 0usize)
            .unwrap_or_else(|error| panic!("patch source chart edit failed: {error}"))
            .set(requested)
            .commit()
            .unwrap_or_else(|error| panic!("patch source chart commit failed: {error}"));
        let patch = commit.patch().clone();
        assert!(!patch.is_noop(), "patch-conflict source unexpectedly no-op");

        for destination in source_packages().iter().skip(1) {
            let destination_before = package_bytes(destination);
            let result = destination.apply_slide_chart_arrangement(&patch);
            assert!(result.is_err(), "cross-package chart patch was accepted");
            if let Err(error) = result {
                observe_error(error);
            }
            assert_eq!(package_bytes(destination), destination_before);
        }

        for bytes in [MALFORMED_PARENT_PACKAGE, MALFORMED_WIRE_PACKAGE] {
            let destination = Package::from_bytes_with_options(bytes, fuzz_options())
                .unwrap_or_else(|error| panic!("malformed patch destination must open: {error}"));
            let destination_before = package_bytes(&destination);
            let result = destination.apply_slide_chart_arrangement(&patch);
            assert!(
                result.is_err(),
                "chart-arrangement patch was accepted against malformed graph"
            );
            if let Err(error) = result {
                observe_error(error);
            }
            assert_eq!(package_bytes(&destination), destination_before);
        }
    });
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_options(PRIVATE_INPUT, fuzz_options()) {
        Err(error) => observe_redacted_bytes(error, PRIVATE_INPUT),
        Ok(_) => panic!("private malformed Keynote chart input was accepted"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a chart package to memory must succeed: {error}"));
    bytes
}

fn read_u16(data: &[u8], offset: usize) -> usize {
    usize::from(data.get(offset).copied().unwrap_or_default())
        | (usize::from(data.get(offset + 1).copied().unwrap_or_default()) << 8)
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn observe_result<T, E>(result: Result<T, E>)
where
    T: Debug,
    E: Debug + Display,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}

fn observe_redacted_bytes(error: impl Debug + Display, private: &[u8]) {
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private chart sentinel is UTF-8: {error}"));
    observe_redacted(error, private);
}
