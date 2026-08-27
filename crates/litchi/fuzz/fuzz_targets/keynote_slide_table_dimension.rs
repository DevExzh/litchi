#![no_main]

//! Bounded selector-first Keynote slide-table row/column dimension fuzzing.
//!
//! Arbitrary bytes exercise bounded Keynote package ingress.  The same input
//! is also interpreted as a command stream against the small source-built
//! table packages used by the neighboring header and lock targets, keeping
//! both dimension axes, no-op/reset/set, apply/conflict/inverse, reopen, lock,
//! selector, and source-atomic failure paths reachable.  The checked-in
//! inputs are command recipes, not native package members.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::table::{
        TableSelector,
        dimension::{Dimension, Points, Size},
    },
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
const MAX_COMMAND_BYTES: usize = 1024;
const PRIVATE_SLIDE_NAME: &str = "__litchi_private_keynote_dimension_slide_missing__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_dimension_input_108__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

fuzz_target!(|data: &[u8]| {
    // Keep arbitrary input on the bounded package-ingress path even when the
    // input is also a command recipe for the fixed source packages below.
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, data),
        Err(error) => observe_error(error),
    }

    let command = command_input(data);
    for package in source_packages() {
        exercise_package(package, &command);
    }
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_archive_limit();
    exercise_semantic_limit(&command);
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
        .unwrap_or_else(|error| unreachable!("valid Keynote dimension archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote dimension semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [source_built_bytes(), locked_bytes()]
            .into_iter()
            .map(|source| {
                Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(|error| {
                    panic!("source-built Keynote dimension package must open: {error}")
                })
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    })
}

fn source_built_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            decode_hex(SOURCE_BUILT_PACKAGE)
                .unwrap_or_else(|| panic!("source-built Keynote dimension package has invalid hex"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            decode_hex(LOCKED_PACKAGE)
                .unwrap_or_else(|| panic!("locked Keynote dimension package has invalid hex"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES as usize {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn exercise_untrusted_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(usize::from(read_u16(data, 0)));
    let table = TableSelector::index(usize::from(read_u16(data, 2)));
    observe_result(package.slide_table_dimension_size(slide, table, Dimension::Row(0)));
    observe_result(package.slide_table_dimension_size(slide, table, Dimension::Column(0)));
    assert_eq!(package_bytes(package), source);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    exercise_selector_failures(package, &source);

    // Row(0) and Column(0) deliberately use valid-looking positions so both
    // axes reach the owner when the source package admits the table graph.
    exercise_dimension(package, Dimension::Row(0), data, &source);
    exercise_dimension(package, Dimension::Column(0), data, &source);
    assert_eq!(package_bytes(package), source);
}

fn exercise_dimension(package: &Package, dimension: Dimension, data: &[u8], source: &[u8]) {
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);
    let before = match package.slide_table_dimension_size(slide, table, dimension) {
        Ok(size) => size,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(package_bytes(package), source);

    exercise_noop(package, slide, table, dimension, before, source);
    exercise_changed(package, slide, table, dimension, before, data, source);
    exercise_reset(package, slide, table, dimension, before, source);
}

fn exercise_noop(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    dimension: Dimension,
    before: Size,
    source: &[u8],
) {
    let edit = match package.edit_slide_table_dimension_size(slide, table, dimension) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.before(), before);
    assert_eq!(edit.dimension(), dimension);
    let commit = match edit.set(before).commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    assert!(patch.is_noop());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), before);
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(commit.package()), source);
    let applied = match package.apply_slide_table_dimension_size(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), source);
}

fn exercise_changed(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    dimension: Dimension,
    before: Size,
    data: &[u8],
    source: &[u8],
) {
    let after = changed_size(before, data, dimension);
    let edit = match package.edit_slide_table_dimension_size(slide, table, dimension) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let commit = match edit.set(after).commit() {
        Ok(commit) => commit,
        Err(error) => {
            // A locked table and a source whose strict graph is unsupported
            // are expected typed failures; every failure remains atomic.
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    let target = package_bytes(commit.package());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert!(!patch.is_noop());
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().touched_components() > 0);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(package), source);
    assert_eq!(
        commit
            .package()
            .slide_table_dimension_size(slide, table, dimension)
            .unwrap_or_else(|error| panic!("dimension candidate readback failed: {error}")),
        after,
    );

    let applied = match package.apply_slide_table_dimension_size(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), target);
    match commit.package().apply_slide_table_dimension_size(&patch) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("changed dimension patch unexpectedly applied twice"),
    }
    match package.apply_slide_table_dimension_size(&patch.inverse()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("dimension inverse unexpectedly applied to source"),
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_dimension_size(&inverse)
        .unwrap_or_else(|error| panic!("dimension inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_table_dimension_size(slide, table, dimension)
            .unwrap_or_else(|error| panic!("dimension inverse readback failed: {error}")),
        before,
    );
    black_box((
        patch.path(),
        patch.dimension(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
    ));
}

fn exercise_reset(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    dimension: Dimension,
    before: Size,
    source: &[u8],
) {
    let edit = match package.edit_slide_table_dimension_size(slide, table, dimension) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let commit = match edit.reset().commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), Size::Default);
    assert_eq!(patch.is_noop(), before == Size::Default);
    assert_eq!(commit.diagnostics().changed(), before != Size::Default);
    assert_eq!(package_bytes(package), source);
    if !patch.is_noop() {
        assert_eq!(
            commit
                .package()
                .slide_table_dimension_size(slide, table, dimension)
                .unwrap_or_else(|error| panic!("reset candidate readback failed: {error}")),
            Size::Default,
        );
    }
}

fn changed_size(before: Size, data: &[u8], dimension: Dimension) -> Size {
    const VALUES: [f32; 6] = [8.0, 16.0, 32.0, 64.0, 98.0, 124.0];
    let start = usize::from(data.first().copied().unwrap_or_default())
        .wrapping_add(dimension.index())
        % VALUES.len();
    for offset in 0..VALUES.len() {
        let value = VALUES[(start + offset) % VALUES.len()];
        let candidate = Size::Points(
            Points::new(value)
                .unwrap_or_else(|error| panic!("finite positive points are valid: {error}")),
        );
        if candidate != before {
            return candidate;
        }
    }
    Size::Default
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    for slide in [
        SlideSelector::index(usize::MAX),
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        SlideSelector::name(""),
    ] {
        observe_result(package.slide_table_dimension_size(
            slide,
            TableSelector::index(0),
            Dimension::Row(0),
        ));
        match package.edit_slide_table_dimension_size(
            slide,
            TableSelector::index(0),
            Dimension::Column(0),
        ) {
            Ok(edit) => {
                black_box(edit.path());
            },
            Err(error) => observe_error(error),
        }
        assert_eq!(package_bytes(package), source);
    }
    observe_result(package.slide_table_dimension_size(
        SlideSelector::index(0),
        TableSelector::index(usize::MAX),
        Dimension::Row(0),
    ));
    assert_eq!(package_bytes(package), source);
}

fn exercise_redacted_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, fuzz_options()) {
        let rendered = error.to_string();
        assert!(!rendered.contains(std::str::from_utf8(PRIVATE_INPUT).unwrap_or_default()));
        observe_error(error);
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0_u8; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Ok(_) => panic!("oversized Keynote dimension input must be rejected"),
        Err(error) => observe_error(error),
    }
}

fn exercise_archive_limit() {
    let archive = Limits::new(
        MAX_INPUT_BYTES,
        1,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid Keynote dimension entry limits: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(_) => panic!("a one-entry Keynote dimension limit accepted the package"),
        Err(error) => observe_error(error),
    }
}

fn exercise_semantic_limit(data: &[u8]) {
    let max_objects = if data.first().copied().unwrap_or_default() & 1 == 0 {
        1
    } else {
        MAX_OBJECTS
    };
    let semantic = SemanticLimits::new(
        max_objects,
        1,
        1,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    );
    let Ok(semantic) = semantic else {
        return;
    };
    let options = ReadOptions::new(fuzz_options().archive(), semantic);
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(package) => {
            observe_result(package.slide_table_dimension_size(
                SlideSelector::index(0),
                TableSelector::index(0),
                Dimension::Row(0),
            ));
        },
        Err(error) => observe_error(error),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote dimension package must succeed: {error}"));
    bytes
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from(data.get(offset).copied().unwrap_or_default())
        | (u16::from(data.get(offset + 1).copied().unwrap_or_default()) << 8)
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

fn observe_error<E>(error: E)
where
    E: Debug + Display,
{
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
