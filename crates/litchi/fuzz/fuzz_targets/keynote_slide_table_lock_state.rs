#![no_main]

//! Bounded selector-first Keynote slide-table lock lifecycle fuzzing.
//!
//! Arbitrary bytes exercise bounded Keynote package ingress.  The same input
//! is also interpreted as a command stream against the small source-built
//! table packages used by the neighboring header target, so successful
//! selector/read/edit/apply paths remain reachable when arbitrary ZIP bytes do
//! not survive physical validation.  The checked-in inputs are command
//! recipes, not native package members.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadError, ReadOptions, SemanticLimits, SlideSelector,
    slide::table::{TableSelector, lock::State},
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
const PRIVATE_SLIDE_NAME: &str = "__litchi_private_slide_lock_state_missing__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => observe_untrusted_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Keep the complete lifecycle reachable from every input, independent of
    // whether that input happens to be a valid CRC-protected ZIP package.
    for package in source_packages() {
        exercise_package(package, data);
    }
    exercise_input_limit();
    exercise_archive_limit();
    exercise_semantic_limit(data);
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
        .unwrap_or_else(|error| unreachable!("valid Keynote lock-state archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote lock-state semantic limits: {error}"));
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
                    panic!("source-built Keynote lock-state package must open: {error}")
                })
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    })
}

fn source_built_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| decode_hex(SOURCE_BUILT_PACKAGE).into_boxed_slice())
        .as_ref()
}

fn locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| decode_hex(LOCKED_PACKAGE).into_boxed_slice())
        .as_ref()
}

fn exercise_package(package: &Package, _data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);

    let before = match package.slide_table_lock_state(slide, table) {
        Ok(state) => state,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            exercise_selector_failures(package, &source);
            return;
        },
    };
    black_box(before);
    assert_eq!(package_bytes(package), source);

    // Name and position are two public routes to the same selected slide;
    // neither route exposes the table's native object identifier.
    if let Ok(named) =
        package.slide_table_lock_state(SlideSelector::name("Tables"), TableSelector::index(0))
    {
        assert_eq!(named, before);
    }
    if let Err(error) = package.slide_table_lock_state(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_error(error);
    }

    exercise_noop(package, slide, table, before, &source);
    exercise_changed(package, slide, table, before, &source);
    exercise_selector_failures(package, &source);
}

fn observe_untrusted_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = usize::from(read_u16(data, 0));
    let table = usize::from(read_u16(data, 2));
    observe_result(
        package.slide_table_lock_state(SlideSelector::index(slide), TableSelector::index(table)),
    );
    assert_eq!(package_bytes(package), source);
}

fn exercise_noop(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: State,
    source: &[u8],
) {
    let mut edit = match package.edit_slide_table_lock_state(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.before(), before);
    if before == State::Locked {
        edit.lock();
    } else {
        edit.unlock();
    }
    assert_eq!(edit.state(), before);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch();
    assert!(patch.is_noop());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), before);
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(commit.package()), source);
    let applied = package
        .apply_slide_table_lock_state(patch)
        .unwrap_or_else(|error| panic!("fresh no-op lock patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), source);
}

fn exercise_changed(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: State,
    source: &[u8],
) {
    let mut edit = match package.edit_slide_table_lock_state(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    if before == State::Locked {
        edit.unlock();
    } else {
        edit.lock();
    }
    let after = edit.state();
    assert_ne!(after, before);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    let candidate = package_bytes(commit.package());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert!(!patch.is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit
            .package()
            .slide_table_lock_state(slide, table)
            .unwrap_or_else(|error| panic!("changed lock candidate readback failed: {error}")),
        after,
    );
    assert_eq!(package_bytes(package), source);

    let applied = package
        .apply_slide_table_lock_state(&patch)
        .unwrap_or_else(|error| panic!("fresh lock patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), candidate);
    match commit.package().apply_slide_table_lock_state(&patch) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("changed lock patch unexpectedly applied twice"),
    }
    match package.apply_slide_table_lock_state(&patch.inverse()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("lock inverse unexpectedly applied to source"),
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_lock_state(&inverse)
        .unwrap_or_else(|error| panic!("fresh lock inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .slide_table_lock_state(slide, table)
            .unwrap_or_else(|error| panic!("lock inverse readback failed: {error}")),
        before,
    );
    assert_eq!(package_bytes(restored.package()), source);
    black_box((patch.source_fingerprint(), patch.target_fingerprint()));
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    for selector in [
        SlideSelector::index(usize::MAX),
        SlideSelector::name(PRIVATE_SLIDE_NAME),
    ] {
        observe_result(package.slide_table_lock_state(selector, TableSelector::index(0)));
        if let Err(error) = package.edit_slide_table_lock_state(selector, TableSelector::index(0)) {
            observe_error(error);
        }
    }
    observe_result(
        package.slide_table_lock_state(SlideSelector::index(0), TableSelector::index(usize::MAX)),
    );
    assert_eq!(package_bytes(package), source);
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Keynote lock-state input must be rejected"),
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
    .unwrap_or_else(|error| panic!("valid Keynote entry-limit profile: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("one-entry Keynote archive ceiling accepted the source-built package"),
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
        MAX_REFERENCES,
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
            let _ = black_box(package.show().map(|show| show.slides().len()));
        },
        Err(error) => observe_error(error),
    }
}

fn decode_hex(source: &[u8]) -> Vec<u8> {
    let source = source
        .strip_prefix(b"hex:")
        .unwrap_or_else(|| panic!("source-built Keynote lock-state corpus must start with hex:"));
    let mut output = Vec::with_capacity(source.len() / 2);
    let mut high = None;
    for byte in source.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = match byte {
            b'0'..=b'9' => byte - b'0',
            b'a'..=b'f' => byte - b'a' + 10,
            b'A'..=b'F' => byte - b'A' + 10,
            _ => panic!("source-built Keynote lock-state corpus contains non-hex data"),
        };
        if let Some(high) = high.take() {
            output.push((high << 4) | nibble);
        } else {
            high = Some(nibble);
        }
    }
    assert!(
        high.is_none(),
        "source-built Keynote lock-state corpus has odd hex length"
    );
    output
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote lock-state package must succeed: {error}"));
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
            let _ = black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
