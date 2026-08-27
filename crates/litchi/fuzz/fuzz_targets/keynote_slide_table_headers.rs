#![no_main]

//! Selector-first fuzzing for Keynote slide-table header transactions.
//!
//! Arbitrary bytes exercise bounded package ingress. The same bytes are also
//! interpreted as a small command prefix against tiny source-built package
//! corpora, so CRC-protected ZIP mutations do not starve the deep
//! read/commit/apply paths. No private native fixture is used.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadError, ReadOptions, SemanticLimits, SlideSelector,
    slide::table::{
        TableSelector,
        headers::{Count, Settings},
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
const PRIVATE_SELECTOR: &str = "__litchi_private_slide_table_header_selector_b4e1__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => observe_package(&package),
        Err(error) => observe_error(error),
    }

    // Use the arbitrary input as a command stream against source-built
    // packages. This keeps successful selector and transaction coverage
    // available even when physical ZIP mutation is rejected, and the locked
    // variant exercises changed-edit atomicity without a private fixture.
    exercise_package(source_built_package(), data);
    exercise_package(locked_package(), data);
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

fn source_built_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(source_built_bytes(), fuzz_options()).unwrap_or_else(
            |error| panic!("source-built Keynote header package must open: {error}"),
        )
    })
}

fn locked_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(locked_bytes(), fuzz_options()).unwrap_or_else(|error| {
            panic!("source-built locked Keynote header package must open: {error}")
        })
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let (slide_index, table_index) = selectors(data, package);
    let selector = SlideSelector::index(slide_index);
    let table = TableSelector::index(table_index);

    // Read failures from out-of-range and name selectors are intentionally
    // observed as data, and must never mutate the immutable source snapshot.
    observe_result(package.slide_table_header_settings(selector, table));
    if let Err(error) = package.slide_table_header_settings(
        SlideSelector::name(PRIVATE_SELECTOR),
        TableSelector::index(0),
    ) {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source);

    // The source-built corpus has a stable table on the "Tables" slide.
    // Resolve by both semantic position and name when available, without
    // exposing any native object identifier.
    let (slide, table) = if package
        .slide_table_header_settings(SlideSelector::index(0), TableSelector::index(0))
        .is_ok()
    {
        (SlideSelector::index(0), TableSelector::index(0))
    } else {
        (SlideSelector::name("Tables"), TableSelector::index(0))
    };
    let before = match package.slide_table_header_settings(slide, table) {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_settings(before);
    match package.slide_table_header_settings(SlideSelector::name("Tables"), table) {
        Ok(named) => assert_eq!(named, before),
        Err(error) => observe_error(error),
    }
    let edit_slide = if data.get(7).copied().unwrap_or_default() & 1 == 0 {
        slide
    } else {
        SlideSelector::name("Tables")
    };

    exercise_noop(package, edit_slide, table, before, &source);
    exercise_invalid_settings(package, edit_slide, table, before, &source);
    exercise_changed_transaction(package, edit_slide, table, before, data, &source);
    exercise_selector_failures(package, &source);
}

fn exercise_noop(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Settings,
    source: &[u8],
) {
    let edit = match package.edit_slide_table_headers(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(edit.before(), before);
    assert_eq!(edit.settings(), before);
    let commit = match edit.set(before).commit() {
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
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(commit.package()), source);
    let applied = package
        .apply_slide_table_headers(patch)
        .unwrap_or_else(|error| panic!("fresh no-op header patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), source);
}

fn exercise_invalid_settings(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Settings,
    source: &[u8],
) {
    // The count type itself is checked at construction. A complete settings
    // value with five header/footer rows and columns is a deterministic
    // semantic rejection for the source-built 8x4 table.
    let invalid = Settings {
        header_rows: Some(Count::FIVE),
        header_columns: Some(Count::FIVE),
        footer_rows: Some(Count::FIVE),
        header_rows_frozen: before.header_rows_frozen,
        header_columns_frozen: before.header_columns_frozen,
        repeating_header_rows_enabled: before.repeating_header_rows_enabled,
        repeating_header_columns_enabled: before.repeating_header_columns_enabled,
    };
    let edit = match package.edit_slide_table_headers(slide, table) {
        Ok(edit) => edit.set(invalid),
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    match edit.commit() {
        Ok(commit) => {
            // A future larger source-built corpus could legitimately admit
            // the value; still require coherent readback and source
            // immutability.
            assert_eq!(commit.patch().after(), invalid);
            assert_eq!(package_bytes(package), source);
            black_box(commit.diagnostics());
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
        },
    }
}

fn exercise_changed_transaction(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Settings,
    data: &[u8],
    source: &[u8],
) {
    let after = changed_settings(before, data);
    let edit = match package.edit_slide_table_headers(slide, table) {
        Ok(edit) => edit.set(after),
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(edit.before(), before);
    assert_eq!(edit.after(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let patch = commit.patch().clone();
    let target = package_bytes(commit.package());
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    assert_eq!(
        diagnostics.touched_components(),
        if before != after { 1 } else { 0 }
    );
    assert_eq!(
        commit
            .package()
            .slide_table_header_settings(slide, table)
            .unwrap_or_else(|error| panic!("header candidate readback failed: {error}")),
        after,
    );
    assert_eq!(package_bytes(package), source);

    let applied = package
        .apply_slide_table_headers(&patch)
        .unwrap_or_else(|error| panic!("fresh header patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target);
    if !patch.is_noop() {
        match commit.package().apply_slide_table_headers(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("changed header patch unexpectedly applied twice"),
        }
        match package.apply_slide_table_headers(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("header inverse unexpectedly applied to source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_headers(&inverse)
        .unwrap_or_else(|error| panic!("fresh header inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .slide_table_header_settings(slide, table)
            .unwrap_or_else(|error| panic!("header inverse readback failed: {error}")),
        before,
    );
    assert_eq!(package_bytes(restored.package()), source);
    black_box((patch.source_fingerprint(), patch.target_fingerprint()));
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    for selector in [
        SlideSelector::index(usize::MAX),
        SlideSelector::name(PRIVATE_SELECTOR),
    ] {
        let result = package.edit_slide_table_headers(selector, TableSelector::index(0));
        if let Err(error) = result {
            observe_error(error);
        }
        assert_eq!(package_bytes(package), source);
    }
}

fn changed_settings(before: Settings, data: &[u8]) -> Settings {
    Settings {
        // Keep generated count values within the source-built table's 8x4
        // bounds while still exploring field removal and every supported
        // non-zero count.
        header_rows: optional_count(data.first().copied().unwrap_or_default(), 4),
        header_columns: optional_count(data.get(1).copied().unwrap_or_default(), 4),
        footer_rows: optional_count(data.get(2).copied().unwrap_or_default(), 4),
        header_rows_frozen: optional_bool(
            data.get(3).copied().unwrap_or_default(),
            before.header_rows_frozen,
        ),
        header_columns_frozen: optional_bool(
            data.get(4).copied().unwrap_or_default(),
            before.header_columns_frozen,
        ),
        repeating_header_rows_enabled: optional_bool(
            data.get(5).copied().unwrap_or_default(),
            before.repeating_header_rows_enabled,
        ),
        repeating_header_columns_enabled: optional_bool(
            data.get(6).copied().unwrap_or_default(),
            before.repeating_header_columns_enabled,
        ),
    }
}

fn optional_count(byte: u8, maximum: u8) -> Option<Count> {
    match byte % (maximum + 1) {
        0 => None,
        1 => Some(Count::ONE),
        2 => Some(Count::TWO),
        3 => Some(Count::THREE),
        4 => Some(Count::FOUR),
        _ => unreachable!("bounded header count selector"),
    }
}

fn optional_bool(byte: u8, before: Option<bool>) -> Option<bool> {
    if byte & 2 != 0 {
        None
    } else if byte & 1 == 0 {
        Some(!before.unwrap_or(false))
    } else {
        before
    }
}

fn selectors(data: &[u8], package: &Package) -> (usize, usize) {
    let slides = package.show().map(|show| show.slides().len()).unwrap_or(1);
    let slide = usize::from(read_u16(data, 0)) % slides.max(1);
    let table = usize::from(read_u16(data, 2));
    (slide, table)
}

fn observe_settings(settings: Settings) {
    black_box((
        settings.header_rows,
        settings.header_columns,
        settings.footer_rows,
        settings.header_rows_frozen,
        settings.header_columns_frozen,
        settings.repeating_header_rows_enabled,
        settings.repeating_header_columns_enabled,
        settings.header_row_count(),
        settings.header_column_count(),
        settings.footer_row_count(),
        settings.header_rows_are_frozen(),
        settings.header_columns_are_frozen(),
        settings.repeats_header_rows(),
        settings.repeats_header_columns(),
    ));
}

fn observe_package(package: &Package) {
    let _ = black_box(package.show().map(|show| show.slides().len()));
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
        .unwrap_or_else(|| panic!("source-built Keynote header corpus must start with hex:"));
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
            _ => panic!("source-built Keynote header corpus contains non-hex data"),
        };
        if let Some(high) = high.take() {
            output.push((high << 4) | nibble);
        } else {
            high = Some(nibble);
        }
    }
    assert!(
        high.is_none(),
        "source-built Keynote header corpus has odd hex length"
    );
    output
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Keynote header input must be rejected"),
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
        Ok(_) => panic!("a one-entry Keynote archive ceiling accepted the source-built package"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote header package must succeed: {error}"));
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

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
