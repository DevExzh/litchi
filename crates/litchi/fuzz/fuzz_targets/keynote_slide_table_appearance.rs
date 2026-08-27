#![no_main]

//! Bounded selector-first Keynote slide-table appearance lifecycle fuzzing.
//!
//! Arbitrary bytes exercise bounded package ingress.  The same input is also
//! decoded as a small command stream and replayed against source-built table
//! packages, keeping no-op, set, apply, conflict, inverse, reopen, selector,
//! and source-atomic failure paths reachable when ZIP mutation is rejected.
//! The command corpus is not a native package fixture.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::table::{
        TableSelector,
        appearance::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing},
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
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_appearance_selector_103__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_appearance_input_103__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

fuzz_target!(|data: &[u8]| {
    // Keep package ingress on the original arbitrary bytes so a valid ZIP
    // supplied by a caller remains a package input, while `hex:` recipes are
    // interpreted only as commands for the fixed source packages below.
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, &command_input(data)),
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
        .unwrap_or_else(|error| unreachable!("valid Keynote appearance archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote appearance semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [source_built_bytes(), NATIVE_KEYNOTE]
            .into_iter()
            .map(|source| {
                Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(|error| {
                    panic!("source-built Keynote appearance package must open: {error}")
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
                .unwrap_or_else(|| {
                    panic!("source-built Keynote appearance package has invalid hex")
                })
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
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(16)
        && encoded
            != SOURCE_BUILT_PACKAGE
                .strip_prefix(b"hex:")
                .unwrap_or_default()
    {
        return None;
    }
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

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let (slide_index, table_index) = selectors(package, data);
    let slide = SlideSelector::index(slide_index);
    let table = TableSelector::index(table_index);

    observe_result(package.slide_table_appearance(slide, table));
    observe_result(package.slide_table_appearance(
        SlideSelector::name(PRIVATE_SELECTOR),
        TableSelector::index(0),
    ));
    assert_eq!(package_bytes(package), source);
    exercise_selector_failures(package, &source);

    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);
    let before = match package.slide_table_appearance(slide, table) {
        Ok(appearance) => appearance,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let after = appearance_from_bytes(before, data);

    let edit = match package.edit_slide_table_appearance(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    black_box(edit.path());
    assert_eq!(edit.appearance(), before);
    let edit = edit.set(after);
    assert_eq!(edit.appearance(), after);

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
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
        assert_eq!(target, source);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_eq!(package_bytes(package), source);
    assert_eq!(
        commit
            .package()
            .slide_table_appearance(slide, table)
            .unwrap_or_else(|error| panic!("appearance candidate readback failed: {error}")),
        after,
    );
    black_box((
        patch.path(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
    ));

    let applied = match package.apply_slide_table_appearance(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), target);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_slide_table_appearance(&patch)
                .is_err()
        );
        assert!(
            package
                .apply_slide_table_appearance(&patch.inverse())
                .is_err()
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_appearance(&inverse)
        .unwrap_or_else(|error| panic!("appearance inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_table_appearance(slide, table)
            .unwrap_or_else(|error| panic!("appearance inverse readback failed: {error}")),
        before,
    );
}

fn selectors(package: &Package, data: &[u8]) -> (usize, usize) {
    let slides = package.show().map(|show| show.slides().len()).unwrap_or(1);
    let slide = usize::from(read_u16(data, 0)) % slides.max(1);
    let table = usize::from(read_u16(data, 2));
    (slide, table)
}

fn appearance_from_bytes(before: Appearance, data: &[u8]) -> Appearance {
    match data.first().copied().unwrap_or_default() & 3 {
        0 => before,
        1 => Appearance {
            row_banding: if bit(data, 1, 0) {
                Banding::Enabled
            } else {
                Banding::Disabled
            },
            row_sizing: if bit(data, 1, 1) {
                RowSizing::FitCellContents
            } else {
                RowSizing::Fixed
            },
            gridlines: Gridlines {
                body_horizontal: visibility(data, 2, 0),
                header_columns_horizontal: visibility(data, 2, 1),
                body_vertical: visibility(data, 2, 2),
                header_rows_vertical: visibility(data, 2, 3),
                footer_rows_vertical: visibility(data, 2, 4),
            },
        },
        2 => Appearance::default(),
        _ => Appearance {
            row_banding: match before.row_banding {
                Banding::Disabled => Banding::Enabled,
                Banding::Enabled => Banding::Disabled,
            },
            row_sizing: match before.row_sizing {
                RowSizing::Fixed => RowSizing::FitCellContents,
                RowSizing::FitCellContents => RowSizing::Fixed,
            },
            gridlines: Gridlines {
                body_horizontal: invert(before.gridlines.body_horizontal),
                header_columns_horizontal: invert(before.gridlines.header_columns_horizontal),
                body_vertical: invert(before.gridlines.body_vertical),
                header_rows_vertical: invert(before.gridlines.header_rows_vertical),
                footer_rows_vertical: invert(before.gridlines.footer_rows_vertical),
            },
        },
    }
}

fn visibility(data: &[u8], offset: usize, bit_offset: u8) -> GridlineVisibility {
    if bit(data, offset, bit_offset) {
        GridlineVisibility::Visible
    } else {
        GridlineVisibility::Hidden
    }
}

fn invert(value: GridlineVisibility) -> GridlineVisibility {
    match value {
        GridlineVisibility::Hidden => GridlineVisibility::Visible,
        GridlineVisibility::Visible => GridlineVisibility::Hidden,
    }
}

fn bit(data: &[u8], offset: usize, bit_offset: u8) -> bool {
    data.get(offset).copied().unwrap_or_default() & (1 << bit_offset) != 0
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    for slide in [
        SlideSelector::index(usize::MAX),
        SlideSelector::name(PRIVATE_SELECTOR),
        SlideSelector::name(""),
    ] {
        let result = package.slide_table_appearance(slide, TableSelector::index(0));
        observe_result(result);
        assert_eq!(package_bytes(package), source);
    }
    let result =
        package.slide_table_appearance(SlideSelector::index(0), TableSelector::index(usize::MAX));
    observe_result(result);
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
    let oversized = vec![0_u8; OVERSIZED_INPUT_BYTES];
    match Package::from_bytes_with_options(&oversized, fuzz_options()) {
        Ok(_) => panic!("oversized Keynote appearance input must be rejected"),
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
    .unwrap_or_else(|error| unreachable!("valid Keynote appearance entry limits: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(_) => panic!("a one-entry Keynote appearance limit accepted the package"),
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

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote appearance package failed: {error}"));
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
    black_box(format_args!("{error:?}"));
}
