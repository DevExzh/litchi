#![no_main]

//! Bounded selector-first fuzzing for Pages body-table appearance settings.
//!
//! The target exercises only the persisted appearance value.  It does not
//! reorder rows or rewrite cell, formula, tile, comment, or view-state data;
//! those operations remain compatibility-host concerns.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    BodyTableSelector, Limits, Package,
    table::appearance::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
const PRIVATE_TABLE: &str = "__litchi_private_pages_appearance_table_104__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_pages_appearance_input_104__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, &command_input(data)),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary mutations unlikely to reach the focused
    // table graph.  Reuse every bounded input as a command against a valid
    // repository-owned Pages source so the transaction paths stay reachable.
    exercise_package(native_package(), &command_input(data));
    exercise_semantic_values(data);
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_archive_limit();
});

fn fuzz_limits() -> Limits {
    static LIMITS: OnceLock<Limits> = OnceLock::new();
    *LIMITS.get_or_init(|| {
        Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Pages appearance limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages appearance seed must open: {error}"))
    })
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
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(16) {
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
            if output.len() > MAX_COMMAND_BYTES {
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
    let source_before = package_bytes(package);
    let selector = BodyTableSelector::index(0);

    observe_result(package.body_table_appearance(selector));
    if let Err(error) = package.body_table_appearance(BodyTableSelector::name(PRIVATE_TABLE)) {
        observe_redacted(error, PRIVATE_TABLE.as_bytes());
    }
    exercise_selector_failures(package, &source_before);

    let before = match package.body_table_appearance(selector) {
        Ok(appearance) => appearance,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    let after = appearance_from_bytes(before, data);

    let edit = match package.edit_body_table_appearance(selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
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
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    assert_source_unchanged(package, &source_before);
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
        assert_eq!(target_bytes, source_before);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_eq!(
        commit
            .package()
            .body_table_appearance(selector)
            .unwrap_or_else(|error| panic!("appearance candidate readback failed: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = match package.apply_body_table_appearance(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .body_table_appearance(selector)
            .unwrap_or_else(|error| panic!("applied appearance readback failed: {error}")),
        after,
    );
    if !patch.is_noop() {
        match applied.package().apply_body_table_appearance(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Pages appearance patch must conflict on its target"),
        }
        match package.apply_body_table_appearance(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Pages appearance inverse must conflict on its source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_body_table_appearance(&inverse)
        .unwrap_or_else(|error| panic!("Pages appearance inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
    assert_eq!(
        restored
            .package()
            .body_table_appearance(selector)
            .unwrap_or_else(|error| panic!("appearance inverse readback failed: {error}")),
        before,
    );
    assert_source_unchanged(package, &source_before);
}

fn appearance_from_bytes(before: Appearance, data: &[u8]) -> Appearance {
    match control(data, 0) & 3 {
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

fn exercise_semantic_values(data: &[u8]) {
    let generated = Appearance {
        row_banding: if bit(data, 5, 0) {
            Banding::Enabled
        } else {
            Banding::Disabled
        },
        row_sizing: if bit(data, 5, 1) {
            RowSizing::FitCellContents
        } else {
            RowSizing::Fixed
        },
        gridlines: Gridlines {
            body_horizontal: visibility(data, 6, 0),
            header_columns_horizontal: visibility(data, 6, 1),
            body_vertical: visibility(data, 6, 2),
            header_rows_vertical: visibility(data, 6, 3),
            footer_rows_vertical: visibility(data, 6, 4),
        },
    };
    black_box(generated);
    black_box(Appearance::default());
}

fn exercise_selector_failures(package: &Package, source_before: &[u8]) {
    for selector in [
        BodyTableSelector::index(usize::MAX),
        BodyTableSelector::name(PRIVATE_TABLE),
        BodyTableSelector::name(""),
    ] {
        observe_result(package.body_table_appearance(selector));
        assert_source_unchanged(package, source_before);
    }
}

fn exercise_redacted_ingress() {
    match Package::from_bytes_with_limits(PRIVATE_INPUT, fuzz_limits()) {
        Err(error) => observe_redacted(error, PRIVATE_INPUT),
        Ok(_) => panic!("a private malformed Pages appearance sentinel must not parse"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Pages appearance input must be rejected"),
    }
}

fn exercise_archive_limit() {
    let limits = Limits::new(
        MAX_INPUT_BYTES,
        1,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid Pages appearance entry limits: {error}"));
    match Package::from_bytes_with_limits(NATIVE_PAGES, limits) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("a one-entry Pages appearance limit accepted the package"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Pages appearance package failed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source_before: &[u8]) {
    assert_eq!(package_bytes(package), source_before);
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn bit(data: &[u8], offset: usize, bit_offset: u8) -> bool {
    control(data, offset) & (1 << bit_offset) != 0
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

fn observe_redacted(error: impl Debug + Display, private: &[u8]) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    let private = String::from_utf8_lossy(private);
    assert!(!display.contains(private.as_ref()));
    assert!(!debug.contains(private.as_ref()));
    black_box((display, debug));
}
