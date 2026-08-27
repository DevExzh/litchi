#![no_main]

//! Bounded Numbers table-appearance lifecycle fuzzing.
//!
//! Arbitrary bytes exercise physical ingress, while the same bytes are also
//! interpreted as a small command against the checked-in native Numbers
//! table.  Keeping the package seed fixed makes the selector-first
//! read/edit/apply/inverse path reachable despite ZIP CRCs rejecting most
//! arbitrary mutations.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    table::appearance::transaction::{Error as AppearanceError, Path as AppearancePath},
    table::appearance::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing},
};

const MAX_INPUT_BYTES: u64 = 512 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 512 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_OBJECTS: usize = 4 * 1024;
const MAX_SHEETS: usize = 128;
const MAX_TABLES: usize = 512;
const MAX_REFERENCES: usize = 8 * 1024;
const MAX_MATERIALIZED_CELLS: usize = 64 * 1024;
const MAX_TEXT_BYTES: usize = 512 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const PRIVATE_SELECTOR: &str = "__litchi_private_numbers_appearance_selector_80a1__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_numbers_appearance_input_80a1__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    let command = command_input(data);
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected native bytes keep arbitrary package mutations shallow.
    // Always run the same bounded command sequence against the valid seed.
    exercise_package(native_package(), &command);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
});

/// Decode checked-in command recipes without changing the package-ingress
/// input.  A recipe is a command stream, never a package fixture; arbitrary
/// bytes still take the normal bounded ingress path above.
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

fn fuzz_options() -> PackageReadOptions {
    static OPTIONS: OnceLock<PackageReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = PackageLimits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Numbers fuzz archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Numbers fuzz semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Numbers fuzz projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options()).unwrap_or_else(|error| {
            panic!("native Numbers appearance fuzz seed must open: {error}")
        })
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let sheet = SheetSelector::index(usize::from(read_u16(data, 0)));
    let table = TableSelector::index(usize::from(read_u16(data, 2)));
    observe_result(package.table_appearance(sheet, table));
    observe_result(package.table_appearance(
        SheetSelector::name(PRIVATE_SELECTOR),
        TableSelector::index(0),
    ));

    let before = match package.table_appearance(SheetSelector::index(0), TableSelector::index(0)) {
        Ok(appearance) => appearance,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let after = command_appearance(before, data);
    let edit = match package.edit_table_appearance(SheetSelector::index(0), TableSelector::index(0))
    {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(edit.path(), AppearancePath::Table { sheet: 0, table: 0 });
    assert_eq!(edit.appearance(), before);
    let edit = edit.set(after);
    assert_eq!(edit.appearance(), after);

    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.path(), AppearancePath::Table { sheet: 0, table: 0 });
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_eq!(
        commit
            .package()
            .table_appearance(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("committed appearance must be readable: {error}")),
        after
    );
    black_box((&patch, diagnostics));

    let source_bytes = package_bytes(package);
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == target_bytes);

    let applied = package
        .apply_table_appearance(&patch)
        .unwrap_or_else(|error| panic!("fresh appearance patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .table_appearance(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("applied appearance must be readable: {error}")),
        after
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        assert!(matches!(
            applied.package().apply_table_appearance(&patch),
            Err(AppearanceError::PatchConflict)
        ));
        assert!(matches!(
            package.apply_table_appearance(&inverse),
            Err(AppearanceError::PatchConflict)
        ));
    }
    let restored = applied
        .package()
        .apply_table_appearance(&inverse)
        .unwrap_or_else(|error| panic!("fresh appearance inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .table_appearance(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("restored appearance must be readable: {error}")),
        before
    );
    assert_eq!(package_bytes(restored.package()), source_bytes);
}

fn command_appearance(before: Appearance, data: &[u8]) -> Appearance {
    // Command 0 is an exact no-op; command 1 sets every field from the next
    // byte; command 2 resets to native defaults; command 3 combines a
    // deterministic toggle with the current value.  The value is always
    // archive-free and finite, so malformed wire is left to package ingress.
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

fn exercise_redacted_malformed_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        observe_redacted(error, PRIVATE_MALFORMED_INPUT);
    }
}

fn exercise_input_limit() {
    let oversized = vec![0u8; OVERSIZED_INPUT_BYTES];
    if let Err(error) = Package::from_bytes_with_options(&oversized, fuzz_options()) {
        observe_error(error);
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("Numbers appearance package write failed: {error}"));
    bytes
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([
        data.get(offset).copied().unwrap_or_default(),
        data.get(offset + 1).copied().unwrap_or_default(),
    ])
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

fn observe_redacted<E>(error: E, private: &[u8])
where
    E: Debug + Display,
{
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private fuzz sentinel is UTF-8: {error}"));
    let rendered = error.to_string();
    assert!(
        !rendered.contains(private),
        "error leaked private selector/input"
    );
    observe_error(error);
}
