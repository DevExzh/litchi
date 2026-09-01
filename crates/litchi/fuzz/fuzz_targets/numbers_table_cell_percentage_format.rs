#![no_main]

//! Bounded selector-first Numbers table-cell Percentage-format fuzzing.
//!
//! The checked `basic.numbers` source is deliberately used as a safe refusal
//! oracle when its stable B3 cell is not a Percentage.  If a source does
//! expose a Percentage at that coordinate, the same command stream exercises
//! the complete public lifecycle: no-op, set, clear/reset, exact patch
//! application, conflict/inverse handling, candidate readback, diagnostics,
//! and source-byte atomicity.  No native IDs or generated messages cross this
//! target's public `litchi::numbers` boundary.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::FixedDecimalPlaces,
    cell::data_format::percentage::{
        DecimalPlaces, NegativeStyle, Percentage, ThousandsSeparator,
        transaction::Error as PercentageError,
    },
};

#[path = "../../../litchi-numbers/tests/support/table_cell_data_format_fixture.rs"]
mod percentage_fixture;

const MAX_INPUT_BYTES: u64 = 512 * 1024;
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
const OVERSIZED_INPUT_BYTES: usize = 512 * 1024 + 1;
const PRIVATE_SHEET: &str = "__litchi_private_percentage_format_sheet_7e29__";
const PRIVATE_TABLE: &str = "__litchi_private_percentage_format_table_7e29__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_percentage_format_input_7e29__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

// B3 is a stable numeric scalar in the repository's native Numbers fixture.
// It is retained as the wrong-family refusal oracle; the deterministic
// test-only fixture below supplies a real Percentage cell for the success
// lifecycle.
const NATIVE_POSITION: CellPosition = CellPosition::new(2, 1);
const PERCENTAGE_POSITION: CellPosition = CellPosition::new(0, 0);

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data, PERCENTAGE_POSITION),
        Err(error) => observe_error(error),
    }

    // Replay every command against a valid source so semantic reads, edits,
    // exact patch conflicts, inverses, and candidate verification are not
    // starved by physical package admission.  Keep the native Number source
    // as a separate, strict wrong-family refusal oracle.
    exercise_wrong_family(native_package());
    exercise_package(percentage_package(), data, PERCENTAGE_POSITION);
    exercise_constructor_edges(data);
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_semantic_limit();
});

fn options() -> PackageReadOptions {
    static OPTIONS: OnceLock<PackageReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = PackageLimits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid percentage-format archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| {
                    unreachable!("valid percentage-format semantic limits: {error}")
                })
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid percentage-format projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options()).unwrap_or_else(|error| {
            panic!("native Numbers Percentage-format fuzz seed must open: {error}")
        })
    })
}

fn percentage_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = percentage_fixture::synthetic_package_for(
            percentage_fixture::FormatFamily::Percentage,
            percentage_fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Percentage fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("deterministic Percentage source must open: {error}"))
    })
}

fn exercise_wrong_family(package: &Package) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        NATIVE_POSITION,
    );
    assert!(
        matches!(read, Err(PercentageError::WrongFormatFamily { .. })),
        "native Number cell must be refused by the Percentage API: {read:?}"
    );
    let edit = package.edit_table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        NATIVE_POSITION,
    );
    assert!(
        matches!(edit, Err(PercentageError::WrongFormatFamily { .. })),
        "native Number cell edit must be refused by the Percentage API: {edit:?}"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_package(package: &Package, data: &[u8], position: CellPosition) {
    exercise_selector_reads(package, data);
    exercise_selector_errors(package, position);

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    ) {
        Ok(format) => format,
        Err(error) => {
            // Unsupported or non-Percentage cells are still useful for
            // source-atomic refusal coverage.  The fixed native Number
            // oracle above separately asserts the exact error variant.
            observe_error(error);
            if let Err(error) = package.edit_table_cell_percentage_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            ) {
                observe_error(error);
            }
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let requested = percentage_from_bytes(data);
    exercise_transaction(package, &source_bytes, before, requested, position, data);
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8]) {
    let position = CellPosition::new(
        u32::from(control(data, 4) % 8),
        u32::from(control(data, 5) % 8),
    );
    observe_result(package.table_cell_percentage_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        position,
    ));

    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        observe_result(package.table_cell_percentage_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        ));
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let sheet_count = package.document().sheets().len();
    if let Err(error) = package.table_cell_percentage_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ) {
        observe_error(error);
    }
    if let Err(error) = package.table_cell_percentage_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(u32::MAX, u32::MAX),
    ) {
        observe_error(error);
    }
}

fn exercise_transaction(
    package: &Package,
    source_bytes: &[u8],
    before: Option<Percentage>,
    requested: Percentage,
    position: CellPosition,
    data: &[u8],
) {
    let edit = match package.edit_table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    ) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    // Four commands cover an exact semantic no-op, set, clear, and reset.
    // Preserve the source Optional state for the no-op command.
    let command = control(data, 0) & 3;
    let (edit, expected_after) = match command {
        0 => match before {
            Some(format) => (edit.set(format), before),
            None => (edit.clear(), None),
        },
        1 => (edit.set(requested), Some(requested)),
        2 => (edit.clear(), None),
        _ => (edit.reset(), None),
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
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == target_bytes);
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), expected_after.as_ref());
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert!(commit.diagnostics().full_reparse_performed());
    }
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    black_box(commit.diagnostics());
    assert_eq!(
        commit
            .package()
            .table_cell_percentage_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("Percentage-format candidate readback failed: {error}")),
        expected_after
    );

    let applied = package
        .apply_table_cell_percentage_format(&patch)
        .unwrap_or_else(|error| panic!("Percentage-format patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_percentage_format(&patch)
                .is_err(),
            "a committed Percentage-format patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_percentage_format(&patch.inverse())
                .is_err(),
            "a Percentage-format inverse must reject the original source"
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_percentage_format(&inverse)
        .unwrap_or_else(|error| panic!("Percentage-format inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_percentage_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("Percentage-format inverse readback failed: {error}")),
        before
    );
    black_box((patch.before(), patch.after(), patch.path()));
}

fn percentage_from_bytes(data: &[u8]) -> Percentage {
    let decimal_places = if control(data, 1) & 1 == 0 {
        DecimalPlaces::Automatic
    } else {
        DecimalPlaces::fixed(control(data, 2) % 31)
            .unwrap_or_else(|error| panic!("bounded decimal places must validate: {error}"))
    };
    let negative_style = match control(data, 3) & 3 {
        0 => NegativeStyle::MinusSign,
        1 => NegativeStyle::Red,
        2 => NegativeStyle::Parentheses,
        _ => NegativeStyle::RedParentheses,
    };
    let thousands_separator = if control(data, 4) & 1 == 0 {
        ThousandsSeparator::Hidden
    } else {
        ThousandsSeparator::Shown
    };
    Percentage::new(decimal_places, negative_style, thousands_separator)
}

fn exercise_constructor_edges(data: &[u8]) {
    assert!(FixedDecimalPlaces::new(31).is_err());
    assert!(DecimalPlaces::fixed(31).is_err());
    for value in [0, 1, 2, 30] {
        let places = FixedDecimalPlaces::new(value)
            .unwrap_or_else(|error| panic!("valid decimal places rejected: {error}"));
        black_box(places);
    }
    let format = percentage_from_bytes(data);
    black_box((
        format.decimal_places(),
        format.negative_style(),
        format.thousands_separator(),
    ));
}

fn exercise_redacted_ingress() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for input in [PRIVATE_INPUT, b"".as_slice(), b"not-a-numbers-package"] {
            if let Err(error) = Package::from_bytes_with_options(input, options()) {
                if input.is_empty() {
                    observe_error(error);
                } else {
                    let private = std::str::from_utf8(input).unwrap_or_else(|error| {
                        unreachable!("malformed sentinel is UTF-8: {error}")
                    });
                    observe_redacted(error, private);
                }
            }
        }
    });
}

fn exercise_input_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
        let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
        match Package::from_bytes_with_options(bytes, options()) {
            Err(PackageError::InputTooLarge { observed, maximum }) => {
                assert_eq!(observed, OVERSIZED_INPUT_BYTES as u64);
                assert_eq!(maximum, MAX_INPUT_BYTES);
                black_box((observed, maximum));
            },
            Err(PackageError::Archive(error)) => {
                assert_eq!(
                    error.to_string(),
                    format!(
                        "iWork archive input bytes limit exceeded: observed {OVERSIZED_INPUT_BYTES}, maximum {MAX_INPUT_BYTES}"
                    )
                );
                black_box(error);
            },
            Err(error) => observe_error(error),
            Ok(_) => panic!("oversized Numbers Percentage-format input must be rejected"),
        }
    });
}

fn exercise_semantic_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let semantic = PackageSemanticLimits::new(1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("one-entry semantic limits must validate: {error}"));
        let constrained = PackageReadOptions::new(options().archive(), semantic);
        let result = Package::from_bytes_with_options(NATIVE_NUMBERS, constrained);
        assert!(
            result.is_err(),
            "one-entry semantic limits unexpectedly admitted native source"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Percentage-format package failed: {error}"));
    bytes
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

fn observe_error<E>(error: E)
where
    E: Debug + Display,
{
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted<E>(error: E, private: &str)
where
    E: Debug + Display,
{
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(
        !display.contains(private),
        "Percentage-format error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Percentage-format debug leaked private input"
    );
    black_box((display, debug));
}
