#![no_main]

//! Bounded selector-first fuzzing for persisted Numbers table sort metadata.
//!
//! This target exercises only the table's stored sort configuration.  It does
//! not invoke physical row sorting: that operation owns tiles, headers,
//! formula dependencies, comments, and view state and remains a separate
//! migration-host path.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    table::sort::{ColumnIndex, Direction, Order, Rule, Scope},
};

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
const PRIVATE_SHEET: &str = "__litchi_private_sort_sheet_83c4__";
const PRIVATE_TABLE: &str = "__litchi_private_sort_table_83c4__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_sort_input_83c4__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    if let Ok(package) = Package::from_bytes_with_options(data, options()) {
        exercise_package(&package, data);
    } else {
        exercise_input_error(data);
    }
    // Arbitrary ZIP mutations almost always fail CRC/structure checks before
    // reaching the focused owner.  Replay each bounded command against the
    // repository-owned native workbook so every fuzz input reaches read,
    // set/clear, prepared publication, conflict, and inverse paths.
    exercise_package(native_package(), data);
    exercise_redacted_ingress();
    exercise_input_limit();
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
        .unwrap_or_else(|error| unreachable!("valid sort archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid sort semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid sort projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers sort seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    observe_result(package.table_sort_order(sheet, table));
    if let Err(error) =
        package.table_sort_order(SheetSelector::name(PRIVATE_SHEET), TableSelector::index(0))
    {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) =
        package.table_sort_order(SheetSelector::index(0), TableSelector::name(PRIVATE_TABLE))
    {
        observe_redacted(error, PRIVATE_TABLE);
    }

    let source_bytes = package_bytes(package);
    let before = match package.table_sort_order(sheet, table) {
        Ok(order) => order,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let desired = order_from_bytes(data);
    let edit = match package.edit_table_sort_order(sheet, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let command = data.first().copied().unwrap_or_default();
    let edit = if command & 3 == 0 {
        if command & 4 == 0 {
            edit.clear()
        } else {
            edit.reset()
        }
    } else {
        edit.set(desired.clone())
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            // Rejected writes must not publish a partial native candidate.
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    let expected_after = if command & 3 == 0 {
        None
    } else {
        Some(desired.clone())
    };
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), expected_after.as_ref());
    assert_eq!(patch.is_noop(), source_bytes == target_bytes);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
    }
    assert_eq!(
        commit
            .package()
            .table_sort_order(sheet, table)
            .unwrap_or_else(|error| panic!("sort candidate readback failed: {error}")),
        expected_after,
    );

    let applied = package
        .apply_table_sort_order(&patch)
        .unwrap_or_else(|error| panic!("fresh sort patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(applied.package().apply_table_sort_order(&patch).is_err());
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_sort_order(&inverse)
        .unwrap_or_else(|error| panic!("fresh sort inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_sort_order(sheet, table)
            .unwrap_or_else(|error| panic!("sort inverse readback failed: {error}")),
        before,
    );
}

fn order_from_bytes(data: &[u8]) -> Order {
    let scope = if data.get(1).copied().unwrap_or_default() & 1 == 0 {
        Scope::EntireTable
    } else {
        Scope::SelectedRows
    };
    let count = usize::from(data.first().copied().unwrap_or_default() % 3) + 1;
    let mut rules = Vec::with_capacity(count);
    let mut used = [false; 4];
    for priority in 0..count {
        let raw_column = usize::from(data.get(priority + 2).copied().unwrap_or(priority as u8) % 4);
        let mut selected = raw_column;
        while used[selected] {
            selected = (selected + 1) % used.len();
        }
        used[selected] = true;
        let column = ColumnIndex::new(selected).unwrap_or_else(|error| {
            unreachable!("generated sort column is native-compatible: {error}")
        });
        let direction = if data.get(priority + count + 2).copied().unwrap_or_default() & 1 == 0 {
            Direction::Ascending
        } else {
            Direction::Descending
        };
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules)
        .unwrap_or_else(|error| unreachable!("generated sort order is valid: {error}"))
}

fn exercise_input_error(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        black_box(data.len());
    }
}

fn exercise_redacted_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, options()) {
        observe_redacted(
            error,
            std::str::from_utf8(PRIVATE_INPUT).unwrap_or("sort-input"),
        );
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, options()) {
        Err(PackageError::InputTooLarge { observed, maximum }) => {
            assert_eq!(observed, OVERSIZED_INPUT_BYTES as u64);
            assert_eq!(maximum, MAX_INPUT_BYTES);
            black_box((observed, maximum));
        },
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized sort input must be rejected"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing sort package must succeed: {error}"));
    bytes
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
