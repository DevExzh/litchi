#![no_main]

//! Bounded selector-first Numbers Pop-Up Menu lifecycle fuzzing.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::data_format::pop_up_menu::{InitialSelection, Item, PopUpMenu},
    table::CellPosition,
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
const PRIVATE_SHEET: &str = "__litchi_private_popup_sheet_81b7__";
const PRIVATE_TABLE: &str = "__litchi_private_popup_table_81b7__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_popup_input_81b7__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    if let Ok(package) = Package::from_bytes_with_options(data, options()) {
        exercise_package(&package, data);
    } else {
        exercise_input_error(data);
    }
    exercise_package(native_package(), data);
    exercise_constructor_edges(data);
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
        .unwrap_or_else(|error| unreachable!("valid popup archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid popup semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid popup projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers popup seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let position = CellPosition::new(
        u32::from(read_u16(data, 0) % 8),
        u32::from(read_u16(data, 2) % 8),
    );
    observe_result(package.table_cell_pop_up_menu_format(sheet, table, position));
    if let Err(error) = package.table_cell_pop_up_menu_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_pop_up_menu_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_pop_up_menu_format(sheet, table, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let desired = menu(data);
    let edit = match package.edit_table_cell_pop_up_menu_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let commit = match edit.set(desired.clone()).commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), Some(&desired));
    assert_eq!(patch.is_noop(), source_bytes == target_bytes);
    assert_eq!(
        commit.diagnostics().changed(),
        before != Some(desired.clone())
    );
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert!(commit.diagnostics().deleted_previews() <= 3);
    }
    assert_eq!(
        commit
            .package()
            .table_cell_pop_up_menu_format(sheet, table, position)
            .unwrap_or_else(|error| panic!("popup candidate readback failed: {error}")),
        Some(desired.clone())
    );
    let applied = package
        .apply_table_cell_pop_up_menu_format(&patch)
        .unwrap_or_else(|error| panic!("popup patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_pop_up_menu_format(&patch)
                .is_err()
        );
    }
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_pop_up_menu_format(&inverse)
        .unwrap_or_else(|error| panic!("popup inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_pop_up_menu_format(sheet, table, position)
            .unwrap_or_else(|error| panic!("popup inverse readback failed: {error}")),
        before
    );

    // A second cell with the same menu should reuse the existing control
    // model when the source has a writable number cell at this coordinate.
    // Sparse or non-number cells are valid refusals, so this branch remains
    // observational and never turns an unsupported source into a failure.
    let other_position = CellPosition::new(position.row().saturating_add(1), position.column());
    if other_position != position
        && let Ok(second_edit) =
            commit
                .package()
                .edit_table_cell_pop_up_menu_format(sheet, table, other_position)
        && let Ok(second_commit) = second_edit.set(desired.clone()).commit()
    {
        assert_eq!(
            second_commit
                .package()
                .table_cell_pop_up_menu_format(sheet, table, other_position)
                .unwrap_or_else(|error| panic!("popup reuse readback failed: {error}")),
            Some(desired.clone())
        );
        let second_target = package_bytes(second_commit.package());
        let second_restored = second_commit
            .package()
            .apply_table_cell_pop_up_menu_format(&second_commit.patch().inverse())
            .unwrap_or_else(|error| panic!("popup reuse inverse must apply: {error}"));
        assert_eq!(package_bytes(second_restored.package()), target_bytes);
        black_box(second_target);
    }

    // Exercise reset and final cull when the source admitted the graph.
    if !patch.is_noop()
        && let Ok(edit) = commit
            .package()
            .edit_table_cell_pop_up_menu_format(sheet, table, position)
    {
        let reset = if data.get(3).copied().unwrap_or_default() & 1 == 0 {
            edit.clear()
        } else {
            edit.reset()
        };
        if let Ok(reset_commit) = reset.commit() {
            let reset_patch = reset_commit.patch().clone();
            assert_eq!(reset_patch.after(), None);
            assert_eq!(
                reset_commit
                    .package()
                    .table_cell_pop_up_menu_format(sheet, table, position)
                    .unwrap_or_else(|error| panic!("popup reset readback failed: {error}")),
                None
            );
            let reset_inverse = reset_commit
                .package()
                .apply_table_cell_pop_up_menu_format(&reset_patch.inverse())
                .unwrap_or_else(|error| panic!("popup reset inverse must apply: {error}"));
            assert_eq!(package_bytes(reset_inverse.package()), target_bytes);
        }
    }
}

fn menu(data: &[u8]) -> PopUpMenu {
    let count = usize::from(data.first().copied().unwrap_or_default() % 3) + 1;
    let items = (0..count)
        .map(|index| {
            let first = data.get(index + 1).copied().unwrap_or_default();
            let second = data.get(index + count + 1).copied().unwrap_or_default();
            format!("Choice {first:02x}-{second:02x}-{index}")
        })
        .collect::<Vec<_>>();
    PopUpMenu::new(items)
        .unwrap_or_else(|error| unreachable!("generated popup item is valid: {error}"))
        .with_initial_selection(if data.get(2).copied().unwrap_or_default() & 1 == 0 {
            InitialSelection::FirstItem
        } else {
            InitialSelection::Blank
        })
}

fn exercise_constructor_edges(data: &[u8]) {
    observe_result(PopUpMenu::new(Vec::<&str>::new()));
    observe_result(Item::new("invalid\ncontrol"));
    if data.first().copied().unwrap_or_default() & 1 != 0 {
        let oversized = "x".repeat(4 * 1024 + 1);
        observe_result(Item::new(&oversized));
    }
}

fn exercise_input_error(data: &[u8]) {
    if data.len() <= 512 * 1024 {
        return;
    }
    black_box(data.len());
}

fn exercise_redacted_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, options()) {
        observe_redacted(
            error,
            std::str::from_utf8(PRIVATE_INPUT).unwrap_or("popup-input"),
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
        Ok(_) => panic!("oversized popup input must be rejected"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("popup package write failed: {error}"));
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
    black_box(format!("{error:?}"));
}

fn observe_redacted<E>(error: E, private: &str)
where
    E: Debug + Display,
{
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
