#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::table::lock::State as LockState;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector, TableSelector,
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
const MAX_FORMULA_WORK: usize = 8 * 1024;
const MAX_FORMULA_DEPTH: usize = 32;
const CONTROL_BYTES: usize = 8;
const MAX_SELECTOR_BYTES: usize = 512;
const PRIVATE_SHEET: &str = "__litchi_private_sheet_selector_2af1__";
const PRIVATE_TABLE: &str = "__litchi_private_table_selector_2af1__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    if let Ok(package) = Package::from_bytes_with_options(data, fuzz_options()) {
        exercise_package(&package, data);
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the focused
    // codec. Reuse the bytes as bounded semantic commands against a genuine
    // repository-owned package so every input reaches the transaction API.
    exercise_package(native_package(), data);
});

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
        .unwrap_or_else(|error| unreachable!("valid fuzz archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid fuzz semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid fuzz projection limits: {error}"))
                .with_formula_render_limits(MAX_FORMULA_WORK, MAX_FORMULA_DEPTH)
                .unwrap_or_else(|error| unreachable!("valid fuzz formula limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options())
            .unwrap_or_else(|error| panic!("native Numbers fuzz seed must open: {error}"));
        package
            .table_lock(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("native Numbers fuzz seed must expose a lock: {error}"));
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let sheet_index = usize::from(read_u16(data, 2));
    let table_index = usize::from(read_u16(data, 4));
    let name = selector_name(data);

    observe_result(package.table_lock(SheetSelector::index(0), TableSelector::index(0)));
    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        observe_result(package.table_lock(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
        ));
    }
    observe_result(package.table_lock(
        SheetSelector::index(sheet_index),
        TableSelector::index(table_index),
    ));
    observe_result(package.table_lock(
        SheetSelector::name(name.as_ref()),
        TableSelector::name(name.as_ref()),
    ));
    exercise_redacted_selector_errors(package);
    exercise_transaction(package, data);
}

fn exercise_transaction(package: &Package, data: &[u8]) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let Ok(before) = package.table_lock(sheet, table) else {
        return;
    };
    let after = if control(data, 0) & 1 == 0 {
        before
    } else {
        LockState::from_locked(!before.is_locked())
    };
    let Ok(mut edit) = package.edit_table_lock(sheet, table) else {
        return;
    };
    edit.set_state(after);
    assert_eq!(edit.state(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(
        diagnostics.touched_components(),
        usize::from(before != after)
    );
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    assert_eq!(
        commit
            .package()
            .table_lock(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("committed table lock must be readable: {error}")),
        after,
    );
    black_box(&patch);

    let applied = package
        .apply_table_lock(&patch)
        .unwrap_or_else(|error| panic!("fresh table-lock patch must apply: {error}"));
    assert_eq!(
        applied
            .package()
            .table_lock(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("applied table lock must be readable: {error}")),
        after,
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if before != after {
        match applied.package().apply_table_lock(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed table-lock patch must conflict with its target"),
        }
        match package.apply_table_lock(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed table-lock inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_table_lock(&inverse)
        .unwrap_or_else(|error| panic!("fresh table-lock inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .table_lock(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("restored table lock must be readable: {error}")),
        before,
    );
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
}

fn exercise_redacted_selector_errors(package: &Package) {
    if let Err(error) =
        package.table_lock(SheetSelector::name(PRIVATE_SHEET), TableSelector::index(0))
    {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) =
        package.table_lock(SheetSelector::index(0), TableSelector::name(PRIVATE_TABLE))
    {
        observe_redacted(error, PRIVATE_TABLE);
    }
}

fn selector_name(data: &[u8]) -> Cow<'_, str> {
    let start = data.len().min(CONTROL_BYTES);
    let end = data.len().min(start.saturating_add(MAX_SELECTOR_BYTES));
    String::from_utf8_lossy(&data[start..end])
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
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
