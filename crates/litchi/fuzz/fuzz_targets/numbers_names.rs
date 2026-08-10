#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
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
const MAX_FORMULA_WORK: usize = 8 * 1024;
const MAX_FORMULA_DEPTH: usize = 32;
const MAX_NAME_BYTES: usize = 256;
const PRIVATE_SELECTOR: &str = "__litchi_private_numbers_name_selector_6d1a__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_numbers_name_input_6d1a__";
const UNICODE_NAME: &str = "Líneas 你好 🧪";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the strict name
    // transaction. Reuse them as bounded commands against a native package.
    exercise_package(native_package(), data);
    exercise_unicode(native_package());
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
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
        .unwrap_or_else(|error| unreachable!("valid Numbers fuzz archive limits: {error}"));
        let semantic = PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
            .unwrap_or_else(|error| unreachable!("valid Numbers fuzz semantic limits: {error}"))
            .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
            .unwrap_or_else(|error| unreachable!("valid Numbers fuzz projection limits: {error}"))
            .with_formula_render_limits(MAX_FORMULA_WORK, MAX_FORMULA_DEPTH)
            .unwrap_or_else(|error| unreachable!("valid Numbers fuzz formula limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options())
            .unwrap_or_else(|error| panic!("native Numbers names fuzz seed must open: {error}"));
        let sheet = package
            .document()
            .sheets()
            .first()
            .unwrap_or_else(|| panic!("native Numbers names fuzz seed must have a sheet"));
        assert!(sheet.tables().next().is_some(), "native Numbers names fuzz seed must have a table");
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let Some(sheet) = package.document().sheets().first() else {
        return;
    };
    let Some(table) = sheet.tables().next() else {
        return;
    };
    let sheet_name = sheet.name().to_owned();
    let table_name = table.name().to_owned();
    observe_result(package.edit_names().rename_sheet(
        SheetSelector::index(usize::from(read_u16(data, 1))),
        replacement(data, "sheet").as_ref(),
    ));
    observe_result(package.edit_names().rename_table(
        SheetSelector::name(sheet_name.as_str()),
        TableSelector::name(table_name.as_str()),
        replacement(data, "table").as_ref(),
    ));
    exercise_redacted_selector_error(package);

    let command = control(data, 0) & 3;
    let sheet_after = replacement(data, "sheet");
    let table_after = replacement(data, "table");
    let edit = match command {
        0 => package.edit_names().rename_sheet(SheetSelector::index(0), &sheet_name),
        1 => package.edit_names().rename_sheet(SheetSelector::index(0), sheet_after.as_ref()),
        2 => package.edit_names().rename_table(
            SheetSelector::name(sheet_name.as_str()),
            TableSelector::name(table_name.as_str()),
            table_after.as_ref(),
        ),
        _ => package
            .edit_names()
            .rename_sheet(SheetSelector::index(0), sheet_after.as_ref())
            .and_then(|edit| {
                edit.rename_table(
                    SheetSelector::name(sheet_name.as_str()),
                    TableSelector::name(table_name.as_str()),
                    table_after.as_ref(),
                )
            }),
    };
    let Ok(edit) = edit else {
        return;
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(diagnostics.operations(), patch.operation_count());
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_names(commit.package(), command, &sheet_name, &table_name, &sheet_after, &table_after);
    black_box((&patch, diagnostics));

    let applied = package
        .apply_names(&patch)
        .unwrap_or_else(|error| panic!("fresh Numbers names patch must apply: {error}"));
    assert_names(applied.package(), command, &sheet_name, &table_name, &sheet_after, &table_after);
    assert_eq!(package_bytes(applied.package()), package_bytes(commit.package()));

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        assert!(applied.package().apply_names(&patch).is_err());
        assert!(package.apply_names(&inverse).is_err());
    }
    let restored = applied
        .package()
        .apply_names(&inverse)
        .unwrap_or_else(|error| panic!("fresh Numbers names inverse must apply: {error}"));
    assert_names(restored.package(), 0, &sheet_name, &table_name, &sheet_after, &table_after);
    assert_eq!(package_bytes(restored.package()), package_bytes(package));
}

fn assert_names(
    package: &Package,
    command: u8,
    sheet_before: &str,
    table_before: &str,
    sheet_after: &str,
    table_after: &str,
) {
    let sheet = package
        .document()
        .sheets()
        .first()
        .unwrap_or_else(|| panic!("names transaction must retain its first sheet"));
    let table = sheet
        .tables()
        .next()
        .unwrap_or_else(|| panic!("names transaction must retain its first table"));
    assert_eq!(sheet.name(), if matches!(command, 1 | 3) { sheet_after } else { sheet_before });
    assert_eq!(table.name(), if matches!(command, 2 | 3) { table_after } else { table_before });
}

fn exercise_unicode(package: &Package) {
    let result = package
        .edit_names()
        .rename_sheet(SheetSelector::index(0), UNICODE_NAME)
        .and_then(|edit| edit.commit());
    observe_result(result);
}

fn exercise_redacted_selector_error(package: &Package) {
    if let Err(error) = package
        .edit_names()
        .rename_sheet(SheetSelector::name(PRIVATE_SELECTOR), "safe")
    {
        observe_redacted(error, PRIVATE_SELECTOR);
    }
}

fn exercise_redacted_malformed_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        observe_redacted(error, std::str::from_utf8(PRIVATE_MALFORMED_INPUT).unwrap());
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, fuzz_options()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Numbers input must be rejected"),
    }
}

fn replacement(data: &[u8], role: &str) -> String {
    let start = data.len().min(8);
    let end = data.len().min(start.saturating_add(MAX_NAME_BYTES));
    let value = String::from_utf8_lossy(&data[start..end]);
    let sanitized: String = value.chars().filter(|character| *character != '\0').collect();
    format!("Litchi {role} 名 {sanitized}")
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a Numbers package to memory must succeed: {error}"));
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
