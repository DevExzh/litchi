#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    table::headers::{
        Count, Settings,
        transaction::{Error as HeaderError, Path},
    },
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
const PRIVATE_SHEET: &str = "__litchi_private_header_sheet_b713__";
const PRIVATE_TABLE: &str = "__litchi_private_header_table_b713__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_numbers_headers_input_b713__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Reuse the bounded prefix as semantic commands against a genuine
    // Numbers package so CRC-protected arbitrary ingress does not starve the
    // focused header/footer transaction of successful deep operations.
    exercise_package(native_package(), data);
    exercise_count_validation(data);
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
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Numbers fuzz semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Numbers fuzz projection limits: {error}")
                })
                .with_formula_render_limits(MAX_FORMULA_WORK, MAX_FORMULA_DEPTH)
                .unwrap_or_else(|error| unreachable!("valid Numbers fuzz formula limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options())
            .unwrap_or_else(|error| panic!("native Numbers headers fuzz seed must open: {error}"));
        package
            .table_header_settings(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| {
                panic!("native Numbers fuzz seed must expose table headers: {error}")
            });
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    observe_result(package.table_header_settings(
        SheetSelector::index(usize::from(read_u16(data, 2))),
        TableSelector::index(usize::from(read_u16(data, 4))),
    ));
    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        observe_result(package.table_header_settings(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
        ));
    }
    exercise_redacted_selectors(package);

    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let before = match package.table_header_settings(sheet, table) {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_settings(before);
    let after = if control(data, 0) & 1 == 0 {
        before
    } else {
        changed_settings(before, data)
    };

    let Ok(edit) = package.edit_table_headers(sheet, table) else {
        return;
    };
    assert_eq!(edit.path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(edit.settings(), before);
    let edit = edit.set(after);
    assert_eq!(edit.settings(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert_eq!(diagnostics.touched_components(), 1);
        assert!(diagnostics.deleted_previews() <= 3);
    }
    assert_eq!(
        commit
            .package()
            .table_header_settings(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("committed table headers must be readable: {error}")),
        after,
    );
    black_box((&patch, diagnostics));

    let source_bytes = package_bytes(package);
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == committed_bytes);
    let applied = package
        .apply_table_headers(&patch)
        .unwrap_or_else(|error| panic!("fresh table-header patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    assert_eq!(
        applied
            .package()
            .table_header_settings(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("applied table headers must be readable: {error}")),
        after,
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        assert!(matches!(
            applied.package().apply_table_headers(&patch),
            Err(HeaderError::PatchConflict)
        ));
        assert!(matches!(
            package.apply_table_headers(&inverse),
            Err(HeaderError::PatchConflict)
        ));
    }
    let restored = applied
        .package()
        .apply_table_headers(&inverse)
        .unwrap_or_else(|error| panic!("fresh table-header inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .table_header_settings(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| panic!("restored table headers must be readable: {error}")),
        before,
    );
    assert_eq!(package_bytes(restored.package()), source_bytes);
}

fn changed_settings(before: Settings, data: &[u8]) -> Settings {
    Settings {
        header_rows: Some(count(data, 1)),
        header_columns: Some(count(data, 2)),
        footer_rows: Some(count(data, 3)),
        header_rows_frozen: Some(!before.header_rows_are_frozen()),
        header_columns_frozen: Some(!before.header_columns_are_frozen()),
        repeating_header_rows_enabled: Some(!before.repeats_header_rows()),
        repeating_header_columns_enabled: Some(!before.repeats_header_columns()),
    }
}

fn count(data: &[u8], offset: usize) -> Count {
    Count::new(usize::from(control(data, offset) % 5 + 1))
        .unwrap_or_else(|error| unreachable!("1..=5 is a valid Numbers header count: {error}"))
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

fn exercise_count_validation(data: &[u8]) {
    observe_result(Count::new(usize::from(control(data, 6))));
}

fn exercise_redacted_selectors(package: &Package) {
    if let Err(error) =
        package.table_header_settings(SheetSelector::name(PRIVATE_SHEET), TableSelector::index(0))
    {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) =
        package.table_header_settings(SheetSelector::index(0), TableSelector::name(PRIVATE_TABLE))
    {
        observe_redacted(error, PRIVATE_TABLE);
    }
}

fn exercise_redacted_malformed_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        let private = std::str::from_utf8(PRIVATE_MALFORMED_INPUT)
            .unwrap_or_else(|error| unreachable!("private sentinel is UTF-8: {error}"));
        observe_redacted(error, private);
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, fuzz_options()) {
        Err(PackageError::InputTooLarge { observed, maximum }) => {
            assert_eq!(observed, OVERSIZED_INPUT_BYTES as u64);
            assert_eq!(maximum, MAX_INPUT_BYTES);
            black_box((observed, maximum));
        },
        Err(PackageError::Archive(error)) => {
            // Borrowed in-memory ingress is rejected by the shared archive
            // layer before the Numbers streaming adapter can construct its
            // equivalent `InputTooLarge` variant.
            assert_eq!(
                error.to_string(),
                format!(
                    "iWork archive input bytes limit exceeded: observed {OVERSIZED_INPUT_BYTES}, maximum {MAX_INPUT_BYTES}"
                )
            );
            black_box(error);
        },
        Err(error) => panic!("an oversized Numbers input must return a typed input-byte limit: {error}"),
        Ok(_) => panic!("an oversized Numbers input must be rejected"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap_or_else(|error| {
        panic!("writing a Numbers package to memory must succeed: {error}")
    });
    bytes
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
