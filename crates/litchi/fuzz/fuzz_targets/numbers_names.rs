#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    sheet::order::{Commit as SheetOrderCommit, Error as SheetOrderError},
    table::title::{Error as TitleError, Path as TitlePath, Settings as TitleSettings},
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
const PRIVATE_TITLE_SHEET: &str = "__litchi_private_title_sheet_6d1a__";
const PRIVATE_TITLE_TABLE: &str = "__litchi_private_title_table_6d1a__";
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
            .unwrap_or_else(|error| panic!("native Numbers names fuzz seed must open: {error}"));
        let sheet = package
            .document()
            .sheets()
            .first()
            .unwrap_or_else(|| panic!("native Numbers names fuzz seed must have a sheet"));
        assert!(
            sheet.tables().next().is_some(),
            "native Numbers names fuzz seed must have a table"
        );
        package
            .table_title_settings(SheetSelector::index(0), TableSelector::index(0))
            .unwrap_or_else(|error| {
                panic!("native Numbers names fuzz seed must expose a table title: {error}")
            });
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    exercise_sheet_order(package, data);
    exercise_table_title(package, data);

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
        0 => package
            .edit_names()
            .rename_sheet(SheetSelector::index(0), &sheet_name),
        1 => package
            .edit_names()
            .rename_sheet(SheetSelector::index(0), sheet_after.as_ref()),
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
    assert_names(
        commit.package(),
        command,
        &sheet_name,
        &table_name,
        &sheet_after,
        &table_after,
    );
    black_box((&patch, diagnostics));

    let applied = package
        .apply_names(&patch)
        .unwrap_or_else(|error| panic!("fresh Numbers names patch must apply: {error}"));
    assert_names(
        applied.package(),
        command,
        &sheet_name,
        &table_name,
        &sheet_after,
        &table_after,
    );
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(commit.package())
    );

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
    assert_names(
        restored.package(),
        0,
        &sheet_name,
        &table_name,
        &sheet_after,
        &table_after,
    );
    assert_eq!(package_bytes(restored.package()), package_bytes(package));
}

fn exercise_table_title(package: &Package, data: &[u8]) {
    observe_result(package.table_title_settings(
        SheetSelector::index(usize::from(read_u16(data, 3))),
        TableSelector::index(usize::from(read_u16(data, 5))),
    ));
    if let Some(sheet) = package.document().sheets().first() {
        if let Some(table) = sheet.tables().next() {
            observe_result(package.table_title_settings(
                SheetSelector::name(sheet.name()),
                TableSelector::name(table.name()),
            ));
        }
    }
    exercise_title_selector_errors(package);
    exercise_title_presence(data);

    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let before = match package.table_title_settings(sheet, table) {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_title_settings(before);

    // A transaction that preserves both optional-field presences must retain
    // the exact source package and publish no touched components.
    let no_op = match package.edit_table_title(sheet, table) {
        Ok(edit) => edit.set(before),
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(no_op.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(no_op.settings(), before);
    let no_op_commit = match no_op.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let no_op_patch = no_op_commit.patch().clone();
    let no_op_diagnostics = no_op_commit.diagnostics();
    assert!(no_op_patch.is_noop());
    assert_eq!(no_op_patch.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(no_op_patch.before(), before);
    assert_eq!(no_op_patch.after(), before);
    assert!(!no_op_diagnostics.changed());
    assert_eq!(no_op_diagnostics.touched_components(), 0);
    assert_eq!(no_op_diagnostics.deleted_previews(), 0);
    assert!(!no_op_diagnostics.full_reparse_performed());
    let source_bytes = package_bytes(package);
    assert_eq!(package_bytes(no_op_commit.package()), source_bytes);
    let no_op_applied = package
        .apply_table_title(&no_op_patch)
        .unwrap_or_else(|error| panic!("fresh no-op table-title patch must apply: {error}"));
    assert_eq!(package_bytes(no_op_applied.package()), source_bytes);
    assert_eq!(no_op_patch.inverse().inverse(), no_op_patch);
    black_box((&no_op_patch, no_op_diagnostics));

    let after = changed_title_settings(before, data);
    assert_ne!(after, before);
    let edit = match package.edit_table_title(sheet, table) {
        Ok(edit) => edit.set(after),
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(edit.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(edit.settings(), after);
    let commit = match edit.commit() {
        // A supported exact profile reaches apply/inverse/conflict checks.
        Ok(commit) => commit,
        // Other valid Numbers inputs may be intentionally outside the focused
        // producer profile; still retain this bounded public-API error path.
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert!(!patch.is_noop());
    assert_eq!(patch.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert!(diagnostics.changed());
    assert_eq!(diagnostics.touched_components(), 1);
    assert!(diagnostics.deleted_previews() <= 3);
    assert!(diagnostics.full_reparse_performed());
    assert_eq!(
        commit
            .package()
            .table_title_settings(sheet, table)
            .unwrap_or_else(|error| panic!("committed table title must be readable: {error}")),
        after,
    );
    let committed_bytes = package_bytes(commit.package());
    assert_ne!(committed_bytes, source_bytes);
    black_box((&patch, diagnostics));

    let applied = package
        .apply_table_title(&patch)
        .unwrap_or_else(|error| panic!("fresh table-title patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    assert_eq!(
        applied
            .package()
            .table_title_settings(sheet, table)
            .unwrap_or_else(|error| panic!("applied table title must be readable: {error}")),
        after,
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    assert!(matches!(
        applied.package().apply_table_title(&patch),
        Err(TitleError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_table_title(&inverse),
        Err(TitleError::PatchConflict)
    ));
    let restored = applied
        .package()
        .apply_table_title(&inverse)
        .unwrap_or_else(|error| panic!("fresh table-title inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_title_settings(sheet, table)
            .unwrap_or_else(|error| panic!("restored table title must be readable: {error}")),
        before,
    );
}

fn changed_title_settings(before: TitleSettings, data: &[u8]) -> TitleSettings {
    let values = [None, Some(false), Some(true)];
    let candidate = TitleSettings::new(
        values[usize::from(control(data, 0)) % values.len()],
        values[usize::from(control(data, 1)) % values.len()],
    );
    if candidate != before {
        candidate
    } else {
        TitleSettings::new(Some(!before.is_visible()), before.outlined())
    }
}

fn exercise_title_presence(data: &[u8]) {
    let values = [None, Some(false), Some(true)];
    let visible = values[usize::from(control(data, 6)) % values.len()];
    let outlined = values[usize::from(control(data, 7)) % values.len()];
    let settings = TitleSettings::new(visible, outlined);
    assert_eq!(settings.visible(), visible);
    assert_eq!(settings.outlined(), outlined);
    assert_eq!(settings.is_visible(), visible == Some(true));
    assert_eq!(settings.is_outlined(), outlined == Some(true));
    observe_title_settings(settings);
}

fn observe_title_settings(settings: TitleSettings) {
    black_box((
        settings.visible(),
        settings.outlined(),
        settings.is_visible(),
        settings.is_outlined(),
    ));
}

fn exercise_title_selector_errors(package: &Package) {
    assert!(matches!(
        package.table_title_settings(
            SheetSelector::index(package.document().sheets().len()),
            TableSelector::index(0),
        ),
        Err(TitleError::SheetNotFound)
    ));
    if let Err(error) = package.table_title_settings(
        SheetSelector::name(PRIVATE_TITLE_SHEET),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_TITLE_SHEET);
    }
    if let Err(error) = package.table_title_settings(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TITLE_TABLE),
    ) {
        observe_redacted(error, PRIVATE_TITLE_TABLE);
    }
}

fn exercise_sheet_order(package: &Package, data: &[u8]) {
    let sheets = package.document().sheets();
    let sheet_count = sheets.len();
    let Some(_) = sheets.first() else {
        return;
    };
    let source = usize::from(read_u16(data, 3)) % sheet_count;
    let source_name = sheets[source].name();
    let before: Vec<String> = sheets.iter().map(|sheet| sheet.name().to_owned()).collect();

    if let Err(error) = package
        .edit_sheet_order()
        .move_sheet(SheetSelector::name(PRIVATE_SELECTOR), 0)
    {
        observe_redacted(error, PRIVATE_SELECTOR);
    }
    assert!(matches!(
        package
            .edit_sheet_order()
            .move_sheet(SheetSelector::index(sheet_count), 0),
        Err(SheetOrderError::SheetNotFound)
    ));
    assert!(matches!(
        package
            .edit_sheet_order()
            .move_sheet(SheetSelector::index(source), sheet_count),
        Err(SheetOrderError::DestinationOutOfRange { .. })
    ));

    let command = control(data, 0) % 3;
    let destination = if command == 0 || sheet_count == 1 {
        source
    } else {
        (source + 1) % sheet_count
    };
    let staged = if command == 2 {
        package
            .edit_sheet_order()
            .move_sheet(SheetSelector::name(source_name), destination)
    } else {
        package
            .edit_sheet_order()
            .move_sheet(SheetSelector::index(source), destination)
    };
    let edit = match staged {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    publish_sheet_order_and_reverse(package, &before, source, destination, commit);
}

fn publish_sheet_order_and_reverse(
    package: &Package,
    before: &[String],
    source: usize,
    destination: usize,
    commit: SheetOrderCommit,
) {
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.source_position(), source);
    assert_eq!(patch.destination_position(), destination);
    assert_eq!(patch.is_noop(), source == destination);
    assert_eq!(diagnostics.changed(), source != destination);
    assert_eq!(diagnostics.full_reparse_performed(), source != destination);
    if source == destination {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    let expected = moved_sheet_order(before, source, destination);
    assert_sheet_order(commit.package(), &expected);
    let source_bytes = package_bytes(package);
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == committed_bytes);
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_sheet_order(&patch)
        .unwrap_or_else(|error| panic!("fresh sheet-order patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    assert_sheet_order(applied.package(), &expected);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        assert!(applied.package().apply_sheet_order(&patch).is_err());
        assert!(package.apply_sheet_order(&inverse).is_err());
    }
    let restored = applied
        .package()
        .apply_sheet_order(&inverse)
        .unwrap_or_else(|error| panic!("fresh sheet-order inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_sheet_order(restored.package(), before);
}

fn moved_sheet_order(before: &[String], source: usize, destination: usize) -> Vec<String> {
    let mut expected = before.to_vec();
    let moved = expected.remove(source);
    expected.insert(destination, moved);
    expected
}

fn assert_sheet_order(package: &Package, expected: &[String]) {
    let actual: Vec<&str> = package
        .document()
        .sheets()
        .iter()
        .map(|sheet| sheet.name())
        .collect();
    assert_eq!(actual.len(), expected.len());
    assert!(
        actual
            .iter()
            .zip(expected)
            .all(|(actual, expected)| *actual == expected)
    );
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
    assert_eq!(
        sheet.name(),
        if matches!(command, 1 | 3) {
            sheet_after
        } else {
            sheet_before
        }
    );
    assert_eq!(
        table.name(),
        if matches!(command, 2 | 3) {
            table_after
        } else {
            table_before
        }
    );
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
    let sanitized: String = value
        .chars()
        .filter(|character| *character != '\0')
        .collect();
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
    package.write_to(&mut bytes).unwrap_or_else(|error| {
        panic!("writing a Numbers package to memory must succeed: {error}")
    });
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
