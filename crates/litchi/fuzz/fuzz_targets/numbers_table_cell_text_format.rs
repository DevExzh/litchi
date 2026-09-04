#![no_main]

//! Bounded selector-first fuzzing for the Numbers table-cell Text format.
//!
//! Text is a marker format, but its native owner still has to resolve a cell,
//! preserve the exact format-list graph, and publish reversible patches.  The
//! target uses only selectors, checked positions, the archive-free `Text`
//! value, and the public package transaction API.  The deterministic fixture
//! keeps the lifecycle reachable even when arbitrary ZIP input is rejected at
//! ingress.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::{Text, text::transaction::Error as TextError},
};

#[path = "../../../litchi-numbers/tests/support/table_cell_data_format_fixture.rs"]
mod text_fixture;

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
const PRIVATE_SHEET: &str = "__litchi_private_text_format_sheet_2d73__";
const PRIVATE_TABLE: &str = "__litchi_private_text_format_table_2d73__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_text_format_input_2d73__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

const TEXT_POSITION: CellPosition = CellPosition::new(0, 0);
const NATIVE_NUMBER_POSITION: CellPosition = CellPosition::new(2, 1);

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data, TEXT_POSITION),
        Err(error) => observe_error(error),
    }

    // Replay every command against valid Text sources so package-level
    // reads, edits, exact apply, inverse, and candidate verification are not
    // starved by CRC-protected ZIP mutation.
    exercise_package(shared_text_package(), data, TEXT_POSITION);
    exercise_package(inherited_text_package(), data, TEXT_POSITION);
    exercise_converted_noop();
    exercise_wrong_family(number_package(), TEXT_POSITION);
    exercise_wrong_family(percentage_package(), TEXT_POSITION);
    exercise_wrong_family(currency_package(), TEXT_POSITION);
    exercise_wrong_family(scientific_package(), TEXT_POSITION);
    exercise_wrong_family(fraction_package(), TEXT_POSITION);
    exercise_wrong_family(native_package(), NATIVE_NUMBER_POSITION);
    exercise_locked_source();
    exercise_malformed_sources();
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
        .unwrap_or_else(|error| unreachable!("valid Text-format archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Text-format semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Text-format projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn shared_text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = text_fixture::synthetic_package_for(
            text_fixture::FormatFamily::Text,
            text_fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Text fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("deterministic shared Text source must open: {error}"))
    })
}

fn inherited_text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = text_fixture::text_inherited_first_package()
            .unwrap_or_else(|error| panic!("inherited Text fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic inherited Text source must open: {error}")
        })
    })
}

fn converted_text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = text_fixture::text_converted_package()
            .unwrap_or_else(|error| panic!("converted Text fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic converted Text source must open: {error}")
        })
    })
}

fn exercise_converted_noop() {
    let package = converted_text_package();
    let source_bytes = package_bytes(package);
    let before = package
        .table_cell_text_format(0usize, 0usize, TEXT_POSITION)
        .unwrap_or_else(|error| panic!("converted Text source must read: {error}"));
    assert_eq!(before, Some(Text));

    let commit = package
        .edit_table_cell_text_format(0usize, 0usize, TEXT_POSITION)
        .unwrap_or_else(|error| panic!("converted Text source must edit: {error}"))
        .set(Text)
        .commit()
        .unwrap_or_else(|error| panic!("converted Text no-op must commit: {error}"));
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), Some(&Text));
    assert_eq!(commit.patch().after(), Some(&Text));
    assert_eq!(package_bytes(commit.package()), source_bytes);
    assert_eq!(package_bytes(package), source_bytes);
    black_box(commit.diagnostics());
}

fn number_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(text_fixture::FormatFamily::Number))
}

fn percentage_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(text_fixture::FormatFamily::Percentage))
}

fn currency_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(text_fixture::FormatFamily::Currency))
}

fn scientific_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(text_fixture::FormatFamily::Scientific))
}

fn fraction_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(text_fixture::FormatFamily::Fraction))
}

fn package_for_family(family: text_fixture::FormatFamily) -> Package {
    let source = text_fixture::synthetic_package_for(family, text_fixture::FormatSharing::Shared)
        .unwrap_or_else(|error| panic!("{family:?} fixture must build: {error}"));
    Package::from_bytes_with_options(&source, options())
        .unwrap_or_else(|error| panic!("deterministic {family:?} source must open: {error}"))
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers Text-format seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8], position: CellPosition) {
    exercise_selector_reads(package, data, position);
    exercise_selector_errors(package, position);

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_text_format(0usize, 0usize, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            if let Err(error) = package.edit_table_cell_text_format(0usize, 0usize, position) {
                observe_error(error);
            }
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    exercise_transaction(package, &source_bytes, before, data, position);
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8], position: CellPosition) {
    let random_position = CellPosition::new(
        u32::from(control(data, 4) % 8),
        u32::from(control(data, 5) % 8),
    );
    observe_result(package.table_cell_text_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        random_position,
    ));
    observe_result(package.table_cell_text_format(0usize, 0usize, position));

    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        observe_result(package.table_cell_text_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        ));
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let sheet_count = package.document().sheets().len();
    if let Err(error) = package.table_cell_text_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ) {
        observe_error(error);
    }
    if let Err(error) = package.table_cell_text_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_text_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_text_format(
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
    before: Option<Text>,
    data: &[u8],
    position: CellPosition,
) {
    let edit = match package.edit_table_cell_text_format(0usize, 0usize, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    // Preserve the source's Optional state for an exact no-op command, then
    // cover explicit set, clear, and reset in the remaining command classes.
    let command = control(data, 0) & 3;
    let (edit, expected_after) = match command {
        0 => match before {
            Some(value) => (edit.set(value), before),
            None => (edit.clear(), None),
        },
        1 => (edit.set(Text), Some(Text)),
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
    assert_eq!(
        commit.diagnostics().touched_components(),
        usize::from(!patch.is_noop())
    );
    assert_eq!(
        commit.diagnostics().full_reparse_performed(),
        !patch.is_noop()
    );
    assert_eq!(
        commit
            .package()
            .table_cell_text_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Text-format candidate readback failed: {error}")),
        expected_after
    );
    black_box(commit.diagnostics());

    let applied = package
        .apply_table_cell_text_format(&patch)
        .unwrap_or_else(|error| panic!("fresh Text-format patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .table_cell_text_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("applied Text-format patch readback failed: {error}")),
        expected_after
    );

    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_text_format(&patch)
                .is_err(),
            "a changed Text-format patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_text_format(&patch.inverse())
                .is_err(),
            "a Text-format inverse must reject the original source"
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_text_format(&inverse)
        .unwrap_or_else(|error| panic!("Text-format inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_text_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Text-format inverse readback failed: {error}")),
        before
    );
    black_box((patch.before(), patch.after(), patch.path()));
}

fn exercise_wrong_family(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_text_format(0usize, 0usize, position);
    assert!(
        matches!(read, Err(TextError::WrongFormatFamily { .. })),
        "non-Text cell must be refused by the Text API: {read:?}"
    );
    let edit = package.edit_table_cell_text_format(0usize, 0usize, position);
    assert!(
        matches!(edit, Err(TextError::WrongFormatFamily { .. })),
        "non-Text cell edit must be refused by the Text API: {edit:?}"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_locked_source() {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = text_fixture::text_inherited_first_package()
            .unwrap_or_else(|error| panic!("locked Text source must build: {error}"));
        let locked = text_fixture::locked_table_package(&source)
            .unwrap_or_else(|error| panic!("locked Text source must build: {error}"));
        Package::from_bytes_with_options(&locked, options())
            .unwrap_or_else(|error| panic!("locked Text source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let before = package.table_cell_text_format(0usize, 0usize, TEXT_POSITION);
    let edit = package.edit_table_cell_text_format(0usize, 0usize, TEXT_POSITION);
    if let (Ok(_before), Ok(edit)) = (before, edit) {
        let result = edit.set(Text).commit();
        if let Err(error) = result {
            observe_error(error);
        }
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_malformed_sources() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for corruption in [
            text_fixture::Corruption::DuplicateFormatKey,
            text_fixture::Corruption::MissingFormatEntry,
            text_fixture::Corruption::RefcountMismatch,
            text_fixture::Corruption::MalformedFormatPayload,
            text_fixture::Corruption::UnterminatedUnknownGroup,
            text_fixture::Corruption::UnsupportedFormatType,
            text_fixture::Corruption::WrongCellFormatKey,
        ] {
            let source =
                text_fixture::corrupted_package_for(text_fixture::FormatFamily::Text, corruption)
                    .unwrap_or_else(|error| panic!("malformed Text source must build: {error}"));
            match Package::from_bytes_with_options(&source, options()) {
                Ok(package) => {
                    let before = package_bytes(&package);
                    observe_result(package.table_cell_text_format(0usize, 0usize, TEXT_POSITION));
                    observe_result(package.edit_table_cell_text_format(
                        0usize,
                        0usize,
                        TEXT_POSITION,
                    ));
                    assert_eq!(package_bytes(&package), before);
                },
                Err(error) => observe_error(error),
            }
        }
    });
}

fn exercise_redacted_ingress() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for input in [PRIVATE_INPUT, b"".as_slice(), b"not-a-numbers-package"] {
            if let Err(error) = Package::from_bytes_with_options(input, options()) {
                if input.is_empty() {
                    observe_error(error);
                } else {
                    let private = std::str::from_utf8(input)
                        .unwrap_or_else(|error| unreachable!("sentinel is UTF-8: {error}"));
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
                black_box(error);
            },
            Err(error) => observe_error(error),
            Ok(_) => panic!("oversized Numbers Text-format input must be rejected"),
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
        .unwrap_or_else(|error| panic!("writing Text-format package failed: {error}"));
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
        "Text-format error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Text-format debug leaked private input"
    );
    black_box((display, debug));
}
