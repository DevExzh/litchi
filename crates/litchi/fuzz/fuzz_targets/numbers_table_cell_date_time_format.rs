#![no_main]

//! Bounded selector-first fuzzing for Numbers table-cell Date & Time formats.
//!
//! The target offers arbitrary bytes to bounded Numbers ingress and replays
//! the same command prefix against a deterministic source-built Date & Time
//! package.  This keeps the semantic owner reachable even when CRC-protected
//! ZIP mutation is rejected before a table can be selected.  Only selectors,
//! checked cell positions, archive-free [`DateTime`] values, and exact-source
//! patches cross the package boundary; native identifiers and wire messages
//! remain inside the fixture module and the Numbers adapter.

use std::{
    fmt::{Debug, Display},
    hint::black_box,
    sync::OnceLock,
};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::{DateTime, date_time::MAX_PATTERN_BYTES, date_time::transaction::Error},
};

#[path = "../../../litchi-numbers/tests/support/table_cell_data_format_fixture.rs"]
mod fixture;

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
const PRIVATE_SHEET: &str = "__litchi_private_datetime_sheet_1f20__";
const PRIVATE_TABLE: &str = "__litchi_private_datetime_table_1f20__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_datetime_input_1f20__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

const FIXTURE_POSITION: CellPosition = CellPosition::new(0, 0);
const FIXTURE_SIBLING_POSITION: CellPosition = CellPosition::new(0, 1);
// B3 is a stable Number scalar in the checked-in native Numbers seed.
const NATIVE_POSITION: CellPosition = CellPosition::new(2, 1);

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data, NATIVE_POSITION, false),
        Err(error) => observe_error(error),
    }

    // Replay commands against a valid source so reads, edits, patch replay,
    // inverse conflicts, and output verification are not starved by ingress.
    exercise_package(date_time_package(), data, FIXTURE_POSITION, true);
    exercise_wrong_family(number_package(), FIXTURE_POSITION);
    exercise_wrong_family(text_package(), FIXTURE_POSITION);
    exercise_wrong_family(native_package(), NATIVE_POSITION);
    exercise_shared_format_cow(data);
    exercise_locked_source(data);
    exercise_malformed_sources();
    exercise_constructor_boundaries(data);
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_semantic_limit();
    exercise_output_limit();
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
        .unwrap_or_else(|error| unreachable!("valid Date & Time archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Date & Time semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Date & Time projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn date_time_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::DateTime,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Date & Time fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("deterministic Date & Time source must open: {error}"))
    })
}

fn unshared_date_time_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::DateTime,
            fixture::FormatSharing::Unshared,
        )
        .unwrap_or_else(|error| panic!("unshared Date & Time fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic unshared Date & Time source must open: {error}")
        })
    })
}

fn number_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Number))
}

fn text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Text))
}

fn package_for_family(family: fixture::FormatFamily) -> Package {
    let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)
        .unwrap_or_else(|error| panic!("{family:?} fixture must build: {error}"));
    Package::from_bytes_with_options(&source, options())
        .unwrap_or_else(|error| panic!("deterministic {family:?} source must open: {error}"))
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers Date & Time seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8], position: CellPosition, fixture_source: bool) {
    exercise_selector_reads(package, data, position);
    exercise_selector_errors(package, position);

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_date_time_format(0usize, 0usize, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            if let Err(error) = package.edit_table_cell_date_time_format(0usize, 0usize, position) {
                observe_error(error);
            }
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    let requested = requested_pattern(data);
    exercise_transaction(
        package,
        &source_bytes,
        before,
        requested,
        position,
        data,
        fixture_source,
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8], position: CellPosition) {
    let random_position = CellPosition::new(
        u32::from(control(data, 4) % 8),
        u32::from(control(data, 5) % 8),
    );
    observe_result(package.table_cell_date_time_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        random_position,
    ));
    observe_result(package.table_cell_date_time_format(0usize, 0usize, position));

    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        let by_name = package.table_cell_date_time_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        );
        let by_index = package.table_cell_date_time_format(0usize, 0usize, position);
        assert_eq!(by_name, by_index, "Date & Time selector forms disagree");
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let sheet_count = package.document().sheets().len();
    if let Err(error) = package.table_cell_date_time_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ) {
        observe_error(error);
    }
    if let Err(error) = package.table_cell_date_time_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_date_time_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_date_time_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(u32::MAX, u32::MAX),
    ) {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_transaction(
    package: &Package,
    source_bytes: &[u8],
    before: Option<DateTime>,
    requested: DateTime,
    position: CellPosition,
    data: &[u8],
    fixture_source: bool,
) {
    let edit = match package.edit_table_cell_date_time_format(0usize, 0usize, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    // Keep the source's Optional state for the no-op command and cover the
    // complete public lifecycle: set, clear, reset, apply, and inverse.
    let command = control(data, 0) & 3;
    let (edit, expected_after) = match command {
        0 => match before.clone() {
            Some(value) => (edit.set(value), before.clone()),
            None => (edit.clear(), None),
        },
        1 => (edit.set(requested.clone()), Some(requested)),
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
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(
        commit.diagnostics().full_reparse_performed(),
        !patch.is_noop()
    );
    assert_eq!(
        commit
            .package()
            .table_cell_date_time_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Date & Time candidate readback failed: {error}")),
        expected_after
    );
    black_box(commit.diagnostics());

    let applied = package
        .apply_table_cell_date_time_format(&patch)
        .unwrap_or_else(|error| panic!("fresh Date & Time patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_date_time_format(&patch)
                .is_err(),
            "changed Date & Time patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_date_time_format(&patch.inverse())
                .is_err(),
            "Date & Time inverse must reject the original source"
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_date_time_format(&inverse)
        .unwrap_or_else(|error| panic!("Date & Time inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_date_time_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Date & Time inverse readback failed: {error}")),
        before
    );

    // A changed transaction must preserve the sibling's semantic format and
    // the untouched source package.  This is deliberately package-level and
    // does not inspect native IDs or generated messages.
    if fixture_source && !patch.is_noop() {
        let sibling_before = package
            .table_cell_date_time_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("Date & Time sibling read failed: {error}"));
        assert_eq!(
            restored
                .package()
                .table_cell_date_time_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
                .unwrap_or_else(|error| panic!("Date & Time sibling inverse read failed: {error}")),
            sibling_before
        );
    }
    black_box((patch.before(), patch.after(), patch.path()));
}

fn exercise_wrong_family(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_date_time_format(0usize, 0usize, position);
    assert!(
        matches!(read, Err(Error::WrongFormatFamily { .. })),
        "non-DateTime cell must be refused by the Date & Time API: {read:?}"
    );
    let edit = package.edit_table_cell_date_time_format(0usize, 0usize, position);
    assert!(
        matches!(edit, Err(Error::WrongFormatFamily { .. })),
        "non-DateTime edit must be refused by the Date & Time API: {edit:?}"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_shared_format_cow(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let package = unshared_date_time_package();
        let source_bytes = package_bytes(package);
        let before = package
            .table_cell_date_time_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("unshared Date & Time source read failed: {error}"));
        let requested = requested_pattern(data);
        let result = package
            .edit_table_cell_date_time_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("unshared Date & Time edit open failed: {error}"))
            .set(requested)
            .commit();
        match result {
            Ok(commit) => {
                assert_eq!(commit.patch().before(), before.as_ref());
                assert_ne!(package_bytes(commit.package()), source_bytes);
                assert_eq!(package_bytes(package), source_bytes);
            },
            Err(error) => {
                observe_error(error);
                assert_eq!(package_bytes(package), source_bytes);
            },
        }
    });
}

fn exercise_locked_source(data: &[u8]) {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::DateTime,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("locked Date & Time source must build: {error}"));
        let locked = fixture::locked_table_package(&source)
            .unwrap_or_else(|error| panic!("locked Date & Time source must build: {error}"));
        Package::from_bytes_with_options(&locked, options())
            .unwrap_or_else(|error| panic!("locked Date & Time source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let result = package
        .edit_table_cell_date_time_format(0usize, 0usize, FIXTURE_POSITION)
        .and_then(|edit| edit.set(requested_pattern(data)).commit());
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_malformed_sources() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for corruption in [
            fixture::Corruption::DuplicateFormatKey,
            fixture::Corruption::MissingFormatEntry,
            fixture::Corruption::RefcountMismatch,
            fixture::Corruption::MalformedFormatPayload,
            fixture::Corruption::UnsupportedFormatType,
            fixture::Corruption::WrongCellFormatKey,
        ] {
            let source =
                fixture::corrupted_package_for(fixture::FormatFamily::DateTime, corruption)
                    .unwrap_or_else(|error| {
                        panic!("malformed Date & Time source must build: {error}")
                    });
            match Package::from_bytes_with_options(&source, options()) {
                Ok(package) => {
                    let before = package_bytes(&package);
                    observe_result(package.table_cell_date_time_format(
                        0usize,
                        0usize,
                        FIXTURE_POSITION,
                    ));
                    observe_result(package.edit_table_cell_date_time_format(
                        0usize,
                        0usize,
                        FIXTURE_POSITION,
                    ));
                    assert_eq!(package_bytes(&package), before);
                },
                Err(error) => observe_error(error),
            }
        }
    });
}

fn requested_pattern(data: &[u8]) -> DateTime {
    match control(data, 1) % 5 {
        0 => DateTime::iso_date(),
        1 => DateTime::time_24_hour_with_seconds(),
        2 => DateTime::iso_date_time_24_hour_with_seconds(),
        _ => {
            let mut value = String::new();
            let source = data.get(2..).unwrap_or_default();
            let source = &source[..source.len().min(512)];
            for character in String::from_utf8_lossy(source).chars() {
                if character != '\0' {
                    value.push(character);
                }
            }
            let value = value.trim();
            DateTime::new(value).unwrap_or_else(|_| DateTime::iso_date_time_24_hour_with_seconds())
        },
    }
}

fn exercise_constructor_boundaries(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        assert!(matches!(DateTime::new(""), Err(_)));
        assert!(matches!(DateTime::new("   \t"), Err(_)));
        assert!(matches!(DateTime::new("yyyy\0MM"), Err(_)));
        let exact = "x".repeat(MAX_PATTERN_BYTES);
        assert!(DateTime::new(&exact).is_ok());
        let oversized = "x".repeat(MAX_PATTERN_BYTES + 1);
        assert!(DateTime::new(&oversized).is_err());
    });
    black_box(requested_pattern(data));
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
            Err(error) => observe_error(error),
            Ok(_) => panic!("oversized Date & Time input must be rejected"),
        }
    });
}

fn exercise_semantic_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let semantic = PackageSemanticLimits::new(1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("one-entry Date & Time limits must validate: {error}"));
        let constrained = PackageReadOptions::new(options().archive(), semantic);
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::DateTime,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Date & Time semantic-limit source: {error}"));
        let result = Package::from_bytes_with_options(&source, constrained);
        assert!(
            result.is_err(),
            "one-entry limits admitted Date & Time source"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn exercise_output_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::DateTime,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Date & Time output-limit source: {error}"));
        let limits = PackageLimits::new(
            u64::try_from(source.len()).unwrap_or(u64::MAX),
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("Date & Time output limits must validate: {error}"));
        let package = Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(limits, PackageSemanticLimits::default()),
        )
        .unwrap_or_else(|error| panic!("Date & Time exact-limit source must open: {error}"));
        let before = package_bytes(&package);
        let result = package
            .edit_table_cell_date_time_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Date & Time output-limit edit open failed: {error}"))
            .set(
                DateTime::new(&"x".repeat(MAX_PATTERN_BYTES)).unwrap_or_else(|error| {
                    panic!("maximum Date & Time pattern must construct: {error}")
                }),
            )
            .commit();
        assert!(
            result.is_err(),
            "exact Date & Time output limits unexpectedly admitted maximum growth"
        );
        black_box(result.err().map(|error| error.to_string()));
        assert_eq!(package_bytes(&package), before);
    });
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Date & Time package failed: {error}"));
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
        "Date & Time error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Date & Time debug leaked private input"
    );
    black_box((display, debug));
}
