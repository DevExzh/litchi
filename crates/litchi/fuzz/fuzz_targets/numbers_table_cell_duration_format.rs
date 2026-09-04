#![no_main]

//! Bounded selector-first fuzzing for Numbers table-cell Duration formats.
//!
//! Arbitrary bytes first exercise bounded Numbers ingress, then the same
//! command prefix is replayed against deterministic Duration packages.  The
//! focused path keeps all three presentation styles, every ordered unit
//! range, automatic/custom units, exact-source patches, COW/refcount edges,
//! and malformed BNC shapes reachable even when CRC-protected ZIP mutation is
//! rejected before a table can be selected.  Native identifiers and wire
//! messages stay inside the fixture and the adapter; only archive-free
//! Duration values and selectors cross the package boundary.

use std::{
    fmt::{Debug, Display},
    hint::black_box,
    io,
    sync::OnceLock,
};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::{
        Duration,
        duration::{Style, Unit, UnitRange, Units},
    },
};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::{tsk, tst};
use litchi_numbers_wire::{
    BncCell, CellDataFormatKind, DATE_TIME_CELL_FORMAT_KIND, EXPLICIT_DECIMAL_FORMAT,
    EXPLICIT_DURATION_FORMAT, EXPLICIT_DURATION_WITH_NUMBER_FORMAT,
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
const PRIVATE_SHEET: &str = "__litchi_private_duration_format_sheet_7c42__";
const PRIVATE_TABLE: &str = "__litchi_private_duration_format_table_7c42__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_duration_format_input_7c42__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

const FIRST_FORMAT_KEY: u32 = 1;
const SECOND_FORMAT_KEY: u32 = 2;
const FIXTURE_POSITION: CellPosition = CellPosition::new(0, 0);
const FIXTURE_SIBLING_POSITION: CellPosition = CellPosition::new(0, 1);
// B3 is a stable explicit Number cell in the checked-in native Numbers seed.
const NATIVE_POSITION: CellPosition = CellPosition::new(2, 1);

const ALL_STYLES: [Style; 3] = [Style::Colon, Style::Abbreviated, Style::FullNames];
const ALL_UNITS: [Unit; 6] = [
    Unit::Weeks,
    Unit::Days,
    Unit::Hours,
    Unit::Minutes,
    Unit::Seconds,
    Unit::Milliseconds,
];

fuzz_target!(|data: &[u8]| {
    // Arbitrary ingress remains useful for package parser coverage, but the
    // semantic owner is exercised independently below on valid sources.
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_arbitrary_package(&package, data),
        Err(error) => observe_error(error),
    }

    exercise_package(duration_package(), data, FIXTURE_POSITION, true);

    // Duration is a typed family.  Every other focused display family and a
    // real native Number cell must be refused without publishing bytes.
    exercise_wrong_family(number_package(), FIXTURE_POSITION);
    exercise_wrong_family(percentage_package(), FIXTURE_POSITION);
    exercise_wrong_family(currency_package(), FIXTURE_POSITION);
    exercise_wrong_family(scientific_package(), FIXTURE_POSITION);
    exercise_wrong_family(fraction_package(), FIXTURE_POSITION);
    exercise_wrong_family(date_time_package(), FIXTURE_POSITION);
    exercise_wrong_family(text_package(), FIXTURE_POSITION);
    exercise_wrong_family(native_package(), NATIVE_POSITION);

    exercise_all_styles_and_units();
    exercise_shared_format_cow(data);
    exercise_secondary_number(data);
    exercise_inherited_marker(data);
    exercise_malformed_sources();
    exercise_locked_source(data);
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
        .unwrap_or_else(|error| unreachable!("valid Duration archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Duration semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid Duration projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn duration_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Duration,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Duration fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("deterministic Duration source must open: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers Duration seed must open: {error}"))
    })
}

fn package_for_family(family: fixture::FormatFamily) -> Package {
    let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)
        .unwrap_or_else(|error| panic!("{family:?} fixture must build: {error}"));
    Package::from_bytes_with_options(&source, options())
        .unwrap_or_else(|error| panic!("deterministic {family:?} source must open: {error}"))
}

fn number_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Number))
}

fn percentage_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Percentage))
}

fn currency_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Currency))
}

fn scientific_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Scientific))
}

fn fraction_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Fraction))
}

fn date_time_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::DateTime))
}

fn text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_family(fixture::FormatFamily::Text))
}

fn exercise_arbitrary_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let position = random_position(data);
    exercise_selector_reads(package, data, position);
    let result = package
        .edit_table_cell_duration_format(
            SheetSelector::index(usize::from(control(data, 0))),
            TableSelector::index(usize::from(control(data, 1))),
            position,
        )
        .and_then(|edit| edit.set(duration_from_bytes(data)).commit());
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_package(package: &Package, data: &[u8], position: CellPosition, fixture_source: bool) {
    exercise_selector_reads(package, data, position);
    exercise_selector_errors(package, position);

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_duration_format(0usize, 0usize, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            observe_result(package.edit_table_cell_duration_format(0usize, 0usize, position));
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    exercise_transaction(
        package,
        &source_bytes,
        before,
        duration_from_bytes(data),
        position,
        data,
        fixture_source,
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8], position: CellPosition) {
    let random_position = random_position(data);
    observe_result(package.table_cell_duration_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        random_position,
    ));
    observe_result(package.table_cell_duration_format(0usize, 0usize, position));

    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        let by_name = package.table_cell_duration_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        );
        let by_index = package.table_cell_duration_format(0usize, 0usize, position);
        assert_eq!(by_name, by_index, "Duration selector forms disagree");
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let sheet_count = package.document().sheets().len();
    if let Err(error) = package.table_cell_duration_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ) {
        observe_error(error);
    }
    if let Err(error) = package.table_cell_duration_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_duration_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_duration_format(
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
    before: Option<Duration>,
    requested: Duration,
    position: CellPosition,
    data: &[u8],
    fixture_source: bool,
) {
    let edit = match package.edit_table_cell_duration_format(0usize, 0usize, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    // Preserve the source Optional state for an exact no-op, then cover set,
    // clear, and reset independently.  The command prefix is intentionally
    // tiny; deterministic style/range sweeps below cover the complete value
    // matrix once per process.
    let command = control(data, 0) & 3;
    let (edit, expected_after) = match command {
        0 => match before {
            Some(value) => (edit.set(value), before),
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
    assert_eq!(
        commit.diagnostics().full_reparse_performed(),
        !patch.is_noop()
    );
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        if fixture_source {
            assert_exact_locality(source_bytes, &target_bytes);
            assert_non_format_bnc_bytes(source_bytes, &target_bytes);
        }
    }
    assert_eq!(
        commit
            .package()
            .table_cell_duration_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Duration candidate readback failed: {error}")),
        expected_after
    );
    black_box(commit.diagnostics());

    let applied = package
        .apply_table_cell_duration_format(&patch)
        .unwrap_or_else(|error| panic!("fresh Duration patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_duration_format(&patch)
                .is_err(),
            "changed Duration patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_duration_format(&patch.inverse())
                .is_err(),
            "Duration inverse must reject the original source"
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_duration_format(&inverse)
        .unwrap_or_else(|error| panic!("Duration inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_duration_format(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("Duration inverse readback failed: {error}")),
        before
    );

    if fixture_source && !patch.is_noop() {
        let sibling_before = package
            .table_cell_duration_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("Duration sibling read failed: {error}"));
        assert_eq!(
            restored
                .package()
                .table_cell_duration_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
                .unwrap_or_else(|error| panic!("Duration sibling inverse read failed: {error}")),
            sibling_before
        );
    }

    // A patch is bound to the exact source artifact, so a separately built
    // Duration package is a foreign/stale conflict even when its selectors
    // and semantic values look identical.
    if !patch.is_noop() {
        let foreign_source = fixture::synthetic_package_for(
            fixture::FormatFamily::Duration,
            fixture::FormatSharing::Unshared,
        )
        .unwrap_or_else(|error| panic!("foreign Duration source must build: {error}"));
        if let Ok(foreign) = Package::from_bytes_with_options(&foreign_source, options()) {
            assert!(
                foreign.apply_table_cell_duration_format(&patch).is_err(),
                "Duration patch unexpectedly applied to a foreign source"
            );
        }
    }
    black_box((patch.before(), patch.after(), patch.path()));
}

fn exercise_wrong_family(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_duration_format(0usize, 0usize, position);
    assert!(
        matches!(
            read,
            Err(litchi::numbers::cell::data_format::duration::transaction::Error::WrongFormatFamily { .. })
        ),
        "non-Duration cell must be refused by the Duration API: {read:?}"
    );
    let edit = package.edit_table_cell_duration_format(0usize, 0usize, position);
    assert!(
        matches!(
            edit,
            Err(litchi::numbers::cell::data_format::duration::transaction::Error::WrongFormatFamily { .. })
        ),
        "non-Duration edit must be refused by the Duration API: {edit:?}"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_all_styles_and_units() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = package_bytes(duration_package());
        for style in ALL_STYLES {
            // Explicit automatic-unit constructors are part of the public
            // value matrix even though they retain the same persisted range.
            let automatic = Duration::automatic(style);
            assert_eq!(automatic.style(), style);
            assert!(automatic.units().is_automatic());
            exercise_one_value(&source, automatic);

            for (largest_index, largest) in ALL_UNITS.iter().copied().enumerate() {
                for smallest in ALL_UNITS[largest_index..].iter().copied() {
                    let range = UnitRange::new(largest, smallest)
                        .unwrap_or_else(|error| panic!("ordered Duration range rejected: {error}"));
                    for units in [Units::Automatic(range), Units::Custom(range)] {
                        let value = Duration::new(style, units);
                        assert_eq!(value.style(), style);
                        assert_eq!(value.units(), units);
                        assert_eq!(value.units().range(), range);
                        exercise_one_value(&source, value);
                    }
                }
            }
        }
    });
}

fn exercise_one_value(source: &[u8], value: Duration) {
    let package = Package::from_bytes_with_options(source, options())
        .unwrap_or_else(|error| panic!("Duration matrix source must reopen: {error}"));
    let before = package
        .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration matrix source read failed: {error}"));
    let result = package
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration matrix edit open failed: {error}"))
        .set(value)
        .commit();
    match result {
        Ok(commit) => {
            assert_eq!(commit.patch().after(), Some(&value));
            assert_eq!(
                commit
                    .package()
                    .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
                    .unwrap_or_else(|error| panic!(
                        "Duration matrix candidate read failed: {error}"
                    )),
                Some(value)
            );
            let _ = black_box(commit);
        },
        Err(error) => panic!("valid Duration style/unit subset was refused: {error}"),
    }
    assert_eq!(
        package_bytes(&package),
        source,
        "Duration matrix edit mutated its source"
    );
    black_box(before);
}

fn exercise_shared_format_cow(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let package = duration_package();
        let source_bytes = package_bytes(package);
        let before = package
            .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("shared Duration source read failed: {error}"));
        let sibling_before = package
            .table_cell_duration_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("shared Duration sibling read failed: {error}"));
        let source_facts = fixture::format_entry_facts(&source_bytes)
            .unwrap_or_else(|error| panic!("shared Duration format facts: {error}"));
        assert_eq!(
            fixture::format_keys(&source_bytes)
                .unwrap_or_else(|error| panic!("shared Duration format keys: {error}")),
            vec![Some(FIRST_FORMAT_KEY), Some(FIRST_FORMAT_KEY)]
        );
        assert_eq!(source_facts, vec![(FIRST_FORMAT_KEY, 2)]);
        let requested = Duration::custom(
            Style::Colon,
            UnitRange::new(Unit::Weeks, Unit::Seconds)
                .unwrap_or_else(|error| panic!("Duration COW range: {error}")),
        );
        let result = package
            .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("shared Duration edit open failed: {error}"))
            .set(requested)
            .commit();
        match result {
            Ok(commit) => {
                let target_bytes = package_bytes(commit.package());
                let target_facts = fixture::format_entry_facts(&target_bytes)
                    .unwrap_or_else(|error| panic!("target Duration format facts: {error}"));
                assert!(
                    commit.patch().is_noop() || target_facts.len() >= source_facts.len(),
                    "a changed shared Duration format must retain its source entry"
                );
                assert_eq!(
                    target_facts.iter().map(|(_, count)| *count).sum::<u32>(),
                    source_facts.iter().map(|(_, count)| *count).sum::<u32>(),
                    "Duration COW changed total format references"
                );
                if !commit.patch().is_noop() {
                    let target_keys = fixture::format_keys(&target_bytes)
                        .unwrap_or_else(|error| panic!("target Duration format keys: {error}"));
                    assert_ne!(target_keys[0], target_keys[1]);
                    assert_eq!(target_keys[1], Some(FIRST_FORMAT_KEY));
                    let new_key = target_keys[0]
                        .unwrap_or_else(|| panic!("changed Duration COW key is missing"));
                    assert_eq!(target_facts, vec![(FIRST_FORMAT_KEY, 1), (new_key, 1)]);
                }
                assert_eq!(commit.patch().before(), before.as_ref());
                assert_eq!(
                    commit
                        .package()
                        .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
                        .unwrap_or_else(|error| panic!(
                            "Duration COW candidate read failed: {error}"
                        )),
                    Some(requested)
                );
                assert_eq!(
                    commit
                        .package()
                        .table_cell_duration_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
                        .unwrap_or_else(|error| panic!(
                            "Duration COW sibling read failed: {error}"
                        )),
                    sibling_before
                );
                let restored = commit
                    .package()
                    .apply_table_cell_duration_format(&commit.patch().inverse())
                    .unwrap_or_else(|error| panic!("Duration COW inverse failed: {error}"));
                assert_eq!(package_bytes(restored.package()), source_bytes);
                assert_eq!(package_bytes(package), source_bytes);
            },
            Err(error) => {
                observe_error(error);
                assert_eq!(package_bytes(package), source_bytes);
            },
        }
    });
    black_box(duration_from_bytes(data));
}

fn exercise_secondary_number(data: &[u8]) {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = duration_secondary_source();
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("Duration secondary source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let before = package
        .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration secondary source read failed: {error}"));
    let requested = duration_from_bytes(data);
    let result = package
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .and_then(|edit| edit.set(requested).commit());
    match result {
        Ok(commit) => {
            let target_bytes = package_bytes(commit.package());
            let source_facts = fixture::format_entry_facts(&source_bytes)
                .unwrap_or_else(|error| panic!("Duration secondary source facts: {error}"));
            let target_facts = fixture::format_entry_facts(&target_bytes)
                .unwrap_or_else(|error| panic!("Duration secondary target facts: {error}"));
            assert_secondary_number_refcount(&source_facts);
            assert_secondary_number_refcount(&target_facts);
            assert_eq!(commit.patch().before(), before.as_ref());
            assert_eq!(
                commit
                    .package()
                    .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
                    .unwrap_or_else(|error| panic!(
                        "Duration secondary candidate read failed: {error}"
                    )),
                Some(requested)
            );
            let changed_cell = fixture::tile_cells(&target_bytes)
                .unwrap_or_else(|error| panic!("Duration secondary target cells: {error}"))
                .into_iter()
                .next()
                .and_then(|bytes| BncCell::parse(&bytes).ok())
                .unwrap_or_else(|| panic!("Duration secondary target cell is missing"));
            assert_eq!(
                changed_cell.secondary_format_identifier(),
                Some(SECOND_FORMAT_KEY)
            );
            if !commit.patch().is_noop() {
                let restored = commit
                    .package()
                    .apply_table_cell_duration_format(&commit.patch().inverse())
                    .unwrap_or_else(|error| panic!("Duration secondary inverse failed: {error}"));
                assert_eq!(package_bytes(restored.package()), source_bytes);
            }
            let cleared = commit
                .package()
                .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
                .unwrap_or_else(|error| panic!("Duration secondary clear open failed: {error}"))
                .clear()
                .commit()
                .unwrap_or_else(|error| panic!("Duration secondary clear failed: {error}"));
            let cleared_bytes = package_bytes(cleared.package());
            let cleared_cell = fixture::tile_cells(&cleared_bytes)
                .unwrap_or_else(|error| panic!("Duration secondary cleared cells: {error}"))
                .into_iter()
                .next()
                .and_then(|bytes| BncCell::parse(&bytes).ok())
                .unwrap_or_else(|| panic!("Duration secondary cleared cell is missing"));
            assert_eq!(cleared_cell.secondary_format_identifier(), None);
            assert_eq!(
                fixture::format_keys(&cleared_bytes)
                    .unwrap_or_else(|error| panic!("Duration secondary cleared keys: {error}")),
                vec![None, Some(FIRST_FORMAT_KEY)]
            );
            assert_eq!(
                fixture::format_entry_facts(&cleared_bytes)
                    .unwrap_or_else(|error| panic!("Duration secondary cleared facts: {error}")),
                vec![(FIRST_FORMAT_KEY, 1), (SECOND_FORMAT_KEY, 1)]
            );
            assert_eq!(package_bytes(package), source_bytes);
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
        },
    }
}

fn assert_secondary_number_refcount(facts: &[(u32, u32)]) {
    assert_eq!(
        facts
            .iter()
            .filter(|(key, _)| *key == SECOND_FORMAT_KEY)
            .count(),
        1,
        "Duration secondary Number entry was culled or duplicated"
    );
    assert_eq!(
        facts
            .iter()
            .find(|(key, _)| *key == SECOND_FORMAT_KEY)
            .map(|(_, count)| *count),
        Some(2),
        "Duration secondary Number refcount changed"
    );
}

fn duration_secondary_source() -> Vec<u8> {
    let source = fixture::synthetic_package_for(
        fixture::FormatFamily::Duration,
        fixture::FormatSharing::Shared,
    )
    .unwrap_or_else(|error| panic!("Duration secondary fixture must build: {error}"));
    let source = fixture::rewrite_tile_cells(&source, |cells| {
        for cell_bytes in cells.iter_mut() {
            let mut encoded = BncCell::parse(cell_bytes)?.encode();
            let flags = u32::from_le_bytes(
                encoded
                    .get(8..12)
                    .ok_or_else(|| io::Error::other("Duration flags are truncated"))?
                    .try_into()
                    .map_err(|_| io::Error::other("Duration flags have an invalid width"))?,
            );
            let marker = u16::from_le_bytes(
                encoded
                    .get(6..8)
                    .ok_or_else(|| io::Error::other("Duration marker is truncated"))?
                    .try_into()
                    .map_err(|_| io::Error::other("Duration marker has an invalid width"))?,
            );
            assert_eq!(marker, EXPLICIT_DURATION_FORMAT);
            // A primary-only Duration cell stores kind at 20..24 followed by
            // its primary key. Insert the generic Number edge immediately
            // before that key, preserving the scalar and all other fields.
            encoded[8..12].copy_from_slice(&(flags | 0x0000_2000).to_le_bytes());
            encoded[6..8].copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());
            encoded.splice(24..24, SECOND_FORMAT_KEY.to_le_bytes());
            *cell_bytes = encoded;
        }
        Ok(())
    })
    .unwrap_or_else(|error| panic!("Duration secondary cell fixture: {error}"));
    fixture::rewrite_format_list_payload_for_test(&source, |list| {
        let primary = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("Duration primary entry is missing"))?;
        // The sibling still points at the primary Duration entry, so adding
        // one secondary edge must not decrement the primary refcount.
        primary.refcount = 2;
        list.entries.push(tst::table_data_list::ListEntry {
            key: SECOND_FORMAT_KEY,
            refcount: 2,
            format: Some(tsk::FormatStructArchive {
                format_type: Some(fixture::NATIVE_NUMBER_FORMAT_TYPE),
                decimal_places: Some(2),
                negative_style: Some(0),
                show_thousands_separator: Some(true),
                ..Default::default()
            }),
            ..Default::default()
        });
        Ok(())
    })
    .unwrap_or_else(|error| panic!("Duration secondary format list fixture: {error}"))
}

fn exercise_inherited_marker(data: &[u8]) {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = fixture::rewrite_tile_cells(&duration_source(), |cells| {
            let first = cells
                .first_mut()
                .ok_or_else(|| io::Error::other("Duration inherited cell is missing"))?;
            if first.len() < 8 {
                return Err(io::Error::other("Duration inherited cell is truncated").into());
            }
            // Keep the valid Duration kind and primary key while removing
            // only the explicit marker. This ambiguous native tuple must
            // fail closed; it is not the same as a genuinely absent cell.
            first[6..8].copy_from_slice(&0_u16.to_le_bytes());
            Ok(())
        })
        .unwrap_or_else(|error| panic!("Duration inherited source must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("Duration inherited source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let read = package.table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION);
    assert!(read.is_err(), "ambiguous Duration marker was admitted");
    observe_result(read);
    let edit = package.edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION);
    assert!(edit.is_err(), "ambiguous Duration edit was admitted");
    observe_result(edit);
    assert_eq!(package_bytes(package), source_bytes);

    // True absence is represented by clearing a valid source, which removes
    // the marker, kind, and primary reference together. It remains an
    // ordinary Optional::None state and can be attached and cleared again.
    let absent = duration_package()
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration absent edit open failed: {error}"))
        .clear()
        .commit()
        .unwrap_or_else(|error| panic!("Duration absent clear failed: {error}"));
    assert_eq!(
        absent
            .package()
            .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Duration absent read failed: {error}")),
        None
    );
    let desired = duration_from_bytes(data);
    let attached = absent
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration absent attach open failed: {error}"))
        .set(desired)
        .commit()
        .unwrap_or_else(|error| panic!("Duration absent attach failed: {error}"));
    assert_eq!(
        attached
            .package()
            .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Duration absent attach read failed: {error}")),
        Some(desired)
    );
    let cleared = attached
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("Duration absent final clear open failed: {error}"))
        .clear()
        .commit()
        .unwrap_or_else(|error| panic!("Duration absent final clear failed: {error}"));
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Duration absent final read failed: {error}")),
        None
    );
    assert_eq!(
        package_bytes(absent.package()),
        package_bytes(cleared.package())
    );
}

fn exercise_malformed_sources() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = duration_source();
        let mut malformed = Vec::new();

        malformed.push((
            "duplicate format key",
            fixture::rewrite_format_list_payload_for_test(&source, |list| {
                let entry = list
                    .entries
                    .first()
                    .cloned()
                    .ok_or_else(|| io::Error::other("Duration format entry is missing"))?;
                list.entries.push(entry);
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration duplicate key source: {error}")),
        ));
        malformed.push((
            "missing format entry",
            fixture::rewrite_format_list_payload_for_test(&source, |list| {
                list.entries.retain(|entry| entry.key != FIRST_FORMAT_KEY);
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration missing key source: {error}")),
        ));
        malformed.push((
            "refcount mismatch",
            fixture::rewrite_format_list_payload_for_test(&source, |list| {
                let entry = list
                    .entries
                    .first_mut()
                    .ok_or_else(|| io::Error::other("Duration format entry is missing"))?;
                entry.refcount = 0;
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration refcount source: {error}")),
        ));
        malformed.push((
            "malformed format payload",
            fixture::rewrite_format_payload_by_key(&source, FIRST_FORMAT_KEY, &[0x80])
                .unwrap_or_else(|error| panic!("Duration malformed payload source: {error}")),
        ));
        malformed.push((
            "unsupported format type",
            fixture::rewrite_format_varint_by_key(&source, FIRST_FORMAT_KEY, 1, 65_535)
                .unwrap_or_else(|error| panic!("Duration unsupported type source: {error}")),
        ));
        malformed.push((
            "wrong cell family",
            fixture::rewrite_tile_cells(&source, |cells| {
                let first = cells
                    .first_mut()
                    .ok_or_else(|| io::Error::other("Duration first cell is missing"))?;
                let mut cell = BncCell::minimal();
                cell.set_number(7.0)?;
                cell.set_data_format_identifier(
                    FIRST_FORMAT_KEY,
                    CellDataFormatKind::NumberOrPercentage,
                    None,
                )?;
                *first = cell.encode();
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration wrong-family cell source: {error}")),
        ));

        let secondary = duration_secondary_source();
        malformed.push((
            "automatic Duration secondary tuple",
            duration_cell_source(&secondary, |bytes| {
                bytes[6..8].copy_from_slice(&0_u16.to_le_bytes());
            }),
        ));
        malformed.push((
            "wrong Duration kind",
            duration_cell_source(&secondary, |bytes| {
                bytes[20..24].copy_from_slice(&DATE_TIME_CELL_FORMAT_KIND.to_le_bytes());
            }),
        ));
        malformed.push((
            "wrong Duration marker",
            duration_cell_source(&secondary, |bytes| {
                bytes[6..8].copy_from_slice(&EXPLICIT_DECIMAL_FORMAT.to_le_bytes());
            }),
        ));
        malformed.push((
            "zero Duration identifier",
            duration_cell_source(&secondary, |bytes| {
                bytes[28..32].fill(0);
            }),
        ));
        malformed.push((
            "reserved BNC field",
            duration_cell_source(&secondary, |bytes| {
                let flags = u32::from_le_bytes(bytes[8..12].try_into().unwrap_or([0; 4]));
                bytes[8..12].copy_from_slice(&(flags | 0x0010_0000).to_le_bytes());
                bytes.extend_from_slice(&[0; 4]);
            }),
        ));
        malformed.push((
            "secondary marker without Number",
            fixture::rewrite_tile_cells(&secondary, |cells| {
                let first = cells
                    .first_mut()
                    .ok_or_else(|| io::Error::other("Duration secondary cell is missing"))?;
                let mut bytes = first.clone();
                bytes[6..8].copy_from_slice(&EXPLICIT_DURATION_FORMAT.to_le_bytes());
                *first = bytes;
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration secondary marker source: {error}")),
        ));
        malformed.push((
            "secondary wrong family",
            fixture::rewrite_format_varint_by_key(
                &secondary,
                SECOND_FORMAT_KEY,
                1,
                u64::from(fixture::NATIVE_PERCENTAGE_FORMAT_TYPE),
            )
            .unwrap_or_else(|error| panic!("Duration secondary wrong-family source: {error}")),
        ));
        malformed.push((
            "secondary refcount mismatch",
            fixture::rewrite_format_list_payload_for_test(&secondary, |list| {
                let entry = list
                    .entries
                    .iter_mut()
                    .find(|entry| entry.key == SECOND_FORMAT_KEY)
                    .ok_or_else(|| io::Error::other("Duration secondary entry is missing"))?;
                entry.refcount = 0;
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration secondary refcount source: {error}")),
        ));
        malformed.push((
            "formula cache family mismatch",
            fixture::rewrite_tile_cells(&source, |cells| {
                let first = cells
                    .first_mut()
                    .ok_or_else(|| io::Error::other("Duration first cell is missing"))?;
                let mut cell = BncCell::minimal();
                cell.set_duration(7.0)?;
                cell.set_data_format_identifier(
                    FIRST_FORMAT_KEY,
                    CellDataFormatKind::Duration,
                    None,
                )?;
                cell.set_formula_reference(11);
                cell.set_formula_cached_boolean(true)?;
                *first = cell.encode();
                Ok(())
            })
            .unwrap_or_else(|error| panic!("Duration formula mismatch source: {error}")),
        ));

        for (label, source) in malformed {
            let before = source.clone();
            match Package::from_bytes_with_options(&source, options()) {
                Ok(package) => {
                    let package_before = package_bytes(&package);
                    let read = package.table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION);
                    assert!(
                        read.is_err(),
                        "malformed {label} admitted Duration read: {read:?}"
                    );
                    let edit =
                        package.edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION);
                    assert!(
                        edit.is_err(),
                        "malformed {label} admitted Duration edit: {edit:?}"
                    );
                    observe_result(edit);
                    assert_eq!(package_bytes(&package), package_before);
                },
                Err(error) => observe_error(error),
            }
            assert_eq!(source, before, "malformed {label} builder mutated source");
        }
    });
}

fn duration_source() -> Vec<u8> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::Duration,
        fixture::FormatSharing::Shared,
    )
    .unwrap_or_else(|error| panic!("Duration source must build: {error}"))
}

fn duration_cell_source(source: &[u8], mutate: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
    fixture::rewrite_tile_cells(source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration first cell is missing"))?;
        mutate(first);
        Ok(())
    })
    .unwrap_or_else(|error| panic!("Duration malformed cell source: {error}"))
}

fn exercise_locked_source(data: &[u8]) {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = duration_source();
        let locked = fixture::locked_table_package(&source)
            .unwrap_or_else(|error| panic!("locked Duration source must build: {error}"));
        Package::from_bytes_with_options(&locked, options())
            .unwrap_or_else(|error| panic!("locked Duration source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let result = package
        .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
        .and_then(|edit| edit.set(duration_from_bytes(data)).commit());
    assert!(result.is_err(), "changed Duration edit bypassed table lock");
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_constructor_boundaries(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        assert!(matches!(
            UnitRange::new(Unit::Seconds, Unit::Hours),
            Err(litchi::numbers::cell::data_format::duration::Error::ReversedRange { .. })
        ));
        for style in ALL_STYLES {
            for (largest_index, largest) in ALL_UNITS.iter().copied().enumerate() {
                for smallest in ALL_UNITS[largest_index..].iter().copied() {
                    let range = UnitRange::new(largest, smallest)
                        .unwrap_or_else(|error| panic!("valid Duration range rejected: {error}"));
                    black_box((
                        Duration::new(style, Units::Automatic(range)),
                        Duration::custom(style, range),
                    ));
                }
            }
        }
        assert_eq!(UnitRange::all().largest(), Unit::Weeks);
        assert_eq!(
            UnitRange::hours_to_milliseconds().smallest(),
            Unit::Milliseconds
        );
    });
    black_box(duration_from_bytes(data));
}

fn duration_from_bytes(data: &[u8]) -> Duration {
    let style = ALL_STYLES[usize::from(control(data, 1)) % ALL_STYLES.len()];
    let largest_index = usize::from(control(data, 2)) % ALL_UNITS.len();
    let largest = ALL_UNITS[largest_index];
    let smallest_index =
        largest_index + (usize::from(control(data, 3)) % (ALL_UNITS.len() - largest_index));
    let smallest = ALL_UNITS[smallest_index];
    let range = UnitRange::new(largest, smallest)
        .unwrap_or_else(|error| panic!("bounded Duration range must validate: {error}"));
    let units = if control(data, 4) & 1 == 0 {
        Units::Automatic(range)
    } else {
        Units::Custom(range)
    };
    Duration::new(style, units)
}

fn random_position(data: &[u8]) -> CellPosition {
    CellPosition::new(
        u32::from(control(data, 5) % 8),
        u32::from(control(data, 6) % 8),
    )
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
                        unreachable!("Duration sentinel is UTF-8: {error}")
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
            Err(error) => observe_error(error),
            Ok(_) => panic!("oversized Duration input must be rejected"),
        }
    });
}

fn exercise_semantic_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let semantic = PackageSemanticLimits::new(1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("one-entry Duration limits: {error}"));
        let constrained = PackageReadOptions::new(options().archive(), semantic);
        let result = Package::from_bytes_with_options(&duration_source(), constrained);
        assert!(result.is_err(), "one-entry limits admitted Duration source");
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn exercise_output_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = duration_source();
        let limits = PackageLimits::new(
            u64::try_from(source.len()).unwrap_or(u64::MAX),
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("Duration output limits: {error}"));
        let package = Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(limits, PackageSemanticLimits::default()),
        )
        .unwrap_or_else(|error| panic!("Duration exact-limit source must open: {error}"));
        let before = package_bytes(&package);
        let requested = Duration::custom(
            Style::FullNames,
            UnitRange::new(Unit::Weeks, Unit::Milliseconds)
                .unwrap_or_else(|error| panic!("Duration output range: {error}")),
        );
        let result = package
            .edit_table_cell_duration_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Duration output-limit edit open failed: {error}"))
            .set(requested)
            .commit();
        assert!(
            matches!(
                result.as_ref(),
                Err(litchi::numbers::cell::data_format::duration::transaction::Error::LimitExceeded { .. })
            ),
            "exact Duration output limit did not produce a bounded refusal: {result:?}"
        );
        black_box(result.err().map(|error| error.to_string()));
        assert_eq!(package_bytes(&package), before);
    });
}

fn assert_exact_locality(source: &[u8], target: &[u8]) {
    let before = Catalog::from_bytes(source)
        .unwrap_or_else(|error| panic!("Duration source catalog: {error}"));
    let after = Catalog::from_bytes(target)
        .unwrap_or_else(|error| panic!("Duration target catalog: {error}"));
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .unwrap_or_else(|| panic!("Duration candidate removed {}", entry.name()));
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged Duration member {} lost its exact local ZIP record",
                entry.name()
            );
        }
        if entry.name().contains("Metadata")
            || fixture::PREVIEW_MEMBERS.contains(&entry.name())
            || entry.name() == fixture::UNRELATED_MEMBER
            || entry.name() == fixture::SENTINEL_MEMBER
        {
            assert_eq!(
                entry.data(),
                candidate.data(),
                "unrelated Duration member changed"
            );
        }
    }
    assert_eq!(changed, [fixture::TABLES_MEMBER.to_owned()]);
    assert_eq!(before.len(), after.len());
}

fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) {
    let before = fixture::tile_cells(source)
        .unwrap_or_else(|error| panic!("Duration source tile cells: {error}"));
    let after = fixture::tile_cells(target)
        .unwrap_or_else(|error| panic!("Duration target tile cells: {error}"));
    assert_eq!(before.len(), after.len());
    for (left, right) in before.iter().zip(after.iter()) {
        let left_cell =
            BncCell::parse(left).unwrap_or_else(|error| panic!("Duration source BNC: {error}"));
        let right_cell =
            BncCell::parse(right).unwrap_or_else(|error| panic!("Duration target BNC: {error}"));
        assert_eq!(
            left_cell
                .cached_scalar()
                .unwrap_or_else(|error| panic!("Duration source scalar: {error}")),
            right_cell
                .cached_scalar()
                .unwrap_or_else(|error| panic!("Duration target scalar: {error}")),
            "Duration edit changed the cached scalar"
        );
        assert_eq!(normalized_cell(left), normalized_cell(right));
    }
}

fn normalized_cell(cell: &[u8]) -> Vec<u8> {
    let mut parsed =
        BncCell::parse(cell).unwrap_or_else(|error| panic!("Duration BNC normalization: {error}"));
    parsed.clear_explicit_format();
    parsed.encode()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Duration package failed: {error}"));
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
        "Duration error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Duration error debug leaked private input"
    );
    black_box((display, debug));
}
