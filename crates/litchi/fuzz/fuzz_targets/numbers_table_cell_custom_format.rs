#![no_main]

//! Bounded selector-first fuzzing for Numbers table-cell Custom formats.
//!
//! The target gives arbitrary bytes to bounded Numbers ingress and separately
//! replays a small command prefix against deterministic Custom Number, Text,
//! and Date & Time packages.  The checked-in corpus is made of command bytes,
//! not native package data.  Native registry objects and UUIDs stay in the
//! test-only fixture; this target crosses the package boundary only with
//! archive-free semantic values and selectors.

use std::{
    fmt::{Debug, Display},
    hint::black_box,
    sync::OnceLock,
};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::custom::{
        Condition, ConditionValue, Custom, DateTime as CustomDateTime, DateTimePattern,
        MAX_NAME_BYTES, MAX_PATTERN_BYTES, MAX_RULES, Name, Number as CustomNumber, NumberPattern,
        NumberRule, Text as CustomText,
    },
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
const PRIVATE_SHEET: &str = "__litchi_private_custom_format_sheet_9e31__";
const PRIVATE_TABLE: &str = "__litchi_private_custom_format_table_9e31__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_custom_format_input_9e31__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

const FIXTURE_POSITION: CellPosition = CellPosition::new(0, 0);
const FIXTURE_SIBLING_POSITION: CellPosition = CellPosition::new(0, 1);
const NATIVE_POSITION: CellPosition = CellPosition::new(2, 1);

fuzz_target!(|data: &[u8]| {
    // Arbitrary ingress remains useful for package-level parser coverage, but
    // all semantic command construction below is independently bounded.
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_arbitrary_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Replay every lifecycle command against valid source packages.  This
    // keeps Custom reads, edits, exact apply, inverse, and candidate reopen
    // reachable even when fuzzed ZIP/CRC input is rejected at ingress.
    exercise_custom_package(custom_number_package(), fixture::CustomFamily::Number, data);
    exercise_custom_package(custom_text_package(), fixture::CustomFamily::Text, data);
    exercise_custom_package(
        custom_date_time_package(),
        fixture::CustomFamily::DateTime,
        data,
    );

    exercise_wrong_families(data);
    exercise_shared_registry_cow(data);
    exercise_foreign_patch_conflicts(data);
    exercise_locked_source(data);
    exercise_malformed_registries(data);
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
        .unwrap_or_else(|error| unreachable!("valid Custom archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Custom semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid Custom projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn custom_number_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_custom(fixture::CustomFamily::Number))
}

fn custom_text_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_custom(fixture::CustomFamily::Text))
}

fn custom_date_time_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| package_for_custom(fixture::CustomFamily::DateTime))
}

fn package_for_custom(family: fixture::CustomFamily) -> Package {
    let source = fixture::custom_package(family)
        .unwrap_or_else(|error| panic!("deterministic {family:?} Custom source: {error}"));
    Package::from_bytes_with_options(&source, options())
        .unwrap_or_else(|error| panic!("deterministic {family:?} Custom source must open: {error}"))
}

fn custom_unshared_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::custom_package_for(
            fixture::CustomFamily::Number,
            fixture::FormatSharing::Unshared,
        )
        .unwrap_or_else(|error| panic!("unshared Custom source: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("unshared Custom source must open: {error}"))
    })
}

fn exercise_arbitrary_package(package: &Package, data: &[u8]) {
    let position = random_position(data);
    exercise_selector_reads(package, data, position);
    let source_bytes = package_bytes(package);
    let before = match package.table_cell_custom_format(0usize, 0usize, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            observe_result(package.edit_table_cell_custom_format(0usize, 0usize, position));
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let edit = match package.edit_table_cell_custom_format(0usize, 0usize, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let result = edit.set(requested_custom(data)).commit();
    if let Err(error) = result {
        observe_error(error);
    }
    // A failed or unsupported arbitrary operation must never publish through
    // the source handle.  Deterministic sources below perform full replay.
    assert_eq!(package_bytes(package), source_bytes);
    black_box(before);
}

fn exercise_custom_package(package: &Package, family: fixture::CustomFamily, data: &[u8]) {
    exercise_selector_reads(package, data, FIXTURE_POSITION);
    exercise_selector_errors(package, FIXTURE_POSITION);

    let source_bytes = package_bytes(package);
    let before = package
        .table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("{family:?} Custom source read failed: {error}"));
    let requested = requested_custom_for(family, data);
    let edit = package
        .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("{family:?} Custom source edit failed: {error}"));

    // Four operations deliberately retain the source's Optional state for a
    // true no-op, then cover explicit set, clear, and reset.  The fifth class
    // is another set to make replacement common in a small smoke run.
    let command = control(data, 0) % 5;
    let (edit, expected_after) = match command {
        0 => match before.clone() {
            Some(value) => (edit.set(value), before.clone()),
            None => (edit.clear(), None),
        },
        1 | 4 => (edit.set(requested.clone()), Some(requested)),
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
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
    } else {
        // Replacing a Custom value updates both the table-local format list
        // and the document-scoped registry. Clearing one of two shared cells
        // leaves the registry entry live and therefore touches only the table
        // component.
        assert_eq!(
            commit.diagnostics().touched_components(),
            if expected_after.is_some() { 2 } else { 1 }
        );
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
    }
    assert_eq!(
        commit
            .package()
            .table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("{family:?} Custom candidate read failed: {error}")),
        expected_after
    );

    let applied = package
        .apply_table_cell_custom_format(&patch)
        .unwrap_or_else(|error| panic!("fresh {family:?} Custom patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("applied {family:?} Custom read failed: {error}")),
        expected_after
    );
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_custom_format(&patch)
                .is_err(),
            "a changed Custom patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_custom_format(&patch.inverse())
                .is_err(),
            "a Custom inverse must reject the original source"
        );
    }
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_custom_format(&inverse)
        .unwrap_or_else(|error| panic!("{family:?} Custom inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("{family:?} Custom inverse read failed: {error}")),
        before
    );
    assert_eq!(
        restored
            .package()
            .table_cell_custom_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!(
                "{family:?} Custom sibling inverse read failed: {error}"
            )),
        package
            .table_cell_custom_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("{family:?} Custom sibling read failed: {error}"))
    );
    black_box((
        patch.before(),
        patch.after(),
        patch.path(),
        commit.diagnostics(),
    ));
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8], position: CellPosition) {
    let random = random_position(data);
    observe_result(package.table_cell_custom_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        random,
    ));
    observe_result(package.table_cell_custom_format(0usize, 0usize, position));
    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        let by_name = package.table_cell_custom_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        );
        let by_index = package.table_cell_custom_format(0usize, 0usize, position);
        assert_eq!(by_name, by_index, "Custom selector forms disagree");
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let sheet_count = package.document().sheets().len();
    observe_result(package.table_cell_custom_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ));
    if let Err(error) = package.table_cell_custom_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_custom_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    observe_result(package.table_cell_custom_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(u32::MAX, u32::MAX),
    ));
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_wrong_families(data: &[u8]) {
    for (family, package) in wrong_family_packages() {
        black_box(family);
        assert_custom_refusal(package, data, FIXTURE_POSITION);
    }
    let package = Package::from_bytes_with_options(NATIVE_NUMBERS, options())
        .unwrap_or_else(|error| panic!("native Numbers Custom source must open: {error}"));
    assert_custom_refusal(&package, data, NATIVE_POSITION);
}

fn wrong_family_packages() -> &'static [(fixture::FormatFamily, Package)] {
    static PACKAGES: OnceLock<Vec<(fixture::FormatFamily, Package)>> = OnceLock::new();
    PACKAGES
        .get_or_init(|| {
            [
                fixture::FormatFamily::Number,
                fixture::FormatFamily::Percentage,
                fixture::FormatFamily::Currency,
                fixture::FormatFamily::Scientific,
                fixture::FormatFamily::Fraction,
                fixture::FormatFamily::DateTime,
                fixture::FormatFamily::Text,
            ]
            .into_iter()
            .map(|family| {
                let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)
                    .unwrap_or_else(|error| panic!("{family:?} wrong-family source: {error}"));
                let package =
                    Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
                        panic!("{family:?} wrong-family source must open: {error}")
                    });
                (family, package)
            })
            .collect()
        })
        .as_slice()
}

fn assert_custom_refusal(package: &Package, data: &[u8], position: CellPosition) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_custom_format(0usize, 0usize, position);
    assert!(
        read.is_err(),
        "non-Custom cell was accepted by Custom read: {read:?}"
    );
    if let Err(error) = read {
        observe_error(error);
    }
    let result = package
        .edit_table_cell_custom_format(0usize, 0usize, position)
        .and_then(|edit| edit.set(requested_custom(data)).commit());
    assert!(
        result.is_err(),
        "non-Custom cell was accepted by Custom edit"
    );
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_shared_registry_cow(data: &[u8]) {
    let package = custom_unshared_package();
    let source_bytes = package_bytes(package);
    let before = package
        .table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("unshared Custom read failed: {error}"));
    let facts = fixture::format_entry_facts(&source_bytes)
        .unwrap_or_else(|error| panic!("Custom format refcounts must decode: {error}"));
    assert_eq!(
        facts,
        vec![
            (fixture::FIRST_FORMAT_KEY, 1),
            (fixture::SECOND_FORMAT_KEY, 1)
        ]
    );
    let registry = fixture::custom_registry_facts(&source_bytes)
        .unwrap_or_else(|error| panic!("Custom registry must decode: {error}"));
    assert!(
        registry.len() >= 2,
        "unshared Custom registry lost its second UUID"
    );

    let result = package
        .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
        .unwrap_or_else(|error| panic!("unshared Custom edit failed: {error}"))
        .set(requested_custom_for(fixture::CustomFamily::Number, data))
        .commit();
    match result {
        Ok(commit) => {
            assert_ne!(package_bytes(commit.package()), source_bytes);
            assert_eq!(commit.patch().before(), before.as_ref());
            assert_eq!(
                commit
                    .package()
                    .table_cell_custom_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
                    .unwrap_or_else(|error| panic!("Custom sibling COW read failed: {error}")),
                package
                    .table_cell_custom_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
                    .unwrap_or_else(|error| panic!("Custom sibling read failed: {error}"))
            );
            let target_facts = fixture::format_entry_facts(&package_bytes(commit.package()))
                .unwrap_or_else(|error| panic!("Custom target refcounts must decode: {error}"));
            assert_eq!(target_facts.len(), 2);
            assert!(
                !target_facts
                    .iter()
                    .any(|(key, _)| *key == fixture::FIRST_FORMAT_KEY),
                "the unshared replaced format entry was not culled"
            );
            assert!(
                target_facts
                    .iter()
                    .any(|(key, count)| { *key == fixture::SECOND_FORMAT_KEY && *count == 1 }),
                "the unselected sibling format entry was not retained"
            );
            assert!(
                target_facts.iter().any(|(key, count)| {
                    *key != fixture::FIRST_FORMAT_KEY
                        && *key != fixture::SECOND_FORMAT_KEY
                        && *count == 1
                }),
                "the replacement format entry was not installed"
            );
            black_box(commit.diagnostics());
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_foreign_patch_conflicts(_data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = fixture::custom_package(fixture::CustomFamily::Number)
            .unwrap_or_else(|error| panic!("Custom conflict source: {error}"));
        let package = Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("Custom conflict source must open: {error}"));
        let commit = package
            .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Custom conflict edit: {error}"))
            .set(
                CustomNumber::new(
                    Name::new("Conflict Replacement")
                        .unwrap_or_else(|error| panic!("Custom conflict name: {error}")),
                    NumberPattern::new("#,##0.000")
                        .unwrap_or_else(|error| panic!("Custom conflict pattern: {error}")),
                )
                .into(),
            )
            .commit()
            .unwrap_or_else(|error| panic!("Custom conflict patch must commit: {error}"));

        let target = package_bytes(commit.package());
        let stale = Package::from_bytes_with_options(&target, options())
            .unwrap_or_else(|error| panic!("Custom stale target must open: {error}"));
        let stale_before = package_bytes(&stale);
        assert!(
            stale
                .apply_table_cell_custom_format(commit.patch())
                .is_err()
        );
        assert_eq!(package_bytes(&stale), stale_before);

        let foreign_source = fixture::custom_package_for(
            fixture::CustomFamily::Number,
            fixture::FormatSharing::Unshared,
        )
        .unwrap_or_else(|error| panic!("Custom foreign source: {error}"));
        let foreign = Package::from_bytes_with_options(&foreign_source, options())
            .unwrap_or_else(|error| panic!("Custom foreign source must open: {error}"));
        let foreign_before = package_bytes(&foreign);
        assert!(
            foreign
                .apply_table_cell_custom_format(commit.patch())
                .is_err()
        );
        assert_eq!(package_bytes(&foreign), foreign_before);
    });
}

fn exercise_locked_source(data: &[u8]) {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    let package = PACKAGE.get_or_init(|| {
        let source = fixture::custom_package(fixture::CustomFamily::Number)
            .unwrap_or_else(|error| panic!("locked Custom source: {error}"));
        let locked = fixture::locked_table_package(&source)
            .unwrap_or_else(|error| panic!("locked Custom source: {error}"));
        Package::from_bytes_with_options(&locked, options())
            .unwrap_or_else(|error| panic!("locked Custom source must open: {error}"))
    });
    let source_bytes = package_bytes(package);
    let result = package
        .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
        .and_then(|edit| {
            edit.set(requested_custom_for(fixture::CustomFamily::Number, data))
                .commit()
        });
    assert!(
        result.is_err(),
        "locked Custom table unexpectedly committed"
    );
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_malformed_registries(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for family in [
            fixture::CustomFamily::Number,
            fixture::CustomFamily::Text,
            fixture::CustomFamily::DateTime,
        ] {
            for corruption in [
                fixture::CustomCorruption::StaleUuid,
                fixture::CustomCorruption::ForeignUuid,
                fixture::CustomCorruption::DuplicateUuid,
                fixture::CustomCorruption::ZeroUuid,
                fixture::CustomCorruption::MissingRegistry,
                fixture::CustomCorruption::DetachedRegistry,
                fixture::CustomCorruption::Field8CustomEntry,
                fixture::CustomCorruption::DeprecatedTableRegistry,
                fixture::CustomCorruption::DuplicateFormatEntry,
                fixture::CustomCorruption::MissingFormatEntry,
                fixture::CustomCorruption::RefcountMismatch,
                fixture::CustomCorruption::MalformedRegistryPayload,
            ] {
                let source = fixture::corrupted_custom_package(family, corruption)
                    .unwrap_or_else(|error| panic!("{family:?} {corruption:?} source: {error}"));
                match Package::from_bytes_with_options(&source, options()) {
                    Err(error) => observe_error(error),
                    Ok(package) => {
                        let source_bytes = package_bytes(&package);
                        let result = package
                            .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
                            .and_then(|edit| edit.set(requested_custom_for(family, data)).commit());
                        assert!(
                            result.is_err(),
                            "malformed {family:?} registry {corruption:?} committed"
                        );
                        if let Err(error) = result {
                            observe_error(error);
                        }
                        assert_eq!(package_bytes(&package), source_bytes);
                    },
                }
            }
        }
    });
}

fn requested_custom(data: &[u8]) -> Custom {
    match control(data, 1) % 3 {
        0 => requested_custom_for(fixture::CustomFamily::Number, data),
        1 => requested_custom_for(fixture::CustomFamily::Text, data),
        _ => requested_custom_for(fixture::CustomFamily::DateTime, data),
    }
}

fn requested_custom_for(family: fixture::CustomFamily, data: &[u8]) -> Custom {
    let name = requested_name(family, data);
    match family {
        fixture::CustomFamily::Number => {
            let default = if control(data, 2) & 1 == 0 {
                "#,##0.00"
            } else {
                "0.00;[Red]-0.00"
            };
            let default = NumberPattern::new(default)
                .unwrap_or_else(|error| unreachable!("bounded Number pattern: {error}"));
            let count = usize::from(control(data, 3) % 4);
            let mut rules = Vec::with_capacity(count);
            for index in 0..count {
                let threshold = ConditionValue::try_new(
                    f64::from(control(data, 4 + index)) - 64.0 + index as f64 / 10.0,
                )
                .unwrap_or_else(|error| unreachable!("bounded threshold: {error}"));
                let condition = match index % 5 {
                    0 => Condition::EqualTo(threshold),
                    1 => Condition::LessThan(threshold),
                    2 => Condition::LessThanOrEqualTo(threshold),
                    3 => Condition::GreaterThan(threshold),
                    _ => Condition::GreaterThanOrEqualTo(threshold),
                };
                let pattern = NumberPattern::new(if index & 1 == 0 {
                    "(#,##0.00)"
                } else {
                    ">#,##0.00"
                })
                .unwrap_or_else(|error| unreachable!("bounded rule pattern: {error}"));
                rules.push(NumberRule::new(condition, pattern));
            }
            CustomNumber::try_with_rules(name, default, rules)
                .unwrap_or_else(|error| unreachable!("bounded Custom Number: {error}"))
                .into()
        },
        fixture::CustomFamily::Text => {
            let prefix = if control(data, 2) & 1 == 0 {
                "ID: "
            } else {
                "["
            };
            let suffix = if control(data, 3) & 1 == 0 { " !" } else { "]" };
            if control(data, 4) & 1 == 0 {
                CustomText::try_new(name, prefix, suffix)
                    .unwrap_or_else(|error| unreachable!("bounded Custom Text: {error}"))
                    .into()
            } else {
                let literal = bounded_ascii(data, 5, 32);
                let literal = if literal.is_empty() {
                    "literal".to_owned()
                } else {
                    literal
                };
                CustomText::try_literal(name, literal)
                    .unwrap_or_else(|error| unreachable!("bounded Custom literal: {error}"))
                    .into()
            }
        },
        fixture::CustomFamily::DateTime => {
            let pattern = match control(data, 2) % 3 {
                0 => "yyyy-MM-dd",
                1 => "EEEE, MMMM d, y",
                _ => "yyyy-MM-dd HH:mm:ss",
            };
            CustomDateTime::new(
                name,
                DateTimePattern::new(pattern)
                    .unwrap_or_else(|error| unreachable!("bounded Custom Date & Time: {error}")),
            )
            .into()
        },
    }
}

fn requested_name(family: fixture::CustomFamily, data: &[u8]) -> Name {
    let family = match family {
        fixture::CustomFamily::Number => "Number",
        fixture::CustomFamily::Text => "Text",
        fixture::CustomFamily::DateTime => "DateTime",
    };
    Name::new(&format!("Fuzz {family} {}", control(data, 6)))
        .unwrap_or_else(|error| unreachable!("bounded Custom name: {error}"))
}

fn exercise_constructor_boundaries(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        assert!(matches!(Name::new(""), Err(_)));
        assert!(matches!(Name::new(" name"), Err(_)));
        assert!(matches!(Name::new("name "), Err(_)));
        assert!(matches!(Name::new("name\0"), Err(_)));
        assert!(Name::new(&"N".repeat(MAX_NAME_BYTES)).is_ok());
        assert!(Name::new(&"N".repeat(MAX_NAME_BYTES + 1)).is_err());

        assert!(NumberPattern::new("literal").is_err());
        assert!(NumberPattern::new("0\n").is_err());
        assert!(NumberPattern::new(&format!("0{}", "x".repeat(MAX_PATTERN_BYTES - 1))).is_ok());
        assert!(NumberPattern::new(&format!("0{}", "x".repeat(MAX_PATTERN_BYTES))).is_err());

        assert!(DateTimePattern::new("").is_err());
        assert!(DateTimePattern::new("qqq").is_err());
        assert!(DateTimePattern::new("yyyy\tMM").is_err());
        assert!(DateTimePattern::new(&format!("y{}", "x".repeat(MAX_PATTERN_BYTES - 1))).is_ok());
        assert!(DateTimePattern::new(&format!("y{}", "x".repeat(MAX_PATTERN_BYTES))).is_err());

        assert!(ConditionValue::try_new(f64::NAN).is_err());
        assert!(ConditionValue::try_new(f64::INFINITY).is_err());
        assert!(ConditionValue::try_new(f64::NEG_INFINITY).is_err());
        assert_eq!(
            ConditionValue::try_new(-0.0).unwrap().value().to_bits(),
            0.0f64.to_bits()
        );

        let name = Name::new("Boundary").unwrap();
        let default = NumberPattern::new("0").unwrap();
        let mut rules = Vec::with_capacity(MAX_RULES);
        for index in 0..MAX_RULES {
            let threshold = ConditionValue::try_new(index as f64).unwrap();
            rules.push(NumberRule::new(
                Condition::EqualTo(threshold),
                NumberPattern::new("0.0").unwrap(),
            ));
        }
        assert!(CustomNumber::try_with_rules(name.clone(), default.clone(), rules.clone()).is_ok());
        rules.push(NumberRule::new(
            Condition::EqualTo(ConditionValue::try_new(999.0).unwrap()),
            NumberPattern::new("0").unwrap(),
        ));
        assert!(CustomNumber::try_with_rules(name.clone(), default.clone(), rules).is_err());
        let duplicate = NumberRule::new(
            Condition::EqualTo(ConditionValue::try_new(1.0).unwrap()),
            NumberPattern::new("0").unwrap(),
        );
        assert!(
            CustomNumber::try_with_rules(name, default, [duplicate.clone(), duplicate]).is_err()
        );

        assert!(CustomText::try_new(Name::new("Affix").unwrap(), "x", "y").is_ok());
        assert!(CustomText::try_new(Name::new("Affix").unwrap(), "x\0", "").is_err());
        assert!(
            CustomText::try_new(
                Name::new("Affix").unwrap(),
                "x".repeat(MAX_PATTERN_BYTES - 1),
                "",
            )
            .is_ok()
        );
        assert!(
            CustomText::try_new(
                Name::new("Affix").unwrap(),
                "x".repeat(MAX_PATTERN_BYTES),
                "",
            )
            .is_err()
        );
        assert!(CustomText::try_literal(Name::new("Literal").unwrap(), "").is_err());
    });
    black_box(requested_custom(data));
}

fn exercise_redacted_ingress() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        for input in [PRIVATE_INPUT, b"not-a-numbers-package", b"".as_slice()] {
            if let Err(error) = Package::from_bytes_with_options(input, options()) {
                if input == b"" {
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
            Ok(_) => panic!("oversized Custom input must be rejected"),
        }
    });
}

fn exercise_semantic_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let semantic = PackageSemanticLimits::new(1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("one-entry Custom limits: {error}"));
        let constrained = PackageReadOptions::new(options().archive(), semantic);
        let source = fixture::custom_package(fixture::CustomFamily::Number)
            .unwrap_or_else(|error| panic!("Custom semantic-limit source: {error}"));
        let result = Package::from_bytes_with_options(&source, constrained);
        assert!(result.is_err(), "one-entry limits admitted Custom source");
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn exercise_output_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = fixture::custom_package(fixture::CustomFamily::Number)
            .unwrap_or_else(|error| panic!("Custom output-limit source: {error}"));
        let limits = PackageLimits::new(
            u64::try_from(source.len()).unwrap_or(u64::MAX),
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("Custom output limits: {error}"));
        let package = Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(limits, options().semantic()),
        )
        .unwrap_or_else(|error| panic!("Custom exact-limit source must open: {error}"));
        let before = package_bytes(&package);
        let name = Name::new(&"N".repeat(MAX_NAME_BYTES)).unwrap();
        let pattern =
            NumberPattern::new(&format!("0{}", "x".repeat(MAX_PATTERN_BYTES - 1))).unwrap();
        let result = package
            .edit_table_cell_custom_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Custom output-limit edit: {error}"))
            .set(CustomNumber::new(name, pattern).into())
            .commit();
        assert!(result.is_err(), "exact Custom output limit admitted growth");
        black_box(result.err().map(|error| error.to_string()));
        assert_eq!(package_bytes(&package), before);
    });
}

fn random_position(data: &[u8]) -> CellPosition {
    CellPosition::new(
        u32::from(control(data, 4) % 8),
        u32::from(control(data, 5) % 8),
    )
}

fn bounded_ascii(data: &[u8], offset: usize, maximum: usize) -> String {
    data.get(offset..)
        .unwrap_or_default()
        .iter()
        .take(maximum)
        .map(|byte| match byte {
            b' '..=b'~' => char::from(*byte),
            _ => '_',
        })
        .collect()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Custom package failed: {error}"));
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
        "Custom error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Custom debug leaked private input"
    );
    black_box((display, debug));
}
