#![no_main]

//! Bounded selector-first fuzzing for Numbers table-cell Scientific formats.
//!
//! Arbitrary bytes first exercise bounded Numbers ingress, then the same
//! command prefix is replayed against a deterministic source-built Scientific
//! package.  This keeps package-level semantic coverage reachable even when a
//! mutated ZIP is rejected before its rooted table graph can be selected.  The
//! focused path covers selector-safe reads, no-op/set/clear/reset transactions,
//! exact patch application and inverse replay, stale and foreign conflicts,
//! copy-on-write format-list ownership, malformed and locked graphs, finite
//! limits, and scalar/source-locality invariants.

use std::{
    fmt::{Debug, Display},
    hint::black_box,
    sync::OnceLock,
};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    SheetSelector, TableSelector,
    cell::data_format::scientific::{
        FixedDecimalPlaces, Scientific, transaction::Error as ScientificError,
    },
};
use litchi_iwa_archive::package::Catalog;
use litchi_numbers_wire::BncCell;

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
const PRIVATE_SHEET: &str = "__litchi_private_scientific_format_sheet_4b91__";
const PRIVATE_TABLE: &str = "__litchi_private_scientific_format_table_4b91__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_scientific_format_input_4b91__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

// B3 is a stable explicit Number cell in the checked-in Numbers source.  It
// is a real-package wrong-family oracle; the source-built fixture below is
// the success source for Scientific transactions.
const NATIVE_POSITION: CellPosition = CellPosition::new(2, 1);
const FIXTURE_POSITION: CellPosition = CellPosition::new(0, 0);
const FIXTURE_SIBLING_POSITION: CellPosition = CellPosition::new(0, 1);

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data, NATIVE_POSITION, false),
        Err(error) => observe_error(error),
    }

    // Replay the command prefix against a stable valid source so semantic
    // operations are not starved by physical package admission.
    exercise_package(scientific_package(), data, FIXTURE_POSITION, true);

    // Number, Percentage, Currency, and a real native Number source all must
    // be refused by the Scientific owner.
    exercise_wrong_family(number_package(), FIXTURE_POSITION);
    exercise_wrong_family(percentage_package(), FIXTURE_POSITION);
    exercise_wrong_family(currency_package(), FIXTURE_POSITION);
    exercise_wrong_family(native_package(), NATIVE_POSITION);

    exercise_copy_on_write_and_refcounts();
    exercise_malformed_graphs();
    exercise_locked_graph();
    exercise_constructor_edges(data);
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_semantic_limit();
    exercise_operation_limit();
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
        .unwrap_or_else(|error| unreachable!("valid Scientific archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid Scientific semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Scientific projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers Scientific seed must open: {error}"))
    })
}

fn scientific_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Scientific,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Scientific fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("deterministic Scientific source must open: {error}"))
    })
}

fn number_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Number,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Number fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic Number wrong-family source must open: {error}")
        })
    })
}

fn percentage_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Percentage,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Percentage fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic Percentage wrong-family source must open: {error}")
        })
    })
}

fn currency_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Currency,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Currency fixture must build: {error}"));
        Package::from_bytes_with_options(&source, options()).unwrap_or_else(|error| {
            panic!("deterministic Currency wrong-family source must open: {error}")
        })
    })
}

fn exercise_package(
    package: &Package,
    data: &[u8],
    position: CellPosition,
    fixture_invariants: bool,
) {
    exercise_selector_reads(package, data, position);
    exercise_selector_errors(package, position);

    let source_bytes = package_bytes(package);
    let before = match package.table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    ) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            if let Err(error) = package.edit_table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            ) {
                observe_error(error);
            }
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    let requested = scientific_from_bytes(data);
    exercise_transaction(
        package,
        &source_bytes,
        before,
        requested,
        position,
        data,
        fixture_invariants,
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8], position: CellPosition) {
    let random_position = CellPosition::new(
        u32::from(control(data, 4) % 8),
        u32::from(control(data, 5) % 8),
    );
    observe_result(package.table_cell_scientific_format(
        SheetSelector::index(usize::from(control(data, 0))),
        TableSelector::index(usize::from(control(data, 1))),
        random_position,
    ));
    observe_result(package.table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    ));

    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        let by_name = package.table_cell_scientific_format(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
            position,
        );
        let by_index = package.table_cell_scientific_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        );
        assert_eq!(by_name, by_index, "Scientific selector forms disagree");
    }
}

fn exercise_selector_errors(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let sheet_count = package.document().sheets().len();
    if let Err(error) = package.table_cell_scientific_format(
        SheetSelector::index(sheet_count),
        TableSelector::index(0),
        position,
    ) {
        observe_error(error);
    }
    if let Err(error) = package.table_cell_scientific_format(
        SheetSelector::name(PRIVATE_SHEET),
        TableSelector::index(0),
        position,
    ) {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) = package.table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::name(PRIVATE_TABLE),
        position,
    ) {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) = package.table_cell_scientific_format(
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
    before: Option<Scientific>,
    requested: Scientific,
    position: CellPosition,
    data: &[u8],
    fixture_invariants: bool,
) {
    let edit = match package.edit_table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    ) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    // Four commands cover an exact semantic no-op, set, clear, and reset.
    // Preserve the source Optional state for the no-op command.
    let command = control(data, 0) & 3;
    let (edit, expected_after) = match command {
        0 => match before {
            Some(format) => (edit.set(format), before),
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
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert!(commit.diagnostics().full_reparse_performed());
        if fixture_invariants {
            assert_exact_locality(source_bytes, &target_bytes);
            assert_non_format_bnc_bytes(source_bytes, &target_bytes);
        }
    }
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    black_box(commit.diagnostics());
    assert_eq!(
        commit
            .package()
            .table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("Scientific candidate readback failed: {error}")),
        expected_after
    );

    let applied = package
        .apply_table_cell_scientific_format(&patch)
        .unwrap_or_else(|error| panic!("Scientific patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_table_cell_scientific_format(&patch)
                .is_err(),
            "a committed Scientific patch must reject a second application"
        );
        assert!(
            package
                .apply_table_cell_scientific_format(&patch.inverse())
                .is_err(),
            "a Scientific inverse must reject the original source"
        );

        // Reopening the target creates a stale artifact from the same logical
        // source lineage. Its current value is already `after`, so the forward
        // patch must refuse it without publication.
        let stale = Package::from_bytes_with_options(&target_bytes, options())
            .unwrap_or_else(|error| panic!("Scientific stale candidate must reopen: {error}"));
        let stale_before = package_bytes(&stale);
        assert!(stale.apply_table_cell_scientific_format(&patch).is_err());
        assert_eq!(package_bytes(&stale), stale_before);

        // A patch from a Scientific source is not transferable to a distinct
        // Number artifact, even when selectors happen to line up.
        let foreign = number_package();
        let foreign_before = package_bytes(foreign);
        assert!(foreign.apply_table_cell_scientific_format(&patch).is_err());
        assert_eq!(package_bytes(foreign), foreign_before);
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_cell_scientific_format(&inverse)
        .unwrap_or_else(|error| panic!("Scientific inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("Scientific inverse readback failed: {error}")),
        before
    );
    black_box((patch.before(), patch.after(), patch.path()));
}

fn exercise_wrong_family(package: &Package, position: CellPosition) {
    let source_bytes = package_bytes(package);
    let read = package.table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    );
    assert!(
        matches!(read, Err(ScientificError::WrongFormatFamily { .. })),
        "non-Scientific cell must be refused by the Scientific API: {read:?}"
    );
    let edit = package.edit_table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    );
    assert!(
        matches!(edit, Err(ScientificError::WrongFormatFamily { .. })),
        "non-Scientific cell edit must be refused by the Scientific API: {edit:?}"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_copy_on_write_and_refcounts() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = package_bytes(scientific_package());
        assert_eq!(
            fixture::format_entry_facts(&source)
                .unwrap_or_else(|error| panic!("Scientific format facts: {error}")),
            vec![(fixture::FIRST_FORMAT_KEY, 2)]
        );
        assert_eq!(
            fixture::format_keys(&source)
                .unwrap_or_else(|error| panic!("Scientific format keys: {error}")),
            vec![
                Some(fixture::FIRST_FORMAT_KEY),
                Some(fixture::FIRST_FORMAT_KEY)
            ]
        );

        let replacement = Scientific::new(
            FixedDecimalPlaces::new(7)
                .unwrap_or_else(|error| panic!("valid Scientific precision: {error}")),
        );
        let first = scientific_package()
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Scientific COW edit open: {error}"))
            .set(replacement)
            .commit()
            .unwrap_or_else(|error| panic!("Scientific COW first commit: {error}"));
        let first_bytes = package_bytes(first.package());
        let first_keys = fixture::format_keys(&first_bytes)
            .unwrap_or_else(|error| panic!("Scientific COW first keys: {error}"));
        assert_ne!(
            first_keys[0], first_keys[1],
            "shared Scientific entry must COW"
        );
        assert_eq!(first_keys[1], Some(fixture::FIRST_FORMAT_KEY));
        let new_key = first_keys[0].unwrap_or_else(|| panic!("Scientific COW key missing"));
        assert_eq!(
            fixture::format_entry_facts(&first_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW first facts: {error}")),
            vec![(fixture::FIRST_FORMAT_KEY, 1), (new_key, 1)]
        );
        assert_non_format_bnc_bytes(&source, &first_bytes);

        let both = first
            .package()
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("Scientific COW sibling edit open: {error}"))
            .set(replacement)
            .commit()
            .unwrap_or_else(|error| panic!("Scientific COW sibling commit: {error}"));
        let both_bytes = package_bytes(both.package());
        assert_eq!(
            fixture::format_keys(&both_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW shared keys: {error}")),
            vec![Some(new_key), Some(new_key)]
        );
        assert_eq!(
            fixture::format_entry_facts(&both_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW shared facts: {error}")),
            vec![(new_key, 2)]
        );
        assert_non_format_bnc_bytes(&first_bytes, &both_bytes);

        let one_cleared = both
            .package()
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Scientific COW clear edit open: {error}"))
            .clear()
            .commit()
            .unwrap_or_else(|error| panic!("Scientific COW clear commit: {error}"));
        let one_cleared_bytes = package_bytes(one_cleared.package());
        assert_eq!(
            fixture::format_keys(&one_cleared_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW one-cleared keys: {error}")),
            vec![None, Some(new_key)]
        );
        assert_eq!(
            fixture::format_entry_facts(&one_cleared_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW one-cleared facts: {error}")),
            vec![(new_key, 1)]
        );
        assert_non_format_bnc_bytes(&both_bytes, &one_cleared_bytes);

        let all_cleared = one_cleared
            .package()
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_SIBLING_POSITION)
            .unwrap_or_else(|error| panic!("Scientific COW final clear edit open: {error}"))
            .reset()
            .commit()
            .unwrap_or_else(|error| panic!("Scientific COW final reset commit: {error}"));
        let all_cleared_bytes = package_bytes(all_cleared.package());
        assert_eq!(
            fixture::format_keys(&all_cleared_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW final keys: {error}")),
            vec![None, None]
        );
        assert!(
            fixture::format_entry_facts(&all_cleared_bytes)
                .unwrap_or_else(|error| panic!("Scientific COW final facts: {error}"))
                .is_empty()
        );
        assert_non_format_bnc_bytes(&one_cleared_bytes, &all_cleared_bytes);
    });
}

fn exercise_malformed_graphs() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let corruptions = [
            fixture::Corruption::DuplicateFormatKey,
            fixture::Corruption::MissingFormatEntry,
            fixture::Corruption::RefcountMismatch,
            fixture::Corruption::DuplicateFormatList,
            fixture::Corruption::AliasedFormatList,
            fixture::Corruption::UnsupportedFormatType,
            fixture::Corruption::MalformedFormatPayload,
            fixture::Corruption::WrongCellFormatKey,
            fixture::Corruption::UnterminatedUnknownGroup,
            fixture::Corruption::UnexpectedFieldReference,
        ];
        for corruption in corruptions {
            let source =
                fixture::corrupted_package_for(fixture::FormatFamily::Scientific, corruption)
                    .unwrap_or_else(|error| panic!("Scientific malformed fixture: {error}"));
            let package = match Package::from_bytes_with_options(&source, options()) {
                Ok(package) => package,
                Err(error) => {
                    observe_error(error);
                    continue;
                },
            };
            let before = package_bytes(&package);
            observe_result(package.table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                FIXTURE_POSITION,
            ));
            let edit = package.edit_table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                FIXTURE_POSITION,
            );
            let result = match edit {
                Ok(edit) => edit
                    // The canonical fixture value is Scientific::default()
                    // (two places), so use a definite change here; otherwise
                    // a malformed source could be accepted by the intentional
                    // no-op fast path instead of reaching validation.
                    .set(Scientific::new(FixedDecimalPlaces::new(3).unwrap_or_else(
                        |error| panic!("Scientific malformed precision: {error}"),
                    )))
                    .commit()
                    .map(|commit| package_bytes(commit.package())),
                Err(error) => {
                    observe_error(error);
                    Err(ScientificError::Verification)
                },
            };
            assert!(result.is_err(), "malformed Scientific graph was published");
            if let Err(error) = result {
                observe_error(error);
            }
            assert_eq!(package_bytes(&package), before);
        }
    });
}

fn exercise_locked_graph() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = fixture::locked_table_package(
            &fixture::synthetic_package_for(
                fixture::FormatFamily::Scientific,
                fixture::FormatSharing::Shared,
            )
            .unwrap_or_else(|error| panic!("Scientific lock source: {error}")),
        )
        .unwrap_or_else(|error| panic!("Scientific locked fixture: {error}"));
        let package = Package::from_bytes_with_options(&source, options())
            .unwrap_or_else(|error| panic!("locked Scientific package must open: {error}"));
        let before = package_bytes(&package);
        let result = package
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("locked Scientific edit open: {error}"))
            .set(Scientific::new(FixedDecimalPlaces::new(30).unwrap_or_else(
                |error| panic!("Scientific lock precision: {error}"),
            )))
            .commit();
        assert!(
            matches!(result, Err(ScientificError::TableLocked { .. })),
            "changed Scientific edit must refuse a locked table: {result:?}"
        );
        assert_eq!(package_bytes(&package), before);
    });
}

fn exercise_constructor_edges(data: &[u8]) {
    assert!(FixedDecimalPlaces::new(31).is_err());
    for value in [0, 1, 2, 30] {
        let places = FixedDecimalPlaces::new(value)
            .unwrap_or_else(|error| panic!("valid Scientific precision rejected: {error}"));
        black_box(places);
    }
    let format = scientific_from_bytes(data);
    black_box((format.decimal_places(), format));
}

fn scientific_from_bytes(data: &[u8]) -> Scientific {
    Scientific::new(
        FixedDecimalPlaces::new(control(data, 1) % 31)
            .unwrap_or_else(|error| panic!("bounded Scientific precision must validate: {error}")),
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
                        unreachable!("malformed sentinel is UTF-8: {error}")
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
            Err(PackageError::Archive(error)) => {
                assert_eq!(
                    error.to_string(),
                    format!(
                        "iWork archive input bytes limit exceeded: observed {OVERSIZED_INPUT_BYTES}, maximum {MAX_INPUT_BYTES}"
                    )
                );
                black_box(error);
            },
            Err(error) => observe_error(error),
            Ok(_) => panic!("oversized Numbers Scientific-format input must be rejected"),
        }
    });
}

fn exercise_semantic_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let semantic = PackageSemanticLimits::new(1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("one-entry Scientific semantic limits: {error}"));
        let constrained = PackageReadOptions::new(options().archive(), semantic);
        let result = Package::from_bytes_with_options(
            &fixture::synthetic_package_for(
                fixture::FormatFamily::Scientific,
                fixture::FormatSharing::Shared,
            )
            .unwrap_or_else(|error| panic!("Scientific semantic-limit source: {error}")),
            constrained,
        );
        assert!(
            result.is_err(),
            "one-entry semantic limits admitted Scientific source"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn exercise_operation_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source = fixture::synthetic_package_for(
            fixture::FormatFamily::Scientific,
            fixture::FormatSharing::Shared,
        )
        .unwrap_or_else(|error| panic!("Scientific operation-limit source: {error}"));
        let limits = PackageLimits::new(
            u64::try_from(source.len()).unwrap_or(u64::MAX),
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("Scientific operation limits: {error}"));
        let package = Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(limits, PackageSemanticLimits::default()),
        )
        .unwrap_or_else(|error| panic!("Scientific exact-limit source must open: {error}"));
        let before = package_bytes(&package);
        let result = package
            .edit_table_cell_scientific_format(0usize, 0usize, FIXTURE_POSITION)
            .unwrap_or_else(|error| panic!("Scientific operation-limit edit open: {error}"))
            .set(Scientific::new(FixedDecimalPlaces::new(30).unwrap_or_else(
                |error| panic!("Scientific operation precision: {error}"),
            )))
            .commit();
        assert!(
            result.is_err(),
            "Scientific operation limit admitted growth"
        );
        black_box(result.err().map(|error| error.to_string()));
        assert_eq!(package_bytes(&package), before);
    });
}

fn assert_exact_locality(source: &[u8], target: &[u8]) {
    let before =
        Catalog::from_bytes(source).unwrap_or_else(|error| panic!("source catalog: {error}"));
    let after =
        Catalog::from_bytes(target).unwrap_or_else(|error| panic!("target catalog: {error}"));
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .unwrap_or_else(|| panic!("candidate removed source member {}", entry.name()));
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged Scientific member {} lost its exact local ZIP record",
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
                "unrelated Scientific member changed"
            );
        }
    }
    assert_eq!(changed, [fixture::TABLES_MEMBER.to_owned()]);
    assert_eq!(before.len(), after.len());
}

fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) {
    let before =
        fixture::tile_cells(source).unwrap_or_else(|error| panic!("source tile cells: {error}"));
    let after =
        fixture::tile_cells(target).unwrap_or_else(|error| panic!("target tile cells: {error}"));
    assert_eq!(before.len(), after.len());
    for (left, right) in before.iter().zip(after.iter()) {
        let left_cell =
            BncCell::parse(left).unwrap_or_else(|error| panic!("source BNC cell: {error}"));
        let right_cell =
            BncCell::parse(right).unwrap_or_else(|error| panic!("target BNC cell: {error}"));
        assert_eq!(
            left_cell
                .cached_scalar()
                .unwrap_or_else(|error| panic!("source scalar: {error}")),
            right_cell
                .cached_scalar()
                .unwrap_or_else(|error| panic!("target scalar: {error}")),
            "Scientific edit changed the cached scalar"
        );
        assert_eq!(
            normalized_cell(left),
            normalized_cell(right),
            "Scientific edit changed non-format BNC bytes"
        );
    }
}

fn normalized_cell(cell: &[u8]) -> Vec<u8> {
    let mut parsed =
        BncCell::parse(cell).unwrap_or_else(|error| panic!("BNC normalization: {error}"));
    parsed.clear_explicit_format();
    parsed.encode()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Scientific package failed: {error}"));
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
        "Scientific error leaked private input"
    );
    assert!(
        !debug.contains(private),
        "Scientific error debug leaked private input"
    );
    black_box((display, debug));
}
