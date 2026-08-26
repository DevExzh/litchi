#![no_main]

//! Bounded selector-first Numbers Pop-Up Menu lifecycle fuzzing.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::data_format::pop_up_menu::{InitialSelection, Item, PopUpMenu, transaction::Patch},
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
/// The split-owner source is shared with the unified-control target.  It has
/// a rooted Pop-Up cell whose model is in a distinct current component while
/// the table/tile/list graph remains in its own members.  Keeping one
/// file-backed source avoids inventing a second synthetic archive and makes
/// this direct popup target exercise the same metadata/token closure.
const SPLIT_COMPONENT_NUMBERS: &[u8] =
    include_bytes!("../corpus/numbers_table_cell_control/split_component_source.numbers");

fuzz_target!(|data: &[u8]| {
    if let Ok(package) = Package::from_bytes_with_options(data, options()) {
        exercise_package(&package, data);
    } else {
        exercise_input_error(data);
    }
    exercise_package(native_package(), data);
    exercise_split_component_source(data);
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

fn split_component_package() -> Option<&'static Package> {
    static PACKAGE: OnceLock<Option<Package>> = OnceLock::new();
    PACKAGE
        .get_or_init(|| {
            match Package::from_bytes_with_options(SPLIT_COMPONENT_NUMBERS, options()) {
                Ok(package) => Some(package),
                Err(error) => {
                    observe_error(error);
                    None
                },
            }
        })
        .as_ref()
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

/// Drive the file-backed split graph independently of arbitrary ZIP ingress.
/// The source has an existing Pop-Up model in a separate current component;
/// every successful branch must therefore report the full member closure,
/// reopen the serialized candidate, and preserve exact patch/inverse bytes.
/// Unsupported or malformed ownership remains an observed, source-atomic
/// error so this target stays useful while the graph owner evolves.
fn exercise_split_component_source(data: &[u8]) {
    let Some(package) = split_component_package() else {
        return;
    };
    let popup_position = CellPosition::new(4, 0);
    let create_position = CellPosition::new(0, 1);
    let sibling_position = CellPosition::new(1, 1);
    let desired = menu(data);

    if let Ok(Some(before)) = package.table_cell_pop_up_menu_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        popup_position,
    ) {
        exercise_split_existing_popup(package, popup_position, before, desired.clone(), data);
    }
    exercise_split_create_reuse(package, create_position, sibling_position, desired, data);
    exercise_split_component_limits();
    exercise_split_corruption_probes(data);
}

fn exercise_split_existing_popup(
    package: &Package,
    position: CellPosition,
    before: PopUpMenu,
    desired: PopUpMenu,
    data: &[u8],
) {
    let source_bytes = package_bytes(package);
    let no_op = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )
        .and_then(|edit| edit.set(before.clone()).commit());
    match no_op {
        Ok(commit) => {
            assert!(commit.patch().is_noop());
            assert_eq!(commit.diagnostics().touched_components(), 0);
            assert_eq!(package_bytes(commit.package()), source_bytes);
        },
        Err(error) => observe_error(error),
    }

    if before == desired {
        return;
    }
    let result = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )
        .and_then(|edit| edit.set(desired.clone()).commit());
    let Ok(commit) = result else {
        assert_eq!(package_bytes(package), source_bytes);
        return;
    };
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), Some(&before));
    assert_eq!(patch.after(), Some(&desired));
    assert!(!patch.is_noop());
    assert!(commit.diagnostics().touched_components() >= 2);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit
            .package()
            .table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("split popup candidate readback failed: {error}")),
        Some(desired.clone())
    );
    let applied = package
        .apply_table_cell_pop_up_menu_format(&patch)
        .unwrap_or_else(|error| panic!("split popup patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert!(package.apply_table_cell_pop_up_menu_format(&patch).is_err());
    let restored = commit
        .package()
        .apply_table_cell_pop_up_menu_format(&patch.inverse())
        .unwrap_or_else(|error| panic!("split popup inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    exercise_split_reopened_replay(
        &source_bytes,
        &target_bytes,
        &patch,
        position,
        commit.diagnostics().touched_components(),
    );
    black_box(data);
}

fn exercise_split_create_reuse(
    package: &Package,
    create_position: CellPosition,
    sibling_position: CellPosition,
    desired: PopUpMenu,
    data: &[u8],
) {
    let source_bytes = package_bytes(package);
    let result = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            create_position,
        )
        .and_then(|edit| edit.set(desired.clone()).commit());
    let Ok(created) = result else {
        assert_eq!(package_bytes(package), source_bytes);
        return;
    };
    let create_patch = created.patch().clone();
    let created_bytes = package_bytes(created.package());
    assert_eq!(create_patch.before(), None);
    assert_eq!(create_patch.after(), Some(&desired));
    assert!(!create_patch.is_noop());
    assert!(created.diagnostics().touched_components() >= 2);
    assert_eq!(
        created
            .package()
            .table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                create_position,
            )
            .unwrap_or_else(|error| panic!("split popup create readback failed: {error}")),
        Some(desired.clone())
    );
    let restored = created
        .package()
        .apply_table_cell_pop_up_menu_format(&create_patch.inverse())
        .unwrap_or_else(|error| panic!("split popup create inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    exercise_split_reopened_replay(
        &source_bytes,
        &created_bytes,
        &create_patch,
        create_position,
        created.diagnostics().touched_components(),
    );

    let reuse_source = package_bytes(created.package());
    let reuse = created
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            sibling_position,
        )
        .and_then(|edit| edit.set(desired.clone()).commit());
    let Ok(reused) = reuse else {
        assert_eq!(package_bytes(created.package()), reuse_source);
        return;
    };
    let reuse_patch = reused.patch().clone();
    let reused_bytes = package_bytes(reused.package());
    assert_eq!(reuse_patch.before(), None);
    assert_eq!(reuse_patch.after(), Some(&desired));
    assert!(!reuse_patch.is_noop());
    assert!(reused.diagnostics().touched_components() >= 2);
    assert_eq!(
        reused
            .package()
            .table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                sibling_position,
            )
            .unwrap_or_else(|error| panic!("split popup reuse readback failed: {error}")),
        Some(desired.clone())
    );
    let clear_first = reused
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            create_position,
        )
        .and_then(|edit| edit.clear().commit());
    if let Ok(cleared) = clear_first {
        assert_eq!(cleared.patch().after(), None);
        assert!(cleared.diagnostics().touched_components() >= 2);
        assert_eq!(
            cleared
                .package()
                .table_cell_pop_up_menu_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    sibling_position,
                )
                .unwrap_or_else(|error| panic!("split popup sibling readback failed: {error}")),
            Some(desired.clone())
        );
        let clear_source = reused_bytes.clone();
        let clear_target = package_bytes(cleared.package());
        let clear_restored = cleared
            .package()
            .apply_table_cell_pop_up_menu_format(&cleared.patch().inverse())
            .unwrap_or_else(|error| panic!("split popup clear inverse must apply: {error}"));
        assert_eq!(package_bytes(clear_restored.package()), clear_source);
        exercise_split_reopened_replay(
            &clear_source,
            &clear_target,
            cleared.patch(),
            create_position,
            cleared.diagnostics().touched_components(),
        );

        let final_clear = cleared
            .package()
            .edit_table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                sibling_position,
            )
            .and_then(|edit| edit.reset().commit());
        if let Ok(final_clear) = final_clear {
            assert_eq!(final_clear.patch().after(), None);
            assert!(final_clear.diagnostics().touched_components() >= 2);
            assert_eq!(
                final_clear
                    .package()
                    .table_cell_pop_up_menu_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        sibling_position,
                    )
                    .unwrap_or_else(|error| panic!(
                        "split popup final cull readback failed: {error}"
                    )),
                None
            );
            let final_source = clear_target;
            let final_target = package_bytes(final_clear.package());
            let final_restored = final_clear
                .package()
                .apply_table_cell_pop_up_menu_format(&final_clear.patch().inverse())
                .unwrap_or_else(|error| panic!("split popup final inverse must apply: {error}"));
            assert_eq!(package_bytes(final_restored.package()), final_source);
            exercise_split_reopened_replay(
                &final_source,
                &final_target,
                final_clear.patch(),
                sibling_position,
                final_clear.diagnostics().touched_components(),
            );
        } else {
            assert_eq!(package_bytes(cleared.package()), clear_target);
        }
    } else {
        assert_eq!(package_bytes(reused.package()), reused_bytes);
    }
    black_box(data);
}

fn exercise_split_reopened_replay(
    source_bytes: &[u8],
    target_bytes: &[u8],
    patch: &Patch,
    position: CellPosition,
    touched_components: usize,
) {
    let source = Package::from_bytes_with_options(source_bytes, options())
        .unwrap_or_else(|error| panic!("split popup source must reopen: {error}"));
    let target = Package::from_bytes_with_options(target_bytes, options())
        .unwrap_or_else(|error| panic!("split popup target must reopen: {error}"));
    assert_eq!(
        target
            .table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .unwrap_or_else(|error| panic!("split popup reopened read failed: {error}")),
        patch.after().cloned()
    );
    let restored = target
        .apply_table_cell_pop_up_menu_format(&patch.inverse())
        .unwrap_or_else(|error| panic!("split popup reopened inverse failed: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    let replayed = source
        .apply_table_cell_pop_up_menu_format(patch)
        .unwrap_or_else(|error| panic!("split popup reopened patch failed: {error}"));
    assert_eq!(package_bytes(replayed.package()), target_bytes);
    assert!(source.apply_table_cell_pop_up_menu_format(patch).is_err());
    if touched_components > 1 {
        black_box(touched_components);
    }
}

fn exercise_split_component_limits() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source_len = u64::try_from(SPLIT_COMPONENT_NUMBERS.len())
            .unwrap_or_else(|error| panic!("split popup source length conversion failed: {error}"));
        let exact_archive = PackageLimits::new(
            source_len,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("split popup exact limits invalid: {error}"));
        let exact = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(exact_archive, PackageSemanticLimits::default()),
        )
        .unwrap_or_else(|error| {
            panic!("split popup source rejected at exact input ceiling: {error}")
        });
        black_box(exact);

        let tight_archive = PackageLimits::new(
            source_len.saturating_sub(1),
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("split popup tight limits invalid: {error}"));
        let tight = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(tight_archive, PackageSemanticLimits::default()),
        );
        assert!(tight.is_err(), "split popup input-minus-one was admitted");
        black_box(tight.err().map(|error| error.to_string()));

        for (label, limits) in [
            (
                "entries",
                PackageLimits::new(
                    source_len,
                    1,
                    MAX_ENTRY_BYTES,
                    MAX_EXPANDED_BYTES,
                    MAX_IWA_STREAM_BYTES,
                ),
            ),
            (
                "entry-bytes",
                PackageLimits::new(
                    source_len,
                    MAX_ENTRIES,
                    1,
                    MAX_EXPANDED_BYTES,
                    MAX_IWA_STREAM_BYTES,
                ),
            ),
            (
                "total-bytes",
                PackageLimits::new(
                    source_len,
                    MAX_ENTRIES,
                    MAX_ENTRY_BYTES,
                    1,
                    MAX_IWA_STREAM_BYTES,
                ),
            ),
            (
                "iwa-stream-bytes",
                PackageLimits::new(
                    source_len,
                    MAX_ENTRIES,
                    MAX_ENTRY_BYTES,
                    MAX_EXPANDED_BYTES,
                    1,
                ),
            ),
        ] {
            let limits = limits
                .unwrap_or_else(|error| panic!("split popup {label} limits invalid: {error}"));
            let result = Package::from_bytes_with_options(
                SPLIT_COMPONENT_NUMBERS,
                PackageReadOptions::new(limits, PackageSemanticLimits::default()),
            );
            assert!(result.is_err(), "split popup {label} limit admitted source");
            black_box(result.err().map(|error| error.to_string()));
        }

        let object_limit = PackageSemanticLimits::new(1, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
            .unwrap_or_else(|error| panic!("split popup object limit invalid: {error}"));
        let result = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(PackageLimits::default(), object_limit),
        );
        assert!(result.is_err(), "split popup object-minus-one was admitted");
        black_box(result.err().map(|error| error.to_string()));
    });
}

/// Mutate a few bounded ZIP bytes to exercise malformed metadata, alias, and
/// framing ingress.  Arbitrary fuzz input also reaches this path; the fixed
/// source probes make the malformed route deterministic in every campaign.
fn exercise_split_corruption_probes(data: &[u8]) {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let offsets = [
            usize::from(data.first().copied().unwrap_or_default()) % SPLIT_COMPONENT_NUMBERS.len(),
            usize::from(data.get(1).copied().unwrap_or_default())
                .saturating_mul(257)
                .saturating_add(31)
                % SPLIT_COMPONENT_NUMBERS.len(),
            SPLIT_COMPONENT_NUMBERS.len().saturating_sub(1),
        ];
        for (index, offset) in offsets.into_iter().enumerate() {
            let mut corrupted = SPLIT_COMPONENT_NUMBERS.to_vec();
            corrupted[offset] ^= [0x01, 0x20, 0x80][index];
            match Package::from_bytes_with_options(&corrupted, options()) {
                Ok(package) => {
                    // A mutation that happens to preserve a valid ZIP must
                    // still be safe to read and must not publish anything.
                    observe_result(package.table_cell_pop_up_menu_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        CellPosition::new(4, 0),
                    ));
                    black_box(package_bytes(&package));
                },
                Err(error) => observe_error(error),
            }
        }
    });
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
