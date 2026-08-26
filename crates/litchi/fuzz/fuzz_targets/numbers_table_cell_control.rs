#![no_main]

//! Bounded selector-first fuzzing for the unified Numbers cell-control owner.
//!
//! The command stream deliberately drives all five control variants through
//! one transaction surface.  A source that does not contain an admitted
//! control graph is still useful: selector errors, constructor validation,
//! ingress limits, and source-atomic failures remain covered without making
//! native identifiers part of this harness.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::data_format::control::transaction::Patch,
    cell::data_format::control::{CellControl, DisplayFormat, Range, Slider, Stepper},
    cell::data_format::{Checkbox, Number, PopUpMenu, StarRating},
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
const PRIVATE_SHEET: &str = "__litchi_private_control_sheet_85__";
const PRIVATE_TABLE: &str = "__litchi_private_control_table_85__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_control_input_85__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");
/// The committed Wave86 split-owner source is retained as the Wave88
/// multi-member write seed.  Its rooted model, tile, format/control lists,
/// and Pop-Up model are separate current components with metadata edges.
const SPLIT_COMPONENT_NUMBERS: &[u8] =
    include_bytes!("../corpus/numbers_table_cell_control/split_component_source.numbers");
const ZIP_LOCAL_HEADER: &[u8] = b"PK\x03\x04";

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected native bytes make arbitrary ZIP mutation unlikely to
    // reach a semantic table.  Reuse the same command against a fixed native
    // source so every campaign still exercises selectors and transactions.
    exercise_package(native_package(), data);
    // A valid ZIP supplied as fuzz input may be a multi-member native source
    // whose control lists live outside the selected model component.  Probe
    // its first bounded table window as well: unsupported split ownership
    // must remain source-atomic, while any admitted route must survive a
    // serialized candidate reopen and exact forward/inverse replay.
    if data.starts_with(ZIP_LOCAL_HEADER)
        && let Ok(package) = Package::from_bytes_with_options(data, options())
    {
        exercise_split_component_window(&package, data);
    }
    exercise_split_component_source(data);
    exercise_constructors(data);
    exercise_selector_errors(native_package(), data);
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
        .unwrap_or_else(|error| unreachable!("valid control archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid control semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| unreachable!("valid control projection limits: {error}"));
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers control seed must open: {error}"))
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
    let position = CellPosition::new(read_u16(data, 0) as u32 % 8, read_u16(data, 2) as u32 % 8);
    let source_bytes = package_bytes(package);
    let before = match package.table_cell_control_format(sheet, table, position) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            return;
        },
    };

    // Every fifth command selects one of the five unified semantic controls.
    // The same command is then used as a second kind for an admitted source,
    // exercising cross-kind replacement without native IDs in the harness.
    let first_kind = usize::from(data.first().copied().unwrap_or_default() % 5);
    let first = control(first_kind, data);
    exercise_set(package, data, position, before, source_bytes, first);
}

fn exercise_set(
    package: &Package,
    data: &[u8],
    position: CellPosition,
    before: Option<CellControl>,
    source_bytes: Vec<u8>,
    desired: CellControl,
) {
    let sheet = SheetSelector::index(0);
    let table = TableSelector::index(0);
    let edit = match package.edit_table_cell_control_format(sheet, table, position) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
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
    if !patch.is_noop() {
        assert!(
            commit.diagnostics().full_reparse_performed(),
            "a published control candidate must be reopened before publication"
        );
    }
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), Some(&desired));
    assert_eq!(patch.is_noop(), target_bytes == source_bytes);
    black_box(commit.diagnostics());
    assert_eq!(
        commit
            .package()
            .table_cell_control_format(sheet, table, position)
            .unwrap_or_else(|error| panic!("control candidate readback failed: {error}")),
        Some(desired.clone())
    );

    let applied = package
        .apply_table_cell_control_format(&patch)
        .unwrap_or_else(|error| panic!("control patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(package.apply_table_cell_control_format(&patch).is_err());
    }
    let restored = commit
        .package()
        .apply_table_cell_control_format(&patch.inverse())
        .unwrap_or_else(|error| panic!("control inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(patch.inverse().inverse(), patch);
    exercise_reopened_replay(
        &source_bytes,
        &target_bytes,
        &patch,
        position,
        commit.diagnostics().touched_components(),
    );
    let first_target = commit.package();
    let second_kind = usize::from(data.get(1).copied().unwrap_or_default() % 5);
    if second_kind != control_kind(&desired) {
        let second = control(second_kind, data);
        if let Ok(edit) = first_target.edit_table_cell_control_format(sheet, table, position)
            && let Ok(second_commit) = edit.set(second.clone()).commit()
        {
            let second_patch = second_commit.patch().clone();
            let second_source_bytes = package_bytes(first_target);
            let second_bytes = package_bytes(second_commit.package());
            assert_eq!(second_patch.after(), Some(&second));
            assert_eq!(
                second_commit
                    .package()
                    .table_cell_control_format(sheet, table, position)
                    .unwrap_or_else(|error| panic!("control transition readback failed: {error}")),
                Some(second.clone())
            );
            let applied = first_target
                .apply_table_cell_control_format(&second_patch)
                .unwrap_or_else(|error| panic!("control transition patch must apply: {error}"));
            assert_eq!(package_bytes(applied.package()), second_bytes);
            let restored = second_commit
                .package()
                .apply_table_cell_control_format(&second_patch.inverse())
                .unwrap_or_else(|error| panic!("control transition inverse must apply: {error}"));
            assert_eq!(
                package_bytes(restored.package()),
                package_bytes(first_target)
            );
            assert!(
                first_target
                    .apply_table_cell_control_format(&second_patch)
                    .is_err()
            );
            exercise_reopened_replay(
                &second_source_bytes,
                &second_bytes,
                &second_patch,
                position,
                second_commit.diagnostics().touched_components(),
            );
        }
    }

    // Clear and reset are deliberately equivalent commands at this generic
    // boundary. They may be refused for unsupported/locked native graphs, but
    // any successful clear must be exactly invertible.
    if let Ok(edit) = first_target.edit_table_cell_control_format(sheet, table, position) {
        let reset_edit = if data.get(2).copied().unwrap_or_default() & 1 != 0 {
            edit.clear()
        } else {
            edit.reset()
        };
        let Ok(clear_commit) = reset_edit.commit() else {
            return;
        };
        let clear_patch = clear_commit.patch().clone();
        let clear_source_bytes = package_bytes(first_target);
        let clear_target_bytes = package_bytes(clear_commit.package());
        assert_eq!(clear_patch.after(), None);
        assert_eq!(
            clear_commit
                .package()
                .table_cell_control_format(sheet, table, position)
                .unwrap_or_else(|error| panic!("control clear readback failed: {error}")),
            None
        );
        let restored = clear_commit
            .package()
            .apply_table_cell_control_format(&clear_patch.inverse())
            .unwrap_or_else(|error| panic!("control clear inverse must apply: {error}"));
        assert_eq!(
            package_bytes(restored.package()),
            package_bytes(first_target)
        );
        exercise_reopened_replay(
            &clear_source_bytes,
            &clear_target_bytes,
            &clear_patch,
            position,
            clear_commit.diagnostics().touched_components(),
        );
    }
}

/// Replay a successful patch through serialized source/target packages.
///
/// This deliberately works from package bytes rather than retaining the
/// private candidate returned by `commit`. Sources with missing, duplicate,
/// versioned, or wrong-locator ownership must fail before this function is
/// reached and leave their source bytes unchanged.
fn exercise_reopened_replay(
    source_bytes: &[u8],
    target_bytes: &[u8],
    patch: &Patch,
    position: CellPosition,
    touched_components: usize,
) {
    let source = Package::from_bytes_with_options(source_bytes, options())
        .unwrap_or_else(|error| panic!("published control source must reopen: {error}"));
    let target = Package::from_bytes_with_options(target_bytes, options())
        .unwrap_or_else(|error| panic!("published control target must reopen: {error}"));
    let selected = target
        .table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)
        .unwrap_or_else(|error| panic!("reopened control target read failed: {error}"));
    assert_eq!(selected, patch.after().cloned());

    let restored = target
        .apply_table_cell_control_format(&patch.inverse())
        .unwrap_or_else(|error| panic!("reopened control inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);

    let replayed = source
        .apply_table_cell_control_format(patch)
        .unwrap_or_else(|error| panic!("reopened control patch must apply: {error}"));
    assert_eq!(package_bytes(replayed.package()), target_bytes);
    if !patch.is_noop() {
        assert!(source.apply_table_cell_control_format(patch).is_err());
    }
    assert_eq!(package_bytes(&source), source_bytes);

    // Keep the diagnostic cardinality visible to coverage without assuming a
    // particular physical layout for successful same-owner candidates.
    if touched_components > 1 {
        black_box(touched_components);
    }
}

/// Exercise the committed split-component graph independently from mutated
/// fuzz input.  This is the Wave88 write gate: a successful changed
/// transaction must rewrite at least two native members, preserve the
/// metadata external-edge/token closure, reopen the candidate, and support
/// exact patch/inverse/conflict replay.  Older owners may reject the changed
/// route; that is still required to be atomic and is deliberately accepted
/// here so the target remains useful across the migration boundary.
fn exercise_split_component_source(data: &[u8]) {
    let Some(package) = split_component_package() else {
        return;
    };
    let positions = [
        CellPosition::new(0, 0),
        CellPosition::new(1, 0),
        CellPosition::new(2, 0),
        CellPosition::new(3, 0),
        CellPosition::new(4, 0),
    ];
    let mut observed = Vec::new();
    for (index, position) in positions.into_iter().enumerate() {
        let Ok(Some(before)) = package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        ) else {
            continue;
        };
        observed.push((position, before.clone()));
        let desired = control(
            usize::from(data.get(index.wrapping_add(1)).copied().unwrap_or_default()) % 5,
            data,
        );
        exercise_split_write(package, position, before, desired, data);
    }
    exercise_split_shared_refcount(package, &observed, data);
    exercise_split_component_limits();
}

/// Run one split-component no-op/change/clear transaction and all exact
/// source-bound patch invariants.  The metadata UUID/token and external-edge
/// ownership is intentionally observed through candidate reopen/locality and
/// the actual touched-member cardinality rather than leaking physical types
/// into this fuzz boundary.
fn exercise_split_write(
    package: &Package,
    position: CellPosition,
    before: CellControl,
    desired: CellControl,
    data: &[u8],
) {
    let source_bytes = package_bytes(package);
    let no_op = package
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)
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
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)
        .and_then(|edit| edit.set(desired.clone()).commit());
    let Ok(commit) = result else {
        if let Err(error) = package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )
            .and_then(|edit| edit.set(desired).commit())
        {
            observe_error(error);
        }
        assert_eq!(package_bytes(package), source_bytes);
        return;
    };
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), Some(&before));
    assert_eq!(patch.after(), Some(&desired));
    assert!(!patch.is_noop());
    assert!(
        commit.diagnostics().touched_components() >= 2,
        "split control write must report every changed native component"
    );
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit
            .package()
            .table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)
            .unwrap_or_else(|error| panic!("split control candidate readback failed: {error}")),
        Some(desired.clone())
    );
    let applied = package
        .apply_table_cell_control_format(&patch)
        .unwrap_or_else(|error| panic!("split control patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert!(package.apply_table_cell_control_format(&patch).is_err());
    let restored = commit
        .package()
        .apply_table_cell_control_format(&patch.inverse())
        .unwrap_or_else(|error| panic!("split control inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    exercise_reopened_replay(
        &source_bytes,
        &target_bytes,
        &patch,
        position,
        commit.diagnostics().touched_components(),
    );

    // A changed Some value is followed by a clear/reset attempt.  When the
    // format/control refcount census admits it, the clear must remove only
    // the selected ownership and preserve every sibling; when it is refused,
    // the candidate remains byte-identical and the fuzz target records the
    // typed error without treating it as a crash.
    let clear_result = commit
        .package()
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)
        .and_then(|edit| {
            if data.first().copied().unwrap_or_default() & 1 == 0 {
                edit.clear().commit()
            } else {
                edit.reset().commit()
            }
        });
    match clear_result {
        Ok(clear) => {
            assert_eq!(clear.patch().after(), None);
            assert!(clear.diagnostics().touched_components() >= 2);
            assert_eq!(
                clear
                    .package()
                    .table_cell_control_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        position,
                    )
                    .unwrap_or_else(|error| panic!("split clear readback failed: {error}")),
                None
            );
            let clear_source = target_bytes.clone();
            let clear_target = package_bytes(clear.package());
            let restored = clear
                .package()
                .apply_table_cell_control_format(&clear.patch().inverse())
                .unwrap_or_else(|error| panic!("split clear inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), clear_source);
            exercise_reopened_replay(
                &clear_source,
                &clear_target,
                clear.patch(),
                position,
                clear.diagnostics().touched_components(),
            );
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(commit.package()), target_bytes);
        },
    }
}

/// Find two equal controls in the split source and clear one.  This exercises
/// the format/control-list refcount path when a native seed shares a model;
/// if this particular source has no equal pair, the scan is still bounded.
fn exercise_split_shared_refcount(
    package: &Package,
    observed: &[(CellPosition, CellControl)],
    data: &[u8],
) {
    let Some((first, value)) = observed.iter().find_map(|(position, value)| {
        observed
            .iter()
            .find(|(other_position, other)| other_position != position && other == value)
            .map(|_| (*position, value.clone()))
    }) else {
        return;
    };
    let Some((sibling, _)) = observed
        .iter()
        .find(|(position, other)| *position != first && *other == value)
    else {
        return;
    };
    let source_bytes = package_bytes(package);
    let result = package
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), first)
        .and_then(|edit| edit.clear().commit());
    match result {
        Ok(commit) => {
            assert!(commit.diagnostics().touched_components() >= 2);
            assert_eq!(
                commit
                    .package()
                    .table_cell_control_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        *sibling,
                    )
                    .unwrap_or_else(|error| panic!("shared control readback failed: {error}")),
                Some(value.clone())
            );
            let target_bytes = package_bytes(commit.package());
            let restored = commit
                .package()
                .apply_table_cell_control_format(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("shared control inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_bytes);
            exercise_reopened_replay(
                &source_bytes,
                &target_bytes,
                commit.patch(),
                first,
                commit.diagnostics().touched_components(),
            );
        },
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
        },
    }
    black_box(data);
}

/// Replay exact and required-minus-one physical/semantic ingress profiles for
/// the split source.  The source must be admitted at its exact byte ceiling,
/// while every one-byte-tight profile fails before any candidate publication.
fn exercise_split_component_limits() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let source_len = u64::try_from(SPLIT_COMPONENT_NUMBERS.len())
            .unwrap_or_else(|error| panic!("split source length conversion failed: {error}"));
        let exact_archive = PackageLimits::new(
            source_len,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("split exact limits invalid: {error}"));
        let exact = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(exact_archive, PackageSemanticLimits::default()),
        )
        .unwrap_or_else(|error| panic!("split source rejected at exact input ceiling: {error}"));
        black_box(exact);
        let tight_archive = PackageLimits::new(
            source_len.saturating_sub(1),
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| panic!("split tight limits invalid: {error}"));
        let result = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(tight_archive, PackageSemanticLimits::default()),
        );
        assert!(
            result.is_err(),
            "split source exceeded input-minus-one gate"
        );
        black_box(result.err().map(|error| error.to_string()));

        // Exercise the remaining public physical axes with deliberately
        // tight, checked profiles.  These are required-minus-one style
        // probes for entries, uncompressed member bytes, aggregate bytes,
        // and IWA stream bytes; every failure must occur before a write can
        // publish a split candidate.
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
            let limits =
                limits.unwrap_or_else(|error| panic!("split {label} limits invalid: {error}"));
            let result = Package::from_bytes_with_options(
                SPLIT_COMPONENT_NUMBERS,
                PackageReadOptions::new(limits, PackageSemanticLimits::default()),
            );
            assert!(
                result.is_err(),
                "split {label} limit unexpectedly admitted source"
            );
            black_box(result.err().map(|error| error.to_string()));
        }

        let object_limit = PackageSemanticLimits::new(1, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
            .unwrap_or_else(|error| panic!("split object limit invalid: {error}"));
        let result = Package::from_bytes_with_options(
            SPLIT_COMPONENT_NUMBERS,
            PackageReadOptions::new(PackageLimits::default(), object_limit),
        );
        assert!(
            result.is_err(),
            "split object-minus-one limit admitted source"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

/// Probe a bounded region of a file-backed fuzz input.  Valid ZIP mutations
/// may carry split model/list/control members; every admitted changed route
/// is sent through the same multi-member replay checks, while unsupported or
/// malformed ownership remains source-atomic.
fn exercise_split_component_window(package: &Package, data: &[u8]) {
    for row in 0..8 {
        for column in 0..8 {
            let position = CellPosition::new(row, column);
            let Ok(before) = package.table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            ) else {
                continue;
            };
            let Some(before) = before else {
                continue;
            };
            let source_bytes = package_bytes(package);
            let desired = control(
                usize::from(
                    data.get((row as usize + column as usize) % data.len().max(1))
                        .copied()
                        .unwrap_or_default(),
                ) % 5,
                data,
            );
            exercise_set(package, data, position, Some(before), source_bytes, desired);
        }
    }
}

fn control_kind(control: &CellControl) -> usize {
    match control {
        CellControl::Checkbox(_) => 0,
        CellControl::StarRating(_) => 1,
        CellControl::Slider(_) => 2,
        CellControl::Stepper(_) => 3,
        CellControl::PopUpMenu(_) => 4,
    }
}

fn control(kind: usize, data: &[u8]) -> CellControl {
    let minimum = f64::from((data.get(3).copied().unwrap_or_default() % 8) + 1);
    let maximum = minimum + f64::from((data.get(4).copied().unwrap_or_default() % 16) + 2);
    let increment = f64::from((data.get(5).copied().unwrap_or_default() % 4) + 1);
    let range = Range::new(minimum, maximum, increment)
        .unwrap_or_else(|error| unreachable!("generated control range is valid: {error}"));
    match kind {
        0 => CellControl::Checkbox(Checkbox),
        1 => CellControl::StarRating(StarRating),
        2 => CellControl::Slider(Slider::new(range, DisplayFormat::Number(Number::default()))),
        3 => CellControl::Stepper(Stepper::new(
            range,
            DisplayFormat::Number(Number::default()),
        )),
        _ => {
            let items = (0..(usize::from(data.get(6).copied().unwrap_or_default() % 3) + 1))
                .map(|index| {
                    format!(
                        "Choice-{index}-{:02x}",
                        data.get(index + 7).copied().unwrap_or_default()
                    )
                })
                .collect::<Vec<_>>();
            let menu = PopUpMenu::new(items)
                .unwrap_or_else(|error| unreachable!("generated popup control is valid: {error}"));
            CellControl::PopUpMenu(menu)
        },
    }
}

fn exercise_constructors(data: &[u8]) {
    observe_result(Range::new(f64::NAN, 1.0, 1.0));
    observe_result(Range::new(1.0, 1.0, 1.0));
    observe_result(Range::new(0.0, 1.0, 0.0));
    observe_result(PopUpMenu::new(Vec::<&str>::new()));
    observe_result(litchi::numbers::cell::data_format::pop_up_menu::Item::new(
        "invalid\ncontrol",
    ));
    if data.first().copied().unwrap_or_default() & 1 != 0 {
        let oversized = "x".repeat(4 * 1024 + 1);
        observe_result(litchi::numbers::cell::data_format::pop_up_menu::Item::new(
            &oversized,
        ));
    }
    black_box(CellControl::Checkbox(Checkbox).to_data_format());
    black_box(CellControl::StarRating(StarRating).to_data_format());
}

fn exercise_selector_errors(package: &Package, data: &[u8]) {
    let position = CellPosition::new(read_u16(data, 0) as u32, read_u16(data, 2) as u32);
    observe_redacted(
        package.table_cell_control_format(
            SheetSelector::name(PRIVATE_SHEET),
            TableSelector::index(0),
            position,
        ),
        PRIVATE_SHEET,
    );
    observe_redacted(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::name(PRIVATE_TABLE),
            position,
        ),
        PRIVATE_TABLE,
    );
}

fn exercise_redacted_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, options()) {
        observe_redacted(
            Err::<(), _>(error),
            std::str::from_utf8(PRIVATE_INPUT).unwrap_or("control-input"),
        );
    }
}

fn exercise_input_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let oversized = vec![0_u8; OVERSIZED_INPUT_BYTES];
        let result = Package::from_bytes_with_options(&oversized, options());
        assert!(
            result.is_err(),
            "oversized control input unexpectedly accepted"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("control package write failed: {error}"));
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
        Err(error) => {
            observe_error(error);
        },
    }
}

fn observe_error<E>(error: E)
where
    E: Debug + Display,
{
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted<T, E>(result: Result<T, E>, private: &str)
where
    T: Debug,
    E: Debug + Display,
{
    if let Err(error) = result {
        let display = error.to_string();
        let debug = format!("{error:?}");
        assert!(!display.contains(private));
        assert!(!debug.contains(private));
        black_box((display, debug));
    }
}
