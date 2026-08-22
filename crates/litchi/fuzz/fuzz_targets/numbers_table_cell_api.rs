#![no_main]

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::Value,
    table::cells::{Error as CellError, Input, Storage},
    table::{CellPosition, CellRange, Dimensions},
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
const CONTROL_BYTES: usize = 16;
const PRIVATE_SELECTOR: &str = "__litchi_private_table_cell_api_selector_4e8f__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary bytes unlikely to reach semantic table
    // reads.  Reuse every input as a bounded command against the fixed native
    // package so A1 parsing, selector errors, dense limits, and source
    // atomicity remain covered in every campaign.
    exercise_package(native_package(), data);
    exercise_checked_scalar(native_package(), data);
    exercise_table_range_edges(native_package(), data);
    exercise_exact_noop(native_package());
    exercise_constructor_limits();
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
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_NUMBERS, fuzz_options())
            .unwrap_or_else(|error| panic!("native Numbers table-cell seed must open: {error}"));
        let table = package
            .document()
            .sheets()
            .first()
            .and_then(|sheet| sheet.tables().next())
            .unwrap_or_else(|| panic!("native Numbers table-cell seed must expose a table"));
        assert!(table.dimensions().rows() > 0);
        assert!(table.dimensions().columns() > 0);
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
    let dimensions = table.dimensions();
    let position = position(data, dimensions);
    let address =
        ["A1", "B2", "$A$1", "C3", "bad", "A0", "A1048577"][usize::from(control(data, 0)) % 7];

    observe_result(package.table_cell(0usize, 0usize, position));
    observe_result(package.table_cell_a1(0usize, 0usize, address));
    observe_result(package.table_cell(
        SheetSelector::name(sheet.name()),
        TableSelector::name(table.name()),
        position,
    ));
    observe_result(package.table_cell(
        SheetSelector::name(PRIVATE_SELECTOR),
        TableSelector::index(0),
        position,
    ));

    let range = CellRange::single(position)
        .unwrap_or_else(|error| panic!("checked position must form a range: {error}"));
    observe_result(package.table_cells(0usize, 0usize, range));
    observe_result(package.table_cells_a1(0usize, 0usize, address));
    observe_result(package.table_cells_a1(0usize, 0usize, "A1:B2"));
    observe_result(package.table_cells_a1(0usize, 0usize, "bad:range"));

    exercise_presence(package, dimensions);
    exercise_bounded_batch(package, data, position, address);
}

fn exercise_presence(package: &Package, dimensions: Dimensions) {
    let first = CellPosition::new(0, 0);
    if let Ok(state) = package.table_cell(0usize, 0usize, first) {
        match state.storage() {
            Storage::Missing => {
                black_box("missing");
            },
            Storage::Stored(Value::Empty) => {
                black_box("stored-empty");
            },
            Storage::Stored(value) => {
                black_box(value.cell_type());
            },
            _ => {
                black_box("unknown-storage");
            },
        }
    }

    // The one-over dimensions are guaranteed to trigger a typed bounds error
    // whenever the selected table has a positive extent.  This probes the
    // dense API's pre-allocation check without retaining a partial Vec.
    let out_of_bounds = CellPosition::new(dimensions.rows(), dimensions.columns());
    observe_result(package.table_cell(0usize, 0usize, out_of_bounds));
    let malformed = CellRange::new(CellPosition::new(0, 0), out_of_bounds);
    if let Ok(range) = malformed {
        observe_result(package.table_cells(0usize, 0usize, range));
    }
}

fn exercise_bounded_batch(package: &Package, data: &[u8], position: CellPosition, address: &str) {
    let source_bytes = package_bytes(package);
    let before_state = CellPosition::from_a1(address)
        .ok()
        .and_then(|position| package.table_cell(0usize, 0usize, position).ok());
    let input = match control(data, 1) % 4 {
        0 => Input::boolean(control(data, 2) & 1 != 0),
        1 => Input::number(f64::from(control(data, 3)) + 0.5)
            .unwrap_or_else(|error| panic!("finite fuzz number: {error}")),
        2 => Input::text(replacement(data))
            .unwrap_or_else(|error| panic!("bounded fuzz text allocation: {error:?}")),
        _ => Input::duration(f64::from(control(data, 4)) + 1.0)
            .unwrap_or_else(|error| panic!("finite fuzz duration: {error}")),
    };

    let result = package.set_table_cell_a1(0usize, 0usize, address, input);
    match result {
        Ok(commit) => {
            let patch = commit.patch().clone();
            let target = commit.package();
            let target_bytes = package_bytes(target);
            assert_eq!(patch.is_noop(), target_bytes == source_bytes);
            black_box((
                commit.diagnostics(),
                patch.source_fingerprint(),
                patch.target_fingerprint(),
            ));
            let applied = package
                .apply_table_cells(&patch)
                .unwrap_or_else(|error| panic!("table-cell patch must apply: {error}"));
            assert_eq!(package_bytes(applied.package()), target_bytes);
            let restored = target
                .apply_table_cells(&patch.inverse())
                .unwrap_or_else(|error| panic!("table-cell inverse must apply: {error}"));
            assert_eq!(
                package_bytes(restored.package()),
                source_bytes,
                "table-cell inverse did not restore exact source bytes"
            );
            if let (Some(before), Ok(position)) =
                (before_state.as_ref(), CellPosition::from_a1(address))
            {
                let after = restored
                    .package()
                    .table_cell(0usize, 0usize, position)
                    .unwrap_or_else(|error| {
                        panic!("restored table-cell read must succeed: {error}")
                    });
                assert_eq!(after, *before);
            }
        },
        Err(error) => {
            observe_error(error);
        },
    }

    // Exercise checked-coordinate and clear shortcuts independently.  Their
    // errors are expected on locked/unsupported native routes; a successful
    // commit must still invert to the exact source.
    let before = package.table_cell(0usize, 0usize, position).ok();
    match package.clear_table_cell(0usize, 0usize, position) {
        Ok(commit) => {
            let patch = commit.patch().clone();
            let target_bytes = package_bytes(commit.package());
            assert_eq!(patch.is_noop(), target_bytes == source_bytes);
            black_box((patch.source_fingerprint(), patch.target_fingerprint()));
            let inverse = patch.inverse();
            let restored = commit
                .package()
                .apply_table_cells(&inverse)
                .unwrap_or_else(|error| panic!("clear inverse must apply: {error}"));
            assert_eq!(
                package_bytes(restored.package()),
                source_bytes,
                "clear inverse did not restore exact source bytes"
            );
            if let Some(before) = before.as_ref() {
                let after = restored
                    .package()
                    .table_cell(0usize, 0usize, position)
                    .unwrap_or_else(|error| {
                        panic!("restored cleared cell read must succeed: {error}")
                    });
                assert_eq!(after, *before);
            }
        },
        Err(error) => {
            observe_error(error);
        },
    }
    observe_result(package.clear_table_cell_a1(0usize, 0usize, "A0"));
}

/// Keep the checked-coordinate scalar convenience form on a separate route
/// from the A1 command above. The native fixture's B3 cell is a stable,
/// writable scalar target; the command selects every finite scalar variant,
/// including the date path that the A1 batch intentionally does not use.
fn exercise_checked_scalar(package: &Package, data: &[u8]) {
    let position = CellPosition::new(2, 1);
    let source_bytes = package_bytes(package);
    let input = match control(data, 7) % 5 {
        0 => Input::boolean(control(data, 8) & 1 != 0),
        1 => Input::number(f64::from(control(data, 9)) + 0.25)
            .unwrap_or_else(|error| panic!("finite checked scalar number: {error}")),
        2 => Input::text(replacement(data))
            .unwrap_or_else(|error| panic!("bounded checked scalar text: {error:?}")),
        3 => Input::date(f64::from(control(data, 10)) + 0.25)
            .unwrap_or_else(|error| panic!("finite checked scalar date: {error}")),
        _ => Input::duration(f64::from(control(data, 11)) + 0.25)
            .unwrap_or_else(|error| panic!("finite checked scalar duration: {error}")),
    };
    let before = package.table_cell(0usize, 0usize, position).ok();
    let commit = match package.set_table_cell(0usize, 0usize, position, input) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };

    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), target_bytes == source_bytes);
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.inverse().inverse(), patch);
    assert_eq!(diagnostics.requested_cells(), 1);
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    assert_eq!(diagnostics.changed_cells(), usize::from(!patch.is_noop()));
    if patch.is_noop() {
        assert_eq!(patch.source_fingerprint(), patch.target_fingerprint());
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
        assert!(!diagnostics.full_reparse_performed());
    } else {
        assert!(diagnostics.touched_components() > 0);
        assert!(diagnostics.full_reparse_performed());
        assert!(matches!(
            commit.package().apply_table_cells(&patch),
            Err(CellError::PatchConflict)
        ));
    }

    let restored = commit
        .package()
        .apply_table_cells(&patch.inverse())
        .unwrap_or_else(|error| panic!("checked scalar inverse must apply: {error}"));
    assert_eq!(
        package_bytes(restored.package()),
        source_bytes,
        "checked scalar inverse did not restore exact source bytes"
    );
    if let Some(before) = before {
        let after = restored
            .package()
            .table_cell(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("checked scalar inverse read must succeed: {error}"));
        assert_eq!(after, before);
    }
}

/// Exercise range construction and the dense reader's zero-allocation edges.
/// These commands are deliberately separate from the ordinary one-cell and
/// A1 ranges so empty and inverted coordinates reach their own checks.
fn exercise_table_range_edges(package: &Package, data: &[u8]) {
    let Some(table) = package
        .document()
        .sheets()
        .first()
        .and_then(|sheet| sheet.tables().next())
    else {
        return;
    };
    let dimensions = table.dimensions();
    if dimensions.rows() == 0 || dimensions.columns() == 0 {
        return;
    }

    let sheet_name = package
        .document()
        .sheets()
        .first()
        .map(|sheet| sheet.name())
        .unwrap_or_default();
    let table_name = table.name();

    let empty = CellRange::new(
        CellPosition::new(dimensions.rows(), 0),
        CellPosition::new(dimensions.rows(), 0),
    )
    .unwrap_or_else(|error| panic!("empty table-cell range must construct: {error}"));
    let empty_states = package.table_cells(
        SheetSelector::name(sheet_name),
        TableSelector::name(table_name),
        empty,
    );
    match empty_states {
        Ok(states) => assert!(states.is_empty(), "empty range returned cells"),
        Err(error) => observe_error(error),
    }

    let inverted = CellRange::new(CellPosition::new(1, 1), CellPosition::new(0, 0));
    observe_result(inverted);

    let overflow = CellRange::single(CellPosition::new(u32::MAX, u32::MAX));
    observe_result(overflow);

    // Keep one name-selected dense request in this route as well. Its shape
    // is bounded by the source table and the command bytes, so it cannot
    // bypass the package's materialized-cell guard.
    let end = CellPosition::new(dimensions.rows().min(2), dimensions.columns().min(2));
    let bounded = CellRange::new(CellPosition::new(0, 0), end)
        .unwrap_or_else(|error| panic!("bounded table-cell range must construct: {error}"));
    if bounded.area().is_some_and(|area| area > 0) && control(data, 12) & 1 != 0 {
        observe_result(package.table_cells(
            SheetSelector::name(sheet_name),
            TableSelector::name(table_name),
            bounded,
        ));
    }
}

/// B2 is the documented text cell in the native Numbers seed. Re-entering
/// that exact value must stay a byte-level no-op, including when its inverse
/// is applied as a directional patch.
fn exercise_exact_noop(package: &Package) {
    let source_bytes = package_bytes(package);
    let position = CellPosition::new(1, 1);
    let before = package
        .table_cell(0usize, 0usize, position)
        .unwrap_or_else(|error| panic!("native no-op source cell must read: {error}"));
    if !matches!(before.storage(), Storage::Stored(Value::Text(value)) if value == "Litchi native Numbers fixture")
    {
        return;
    }

    let commit = package
        .set_table_cell(
            0usize,
            0usize,
            position,
            Input::text("Litchi native Numbers fixture")
                .unwrap_or_else(|error| panic!("native no-op text must allocate: {error:?}")),
        )
        .unwrap_or_else(|error| panic!("native scalar no-op must commit: {error}"));
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert!(patch.is_noop());
    assert_eq!(patch.len(), 1);
    assert_eq!(patch.source_fingerprint(), patch.target_fingerprint());
    assert_eq!(patch.inverse().inverse(), patch);
    assert!(!diagnostics.changed());
    assert_eq!(diagnostics.changed_cells(), 0);
    assert_eq!(diagnostics.touched_components(), 0);
    assert_eq!(diagnostics.deleted_previews(), 0);
    assert!(!diagnostics.full_reparse_performed());

    let applied = package
        .apply_table_cells(&patch)
        .unwrap_or_else(|error| panic!("native scalar no-op patch must apply: {error}"));
    assert!(!applied.diagnostics().changed());
    assert_eq!(package_bytes(applied.package()), source_bytes);
    let restored = commit
        .package()
        .apply_table_cells(&patch.inverse())
        .unwrap_or_else(|error| panic!("native scalar no-op inverse must apply: {error}"));
    assert_eq!(package_bytes(commit.package()), source_bytes);
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_cell(0usize, 0usize, position)
            .unwrap_or_else(|error| panic!("native scalar no-op inverse read: {error}")),
        before
    );
}

/// Exercise the checked finite-scalar and resource-profile constructors once
/// per process. This keeps invalid-input branches hot without allocating an
/// oversized package or relying on an OS-level memory failure.
fn exercise_constructor_limits() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        assert!(
            PackageLimits::new(
                0,
                MAX_ENTRIES,
                MAX_ENTRY_BYTES,
                MAX_EXPANDED_BYTES,
                MAX_IWA_STREAM_BYTES
            )
            .is_err()
        );
        assert!(
            PackageLimits::new(
                MAX_INPUT_BYTES,
                0,
                MAX_ENTRY_BYTES,
                MAX_EXPANDED_BYTES,
                MAX_IWA_STREAM_BYTES
            )
            .is_err()
        );

        let semantic = PackageSemanticLimits::default();
        assert!(PackageSemanticLimits::new(0, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES).is_err());
        assert!(PackageSemanticLimits::new(MAX_OBJECTS, 0, MAX_TABLES, MAX_REFERENCES).is_err());
        assert!(semantic.with_projection_limits(0, 1).is_err());
        assert!(semantic.with_projection_limits(1, 0).is_err());
        assert!(semantic.with_projection_limits(usize::MAX, 1).is_err());
        assert!(semantic.with_projection_limits(1, usize::MAX).is_err());

        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(Input::number(value).is_err());
            assert!(Input::date(value).is_err());
            assert!(Input::duration(value).is_err());
        }

        assert!(matches!(CellPosition::from_a1(""), Err(_)));
    });
}

fn exercise_input_limit() {
    static ONCE: OnceLock<()> = OnceLock::new();
    ONCE.get_or_init(|| {
        let oversized = vec![0_u8; usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) + 1];
        let result = Package::from_bytes_with_options(&oversized, fuzz_options());
        assert!(
            result.is_err(),
            "oversized Numbers input unexpectedly accepted"
        );
        black_box(result.err().map(|error| error.to_string()));
    });
}

fn position(data: &[u8], dimensions: Dimensions) -> CellPosition {
    CellPosition::new(
        u32::from(control(data, 5)) % dimensions.rows().max(1),
        u32::from(control(data, 6)) % dimensions.columns().max(1),
    )
}

fn replacement(data: &[u8]) -> String {
    let start = data.len().min(CONTROL_BYTES);
    let end = data.len().min(start.saturating_add(128));
    let mut value = String::from_utf8_lossy(&data[start..end]).into_owned();
    value.retain(|character| character != '\0');
    if value.is_empty() {
        value.push('x');
    }
    value
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index % CONTROL_BYTES).copied().unwrap_or_default()
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
        Err(error) => {
            observe_error(error);
        },
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
