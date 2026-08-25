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

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected native bytes make arbitrary ZIP mutation unlikely to
    // reach a semantic table.  Reuse the same command against a fixed native
    // source so every campaign still exercises selectors and transactions.
    exercise_package(native_package(), data);
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
    let first_target = commit.package();
    let second_kind = usize::from(data.get(1).copied().unwrap_or_default() % 5);
    if second_kind != control_kind(&desired) {
        let second = control(second_kind, data);
        if let Ok(edit) = first_target.edit_table_cell_control_format(sheet, table, position)
            && let Ok(second_commit) = edit.set(second.clone()).commit()
        {
            let second_patch = second_commit.patch().clone();
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
